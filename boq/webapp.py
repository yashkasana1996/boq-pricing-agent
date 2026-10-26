"""BOQ Intelligence Console — web UI over the same pipeline the CLI uses.

Run:  uvicorn boq.webapp:app          (pip install fastapi uvicorn python-multipart)

Intake accepts Excel/CSV quotes, PDF quotations, Outlook .msg emails (their
attachments are pulled automatically) and zips of any of the above, nested
zips included. Whatever can't yield rate lines is reported with a reason —
nothing is silently dropped.
"""

from __future__ import annotations

import html as _html
import json
import os
import shutil
import traceback
import zipfile
from functools import lru_cache
from pathlib import Path

try:
    from fastapi import FastAPI, File, Form, UploadFile
    from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse
except ImportError as e:  # keep the core pipeline importable without fastapi
    raise ImportError("web UI needs: pip install fastapi uvicorn python-multipart") from e

from .ingest import EXCEL_EXTS, is_junk
from .pipeline import fill_pipeline
from .report import render_report

RUNS_DIR = Path("runs")
UPLOADS = Path("uploads")
KB_DIR = "knowledge"
DIRECT_EXTS = set(EXCEL_EXTS) | {".csv", ".pdf"}  # go straight to the pipeline
EXPAND_EXTS = {".zip", ".msg"}                     # unpacked into more files
MAX_EXPAND_DEPTH = 3

app = FastAPI(title="BOQ Intelligence Console")

# ---------------------------------------------------------------------------
# presentation
# ---------------------------------------------------------------------------

STYLE = """
:root{--ink:#1a2233;--muted:#5b6472;--line:#e2e5ea;--accent:#1f3a5f;
      --accent-ink:#ffffff;--bg:#f6f7f9;--card:#ffffff;--ok:#1f6f43;--warn:#8a5a00}
*{box-sizing:border-box}
body{font-family:"Segoe UI",system-ui,-apple-system,sans-serif;margin:0;background:var(--bg);
     color:var(--ink);line-height:1.55}
header.top{background:var(--accent);color:var(--accent-ink);padding:.8rem 2rem;
     display:flex;align-items:baseline;gap:1rem}
header.top .brand{font-size:1.05rem;font-weight:600;letter-spacing:.04em}
header.top .sub{font-size:.78rem;opacity:.75}
main{max-width:1280px;margin:1.6rem auto;padding:0 1.5rem}
h1{font-size:1.25rem;font-weight:600;margin:.2rem 0 1rem}
h2{font-size:.8rem;font-weight:600;letter-spacing:.09em;text-transform:uppercase;
   color:var(--muted);margin:1.8rem 0 .6rem}
.card{background:var(--card);border:1px solid var(--line);border-radius:8px;
      padding:1.1rem 1.3rem;margin-bottom:1.2rem}
table{border-collapse:collapse;width:100%;font-size:.86rem}
td,th{border-bottom:1px solid var(--line);padding:.5rem .65rem;text-align:left;
      vertical-align:top}
th{font-size:.72rem;letter-spacing:.06em;text-transform:uppercase;color:var(--muted);
   font-weight:600;background:transparent}
tr:hover td{background:#fafbfc}
.num{text-align:right;font-variant-numeric:tabular-nums}
.btn{display:inline-block;padding:.45rem 1rem;border:1px solid var(--accent);
     border-radius:6px;text-decoration:none;color:var(--accent);background:#fff;
     font-size:.85rem;margin:.15rem .2rem .15rem 0;cursor:pointer}
.btn.primary{background:var(--accent);color:var(--accent-ink)}
.btn:hover{filter:brightness(1.08)}
.note{color:var(--warn);font-size:.84rem}
.help{color:var(--muted);font-size:.82rem}
input,select{padding:.45rem .5rem;border:1px solid var(--line);border-radius:6px;
     font-size:.88rem;background:#fff}
label{font-size:.82rem;color:var(--muted);display:block;margin-bottom:.15rem}
.field{display:inline-block;margin:0 1.1rem .8rem 0;vertical-align:top}
.grid{font-size:11.5px}
.grid td{padding:2px 6px;max-width:430px;overflow:hidden;text-overflow:ellipsis;
     white-space:nowrap;border:1px solid #eceef1}
.grid td.wrap{white-space:normal;max-width:540px}
.grid .filled{background:#eef6ef;font-weight:600;color:var(--ok)}
.grid .rownum{background:#f4f5f7;color:#9aa1ab;text-align:right;font-size:10px}
.sheettabs{margin:.5rem 0}
.sheettabs a{display:inline-block;margin:2px 3px 2px 0;padding:3px 10px;
     border:1px solid var(--line);border-radius:4px;text-decoration:none;
     color:var(--ink);font-size:.76rem;background:#fff}
.sheettabs a.active{background:var(--accent);color:#fff;border-color:var(--accent)}
.kpis{display:flex;gap:1rem;flex-wrap:wrap;margin:.4rem 0 1rem}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:8px;
     padding:.7rem 1.1rem;min-width:150px}
.kpi .v{font-size:1.15rem;font-weight:600;font-variant-numeric:tabular-nums}
.kpi .k{font-size:.7rem;letter-spacing:.07em;text-transform:uppercase;color:var(--muted)}
details summary{cursor:pointer;color:var(--muted);font-size:.84rem}
pre{white-space:pre-wrap;font-size:.82rem}
.aibadge{margin-left:auto;font-size:.72rem;padding:.2rem .6rem;border-radius:999px;
     border:1px solid rgba(255,255,255,.35)}
.aibadge.on{background:rgba(120,220,160,.18);color:#d6f5e2;border-color:rgba(120,220,160,.5)}
.aibadge.off{background:rgba(255,255,255,.08);color:#c3c9d2}
"""


def ai_enabled() -> bool:
    """True when an Anthropic API key is present in the environment."""
    return bool(os.environ.get("ANTHROPIC_API_KEY"))


def _ai_badge() -> str:
    if ai_enabled():
        return "<span class='aibadge on' title='ANTHROPIC_API_KEY detected'>AI reading: ON</span>"
    return ("<span class='aibadge off' title='Set ANTHROPIC_API_KEY to enable "
            "AI reading of tricky PDFs and leftover matches'>AI reading: OFF</span>")


def _page(title: str, body: str) -> str:
    return f"""<!doctype html><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{title}</title><style>{STYLE}</style>
<header class="top"><span class="brand">BOQ INTELLIGENCE CONSOLE</span>
<span class="sub">estimation · pricing · audit</span>{_ai_badge()}</header>
<main><h1>{title}</h1>{body}</main>"""


def _error_page(title: str, err: Exception) -> HTMLResponse:
    tb = traceback.format_exc(limit=4)
    body = f"""<div class="card"><p class="note"><b>The run could not be completed:</b>
{_html.escape(str(err))}</p>
<p class="help">Common causes: the BOQ must be an Excel workbook (.xlsx) or a zip
containing one; quotes may be Excel, CSV, PDF or Outlook .msg files, individually
or zipped. Catalogues and datasheets are recognised and set aside automatically.</p>
<details><summary>Technical detail</summary><pre>{_html.escape(tb)}</pre></details>
<p><a class="btn" href="/">Back to dashboard</a></p></div>"""
    return HTMLResponse(_page(title, body), status_code=500)


# ---------------------------------------------------------------------------
# intake: recursive expansion of zips and Outlook emails
# ---------------------------------------------------------------------------


def _save_upload(up: UploadFile, dest_dir: Path) -> Path:
    dest_dir.mkdir(parents=True, exist_ok=True)
    p = dest_dir / Path(up.filename).name
    with open(p, "wb") as f:
        shutil.copyfileobj(up.file, f)
    return p


def _expand_files(paths: list[Path]) -> tuple[list[Path], list[str]]:
    """Resolve every upload to pipeline-ready files. Zips and .msg emails are
    unpacked recursively (nested zips included); everything unusable gets a
    reasoned note."""
    out, notes = [], []
    queue: list[tuple[Path, int]] = [(p, 0) for p in paths]
    while queue:
        p, depth = queue.pop(0)
        suffix = p.suffix.lower()
        if is_junk(p):
            continue  # Office lock files etc. — not worth a note
        if suffix == ".zip":
            if depth >= MAX_EXPAND_DEPTH:
                notes.append(f"{p.name}: zip nested too deep — skipped")
                continue
            extract_dir = p.parent / (p.stem + "_x")
            extract_dir.mkdir(exist_ok=True)
            try:
                with zipfile.ZipFile(p) as z:
                    z.extractall(extract_dir)
            except zipfile.BadZipFile:
                notes.append(f"{p.name}: not a valid zip — skipped")
                continue
            queue += [(m, depth + 1) for m in sorted(extract_dir.rglob("*")) if m.is_file()]
        elif suffix == ".msg":
            from .docs_extract import expand_msg

            files, msg_notes = expand_msg(p, p.parent / (p.stem[:40] + "_att"))
            notes += msg_notes
            queue += [(f, depth + 1) for f in files]
        elif suffix in DIRECT_EXTS:
            out.append(p)
        elif suffix in {".txt"}:
            pass  # email bodies kept on disk for reference, not parsed
        else:
            notes.append(f"{p.name}: unsupported type ({suffix or 'no extension'}) — skipped")
    return out, notes


# ---------------------------------------------------------------------------
# routes
# ---------------------------------------------------------------------------


@app.get("/", response_class=HTMLResponse)
def dashboard():
    rows = []
    if RUNS_DIR.exists():
        for rd in sorted(RUNS_DIR.iterdir(), reverse=True):
            sf = rd / "state.json"
            if not sf.exists():
                continue
            s = json.loads(sf.read_text(encoding="utf-8"))
            t = s["totals"]
            rows.append(
                f"<tr><td><a href='/runs/{rd.name}/'>{rd.name}</a></td>"
                f"<td>{_html.escape(Path(s['boq_file']).name)}</td>"
                f"<td class='num'>{t['items_priced']}/{t['items_total']}</td>"
                f"<td class='num'>{t['grand_total']:,.0f}</td>"
                f"<td class='num'>{t['margin_pct'] if t['margin_pct'] is not None else '—'}</td>"
                f"<td class='num'>{len(s['review_queue'])}</td>"
                f"<td><a class='btn' href='/runs/{rd.name}/view'>Open BOQ</a></td></tr>"
            )
    ai_line = ("AI reading is ON — tricky PDF quotes and leftover matches are read by AI "
               "(numbers are verified against the source before use)."
               if ai_enabled() else
               "AI reading is OFF — Excel/CSV and clean PDF quotes still work; set "
               "ANTHROPIC_API_KEY and restart to enable AI reading of awkward PDFs.")
    body = f"""
<div class="card">
<h2>New pricing run</h2>
<form action="/upload" method="post" enctype="multipart/form-data">
 <div class="field"><label>Client BOQ (.xlsx or .zip)</label>
   <input type="file" name="boq" accept=".xlsx,.xlsm,.zip" required></div>
 <div class="field"><label>Vendor quotes — files or one .zip of the folder
   (Excel, CSV, PDF, Outlook .msg all accepted)</label>
   <input type="file" name="quotes" accept=".xlsx,.xlsm,.csv,.pdf,.msg,.zip" multiple></div>
 <br>
 <div class="field"><label>Markup %</label><input name="markup" type="number" step="0.01" value="12"></div>
 <div class="field"><label>GST treatment</label><select name="gst_mode">
   <option value="exclusive">exclusive (added on total)</option>
   <option value="inclusive">inclusive (within rates)</option></select></div>
 <div class="field"><label>GST %</label><input name="gst_pct" type="number" step="0.01" value="18"></div>
 <div class="field"><label>Leasable area (sqft)</label><input name="area" type="number" step="1"></div>
 <br><button class="btn primary">Run pricing</button>
 <p class="help">{ai_line}</p>
</form></div>
<h2>Pricing runs</h2>
<div class="card"><table>
<tr><th>Run</th><th>BOQ</th><th>Priced</th><th>Grand total</th><th>Margin %</th><th>Review</th><th></th></tr>
{''.join(rows) or '<tr><td colspan=7 class="help">No runs yet — start one above.</td></tr>'}</table></div>"""
    return _page("Dashboard", body)


@app.post("/upload")
async def upload(boq: UploadFile = File(...), quotes: list[UploadFile] = File(default=[]),
                 markup: float = Form(12.0), gst_mode: str = Form("exclusive"),
                 gst_pct: float = Form(18.0), area: float | None = Form(None)):
    try:
        boq_raw = _save_upload(boq, UPLOADS)
        boq_paths, boq_notes = _expand_files([boq_raw])
        boq_xlsx = [p for p in boq_paths if p.suffix.lower() in EXCEL_EXTS]
        if not boq_xlsx:
            raise ValueError(f"no usable BOQ Excel file in upload ({boq_raw.name})")
        boq_path = boq_xlsx[0]

        qraw = [_save_upload(q, UPLOADS) for q in quotes if q.filename]
        qpaths, notes = _expand_files(qraw)

        summary = fill_pipeline(str(boq_path), [str(q) for q in qpaths], markup_pct=markup,
                                gst_mode=gst_mode, gst_pct=gst_pct, area_sqft=area,
                                out_dir=str(RUNS_DIR), kb_dir=KB_DIR,
                                use_llm=ai_enabled())
        run_dir = Path(summary["run_dir"])
        state = json.loads((run_dir / "state.json").read_text(encoding="utf-8"))
        state["upload_notes"] = boq_notes + notes
        (run_dir / "state.json").write_text(json.dumps(state, indent=1, default=str),
                                            encoding="utf-8")
        render_report(run_dir)
        return RedirectResponse(f"/runs/{run_dir.name}/", status_code=303)
    except Exception as e:
        return _error_page("Pricing run failed", e)


@app.get("/runs/{run}/", response_class=HTMLResponse)
def run_detail(run: str):
    rd = RUNS_DIR / Path(run).name
    report = rd / "audit-report.md"
    if not report.exists():
        render_report(rd)
    state = json.loads((rd / "state.json").read_text(encoding="utf-8"))
    t = state["totals"]
    notes = (state.get("upload_notes") or []) + (state.get("file_notes") or [])
    notes_html = ""
    if notes:
        items = "".join(f"<li>{_html.escape(n)}</li>" for n in notes)
        notes_html = (f"<div class='card'><h2>Files set aside ({len(notes)})</h2>"
                      f"<ul class='note'>{items}</ul></div>")
    kpis = f"""<div class="kpis">
<div class="kpi"><div class="v">{t['grand_total']:,.0f}</div><div class="k">Grand total</div></div>
<div class="kpi"><div class="v">{t['items_priced']}/{t['items_total']}</div><div class="k">Items priced</div></div>
<div class="kpi"><div class="v">{t['margin_pct'] if t['margin_pct'] is not None else '—'}</div><div class="k">Margin %</div></div>
<div class="kpi"><div class="v">{len(state['review_queue'])}</div><div class="k">Awaiting review</div></div>
</div>"""
    a = state.get("boq_analysis")
    analysis_html = ""
    if a:
        rows = "".join(
            f"<tr><td>{_html.escape(s['sheet'])}</td><td>{s['kind']}</td>"
            f"<td class='num'>{s['items']}</td><td>{', '.join(s['columns_detected'])}</td>"
            f"<td>{'✓' if s['has_qty'] else ''}</td><td>{'✓' if s['has_rate_column'] else ''}</td>"
            f"<td>{'✓' if s['supply_install_split'] else ''}</td></tr>"
            for s in a["sheets"] if s["kind"] in ("boq", "make_list", "price_list") or s["items"])
        analysis_html = f"""<div class="card"><h2>BOQ analysis</h2>
<p class="help">Detected before pricing — {a['boq_type']}, {a['priceable_sheets']} priceable
sheet(s) of {a['total_sheets']}, {a['total_items']} items.</p>
<details><summary>Per-sheet structure</summary>
<table><tr><th>Sheet</th><th>Kind</th><th>Items</th><th>Columns detected</th>
<th>Qty</th><th>Rate</th><th>S/I split</th></tr>{rows}</table></details></div>"""
    body = f"""
<p><a class="btn" href="/">Dashboard</a>
   <a class="btn primary" href="/runs/{run}/view">View filled BOQ</a>
   <a class="btn" href="/runs/{run}/review">Review queue</a>
   <a class="btn" href="/runs/{run}/download/client-copy.xlsx">Export client copy (Excel)</a>
   <a class="btn" href="/runs/{run}/download/internal-workbook.xlsx">Internal workbook</a></p>
{kpis}
{analysis_html}
{notes_html}
<div class="card"><pre>{_html.escape(report.read_text(encoding="utf-8"))}</pre></div>"""
    return _page(f"Run {run}", body)


MAX_VIEW_ROWS = 250


@lru_cache(maxsize=4)
def _load_view_workbook(path_str: str, mtime: float):
    import openpyxl

    return openpyxl.load_workbook(path_str, data_only=True, read_only=False)


@app.get("/runs/{run}/view", response_class=HTMLResponse)
def view_boq(run: str, sheet: str | None = None, start: int = 1):
    try:
        rd = RUNS_DIR / Path(run).name
        copy = rd / "client-copy.xlsx"
        if not copy.exists():
            return HTMLResponse(_page("No client copy", "<p>This run has no client-copy.xlsx</p>"), 404)
        state = json.loads((rd / "state.json").read_text(encoding="utf-8"))
        filled = {(i["sheet"], i["row"]) for i in state["items"] if i["status"] == "priced"}
        wb = _load_view_workbook(str(copy), copy.stat().st_mtime)

        names = wb.sheetnames
        sheet = sheet if sheet in names else (
            next((n for n in names if any(k[0] == n for k in filled)), names[0]))
        ws = wb[sheet]

        tabs = []
        for n in names:
            has_fill = any(k[0] == n for k in filled)
            cls = "active" if n == sheet else ""
            label = _html.escape(n) + (" ●" if has_fill else "")
            tabs.append(f"<a class='{cls}' href='/runs/{run}/view?sheet={_html.escape(n)}'>{label}</a>")

        from openpyxl.utils import get_column_letter

        start = max(1, start)
        end = min(ws.max_row, start + MAX_VIEW_ROWS - 1)
        max_col = min(ws.max_column, 20)
        rows_html = ["<tr><td class='rownum'></td>" + "".join(
            f"<td class='rownum'>{get_column_letter(c)}</td>" for c in range(1, max_col + 1)) + "</tr>"]
        for r in range(start, end + 1):
            cells = [f"<td class='rownum'>{r}</td>"]
            row_filled = (sheet, r) in filled
            for c in range(1, max_col + 1):
                v = ws.cell(row=r, column=c).value
                txt = "" if v is None else str(v)
                if len(txt) > 400:
                    txt = txt[:400] + "…"
                cls = []
                if len(txt) > 60:
                    cls.append("wrap")
                if row_filled and isinstance(v, (int, float)):
                    cls.append("filled")
                cells.append(f"<td class='{' '.join(cls)}'>{_html.escape(txt)}</td>")
            rows_html.append("<tr>" + "".join(cells) + "</tr>")

        nav = []
        if start > 1:
            nav.append(f"<a class='btn' href='/runs/{run}/view?sheet={_html.escape(sheet)}&start={max(1, start - MAX_VIEW_ROWS)}'>Previous rows</a>")
        if end < ws.max_row:
            nav.append(f"<a class='btn' href='/runs/{run}/view?sheet={_html.escape(sheet)}&start={end + 1}'>Next rows</a>")

        body = f"""
<p><a class="btn" href="/runs/{run}/">Run summary</a>
   <a class="btn primary" href="/runs/{run}/download/client-copy.xlsx">Export filled BOQ (Excel)</a></p>
<p class="help">Highlighted cells were filled by this run. Sheets marked ● contain filled
items. Rows {start}–{end} of {ws.max_row}. Formula cells not yet recalculated by Excel
display blank here; they compute when the exported file is opened.</p>
<div class="sheettabs">{''.join(tabs)}</div>
<p>{' '.join(nav)}</p>
<div class="card" style="overflow-x:auto;padding:.4rem"><table class="grid">{''.join(rows_html)}</table></div>
<p>{' '.join(nav)}</p>"""
        return HTMLResponse(_page(f"Filled BOQ — {_html.escape(sheet)}", body))
    except Exception as e:
        return _error_page("Viewer failed", e)


@app.get("/runs/{run}/review", response_class=HTMLResponse)
def review(run: str):
    state = json.loads((RUNS_DIR / Path(run).name / "state.json").read_text(encoding="utf-8"))
    rows = []
    for i, r in enumerate(state["review_queue"]):
        resolved = r.get("resolution")
        action = (f"<b>{resolved}</b>" if resolved else f"""<form action="/runs/{run}/resolve" method="post">
 <input type="hidden" name="idx" value="{i}">
 <select name="verdict"><option value="accept">accept match</option>
 <option value="reject">reject</option></select>
 <button class="btn">Save</button></form>""")
        rows.append(f"""<tr><td>{_html.escape(str(r.get('vendor', '?')))}</td>
<td>{_html.escape(r['quote']['description'][:100])}</td>
<td>{_html.escape(r['boq']['sheet'])} r{r['boq']['row']}: {_html.escape(r['boq']['description'][:100])}</td>
<td class="num">{r['confidence']}</td><td>{action}</td></tr>""")
    body = f"""<p><a class="btn" href="/runs/{run}/">Run summary</a></p>
<p class="help">These matches were not confident enough to apply automatically.
Your decisions are remembered and improve future matching.</p>
<div class="card"><table>
<tr><th>Vendor</th><th>Quote line</th><th>Proposed BOQ item</th><th>Confidence</th><th>Decision</th></tr>
{''.join(rows) or '<tr><td colspan=5 class="help">Queue is empty.</td></tr>'}</table></div>"""
    return _page(f"Review queue — {run}", body)


@app.post("/runs/{run}/resolve")
async def resolve(run: str, idx: int = Form(...), verdict: str = Form(...)):
    sf = RUNS_DIR / Path(run).name / "state.json"
    state = json.loads(sf.read_text(encoding="utf-8"))
    if 0 <= idx < len(state["review_queue"]):
        r = state["review_queue"][idx]
        r["resolution"] = verdict
        if verdict == "accept":
            from .match import resolve_review

            resolve_review(KB_DIR, r["quote"]["description"][:80], r["boq"]["description"][:80])
        sf.write_text(json.dumps(state, indent=1, default=str), encoding="utf-8")
    return RedirectResponse(f"/runs/{run}/review", status_code=303)


@app.get("/runs/{run}/download/{artifact}")
def download(run: str, artifact: str):
    safe = Path(artifact).name  # no path traversal
    path = RUNS_DIR / Path(run).name / safe
    if not path.exists():
        return HTMLResponse("not found", status_code=404)
    return FileResponse(path, filename=safe)
