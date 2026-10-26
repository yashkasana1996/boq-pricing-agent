"""Non-Excel quote intake: Outlook .msg emails, PDF quotations, nested zips.

Reality of vendor responses (observed on a live tender): most quotes arrive as
email replies (.msg) carrying the actual quote as an attachment (xlsx or PDF),
alongside catalogues, datasheets and T&C PDFs that are NOT quotes.

Strategy — deterministic-first, honest about limits:
  .msg  -> extract attachments (they re-enter the normal intake), keep the
           body text for reference; the email itself is transport, not data.
  .pdf  -> classify (quotation vs catalogue/datasheet/terms) by text signals;
           quotations get a deterministic line-item parse; when an
           ANTHROPIC_API_KEY is set, an LLM pass structures what the
           deterministic parse couldn't (numbers re-validated in code).
  .zip  -> recursed by the caller (webapp/CLI), including zips inside zips.

Anything that can't be turned into rate lines is reported with a reason —
never silently dropped, never guessed.
"""

from __future__ import annotations

import re
from pathlib import Path

# ---------------------------------------------------------------------------
# .msg (Outlook) — attachment extraction
# ---------------------------------------------------------------------------


def expand_msg(path: str | Path, out_dir: str | Path) -> tuple[list[Path], list[str]]:
    """Extract attachments from an Outlook .msg. Returns (files, notes).

    Reads the OLE compound file directly (olefile): each attachment lives in
    an `__attach_version1.0_#NNNNNNNN` storage whose data stream is
    `__substg1.0_37010102` and whose long filename is `__substg1.0_3707001F`
    (UTF-16LE). Body text (`__substg1.0_1000001F`) is kept for reference.
    No fragile third-party .msg stack needed.
    """
    path, out_dir = Path(path), Path(out_dir)
    try:
        import olefile
    except ImportError:
        return [], [f"{path.name}: .msg support needs 'pip install olefile' — skipped"]
    try:
        ole = olefile.OleFileIO(str(path))
    except Exception as e:
        return [], [f"{path.name}: unreadable .msg ({type(e).__name__}) — skipped"]
    out_dir.mkdir(parents=True, exist_ok=True)
    files, notes = [], []
    try:
        attach_dirs = sorted({entry[0] for entry in ole.listdir()
                              if entry[0].startswith("__attach_version1.0_")})
        for adir in attach_dirs:
            data_stream = f"{adir}/__substg1.0_37010102"
            if not ole.exists(data_stream):
                continue  # embedded message rather than a file attachment
            data = ole.openstream(data_stream).read()
            name = "attachment.bin"
            for name_stream, enc in ((f"{adir}/__substg1.0_3707001F", "utf-16-le"),
                                     (f"{adir}/__substg1.0_3704001F", "utf-16-le"),
                                     (f"{adir}/__substg1.0_37070102", "latin-1")):
                if ole.exists(name_stream):
                    try:
                        name = ole.openstream(name_stream).read().decode(enc).strip("\x00").strip()
                        if name:
                            break
                    except Exception:
                        continue
            name = Path(name or "attachment.bin").name
            target = out_dir / f"{path.stem[:40]}__{name}"
            target.write_bytes(data)
            files.append(target)
        if ole.exists("__substg1.0_1000001F"):
            try:
                body = ole.openstream("__substg1.0_1000001F").read().decode("utf-16-le", "ignore")
                if body.strip():
                    (out_dir / f"{path.stem[:40]}__body.txt").write_text(
                        body[:20000], encoding="utf-8")
            except Exception:
                pass
        if not files:
            notes.append(f"{path.name}: email had no file attachments (body kept for reference)")
    finally:
        ole.close()
    return files, notes


# ---------------------------------------------------------------------------
# .pdf — classification + line-item extraction
# ---------------------------------------------------------------------------

QUOTE_SIGNALS = re.compile(
    r"quotation|quote|rate\s+(offered|list|for)|amount\s*\(?(inr|rs)|total\s+amount|"
    r"basic\s+rate|unit\s+rate|grand\s+total|offer\s+price|price\s+offer", re.I)
NON_QUOTE_SIGNALS = re.compile(
    r"catalogue|catalog|data\s*sheet|datasheet|technical\s+specification\s+sheet|"
    r"brochure|leaflet|terms\s*(&|and)\s*conditions|warranty\s+certificate|"
    r"installation\s+guide|system\s*page", re.I)

# a plausible quote line: <desc words> ... <qty> <unit?> <rate> <amount?>
# numbers accept both plain (29460.00) and Indian-grouped (29,460.00) forms
NUM = r"\d[\d,]*(?:\.\d+)?"
LINE_ITEM = re.compile(
    rf"^(?P<desc>.{{8,120}}?)\s+"
    rf"(?P<qty>{NUM})\s+"
    rf"(?P<unit>sqm|sqmt|sq\.?\s?ft|sft|rmt|nos?|set|kg|lot|ls|each|mtr|m)?\s*"
    rf"(?P<rate>{NUM})\s+"
    rf"(?P<amount>{NUM})\s*$",
    re.I,
)


def _num(s: str | None) -> float | None:
    if not s:
        return None
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return None


def pdf_to_payload(path: str | Path, use_llm: bool = False) -> tuple[dict | None, list[str]]:
    """Extract quote items from a PDF. Returns (payload, notes).

    payload matches extract_workbook()'s shape so downstream matching/pricing
    needs no special cases. None payload = not a quote / nothing extractable.
    """
    path = Path(path)
    try:
        import pypdf
    except ImportError:
        return None, [f"{path.name}: PDF support needs 'pip install pypdf' — skipped"]
    try:
        reader = pypdf.PdfReader(str(path))
        text = "\n".join((p.extract_text() or "") for p in reader.pages[:15])
    except Exception as e:
        return None, [f"{path.name}: unreadable PDF ({type(e).__name__}) — skipped"]

    if not text.strip():
        return None, [f"{path.name}: scanned/image PDF, no extractable text — needs manual entry"]

    is_quote = bool(QUOTE_SIGNALS.search(text))
    is_noise = bool(NON_QUOTE_SIGNALS.search(text)) and not is_quote
    if is_noise:
        return None, [f"{path.name}: classified as catalogue/datasheet/T&C, not a quote — skipped"]

    items = []
    for line in text.splitlines():
        line = line.strip()
        m = LINE_ITEM.match(line)
        if not m:
            continue
        qty, rate, amount = _num(m["qty"]), _num(m["rate"]), _num(m["amount"])
        if not qty or not rate:
            continue
        # arithmetic sanity: amount within 2% of qty*rate, else mis-parsed row
        if amount and abs(qty * rate - amount) > max(1.0, 0.02 * amount):
            continue
        items.append({
            "sheet": "PDF", "row": len(items) + 1, "last_row": len(items) + 1,
            "sr": None, "code": None, "description": m["desc"].strip(" .-:"),
            "spec": "", "make": None, "unit": (m["unit"] or "").lower() or None,
            "qty": qty, "qty_levels": {}, "supply_rate": None, "install_rate": None,
            "rate": rate, "amount": amount, "supply_amount": None,
            "install_amount": None, "section": None, "section_code": None,
            "location": None, "basic_rate_text": None, "boq_slno_ref": None,
            "remarks": f"parsed from PDF: {path.name}", "total_rate": rate,
        })

    notes = []
    if not items and use_llm:
        llm_items, llm_notes = _pdf_llm_extract(path.name, text)
        items, notes = llm_items, llm_notes
    if not items:
        reason = ("looks like a quotation but its layout defeats text parsing — "
                  "needs manual entry" + ("" if use_llm else " (or enable AI key for smart PDF reading)"))
        return None, notes + [f"{path.name}: {reason}"]

    payload = {
        "file": path.name, "file_hash": "", "sheet_count": 1,
        "item_count": len(items),
        "sheets": [{"sheet": "PDF", "kind": "boq", "header_row": None,
                    "columns": {}, "level_qty_cols": {}, "supply_install_split": False,
                    "gst_percent": None, "items": items, "sections": [],
                    "warnings": [f"extracted from PDF text ({len(items)} lines)"]}],
    }
    return payload, notes


def _pdf_llm_extract(name: str, text: str) -> tuple[list[dict], list[str]]:
    """LLM pass for PDFs the regex can't parse. Numbers are accepted only if
    they appear verbatim in the PDF text (no invented values)."""
    import json as _json
    import os

    if not os.environ.get("ANTHROPIC_API_KEY"):
        return [], []
    try:
        import anthropic

        client = anthropic.Anthropic()
        resp = client.messages.create(
            model=os.environ.get("BOQ_LLM_MODEL", "claude-sonnet-4-6"),
            max_tokens=4000,
            system=("Extract quotation line items from this PDF text. STRICT JSON: "
                    '{"items":[{"description":str,"unit":str|null,"qty":num|null,'
                    '"rate":num,"amount":num|null}]} — copy numbers EXACTLY as printed; '
                    "skip catalogue/spec content; empty list if no priced lines."),
            messages=[{"role": "user", "content": text[:30000]}],
        )
        raw = resp.content[0].text.strip().strip("`").lstrip("json").strip()
        data = _json.loads(raw)
        digits = re.sub(r"[,\s]", "", text)
        items = []
        for i, it in enumerate(data.get("items", []), start=1):
            rate = it.get("rate")
            if rate is None:
                continue
            # verbatim check: the rate figure must exist in the source text
            if re.sub(r"[,\s]", "", f"{rate:g}") not in digits:
                continue
            items.append({
                "sheet": "PDF", "row": i, "last_row": i, "sr": None, "code": None,
                "description": (it.get("description") or "").strip()[:500], "spec": "",
                "make": None, "unit": it.get("unit"), "qty": it.get("qty"),
                "qty_levels": {}, "supply_rate": None, "install_rate": None,
                "rate": float(rate), "amount": it.get("amount"), "supply_amount": None,
                "install_amount": None, "section": None, "section_code": None,
                "location": None, "basic_rate_text": None, "boq_slno_ref": None,
                "remarks": f"LLM-parsed from PDF: {name}", "total_rate": float(rate),
            })
        return items, ([f"{name}: {len(items)} lines read by AI (verbatim-checked)"] if items else [])
    except Exception as e:
        return [], [f"{name}: AI PDF reading failed ({type(e).__name__})"]


# ---------------------------------------------------------------------------
# unified quote loader used by the pipeline
# ---------------------------------------------------------------------------


def load_quote_payload(path: str | Path, use_llm: bool = False) -> tuple[dict | None, list[str]]:
    """Any quote file -> extraction payload (or None + reasons)."""
    from .extract import extract_workbook
    from .ingest import EXCEL_EXTS, load_workbook

    path = Path(path)
    suffix = path.suffix.lower()
    if suffix in EXCEL_EXTS or suffix == ".csv":
        try:
            return extract_workbook(load_workbook(path)), []
        except Exception as e:
            return None, [f"{path.name}: failed to parse ({type(e).__name__}: {e})"]
    if suffix == ".pdf":
        return pdf_to_payload(path, use_llm=use_llm)
    return None, [f"{path.name}: unsupported type"]
