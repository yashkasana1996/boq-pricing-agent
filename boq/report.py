"""Audit report per run (CLAUDE.md non-negotiable: every run explains every
number). Renders a fill run's state.json to Markdown + a self-contained HTML."""

from __future__ import annotations

import html
import json
from pathlib import Path


def render_report(run_dir: str | Path) -> Path:
    run_dir = Path(run_dir)
    state = json.loads((run_dir / "state.json").read_text(encoding="utf-8"))

    md = []
    md.append(f"# Pricing run audit — {run_dir.name}\n")
    md.append(f"- BOQ: `{state['boq_file']}`")
    md.append(f"- Quotes: {', '.join('`' + q + '`' for q in state['quote_files']) or '(none)'}")
    md.append(f"- Markup: **{state['markup_pct']}%** | GST: {state['gst_pct']}% "
              f"({state['gst_mode']}) | Round: {state['round_digits']} digits"
              + (f" | Area: {state['area_sqft']:,.0f} sqft" if state.get("area_sqft") else ""))

    a = state.get("boq_analysis")
    if a:
        md.append("\n## BOQ analysis (detected before pricing)\n")
        md.append(f"- Type: **{a['boq_type']}** | {a['priceable_sheets']} priceable "
                  f"sheet(s) of {a['total_sheets']} | {a['total_items']} items "
                  f"({a['priceable_items']} priceable)")
        md.append("\n| Sheet | Kind | Items | Columns detected | Qty | Rate col | Split |")
        md.append("|---|---|---|---|---|---|---|")
        for s in a["sheets"]:
            if s["kind"] not in ("boq", "make_list", "price_list") and s["items"] == 0:
                continue
            md.append(f"| {s['sheet']} | {s['kind']} | {s['items']} | "
                      f"{', '.join(s['columns_detected'])} | {'✓' if s['has_qty'] else ''} | "
                      f"{'✓' if s['has_rate_column'] else ''} | "
                      f"{'✓' if s['supply_install_split'] else ''} |")

    t = state["totals"]
    md.append("\n## Totals\n")
    md.append(f"| Subtotal (ex GST) | GST | Grand total | Cost basis | Margin | Margin % |")
    md.append("|---|---|---|---|---|---|")
    md.append(f"| {t['subtotal_ex_gst']:,.2f} | {t['gst_amount']:,.2f} | "
              f"**{t['grand_total']:,.2f}** | {t['cost_basis']:,.2f} | "
              f"{t['margin']:,.2f} | {t['margin_pct'] if t['margin_pct'] is not None else '—'} |")
    if t.get("rate_per_sqft_ex_gst"):
        md.append(f"\n₹/sqft (ex GST): **{t['rate_per_sqft_ex_gst']}**")

    md.append("\n## Packages\n")
    md.append("| Package | Amount | Cost | Margin | Margin % | ₹/sqft | Priced/Items |")
    md.append("|---|---|---|---|---|---|---|")
    for pkg, p in sorted(state["package_totals"].items()):
        md.append(f"| {pkg} | {p['amount']:,.2f} | {p['cost']:,.2f} | {p['margin']:,.2f} | "
                  f"{p['margin_pct'] if p['margin_pct'] is not None else '—'} | "
                  f"{p['rate_per_sqft'] if p.get('rate_per_sqft') is not None else '—'} | "
                  f"{p['priced']}/{p['items']} |")

    priced = [i for i in state["items"] if i["status"] == "priced"]
    md.append(f"\n## Priced items ({len(priced)})\n")
    md.append("| Sheet | Row | Item | Qty | L1 vendor | L1 rate | Final rate | Amount | Candidates |")
    md.append("|---|---|---|---|---|---|---|---|---|")
    for i in priced:
        cands = "; ".join(
            f"{c['vendor']}@{c['total_rate']}"
            + (f" ⚠{c['history_flag']['deviation_pct']}% off history" if c.get("history_flag") else "")
            for c in i["candidates"])
        md.append(f"| {i['sheet']} | {i['row']} | {i['description'][:60]} | {i['qty']} | "
                  f"{i['l1_vendor']} | {i['l1_rate']} | {i['final_rate']} | {i['amount']} | {cands} |")

    md.append(f"\n## Review queue ({len(state['review_queue'])}) — needs a human\n")
    for r in state["review_queue"]:
        md.append(f"- **{r.get('vendor', '?')}** {r['quote']['description'][:80]} → "
                  f"{r['boq']['sheet']} r{r['boq']['row']} "
                  f"(confidence {r['confidence']}, signal {r['signal']})")
    if not state["review_queue"]:
        md.append("- (empty)")

    md.append(f"\n## Unmatched quote lines ({len(state['unmatched'])})\n")
    for u in state["unmatched"]:
        md.append(f"- {u.get('vendor', '?')}: {u['description'][:80]} — {u['reason']}")
    if not state["unmatched"]:
        md.append("- (none)")

    if state.get("extraction_errors"):
        md.append(f"\n## Extraction validation errors ({len(state['extraction_errors'])})\n")
        for e in state["extraction_errors"][:50]:
            md.append(f"- {e}")

    unpriced = [i for i in state["items"] if i["status"] == "unpriced"]
    md.append(f"\n## Unpriced BOQ items: {len(unpriced)} "
              f"(no vendor coverage — float or in-house rate needed)\n")

    text = "\n".join(md)
    md_path = run_dir / "audit-report.md"
    md_path.write_text(text, encoding="utf-8")

    body = html.escape(text)
    (run_dir / "audit-report.html").write_text(
        "<!doctype html><meta charset='utf-8'><title>BOQ audit</title>"
        "<style>body{font-family:system-ui;max-width:1100px;margin:2rem auto;"
        "padding:0 1rem;line-height:1.45}pre{white-space:pre-wrap}</style>"
        f"<pre>{body}</pre>"
    , encoding="utf-8")
    return md_path
