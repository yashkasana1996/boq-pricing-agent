"""End-to-end fill pipeline: extract -> match -> price -> write-back -> audit.

Every run emits an audit trail (CLAUDE.md non-negotiable): per item the
candidate rates, chosen L1 vendor, base, markup, final rate, amount, match
confidence; unmatched/queued items with reasons. The run directory is the
single artifact reviewers need.
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from .extract import extract_workbook, validate_extraction
from .ingest import load_workbook
from .knowledge import KnowledgeBase, item_fingerprint
from .match import boq_items, quote_items
from .price import CandidateRate, PricedItem, PricingRun
from .writeback import write_client_copy, write_internal_workbook


def fill_pipeline(boq_file: str, quote_files: list[str], markup_pct: float,
                  gst_mode: str = "exclusive", gst_pct: float = 18.0,
                  area_sqft: float | None = None, round_digits: int = 2,
                  out_dir: str = "runs", kb_dir: str = "knowledge",
                  use_llm: bool = False) -> dict:
    run_dir = Path(out_dir) / datetime.now().strftime("run-%Y%m%d-%H%M%S")
    run_dir.mkdir(parents=True, exist_ok=True)
    kb = KnowledgeBase.load(kb_dir)

    # 1. extract
    boq_wb = load_workbook(boq_file)
    bp = extract_workbook(boq_wb)
    extraction_errors = validate_extraction(bp)
    b_items = boq_items(bp)

    # header column maps per sheet — write-back lands values only here
    rate_cols = {s["sheet"]: s["columns"] for s in bp["sheets"] if s["kind"] == "boq"}

    # analyze the BOQ up front (formats differ per client) so the run records
    # exactly what structure was detected before any pricing happens
    from .specmatch import analyze_boq, match_boq_item
    from .docs_extract import load_quote_payload

    boq_analysis = analyze_boq(bp)

    # 2. load every quote (xlsx/csv/pdf/.msg all normalised upstream), pooling
    # all vendor lines; a single bad file becomes a note, never a dead run
    file_notes: list[str] = []
    quote_lines: list[tuple[str, dict]] = []   # (vendor, quote_item)
    for qf in quote_files:
        qp, notes = load_quote_payload(qf, use_llm=use_llm)
        file_notes += notes
        if qp is None:
            continue
        vendor = Path(qf).stem
        for q in quote_items(qp):
            if q.get("total_rate") not in (None, 0.0):
                quote_lines.append((vendor, q))

    # 3. BOQ-centric matching + pricing: for each BOQ item find matching vendor
    # lines by description AND spec, take L1, else review (uncertain) / unpriced
    # (no candidate). Review is the exception, never the default.
    from .specmatch import ACCEPT as MATCH_ACCEPT

    review, unmatched = [], []
    priced_items: list[PricedItem] = []
    for b in b_items:
        m = match_boq_item(b, quote_lines, kb.spec_synonyms)
        fp = item_fingerprint(b.get("description") or "", b.get("unit"))
        # only CONFIDENT matches are eligible for L1 selection (a cheap but
        # low-confidence line must never undercut the real rate)
        confident = [
            CandidateRate(
                vendor=c.vendor, total_rate=c.total_rate,
                supply_rate=c.quote.get("supply_rate"), install_rate=c.quote.get("install_rate"),
                confidence=c.score, signal=c.why,
                history_flag=kb.sanity_check(fp, c.total_rate) if c.total_rate else None,
            )
            for c in m.candidates if c.score >= MATCH_ACCEPT
        ]
        pi = PricedItem(
            sheet=b["sheet"], row=b["row"], description=b.get("description") or "",
            unit=b.get("unit"), qty=b.get("qty"), package=b["sheet"].strip(),
            candidates=confident if m.status == "priced" else [],
        )
        if m.status == "review":
            pi.status = "queued"
            best = m.candidates[0]
            review.append({
                "boq": {"sheet": b["sheet"], "row": b["row"], "sr": b.get("sr"),
                        "code": b.get("code"), "description": (b.get("description") or "")[:140]},
                "quote": {"vendor": best.vendor, "code": best.quote.get("code"),
                          "description": (best.quote.get("description") or "")[:140],
                          "total_rate": best.total_rate},
                "vendor": best.vendor, "confidence": best.score, "signal": best.why,
            })
        elif m.status == "unpriced":
            pi.status = "unpriced"
        priced_items.append(pi)

    run = PricingRun(items=priced_items, markup_pct=markup_pct, gst_mode=gst_mode,
                     gst_pct=gst_pct, area_sqft=area_sqft, round_digits=round_digits)
    run.price_all()

    # 4. write-back
    client = write_client_copy(boq_file, run.items, run_dir / "client-copy.xlsx", rate_cols)
    internal = write_internal_workbook(run, run_dir / "internal-workbook.xlsx")

    # 5. audit state
    state = {
        "boq_file": str(boq_file), "quote_files": [str(q) for q in quote_files],
        "markup_pct": markup_pct, "gst_mode": gst_mode, "gst_pct": gst_pct,
        "area_sqft": area_sqft, "round_digits": round_digits,
        "extraction_errors": extraction_errors,
        "file_notes": file_notes,
        "boq_analysis": boq_analysis,
        "items": [it.to_dict() for it in run.items],
        "review_queue": review, "unmatched": unmatched,
        "package_totals": run.package_totals(), "totals": run.totals(),
        "outputs": {"client_copy": client, "internal_workbook": internal},
    }
    (run_dir / "state.json").write_text(json.dumps(state, indent=1, default=str), encoding="utf-8")
    summary = {
        "run_dir": str(run_dir), "totals": state["totals"],
        "items_priced": state["totals"]["items_priced"],
        "review_queue": len(review), "unmatched": len(unmatched),
        "outputs": state["outputs"],
    }
    return summary


def load_priced_state(state_file: str) -> PricingRun:
    """Rebuild a PricingRun from a fill run's state.json (for `boq target`)."""
    data = json.loads(Path(state_file).read_text(encoding="utf-8"))
    items = []
    for d in data["items"]:
        pi = PricedItem(
            sheet=d["sheet"], row=d["row"], description=d["description"],
            unit=d.get("unit"), qty=d.get("qty"), package=d["package"],
            candidates=[CandidateRate(vendor=c["vendor"], total_rate=c["total_rate"],
                                      confidence=c.get("confidence", 1.0),
                                      signal=c.get("signal", ""))
                        for c in d.get("candidates", [])],
        )
        items.append(pi)
    run = PricingRun(items=items, markup_pct=data["markup_pct"],
                     gst_mode=data["gst_mode"], gst_pct=data["gst_pct"],
                     area_sqft=data.get("area_sqft"), round_digits=data.get("round_digits", 2))
    run.price_all()
    return run
