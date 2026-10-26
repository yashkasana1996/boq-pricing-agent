"""boq — command line interface.

Commands mirror BUILD_PLAN.md:
  boq extract <file> [...]           M1  validated JSON items per fixture
  boq learn <offer> [--quotes ...]   M2  seed knowledge base from a final offer
  boq match <boq> <quote> [...]      M3  match quotes to BOQ w/ review queue
  boq fill <boq> [options]           M4  price + write client copy & internal wb
  boq diff <r0> <r1>                 M5  revision diff + carry-forward
  boq target <amount> [options]      M6  back-solve markup for a target total
  boq revise --package X --cut n%    M6  line-level cuts
  boq float <boq> [--make-list f]    M7  per-vendor float packages
  boq report <run-dir>               M8  audit report
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


def _cmd_extract(args) -> int:
    from .extract import extract_workbook, validate_extraction
    from .ingest import cache_extraction, is_junk, load_workbook, load_cached

    rc = 0
    for f in args.files:
        p = Path(f)
        if is_junk(p):
            print(f"skip junk: {p.name}", file=sys.stderr)
            continue
        if args.cache:
            cached = load_cached(args.cache, p)
            if cached is not None:
                print(json.dumps(cached, indent=None if args.compact else 1, default=str))
                continue
        wb = load_workbook(p)
        payload = extract_workbook(wb)
        errors = validate_extraction(payload)
        payload["validation_errors"] = errors
        if errors:
            rc = 2
            print(f"VALIDATION: {len(errors)} error(s) in {p.name}", file=sys.stderr)
        if args.cache:
            cache_extraction(args.cache, wb, payload)
        out = json.dumps(payload, indent=None if args.compact else 1, default=str)
        if args.out:
            Path(args.out).write_text(out, encoding="utf-8")
            print(f"wrote {args.out} ({payload['item_count']} items)", file=sys.stderr)
        else:
            print(out)
    return rc


def _cmd_learn(args) -> int:
    from .knowledge import learn_from_offer

    kb = learn_from_offer(args.offer, quote_files=args.quotes or [], kb_dir=args.kb)
    print(json.dumps(kb.summary(), indent=1))
    return 0


def _cmd_match(args) -> int:
    from .match import match_files

    result = match_files(args.boq, args.quotes, kb_dir=args.kb, use_llm=args.llm)
    print(json.dumps(result.to_dict(), indent=1, default=str))
    return 0


def _cmd_fill(args) -> int:
    from .pipeline import fill_pipeline

    out = fill_pipeline(
        boq_file=args.boq,
        quote_files=args.quotes or [],
        markup_pct=args.markup,
        gst_mode=args.gst,
        gst_pct=args.gst_pct,
        area_sqft=args.area,
        round_digits=args.round_digits,
        out_dir=args.out_dir,
        kb_dir=args.kb,
        use_llm=args.llm,
    )
    print(json.dumps(out, indent=1, default=str))
    return 0


def _cmd_diff(args) -> int:
    from .diff import diff_files

    d = diff_files(args.r0, args.r1)
    print(json.dumps(d.to_dict(), indent=1, default=str))
    return 0


def _cmd_target(args) -> int:
    from .price import back_solve_target
    from .pipeline import load_priced_state

    state = load_priced_state(args.state)
    res = back_solve_target(state, args.amount, gst_mode=args.gst, gst_pct=args.gst_pct,
                            round_digits=args.round_digits, margin_floor_pct=args.margin_floor)
    print(json.dumps(res, indent=1, default=str))
    return 0


def _cmd_float(args) -> int:
    from .float_pkg import build_float_packages

    res = build_float_packages(args.boq, make_list=args.make_list, out_dir=args.out_dir)
    print(json.dumps(res, indent=1, default=str))
    return 0


def _cmd_report(args) -> int:
    from .report import render_report

    p = render_report(args.run_dir)
    print(f"report: {p}")
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="boq", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("extract", help="extract items from xlsx/csv to JSON")
    p.add_argument("files", nargs="+")
    p.add_argument("--out")
    p.add_argument("--cache", help="cache dir keyed by file hash")
    p.add_argument("--compact", action="store_true")
    p.set_defaults(fn=_cmd_extract)

    p = sub.add_parser("learn", help="seed knowledge base from a final offer")
    p.add_argument("offer")
    p.add_argument("--quotes", nargs="*")
    p.add_argument("--kb", default="knowledge")
    p.set_defaults(fn=_cmd_learn)

    p = sub.add_parser("match", help="match vendor quotes to a BOQ")
    p.add_argument("boq")
    p.add_argument("quotes", nargs="+")
    p.add_argument("--kb", default="knowledge")
    p.add_argument("--llm", action="store_true", help="enable LLM assist pass")
    p.set_defaults(fn=_cmd_match)

    p = sub.add_parser("fill", help="price a BOQ and write client + internal workbooks")
    p.add_argument("boq")
    p.add_argument("--quotes", nargs="*")
    p.add_argument("--markup", type=float, required=True, help="markup %% e.g. 12")
    p.add_argument("--gst", choices=["inclusive", "exclusive"], default="exclusive")
    p.add_argument("--gst-pct", type=float, default=18.0)
    p.add_argument("--area", type=float, help="leasable area sqft for ₹/sqft")
    p.add_argument("--round-digits", type=int, default=2)
    p.add_argument("--out-dir", default="runs")
    p.add_argument("--kb", default="knowledge")
    p.add_argument("--llm", action="store_true")
    p.set_defaults(fn=_cmd_fill)

    p = sub.add_parser("diff", help="diff two BOQ revisions")
    p.add_argument("r0")
    p.add_argument("r1")
    p.set_defaults(fn=_cmd_diff)

    p = sub.add_parser("target", help="back-solve markup %% for a target total")
    p.add_argument("amount", type=float)
    p.add_argument("--state", required=True, help="priced state JSON from a fill run")
    p.add_argument("--gst", choices=["inclusive", "exclusive"], default="exclusive")
    p.add_argument("--gst-pct", type=float, default=18.0)
    p.add_argument("--round-digits", type=int, default=2)
    p.add_argument("--margin-floor", type=float, default=8.0)
    p.set_defaults(fn=_cmd_target)

    p = sub.add_parser("float", help="build per-vendor float packages")
    p.add_argument("boq")
    p.add_argument("--make-list")
    p.add_argument("--out-dir", default="floats")
    p.set_defaults(fn=_cmd_float)

    p = sub.add_parser("report", help="render audit report for a run dir")
    p.add_argument("run_dir")
    p.set_defaults(fn=_cmd_report)

    args = ap.parse_args(argv)
    return args.fn(args)


if __name__ == "__main__":
    raise SystemExit(main())
