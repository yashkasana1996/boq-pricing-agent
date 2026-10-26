"""The learning layer — a knowledge base built from manually-prepared tenders.

Not model training: plain inspectable files (CLAUDE.md §learning):
  rate_history.jsonl   item fingerprint -> vendor -> rate -> date -> project
  spec_synonyms.json   learned text equivalences, grows with review resolutions
  conventions.yaml     observed house style (rounding, GST mode, summary shape)

(BUILD_PLAN names rate_history.parquet; JSONL is used instead so the KB stays
dependency-free and diffable in git — swap to parquet later if volume demands.)

M2 acceptance: reconstruct >=90% of a final offer's amounts from its own rates
(Amount == round-consistent Qty x Rate), and recover the implied markup for
every BOQ item a vendor quote can be matched to.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from .extract import extract_workbook, norm
from .ingest import load_workbook

STOPWORDS = {
    "providing", "and", "fixing", "laying", "supply", "supplying", "installation",
    "installing", "of", "the", "in", "to", "as", "per", "with", "for", "at", "all",
    "approved", "complete", "including", "etc", "a", "an", "or", "shall", "be",
    "work", "works", "rate", "cost", "position", "specification", "specifications",
    "drawing", "drawings", "required", "necessary",
    # generic dimension/qualifier words carry no discriminating meaning on their
    # own — they must not inflate a match (a "600 mm" joinery item is not a
    # "600 mm" floor tile)
    "size", "high", "height", "deep", "depth", "wide", "width", "thick",
    "thickness", "length", "long", "dia", "diameter", "nos", "make", "makes",
    "colour", "color", "shade", "finish", "finished", "provide", "providing",
}

DEFAULT_SYNONYMS = {
    # seeded from fixtures; grows via review-queue resolutions
    "gyp. bd.": "gypsum board",
    "gyp bd": "gypsum board",
    "bd.": "board",
    "alu": "aluminium",
    "aluminum": "aluminium",
    "sst": "stainless steel",
    "ss": "stainless steel",
    "vit.": "vitrified",
    "flr": "floor",
    "clg": "ceiling",
    "skirting profile": "skirting",
    "pt.": "paint",
    "laminated": "laminate",
    "lamination": "laminate",
    "anti static": "antistatic",
    "anti-static": "antistatic",
    "anodised": "anodized",
    "aluminum": "aluminium",
    "colour": "color",
}

UNIT_ALIASES = {
    "sqm": "sqm", "sqmt": "sqm", "sq.m": "sqm", "sq.mt": "sqm", "sq m": "sqm", "m2": "sqm",
    "sqft": "sqft", "sft": "sqft", "sq.ft": "sqft", "sq ft": "sqft",
    "rmt": "rmt", "rm": "rmt", "mtr": "rmt", "m": "rmt", "meter": "rmt", "metre": "rmt",
    "nos": "nos", "no": "nos", "no.": "nos", "num": "nos", "each": "nos", "ea": "nos",
    "set": "set", "sets": "set", "pair": "pair", "lot": "lot", "ls": "lot", "lumpsum": "lot",
    "kg": "kg", "kgs": "kg", "days": "days", "day": "days",
}


def normalize_unit(u: str | None) -> str | None:
    if not u:
        return None
    return UNIT_ALIASES.get(norm(u).rstrip("s. "), norm(u).rstrip(". "))


def apply_synonyms(text: str, synonyms: dict[str, str]) -> str:
    t = norm(text)
    for k in sorted(synonyms, key=len, reverse=True):
        t = t.replace(k, synonyms[k])
    return t


def fingerprint_tokens(description: str, synonyms: dict[str, str] | None = None) -> set[str]:
    """Order-free token fingerprint of an item description (matching signal)."""
    t = apply_synonyms(description, synonyms or DEFAULT_SYNONYMS)
    tokens = re.findall(r"[a-z]{3,}|\d+(?:\.\d+)?(?:mm|m|kg|w|va|a)?", t)
    return {tok for tok in tokens if tok not in STOPWORDS}


def item_fingerprint(description: str, unit: str | None, synonyms=None) -> str:
    toks = sorted(fingerprint_tokens(description, synonyms))[:24]
    u = normalize_unit(unit) or "?"
    return f"{u}|{' '.join(toks)}"


@dataclass
class KnowledgeBase:
    kb_dir: Path
    rate_history: list[dict] = field(default_factory=list)
    spec_synonyms: dict[str, str] = field(default_factory=lambda: dict(DEFAULT_SYNONYMS))
    conventions: dict = field(default_factory=dict)

    @classmethod
    def load(cls, kb_dir: str | Path) -> "KnowledgeBase":
        kb_dir = Path(kb_dir)
        kb = cls(kb_dir=kb_dir)
        rh = kb_dir / "rate_history.jsonl"
        if rh.exists():
            kb.rate_history = [json.loads(l) for l in rh.read_text(encoding="utf-8").splitlines() if l.strip()]
        syn = kb_dir / "spec_synonyms.json"
        if syn.exists():
            kb.spec_synonyms.update(json.loads(syn.read_text(encoding="utf-8")))
        conv = kb_dir / "conventions.yaml"
        if conv.exists():
            kb.conventions = _read_simple_yaml(conv.read_text(encoding="utf-8"))
        return kb

    def save(self) -> None:
        self.kb_dir.mkdir(parents=True, exist_ok=True)
        (self.kb_dir / "rate_history.jsonl").write_text(
            "\n".join(json.dumps(r, default=str) for r in self.rate_history)
        , encoding="utf-8")
        (self.kb_dir / "spec_synonyms.json").write_text(json.dumps(self.spec_synonyms, indent=1), encoding="utf-8")
        (self.kb_dir / "conventions.yaml").write_text(_write_simple_yaml(self.conventions), encoding="utf-8")

    def add_rate(self, fingerprint: str, vendor: str, rate: float, project: str,
                 unit: str | None = None, when: str | None = None, kind: str = "quote") -> None:
        self.rate_history.append({
            "fingerprint": fingerprint, "vendor": vendor, "rate": rate,
            "unit": unit, "project": project, "date": when or date.today().isoformat(),
            "kind": kind,
        })

    def rates_for(self, fingerprint: str) -> list[dict]:
        return [r for r in self.rate_history if r["fingerprint"] == fingerprint]

    def sanity_check(self, fingerprint: str, rate: float, tolerance_pct: float = 40.0) -> dict | None:
        """Flag a rate that deviates >tolerance from history median. None = fine."""
        hist = [r["rate"] for r in self.rates_for(fingerprint) if r["rate"]]
        if not hist:
            return None
        hist.sort()
        median = hist[len(hist) // 2]
        if median and abs(rate - median) / median * 100 > tolerance_pct:
            return {"median": median, "rate": rate,
                    "deviation_pct": round(abs(rate - median) / median * 100, 1)}
        return None

    def learn_synonym(self, variant: str, canonical: str) -> None:
        """Called when a human resolves a review-queue match (M3 feedback loop)."""
        self.spec_synonyms[norm(variant)] = norm(canonical)

    def summary(self) -> dict:
        return {
            "kb_dir": str(self.kb_dir),
            "rate_history_entries": len(self.rate_history),
            "synonyms": len(self.spec_synonyms),
            "conventions": self.conventions,
        }


def _read_simple_yaml(text: str) -> dict:
    """Flat key: value YAML subset (no external dependency)."""
    out = {}
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#") or ":" not in line:
            continue
        k, v = line.split(":", 1)
        v = v.strip()
        if re.fullmatch(r"-?\d+", v):
            v = int(v)
        elif re.fullmatch(r"-?\d+\.\d+", v):
            v = float(v)
        elif v.lower() in ("true", "false"):
            v = v.lower() == "true"
        out[k.strip()] = v
    return out


def _write_simple_yaml(d: dict) -> str:
    lines = ["# observed house conventions (auto-learned; edit freely)"]
    for k, v in d.items():
        lines.append(f"{k}: {v}")
    return "\n".join(lines) + "\n"


# ---------------------------------------------------------------------------
# learning from a final offer
# ---------------------------------------------------------------------------


def detect_round_digits(rates: list[float]) -> int:
    """Infer the house rounding from observed final rates."""
    if not rates:
        return 2
    if all(abs(r - round(r)) < 1e-9 for r in rates):
        return 0
    if all(abs(r * 100 - round(r * 100)) < 1e-6 for r in rates):
        return 2
    return 2


def reconstruct_offer(offer_payload: dict) -> dict:
    """M2 acceptance core: for every priced item, check Amount == Qty x Rate
    under the detected rounding convention. Returns stats + failures."""
    ok, fail, priced = 0, [], 0
    rates = []
    for s in offer_payload["sheets"]:
        if s["kind"] != "boq":
            continue
        for it in s["items"]:
            q, tr, amt = it.get("qty"), it.get("total_rate"), it.get("amount")
            if q in (None, 0.0) or tr in (None, 0.0) or amt in (None, 0.0):
                continue
            priced += 1
            rates.append(tr)
            recon = q * tr
            if abs(recon - amt) <= max(1.0, 0.005 * abs(amt)):
                ok += 1
            else:
                fail.append({"sheet": s["sheet"], "row": it["row"],
                             "expected": amt, "reconstructed": round(recon, 2)})
    return {
        "priced_items": priced,
        "reconstructed_ok": ok,
        "reconstruction_rate": round(ok / priced, 4) if priced else None,
        "detected_round_digits": detect_round_digits(rates),
        "failures": fail[:50],
    }


def implied_markups(offer_payload: dict, quote_payloads: dict[str, dict],
                    synonyms=None) -> list[dict]:
    """Where a vendor quote matches a final-offer item (by code, embedded BOQ
    ref, or fingerprint), compute implied markup = final/vendor - 1."""
    from .match import match_items, quote_items, boq_items  # local import: no cycle at module load

    offers = boq_items(offer_payload)
    out = []
    for vendor, qp in quote_payloads.items():
        qitems = quote_items(qp)
        res = match_items(offers, qitems, synonyms=synonyms)
        for m in res.accepted:
            b, q = m.boq_item, m.quote_item
            if b.get("total_rate") and q.get("total_rate"):
                markup = b["total_rate"] / q["total_rate"] - 1.0
                out.append({
                    "vendor": vendor,
                    "sheet": b["sheet"], "row": b["row"],
                    "final_rate": b["total_rate"], "vendor_rate": q["total_rate"],
                    "implied_markup_pct": round(markup * 100, 2),
                    "confidence": m.confidence, "signal": m.signal,
                })
    return out


def learn_from_offer(offer_file: str | Path, quote_files: list[str | Path],
                     kb_dir: str | Path = "knowledge", project: str | None = None) -> KnowledgeBase:
    """Seed/extend the KB from a submitted offer and its vendor quotes."""
    kb = KnowledgeBase.load(kb_dir)
    offer_file = Path(offer_file)
    project = project or offer_file.stem

    offer = extract_workbook(load_workbook(offer_file))
    recon = reconstruct_offer(offer)

    # rate history: every priced final-offer item is house ground truth
    n_rates = 0
    for s in offer["sheets"]:
        if s["kind"] != "boq":
            continue
        for it in s["items"]:
            if it.get("total_rate") and it.get("qty"):
                fp = item_fingerprint(it["description"], it.get("unit"), kb.spec_synonyms)
                kb.add_rate(fp, vendor="__final_offer__", rate=it["total_rate"],
                            project=project, unit=it.get("unit"), kind="final_offer")
                n_rates += 1

    quote_payloads = {}
    for qf in quote_files:
        qf = Path(qf)
        qp = extract_workbook(load_workbook(qf))
        quote_payloads[qf.stem] = qp
        for s in qp["sheets"]:
            for it in s["items"]:
                if it.get("total_rate"):
                    fp = item_fingerprint(it["description"], it.get("unit"), kb.spec_synonyms)
                    kb.add_rate(fp, vendor=qf.stem, rate=it["total_rate"],
                                project=project, unit=it.get("unit"), kind="quote")

    markups = implied_markups(offer, quote_payloads, synonyms=kb.spec_synonyms)

    kb.conventions.update({
        "round_digits": recon["detected_round_digits"],
        "reconstruction_rate": recon["reconstruction_rate"],
        "gst_mode": "exclusive",       # observed: CGST/SGST rows added on total
        "gst_split_cgst_sgst": True,   # observed: split placeholder rows
        "summary_area_cell_hardcoded": True,
        "rate_style": "values_only_in_client_copy",
        "last_project": project,
    })
    if markups:
        med = sorted(m["implied_markup_pct"] for m in markups)[len(markups) // 2]
        kb.conventions["typical_markup_pct"] = med
    kb.save()
    (kb.kb_dir / "last_learn_run.json").write_text(json.dumps({
        "project": project, "reconstruction": recon,
        "rates_added": n_rates, "implied_markups": markups,
    }, indent=1, default=str), encoding="utf-8")
    return kb
