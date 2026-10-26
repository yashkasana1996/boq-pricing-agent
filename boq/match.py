"""Spec matching — THE CRUX (CLAUDE.md §5).

Deterministic pass first, in strict signal order:
  1. embedded cross-ref  — vendor item carries "(BOQ Sl No-7)" (hazard #9)
  2. exact code          — vendor echoes the BOQ's item code (ESD-01 style)
  3. fingerprint         — normalized token overlap + unit agreement

Then an optional LLM pass for leftovers (returns candidate + confidence; every
candidate is still threshold-gated in code). Wrong matches silently corrupt
the offer, so: accept >= 0.80, review-queue 0.50-0.80, drop < 0.50 — the agent
never guesses silently.
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from pathlib import Path

from .extract import extract_workbook
from .ingest import load_workbook
from .knowledge import (
    DEFAULT_SYNONYMS,
    fingerprint_tokens,
    normalize_unit,
)

ACCEPT = 0.80
QUEUE = 0.50


def boq_items(payload: dict) -> list[dict]:
    out = []
    for s in payload["sheets"]:
        if s["kind"] in ("boq", "make_list"):
            for it in s["items"]:
                out.append(it)
    return out


def quote_items(payload: dict) -> list[dict]:
    out = []
    for s in payload["sheets"]:
        if s["kind"] in ("boq", "price_list", "make_list"):
            out.extend(s["items"])
    return out


@dataclass
class Match:
    boq_item: dict
    quote_item: dict
    confidence: float
    signal: str  # slno_ref | code | fingerprint | llm

    def to_dict(self) -> dict:
        return {
            "boq": {"sheet": self.boq_item["sheet"], "row": self.boq_item["row"],
                    "sr": self.boq_item.get("sr"), "code": self.boq_item.get("code"),
                    "description": (self.boq_item.get("description") or "")[:120]},
            "quote": {"sheet": self.quote_item["sheet"], "row": self.quote_item["row"],
                      "code": self.quote_item.get("code"),
                      "description": (self.quote_item.get("description") or "")[:120],
                      "total_rate": self.quote_item.get("total_rate")},
            "confidence": self.confidence,
            "signal": self.signal,
        }


@dataclass
class MatchResult:
    accepted: list[Match] = field(default_factory=list)
    review_queue: list[Match] = field(default_factory=list)
    unmatched_quote_items: list[dict] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "accepted": [m.to_dict() for m in self.accepted],
            "review_queue": [m.to_dict() for m in self.review_queue],
            "unmatched_quote_items": [
                {"sheet": q["sheet"], "row": q["row"],
                 "description": (q.get("description") or "")[:120],
                 "reason": q.get("_unmatched_reason", "no candidate above drop threshold")}
                for q in self.unmatched_quote_items
            ],
            "counts": {"accepted": len(self.accepted), "review": len(self.review_queue),
                       "unmatched": len(self.unmatched_quote_items)},
        }


def _norm_sr(v: object) -> str | None:
    if v is None:
        return None
    s = str(v).strip()
    try:
        f = float(s)
        return f"{f:g}"
    except ValueError:
        return s or None


def _fingerprint_score(b: dict, q: dict, synonyms: dict[str, str]) -> float:
    """Spec-aware similarity, shared with the BOQ-centric matcher so the CLI
    `boq match` and the pricing pipeline agree. Dimension conflicts and
    fragment-only overlaps are gated out here too."""
    from .specmatch import spec_match

    return spec_match(b, q, synonyms)[0]


def match_items(boq: list[dict], quotes: list[dict],
                synonyms: dict[str, str] | None = None,
                use_llm: bool = False) -> MatchResult:
    synonyms = synonyms or DEFAULT_SYNONYMS
    res = MatchResult()
    by_sr: dict[str, list[dict]] = {}
    by_code: dict[str, list[dict]] = {}
    for b in boq:
        sr = _norm_sr(b.get("sr"))
        if sr:
            by_sr.setdefault(sr, []).append(b)
        code = (b.get("code") or "").strip().lower()
        if code:
            by_code.setdefault(code, []).append(b)

    leftovers: list[dict] = []
    for q in quotes:
        # 1. embedded "(BOQ Sl No-x)" cross reference. Serial numbers collide
        # across sheets (every trade sheet restarts at 1), so a ref match is
        # only trusted when the descriptions also plausibly agree — a blind
        # sr match is exactly the "silent wrong rate" failure mode.
        ref = _norm_sr(q.get("boq_slno_ref"))
        if ref and ref in by_sr:
            cands = by_sr[ref]
            best = cands[0] if len(cands) == 1 else max(
                cands, key=lambda b: _fingerprint_score(b, q, synonyms))
            fp = _fingerprint_score(best, q, synonyms)
            if fp >= 0.15:
                conf = 0.95 if len(cands) == 1 else 0.85
                res.accepted.append(Match(best, q, conf, "slno_ref" if len(cands) == 1 else "slno_ref+fingerprint"))
            else:
                res.review_queue.append(Match(best, q, 0.55, "slno_ref-unverified"))
            continue
        # 2. exact code echo
        code = (q.get("code") or "").strip().lower()
        if code and code in by_code:
            cands = by_code[code]
            best = cands[0] if len(cands) == 1 else max(
                cands, key=lambda b: _fingerprint_score(b, q, synonyms))
            res.accepted.append(Match(best, q, 0.92, "code"))
            continue
        leftovers.append(q)

    # 3. fingerprint pass over leftovers
    llm_candidates: list[dict] = []
    for q in leftovers:
        scored = sorted(((_fingerprint_score(b, q, synonyms), b) for b in boq),
                        key=lambda t: t[0], reverse=True)
        score, best = scored[0] if scored else (0.0, None)
        runner = scored[1][0] if len(scored) > 1 else 0.0
        if best is not None and score > 0 and score - runner < 0.05:
            # two near-identical BOQ candidates: a coin-flip is exactly the
            # silent-wrong-rate failure — force human review, never accept
            score = min(max(0.0, score - 0.10), ACCEPT - 0.01)
        if best is None or score < QUEUE:
            llm_candidates.append(q)
        elif score >= ACCEPT:
            res.accepted.append(Match(best, q, round(score, 3), "fingerprint"))
        else:
            res.review_queue.append(Match(best, q, round(score, 3), "fingerprint"))

    # 4. optional LLM pass for what deterministic matching couldn't place
    if use_llm and llm_candidates and os.environ.get("ANTHROPIC_API_KEY"):
        from .match_llm import llm_match

        for q, m in llm_match(boq, llm_candidates):
            if m is None:
                q["_unmatched_reason"] = "LLM found no plausible BOQ item"
                res.unmatched_quote_items.append(q)
            elif m.confidence >= ACCEPT:
                res.accepted.append(m)
            elif m.confidence >= QUEUE:
                res.review_queue.append(m)
            else:
                q["_unmatched_reason"] = f"LLM confidence {m.confidence:.2f} below queue threshold"
                res.unmatched_quote_items.append(q)
    else:
        for q in llm_candidates:
            q["_unmatched_reason"] = "no deterministic candidate above 0.50"
            res.unmatched_quote_items.append(q)
    return res


def match_files(boq_file: str, quote_files: list[str], kb_dir: str = "knowledge",
                use_llm: bool = False) -> MatchResult:
    from .knowledge import KnowledgeBase

    kb = KnowledgeBase.load(kb_dir)
    bp = extract_workbook(load_workbook(boq_file))
    combined = MatchResult()
    b_items = boq_items(bp)
    for qf in quote_files:
        qp = extract_workbook(load_workbook(qf))
        r = match_items(b_items, quote_items(qp), synonyms=kb.spec_synonyms, use_llm=use_llm)
        for m in r.accepted:
            m.quote_item["_vendor"] = Path(qf).stem
        combined.accepted += r.accepted
        combined.review_queue += r.review_queue
        combined.unmatched_quote_items += r.unmatched_quote_items
    return combined


def resolve_review(kb_dir: str, variant: str, canonical: str) -> None:
    """Human resolved a queued match: record the equivalence so the
    deterministic matcher gets it right next time (feedback loop)."""
    from .knowledge import KnowledgeBase

    kb = KnowledgeBase.load(kb_dir)
    kb.learn_synonym(variant, canonical)
    kb.save()
