"""Spec-aware, BOQ-centric matching and pricing.

Encodes the estimator's actual process (verbatim, from the user):
  for each BOQ item, taken exactly as written with its specification:
    - find the vendor-quote line(s) that match by description AND spec
      (size, height, thickness, make, grade, unit)
    - among genuine matches take the LOWEST total rate (L1)
    - final rate = ROUND(L1 x (1 + markup%)); amount = qty x final rate
  sum amounts -> total. The client BOQ format is never changed.

The decisive lesson from real data: codes and serials are reused across
vendors and revisions (a vendor's "RF-02 @ 4738, 450mm anti-static" is NOT the
BOQ's "RF-02 @ 3922, 150mm laminate"). So a code/serial match is only trusted
when the DIMENSION/SPEC agrees. A conflicting dimension (150mm vs 450mm) blocks
the match outright — that is precisely how a silent wrong rate was getting in.

Review queue is the exception, not the rule: an item with a confident match is
priced automatically; an item with only a weak/ambiguous candidate goes to
review; an item with no candidate at all is simply 'unpriced' (needs a float or
in-house rate) — it is NOT dumped into the review queue.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .knowledge import DEFAULT_SYNONYMS, fingerprint_tokens, normalize_unit

# confidence tiers
ACCEPT = 0.72   # >= this and unambiguous -> auto-priced
QUEUE = 0.50    # [QUEUE, ACCEPT) -> human review; < QUEUE with no candidate -> unpriced

# ---------------------------------------------------------------------------
# specification signature — the decisive discriminator
# ---------------------------------------------------------------------------

_SIZE = re.compile(r"(\d{2,4})\s*(?:mm)?\s*[x×*]\s*(\d{2,4})", re.I)
_MM = re.compile(r"(\d{1,4})\s*mm\b", re.I)
_HIGH = re.compile(r"(\d{1,4})\s*mm?\s*(?:high|ht|height|thk|thick)", re.I)
_PLAIN_HIGH = re.compile(r"(\d{2,4})\s*(?:mm)?\s*(?:high|height)", re.I)


def spec_signature(text: str) -> set[str]:
    """Decisive spec tokens: sizes (600x600), heights/thicknesses (150mm)."""
    t = (text or "").lower()
    sig: set[str] = set()
    for a, b in _SIZE.findall(t):
        lo, hi = sorted((int(a), int(b)))
        sig.add(f"{lo}x{hi}")
    # remove size spans so their component numbers aren't re-counted as heights
    t_no_size = _SIZE.sub(" ", t)
    for m in _MM.findall(t_no_size):
        sig.add(f"{int(m)}mm")
    for m in _HIGH.findall(t_no_size):
        sig.add(f"{int(m)}mm")
    for m in _PLAIN_HIGH.findall(t_no_size):
        sig.add(f"{int(m)}mm")
    return sig


# material / finish qualifiers that materially change the product. Stored as
# stems so laminate/laminated/lamination collapse together. If a BOQ item names
# a qualifier from a group and the quote names a DIFFERENT one from that group
# (or none), the product differs (bare vs laminate vs anti-static flooring).
QUALIFIER_GROUPS = [
    {"bare", "laminat", "antistat", "esd", "vinyl", "ceramic", "vitrif", "granit",
     "epoxy", "wooden", "carpet", "rubber", "kota", "marble"},          # flooring/finish
    {"gypsum", "calcium", "mineral", "metal", "wooden", "acoustic",
     "perforat", "grid", "baffle", "stretch"},                          # ceiling
    {"glass", "gypsum", "block", "brick", "drywall", "modular"},        # partition
    {"stainless", "alumin", "brass", "galvan", "powder", "anodis"},     # metal/finish
    {"firerat", "fireresist", "acoustic", "smoke"},                     # rating
]
_ALL_QUALS = sorted({q for g in QUALIFIER_GROUPS for q in g}, key=len, reverse=True)


def _norm_quals(text: str) -> str:
    t = re.sub(r"anti[\s-]*static", "antistat", (text or "").lower())
    t = re.sub(r"fire[\s-]*rat", "firerat", t)
    t = re.sub(r"fire[\s-]*resist", "fireresist", t)
    return t


def qualifier_set(text: str) -> set[str]:
    t = _norm_quals(text)
    return {q for q in _ALL_QUALS if q in t}


def qualifier_conflict(boq_text: str, quote_text: str) -> tuple[int, int]:
    """Returns (missing, shared): how many BOQ qualifiers the quote lacks vs
    shares, considering only groups the BOQ item actually invokes."""
    bq, qq = qualifier_set(boq_text), qualifier_set(quote_text)
    if not bq:
        return 0, 0
    shared = bq & qq
    return len(bq - qq), len(shared)


def _both_text(it: dict) -> str:
    return (it.get("description") or "") + " " + (it.get("spec") or "")


def spec_match(boq: dict, quote: dict, synonyms: dict[str, str]) -> tuple[float, str]:
    """Score how well a quote line satisfies a BOQ item. Returns (score, why).

    score in [0,1]; a dimension conflict caps the score below ACCEPT so it can
    never be auto-priced.
    """
    bt = fingerprint_tokens(_both_text(boq), synonyms)
    qt = fingerprint_tokens(_both_text(quote), synonyms)
    if not bt or not qt:
        return 0.0, "no comparable text"
    inter = len(bt & qt)
    if inter == 0:
        return 0.0, "no shared words"
    containment = inter / min(len(bt), len(qt))   # how fully one covers the other
    jaccard = inter / len(bt | qt)
    score = 0.60 * containment + 0.40 * jaccard
    why = [f"words {int(containment*100)}%"]
    ceiling = 1.0   # hard safety ceilings applied at the very end

    # unit agreement
    bu, qu = normalize_unit(boq.get("unit")), normalize_unit(quote.get("unit"))
    if bu and qu:
        if bu == qu:
            score += 0.08
        else:
            score -= 0.18
            why.append(f"unit {bu}≠{qu}")

    # DIMENSION / SPEC bonus, with a conflict ceiling
    bsig, qsig = spec_signature(_both_text(boq)), spec_signature(_both_text(quote))
    if bsig:
        shared = bsig & qsig
        if not shared:
            ceiling = min(ceiling, ACCEPT - 0.15)   # BOQ demands a size the quote lacks
            why.append(f"spec conflict {sorted(bsig)}≠{sorted(qsig) or '∅'}")
        else:
            cover = len(shared) / len(bsig)
            score += 0.18 * cover
            if cover < 1.0:
                ceiling = min(ceiling, ACCEPT + 0.10)
            why.append(f"spec {sorted(shared)}")

    # material/finish qualifier agreement (bare vs laminate vs anti-static, etc.)
    missing, shared_q = qualifier_conflict(_both_text(boq), _both_text(quote))
    if missing:
        ceiling = min(ceiling, ACCEPT - 0.05)
        score -= 0.10 * missing
        why.append(f"finish differs (−{missing})")
    elif shared_q:
        score += 0.06
        why.append("finish matches")

    # exact item-code agreement corroborates (never overrides a conflict ceiling)
    bc = (boq.get("code") or "").strip().lower()
    qc = (quote.get("code") or "").strip().lower()
    if bc and qc and bc == qc and (not bsig or (bsig & qsig)) and not missing:
        score += 0.12
        why.append(f"code {bc}")

    # HARD GATE (applied last so no later bonus can lift a weak match over the
    # line): a confident match needs >= 2 shared CONTENT words, not just a
    # shared number — stops a fragment like "Size:- 600 mm" matching anything.
    shared_words = {t for t in (bt & qt) if any(c.isalpha() for c in t)}
    if len(shared_words) < 2:
        ceiling = min(ceiling, ACCEPT - 0.12)
        why.append(f"only {len(shared_words)} shared word(s)")

    return max(0.0, min(min(1.0, ceiling), score)), "; ".join(why)


# ---------------------------------------------------------------------------
# BOQ-centric pricing
# ---------------------------------------------------------------------------


@dataclass
class Candidate:
    vendor: str
    quote: dict
    score: float
    why: str

    @property
    def total_rate(self):
        return self.quote.get("total_rate")


@dataclass
class ItemMatch:
    boq: dict
    status: str                       # priced | review | unpriced
    l1: Candidate | None = None
    candidates: list[Candidate] = field(default_factory=list)   # all considered, best-first
    reason: str = ""


def match_boq_item(boq: dict, quote_lines: list[tuple[str, dict]],
                   synonyms: dict[str, str]) -> ItemMatch:
    """quote_lines: list of (vendor, quote_item). Returns an ItemMatch."""
    scored: list[Candidate] = []
    for vendor, q in quote_lines:
        if q.get("total_rate") in (None, 0.0):
            continue
        s, why = spec_match(boq, q, synonyms)
        if s >= QUEUE:
            scored.append(Candidate(vendor, q, round(s, 3), why))
    scored.sort(key=lambda c: c.score, reverse=True)

    confident = [c for c in scored if c.score >= ACCEPT]
    if confident:
        # L1: lowest total rate among the confident matches (across all vendors)
        l1 = min(confident, key=lambda c: c.total_rate)
        return ItemMatch(boq, "priced", l1=l1, candidates=scored,
                         reason=f"L1 of {len(confident)} confident match(es)")
    if scored:
        return ItemMatch(boq, "review", candidates=scored,
                         reason=f"best candidate only {scored[0].score:.2f} confident")
    return ItemMatch(boq, "unpriced", reason="no vendor quote matched this item")


def analyze_boq(payload: dict) -> dict:
    """Understand the uploaded BOQ before pricing (different clients, different
    formats). Reports per-sheet detected structure so the user can see the
    console read their file correctly."""
    sheets = []
    total_items = 0
    priced_capable = 0
    for s in payload["sheets"]:
        cols = s.get("columns", {})
        n = len(s["items"])
        total_items += n
        has_rate_col = any(k in cols for k in ("rate", "supply_rate", "amount", "supply_amount"))
        has_qty = "qty" in cols or "total_qty" in cols
        if s["kind"] == "boq" and has_rate_col:
            priced_capable += n
        sheets.append({
            "sheet": s["sheet"], "kind": s["kind"], "items": n,
            "columns_detected": sorted(cols.keys()),
            "has_qty": has_qty, "has_rate_column": has_rate_col,
            "supply_install_split": s.get("supply_install_split", False),
        })
    boq_sheets = [s for s in sheets if s["kind"] == "boq"]
    kinds = sorted({s["kind"] for s in sheets})
    return {
        "boq_type": ("item-rate (priceable)" if priced_capable else "specification/lumpsum"),
        "total_sheets": len(sheets),
        "priceable_sheets": len(boq_sheets),
        "total_items": total_items,
        "priceable_items": priced_capable,
        "sheet_kinds": kinds,
        "sheets": sheets,
    }
