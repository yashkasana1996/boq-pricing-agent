"""BOQ revision diffing (CLAUDE.md §2).

Classify every item across two revisions as unchanged / qty-changed /
spec-changed / added / removed; carry forward already-priced rates for
unchanged items; report the scope delta in ₹.

Identity resolution order per sheet: code -> serial -> description
fingerprint (revisions renumber serials freely, so sr alone is weak).
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .extract import extract_workbook
from .ingest import load_workbook
from .knowledge import fingerprint_tokens


@dataclass
class DiffEntry:
    status: str  # unchanged | qty_changed | spec_changed | added | removed
    sheet: str
    r0_row: int | None = None
    r1_row: int | None = None
    description: str = ""
    qty_before: float | None = None
    qty_after: float | None = None
    rate_carried: float | None = None
    amount_delta: float | None = None
    detail: str = ""

    def to_dict(self) -> dict:
        return {k: v for k, v in self.__dict__.items() if v is not None and v != ""}


@dataclass
class DiffResult:
    entries: list[DiffEntry] = field(default_factory=list)

    def counts(self) -> dict:
        out: dict[str, int] = {}
        for e in self.entries:
            out[e.status] = out.get(e.status, 0) + 1
        return out

    def scope_delta(self) -> float:
        return round(sum(e.amount_delta or 0.0 for e in self.entries), 2)

    def to_dict(self) -> dict:
        return {
            "counts": self.counts(),
            "scope_delta_inr": self.scope_delta(),
            "entries": [e.to_dict() for e in self.entries],
        }


def _key_items(items: list[dict]) -> tuple[dict, dict, list[dict]]:
    by_code, by_sr, rest = {}, {}, []
    for it in items:
        code = (it.get("code") or "").strip().lower()
        sr = (str(it.get("sr")) if it.get("sr") is not None else "").strip()
        if code:
            by_code.setdefault(code, []).append(it)
        elif sr:
            by_sr.setdefault(sr, []).append(it)
        else:
            rest.append(it)
    return by_code, by_sr, rest


def _fp(it: dict) -> frozenset:
    return frozenset(fingerprint_tokens(
        (it.get("description") or "") + " " + (it.get("spec") or "")))


def _spec_changed(a: dict, b: dict) -> bool:
    fa, fb = _fp(a), _fp(b)
    if not fa or not fb:
        return (a.get("description") or "").strip() != (b.get("description") or "").strip()
    inter = len(fa & fb)
    return inter / max(len(fa), len(fb)) < 0.7


def _qty(it: dict) -> float | None:
    return it.get("qty")


def diff_items(r0_items: list[dict], r1_items: list[dict]) -> DiffResult:
    res = DiffResult()
    used_r0: set[int] = set()

    c0, s0, rest0 = _key_items(r0_items)

    def find_r0(it1: dict) -> dict | None:
        code = (it1.get("code") or "").strip().lower()
        if code and code in c0:
            for cand in c0[code]:
                if id(cand) not in used_r0:
                    return cand
        sr = (str(it1.get("sr")) if it1.get("sr") is not None else "").strip()
        if sr and sr in s0:
            for cand in s0[sr]:
                if id(cand) not in used_r0 and not _spec_changed(cand, it1):
                    return cand
        # fingerprint fallback across the whole r0 sheet
        f1 = _fp(it1)
        if not f1:
            return None
        best, best_score = None, 0.0
        for cand in r0_items:
            if id(cand) in used_r0:
                continue
            f0 = _fp(cand)
            if not f0:
                continue
            score = len(f0 & f1) / max(len(f0), len(f1))
            if score > best_score:
                best, best_score = cand, score
        return best if best_score >= 0.7 else None

    for it1 in r1_items:
        it0 = find_r0(it1)
        if it0 is None:
            q1 = _qty(it1)
            amt = (it1.get("amount") or ((q1 or 0.0) * (it1.get("total_rate") or 0.0))) or None
            res.entries.append(DiffEntry(
                status="added", sheet=it1["sheet"], r1_row=it1["row"],
                description=(it1.get("description") or "")[:120],
                qty_after=q1, amount_delta=amt))
            continue
        used_r0.add(id(it0))
        q0, q1 = _qty(it0), _qty(it1)
        rate0 = it0.get("total_rate")
        spec_diff = _spec_changed(it0, it1)
        if spec_diff:
            res.entries.append(DiffEntry(
                status="spec_changed", sheet=it1["sheet"],
                r0_row=it0["row"], r1_row=it1["row"],
                description=(it1.get("description") or "")[:120],
                qty_before=q0, qty_after=q1,
                detail="description/spec text diverged >30%; rate NOT carried"))
        elif (q0 or 0.0) != (q1 or 0.0):
            delta = None
            if rate0:
                delta = round(((q1 or 0.0) - (q0 or 0.0)) * rate0, 2)
            res.entries.append(DiffEntry(
                status="qty_changed", sheet=it1["sheet"],
                r0_row=it0["row"], r1_row=it1["row"],
                description=(it1.get("description") or "")[:120],
                qty_before=q0, qty_after=q1,
                rate_carried=rate0, amount_delta=delta,
                detail="rate carried forward to new qty"))
        else:
            res.entries.append(DiffEntry(
                status="unchanged", sheet=it1["sheet"],
                r0_row=it0["row"], r1_row=it1["row"],
                description=(it1.get("description") or "")[:120],
                qty_before=q0, qty_after=q1, rate_carried=rate0))

    for it0 in r0_items:
        if id(it0) not in used_r0:
            q0 = _qty(it0)
            amt = it0.get("amount") or ((q0 or 0.0) * (it0.get("total_rate") or 0.0)) or None
            res.entries.append(DiffEntry(
                status="removed", sheet=it0["sheet"], r0_row=it0["row"],
                description=(it0.get("description") or "")[:120],
                qty_before=q0, amount_delta=-amt if amt else None))
    return res


def diff_payloads(p0: dict, p1: dict) -> DiffResult:
    """Sheet-aligned diff of two extraction payloads."""
    res = DiffResult()
    sheets0 = {s["sheet"]: s for s in p0["sheets"] if s["kind"] == "boq"}
    sheets1 = {s["sheet"]: s for s in p1["sheets"] if s["kind"] == "boq"}
    for name, s1 in sheets1.items():
        s0 = sheets0.get(name)
        if s0 is None:
            for it in s1["items"]:
                res.entries.append(DiffEntry(
                    status="added", sheet=name, r1_row=it["row"],
                    description=(it.get("description") or "")[:120],
                    qty_after=it.get("qty"),
                    amount_delta=it.get("amount"),
                    detail="whole sheet added in revision"))
            continue
        sub = diff_items(s0["items"], s1["items"])
        res.entries.extend(sub.entries)
    for name, s0 in sheets0.items():
        if name not in sheets1:
            for it in s0["items"]:
                res.entries.append(DiffEntry(
                    status="removed", sheet=name, r0_row=it["row"],
                    description=(it.get("description") or "")[:120],
                    qty_before=it.get("qty"),
                    amount_delta=-(it.get("amount") or 0.0) or None,
                    detail="whole sheet removed in revision"))
    return res


def diff_files(r0_file: str, r1_file: str) -> DiffResult:
    p0 = extract_workbook(load_workbook(r0_file))
    p1 = extract_workbook(load_workbook(r1_file))
    return diff_payloads(p0, p1)
