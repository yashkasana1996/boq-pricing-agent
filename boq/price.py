"""Pricing — deterministic arithmetic only (CLAUDE.md non-negotiable).

L1 selection on supply+install TOTALS, house markup/ROUND convention,
GST inclusive/exclusive, ₹/sqft per package, margin table, target back-solve.

Final rate convention (CLAUDE.md §6):  final = ROUND(L1_rate * (1 + markup%), r)
Amount = Qty x final rate. Package subtotals and grand total roll up.
"""

from __future__ import annotations

from dataclasses import dataclass, field


from decimal import ROUND_HALF_UP, Decimal


def excel_round(x: float, digits: int = 2) -> float:
    """Excel's ROUND: half away from zero, decimal-based.

    Uses Decimal(repr(x)) so 1.005 rounds to 1.01 the way Excel shows it,
    instead of falling into binary-float traps (1.005*100 == 100.4999...).
    """
    if x is None:
        return None
    q = Decimal(1).scaleb(-digits)  # 10^-digits
    return float(Decimal(repr(x)).quantize(q, rounding=ROUND_HALF_UP))


@dataclass
class CandidateRate:
    vendor: str
    total_rate: float
    supply_rate: float | None = None
    install_rate: float | None = None
    confidence: float = 1.0
    signal: str = ""
    history_flag: dict | None = None  # sanity-check deviation vs rate history


@dataclass
class PricedItem:
    sheet: str
    row: int
    description: str
    unit: str | None
    qty: float | None
    package: str  # summary package name (usually the sheet)
    candidates: list[CandidateRate] = field(default_factory=list)
    l1: CandidateRate | None = None
    final_rate: float | None = None
    amount: float | None = None
    status: str = "unpriced"  # priced | queued | unpriced | no_qty

    @property
    def cost_amount(self) -> float | None:
        if self.l1 and self.qty:
            return self.l1.total_rate * self.qty
        return None

    def to_dict(self) -> dict:
        return {
            "sheet": self.sheet, "row": self.row,
            "description": self.description[:160], "unit": self.unit,
            "qty": self.qty, "package": self.package, "status": self.status,
            "l1_vendor": self.l1.vendor if self.l1 else None,
            "l1_rate": self.l1.total_rate if self.l1 else None,
            "final_rate": self.final_rate, "amount": self.amount,
            "candidates": [
                {"vendor": c.vendor, "total_rate": c.total_rate,
                 "confidence": c.confidence, "signal": c.signal,
                 "history_flag": c.history_flag}
                for c in self.candidates
            ],
        }


def select_l1(item: PricedItem) -> None:
    """Lowest total (supply+install) rate wins. Ties keep first seen."""
    usable = [c for c in item.candidates if c.total_rate and c.total_rate > 0]
    item.l1 = min(usable, key=lambda c: c.total_rate) if usable else None


def apply_markup(item: PricedItem, markup_pct: float, round_digits: int = 2) -> None:
    if item.l1 is None:
        return
    item.final_rate = excel_round(item.l1.total_rate * (1 + markup_pct / 100.0), round_digits)
    if item.qty:
        item.amount = excel_round(item.qty * item.final_rate, 2)
        item.status = "priced"
    else:
        item.status = "no_qty"


@dataclass
class PricingRun:
    items: list[PricedItem]
    markup_pct: float
    gst_mode: str = "exclusive"  # exclusive: GST added on total; inclusive: rates carry GST
    gst_pct: float = 18.0
    area_sqft: float | None = None
    round_digits: int = 2

    def price_all(self) -> None:
        for it in self.items:
            select_l1(it)
            apply_markup(it, self.markup_pct, self.round_digits)

    # ---- rollups ------------------------------------------------------
    def package_totals(self) -> dict[str, dict]:
        out: dict[str, dict] = {}
        for it in self.items:
            p = out.setdefault(it.package, {"amount": 0.0, "cost": 0.0, "items": 0,
                                            "priced": 0, "rate_per_sqft": None})
            p["items"] += 1
            if it.amount:
                p["amount"] += it.amount
                p["priced"] += 1
            if it.cost_amount:
                p["cost"] += it.cost_amount
        if self.area_sqft:
            for p in out.values():
                p["rate_per_sqft"] = round(p["amount"] / self.area_sqft, 2)
        for p in out.values():
            p["amount"] = round(p["amount"], 2)
            p["cost"] = round(p["cost"], 2)
            p["margin"] = round(p["amount"] - p["cost"], 2)
            p["margin_pct"] = round(p["margin"] / p["amount"] * 100, 2) if p["amount"] else None
        return out

    def totals(self) -> dict:
        subtotal = round(sum(it.amount or 0.0 for it in self.items), 2)
        cost = round(sum(it.cost_amount or 0.0 for it in self.items), 2)
        if self.gst_mode == "exclusive":
            gst = excel_round(subtotal * self.gst_pct / 100.0, 2)
            grand = round(subtotal + gst, 2)
            ex_gst = subtotal
        else:  # rates already include GST
            grand = subtotal
            ex_gst = excel_round(subtotal / (1 + self.gst_pct / 100.0), 2)
            gst = round(grand - ex_gst, 2)
        out = {
            "subtotal_ex_gst": ex_gst,
            "gst_pct": self.gst_pct, "gst_mode": self.gst_mode, "gst_amount": gst,
            "grand_total": grand,
            "cost_basis": cost,
            "margin": round(ex_gst - cost, 2),
            "margin_pct": round((ex_gst - cost) / ex_gst * 100, 2) if ex_gst else None,
            "items_total": len(self.items),
            "items_priced": sum(1 for it in self.items if it.status == "priced"),
        }
        if self.area_sqft:
            out["rate_per_sqft_ex_gst"] = round(ex_gst / self.area_sqft, 2)
        return out


# ---------------------------------------------------------------------------
# techno-commercial: back-solve markup for a target total (M6)
# ---------------------------------------------------------------------------


def total_at_markup(items: list[PricedItem], markup_pct: float, gst_mode: str,
                    gst_pct: float, round_digits: int) -> float:
    """total(m) = Σ ROUND(L1_i x (1+m), r) x Qty_i, GST per mode — the exact
    per-item ROUND matters: a naive total*(1+m) drifts by paise x thousands."""
    subtotal = 0.0
    for it in items:
        if it.l1 is None:
            select_l1(it)
        if it.l1 is None or not it.qty:
            continue
        rate = excel_round(it.l1.total_rate * (1 + markup_pct / 100.0), round_digits)
        subtotal += excel_round(it.qty * rate, 2)
    if gst_mode == "exclusive":
        return round(subtotal + excel_round(subtotal * gst_pct / 100.0, 2), 2)
    return round(subtotal, 2)


def back_solve_markup(items: list[PricedItem], target_total: float, gst_mode: str = "exclusive",
                      gst_pct: float = 18.0, round_digits: int = 2,
                      lo: float = -50.0, hi: float = 200.0, tol: float = 0.5) -> dict:
    """Binary search m so total(m) hits the target (monotone non-decreasing)."""
    for it in items:
        if it.l1 is None:
            select_l1(it)
    f_lo = total_at_markup(items, lo, gst_mode, gst_pct, round_digits)
    f_hi = total_at_markup(items, hi, gst_mode, gst_pct, round_digits)
    if not (f_lo <= target_total <= f_hi):
        return {"feasible": False, "target": target_total,
                "total_at_lo": f_lo, "total_at_hi": f_hi,
                "note": f"target outside achievable band for markup in [{lo}%, {hi}%]"}
    for _ in range(64):
        mid = (lo + hi) / 2
        t = total_at_markup(items, mid, gst_mode, gst_pct, round_digits)
        if abs(t - target_total) <= tol:
            lo = hi = mid
            break
        if t < target_total:
            lo = mid
        else:
            hi = mid
    m = round((lo + hi) / 2, 4)
    return {"feasible": True, "target": target_total, "markup_pct": m,
            "achieved_total": total_at_markup(items, m, gst_mode, gst_pct, round_digits)}


def back_solve_target(state: "PricingRun", target_total: float, gst_mode: str = "exclusive",
                      gst_pct: float = 18.0, round_digits: int = 2,
                      margin_floor_pct: float = 8.0) -> dict:
    """Full M6 readout: solved markup, per-package impact, margin-floor warnings."""
    res = back_solve_markup(state.items, target_total, gst_mode, gst_pct, round_digits)
    if not res.get("feasible"):
        return res
    m = res["markup_pct"]
    baseline = state.package_totals()
    solved = PricingRun(items=state.items, markup_pct=m, gst_mode=gst_mode,
                        gst_pct=gst_pct, area_sqft=state.area_sqft,
                        round_digits=round_digits)
    solved.price_all()
    after = solved.package_totals()
    res["per_package_impact"] = {
        pkg: {"before": baseline.get(pkg, {}).get("amount"),
              "after": after[pkg]["amount"],
              "delta": round(after[pkg]["amount"] - baseline.get(pkg, {}).get("amount", 0.0), 2),
              "margin_pct_after": after[pkg]["margin_pct"]}
        for pkg in after
    }
    res["totals_after"] = solved.totals()
    warn = []
    if res["totals_after"]["margin_pct"] is not None and res["totals_after"]["margin_pct"] < margin_floor_pct:
        warn.append(f"OVERALL margin {res['totals_after']['margin_pct']}% below floor {margin_floor_pct}%")
    for pkg, p in after.items():
        if p["margin_pct"] is not None and p["margin_pct"] < margin_floor_pct:
            warn.append(f"package '{pkg}' margin {p['margin_pct']}% below floor {margin_floor_pct}%")
    res["margin_floor_warnings"] = warn
    # restore original pricing on shared item objects
    state.price_all()
    return res


def revise_cut(state: "PricingRun", package: str, cut_pct: float,
               round_digits: int = 2) -> dict:
    """M6 line-level cut: reduce final rates of one package by cut_pct."""
    changed = []
    for it in state.items:
        if it.package == package and it.final_rate:
            old = it.final_rate
            it.final_rate = excel_round(old * (1 - cut_pct / 100.0), round_digits)
            it.amount = excel_round((it.qty or 0.0) * it.final_rate, 2) if it.qty else None
            changed.append({"sheet": it.sheet, "row": it.row,
                            "rate_before": old, "rate_after": it.final_rate})
    return {"package": package, "cut_pct": cut_pct, "lines_changed": len(changed),
            "changes": changed[:50], "totals_after": state.totals()}
