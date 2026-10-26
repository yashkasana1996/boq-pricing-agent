"""Extract structured BOQ items from sheet grids.

Deterministic-first design: everything in this module works without an LLM.
An optional LLM assist (extract_llm.py) can be layered on for sheets this
module classifies as `unparsed`, but all arithmetic and structure detection
lives here (CLAUDE.md non-negotiable: the LLM never does arithmetic).

Encodes fixtures/INVENTORY.md hazards:
 #3  content-driven multi-row header detection (no fixed header row)
 #4  section codes read from in-sheet headers, never from tab names
 #5  items are multi-row clusters (description continuation rows)
 #6  Sl.No formulas are advisory only — sequence is inferred positionally
 #7  per-level qty sub-columns under a merged Quantity header
 #8  supply/install column pairs detected per sheet
 #9  vendor codes may embed "BOQ Sl No-<n>" cross-refs
 #11 trailing Total/T&C/signature rows terminate the table by content
 #16 rate==0 on section rows; N()-coerced flood-filled amount formulas
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field, asdict

from .formula import resolve_number
from .ingest import SheetGrid, Workbook

# --------------------------------------------------------------------------
# vocabulary
# --------------------------------------------------------------------------

HDR_PATTERNS: dict[str, re.Pattern] = {
    "sr": re.compile(r"^(sr\.?\s*no|s\.?\s*no|sl\.?\s*no|si\.?\s*no|item\s*no|#)\.?\s*$", re.I),
    "code": re.compile(r"^code$|^sku$", re.I),
    "description": re.compile(r"desc|particular|item of work|scope of work|^product$", re.I),
    "spec": re.compile(r"specification", re.I),
    "make": re.compile(r"^(considered\s+)?make(s)?(\s+consid?ered)?\s*$|^brand|^make & model", re.I),
    "unit": re.compile(r"^(unit|uom)\s*$", re.I),
    "total_qty": re.compile(r"^total\s*\(?\s*qty\.?\)?\s*$|^total\s+quantity", re.I),
    "qty": re.compile(r"quantity|^qty|^qnty", re.I),
    "rate": re.compile(r"^(unit\s+)?rate(\s*\(.*\))?\s*$|^rate\s+\(inr\)|^offer(\s*\(.*\))?\s*$", re.I),
    "supply_rate": re.compile(r"^supply\s*(rate)?\s*$", re.I),
    "install_rate": re.compile(r"^installation\s*(rate)?\s*$|^install\s*(rate)?\s*$", re.I),
    "amount": re.compile(r"^(total\s+)?amount(\s*\(.*\))?\s*$|^amount\s+\(inr\)", re.I),
    "install_amount": re.compile(r"^installation\s+amount", re.I),
    "location": re.compile(r"^location", re.I),
    "remarks": re.compile(r"remark", re.I),
    "basic_rate": re.compile(r"^basic\s+rate", re.I),
    "ref_image": re.compile(r"image", re.I),
    "hsn": re.compile(r"^hsn", re.I),
    "mrp": re.compile(r"^mrp", re.I),
}

LEVEL_QTY = re.compile(r"^(lvl|l|level)[\s.\-]*\d+\s*$|^\d+\s*(st|nd|rd|th)(\s+floor)?\s*$", re.I)
TOTAL_ROW = re.compile(r"^(sub\s*-?\s*total|total|grand\s*total|total\s*w/o\s*tax).{0,30}$", re.I)
TRAILING_BLOCK = re.compile(
    r"terms\s*&?\s*conditions|payment\s+terms|thanking\s+you|authori[sz]ed\s+signator|"
    r"validity|force\s+majeure|delivery\s+(period|timeline|schedule)|gst\s+extra|"
    r"(exclusive|exculsive|inclusive)\s+of\s+gst|taxes\s+extra",
    re.I,
)
CARRIED = re.compile(r"carried\s+(over|forward)|brought\s+forward|\bb/f\b|\bc/f\b", re.I)
SECTION_CODE = re.compile(r"^[A-Z]{1,3}[-.]?\d{1,3}$")
BOQ_SLNO_REF = re.compile(r"BOQ\s+Sl\s*\.?\s*No\s*[-:.]?\s*([\d.]+)", re.I)
GST_IN_TEXT = re.compile(r"gst[^%\n]{0,40}?(\d{1,2}(?:\.\d+)?)\s*%|(\d{1,2}(?:\.\d+)?)\s*%\s*gst", re.I)

SHEET_CLASS_HINTS = {
    "cover": re.compile(r"cover\s*sheet", re.I),
    "notes": re.compile(r"general\s+notes|leed", re.I),
    "summary": re.compile(r"summary", re.I),
    "compliance": re.compile(r"technical\s+compliance|technical\s+data", re.I),
}


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def to_number(v: object) -> float | None:
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("₹", "")
    m = re.fullmatch(r"-?\d+(\.\d+)?", s)
    return float(s) if m else None


# --------------------------------------------------------------------------
# data model
# --------------------------------------------------------------------------


@dataclass
class Item:
    sheet: str
    row: int  # anchor row (row carrying the serial/first data)
    last_row: int  # last row of the cluster
    sr: str | None = None
    code: str | None = None
    description: str = ""
    spec: str = ""  # continuation-row text beyond the first description cell
    make: str | None = None
    unit: str | None = None
    qty: float | None = None
    qty_levels: dict[str, float] = field(default_factory=dict)
    supply_rate: float | None = None
    install_rate: float | None = None
    rate: float | None = None  # combined rate when no split
    amount: float | None = None
    supply_amount: float | None = None
    install_amount: float | None = None
    section: str | None = None
    section_code: str | None = None
    location: str | None = None
    basic_rate_text: str | None = None
    boq_slno_ref: str | None = None  # extracted "(BOQ Sl No-7)" style cross-ref
    remarks: str | None = None

    @property
    def total_rate(self) -> float | None:
        """Comparison must ALWAYS be on supply+install total (CLAUDE.md §4)."""
        if self.supply_rate is not None or self.install_rate is not None:
            return (self.supply_rate or 0.0) + (self.install_rate or 0.0)
        return self.rate

    @property
    def full_text(self) -> str:
        return " ".join(t for t in (self.description, self.spec) if t)

    def to_dict(self) -> dict:
        d = asdict(self)
        d["total_rate"] = self.total_rate
        return d


@dataclass
class HeaderMap:
    header_row: int  # last header row (labels resolved through merges)
    data_start: int
    cols: dict[str, int] = field(default_factory=dict)  # field -> column index
    level_qty_cols: dict[str, int] = field(default_factory=dict)
    has_supply_install_split: bool = False
    score: int = 0


@dataclass
class SheetExtract:
    sheet: str
    kind: str  # boq | vendor_quote | price_list | make_list | summary | notes | cover | compliance | unparsed
    header: HeaderMap | None = None
    items: list[Item] = field(default_factory=list)
    section_titles: list[str] = field(default_factory=list)
    gst_percent: float | None = None
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "sheet": self.sheet,
            "kind": self.kind,
            "header_row": self.header.header_row if self.header else None,
            "columns": self.header.cols if self.header else {},
            "level_qty_cols": self.header.level_qty_cols if self.header else {},
            "supply_install_split": bool(self.header and self.header.has_supply_install_split),
            "gst_percent": self.gst_percent,
            "items": [i.to_dict() for i in self.items],
            "sections": self.section_titles,
            "warnings": self.warnings,
        }


# --------------------------------------------------------------------------
# header detection (content-driven, multi-row aware)
# --------------------------------------------------------------------------


def find_header(grid: SheetGrid, scan_rows: int = 40) -> HeaderMap | None:
    """Score candidate header rows; combine a row with the sub-row below it
    (merged multi-row headers put level/supply-install labels one row down)."""
    best: HeaderMap | None = None
    limit = min(grid.max_row, scan_rows)
    for r in range(1, limit + 1):
        cols: dict[str, int] = {}
        level_cols: dict[str, int] = {}
        qty_cols: list[int] = []  # a merged "Quantity" header can span N columns
        score = 0
        for c in range(1, grid.max_col + 1):
            t = grid.text(r, c)
            if not t:
                continue
            tl = t.strip()
            for fieldname, pat in HDR_PATTERNS.items():
                if fieldname == "qty" and pat.search(tl):
                    qty_cols.append(c)
                    if "qty" not in cols:
                        cols["qty"] = c
                        score += 2
                    break
                if fieldname not in cols and pat.search(tl):
                    cols[fieldname] = c
                    score += 2 if fieldname in ("description", "rate", "amount", "sr") else 1
                    break
        # several qty-labelled columns => rightmost is the total, rest are
        # per-level breakdowns (labels come from the sub-row when present)
        if len(qty_cols) > 1:
            cols["qty"] = qty_cols[-1]
            for lc in qty_cols[:-1]:
                lbl = grid.text(r + 1, lc).strip() or f"col{lc}"
                level_cols[lbl] = lc
        # level columns labelled on the header row itself (L5|L6|L7|L8 style)
        for c in range(1, grid.max_col + 1):
            t = grid.text(r, c).strip()
            if t and LEVEL_QTY.match(t) and c not in level_cols.values():
                level_cols[t] = c
        # a bare "Total" column immediately right of the level block is the
        # true total qty (a "Job" item may show 1 per level but total 1)
        if level_cols and "total_qty" not in cols:
            nxt = max(level_cols.values()) + 1
            if grid.text(r, nxt).strip().lower() == "total" and nxt not in cols.values():
                cols["total_qty"] = nxt
        if "description" not in cols and "spec" in cols:
            cols["description"] = cols.pop("spec")
        if score < 4 or "description" not in cols:
            continue

        # sub-row: per-level qty labels and supply/install labels live at r+1
        split = False
        subrow_labels = False
        for c in range(1, grid.max_col + 1):
            t = grid.text(r + 1, c)
            if t and LEVEL_QTY.match(t.strip()):
                level_cols[t.strip()] = c
                subrow_labels = True
            if t and HDR_PATTERNS["total_qty"].match(t.strip()):
                cols["total_qty"] = c
                subrow_labels = True
            if t and re.match(r"^(supply|installation|install)\s*$", t.strip(), re.I):
                split = True
                subrow_labels = True
                is_supply = bool(re.match(r"^supply", t.strip(), re.I))
                rate_key = "supply_rate" if is_supply else "install_rate"
                amt_key = "supply_amount" if is_supply else "install_amount"
                # first occurrence labels the rate pair, second the amount pair
                if rate_key not in cols:
                    cols[rate_key] = c
                elif amt_key not in cols:
                    cols[amt_key] = c
        # same-row split (vendor B style: explicit Installation Rate/Amount columns)
        if "install_rate" in cols or "install_amount" in cols:
            split = True
            if "rate" in cols and "supply_rate" not in cols:
                cols["supply_rate"] = cols.pop("rate")
            if "amount" in cols and "install_amount" in cols and "supply_amount" not in cols:
                cols["supply_amount"] = cols.pop("amount")

        hm = HeaderMap(
            header_row=r + 1 if subrow_labels else r,
            data_start=r + 2 if subrow_labels else r + 1,
            cols=cols,
            level_qty_cols=level_cols,
            has_supply_install_split=split,
            score=score,
        )
        if best is None or score > best.score:
            best = hm
    return best


# --------------------------------------------------------------------------
# sheet classification
# --------------------------------------------------------------------------


def classify_sheet(grid: SheetGrid, header: HeaderMap | None) -> str:
    for kind, pat in SHEET_CLASS_HINTS.items():
        if pat.search(grid.name):
            return {"cover": "cover", "notes": "notes", "summary": "summary", "compliance": "compliance"}[kind]
    if header is None:
        return "unparsed"
    c = header.cols
    if "mrp" in c:
        return "price_list"
    has_pricing = any(k in c for k in ("rate", "supply_rate", "amount", "supply_amount"))
    if "make" in c and not has_pricing:
        return "make_list"
    if has_pricing:
        return "boq"  # vendor quotes are BOQ-shaped; caller distinguishes by file role
    return "make_list" if "make" in c else "unparsed"


# --------------------------------------------------------------------------
# item extraction
# --------------------------------------------------------------------------


def _row_text_joined(grid: SheetGrid, r: int) -> str:
    return " ".join(t for t in grid.row_texts(r) if t)


def _is_section_row(grid: SheetGrid, r: int, hm: HeaderMap) -> tuple[bool, str | None, str | None]:
    """Section header: a code like 'A-2' in the sr column (or col A) plus an
    ALL-CAPS title in the description area, with no qty/rate/amount values."""
    sr_c = hm.cols.get("sr", 1)
    desc_c = hm.cols["description"]
    sr_t = grid.text(r, sr_c)
    code = sr_t.strip() if SECTION_CODE.match(sr_t.strip()) else None
    title = None
    for c in (desc_c, desc_c - 1, desc_c + 1):
        t = grid.text(r, c)
        if t and t.upper() == t and len(t) > 3 and not to_number(t):
            title = t.strip()
            break
    has_numbers = any(
        to_number(grid.value(r, hm.cols[k])) not in (None, 0.0)
        for k in ("qty", "rate", "amount", "supply_rate")
        if k in hm.cols
    )
    if code and title and not has_numbers:
        return True, code, title
    # numbered top-level section like "1 | VITRIFIED FLOORING" with no pricing
    sr_num = to_number(sr_t)
    if sr_num is not None and float(sr_num).is_integer() and title and not has_numbers:
        nxt = _row_text_joined(grid, r + 1)
        if len(nxt) > 120:  # followed by a long spec paragraph
            return True, None, title
    return False, None, None


def extract_sheet(grid: SheetGrid, kind_hint: str | None = None) -> SheetExtract:
    hm = find_header(grid)
    kind = kind_hint or classify_sheet(grid, hm)
    ext = SheetExtract(sheet=grid.name, kind=kind, header=hm)
    if hm is None or kind in ("cover", "notes", "summary", "compliance", "unparsed"):
        return ext

    c = hm.cols
    desc_c = c["description"]
    sr_c = c.get("sr")
    # effective qty column: explicit Total-Qty wins; a plain "Quantity" header
    # sitting on a level sub-column must not be trusted as the total
    qty_c = c.get("total_qty") or c.get("qty")
    if qty_c in hm.level_qty_cols.values() and "total_qty" not in c:
        qty_c = None  # only per-level values exist; total comes from summing
    current: Item | None = None
    section_title: str | None = None
    section_code: str | None = None
    in_trailing = False

    for r in range(hm.data_start, grid.max_row + 1):
        joined = _row_text_joined(grid, r)
        if not joined:
            continue

        desc_t = grid.text(r, desc_c)
        qty_v = resolve_number(grid, r, qty_c) if qty_c else None
        rate_v = resolve_number(grid, r, c["rate"]) if "rate" in c else None
        sup_r = resolve_number(grid, r, c["supply_rate"]) if "supply_rate" in c else None
        ins_r = resolve_number(grid, r, c["install_rate"]) if "install_rate" in c else None

        row_has_numbers = any(v not in (None, 0.0) for v in (qty_v, rate_v, sup_r, ins_r))
        if TRAILING_BLOCK.search(joined[:200]) and not row_has_numbers:
            in_trailing = True
        if in_trailing:
            m = GST_IN_TEXT.search(joined)
            if m and ext.gst_percent is None:
                ext.gst_percent = float(m.group(1) or m.group(2))
            continue

        if TOTAL_ROW.match(desc_t.strip()) or (CARRIED.search(joined) and qty_v in (None, 0.0)):
            continue

        is_sec, sc_code, sc_title = _is_section_row(grid, r, hm)
        if is_sec:
            section_code, section_title = sc_code or section_code, sc_title
            if sc_title:
                ext.section_titles.append(sc_title)
            current = None
            continue

        # vertically-merged header labels bleed into the row below the header
        # (A5:A6 merges put "Description"/"Considered Make" text on row 6 too)
        if HDR_PATTERNS["description"].search(desc_t.strip()) and len(desc_t) < 30 and not row_has_numbers:
            continue

        sr_v = grid.text(r, sr_c) if sr_c else ""
        if not sr_v and sr_c:
            # sr formulas (=A11+0.01 chains) lose their cache on re-save —
            # recover the serial through the evaluator (hazard #6/#17)
            sr_num = resolve_number(grid, r, sr_c)
            if sr_num is not None:
                sr_v = f"{sr_num:g}"
        has_sr = bool(sr_v) and (to_number(sr_v) is not None or SECTION_CODE.match(sr_v.strip()) is not None or bool(re.match(r"^[\w\-/]+", sr_v)))
        has_data = any(v not in (None,) for v in (qty_v, rate_v, sup_r, ins_r)) and any(
            v not in (None, 0.0) for v in (qty_v, rate_v, sup_r, ins_r)
        )

        starts_item = (has_sr and desc_t) or (has_data and desc_t and current is None) or (has_data and desc_t and has_sr)
        if starts_item:
            item = Item(
                sheet=grid.name,
                row=r,
                last_row=r,
                sr=sr_v or None,
                code=grid.text(r, c["code"]) or None if "code" in c else None,
                description=desc_t,
                make=grid.text(r, c["make"]) or None if "make" in c else None,
                unit=grid.text(r, c["unit"]) or None if "unit" in c else None,
                qty=qty_v,
                supply_rate=sup_r,
                install_rate=ins_r,
                rate=rate_v if sup_r is None and ins_r is None else None,
                amount=resolve_number(grid, r, c["amount"]) if "amount" in c else None,
                supply_amount=resolve_number(grid, r, c["supply_amount"]) if "supply_amount" in c else None,
                install_amount=resolve_number(grid, r, c["install_amount"]) if "install_amount" in c else None,
                section=section_title,
                section_code=section_code,
                location=grid.text(r, c["location"]) or None if "location" in c else None,
                basic_rate_text=grid.text(r, c["basic_rate"]) or None if "basic_rate" in c else None,
                remarks=grid.text(r, c["remarks"]) or None if "remarks" in c else None,
            )
            for lvl, lc in hm.level_qty_cols.items():
                lv = resolve_number(grid, r, lc)
                if lv is not None:
                    item.qty_levels[lvl] = lv
            if item.qty is None and item.qty_levels:
                item.qty = sum(item.qty_levels.values())
            src = item.sr or ""
            m = BOQ_SLNO_REF.search(src) or BOQ_SLNO_REF.search(item.description)
            if m:
                item.boq_slno_ref = m.group(1)
            ext.items.append(item)
            current = item
        elif current is not None and desc_t and not has_data:
            # continuation row: only description text — hazard #5
            current.spec = (current.spec + "\n" + desc_t).strip()
            current.last_row = r
        elif current is not None and desc_t and has_data and not has_sr:
            # data row without serial — sub-line of the current item cluster;
            # treat as its own item to keep amounts complete (same full field
            # set as the serial branch: sr formulas with lost caches land here)
            item = Item(
                sheet=grid.name, row=r, last_row=r, description=desc_t,
                code=grid.text(r, c["code"]) or None if "code" in c else None,
                make=grid.text(r, c["make"]) or None if "make" in c else None,
                unit=grid.text(r, c["unit"]) or None if "unit" in c else None,
                qty=qty_v, supply_rate=sup_r, install_rate=ins_r,
                rate=rate_v if sup_r is None and ins_r is None else None,
                amount=resolve_number(grid, r, c["amount"]) if "amount" in c else None,
                supply_amount=resolve_number(grid, r, c["supply_amount"]) if "supply_amount" in c else None,
                install_amount=resolve_number(grid, r, c["install_amount"]) if "install_amount" in c else None,
                section=section_title, section_code=section_code,
                location=grid.text(r, c["location"]) or None if "location" in c else None,
                remarks=grid.text(r, c["remarks"]) or None if "remarks" in c else None,
            )
            for lvl, lc in hm.level_qty_cols.items():
                lv = resolve_number(grid, r, lc)
                if lv is not None:
                    item.qty_levels[lvl] = lv
            if item.qty is None and item.qty_levels:
                item.qty = sum(item.qty_levels.values())
            ext.items.append(item)
            current = item

    # rate==0 flood-fill artifacts: null out zero rates on items with no qty
    for it in ext.items:
        if it.rate == 0.0 and (it.qty in (None, 0.0)):
            it.rate = None
        if it.amount == 0.0 and (it.qty in (None, 0.0)):
            it.amount = None
        # split sheets: combined amount = supply amount + install amount.
        # Also covers the header layout where the merged "Total Amount" label
        # made cols['amount'] alias the supply half of the pair.
        if (it.supply_amount is not None or it.install_amount is not None) and (
            it.amount is None or it.amount == it.supply_amount
        ):
            it.amount = (it.supply_amount or 0.0) + (it.install_amount or 0.0)
    ext.items = [
        it for it in ext.items
        if it.qty not in (None, 0.0) or it.total_rate not in (None, 0.0) or (it.amount not in (None, 0.0))
        or (it.sr and it.description) or (it.code and it.description)
    ]
    return ext


def extract_workbook(wb: Workbook, sheet_kinds: dict[str, str] | None = None) -> dict:
    """Extract every sheet; returns the validated JSON-able payload."""
    sheets = []
    for sn, grid in wb.sheets.items():
        hint = (sheet_kinds or {}).get(sn)
        try:
            sheets.append(extract_sheet(grid, hint).to_dict())
        except Exception as e:  # never let one sheet kill the workbook
            sheets.append({"sheet": sn, "kind": "error", "items": [], "warnings": [f"{type(e).__name__}: {e}"]})
    n_items = sum(len(s["items"]) for s in sheets)
    return {
        "file": str(wb.path.name),
        "file_hash": wb.file_hash,
        "sheet_count": len(sheets),
        "item_count": n_items,
        "sheets": sheets,
    }


def validate_extraction(payload: dict) -> list[str]:
    """Code-side validation per CLAUDE.md extract.py contract."""
    errors = []
    for s in payload["sheets"]:
        for i, it in enumerate(s["items"]):
            if not (it.get("description") or "").strip():
                errors.append(f"{s['sheet']}!row{it['row']}: empty description")
            q, tr, amt = it.get("qty"), it.get("total_rate"), it.get("amount")
            if q is not None and tr is not None and amt is not None and q and tr:
                if abs(q * tr - amt) > max(1.0, 0.01 * abs(amt)):
                    sup_a, ins_a = it.get("supply_amount"), it.get("install_amount")
                    if not (sup_a is not None or ins_a is not None):
                        errors.append(
                            f"{s['sheet']}!row{it['row']}: qty*rate={q * tr:.2f} != amount={amt:.2f}"
                        )
    return errors
