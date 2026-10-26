"""Optional LLM assist for sheets the deterministic extractor can't parse.

Strictly additive: only sheets classified `unparsed` (or explicitly forced)
are sent to the model, and every number the model returns is re-validated in
code against the grid before acceptance (CLAUDE.md: the LLM never does
arithmetic; silent wrong numbers are the one unforgivable failure).

Requires ANTHROPIC_API_KEY. Without it, callers get a clean NoApiKey signal
and the pipeline continues deterministic-only.
"""

from __future__ import annotations

import json
import os

from .extract import Item, SheetExtract, to_number
from .ingest import SheetGrid, grid_to_text

MODEL = os.environ.get("BOQ_LLM_MODEL", "claude-sonnet-4-6")

SYSTEM = """You extract construction BOQ line items from a coordinate-addressed
spreadsheet dump. Return STRICT JSON only, no prose, matching this schema:
{"kind": "boq|vendor_quote|rate_list|make_list|price_list|notes|summary|cover|other",
 "items": [{"row": <int anchor row>, "sr": str|null, "code": str|null,
   "description": str, "make": str|null, "unit": str|null,
   "qty_cell": "<A1 coord or null>", "rate_cell": "<A1 coord or null>",
   "supply_rate_cell": "<A1 coord or null>", "install_rate_cell": "<A1 coord or null>",
   "amount_cell": "<A1 coord or null>", "section": str|null}]}
CRITICAL: never output numeric values for qty/rate/amount — output the CELL
ADDRESS where the number lives. Numbers are read from the sheet by code."""


class NoApiKey(RuntimeError):
    pass


def have_key() -> bool:
    return bool(os.environ.get("ANTHROPIC_API_KEY"))


def llm_extract_sheet(grid: SheetGrid, client=None) -> tuple[SheetExtract, dict]:
    """Ask the model to LOCATE items/cells; values are read back from the grid.

    Returns (SheetExtract, usage_stats). Raises NoApiKey when unconfigured.
    """
    if client is None:
        if not have_key():
            raise NoApiKey("ANTHROPIC_API_KEY not set — LLM assist unavailable")
        import anthropic

        client = anthropic.Anthropic()

    dump = grid_to_text(grid, max_rows=300)
    resp = client.messages.create(
        model=MODEL,
        max_tokens=8000,
        system=SYSTEM,
        messages=[{"role": "user", "content": f"Sheet dump:\n\n{dump}\n\nJSON:"}],
    )
    usage = {
        "model": MODEL,
        "input_tokens": resp.usage.input_tokens,
        "output_tokens": resp.usage.output_tokens,
    }
    text = resp.content[0].text.strip()
    if text.startswith("```"):
        text = text.strip("`").lstrip("json").strip()
    data = json.loads(text)

    ext = SheetExtract(sheet=grid.name, kind=data.get("kind", "unparsed"))
    from openpyxl.utils.cell import coordinate_to_tuple

    def cell_num(coord: str | None) -> float | None:
        if not coord:
            return None
        try:
            r, c = coordinate_to_tuple(coord)
        except ValueError:
            return None
        from .formula import resolve_number

        return resolve_number(grid, r, c)

    for raw in data.get("items", []):
        it = Item(
            sheet=grid.name,
            row=int(raw.get("row") or 0),
            last_row=int(raw.get("row") or 0),
            sr=raw.get("sr"),
            code=raw.get("code"),
            description=(raw.get("description") or "").strip(),
            make=raw.get("make"),
            unit=raw.get("unit"),
            qty=cell_num(raw.get("qty_cell")),
            rate=cell_num(raw.get("rate_cell")),
            supply_rate=cell_num(raw.get("supply_rate_cell")),
            install_rate=cell_num(raw.get("install_rate_cell")),
            amount=cell_num(raw.get("amount_cell")),
            section=raw.get("section"),
        )
        if it.supply_rate is not None or it.install_rate is not None:
            it.rate = None
        if it.description:
            ext.items.append(it)
    ext.warnings.append(f"llm-extracted ({usage['input_tokens']}in/{usage['output_tokens']}out tokens)")
    return ext, usage
