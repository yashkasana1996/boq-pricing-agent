"""LLM-assisted matching pass for quote items the deterministic matcher
couldn't place. The model only proposes (boq_row, confidence) pairs; code
applies the accept/queue/drop thresholds. Requires ANTHROPIC_API_KEY."""

from __future__ import annotations

import json
import os

from .match import Match

MODEL = os.environ.get("BOQ_LLM_MODEL", "claude-sonnet-4-6")

SYSTEM = """You match a vendor quote line to the correct BOQ item by
description AND specification (size, thickness, make, grade, unit).
Return STRICT JSON: {"matches": [{"quote_idx": int, "boq_idx": int|null,
"confidence": 0.0-1.0, "reason": "<=15 words"}]}
A wrong match corrupts a commercial offer: when unsure, use low confidence
or null. Never invent indexes."""


def _brief(it: dict, idx: int) -> str:
    return (f"[{idx}] unit={it.get('unit') or '?'} "
            f"desc={((it.get('description') or '') + ' ' + (it.get('spec') or ''))[:220]}")


def llm_match(boq: list[dict], quotes: list[dict], client=None,
              batch: int = 20) -> list[tuple[dict, Match | None]]:
    if client is None:
        import anthropic

        client = anthropic.Anthropic()

    out: list[tuple[dict, Match | None]] = []
    boq_lines = "\n".join(_brief(b, i) for i, b in enumerate(boq))
    for start in range(0, len(quotes), batch):
        chunk = quotes[start:start + batch]
        q_lines = "\n".join(_brief(q, i) for i, q in enumerate(chunk))
        resp = client.messages.create(
            model=MODEL, max_tokens=2000, system=SYSTEM,
            messages=[{"role": "user", "content":
                       f"BOQ ITEMS:\n{boq_lines}\n\nQUOTE LINES:\n{q_lines}\n\nJSON:"}],
        )
        text = resp.content[0].text.strip()
        if text.startswith("```"):
            text = text.strip("`").lstrip("json").strip()
        try:
            data = json.loads(text)
        except json.JSONDecodeError:
            data = {"matches": []}
        seen = set()
        for m in data.get("matches", []):
            qi = m.get("quote_idx")
            if not isinstance(qi, int) or qi < 0 or qi >= len(chunk) or qi in seen:
                continue
            seen.add(qi)
            bi = m.get("boq_idx")
            conf = float(m.get("confidence") or 0.0)
            if bi is None or not isinstance(bi, int) or bi < 0 or bi >= len(boq):
                out.append((chunk[qi], None))
            else:
                out.append((chunk[qi], Match(boq[bi], chunk[qi], round(conf, 3), "llm")))
        for i, q in enumerate(chunk):
            if i not in seen:
                out.append((q, None))
    return out
