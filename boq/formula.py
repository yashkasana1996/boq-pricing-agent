"""Tiny, safe Excel-formula evaluator for cache-less workbooks.

openpyxl loses cached formula values whenever a workbook has been re-saved
(INVENTORY hazard #17). Real BOQs keep qty/amounts in trivially simple
formulas (`=454-135`, `=SUM(I13:L13)`, `=N13*M13`, `=N(I3)*N($G3)`), so a
small evaluator recovers them deterministically — no LLM, no eval().

Supported: numbers, + - * / ( ), same-sheet cell refs (with $), SUM(range),
N(x) blank-coercion, ROUND(x, n). Anything else -> None (caller treats the
value as unknown, never as zero).
"""

from __future__ import annotations

import re

TOKEN = re.compile(
    r"\s*(?:"
    r"(?P<func>SUM|ROUND|N)\s*\(|"
    r"(?P<range>\$?[A-Z]{1,3}\$?\d+\s*:\s*\$?[A-Z]{1,3}\$?\d+)|"
    r"(?P<cell>\$?[A-Z]{1,3}\$?\d+)|"
    r"(?P<num>\d+(?:\.\d+)?)|"
    r"(?P<op>[+\-*/(),])"
    r")",
    re.I,
)

CELL_REF = re.compile(r"^\$?([A-Z]{1,3})\$?(\d+)$", re.I)

_MAX_DEPTH = 8


def _col_to_idx(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _tokenize(expr: str) -> list[tuple[str, str]] | None:
    out, pos = [], 0
    while pos < len(expr):
        m = TOKEN.match(expr, pos)
        if not m:
            if expr[pos:].strip() == "":
                break
            return None  # unsupported syntax (cross-sheet ref, string, etc.)
        if m.group("func"):
            out.append(("func", m.group("func").upper()))
        elif m.group("range"):
            out.append(("range", m.group("range").replace("$", "").replace(" ", "")))
        elif m.group("cell"):
            out.append(("cell", m.group("cell").replace("$", "")))
        elif m.group("num"):
            out.append(("num", m.group("num")))
        else:
            out.append(("op", m.group("op")))
        pos = m.end()
    return out


class _Parser:
    """Recursive-descent over the token list (precedence: +- < */ < unary)."""

    def __init__(self, tokens, resolver):
        self.toks = tokens
        self.i = 0
        self.resolve = resolver

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def take(self):
        t = self.peek()
        self.i += 1
        return t

    def expr(self):
        v = self.term()
        while self.peek() == ("op", "+") or self.peek() == ("op", "-"):
            _, op = self.take()
            rhs = self.term()
            if v is None or rhs is None:
                return None
            v = v + rhs if op == "+" else v - rhs
        return v

    def term(self):
        v = self.factor()
        while self.peek() == ("op", "*") or self.peek() == ("op", "/"):
            _, op = self.take()
            rhs = self.factor()
            if v is None or rhs is None:
                return None
            if op == "/":
                if rhs == 0:
                    return None
                v = v / rhs
            else:
                v = v * rhs
        return v

    def factor(self):
        kind, val = self.peek()
        if kind == "op" and val == "-":
            self.take()
            v = self.factor()
            return -v if v is not None else None
        if kind == "op" and val == "+":
            self.take()
            return self.factor()
        if kind == "op" and val == "(":
            self.take()
            v = self.expr()
            if self.peek() == ("op", ")"):
                self.take()
            return v
        if kind == "num":
            self.take()
            return float(val)
        if kind == "cell":
            self.take()
            return self.resolve("cell", val)
        if kind == "func":
            self.take()
            return self._call(val)
        return None

    def _call(self, name):
        args, coerce_blank = [], (name == "N")
        while True:
            kind, val = self.peek()
            if kind is None:
                return None
            if kind == "op" and val == ")":
                self.take()
                break
            if kind == "op" and val == ",":
                self.take()
                continue
            if kind == "range":
                self.take()
                args.append(("range", val))
                continue
            v = self.expr()
            args.append(("val", v))
        if name == "SUM":
            total = 0.0
            for k, a in args:
                if k == "range":
                    s = self.resolve("range", a)
                    if s is None:
                        return None
                    total += s
                elif a is not None:
                    total += a
            return total
        if name == "N":
            if not args:
                return 0.0
            k, a = args[0]
            if k == "range":
                return None
            return a if a is not None else 0.0
        if name == "ROUND":
            if not args or args[0][1] is None:
                return None
            digits = int(args[1][1]) if len(args) > 1 and args[1][1] is not None else 0
            return round(args[0][1], digits)
        return None


def evaluate(grid, formula: str, _depth: int = 0) -> float | None:
    """Evaluate a same-sheet formula against a SheetGrid. None = unknown."""
    if _depth > _MAX_DEPTH or not formula or not formula.startswith("="):
        return None
    body = formula[1:]
    if "!" in body or '"' in body or "[" in body:
        return None  # cross-sheet/external/string formulas are out of scope
    tokens = _tokenize(body)
    if tokens is None:
        return None

    def resolver(kind: str, ref: str) -> float | None:
        if kind == "cell":
            m = CELL_REF.match(ref)
            if not m:
                return None
            col, row = _col_to_idx(m.group(1)), int(m.group(2))
            return _cell_number(grid, row, col, _depth)
        if kind == "range":
            a, b = ref.split(":")
            ma, mb = CELL_REF.match(a), CELL_REF.match(b)
            if not (ma and mb):
                return None
            c1, r1 = _col_to_idx(ma.group(1)), int(ma.group(2))
            c2, r2 = _col_to_idx(mb.group(1)), int(mb.group(2))
            total = 0.0
            for rr in range(min(r1, r2), max(r1, r2) + 1):
                for cc in range(min(c1, c2), max(c1, c2) + 1):
                    v = _cell_number(grid, rr, cc, _depth, blank_is_zero=True)
                    if v is None:
                        return None  # a truly unknown cell poisons the sum
                    total += v
            return total
        return None

    p = _Parser(tokens, resolver)
    v = p.expr()
    return v if p.i >= len(p.toks) else None


def _cell_number(grid, row: int, col: int, depth: int, blank_is_zero: bool = False) -> float | None:
    cell = grid.cell(row, col)
    if cell is None:
        return 0.0 if blank_is_zero else 0.0  # empty cell is 0 in Excel math
    if cell.value is not None and not isinstance(cell.value, str):
        try:
            return float(cell.value)
        except (TypeError, ValueError):
            return None
    if cell.value is None and cell.formula:
        return evaluate(grid, cell.formula, depth + 1)
    if isinstance(cell.value, str):
        s = cell.value.strip().replace(",", "")
        if re.fullmatch(r"-?\d+(\.\d+)?", s):
            return float(s)
        return 0.0 if blank_is_zero else None  # text is 0 inside SUM, else unknown
    return 0.0


def resolve_number(grid, row: int, col: int) -> float | None:
    """Public helper: numeric value of a cell, evaluating simple formulas
    when the cached value is missing. None = genuinely unknown."""
    cell = grid.cell(row, col)
    if cell is None:
        return None
    if cell.value is not None:
        if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
            return float(cell.value)
        s = str(cell.value).strip().replace(",", "")
        if re.fullmatch(r"-?\d+(\.\d+)?", s):
            return float(s)
        return None
    if cell.formula:
        return evaluate(grid, cell.formula)
    return None
