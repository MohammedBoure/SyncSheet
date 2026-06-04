"""Small, safe formula evaluator for spreadsheet cells."""

from __future__ import annotations

import ast
import math
import re
from datetime import date, datetime
from typing import Any, Callable

from .cell_address import iter_range, normalize_cell_ref, parse_cell_ref
from .workbook import WorkbookData, WorksheetData

CELL_PATTERN = re.compile(r"(?<![A-Za-z0-9_])(\$?[A-Za-z]{1,4}\$?[1-9][0-9]*)(?![A-Za-z0-9_])")
RANGE_PATTERN = re.compile(
    r"(?<![A-Za-z0-9_])(\$?[A-Za-z]{1,4}\$?[1-9][0-9]*):(\$?[A-Za-z]{1,4}\$?[1-9][0-9]*)(?![A-Za-z0-9_])"
)


class FormulaError(Exception):
    pass


class FormulaCycleError(FormulaError):
    pass


def _split_outside_strings(text: str):
    current: list[str] = []
    quote: str | None = None
    escaped = False
    for char in text:
        if quote:
            current.append(char)
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == quote:
                yield True, "".join(current)
                current = []
                quote = None
        else:
            if char in ("'", '"'):
                if current:
                    yield False, "".join(current)
                    current = []
                current.append(char)
                quote = char
            else:
                current.append(char)
    if current:
        yield quote is not None, "".join(current)


def _replace_outside_strings(text: str, pattern: re.Pattern[str], repl: Callable[[re.Match[str]], str]) -> str:
    pieces = []
    for is_string, part in _split_outside_strings(text):
        pieces.append(part if is_string else pattern.sub(repl, part))
    return "".join(pieces)


def flatten(values: Any) -> list[Any]:
    if isinstance(values, (list, tuple)):
        result: list[Any] = []
        for value in values:
            result.extend(flatten(value))
        return result
    return [values]


def is_blank(value: Any) -> bool:
    return value is None or value == ""


def to_number(value: Any) -> float:
    if is_blank(value):
        return 0.0
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip())
    except ValueError as exc:
        raise FormulaError(f"Cannot convert {value!r} to number") from exc


def coerce_display(value: Any) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, date):
        return value.isoformat()
    return "" if value is None else str(value)


def coerce_reference_value(value: Any) -> Any:
    if is_blank(value):
        return 0
    if isinstance(value, str):
        try:
            number = to_number(value)
            return int(number) if number.is_integer() else number
        except FormulaError:
            return value
    return value


class FormulaEvaluator:
    """Evaluate a focused Excel-like formula subset without exposing Python builtins."""

    allowed_nodes = (
        ast.Expression,
        ast.BinOp,
        ast.UnaryOp,
        ast.BoolOp,
        ast.Compare,
        ast.IfExp,
        ast.Call,
        ast.Name,
        ast.Load,
        ast.Constant,
        ast.List,
        ast.Tuple,
        ast.Add,
        ast.Sub,
        ast.Mult,
        ast.Div,
        ast.FloorDiv,
        ast.Mod,
        ast.Pow,
        ast.USub,
        ast.UAdd,
        ast.Eq,
        ast.NotEq,
        ast.Lt,
        ast.LtE,
        ast.Gt,
        ast.GtE,
        ast.And,
        ast.Or,
        ast.Not,
    )

    def __init__(self, workbook: WorkbookData):
        self.workbook = workbook
        self.functions = {
            "SUM": self.func_sum,
            "AVERAGE": self.func_average,
            "MIN": self.func_min,
            "MAX": self.func_max,
            "COUNT": self.func_count,
            "COUNTA": self.func_counta,
            "IF": self.func_if,
            "ROUND": round,
            "ABS": abs,
            "SQRT": math.sqrt,
            "POWER": pow,
            "CONCAT": self.func_concat,
            "LEN": lambda value: len(str(value)),
            "LEFT": lambda value, count=1: str(value)[: int(count)],
            "RIGHT": lambda value, count=1: str(value)[-int(count) :],
            "UPPER": lambda value: str(value).upper(),
            "LOWER": lambda value: str(value).lower(),
            "AND": self.func_and,
            "OR": self.func_or,
            "NOT": lambda value: not bool(value),
            "TODAY": lambda: date.today(),
            "NOW": lambda: datetime.now(),
        }

    def evaluate_cell(self, sheet: WorksheetData, row: int, column: int, visiting: set[tuple[str, int, int]] | None = None) -> Any:
        raw = sheet.raw_value(row, column)
        return self.evaluate(raw, sheet, visiting=visiting, cell_position=(row, column))

    def evaluate(
        self,
        raw_value: Any,
        sheet: WorksheetData,
        visiting: set[tuple[str, int, int]] | None = None,
        cell_position: tuple[int, int] | None = None,
    ) -> Any:
        if not isinstance(raw_value, str) or not raw_value.startswith("="):
            return raw_value
        visiting = visiting or set()
        if cell_position is not None:
            key = (sheet.name, cell_position[0], cell_position[1])
            if key in visiting:
                raise FormulaCycleError("Circular formula reference")
            visiting.add(key)
        expression = self.prepare_expression(raw_value[1:])
        try:
            tree = ast.parse(expression, mode="eval")
            self._validate_ast(tree)
            return eval(
                compile(tree, "<formula>", "eval"),
                {"__builtins__": {}},
                self._environment(sheet, visiting),
            )
        except FormulaError:
            raise
        except Exception as exc:
            raise FormulaError(str(exc)) from exc
        finally:
            if cell_position is not None:
                visiting.discard((sheet.name, cell_position[0], cell_position[1]))

    def display(self, raw_value: Any, sheet: WorksheetData, row: int | None = None, column: int | None = None) -> str:
        try:
            return coerce_display(self.evaluate(raw_value, sheet, cell_position=(row, column) if row is not None and column is not None else None))
        except FormulaError as exc:
            return f"#ERROR: {exc}"

    def prepare_expression(self, expression: str) -> str:
        prepared = expression.replace("^", "**")
        prepared = prepared.replace("<>", "!=")
        prepared = _replace_outside_strings(prepared, re.compile(r"(?<![<>=!])=(?!=)"), lambda _m: "==")
        prepared = _replace_outside_strings(prepared, re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\s*\("), lambda m: f"{m.group(1).upper()}(")
        prepared = _replace_outside_strings(prepared, RANGE_PATTERN, lambda m: f'RANGE("{normalize_cell_ref(m.group(1))}:{normalize_cell_ref(m.group(2))}")')
        prepared = _replace_outside_strings(prepared, CELL_PATTERN, lambda m: f'CELL("{normalize_cell_ref(m.group(1))}")')
        return prepared

    def _validate_ast(self, tree: ast.AST) -> None:
        for node in ast.walk(tree):
            if not isinstance(node, self.allowed_nodes):
                raise FormulaError(f"Unsupported formula syntax: {node.__class__.__name__}")
            if isinstance(node, ast.Call) and not isinstance(node.func, ast.Name):
                raise FormulaError("Only direct formula functions are allowed")
            if isinstance(node, ast.Name) and node.id not in self.functions and node.id not in {"CELL", "RANGE", "True", "False"}:
                raise FormulaError(f"Unknown formula name: {node.id}")

    def _environment(self, sheet: WorksheetData, visiting: set[tuple[str, int, int]]) -> dict[str, Any]:
        env = dict(self.functions)

        def cell(address: str) -> Any:
            ref = parse_cell_ref(address)
            value = self.evaluate_cell(sheet, ref.row, ref.column, visiting)
            return coerce_reference_value(value)

        def cell_range(address: str) -> list[Any]:
            start, end = address.split(":", 1)
            values = []
            for ref in iter_range(start, end):
                value = self.evaluate_cell(sheet, ref.row, ref.column, visiting)
                values.append(coerce_reference_value(value))
            return values

        env["CELL"] = cell
        env["RANGE"] = cell_range
        env["True"] = True
        env["False"] = False
        return env

    def func_sum(self, *values: Any) -> float:
        return sum(to_number(value) for value in flatten(values))

    def func_average(self, *values: Any) -> float:
        numbers = [to_number(value) for value in flatten(values) if not is_blank(value)]
        return sum(numbers) / len(numbers) if numbers else 0

    def func_min(self, *values: Any) -> float:
        numbers = [to_number(value) for value in flatten(values) if not is_blank(value)]
        return min(numbers) if numbers else 0

    def func_max(self, *values: Any) -> float:
        numbers = [to_number(value) for value in flatten(values) if not is_blank(value)]
        return max(numbers) if numbers else 0

    def func_count(self, *values: Any) -> int:
        count = 0
        for value in flatten(values):
            if is_blank(value):
                continue
            try:
                to_number(value)
            except FormulaError:
                continue
            count += 1
        return count

    def func_counta(self, *values: Any) -> int:
        return sum(1 for value in flatten(values) if not is_blank(value))

    def func_if(self, condition: Any, true_value: Any, false_value: Any = "") -> Any:
        return true_value if bool(condition) else false_value

    def func_concat(self, *values: Any) -> str:
        return "".join(coerce_display(value) for value in flatten(values))

    def func_and(self, *values: Any) -> bool:
        return all(bool(value) for value in flatten(values))

    def func_or(self, *values: Any) -> bool:
        return any(bool(value) for value in flatten(values))
