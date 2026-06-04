"""Small, safe formula evaluator for spreadsheet cells."""

from __future__ import annotations

import ast
import math
import random
import re
import statistics
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Any, Callable

from .cell_address import normalize_cell_ref, parse_cell_ref
from .workbook import WorkbookData, WorksheetData

CELL_PATTERN = re.compile(r"(?<![A-Za-z0-9_])(\$?[A-Za-z]{1,4}\$?[1-9][0-9]*)(?![A-Za-z0-9_])")
RANGE_PATTERN = re.compile(
    r"(?<![A-Za-z0-9_])(\$?[A-Za-z]{1,4}\$?[1-9][0-9]*):(\$?[A-Za-z]{1,4}\$?[1-9][0-9]*)(?![A-Za-z0-9_])"
)
FUNCTION_PATTERN = re.compile(r"\b([A-Za-z_][A-Za-z0-9_\.]*)\s*\(")
BOOL_PATTERN = re.compile(r"\b(TRUE|FALSE)\b", re.IGNORECASE)
EQUALS_PATTERN = re.compile(r"(?<![<>=!])=(?!=)")
PERCENT_PATTERN = re.compile(r'(\d+(?:\.\d+)?|CELL\("[^"]+"\)|\([^()]*\))%')
CRITERIA_PATTERN = re.compile(r"^(<=|>=|<>|=|<|>)(.*)$")


class FormulaError(Exception):
    pass


class FormulaCycleError(FormulaError):
    pass


@dataclass(frozen=True)
class ExcelRange:
    rows: tuple[tuple[Any, ...], ...]

    @property
    def values(self) -> list[Any]:
        return [value for row in self.rows for value in row]

    def column_values(self, column: int) -> list[Any]:
        return [row[column] for row in self.rows if column < len(row)]

    def row_values(self, row: int) -> list[Any]:
        if row < 0 or row >= len(self.rows):
            return []
        return list(self.rows[row])


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
    if isinstance(values, ExcelRange):
        return values.values
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
    if isinstance(value, datetime):
        value = value.date()
    if isinstance(value, date):
        return float((value - date(1899, 12, 30)).days)
    text = str(value).strip()
    if text.endswith("%"):
        return to_number(text[:-1]) / 100
    try:
        return float(text)
    except ValueError as exc:
        raise FormulaError(f"Cannot convert {value!r} to number") from exc


def try_number(value: Any) -> float | None:
    try:
        return to_number(value)
    except FormulaError:
        return None


def numeric_values(*values: Any, ignore_text: bool = True) -> list[float]:
    numbers: list[float] = []
    for value in flatten(values):
        if is_blank(value):
            continue
        try:
            numbers.append(to_number(value))
        except FormulaError:
            if not ignore_text:
                raise
    return numbers


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


def coerce_bool(value: Any) -> bool:
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in {"true", "yes", "1"}:
            return True
        if lowered in {"false", "no", "0", ""}:
            return False
    return bool(value)


def as_date(value: Any) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        return date(1899, 12, 30) + timedelta(days=int(value))
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise FormulaError(f"Cannot convert {value!r} to date")


def excel_round(value: Any, digits: Any = 0) -> float:
    number = to_number(value)
    places = int(to_number(digits))
    sign = -1 if number < 0 else 1
    absolute = abs(number)
    if places >= 0:
        factor = 10**places
        return sign * math.floor(absolute * factor + 0.5) / factor
    factor = 10 ** abs(places)
    return sign * math.floor(absolute / factor + 0.5) * factor


def criteria_match(value: Any, criteria: Any) -> bool:
    if callable(criteria):
        return bool(criteria(value))
    operator = "="
    target = criteria
    if isinstance(criteria, str):
        match = CRITERIA_PATTERN.match(criteria.strip())
        if match:
            operator, target = match.groups()
        elif "*" in criteria or "?" in criteria:
            pattern = "^" + re.escape(criteria).replace("\\*", ".*").replace("\\?", ".") + "$"
            return re.match(pattern, str(value), re.IGNORECASE) is not None
    left_number = try_number(value)
    right_number = try_number(target)
    if left_number is not None and right_number is not None:
        left: Any = left_number
        right: Any = right_number
    else:
        left = str(value).lower()
        right = str(target).lower()
    if operator in ("=", "=="):
        return left == right
    if operator in ("<>", "!="):
        return left != right
    if operator == ">":
        return left > right
    if operator == ">=":
        return left >= right
    if operator == "<":
        return left < right
    if operator == "<=":
        return left <= right
    return False


class FormulaEvaluator:
    """Evaluate an Excel-like formula subset without exposing Python builtins."""

    allowed_nodes = (
        ast.Expression,
        ast.BinOp,
        ast.UnaryOp,
        ast.BoolOp,
        ast.Compare,
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
        ast.BitAnd,
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
        self.functions = self._build_functions()

    def _build_functions(self) -> dict[str, Callable[..., Any]]:
        return {
            "SUM": self.func_sum,
            "SUMIF": self.func_sumif,
            "SUMIFS": self.func_sumifs,
            "AVERAGE": self.func_average,
            "AVERAGEIF": self.func_averageif,
            "MIN": self.func_min,
            "MAX": self.func_max,
            "COUNT": self.func_count,
            "COUNTA": self.func_counta,
            "COUNTIF": self.func_countif,
            "PRODUCT": self.func_product,
            "MEDIAN": self.func_median,
            "MODE": self.func_mode,
            "STDEV": self.func_stdev_s,
            "STDEV_S": self.func_stdev_s,
            "STDEV_P": self.func_stdev_p,
            "VAR": self.func_var_s,
            "VAR_S": self.func_var_s,
            "VAR_P": self.func_var_p,
            "IF": self.func_if,
            "IFS": self.func_ifs,
            "IFERROR": self.func_iferror,
            "ROUND": excel_round,
            "ROUNDUP": self.func_roundup,
            "ROUNDDOWN": self.func_rounddown,
            "ABS": lambda value: abs(to_number(value)),
            "SQRT": lambda value: math.sqrt(to_number(value)),
            "POWER": lambda value, power: to_number(value) ** to_number(power),
            "EXP": lambda value: math.exp(to_number(value)),
            "LN": lambda value: math.log(to_number(value)),
            "LOG": self.func_log,
            "LOG10": lambda value: math.log10(to_number(value)),
            "PI": lambda: math.pi,
            "SIN": lambda value: math.sin(to_number(value)),
            "COS": lambda value: math.cos(to_number(value)),
            "TAN": lambda value: math.tan(to_number(value)),
            "ASIN": lambda value: math.asin(to_number(value)),
            "ACOS": lambda value: math.acos(to_number(value)),
            "ATAN": lambda value: math.atan(to_number(value)),
            "ATAN2": lambda x, y: math.atan2(to_number(x), to_number(y)),
            "RADIANS": lambda value: math.radians(to_number(value)),
            "DEGREES": lambda value: math.degrees(to_number(value)),
            "INT": lambda value: math.floor(to_number(value)),
            "TRUNC": self.func_trunc,
            "CEILING": self.func_ceiling,
            "FLOOR": self.func_floor,
            "MOD": lambda number, divisor: to_number(number) % to_number(divisor),
            "SIGN": self.func_sign,
            "FACT": lambda value: math.factorial(int(to_number(value))),
            "COMBIN": lambda number, chosen: math.comb(int(to_number(number)), int(to_number(chosen))),
            "PERMUT": lambda number, chosen: math.perm(int(to_number(number)), int(to_number(chosen))),
            "GCD": self.func_gcd,
            "LCM": self.func_lcm,
            "RAND": lambda: random.random(),
            "RANDBETWEEN": lambda low, high: random.randint(int(to_number(low)), int(to_number(high))),
            "CONCAT": self.func_concat,
            "CONCATENATE": self.func_concat,
            "TEXTJOIN": self.func_textjoin,
            "LEN": lambda value: len(str(value)),
            "LEFT": lambda value, count=1: str(value)[: int(to_number(count))],
            "RIGHT": lambda value, count=1: str(value)[-int(to_number(count)) :],
            "MID": lambda value, start, count: str(value)[int(to_number(start)) - 1 : int(to_number(start)) - 1 + int(to_number(count))],
            "UPPER": lambda value: str(value).upper(),
            "LOWER": lambda value: str(value).lower(),
            "TRIM": lambda value: " ".join(str(value).split()),
            "PROPER": lambda value: str(value).title(),
            "FIND": self.func_find,
            "SEARCH": self.func_search,
            "SUBSTITUTE": self.func_substitute,
            "REPLACE": self.func_replace,
            "REPT": lambda value, count: str(value) * int(to_number(count)),
            "AND": self.func_and,
            "OR": self.func_or,
            "NOT": lambda value: not coerce_bool(value),
            "TRUE": lambda: True,
            "FALSE": lambda: False,
            "TODAY": lambda: date.today(),
            "NOW": lambda: datetime.now(),
            "DATE": self.func_date,
            "DATEVALUE": as_date,
            "YEAR": lambda value: as_date(value).year,
            "MONTH": lambda value: as_date(value).month,
            "DAY": lambda value: as_date(value).day,
            "DAYS": lambda end_date, start_date: (as_date(end_date) - as_date(start_date)).days,
            "WEEKDAY": self.func_weekday,
            "HOUR": lambda value: datetime.fromisoformat(str(value)).hour if not isinstance(value, datetime) else value.hour,
            "MINUTE": lambda value: datetime.fromisoformat(str(value)).minute if not isinstance(value, datetime) else value.minute,
            "SECOND": lambda value: datetime.fromisoformat(str(value)).second if not isinstance(value, datetime) else value.second,
            "INDEX": self.func_index,
            "MATCH": self.func_match,
            "VLOOKUP": self.func_vlookup,
            "HLOOKUP": self.func_hlookup,
            "CHOOSE": self.func_choose,
            "ISBLANK": is_blank,
            "ISNUMBER": lambda value: try_number(value) is not None,
            "ISTEXT": lambda value: isinstance(value, str) and try_number(value) is None,
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
            return self._eval_node(tree.body, self._environment(sheet, visiting))
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
        prepared = _replace_outside_strings(prepared, re.compile(r";"), lambda _m: ",")
        prepared = prepared.replace("<>", "!=")
        prepared = _replace_outside_strings(prepared, BOOL_PATTERN, lambda m: m.group(1).upper())
        prepared = _replace_outside_strings(prepared, EQUALS_PATTERN, lambda _m: "==")
        prepared = _replace_outside_strings(prepared, FUNCTION_PATTERN, lambda m: f"{m.group(1).upper().replace('.', '_')}(")
        prepared = _replace_outside_strings(prepared, RANGE_PATTERN, lambda m: f'RANGE("{normalize_cell_ref(m.group(1))}:{normalize_cell_ref(m.group(2))}")')
        prepared = _replace_outside_strings(prepared, CELL_PATTERN, lambda m: f'CELL("{normalize_cell_ref(m.group(1))}")')
        prepared = _replace_outside_strings(prepared, PERCENT_PATTERN, lambda m: f"({m.group(1)}/100)")
        return prepared

    def _validate_ast(self, tree: ast.AST) -> None:
        for node in ast.walk(tree):
            if not isinstance(node, self.allowed_nodes):
                raise FormulaError(f"Unsupported formula syntax: {node.__class__.__name__}")
            if isinstance(node, ast.Call) and not isinstance(node.func, ast.Name):
                raise FormulaError("Only direct formula functions are allowed")
            if isinstance(node, ast.Name) and node.id not in self.functions and node.id not in {"CELL", "RANGE", "TRUE", "FALSE", "True", "False"}:
                raise FormulaError(f"Unknown formula name: {node.id}")

    def _environment(self, sheet: WorksheetData, visiting: set[tuple[str, int, int]]) -> dict[str, Any]:
        env = dict(self.functions)

        def cell(address: str) -> Any:
            ref = parse_cell_ref(address)
            value = self.evaluate_cell(sheet, ref.row, ref.column, visiting)
            return coerce_reference_value(value)

        def cell_range(address: str) -> ExcelRange:
            start, end = address.split(":", 1)
            first = parse_cell_ref(start)
            last = parse_cell_ref(end)
            row_from, row_to = sorted((first.row, last.row))
            col_from, col_to = sorted((first.column, last.column))
            rows = []
            for row in range(row_from, row_to + 1):
                values = []
                for column in range(col_from, col_to + 1):
                    value = self.evaluate_cell(sheet, row, column, visiting)
                    values.append(coerce_reference_value(value))
                rows.append(tuple(values))
            return ExcelRange(tuple(rows))

        env["CELL"] = cell
        env["RANGE"] = cell_range
        env["TRUE"] = True
        env["FALSE"] = False
        env["True"] = True
        env["False"] = False
        return env

    def _eval_node(self, node: ast.AST, env: dict[str, Any]) -> Any:
        if isinstance(node, ast.Constant):
            return node.value
        if isinstance(node, ast.Name):
            if node.id in env:
                return env[node.id]
            raise FormulaError(f"Unknown formula name: {node.id}")
        if isinstance(node, ast.List):
            return [self._eval_node(item, env) for item in node.elts]
        if isinstance(node, ast.Tuple):
            return tuple(self._eval_node(item, env) for item in node.elts)
        if isinstance(node, ast.UnaryOp):
            value = self._eval_node(node.operand, env)
            if isinstance(node.op, ast.USub):
                return -to_number(value)
            if isinstance(node.op, ast.UAdd):
                return to_number(value)
            if isinstance(node.op, ast.Not):
                return not coerce_bool(value)
        if isinstance(node, ast.BoolOp):
            if isinstance(node.op, ast.And):
                for value_node in node.values:
                    if not coerce_bool(self._eval_node(value_node, env)):
                        return False
                return True
            if isinstance(node.op, ast.Or):
                for value_node in node.values:
                    if coerce_bool(self._eval_node(value_node, env)):
                        return True
                return False
        if isinstance(node, ast.BinOp):
            left = self._eval_node(node.left, env)
            right = self._eval_node(node.right, env)
            if isinstance(node.op, ast.BitAnd):
                return coerce_display(left) + coerce_display(right)
            if isinstance(node.op, ast.Add):
                return to_number(left) + to_number(right)
            if isinstance(node.op, ast.Sub):
                return to_number(left) - to_number(right)
            if isinstance(node.op, ast.Mult):
                return to_number(left) * to_number(right)
            if isinstance(node.op, ast.Div):
                divisor = to_number(right)
                if divisor == 0:
                    raise FormulaError("Division by zero")
                return to_number(left) / divisor
            if isinstance(node.op, ast.FloorDiv):
                divisor = to_number(right)
                if divisor == 0:
                    raise FormulaError("Division by zero")
                return to_number(left) // divisor
            if isinstance(node.op, ast.Mod):
                divisor = to_number(right)
                if divisor == 0:
                    raise FormulaError("Division by zero")
                return to_number(left) % divisor
            if isinstance(node.op, ast.Pow):
                return to_number(left) ** to_number(right)
        if isinstance(node, ast.Compare):
            left = self._eval_node(node.left, env)
            for operator, comparator in zip(node.ops, node.comparators):
                right = self._eval_node(comparator, env)
                if not self._compare(left, right, operator):
                    return False
                left = right
            return True
        if isinstance(node, ast.Call):
            return self._eval_call(node, env)
        raise FormulaError(f"Unsupported formula syntax: {node.__class__.__name__}")

    def _eval_call(self, node: ast.Call, env: dict[str, Any]) -> Any:
        name = node.func.id
        if name == "IF":
            if len(node.args) < 2:
                raise FormulaError("IF expects at least 2 arguments")
            condition = self._eval_node(node.args[0], env)
            branch = node.args[1] if coerce_bool(condition) else node.args[2] if len(node.args) > 2 else ast.Constant(value="")
            return self._eval_node(branch, env)
        if name == "IFERROR":
            if len(node.args) != 2:
                raise FormulaError("IFERROR expects 2 arguments")
            try:
                return self._eval_node(node.args[0], env)
            except FormulaError:
                return self._eval_node(node.args[1], env)
        if name == "IFS":
            if len(node.args) % 2 != 0:
                raise FormulaError("IFS expects condition/value pairs")
            for index in range(0, len(node.args), 2):
                if coerce_bool(self._eval_node(node.args[index], env)):
                    return self._eval_node(node.args[index + 1], env)
            raise FormulaError("No IFS condition matched")
        if name not in env or not callable(env[name]):
            raise FormulaError(f"Unknown formula function: {name}")
        return env[name](*(self._eval_node(argument, env) for argument in node.args))

    def _compare(self, left: Any, right: Any, operator: ast.cmpop) -> bool:
        left_number = try_number(left)
        right_number = try_number(right)
        if left_number is not None and right_number is not None:
            first: Any = left_number
            second: Any = right_number
        else:
            first = str(left).lower()
            second = str(right).lower()
        if isinstance(operator, ast.Eq):
            return first == second
        if isinstance(operator, ast.NotEq):
            return first != second
        if isinstance(operator, ast.Lt):
            return first < second
        if isinstance(operator, ast.LtE):
            return first <= second
        if isinstance(operator, ast.Gt):
            return first > second
        if isinstance(operator, ast.GtE):
            return first >= second
        raise FormulaError("Unsupported comparison")

    def func_sum(self, *values: Any) -> float:
        return sum(numeric_values(*values))

    def func_sumif(self, criteria_range: Any, criteria: Any, sum_range: Any | None = None) -> float:
        criteria_values = flatten(criteria_range)
        sum_values = flatten(criteria_range if sum_range is None else sum_range)
        return sum(to_number(value) for check, value in zip(criteria_values, sum_values) if criteria_match(check, criteria) and try_number(value) is not None)

    def func_sumifs(self, sum_range: Any, *criteria_pairs: Any) -> float:
        if len(criteria_pairs) % 2 != 0:
            raise FormulaError("SUMIFS expects criteria range / criteria pairs")
        sum_values = flatten(sum_range)
        checks = [(flatten(criteria_pairs[index]), criteria_pairs[index + 1]) for index in range(0, len(criteria_pairs), 2)]
        total = 0.0
        for offset, value in enumerate(sum_values):
            if all(offset < len(values) and criteria_match(values[offset], criteria) for values, criteria in checks):
                number = try_number(value)
                if number is not None:
                    total += number
        return total

    def func_average(self, *values: Any) -> float:
        numbers = numeric_values(*values)
        return sum(numbers) / len(numbers) if numbers else 0

    def func_averageif(self, criteria_range: Any, criteria: Any, average_range: Any | None = None) -> float:
        criteria_values = flatten(criteria_range)
        average_values = flatten(criteria_range if average_range is None else average_range)
        numbers = [to_number(value) for check, value in zip(criteria_values, average_values) if criteria_match(check, criteria) and try_number(value) is not None]
        return sum(numbers) / len(numbers) if numbers else 0

    def func_min(self, *values: Any) -> float:
        numbers = numeric_values(*values)
        return min(numbers) if numbers else 0

    def func_max(self, *values: Any) -> float:
        numbers = numeric_values(*values)
        return max(numbers) if numbers else 0

    def func_count(self, *values: Any) -> int:
        return len(numeric_values(*values))

    def func_counta(self, *values: Any) -> int:
        return sum(1 for value in flatten(values) if not is_blank(value))

    def func_countif(self, values: Any, criteria: Any) -> int:
        return sum(1 for value in flatten(values) if criteria_match(value, criteria))

    def func_product(self, *values: Any) -> float:
        numbers = numeric_values(*values)
        if not numbers:
            return 0
        result = 1.0
        for number in numbers:
            result *= number
        return result

    def func_median(self, *values: Any) -> float:
        numbers = numeric_values(*values)
        return statistics.median(numbers) if numbers else 0

    def func_mode(self, *values: Any) -> Any:
        flattened = [value for value in flatten(values) if not is_blank(value)]
        if not flattened:
            return 0
        return Counter(flattened).most_common(1)[0][0]

    def func_stdev_s(self, *values: Any) -> float:
        numbers = numeric_values(*values)
        return statistics.stdev(numbers) if len(numbers) > 1 else 0

    def func_stdev_p(self, *values: Any) -> float:
        numbers = numeric_values(*values)
        return statistics.pstdev(numbers) if numbers else 0

    def func_var_s(self, *values: Any) -> float:
        numbers = numeric_values(*values)
        return statistics.variance(numbers) if len(numbers) > 1 else 0

    def func_var_p(self, *values: Any) -> float:
        numbers = numeric_values(*values)
        return statistics.pvariance(numbers) if numbers else 0

    def func_if(self, condition: Any, true_value: Any, false_value: Any = "") -> Any:
        return true_value if coerce_bool(condition) else false_value

    def func_ifs(self, *values: Any) -> Any:
        for condition, result in zip(values[0::2], values[1::2]):
            if coerce_bool(condition):
                return result
        raise FormulaError("No IFS condition matched")

    def func_iferror(self, value: Any, fallback: Any) -> Any:
        return value if not isinstance(value, FormulaError) else fallback

    def func_roundup(self, value: Any, digits: Any = 0) -> float:
        number = to_number(value)
        places = int(to_number(digits))
        factor = 10**abs(places)
        sign = -1 if number < 0 else 1
        absolute = abs(number)
        if places >= 0:
            return sign * math.ceil(absolute * factor) / factor
        return sign * math.ceil(absolute / factor) * factor

    def func_rounddown(self, value: Any, digits: Any = 0) -> float:
        number = to_number(value)
        places = int(to_number(digits))
        factor = 10**abs(places)
        sign = -1 if number < 0 else 1
        absolute = abs(number)
        if places >= 0:
            return sign * math.floor(absolute * factor) / factor
        return sign * math.floor(absolute / factor) * factor

    def func_log(self, value: Any, base: Any = 10) -> float:
        return math.log(to_number(value), to_number(base))

    def func_trunc(self, value: Any, digits: Any = 0) -> float:
        number = to_number(value)
        places = int(to_number(digits))
        factor = 10**abs(places)
        sign = -1 if number < 0 else 1
        absolute = abs(number)
        if places >= 0:
            return sign * math.floor(absolute * factor) / factor
        return sign * math.floor(absolute / factor) * factor

    def func_ceiling(self, value: Any, significance: Any = 1) -> float:
        number = to_number(value)
        step = abs(to_number(significance)) or 1
        return math.ceil(number / step) * step

    def func_floor(self, value: Any, significance: Any = 1) -> float:
        number = to_number(value)
        step = abs(to_number(significance)) or 1
        return math.floor(number / step) * step

    def func_sign(self, value: Any) -> int:
        number = to_number(value)
        return 1 if number > 0 else -1 if number < 0 else 0

    def func_gcd(self, *values: Any) -> int:
        numbers = [abs(int(number)) for number in numeric_values(*values)]
        result = 0
        for number in numbers:
            result = math.gcd(result, number)
        return result

    def func_lcm(self, *values: Any) -> int:
        numbers = [abs(int(number)) for number in numeric_values(*values)]
        result = 1
        for number in numbers:
            result = abs(result * number) // math.gcd(result, number) if number else 0
        return result

    def func_concat(self, *values: Any) -> str:
        return "".join(coerce_display(value) for value in flatten(values))

    def func_textjoin(self, delimiter: Any, ignore_empty: Any, *values: Any) -> str:
        parts = [coerce_display(value) for value in flatten(values)]
        if coerce_bool(ignore_empty):
            parts = [part for part in parts if part != ""]
        return str(delimiter).join(parts)

    def func_find(self, needle: Any, haystack: Any, start: Any = 1) -> int:
        offset = int(to_number(start)) - 1
        position = str(haystack).find(str(needle), offset)
        if position < 0:
            raise FormulaError("Text not found")
        return position + 1

    def func_search(self, needle: Any, haystack: Any, start: Any = 1) -> int:
        offset = int(to_number(start)) - 1
        position = str(haystack).lower().find(str(needle).lower(), offset)
        if position < 0:
            raise FormulaError("Text not found")
        return position + 1

    def func_substitute(self, text: Any, old_text: Any, new_text: Any, instance_num: Any | None = None) -> str:
        source = str(text)
        old = str(old_text)
        new = str(new_text)
        if instance_num is None:
            return source.replace(old, new)
        target = int(to_number(instance_num))
        parts = source.split(old)
        if target <= 0 or target >= len(parts):
            return source
        return old.join(parts[:target]) + new + old.join(parts[target:])

    def func_replace(self, old_text: Any, start: Any, chars: Any, new_text: Any) -> str:
        source = str(old_text)
        index = int(to_number(start)) - 1
        count = int(to_number(chars))
        return source[:index] + str(new_text) + source[index + count :]

    def func_and(self, *values: Any) -> bool:
        return all(coerce_bool(value) for value in flatten(values))

    def func_or(self, *values: Any) -> bool:
        return any(coerce_bool(value) for value in flatten(values))

    def func_date(self, year: Any, month: Any, day: Any) -> date:
        y = int(to_number(year))
        m = int(to_number(month))
        d = int(to_number(day))
        y += (m - 1) // 12
        m = (m - 1) % 12 + 1
        return date(y, m, 1) + timedelta(days=d - 1)

    def func_weekday(self, value: Any, return_type: Any = 1) -> int:
        weekday = as_date(value).weekday()
        mode = int(to_number(return_type))
        if mode == 2:
            return weekday + 1
        if mode == 3:
            return weekday
        return ((weekday + 1) % 7) + 1

    def func_index(self, values: Any, row_num: Any, column_num: Any = 1) -> Any:
        row = int(to_number(row_num)) - 1
        column = int(to_number(column_num)) - 1
        if isinstance(values, ExcelRange):
            if row < 0 or row >= len(values.rows):
                raise FormulaError("INDEX row out of range")
            if column < 0 or column >= len(values.rows[row]):
                raise FormulaError("INDEX column out of range")
            return values.rows[row][column]
        flattened = flatten(values)
        if row < 0 or row >= len(flattened):
            raise FormulaError("INDEX out of range")
        return flattened[row]

    def func_match(self, lookup_value: Any, lookup_array: Any, match_type: Any = 0) -> int:
        values = flatten(lookup_array)
        exact = int(to_number(match_type)) == 0
        best_index = None
        for index, value in enumerate(values):
            if criteria_match(value, lookup_value):
                return index + 1
            if not exact and try_number(value) is not None and try_number(lookup_value) is not None and to_number(value) <= to_number(lookup_value):
                best_index = index + 1
        if best_index is not None:
            return best_index
        raise FormulaError("MATCH value not found")

    def func_vlookup(self, lookup_value: Any, table_array: Any, col_index_num: Any, range_lookup: Any = False) -> Any:
        if not isinstance(table_array, ExcelRange):
            raise FormulaError("VLOOKUP expects a range")
        column = int(to_number(col_index_num)) - 1
        approximate = coerce_bool(range_lookup)
        best_row = None
        for row in table_array.rows:
            if not row or column >= len(row):
                continue
            if criteria_match(row[0], lookup_value):
                return row[column]
            if approximate and try_number(row[0]) is not None and try_number(lookup_value) is not None and to_number(row[0]) <= to_number(lookup_value):
                best_row = row
        if best_row is not None:
            return best_row[column]
        raise FormulaError("VLOOKUP value not found")

    def func_hlookup(self, lookup_value: Any, table_array: Any, row_index_num: Any, range_lookup: Any = False) -> Any:
        if not isinstance(table_array, ExcelRange) or not table_array.rows:
            raise FormulaError("HLOOKUP expects a range")
        row_index = int(to_number(row_index_num)) - 1
        approximate = coerce_bool(range_lookup)
        best_column = None
        first_row = table_array.rows[0]
        for column, value in enumerate(first_row):
            if criteria_match(value, lookup_value):
                return table_array.rows[row_index][column]
            if approximate and try_number(value) is not None and try_number(lookup_value) is not None and to_number(value) <= to_number(lookup_value):
                best_column = column
        if best_column is not None:
            return table_array.rows[row_index][best_column]
        raise FormulaError("HLOOKUP value not found")

    def func_choose(self, index_num: Any, *values: Any) -> Any:
        index = int(to_number(index_num)) - 1
        if index < 0 or index >= len(values):
            raise FormulaError("CHOOSE index out of range")
        return values[index]
