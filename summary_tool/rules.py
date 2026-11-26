from collections.abc import Iterator, Mapping
from dataclasses import dataclass
import re

from openpyxl.utils.cell import column_index_from_string, get_column_letter

from .configuration import LEGACY_BINDINGS
from .functions import FunctionSpec, function_names, get_function, is_alias


class RuleError(ValueError):
    pass


@dataclass(frozen=True)
class CellReference:
    """单个参数的引用边界，单格等价于 1×1 区域。"""

    sheet: str
    coordinate: str
    row: int
    column: int
    end_row: int | None = None
    end_column: int | None = None

    @property
    def shape(self) -> tuple[int, int]:
        return (
            (self.end_row or self.row) - self.row + 1,
            (self.end_column or self.column) - self.column + 1,
        )

    def cells(self) -> Iterator["CellReference"]:
        if self.shape == (1, 1):
            yield CellReference(self.sheet, f"{get_column_letter(self.column)}{self.row}", self.row, self.column)
            return
        for row in range(self.row, self.end_row + 1):
            for column in range(self.column, self.end_column + 1):
                yield CellReference(self.sheet, f"{get_column_letter(column)}{row}", row, column)

    @property
    def label(self) -> str:
        return f"'{self.sheet.replace(chr(39), chr(39) * 2)}'!{self.coordinate}"


@dataclass(frozen=True)
class TemplateRule:
    sheet: str
    coordinate: str
    operation: str
    references: tuple[CellReference, ...]
    expression: str
    function_name: str = ""


# 仅将函数调用的中英文括号按同一语法解析，参数中的工作表名保持原文。
_FUNCTION = re.compile(r"([A-Za-z_][A-Za-z_0-9]*)\s*[(（](.*)[)）]", re.DOTALL)
_REFERENCE = re.compile(
    r"(?:(?P<sheet>'(?:[^']|'')+'|[^'!]+)\s*!\s*)?"
    r"\$?(?P<column>[A-Za-z]{1,3})\$?(?P<row>[1-9][0-9]{0,6})"
    r"(?:\s*:\s*\$?(?P<end_column>[A-Za-z]{1,3})\$?(?P<end_row>[1-9][0-9]{0,6}))?"
)


def _arguments(text: str) -> list[str]:
    result = []
    start = 0
    quoted = False
    index = 0
    while index < len(text):
        char = text[index]
        if char == "'":
            if quoted and index + 1 < len(text) and text[index + 1] == "'":
                index += 2
                continue
            quoted = not quoted
        elif char in ",，" and not quoted:
            result.append(text[start:index].strip())
            start = index + 1
        elif char in "()" and not quoted:
            raise RuleError("参数只支持单格或矩形区域引用，不支持嵌套函数。")
        index += 1
    if quoted:
        raise RuleError("工作表名的单引号没有闭合。")
    result.append(text[start:].strip())
    return result


def parse_reference(text: str, current_sheet: str) -> CellReference:
    match = _REFERENCE.fullmatch(text.strip())
    if match is None:
        raise RuleError(f"无效引用“{text}”。请使用 A5、A1:B3 或 '工作表名'!A1:B3，不支持运算表达式。")
    sheet = match.group("sheet")
    if sheet is None:
        sheet = current_sheet
    elif sheet.startswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    else:
        sheet = sheet.strip()
        if any(char.isspace() for char in sheet):
            raise RuleError("包含空格的工作表名请用单引号括起来。")
    if not sheet or len(sheet) > 31 or any(char in sheet for char in "[]:*?/\\"):
        raise RuleError(f"无效工作表名“{sheet}”，不支持外部工作簿引用。")
    column = match.group("column").upper()
    row = int(match.group("row"))
    column_index = column_index_from_string(column)
    end_row = int(match.group("end_row") or row)
    end_column = (match.group("end_column") or column).upper()
    end_column_index = column_index_from_string(end_column)
    if max(row, end_row) > 1048576 or max(column_index, end_column_index) > 16384:
        raise RuleError("单元格地址超出 Excel 范围（最大 XFD1048576）。")
    if end_row < row or end_column_index < column_index:
        raise RuleError("区域终点必须位于起点的右下方或同一位置。")
    coordinate = f"{column}{row}"
    if match.group("end_row") is not None:
        coordinate += f":{end_column}{end_row}"
        return CellReference(sheet, coordinate, row, column_index, end_row, end_column_index)
    return CellReference(sheet, coordinate, row, column_index)


def _resolve_function(name: str, bindings: Mapping[str, str] | None) -> FunctionSpec:
    definition = get_function(name)
    if definition is not None:
        return definition
    if not is_alias(name):
        raise RuleError(f"不支持指令 {name}，当前已实现 {'、'.join(function_names())}，也可使用已配置的 FUN_编号 别名。")
    normalized = {}
    for key, target in (bindings or {}).items():
        normalized_key = key.upper()
        if normalized_key in normalized:
            raise RuleError(f"配置中存在大小写重复的别名 {normalized_key}。")
        normalized[normalized_key] = target
    target = normalized.get(name, LEGACY_BINDINGS.get(name))
    if target is None:
        raise RuleError(f"函数别名 {name} 未配置，请在 config.xlsx 的“函数映射”页 A/B 列绑定实际算法。")
    if not isinstance(target, str):
        raise RuleError(f"别名 {name} 的目标必须是算法名称字符串。")
    definition = get_function(target.strip())
    if definition is None:
        raise RuleError(f"别名 {name} 对应的算法 {target} 尚未实现。")
    return definition


def parse_rule(value: object, current_sheet: str, coordinate: str,
               bindings: Mapping[str, str] | None = None) -> TemplateRule | None:
    if not isinstance(value, str):
        return None
    text = value.strip()
    if text.startswith("="):
        candidate = re.match(r"=\s*([A-Za-z_][A-Za-z_0-9]*)\s*[(（]", text)
        if candidate:
            name = candidate.group(1).upper()
            if is_alias(name) or (name != "SUM" and get_function(name) is not None):
                raise RuleError(f"{name} 是工具指令，请去掉开头的等号。")
        return None
    match = _FUNCTION.fullmatch(text)
    if match is None and not is_alias(text):
        candidate = re.match(r"([A-Za-z_][A-Za-z_0-9]*)(.*)", text, re.DOTALL)
        if (candidate and any(char in candidate.group(2) for char in "()（）")
                and (get_function(candidate.group(1)) is not None or is_alias(candidate.group(1)))):
            raise RuleError(f"模板 '{current_sheet}'!{coordinate}，原始指令“{value}”：汇总指令格式不完整，请使用成对的 () 或 （），并检查参数。")
        return None
    name = match.group(1).upper() if match is not None else text.upper()
    location = f"模板 '{current_sheet}'!{coordinate}，原始指令“{value}”，函数 {name}"
    try:
        arguments = _arguments(match.group(2)) if match is not None else [""]
        definition = _resolve_function(name, bindings)
        location += f" → {definition.name}"
        arguments = [argument or coordinate for argument in arguments]
        if not definition.accepts_count(len(arguments)):
            raise RuleError(f"期望 {definition.parameter_count} 个参数，实际 {len(arguments)} 个参数，请使用 {definition.signature(name)}。")
        references = tuple(parse_reference(argument, current_sheet) for argument in arguments)
        definition.validate_references(references)
    except ValueError as exc:
        raise RuleError(f"{location}：{exc}") from exc
    return TemplateRule(current_sheet, coordinate, definition.name, references, text, name)
