from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
import posixpath
import re
from urllib.parse import unquote
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from python_calamine import CalamineWorkbook

from .rules import CellReference


_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_NUMBER = re.compile(
    r"[+-]?(?:(?:[0-9]{1,3}(?:,[0-9]{3})+|[0-9]+)(?:\.[0-9]*)?|\.[0-9]+)"
    r"(?:[eE][+-]?[0-9]+)?"
)


class SourceReadError(ValueError):
    def __init__(self, message: str, code: str = "invalid_number", sheet: str = "", cell: str = ""):
        super().__init__(message)
        self.code = code
        self.sheet = sheet
        self.cell = cell


@dataclass(frozen=True)
class NumberValue:
    value: Decimal
    blank: bool


def decimal_value(value: object) -> NumberValue:
    if value is None:
        return NumberValue(Decimal(0), True)
    if isinstance(value, bool):
        raise SourceReadError("布尔值不是金额或利率，请填写数字。")
    if isinstance(value, str):
        text = value.replace("\u200b", "").replace("\ufeff", "").strip()
        if not text:
            return NumberValue(Decimal(0), True)
        percent = text.endswith(("%", "％"))
        if percent:
            text = text[:-1].strip()
        if _NUMBER.fullmatch(text) is None:
            raise SourceReadError(f"无法将“{value[:100]}”解析成数字；非空文字不能按 0 处理。")
        try:
            number = Decimal(text.replace(",", ""))
        except InvalidOperation as exc:
            raise SourceReadError("数字格式无效。") from exc
        if percent:
            number /= 100
    elif isinstance(value, (int, float, Decimal)):
        number = Decimal(str(value))
    else:
        raise SourceReadError(f"不支持的数值类型：{type(value).__name__}，请检查是否引用到了日期或表头。")
    if not number.is_finite():
        raise SourceReadError("无穷大或 NaN 不是有效数字。")
    if number and abs(number.adjusted()) > 1000:
        raise SourceReadError("数字的指数过大，无法安全计算。")
    return NumberValue(number, False)


@dataclass(frozen=True)
class _CellInfo:
    error: str = ""
    formula: bool = False
    cache_present: bool = True


class SourceWorkbook:
    """calamine 读取数值，OOXML 只用于区分空白、错误值和无缓存公式。"""

    def __init__(self, path: Path, references: list[CellReference]):
        self.path = path
        self._workbook = None
        self._archive = None
        self._values = {}
        self._metadata = {}
        self._numbers = {}
        self._required = {}
        for reference in references:
            self._required.setdefault(reference.sheet.casefold(), set()).add(reference.coordinate)
        try:
            self._archive = ZipFile(path)
            self._sheet_parts = self._read_sheet_parts()
            self._workbook = CalamineWorkbook.from_path(str(path))
            self._sheet_names = {name.casefold(): name for name in self._workbook.sheet_names}
        except Exception:
            self.close()
            raise

    def _read_sheet_parts(self) -> dict[str, str]:
        workbook = ET.fromstring(self._archive.read("xl/workbook.xml"))
        relations = ET.fromstring(self._archive.read("xl/_rels/workbook.xml.rels"))
        targets = {}
        for relation in relations:
            if relation.get("TargetMode") == "External":
                continue
            target = unquote(relation.get("Target", ""))
            part = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
            targets[relation.get("Id")] = part
        parts = {}
        sheets = workbook.find(f"{_MAIN}sheets")
        if sheets is None:
            raise SourceReadError("工作簿中没有可读取的工作表。", "invalid_workbook")
        for sheet in sheets:
            target = targets.get(sheet.get(f"{_REL}id"))
            if target:
                parts[sheet.get("name", "").casefold()] = target
        return parts

    def _inspect_cells(self, name: str) -> dict[str, _CellInfo]:
        if name in self._metadata:
            return self._metadata[name]
        part = self._sheet_parts.get(name)
        if part is None:
            raise SourceReadError("找不到工作表的内部数据，文件可能已损坏。", "invalid_workbook")
        wanted = self._required.get(name, set())
        found = {}
        with self._archive.open(part) as stream:
            for _, element in ET.iterparse(stream, events=("end",)):
                if element.tag == f"{_MAIN}c":
                    coordinate = element.get("r", "").upper()
                    if coordinate in wanted:
                        formula = element.find(f"{_MAIN}f") is not None
                        cached = element.find(f"{_MAIN}v")
                        kind = element.get("t", "n")
                        present = cached is not None and (cached.text is not None or kind == "str")
                        found[coordinate] = _CellInfo(
                            error=((cached.text if cached is not None else None) or "未知 Excel 错误") if kind == "e" else "",
                            formula=formula,
                            cache_present=present,
                        )
                    element.clear()
                elif element.tag == f"{_MAIN}row":
                    element.clear()
        self._metadata[name] = found
        return found

    def _sheet_values(self, name: str):
        if name not in self._values:
            sheet = self._workbook.get_sheet_by_name(self._sheet_names[name])
            # 不补齐前导空白区域，用 range 的起点还原 Excel 绝对地址。
            self._values[name] = (sheet.start, sheet.to_python())
        return self._values[name]

    def get(self, reference: CellReference) -> NumberValue:
        name = reference.sheet.casefold()
        key = (name, reference.coordinate)
        if key in self._numbers:
            return self._numbers[key]
        if name not in self._sheet_names:
            raise SourceReadError("找不到工作表，不能按空白 0 处理。", "missing_sheet", reference.sheet, reference.coordinate)
        try:
            metadata = self._inspect_cells(name).get(reference.coordinate)
            if metadata is not None:
                if metadata.error:
                    raise SourceReadError(f"源单元格包含 Excel 错误：{metadata.error}。", "excel_error")
                if metadata.formula and not metadata.cache_present:
                    raise SourceReadError(
                        "源单元格是公式，但没有有效的计算缓存。请用 Excel 打开源文件、完成重算并保存后重试。",
                        "formula_cache_missing",
                    )
            start, rows = self._sheet_values(name)
            raw = None
            if start is not None:
                row_index = reference.row - 1 - start[0]
                column_index = reference.column - 1 - start[1]
                if 0 <= row_index < len(rows) and 0 <= column_index < len(rows[row_index]):
                    raw = rows[row_index][column_index]
            result = decimal_value(raw)
            self._numbers[key] = result
            return result
        except SourceReadError as exc:
            if not exc.sheet:
                exc.sheet = reference.sheet
                exc.cell = reference.coordinate
            raise
        except Exception as exc:
            raise SourceReadError(f"读取工作表失败：{exc}", "read_failed", reference.sheet, reference.coordinate) from exc

    @property
    def read_sheet_count(self) -> int:
        return len(self._values)

    def close(self):
        if self._workbook is not None:
            close = getattr(self._workbook, "close", None)
            if close is not None:
                close()
            self._workbook = None
        if self._archive is not None:
            self._archive.close()
            self._archive = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()
