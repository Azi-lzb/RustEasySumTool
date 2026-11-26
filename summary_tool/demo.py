"""Create fictional v0.6 alias workbooks and an independent Decimal oracle.

``base_dir`` is the example-data directory itself, not its parent. Existing
files are left byte-for-byte intact unless ``force=True`` is explicitly passed.
Expected values describe the canonical synthetic inputs and default bindings,
not later user edits. Runtime bindings come from the application's config.xlsx;
no separate configuration is generated alongside the template.
This module never imports or runs the aggregation engine.
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, localcontext
from io import BytesIO
import json
from pathlib import Path
from typing import Callable, Iterable

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.worksheet import Worksheet

from .templates.base import (
    FORMATS,
    align_number,
    align_text,
    auto_fit_columns,
    auto_fit_row_heights,
    font_caption,
    setup_sheet,
    style_data_row,
    style_header_row,
    style_total_row,
)


DETAIL_SHEET = "业务明细"
LEDGER_SHEET = "利率台账"
PROJECT_NAMES = ("普通业务甲", "普通业务乙", "缺利率示例", "全空示例", "真实零利率")
HEADERS = ("项目", "存量余额", "存量利率", "发生额", "发生额利率")
TEMPLATE_HEADERS = (
    "项目", "余额指标", "余额利率指标", "发生额指标", "发生额利率指标", "余额参考指标",
    "余额极值A", "余额极值B", "有效填报数",
)
# Independent diagnostics never become original project records or source totals.
# Balance values are ordered as 机构甲、机构乙、机构丙, just like _RAW_SOURCES.
_DIAGNOSTIC_ROWS = (
    (13, "负数与空白", (Decimal("-5"), None, Decimal("-10"))),
    (14, "并列最大", (Decimal("7"), Decimal("7"), Decimal("3"))),
    (15, "真实0与空白", (Decimal("0"), None, Decimal("5"))),
)


@dataclass(frozen=True)
class DemoAssets:
    root: Path
    template_path: Path
    source_paths: tuple[Path, ...]
    expected_values: dict[tuple[str, str], Decimal]
    error_paths: tuple[Path, ...]


@dataclass(frozen=True)
class _Record:
    balance: Decimal | None
    balance_rate: Decimal | None
    flow: Decimal | None
    flow_rate: Decimal | None
    ledger_balance_rate: Decimal | None
    ledger_flow_rate: Decimal | None


@dataclass(frozen=True)
class _Source:
    name: str
    rows: tuple[_Record, ...]


def _row(*values: str | None) -> _Record:
    return _Record(*(Decimal(value) if value is not None else None for value in values))


# These raw records are the sole inputs to both fixture generation and the
# independent Decimal calculation below. None means blank; Decimal("0") is real 0%.
_RAW_SOURCES = (
    _Source("机构甲", (
        _row("1000000", "0.03", "2000000", "0.04", "0.02", "0.03"),
        _row("2000000", "0.04", "1000000", "0.04", "0.03", "0.04"),
        _row("1000000", None, "2000000", None, None, None),
        _row(None, None, None, None, None, None),
        _row("1000000", "0", "2000000", "0", "0", "0"),
    )),
    _Source("机构乙", (
        _row("9000000", "0.05", "3000000", "0.06", "0.04", "0.05"),
        _row("3000000", "0.05", "3000000", "0.04", "0.05", "0.06"),
        _row("3000000", "0.04", "3000000", "0.06", "0.05", "0.06"),
        _row(None, None, None, None, None, None),
        _row("1000000", "0.02", "2000000", "0.04", "0.04", "0.06"),
    )),
    _Source("机构丙", (
        _row(None, None, "5000000", "0.02", None, "0.07"),
        _row("5000000", "0.06", "6000000", "0.05", "0.07", "0.08"),
        _row(None, None, "1000000", "0.02", None, "0.10"),
        _row(None, None, None, None, None, None),
        _row("2000000", "0", None, None, "0", None),
    )),
)


def _sum(values: Iterable[Decimal | None]) -> Decimal:
    return sum((value for value in values if value is not None), Decimal(0))


def _weighted(pairs: Iterable[tuple[Decimal | None, Decimal | None]]) -> Decimal:
    # Contract: blank rates are 0%, but their amounts remain in the denominator.
    # Decimal(0) rates are equally numeric, without the missing-rate warning.
    valid = [(Decimal(0) if rate is None else rate, weight)
             for rate, weight in pairs if weight is not None]
    denominator = sum((weight for _, weight in valid), Decimal(0))
    if denominator == 0:
        return Decimal(0)
    return sum((rate * weight for rate, weight in valid), Decimal(0)) / denominator


def _balance_statistics(values: Iterable[Decimal | None]) -> tuple[Decimal, Decimal, Decimal]:
    """Independent oracle: blanks are absent, while a real numeric zero counts."""
    valid = [value for value in values if value is not None]
    return max(valid, default=Decimal(0)), min(valid, default=Decimal(0)), Decimal(len(valid))


def _expected_values() -> dict[tuple[str, str], Decimal]:
    """Calculate all 61 values from raw records and independent diagnostics."""
    values: dict[tuple[str, str], Decimal] = {}
    with localcontext() as context:
        context.prec = 40
        source_count = Decimal(len(_RAW_SOURCES))
        for offset in range(len(PROJECT_NAMES)):
            row = offset + 5
            records = [source.rows[offset] for source in _RAW_SOURCES]
            values[DETAIL_SHEET, f"C{row}"] = _sum(record.balance for record in records)
            values[DETAIL_SHEET, f"E{row}"] = _sum(record.flow for record in records)
            if row != 8:
                values[DETAIL_SHEET, f"D{row}"] = _weighted((r.balance_rate, r.balance) for r in records)
                values[DETAIL_SHEET, f"F{row}"] = _weighted((r.flow_rate, r.flow) for r in records)
            # All institutions participate: a missing balance contributes 0,
            # without reducing the denominator to the nonblank count.
            values[DETAIL_SHEET, f"G{row}"] = _sum(record.balance for record in records) / source_count
            for column, value in zip("HIJ", _balance_statistics(record.balance for record in records)):
                values[DETAIL_SHEET, f"{column}{row}"] = value
        # Independently weight every original record, not the rounded row or
        # per-source totals written into workbooks. No engine output is read.
        records = [record for source in _RAW_SOURCES for record in source.rows]
        values[DETAIL_SHEET, "C10"] = _sum(record.balance for record in records)
        values[DETAIL_SHEET, "D10"] = _weighted((record.balance_rate, record.balance) for record in records)
        values[DETAIL_SHEET, "E10"] = _sum(record.flow for record in records)
        values[DETAIL_SHEET, "F10"] = _weighted((record.flow_rate, record.flow) for record in records)
        values[DETAIL_SHEET, "G10"] = _sum(record.balance for record in records) / source_count
        # H/I/J10 compare independently summed institution balances, not the
        # largest/smallest original project and not generated workbook totals.
        source_totals = [_sum(record.balance for record in source.rows) for source in _RAW_SOURCES]
        for column, value in zip("HIJ", _balance_statistics(source_totals)):
            values[DETAIL_SHEET, f"{column}10"] = value
        for row, _, balances in _DIAGNOSTIC_ROWS:
            values[DETAIL_SHEET, f"C{row}"] = _sum(balances)
            values[DETAIL_SHEET, f"G{row}"] = _sum(balances) / source_count
            for column, value in zip("HIJ", _balance_statistics(balances)):
                values[DETAIL_SHEET, f"{column}{row}"] = value
    return values


def _new_workbook(title: str) -> Workbook:
    workbook = Workbook()
    workbook.properties.creator = "Z.ai"
    workbook.properties.title = title
    workbook.properties.subject = "纯虚构中文汇总工具测试数据"
    workbook.properties.description = "所有机构、项目和金额均为虚构，仅供测试，不对应任何真实业务。"
    workbook.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)
    return workbook


def _header(ws: Worksheet, headers: tuple[str, ...]) -> None:
    for column, text in enumerate(headers, start=2):
        ws.cell(4, column, text)
    style_header_row(ws, 4, 2, len(headers) + 1)
    ws.auto_filter.ref = f"B4:{get_column_letter(len(headers) + 1)}9"


def _numeric_cells(ws: Worksheet, row: int, amount_columns: tuple[int, ...], rate_columns: tuple[int, ...],
                   count_columns: tuple[int, ...] = ()) -> None:
    for column in amount_columns:
        ws.cell(row, column).number_format = FORMATS["decimal_2"]
        ws.cell(row, column).alignment = align_number()
    for column in rate_columns:
        ws.cell(row, column).number_format = "0.0000%"
        ws.cell(row, column).alignment = align_number()
    for column in count_columns:
        ws.cell(row, column).number_format = "0"
        ws.cell(row, column).alignment = align_number()


def _caption(ws: Worksheet, row: int, last_col: int, text: str) -> None:
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=last_col)
    cell = ws.cell(row, 2, text)
    cell.font = font_caption()
    cell.alignment = align_text()


def _finish(ws: Worksheet, last_col: str, last_row: int, *, wide: bool = False) -> None:
    auto_fit_columns(ws, min_width=14, max_width=38 if wide else 24, header_row=4, data_start_row=5)
    ws.column_dimensions["B"].width = max(ws.column_dimensions["B"].width, 18)
    auto_fit_row_heights(ws, header_row=4, data_start_row=5, data_end_row=last_row)
    ws.print_area = f"B2:{last_col}{last_row}"


def _make_source(source: _Source, error: str | None = None) -> Workbook:
    workbook = _new_workbook(f"{source.name}报送数据（纯虚构）")
    details = workbook.active
    details.title = DETAIL_SHEET
    setup_sheet(details, f"{source.name}业务明细（纯虚构）", 6)
    _header(details, HEADERS)
    for offset, record in enumerate(source.rows):
        row = offset + 5
        details.cell(row, 2, PROJECT_NAMES[offset])
        for column, value in enumerate((record.balance, record.balance_rate, record.flow, record.flow_rate), start=3):
            details.cell(row, column, value)
        style_data_row(details, row, 2, 6, offset)
        _numeric_cells(details, row, (3, 5), (4, 6))
    details["B10"] = "合计"
    with localcontext() as context:
        context.prec = 40
        totals = (
            _sum(record.balance for record in source.rows),
            _weighted((record.balance_rate, record.balance) for record in source.rows),
            _sum(record.flow for record in source.rows),
            _weighted((record.flow_rate, record.flow) for record in source.rows),
        )
    for column, value in enumerate(totals, start=3):
        details.cell(10, column, value)
    style_total_row(details, 10, 2, 6)
    _numeric_cells(details, 10, (3, 5), (4, 6))
    details["D10"].comment = Comment("本机构第5至9行独立计算的数值合计；空白利率按0%，权重保留，不使用Excel公式。", "Z.ai")
    source_index = _RAW_SOURCES.index(source)
    for offset, (row, label, balances) in enumerate(_DIAGNOSTIC_ROWS):
        details.cell(row, 2, label)
        details.cell(row, 3, balances[source_index])
        style_data_row(details, row, 2, 3, offset)
        _numeric_cells(details, row, (3,), ())
    _caption(details, 18, 6, "纯虚构数据，金额单位元；第7行缺利率，第8行全空，第9行含真实0%；第10行已直接提供数值合计。")
    _caption(details, 20, 6, "第13至15行为独立诊断，仅C列提供余额源值，不参加第10行原项目合计；空白不是真实0，真实0是有效填报。")
    _finish(details, "F", 20)

    ledger = workbook.create_sheet(LEDGER_SHEET)
    setup_sheet(ledger, f"{source.name}利率台账（纯虚构）", 4)
    _header(ledger, ("项目", "存量利率", "发生额利率"))
    for offset, record in enumerate(source.rows):
        row = offset + 5
        ledger.cell(row, 2, PROJECT_NAMES[offset])
        ledger.cell(row, 3, record.ledger_balance_rate)
        ledger.cell(row, 4, record.ledger_flow_rate)
        style_data_row(ledger, row, 2, 4, offset)
        _numeric_cells(ledger, row, (), (3, 4))
    _caption(ledger, 12, 4, "纯虚构独立台账；本版模板未引用此表，不应计入实际参与源表数量。")
    _finish(ledger, "D", 12)

    if error == "缺工作表":
        workbook.remove(details)
        _caption(ledger, 14, 4, "异常示例：故意缺少“业务明细”工作表，汇总必须失败，不能将缺表视为0。")
        _finish(ledger, "D", 14)
    elif error == "非法数值":
        details["C5"] = "这不是合法金额"
        details["C5"].comment = Comment("异常示例：非空文字必须报错，不可按0处理。", "Z.ai")
        _finish(details, "F", 20)
    elif error == "公式无缓存":
        details["C5"] = "=SUM(500000,500000)"
        details["C5"].comment = Comment("异常示例：故意不重算，不提供公式缓存；汇总必须失败。", "Z.ai")
        _finish(details, "F", 20)
    elif error is not None:
        raise ValueError(f"未知异常示例：{error}")
    return workbook


def _make_template() -> Workbook:
    workbook = _new_workbook("v0.6别名汇总模板（纯虚构）")
    details = workbook.active
    details.title = DETAIL_SHEET
    setup_sheet(details, "业务汇总模板（纯虚构，v0.6别名）", 10)
    _header(details, TEMPLATE_HEADERS)
    for offset, project in enumerate(PROJECT_NAMES):
        row = offset + 5
        details.cell(row, 2, project)
        for column, rule in ((3, f"FUN_1(C{row})"), (4, f"FUN_2(D{row},C{row})"),
                             (5, f"FUN_1(E{row})"), (6, f"FUN_2(F{row},E{row})"),
                             (7, f"FUN_3(C{row})"), (8, f"FUN_4(C{row})"),
                             (9, f"FUN_5(C{row})"), (10, f"FUN_6(C{row})")):
            details.cell(row, column, rule)
        style_data_row(details, row, 2, 10, offset)
        _numeric_cells(details, row, (3, 5, 7, 8, 9), (4, 6), (10,))
    # Keep the original short comments intact; aggregation must append any
    # generated provenance rather than replace a template author's notes.
    for coordinate in ("D8", "F8"):
        details[coordinate] = "无有效权重"
        details[coordinate].comment = Comment("此示例的所有源权重为空，故不设置WAVG指令。WAVG总权重为0时会报错。", "Z.ai")
    details["C5"].comment = Comment(
        "工具指令不加等号，从各源文件的同名表、指定坐标取数。默认绑定FUN_1=SUM、FUN_3=AVG、FUN_2=WAVG；"
        "运行时仅由主程序同目录config.xlsx的函数映射页配置，模板目录不另放配置副本。"
        "切换绑定后须核对表头、单位和数值格式。", "Z.ai")
    details["D7"].comment = Comment("FUN_2默认绑定WAVG。有金额但缺利率的机构须产生提醒；空白利率按0%，其权重仍计入该项加权分母。", "Z.ai")
    details["G5"].comment = Comment("FUN_3默认绑定AVG：取三家机构同名表C列余额之和除以3；空白余额按0，三家全部计入分母，不按非空个数计算。切换绑定后须核对表头和单位。", "Z.ai")
    details["B10"] = "合计"
    for column, rule in ((3, "FUN_1(C10)"), (4, "FUN_2(D10,C10)"),
                         (5, "FUN_1(E10)"), (6, "FUN_2(F10,E10)"), (7, "FUN_3(C10)"),
                         (8, "FUN_4(C10)"), (9, "FUN_5(C10)"), (10, "FUN_6(C10)")):
        details.cell(10, column, rule)
    style_total_row(details, 10, 2, 10)
    _numeric_cells(details, 10, (3, 5, 7, 8, 9), (4, 6), (10,))
    details["D10"].comment = Comment("FUN_2默认绑定WAVG：从各机构源表第10行的数值合计加权；空白利率按0%，权重保留，不使用Excel公式。", "Z.ai")
    details["G10"].comment = Comment("FUN_3默认绑定AVG：取各机构源表C10数值合计的算术均值；源表无需增加均值字段，不使用Excel公式。", "Z.ai")
    for offset, (row, label, _) in enumerate(_DIAGNOSTIC_ROWS):
        details.cell(row, 2, label)
        for column, alias in ((3, "FUN_1"), (7, "FUN_3"), (8, "FUN_4"), (9, "FUN_5"), (10, "FUN_6")):
            details.cell(row, column, f"{alias}(C{row})")
        style_data_row(details, row, 2, 3, offset)
        style_data_row(details, row, 7, 10, offset)
        _numeric_cells(details, row, (3, 7, 8, 9), (), (10,))
    short_comments = {
        "H5": "FUN_4默认绑定MAX：忽略空白，真实0有效；结果需追加来源文件与坐标。",
        "I9": "FUN_5默认绑定MIN：此行机构甲、机构乙并列，结果需保留全部来源。",
        "J5": "FUN_6默认绑定COUNT：统计非空有效数值个数，真实0有效；整数显示。",
        "H10": "比较各机构原项目C10数值合计，不比较单个项目，不含独立诊断。",
        "I10": "比较各机构原项目C10数值合计，不含第13至15行独立诊断。",
        "H8": "原项目全空示例；按约定输出0，但不存在有效极值来源。",
        "I8": "原项目全空示例；空白不能成为值为0的极值来源。",
        "H13": "独立诊断：负数与空白，空白不能当作0参与极值比较。",
        "I13": "独立诊断：仅比较实际报送的负数。",
        "H14": "独立诊断：并列最大，须列出机构甲、机构乙全部来源。",
        "I15": "独立诊断：真实0有效，机构乙空白不应列为并列来源。",
    }
    for coordinate, text in short_comments.items():
        details[coordinate].comment = Comment(text, "Z.ai")
    _caption(details, 18, 10, "全部数据纯属虚构。第5至10行与第13至15行共61个不带等号的函数别名指令；绑定仅由主程序同目录config.xlsx配置，模板目录不另放配置副本；切换后须核对表头、单位和格式。默认金额单位元，结果为最终数值，无需Excel重算。")
    _caption(details, 20, 10, "第13至15行为独立诊断，不参加第10行原项目合计，D/E/F留空。默认FUN_4=MAX、FUN_5=MIN、FUN_6=COUNT；极值忽略空白，真实0有效；全空极值为0且无来源，有效填报数为0。")
    _finish(details, "J", 20)
    return workbook


def _write_bytes(path: Path, payload: bytes, *, force: bool) -> None:
    """Exclusive creation closes the existence-check race in the safe default."""
    try:
        with path.open("wb" if force else "xb") as stream:
            stream.write(payload)
    except FileExistsError:
        if force or not path.is_file():
            raise


def _save_workbook(path: Path, factory: Callable[[], Workbook], *, force: bool) -> None:
    if not force and path.is_file():
        return
    workbook = factory()
    try:
        with BytesIO() as buffer:
            workbook.save(buffer)
            payload = buffer.getvalue()
        _write_bytes(path, payload, force=force)
    finally:
        workbook.close()


def _manifest(values: dict[tuple[str, str], Decimal]) -> dict[str, dict[str, str]]:
    nested: dict[str, dict[str, str]] = {}
    for (sheet, coordinate), value in values.items():
        nested.setdefault(sheet, {})[coordinate] = str(value)
    return nested


def create_demo_assets(base_dir: Path, *, force: bool = False) -> DemoAssets:
    """Create a template, three good sources, three bad sources and an oracle.

    ``error_paths`` is ordered as missing-sheet, invalid-number, missing-cache.
    ``expected_values`` contains all 61 numeric targets: the original 28,
    18 project/total extremes and counts, plus 15 independent diagnostic values.
    The template has one sheet and only no-equals function-alias instructions;
    normal sources include independently calculated numeric totals, no formulas.
    Diagnostics never participate in the original project totals at row 10.
    The eight inputs include no separate configuration; expected values assume
    FUN_1=SUM, FUN_3=AVG, FUN_2=WAVG, FUN_4=MAX, FUN_5=MIN and FUN_6=COUNT.
    No aggregation result file is produced by this function.
    """
    root = Path(base_dir).expanduser().resolve()
    normal_dir, error_dir = root / "正常数据", root / "异常数据"
    normal_dir.mkdir(parents=True, exist_ok=True)
    error_dir.mkdir(parents=True, exist_ok=True)
    template_path = root / "汇总模板.xlsx"
    source_paths = tuple(normal_dir / f"{source.name}.xlsx" for source in _RAW_SOURCES)
    error_names = ("缺工作表", "非法数值", "公式无缓存")
    error_paths = tuple(error_dir / f"{name}.xlsx" for name in error_names)
    _save_workbook(template_path, _make_template, force=force)
    for path, source in zip(source_paths, _RAW_SOURCES):
        _save_workbook(path, lambda source=source: _make_source(source), force=force)
    for path, name in zip(error_paths, error_names):
        _save_workbook(path, lambda name=name: _make_source(_RAW_SOURCES[0], error=name), force=force)
    values = _expected_values()
    manifest = {
        "version": "0.6",
        "result_count": len(values),
        "说明": "全部数据纯属虚构。以下61格（含原项目合计及独立诊断，默认绑定下SUM15/AVG9/WAVG10/MAX9/MIN9/COUNT9）是初始样例的独立Decimal预期值；原28格数值保持不变，新增33格。D8/F8显示无有效权重。不适用于后来修改过的源文件或切换过的函数绑定。",
        "别名约定": "默认绑定FUN_1=SUM、FUN_3=AVG、FUN_2=WAVG、FUN_4=MAX、FUN_5=MIN、FUN_6=COUNT；运行时仅由主程序同目录config.xlsx的函数映射页配置，模板目录不另放配置副本。切换绑定后须核对表头、单位和数值格式。",
        "预期值来源": "仅从生成器内的原始记录及独立诊断values按Decimal独立计算；不读取汇总引擎规则、引擎结果或Excel公式缓存，不从engine反推，仅对应上述默认绑定。H/I/J10由每机构原始记录先独立求和后比较或计数，不读取源表已写出的合计。",
        "function_bindings": {"FUN_1": "SUM", "FUN_3": "AVG", "FUN_2": "WAVG",
                              "FUN_4": "MAX", "FUN_5": "MIN", "FUN_6": "COUNT"},
        "function_counts": {"SUM": 15, "AVG": 9, "WAVG": 10, "MAX": 9, "MIN": 9, "COUNT": 9},
        "加权约定": "空白利率按0%，对应权重仍进入分母且须提醒；真实零利率有效且不提醒；总权重为0报错，故D8/F8仅显示无有效权重。",
        "算术均值约定": "FUN_3默认绑定AVG，取三家机构C列余额之和除以3；空白余额按0，三家全部计入分母，不按非空个数计算；G10取各机构C10数值合计的均值。",
        "极值约定": "FUN_4默认MAX、FUN_5默认MIN；忽略None与空白，真实0有效；保留全部并列来源。全空按约定输出0，无有效数值、无极值来源。H/I列金额两位小数。",
        "计数约定": "FUN_6默认COUNT，只统计非空有效数值个数，真实0有效；全空为0；J列数值格式为整数0。",
        "诊断约定": "业务明细第13至15行为独立诊断，不参加第10行原项目合计；每行仅C/G/H/I/J五个目标，D/E/F留空。源表仅C13:C15增加余额数值，机构顺序为甲/乙/丙。常规项目筛选止于第9行，不包含合计与诊断。",
        "diagnostics": {
            str(row): {"label": label, "balances": [str(value) if value is not None else None for value in balances]}
            for row, label, balances in _DIAGNOSTIC_ROWS
        },
        "expected_values": _manifest(values),
    }
    _write_bytes(root / "预期结果.json", (json.dumps(manifest, ensure_ascii=False, indent=2) + "\n").encode("utf-8"), force=force)
    return DemoAssets(root, template_path, source_paths, values, error_paths)


__all__ = ["DemoAssets", "create_demo_assets"]
