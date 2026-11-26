from collections import Counter
from datetime import datetime
from decimal import Decimal, localcontext
import os
from pathlib import Path
import tempfile
import time
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.workbook.properties import CalcProperties

from .configuration import ConfigurationError, load_configuration
from .functions import CalculationNotice, function_names, get_function, is_alias
from .models import AggregationCancelled, AggregationError, AggregationResult, Issue
from .reader import SourceReadError, SourceWorkbook
from .reports import RUN_LOG_SHEET, append_run_log
from .rules import RuleError, TemplateRule, parse_rule


def default_output_path(template_path, output_dir) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    return Path(output_dir) / f"{Path(template_path).stem}_汇总_{stamp}.xlsx"


def _identity(path: Path) -> str:
    return os.path.normcase(str(path.resolve()))


def _same_file(left: Path, right: Path) -> bool:
    if _identity(left) == _identity(right):
        return True
    return left.exists() and right.exists() and left.samefile(right)


def _fingerprint(path: Path):
    stat = path.stat()
    return stat.st_size, stat.st_mtime_ns


def _cancelled(cancel_event):
    if cancel_event is not None and cancel_event.is_set():
        raise AggregationCancelled()


def _check_template_content(path: Path):
    with ZipFile(path) as archive:
        for name in archive.namelist():
            lower = name.lower()
            unsupported = lower.startswith(("xl/media/", "xl/charts/", "xl/activex/", "xl/ctrlprops/", "_xmlsignatures/"))
            if lower.startswith("xl/drawings/") and lower.endswith(".xml"):
                try:
                    drawing = ET.fromstring(archive.read(name))
                    empty_drawing = drawing.tag == "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}wsDr" and not len(drawing)
                except ET.ParseError:
                    empty_drawing = False
                # 某些办公软件会留下空绘图部件；其中没有需要保留的对象。
                unsupported = unsupported or not empty_drawing
            if unsupported or lower == "xl/vbaproject.bin":
                raise AggregationError(
                    "模板包含图片、图表、绘图、控件、宏或签名。当前版本仅支持单元格表样，"
                    "为避免另存时丢失这些对象，本次未处理。请使用不含这些对象的模板。"
                )


def _read_rules(workbook, bindings=None) -> list[TemplateRule]:
    rules = []
    issues = []
    for sheet in workbook.worksheets:
        # 只访问已加载的单元格，避免稀疏模板被展开成巨大的空白矩形。
        cells = sorted(sheet._cells.values(), key=lambda cell: (cell.row, cell.column))
        for cell in cells:
            if isinstance(cell, MergedCell):
                continue
            if cell.data_type == "f":
                issues.append(Issue(
                    "error", "template_formula_not_supported",
                    f"模板含带等号的普通 Excel 公式“{cell.value}”。请使用不带等号的工具指令或 FUN_编号 别名。",
                    template_sheet=sheet.title, template_cell=cell.coordinate,
                ))
                continue
            if cell.data_type == "e":
                issues.append(Issue("error", "template_excel_error", f"模板含有 Excel 错误值：{cell.value}",
                                    template_sheet=sheet.title, template_cell=cell.coordinate))
                continue
            if sheet.title == RUN_LOG_SHEET:
                continue
            try:
                rule = parse_rule(cell.value, sheet.title, cell.coordinate, bindings=bindings)
            except RuleError as exc:
                issues.append(Issue("error", "invalid_rule", str(exc),
                                    template_sheet=sheet.title, template_cell=cell.coordinate))
                continue
            if rule is not None:
                rules.append(rule)
    if issues:
        raise AggregationError("模板规则校验失败，请修正指定位置后重试。", issues)
    if not rules:
        names = "、".join(function_names())
        raise AggregationError(f"模板中没有汇总指令。请使用不带等号的 {names}，或已配置的 FUN_编号 别名。")
    return rules


def _notice_issue(notice: CalculationNotice, rule: TemplateRule, source: Path | None = None, references=None) -> Issue:
    references = rule.references if references is None else references
    reference = None
    if notice.parameter_index is not None:
        if not 0 <= notice.parameter_index < len(references):
            raise AggregationError(f"计算函数 {rule.operation} 返回了无效的提醒参数位置。")
        reference = references[notice.parameter_index]
    return Issue(
        notice.severity, notice.code, notice.message, str(source) if source is not None else "",
        reference.sheet if source is not None and reference is not None else "",
        reference.coordinate if source is not None and reference is not None else "",
        rule.sheet, rule.coordinate,
    )


def _publish(workbook, output: Path, cancel_event):
    output.parent.mkdir(parents=True, exist_ok=True)
    handle, filename = tempfile.mkstemp(prefix=".汇总临时_", suffix=".xlsx", dir=output.parent)
    os.close(handle)
    temporary = Path(filename)
    try:
        workbook.save(temporary)
        with ZipFile(temporary) as archive:
            damaged = archive.testzip()
            if damaged is not None:
                raise AggregationError(f"结果文件完整性检查失败：{damaged}")
        _cancelled(cancel_event)
        if os.name == "nt":
            # Windows 的 rename 不会替换已存在的目标，避免检查与保存之间的竞争。
            os.rename(temporary, output)
        else:
            os.link(temporary, output)
    except FileExistsError as exc:
        raise AggregationError("输出文件已经存在，为保护已有文件，本次没有覆盖它。请使用新的文件名。") from exc
    finally:
        if temporary.exists():
            temporary.unlink()


def aggregate(template_path, source_paths, output_path, progress=None, cancel_event=None, *, config_path=None) -> AggregationResult:
    started = time.perf_counter()
    run_time = datetime.now()
    template = Path(template_path).expanduser().resolve()
    output = Path(output_path).expanduser().resolve()
    issues = []
    if template.suffix.lower() != ".xlsx" or not template.is_file():
        raise AggregationError("请选择实际存在的 .xlsx 模板文件。")
    if output.suffix.lower() != ".xlsx":
        raise AggregationError("结果文件扩展名必须为 .xlsx。")
    if _same_file(output, template):
        raise AggregationError("输出路径不能与模板相同。模板不会被覆盖。")
    if output.exists():
        raise AggregationError("输出文件已经存在，请换一个文件名；工具不会覆盖已有结果。")
    if output.parent.exists() and not output.parent.is_dir():
        raise AggregationError("输出目录不是一个文件夹。")
    sources = []
    seen = set()
    for candidate in source_paths:
        source = Path(candidate).expanduser().resolve()
        identity = _identity(source)
        if identity in seen or any(_same_file(source, existing) for existing in sources):
            issues.append(Issue("info", "duplicate_source", "重复选择的同一源文件已去重，只计算一次。", str(source)))
            continue
        seen.add(identity)
        if _same_file(source, template):
            raise AggregationError("模板不能同时作为待汇总源文件。")
        if _same_file(source, output):
            raise AggregationError("输出路径不能与源文件相同。")
        if source.name.startswith("~$"):
            issues.append(Issue("info", "lock_file", "Excel 临时锁文件已忽略。", str(source)))
            continue
        if source.suffix.lower() != ".xlsx" or not source.is_file():
            raise AggregationError(f"源文件不存在或不是 .xlsx：{source}")
        sources.append(source)
    if not sources:
        raise AggregationError("请至少选择一份有效的 .xlsx 源工作簿。")
    _cancelled(cancel_event)
    try:
        configuration = load_configuration(config_path)
    except ConfigurationError as exc:
        raise AggregationError(f"配置校验失败：{exc}") from exc
    try:
        template_fingerprint = _fingerprint(template)
        _check_template_content(template)
        workbook = load_workbook(template, data_only=False, keep_links=True)
    except AggregationError:
        raise
    except Exception as exc:
        raise AggregationError(f"模板打开失败：{exc}") from exc
    try:
        rules = _read_rules(workbook, configuration.bindings)
        template_sheet_count = sum(sheet.title != RUN_LOG_SHEET for sheet in workbook.worksheets)
        # 仅取数缓存去重；规则中的参数及各区域边界、重复项始终保留。
        references = list(dict.fromkeys(
            cell for rule in rules for reference in rule.references for cell in reference.cells()
        ))
        definitions = []
        calculators = []
        for rule in rules:
            definition = get_function(rule.operation)
            if definition is None:
                raise AggregationError(f"计算函数 {rule.operation} 未注册。")
            calculators.append(definition.factory())
            definitions.append(definition)
        source_states = {}
        source_sheet_count = 0
        if progress is not None:
            progress(0, len(sources) + 1, f"已识别 {len(rules)} 个汇总指令，开始读取 {len(sources)} 份源文件。")
        from .provenance import append_result_provenance
        with localcontext() as context:
            context.prec = 50
            for index, source in enumerate(sources):
                _cancelled(cancel_event)
                seen_errors = set()
                try:
                    source_states[source] = _fingerprint(source)
                    with SourceWorkbook(source, references) as reader:
                        for rule, definition, calculator in zip(rules, definitions, calculators):
                            _cancelled(cancel_event)
                            try:
                                for reference_row in definition.reference_rows(rule.references):
                                    _cancelled(cancel_event)
                                    values = [reader.get(reference) for reference in reference_row]
                                    for notice in calculator.add(values):
                                        issues.append(_notice_issue(notice, rule, source, reference_row))
                            except SourceReadError as exc:
                                error_key = (exc.code, exc.sheet.casefold(), exc.cell)
                                if error_key not in seen_errors:
                                    seen_errors.add(error_key)
                                    issues.append(Issue("error", exc.code, str(exc), str(source), exc.sheet, exc.cell,
                                                        rule.sheet, rule.coordinate))
                                continue
                        source_sheet_count += reader.read_sheet_count
                    if _fingerprint(source) != source_states[source]:
                        issues.append(Issue("error", "source_changed", "读取期间源文件发生变化，请关闭编辑后重试。", str(source)))
                except AggregationCancelled:
                    raise
                except Exception as exc:
                    issues.append(Issue("error", "open_failed", f"源文件处理失败：{exc}", str(source)))
                if progress is not None:
                    progress(index + 1, len(sources) + 1, f"已检查 {source.name}（{index + 1}/{len(sources)}）。")
            errors = [issue for issue in issues if issue.severity == "error"]
            if errors:
                raise AggregationError(f"发现 {len(errors)} 项错误，未生成汇总结果。请修正后重试。", issues)
            for rule, calculator in zip(rules, calculators):
                calculated = calculator.finish(len(sources))
                for notice in calculated.notices:
                    issues.append(_notice_issue(notice, rule))
                value = calculated.value
                if not value.is_finite() or abs(value) > Decimal("1.7976931348623157e308"):
                    raise AggregationError("计算结果超出 Excel 数值范围。", [Issue(
                        "error", "numeric_range", "结果不是可写入 Excel 的有限数值。",
                        template_sheet=rule.sheet, template_cell=rule.coordinate,
                    )])
                cell = workbook[rule.sheet][rule.coordinate]
                cell.value = value
                issues.extend(append_result_provenance(cell, rule, calculated, sources))
            errors = [issue for issue in issues if issue.severity == "error"]
            if errors:
                raise AggregationError(f"最终计算发现 {len(errors)} 项错误，未生成汇总结果。", issues)
        _cancelled(cancel_event)
        counts = dict(Counter(rule.operation for rule in rules))
        used_bindings = {
            rule.function_name: rule.operation for rule in rules
            if is_alias(rule.function_name)
        }
        result = AggregationResult(
            output_path=output, source_count=len(sources), rule_count=len(rules), issues=issues,
            duration_seconds=time.perf_counter() - started, source_sheet_count=source_sheet_count,
            template_sheet_count=template_sheet_count,
            sum_rule_count=counts.get("SUM", 0), wavg_rule_count=counts.get("WAVG", 0),
            avg_rule_count=counts.get("AVG", 0), function_counts=counts,
            function_bindings=used_bindings, config_path=configuration.path,
        )
        append_run_log(workbook, template, sources, run_time, result)
        workbook.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=False, forceFullCalc=False)
        if _fingerprint(template) != template_fingerprint:
            raise AggregationError("汇总期间模板发生变化，请保存好模板后重新运行。")
        for source, fingerprint in source_states.items():
            if _fingerprint(source) != fingerprint:
                raise AggregationError(f"汇总期间源文件发生变化，请重新运行：{source.name}")
        if progress is not None:
            progress(len(sources), len(sources) + 1, "计算完成，正在生成独立结果文件……")
        try:
            _publish(workbook, output, cancel_event)
        except (AggregationError, AggregationCancelled):
            raise
        except OSError as exc:
            raise AggregationError(f"结果保存失败，请检查目录权限或文件占用：{exc}") from exc
        result.duration_seconds = time.perf_counter() - started
        if progress is not None:
            progress(len(sources) + 1, len(sources) + 1, f"汇总完成：{output.name}")
        return result
    finally:
        workbook.close()
