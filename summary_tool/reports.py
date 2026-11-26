from . import __version__
from .functions import get_function
from .models import AggregationError, result_function_counts
from .templates.base import (
    align_text, auto_fit_columns, auto_fit_row_heights, font_body,
    setup_sheet, style_data_row, style_header_row,
)


RUN_LOG_SHEET = "运行日志"
TEXT_LIMIT = 32767


def split_log_text(value):
    if not isinstance(value, str) or len(value.encode("utf-16-le")) // 2 <= TEXT_LIMIT:
        return [value]
    result = []
    part = []
    units = 0
    for character in value:
        width = 2 if ord(character) > 65535 else 1
        if units + width > TEXT_LIMIT:
            result.append("".join(part))
            part = []
            units = 0
        part.append(character)
        units += width
    if part:
        result.append("".join(part))
    return result


def _expand_rows(rows):
    expanded = []
    for row in rows:
        parts = [split_log_text(value) for value in row]
        count = max(map(len, parts))
        for index in range(count):
            line = [values[index] if index < len(values) else "" for values in parts]
            if index:
                line[0], line[1] = row[0], row[1]
                if len(parts[2]) == 1:
                    label = f"{row[2]}（续 {index + 1}/{count}）"
                    line[2] = label if len(label.encode("utf-16-le")) // 2 <= TEXT_LIMIT else row[2]
            expanded.append(line)
    return expanded


def append_run_log(workbook, template, sources, run_time, result):
    stamp = run_time.strftime("%Y-%m-%d %H:%M:%S")
    counts = "；".join(f"{name} {count} 个单元格" for name, count in result_function_counts(result).items())
    config_path = getattr(result, "config_path", None)
    bindings = getattr(result, "function_bindings", {})
    rows = [
        [stamp, "信息", "程序版本", "", __version__],
        [stamp, "信息", "模板", template.name, str(template)],
        [stamp, "信息", "输出文件", result.output_path.name, str(result.output_path)],
        [stamp, "信息", "配置文件", "", str(config_path) if config_path is not None else "默认配置不存在，可使用直接函数及兼容旧别名。"],
        [stamp, "信息", "汇总统计", "", (
            f"源文件 {result.source_count} 个；参与源表 {result.source_sheet_count} 张；"
            f"输出模板页 {result.template_sheet_count} 张；{counts}。"
        )],
        [stamp, "信息", "计算耗时", "", f"{result.duration_seconds:.3f} 秒（不含结果保存）"],
        [stamp, "信息", "计算约定", "", "SUM/AVG/WAVG/SUMPRODUCT/RATIO空白按0；AVG分母为有效源文件总数；空白利率的权重仍进入加权分母；WAVG区域按位置配对且总权重为0报错；SUMPRODUCT先逐文件相乘再求和；RATIO为分子合计除以分母合计，分母合计0报错；MAX/MIN忽略空白、真实0有效，全空0不代表真实极值；COUNT空白不计、真实0计数；输出仅为结果值。"],
    ]
    for name, target in sorted(bindings.items()):
        rows.append([stamp, "信息", "本批函数绑定", name,
                     f"{name} -> {target}；算法签名：{get_function(target).signature()}。本批使用固定快照，配置修改从下一批生效。"])
    if not bindings:
        rows.append([stamp, "信息", "本批函数绑定", "", "本批未使用别名；直接函数名称的固定意义不受配置改变。"])
    rows.extend([stamp, "信息", "源文件", source.name, str(source)] for source in sources)
    labels = {"info": "信息", "warning": "提醒", "error": "错误"}
    for issue in result.issues:
        positions = []
        if issue.template_sheet:
            positions.append(f"模板 {issue.template_sheet}!{issue.template_cell}")
        if issue.source_sheet:
            positions.append(f"源 {issue.source_sheet}!{issue.source_cell}")
        rows.append([stamp, labels.get(issue.severity, issue.severity), "；".join(positions) or issue.code,
                     issue.source_file, issue.message])
    if not result.issues:
        rows.append([stamp, "信息", "检查结果", "", "本次未发现需要提醒的情况。"])
    rows = _expand_rows(rows)
    existing = RUN_LOG_SHEET in workbook.sheetnames
    sheet = workbook[RUN_LOG_SHEET] if existing else workbook.create_sheet(RUN_LOG_SHEET)
    has_content = bool(sheet._cells or sheet.merged_cells.ranges)
    start_row = sheet.max_row + 3 if has_content else 2
    header_row = start_row + 2
    if header_row + len(rows) > 1048576:
        raise AggregationError("运行日志没有足够的剩余行数，请减少单次文件或规则数量，或使用不含历史日志的模板。")
    if not has_content:
        setup_sheet(sheet, "运行日志", 6)
    else:
        cell = sheet.cell(start_row, 2, "本次运行日志")
        cell.font = font_body()
        cell.alignment = align_text()
    headers = ["时间", "级别", "项目／位置", "文件", "内容"]
    for column, header in enumerate(headers, 2):
        sheet.cell(header_row, column, header)
    style_header_row(sheet, header_row, 2, 6)
    for row_number, row in enumerate(rows, header_row + 1):
        for column, value in enumerate(row, 2):
            cell = sheet.cell(row_number, column, value)
            if isinstance(value, str):
                cell.data_type = "s"
        style_data_row(sheet, row_number, 2, 6, row_number - header_row - 1)
    if not has_content:
        auto_fit_columns(sheet, min_width=14, max_width=56, header_row=header_row, data_start_row=header_row + 1)
        sheet.column_dimensions["B"].width = 22
        sheet.column_dimensions["C"].width = 10
        sheet.column_dimensions["D"].width = 36
        sheet.column_dimensions["E"].width = 36
        sheet.column_dimensions["F"].width = 64
        sheet.freeze_panes = f"D{header_row + 1}"
    auto_fit_row_heights(sheet, header_row=header_row, data_start_row=header_row + 1,
                         data_end_row=header_row + len(rows))
    for row_number in range(header_row, header_row + len(rows) + 1):
        height = sheet.row_dimensions[row_number].height
        if height is not None and height > 409.5:
            sheet.row_dimensions[row_number].height = 409.5
    workbook.move_sheet(sheet, offset=len(workbook.worksheets) - 1 - workbook.worksheets.index(sheet))
