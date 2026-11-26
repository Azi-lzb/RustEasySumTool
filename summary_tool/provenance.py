from collections.abc import Iterator, Sequence
from copy import copy
import os
from pathlib import Path

from openpyxl.cell.cell import Cell
from openpyxl.comments import Comment

from .functions import FunctionResult
from .models import AggregationError, Issue
from .rules import CellReference, TemplateRule


_COMMENT_LIMIT = 32767
_PREVIEW_LIMIT = 10
_SECTION_HEADER = "—— 汇总结果来源 ——"
_PreviewSource = tuple[int, Path, CellReference]


def _utf16_length(text: str) -> int:
    return len(text.encode("utf-16-le", errors="surrogatepass")) // 2


def _abbreviate(text: str, limit: int) -> str:
    if _utf16_length(text) <= limit:
        return text
    suffix = "…（完整内容见运行日志）"
    remaining = limit - _utf16_length(suffix)
    prefix = []
    for character in text:
        size = 2 if ord(character) > 0xFFFF else 1
        if size > remaining:
            break
        prefix.append(character)
        remaining -= size
    return "".join(prefix) + suffix


def _algorithm_label(rule: TemplateRule) -> str:
    if rule.function_name and rule.function_name != rule.operation:
        return f"{rule.operation}（原函数/别名：{rule.function_name}）"
    return rule.operation


def _address(reference: CellReference) -> str:
    return f"'{reference.sheet.replace(chr(39), chr(39) * 2)}'!{reference.coordinate}"


def _checked_index(value: object, count: int, label: str, rule: TemplateRule) -> int:
    if type(value) is not int or not 0 <= value < count:
        message = (
            f"计算函数 {rule.operation} 返回了无效的来源{label} {value!r}，"
            f"必须是非布尔整数，且满足 0 <= 索引 < {count}。"
        )
        raise AggregationError(message, [Issue(
            "error", "invalid_result_source", message,
            template_sheet=rule.sheet, template_cell=rule.coordinate,
        )])
    return value


def _comment_candidates(
    algorithm: str, preview: list[_PreviewSource], total: int, note: str,
) -> Iterator[str]:
    if total > 1:
        summary = f"并列共{total}个，完整来源见运行日志"
    elif total:
        summary = "来源共1个，完整来源见运行日志"
    else:
        summary = "无有效数值，完整说明见运行日志"

    lines = [_SECTION_HEADER, f"算法：{algorithm}"]
    for ordinal, (source_index, path, reference) in enumerate(preview, 1):
        lines.extend((
            f"来源 {ordinal}/{total}（源文件序号 {source_index + 1}）：{path.name}",
            f"表格地址：{_address(reference)}",
            f"完整路径：{path}",
        ))
    if not total:
        lines.append("无有效数值，未记录来源文件。")
    if note:
        lines.append(f"说明：{note}")
    if total > len(preview):
        lines.append(summary)
    yield "\n".join(lines)

    # 只缩略新生成的预览；完整路径及说明始终留在 Issue/计算提醒中。
    compact_algorithm = _abbreviate(algorithm, 160)
    lines = [_SECTION_HEADER, f"算法：{compact_algorithm}"]
    for ordinal, (source_index, path, reference) in enumerate(preview, 1):
        lines.extend((
            f"来源 {ordinal}/{total}（源文件序号 {source_index + 1}）：{_abbreviate(path.name, 160)}",
            f"表格地址：{_abbreviate(_address(reference), 160)}",
            "完整路径：批注空间有限，见运行日志",
        ))
    if note:
        lines.append(f"说明：{_abbreviate(note, 512)}")
    lines.append(summary)
    yield "\n".join(lines)
    yield "\n".join((_SECTION_HEADER, f"算法：{compact_algorithm}", summary))
    yield "\n".join((_SECTION_HEADER, summary))
    hint = "完整来源见运行日志" if total else "无有效数值，完整说明见运行日志"
    yield f"——\n{hint}"


def append_result_provenance(
    cell: Cell, rule: TemplateRule, calculated: FunctionResult, source_paths: Sequence[Path],
) -> list[Issue]:
    """将结果自带的来源元数据追加到输出单元格备注，并返回完整日志条目。"""
    sources, note = calculated.sources, calculated.source_note
    if not sources and not note:
        return []

    algorithm = _algorithm_label(rule)
    total = len(sources)
    issues = []
    preview = []
    for ordinal, source in enumerate(sources, 1):
        source_index = _checked_index(
            getattr(source, "source_index", None), len(source_paths), "文件索引", rule,
        )
        parameter_index = _checked_index(
            getattr(source, "parameter_index", None), len(rule.references), "参数索引", rule,
        )
        # 词法绝对化即可；不打开源文件，也不要求超长路径在当前磁盘存在。
        path = Path(os.path.abspath(Path(source_paths[source_index]).expanduser()))
        reference = rule.references[parameter_index]
        issues.append(Issue(
            "info", "result_source",
            f"{algorithm} 结果来源 {ordinal}/{total}"
            f"（源文件序号 {source_index + 1}，参数序号 {parameter_index + 1}）。",
            str(path), reference.sheet, reference.coordinate, rule.sheet, rule.coordinate,
        ))
        if ordinal <= _PREVIEW_LIMIT:
            preview.append((source_index, path, reference))

    if note and not any(notice.message == note for notice in calculated.notices):
        issues.append(Issue(
            "info", "result_source_note", note,
            template_sheet=rule.sheet, template_cell=rule.coordinate,
        ))

    # 全部索引校验及日志构造成功后才接触批注，避免错误元数据留下半份预览。
    original = cell.comment
    original_text = original.text if original is not None else ""
    prefix = original_text + "\n\n" if original_text else ""
    available = _COMMENT_LIMIT - _utf16_length(prefix)
    section = next((
        candidate for candidate in _comment_candidates(algorithm, preview, total, note)
        if _utf16_length(candidate) <= available
    ), None)
    if section is None:
        issues.append(Issue(
            "warning", "source_comment_full",
            "原批注剩余空间不足，未改动原批注；完整来源见运行日志。",
            template_sheet=rule.sheet, template_cell=rule.coordinate,
        ))
        return issues

    text = prefix + section
    updated = copy(original) if original is not None else Comment(text, "汇总工具")
    updated.text = text
    updated.width = max(updated.width, 640)
    updated.height = max(updated.height, min(900, max(120, len(text.splitlines()) * 18 + 24)))
    cell.comment = updated
    return issues
