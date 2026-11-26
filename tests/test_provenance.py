from dataclasses import replace
from decimal import Decimal
from io import BytesIO
from pathlib import Path
from types import SimpleNamespace
import unittest
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment

from summary_tool.functions import CalculationNotice, FunctionResult, ResultSource
from summary_tool.models import AggregationError
from summary_tool.provenance import append_result_provenance
from summary_tool.rules import CellReference, TemplateRule


def _result(sources=(), *, note="", notices=(), value=Decimal("12.5")):
    return FunctionResult(
        value, notices=tuple(notices), sources=tuple(sources), source_note=note,
    )


def _utf16_length(text):
    return len(text.encode("utf-16-le")) // 2


class ResultProvenanceTests(unittest.TestCase):
    def setUp(self):
        self.workbook = Workbook()
        self.addCleanup(self.workbook.close)
        self.sheet = self.workbook.active
        self.sheet.title = "汇总表"
        self.cell = self.sheet["H8"]
        self.cell.value = Decimal("12.5")
        self.cell.number_format = "0.00"
        self.rule = TemplateRule(
            "汇总表", "H8", "CUSTOM", (
                CellReference("原始 数据", "B7", 7, 2),
                CellReference("异表O'Brien", "D19", 19, 4),
            ), "F2_9('原始 数据'!B7,'异表O''Brien'!D19)", "F2_9",
        )
        self.source_root = Path(__file__).resolve().parent / "不存在的来源目录"
        self.paths = [self.source_root / "来源甲.xlsx", self.source_root / "来源乙.xlsx"]

    def test_single_source_uses_classic_comment_and_keeps_number_and_format(self):
        original_value = self.cell.value
        rule = replace(self.rule, operation="CUSTOM", function_name="CUSTOM")
        issues = append_result_provenance(self.cell, rule, _result((ResultSource(0),)), self.paths)

        self.assertEqual(len(issues), 1)
        issue = issues[0]
        self.assertEqual((issue.severity, issue.code), ("info", "result_source"))
        self.assertEqual(issue.source_file, str(self.paths[0]))
        self.assertTrue(Path(issue.source_file).is_absolute())
        self.assertEqual((issue.source_sheet, issue.source_cell), ("原始 数据", "B7"))
        self.assertEqual((issue.template_sheet, issue.template_cell), ("汇总表", "H8"))
        self.assertIn("CUSTOM", issue.message)
        self.assertIn("1/1", issue.message)
        self.assertIsInstance(self.cell.comment, Comment)
        self.assertEqual(self.cell.comment.author, "汇总工具")
        for expected in ("CUSTOM", self.paths[0].name, "原始 数据", "B7", str(self.paths[0])):
            self.assertIn(expected, self.cell.comment.text)
        self.assertIs(self.cell.value, original_value)
        self.assertEqual(self.cell.data_type, "n")
        self.assertEqual(self.cell.number_format, "0.00")

    def test_real_algorithm_and_alias_work_without_function_name_branches(self):
        for operation in ("SUM", "WAVG", "FUTURE_ALGORITHM"):
            with self.subTest(operation=operation):
                self.cell.comment = None
                rule = replace(self.rule, operation=operation)
                issues = append_result_provenance(
                    self.cell, rule, _result((ResultSource(0, 1),)), self.paths,
                )
                for text in (self.cell.comment.text, issues[0].message):
                    self.assertIn(operation, text)
                    self.assertIn("F2_9", text)
                self.assertEqual(issues[0].source_cell, "D19")

    def test_more_than_ten_ties_preview_first_ten_but_log_every_source(self):
        paths = [self.source_root / f"并列来源_{index:02d}.xlsx" for index in range(1, 14)]
        sources = tuple(ResultSource(index) for index in range(len(paths)))
        issues = append_result_provenance(self.cell, self.rule, _result(sources), paths)

        self.assertEqual(len(issues), 13)
        for ordinal, (path, issue) in enumerate(zip(paths, issues), 1):
            self.assertEqual((issue.severity, issue.code), ("info", "result_source"))
            self.assertEqual(issue.source_file, str(path))
            self.assertEqual((issue.source_sheet, issue.source_cell), ("原始 数据", "B7"))
            self.assertEqual((issue.template_sheet, issue.template_cell), ("汇总表", "H8"))
            self.assertIn("CUSTOM", issue.message)
            self.assertIn(f"{ordinal}/13", issue.message)
            if ordinal <= 10:
                self.assertIn(path.name, self.cell.comment.text)
                self.assertIn(str(path), self.cell.comment.text)
            else:
                self.assertNotIn(path.name, self.cell.comment.text)
        self.assertIn("并列共13个，完整来源见运行日志", self.cell.comment.text)
        self.assertLessEqual(_utf16_length(self.cell.comment.text), 32767)

    def test_source_indices_preserve_original_ordinals_and_metadata_order(self):
        paths = [self.source_root / f"有效源列表_{index}.xlsx" for index in range(5)]
        calculated = _result((ResultSource(4), ResultSource(2)))
        issues = append_result_provenance(self.cell, self.rule, calculated, paths)

        self.assertEqual([issue.source_file for issue in issues], [str(paths[4]), str(paths[2])])
        self.assertIn("1/2", issues[0].message)
        self.assertIn("2/2", issues[1].message)
        self.assertIn("5", issues[0].message)
        self.assertIn("3", issues[1].message)
        text = self.cell.comment.text
        self.assertLess(text.index(paths[4].name), text.index(paths[2].name))
        for index in (0, 1, 3):
            self.assertNotIn(paths[index].name, text)

    def test_parameter_indices_map_cross_sheet_and_offset_references(self):
        issues = append_result_provenance(
            self.cell, self.rule, _result((ResultSource(1, 1), ResultSource(0, 0))), self.paths,
        )
        self.assertEqual(
            [(issue.source_file, issue.source_sheet, issue.source_cell) for issue in issues],
            [(str(self.paths[1]), "异表O'Brien", "D19"), (str(self.paths[0]), "原始 数据", "B7")],
        )
        self.assertIn("'异表O''Brien'!D19", self.cell.comment.text)
        self.assertIn("'原始 数据'!B7", self.cell.comment.text)
        self.assertTrue(all(issue.template_sheet == "汇总表" and issue.template_cell == "H8" for issue in issues))

    def test_relative_source_paths_are_logged_as_complete_absolute_paths(self):
        path = Path("相对来源") / ".." / "输入资料" / "来源.xlsx"
        issues = append_result_provenance(
            self.cell, self.rule, _result((ResultSource(0),)), [path],
        )
        self.assertEqual(issues[0].source_file, str(path.resolve()))
        self.assertTrue(Path(issues[0].source_file).is_absolute())
        self.assertIn(issues[0].source_file, self.cell.comment.text)

    def test_existing_comment_text_author_and_larger_dimensions_are_preserved(self):
        old_text = "原作者的业务说明\n第二行，保留末尾换行。\n"
        old = Comment(old_text, "业务审核人", width=1100, height=1200)
        self.cell.comment = old
        append_result_provenance(self.cell, self.rule, _result((ResultSource(0),)), self.paths)

        updated = self.cell.comment
        self.assertIsNot(updated, old)
        self.assertTrue(updated.text.startswith(old_text + "\n\n"))
        self.assertEqual(updated.author, "业务审核人")
        self.assertGreaterEqual(updated.width, 1100)
        self.assertGreaterEqual(updated.height, 1200)
        self.assertEqual((old.text, old.author, old.width, old.height), (old_text, "业务审核人", 1100, 1200))
        self.assertIn(str(self.paths[0]), updated.text)

    def test_small_existing_comment_can_grow_without_mutating_original(self):
        old = Comment("业务说明", "", width=80, height=30)
        self.cell.comment = old
        append_result_provenance(
            self.cell, self.rule, _result((ResultSource(0), ResultSource(1))), self.paths,
        )
        self.assertGreater(self.cell.comment.width, 80)
        self.assertGreater(self.cell.comment.height, 30)
        self.assertEqual(self.cell.comment.author, "")
        self.assertEqual((old.text, old.width, old.height), ("业务说明", 80, 30))

    def test_note_only_marks_no_valid_number_without_inventing_a_source_file(self):
        note = "所有输入均为空，结果按约定显示0。"
        self.cell.value = Decimal(0)
        issues = append_result_provenance(
            self.cell, self.rule, _result(note=note, value=Decimal(0)), self.paths,
        )
        self.assertEqual(len(issues), 1)
        issue = issues[0]
        self.assertEqual((issue.severity, issue.code, issue.message), ("info", "result_source_note", note))
        self.assertEqual((issue.source_file, issue.source_sheet, issue.source_cell), ("", "", ""))
        self.assertEqual((issue.template_sheet, issue.template_cell), ("汇总表", "H8"))
        for expected in ("无有效数值", note, "CUSTOM", "F2_9"):
            self.assertIn(expected, self.cell.comment.text)
        for path in self.paths:
            self.assertNotIn(path.name, self.cell.comment.text)
        self.assertNotIn("完整路径：", self.cell.comment.text)
        self.assertEqual(self.cell.value, Decimal(0))
        self.assertEqual(self.cell.data_type, "n")

    def test_note_matching_a_calculation_notice_is_not_logged_twice(self):
        note = "全部来源为空，显示0。"
        notices = (
            CalculationNotice("info", "other", "另一条说明。"),
            CalculationNotice("warning", "all_blank", note),
        )
        for sources in ((), (ResultSource(0),)):
            with self.subTest(sources=sources):
                self.cell.comment = None
                issues = append_result_provenance(
                    self.cell, self.rule, _result(sources, note=note, notices=notices), self.paths,
                )
                self.assertEqual(len(issues), len(sources))
                self.assertTrue(all(issue.code == "result_source" for issue in issues))
                self.assertIn(note, self.cell.comment.text)

    def test_distinct_note_is_logged_once_alongside_all_result_sources(self):
        note = "这是来源层的补充说明。"
        calculated = _result(
            (ResultSource(0), ResultSource(1)), note=note,
            notices=(CalculationNotice("info", "notice", "这是另一条计算提醒。"),),
        )
        issues = append_result_provenance(self.cell, self.rule, calculated, self.paths)
        self.assertEqual([issue.code for issue in issues], ["result_source", "result_source", "result_source_note"])
        self.assertEqual(issues[-1].message, note)
        self.assertEqual((issues[-1].source_file, issues[-1].source_sheet, issues[-1].source_cell), ("", "", ""))
        self.assertIn(note, self.cell.comment.text)

    def test_no_metadata_is_immediate_noop_even_with_notices(self):
        calculated = _result(notices=(CalculationNotice("info", "zero_weight", "仅计算提醒。"),))
        self.assertEqual(append_result_provenance(object(), object(), calculated, object()), [])
        old = Comment("原批注不能改动", "原作者", width=700, height=300)
        self.cell.comment = old
        self.assertEqual(append_result_provenance(self.cell, self.rule, calculated, self.paths), [])
        self.assertIs(self.cell.comment, old)
        self.assertEqual((old.text, old.author, old.width, old.height), ("原批注不能改动", "原作者", 700, 300))
        self.cell.comment = None
        self.assertEqual(append_result_provenance(self.cell, self.rule, calculated, self.paths), [])
        self.assertIsNone(self.cell.comment)
        self.assertEqual(self.cell.value, Decimal("12.5"))
        self.assertEqual(self.cell.data_type, "n")

    def test_overlong_path_uses_summary_but_issue_keeps_complete_path(self):
        path = self.source_root.joinpath(*(["超长目录" * 12] * 700), "长路径来源.xlsx")
        self.assertGreater(_utf16_length(str(path)), 32767)
        issues = append_result_provenance(
            self.cell, self.rule, _result((ResultSource(0),)), [path],
        )
        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0].source_file, str(path))
        self.assertIn(path.name, self.cell.comment.text)
        self.assertIn("'原始 数据'!B7", self.cell.comment.text)
        self.assertIn("CUSTOM", self.cell.comment.text)
        self.assertIn("运行日志", self.cell.comment.text)
        self.assertNotIn(str(path), self.cell.comment.text)
        self.assertLessEqual(_utf16_length(self.cell.comment.text), 32767)

    def test_overlong_filename_and_note_are_explicitly_summarized(self):
        path = self.source_root / ("很长的文件名" * 6000 + ".xlsx")
        note = "很长的来源说明。" * 5000
        issues = append_result_provenance(
            self.cell, self.rule, _result((ResultSource(0),), note=note), [path],
        )
        self.assertEqual(issues[0].source_file, str(path))
        self.assertEqual(issues[1].message, note)
        self.assertLessEqual(_utf16_length(self.cell.comment.text), 32767)
        self.assertIn("运行日志", self.cell.comment.text)
        self.assertIn("…", self.cell.comment.text)
        self.assertNotIn(path.name, self.cell.comment.text)
        self.assertNotIn(note, self.cell.comment.text)

    def test_utf16_budget_preserves_original_astral_text_without_truncation(self):
        original = "\U00020000" * 16310
        old = Comment(original, "原作者")
        self.cell.comment = old
        issues = append_result_provenance(
            self.cell, self.rule, _result((ResultSource(0),)), self.paths,
        )
        self.assertTrue(self.cell.comment.text.startswith(original + "\n\n"))
        self.assertEqual(self.cell.comment.author, "原作者")
        self.assertEqual(old.text, original)
        self.assertIn("运行日志", self.cell.comment.text)
        self.assertLessEqual(_utf16_length(self.cell.comment.text), 32767)
        self.assertFalse(any(issue.code == "source_comment_full" for issue in issues))

    def test_full_original_comment_stays_unchanged_and_all_sources_are_logged(self):
        old = Comment("旧" * 32760, "原作者", width=900, height=1000)
        self.cell.comment = old
        paths = self.paths + [self.source_root / "来源丙.xlsx"]
        issues = append_result_provenance(
            self.cell, self.rule, _result(tuple(ResultSource(index) for index in range(3))), paths,
        )
        self.assertIs(self.cell.comment, old)
        self.assertEqual((old.text, old.author, old.width, old.height), ("旧" * 32760, "原作者", 900, 1000))
        self.assertEqual([issue.source_file for issue in issues if issue.code == "result_source"], [str(path) for path in paths])
        warnings = [issue for issue in issues if issue.code == "source_comment_full"]
        self.assertEqual(len(warnings), 1)
        self.assertEqual(warnings[0].severity, "warning")
        self.assertIn("完整来源见运行日志", warnings[0].message)
        self.assertEqual((warnings[0].template_sheet, warnings[0].template_cell), ("汇总表", "H8"))

    def test_already_oversize_original_is_never_silently_truncated(self):
        old = Comment("已有超长原文" * 7000, "原作者")
        original = old.text
        self.cell.comment = old
        note = "没有有效数值。"
        issues = append_result_provenance(self.cell, self.rule, _result(note=note), self.paths)
        self.assertIs(self.cell.comment, old)
        self.assertEqual(old.text, original)
        self.assertEqual([issue.code for issue in issues], ["result_source_note", "source_comment_full"])
        self.assertEqual(issues[0].message, note)
        self.assertEqual(self.cell.value, Decimal("12.5"))

    def test_invalid_source_indices_raise_aggregation_error_without_mutation(self):
        old = Comment("不允许部分追加", "原作者")
        self.cell.comment = old
        for index in (-1, len(self.paths), 0.0, "0", True, False, None, Decimal(0)):
            with self.subTest(source_index=index):
                source = SimpleNamespace(source_index=index, parameter_index=0)
                calculated = _result((ResultSource(0), source))
                with self.assertRaises(AggregationError):
                    append_result_provenance(self.cell, self.rule, calculated, self.paths)
                self.assertIs(self.cell.comment, old)
                self.assertEqual(old.text, "不允许部分追加")
        with self.assertRaises(AggregationError):
            append_result_provenance(self.cell, self.rule, _result((ResultSource(0),)), [])
        self.assertIs(self.cell.comment, old)

    def test_invalid_parameter_indices_raise_aggregation_error_without_mutation(self):
        old = Comment("不允许部分追加", "原作者")
        self.cell.comment = old
        for index in (-1, len(self.rule.references), 0.0, "0", True, False, None, Decimal(0)):
            with self.subTest(parameter_index=index):
                source = SimpleNamespace(source_index=1, parameter_index=index)
                calculated = _result((ResultSource(0), source))
                with self.assertRaises(AggregationError):
                    append_result_provenance(self.cell, self.rule, calculated, self.paths)
                self.assertIs(self.cell.comment, old)
                self.assertEqual(old.text, "不允许部分追加")
        with self.assertRaises(AggregationError):
            append_result_provenance(
                self.cell, replace(self.rule, references=()), _result((ResultSource(0),)), self.paths,
            )
        self.assertIs(self.cell.comment, old)

    def test_classic_comment_round_trip_in_memory_keeps_numeric_cell(self):
        self.cell.comment = Comment("模板的原备注", "模板作者")
        append_result_provenance(self.cell, self.rule, _result((ResultSource(1, 1),)), self.paths)
        with BytesIO() as buffer:
            self.workbook.save(buffer)
            buffer.seek(0)
            with ZipFile(buffer) as archive:
                names = archive.namelist()
                self.assertTrue(any(name.startswith("xl/comments/") and name.endswith(".xml") for name in names))
                self.assertFalse(any("threadedcomments" in name.lower() for name in names))
            buffer.seek(0)
            reloaded = load_workbook(buffer)
            try:
                cell = reloaded["汇总表"]["H8"]
                self.assertEqual(cell.value, 12.5)
                self.assertEqual(cell.data_type, "n")
                self.assertEqual(cell.number_format, "0.00")
                self.assertIsInstance(cell.comment, Comment)
                self.assertEqual(cell.comment.author, "模板作者")
                self.assertTrue(cell.comment.text.startswith("模板的原备注\n\n"))
                self.assertIn(str(self.paths[1]), cell.comment.text)
                self.assertIn("'异表O''Brien'!D19", cell.comment.text)
                self.assertEqual(reloaded.sheetnames, ["汇总表"])
            finally:
                reloaded.close()


if __name__ == "__main__":
    unittest.main()
