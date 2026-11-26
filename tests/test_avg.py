from decimal import Decimal, localcontext
from hashlib import sha256
import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

from summary_tool.engine import aggregate
from summary_tool.models import AggregationError
from summary_tool.rules import RuleError, parse_rule


class AvgRuleTests(unittest.TestCase):
    def test_case_insensitive_single_argument_and_absolute_address(self):
        for expression in ("AVG(A5)", "avg(a5)", " AvG( $a$5 ) "):
            with self.subTest(expression=expression):
                rule = parse_rule(expression, "业务明细", "C5")
                self.assertIsNotNone(rule)
                self.assertEqual(rule.operation, "AVG")
                self.assertEqual((rule.sheet, rule.coordinate), ("业务明细", "C5"))
                self.assertEqual(len(rule.references), 1)
                self.assertEqual(rule.references[0].sheet, "业务明细")
                self.assertEqual(rule.references[0].coordinate, "A5")

    def test_single_cross_sheet_reference(self):
        cases = (
            ("AVG(余额表!$b$7)", "余额表", "B7"),
            ("avg('利率, 一'!$C$8)", "利率, 一", "C8"),
            ("AvG('O''Brien 表'!D9)", "O'Brien 表", "D9"),
        )
        for expression, sheet, coordinate in cases:
            with self.subTest(expression=expression):
                rule = parse_rule(expression, "结果", "E6")
                self.assertEqual(rule.operation, "AVG")
                self.assertEqual(len(rule.references), 1)
                self.assertEqual(rule.references[0].sheet, sheet)
                self.assertEqual(rule.references[0].coordinate, coordinate)

    def test_reject_multiple_range_and_product_arguments(self):
        expressions = (
            "AVG(A5,A6)", "AVG(A5，A6)",
            "AVG(A5,)", "AVG(,A5)", "AVG(A5:A6)", "AVG(A5*B5)",
        )
        for expression in expressions:
            with self.subTest(expression=expression), self.assertRaises(RuleError):
                parse_rule(expression, "业务明细", "C5")

    def test_reject_incomplete_external_and_equals_prefixed_rules(self):
        expressions = (
            "AVG(A5", "avg(", "AVG(A5))", "AVG((A5)",
            "AVG('[源.xlsx]业务明细'!A5)", "AVG([源.xlsx]业务明细!A5)",
            "=AVG(A5)", " = avg( A5 ) ",
        )
        for expression in expressions:
            with self.subTest(expression=expression), self.assertRaises(RuleError):
                parse_rule(expression, "业务明细", "C5")


class AvgFileTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory(prefix="summary_avg_test_")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.template = self.root / "模板.xlsx"
        self.output = self.root / "结果.xlsx"
        self.make_template()

    def make_template(self, rules=None, *, include_notes=False):
        workbook = Workbook()
        try:
            sheet = workbook.active
            sheet.title = "业务明细"
            sheet.merge_cells("B2:F2")
            sheet["B2"] = "AVG 独立测试模板"
            sheet["G5"] = "固定说明"
            for coordinate, expression in (rules or {"C5": "AVG(A5)"}).items():
                sheet[coordinate] = expression
            sheet["C5"].number_format = "#,##0.00"
            sheet["C5"].font = Font(bold=True, color="FF14532D")
            sheet["C5"].fill = PatternFill(fill_type="solid", fgColor="FFE8F5E9")
            sheet.column_dimensions["C"].width = 19
            sheet.row_dimensions[5].height = 28
            if include_notes:
                workbook.create_sheet("模板附页")["B2"] = "仅保留模板自带附页"
            workbook.save(self.template)
        finally:
            workbook.close()

    def make_source(self, name, value=None, *, rate=None,
                    sheet_name="业务明细", include_extra_sheet=False):
        path = self.root / name
        workbook = Workbook()
        try:
            sheet = workbook.active
            sheet.title = sheet_name
            if value is not None:
                sheet["A5"] = value
            if rate is not None:
                sheet["B5"] = rate
            if include_extra_sheet:
                workbook.create_sheet("源独有")["B2"] = "不可复制到结果"
            workbook.save(path)
        finally:
            workbook.close()
        return path

    def read_average(self, output=None):
        workbook = load_workbook(output or self.output, data_only=False)
        try:
            return workbook["业务明细"]["C5"].value
        finally:
            workbook.close()

    def aggregate_average(self, sources, expected, *, source_count, output=None):
        target = output or self.output
        result = aggregate(self.template, sources, target)
        self.assertEqual(self.read_average(target), expected)
        self.assertEqual(result.source_count, source_count)
        self.assertEqual(result.rule_count, 1)
        self.assertEqual(result.avg_rule_count, 1)
        self.assertEqual((result.sum_rule_count, result.wavg_rule_count), (0, 0))
        return result

    def assert_batch_fails(self, sources, *, code=None):
        with self.assertRaises(AggregationError) as caught:
            aggregate(self.template, sources, self.output)
        self.assertFalse(self.output.exists())
        self.assertEqual(list(self.root.glob(".汇总临时_*")), [])
        if code is not None:
            self.assertTrue(
                any(issue.code == code and issue.severity == "error"
                    for issue in caught.exception.issues),
                caught.exception.issues,
            )
        return caught.exception

    def test_two_files_use_arithmetic_average_not_sum_or_weighted_average(self):
        first = self.make_source("甲.xlsx", 100)
        second = self.make_source("乙.xlsx", 900)
        result = self.aggregate_average([first, second], 500, source_count=2)
        self.assertEqual(result.source_sheet_count, 2)
        self.assertEqual(result.template_sheet_count, 1)

    def test_blank_is_zero_and_still_counts_in_denominator(self):
        first = self.make_source("有值.xlsx", 100)
        for index, value in enumerate((None, "", " \t\u200b")):
            with self.subTest(blank=value):
                blank = self.make_source(f"空白{index}.xlsx", value)
                result = self.aggregate_average(
                    [first, blank], 50, source_count=2,
                    output=self.root / f"空白结果{index}.xlsx",
                )
                self.assertFalse(any(issue.code == "blank_rate_with_weight"
                                     for issue in result.issues))

    def test_all_blank_is_zero_without_wavg_zero_weight_issue(self):
        sources = [self.make_source("全空甲.xlsx"),
                   self.make_source("全空乙.xlsx", " \t")]
        result = self.aggregate_average(sources, 0, source_count=2)
        self.assertFalse(any(issue.code in {"zero_weight", "blank_rate_with_weight"}
                             for issue in result.issues))
        self.assertEqual(result.warning_count, 0)

    def test_single_file_preserves_value_with_absolute_cross_sheet_reference(self):
        self.make_template({"C5": "aVg('源 数据'!$A$5)"})
        source = self.make_source("唯一.xlsx", Decimal("137.25"), sheet_name="源 数据")
        self.aggregate_average([source], 137.25, source_count=1)

    def test_real_zero_and_negative_values_are_included(self):
        sources = [self.make_source("正值.xlsx", 100),
                   self.make_source("真实零.xlsx", 0),
                   self.make_source("负值.xlsx", -250)]
        result = self.aggregate_average(sources, -50, source_count=3)
        self.assertEqual(result.warning_count, 0)
        self.assertFalse(any(issue.code in {"zero_weight", "negative_weight"}
                             for issue in result.issues))

    def test_percentage_text_is_converted_before_average(self):
        sources = [self.make_source("百分号.xlsx", "3.5%"),
                   self.make_source("全角百分号.xlsx", "4.5％")]
        result = aggregate(self.template, sources, self.output)
        self.assertAlmostEqual(self.read_average(), 0.04, places=15)
        self.assertEqual(result.source_count, 2)
        self.assertEqual(result.avg_rule_count, 1)

    def test_nonterminating_decimal_average_keeps_excel_precision(self):
        values = (Decimal("0.1"), Decimal("0.2"), Decimal("0.2"))
        sources = [self.make_source(f"小数{index}.xlsx", value)
                   for index, value in enumerate(values)]
        result = aggregate(self.template, sources, self.output)
        actual = self.read_average()
        self.assertIsInstance(actual, (int, float))
        with localcontext() as context:
            context.prec = 50
            expected = sum(values, Decimal(0)) / Decimal(len(values))
            self.assertLess(abs(Decimal(str(actual)) - expected), Decimal("1e-16"))
        self.assertEqual(result.source_count, 3)
        self.assertEqual(result.avg_rule_count, 1)

    def test_duplicate_paths_do_not_inflate_denominator(self):
        first = self.make_source("甲.xlsx", 100)
        second = self.make_source("乙.xlsx", 900)
        result = self.aggregate_average(
            [first, second, first, str(first.resolve())], 500, source_count=2,
        )
        self.assertEqual(sum(issue.code == "duplicate_source" for issue in result.issues), 2)

    def test_hard_links_to_same_file_count_once(self):
        first = self.make_source("原文件.xlsx", 100)
        alias = self.root / "同文件硬链接.xlsx"
        try:
            os.link(first, alias)
        except (OSError, NotImplementedError) as exc:
            self.skipTest(f"当前临时目录不支持创建硬链接，无法验证同文件身份去重：{exc}")
        second = self.make_source("另一文件.xlsx", 900)
        result = self.aggregate_average([first, alias, second], 500, source_count=2)
        self.assertTrue(any(issue.code == "duplicate_source" for issue in result.issues))

    def test_corrupt_and_nonexistent_excel_lock_files_are_excluded(self):
        first = self.make_source("甲.xlsx", 100)
        second = self.make_source("乙.xlsx", 900)
        corrupt_lock = self.root / "~$损坏.xlsx"
        corrupt_lock.write_bytes(b"not an xlsx archive")
        missing_lock = self.root / "~$不存在.xlsx"
        result = self.aggregate_average(
            [corrupt_lock, first, missing_lock, second], 500, source_count=2,
        )
        self.assertFalse(missing_lock.exists())
        self.assertEqual(sum(issue.code == "lock_file" for issue in result.issues), 2)
        self.assertFalse(any(issue.severity == "error" for issue in result.issues))

    def test_no_effective_source_files_fails_without_output(self):
        corrupt_lock = self.root / "~$损坏.xlsx"
        corrupt_lock.write_bytes(b"not an xlsx archive")
        missing_lock = self.root / "~$不存在.xlsx"
        for sources in ([], [corrupt_lock], [missing_lock], [corrupt_lock, missing_lock]):
            with self.subTest(sources=sources):
                self.assert_batch_fails(sources)

    def test_duplicates_blanks_and_locks_use_two_or_three_effective_files(self):
        first = self.make_source("有值.xlsx", 100)
        blank = self.make_source("空白.xlsx")
        third = self.make_source("第三份.xlsx", 200)
        corrupt_lock = self.root / "~$损坏.xlsx"
        corrupt_lock.write_bytes(b"not an xlsx archive")
        missing_lock = self.root / "~$不存在.xlsx"
        base = [first, first, blank, blank, corrupt_lock, missing_lock]
        for sources, count, expected in ((base, 2, 50), (base + [third], 3, 100)):
            with self.subTest(effective_source_count=count):
                result = self.aggregate_average(
                    sources, expected, source_count=count,
                    output=self.root / f"组合结果{count}.xlsx",
                )
                self.assertEqual(sum(issue.code == "duplicate_source"
                                     for issue in result.issues), 2)
                self.assertEqual(sum(issue.code == "lock_file"
                                     for issue in result.issues), 2)

    def test_equals_avg_in_template_fails_without_output(self):
        self.make_template({"C5": "=AVG(A5)"})
        source = self.make_source("正常.xlsx", 100)
        self.assert_batch_fails([source], code="template_formula_not_supported")

    def test_missing_sheet_rejects_entire_batch(self):
        good = self.make_source("正常.xlsx", 100)
        invalid = self.make_source("缺表.xlsx", 900, sheet_name="其他表")
        error = self.assert_batch_fails([good, invalid], code="missing_sheet")
        issue = next(issue for issue in error.issues if issue.code == "missing_sheet")
        self.assertEqual(Path(issue.source_file), invalid.resolve())
        self.assertEqual((issue.source_sheet, issue.source_cell), ("业务明细", "A5"))
        self.assertEqual((issue.template_sheet, issue.template_cell), ("业务明细", "C5"))

    def test_nonnumeric_source_rejects_entire_batch(self):
        good = self.make_source("正常.xlsx", 100)
        for index, value in enumerate(("不是数字", True)):
            with self.subTest(value=value):
                invalid = self.make_source(f"非数值{index}.xlsx", value)
                self.assert_batch_fails([good, invalid], code="invalid_number")

    def test_excel_error_rejects_entire_batch(self):
        good = self.make_source("正常.xlsx", 100)
        invalid = self.make_source("Excel错误.xlsx", "#DIV/0!")
        self.assert_batch_fails([good, invalid], code="excel_error")

    def test_formula_without_cache_rejects_entire_batch(self):
        good = self.make_source("正常.xlsx", 100)
        invalid = self.make_source("无缓存公式.xlsx", "=1+2")
        self.assert_batch_fails([good, invalid], code="formula_cache_missing")

    def test_corrupt_source_rejects_entire_batch(self):
        good = self.make_source("正常.xlsx", 100)
        invalid = self.root / "损坏.xlsx"
        invalid.write_bytes(b"not an xlsx archive")
        self.assert_batch_fails([good, invalid], code="open_failed")

    def test_avg_sum_and_wavg_share_template_without_changing_other_results(self):
        self.make_template({"C5": "AVG(A5)", "D5": "SUM(A5)", "E5": "WAVG(B5,A5)"})
        first = self.make_source("甲.xlsx", 100, rate="3%")
        second = self.make_source("乙.xlsx", 900, rate="5%")
        result = aggregate(self.template, [first, second], self.output)
        workbook = load_workbook(self.output, data_only=False)
        try:
            sheet = workbook["业务明细"]
            self.assertEqual(sheet["C5"].value, 500)
            self.assertEqual(sheet["D5"].value, 1000)
            self.assertAlmostEqual(sheet["E5"].value, 0.048, places=15)
        finally:
            workbook.close()
        self.assertEqual(result.source_count, 2)
        self.assertEqual(result.source_sheet_count, 2)
        self.assertEqual(result.template_sheet_count, 1)
        self.assertEqual(result.rule_count, 3)
        self.assertEqual((result.avg_rule_count, result.sum_rule_count, result.wavg_rule_count),
                         (1, 1, 1))

    def test_style_input_hashes_and_only_template_sheets_plus_log_with_zero_formulas(self):
        self.make_template(include_notes=True)
        first = self.make_source("甲.xlsx", 100, include_extra_sheet=True)
        second = self.make_source("乙.xlsx", 900)
        inputs = (self.template, first, second)
        before = {path: sha256(path.read_bytes()).hexdigest() for path in inputs}
        result = self.aggregate_average([first, second], 500, source_count=2)
        self.assertEqual(result.template_sheet_count, 2)
        self.assertEqual(before, {path: sha256(path.read_bytes()).hexdigest() for path in inputs})
        workbook = load_workbook(self.output, data_only=False)
        try:
            self.assertEqual(workbook.sheetnames, ["业务明细", "模板附页", "运行日志"])
            sheet = workbook["业务明细"]
            self.assertEqual(sheet["B2"].value, "AVG 独立测试模板")
            self.assertEqual(sheet["G5"].value, "固定说明")
            self.assertIn("B2:F2", sheet.merged_cells)
            self.assertTrue(sheet["C5"].font.bold)
            self.assertEqual(sheet["C5"].font.color.rgb, "FF14532D")
            self.assertEqual(sheet["C5"].fill.fgColor.rgb, "FFE8F5E9")
            self.assertEqual(sheet["C5"].number_format, "#,##0.00")
            self.assertEqual(sheet.column_dimensions["C"].width, 19)
            self.assertEqual(sheet.row_dimensions[5].height, 28)
            self.assertEqual(workbook["模板附页"]["B2"].value, "仅保留模板自带附页")
            self.assertEqual(sum(cell.data_type == "f" for ws in workbook
                                 for row in ws for cell in row), 0)
        finally:
            workbook.close()

    def test_avg_does_not_overwrite_inputs_or_existing_result(self):
        source = self.make_source("源文件.xlsx", 100)
        existing = self.make_source("结果.xlsx", 731, sheet_name="已有结果")
        protected = (self.template, source, existing)
        before = {path: sha256(path.read_bytes()).hexdigest() for path in protected}
        for target in protected:
            with self.subTest(target=target.name), self.assertRaises(AggregationError):
                aggregate(self.template, [source], target)
            self.assertEqual(before,
                             {path: sha256(path.read_bytes()).hexdigest() for path in protected})
        self.assertEqual(list(self.root.glob(".汇总临时_*")), [])


if __name__ == "__main__":
    unittest.main()
