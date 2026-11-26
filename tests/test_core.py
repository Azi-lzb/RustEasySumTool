from decimal import Decimal
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from threading import Event
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from summary_tool.engine import AggregationCancelled, AggregationError, aggregate, default_output_path
from summary_tool.reader import SourceReadError, SourceWorkbook, decimal_value
from summary_tool.rules import RuleError, parse_reference, parse_rule


class RuleTests(unittest.TestCase):
    def test_sum_case_and_absolute_address(self):
        rule = parse_rule(" sum( $a$5 ) ", "业务明细", "C8")
        self.assertEqual(rule.operation, "SUM")
        self.assertEqual(rule.references[0].coordinate, "A5")
        self.assertEqual(rule.references[0].sheet, "业务明细")
        self.assertEqual(rule.coordinate, "C8")

    def test_quoted_cross_sheet_and_commas(self):
        rule = parse_rule("WAVG('利率, 一'!B5,'O''Brien 表'!$A$5)", "结果", "D9")
        self.assertEqual(rule.references[0].sheet, "利率, 一")
        self.assertEqual(rule.references[1].sheet, "O'Brien 表")

    def test_chinese_unquoted_sheet(self):
        self.assertEqual(parse_reference("余额表!a5", "结果").sheet, "余额表")

    def test_native_formula_and_plain_text_are_not_rules(self):
        for text in ("=SUM(A5:A9)", "示例：SUM(A5)", "单位：元", 123, None):
            self.assertIsNone(parse_rule(text, "表", "A1"))

    def test_invalid_rules(self):
        for text in ("MEDIAN(A5)", "FOO(A5)", "COUNTA(A5)", "WAVG(B5)", "SUM(A0)",
                     "SUM(A5*B5)", "SUM(A5", "WAVG('表!A5,B5)", "=WAVG(B5,A5)",
                     "SUM(XFE1)", "SUM(A1048577)", "SUM('[源.xlsx]表'!A5)", "SUM(__import__('os'))"):
            with self.subTest(text=text), self.assertRaises(RuleError):
                parse_rule(text, "表", "A1")

    def test_excel_boundary(self):
        reference = parse_reference("XFD1048576", "表")
        self.assertEqual((reference.row, reference.column), (1048576, 16384))


class NumberTests(unittest.TestCase):
    def test_empty_and_zero_are_different(self):
        for value in (None, "", "  ", "\xa0\u200b"):
            self.assertTrue(decimal_value(value).blank)
        self.assertFalse(decimal_value(0).blank)
        self.assertFalse(decimal_value("0%").blank)

    def test_numeric_strings(self):
        cases = {"1,234.50": "1234.50", "3.5%": ".035", "4％": ".04", "-12.3": "-12.3", ".03": ".03"}
        for text, expected in cases.items():
            self.assertEqual(decimal_value(text).value, Decimal(expected))

    def test_invalid_numbers_are_not_zero(self):
        for value in (True, "没有", "1,2", "—", float("inf"), float("nan"), "#DIV/0!"):
            with self.subTest(value=value), self.assertRaises(SourceReadError):
                decimal_value(value)


class FileTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.template = self.root / "模板.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "业务明细"
        sheet.merge_cells("B2:F2")
        sheet["B2"] = "固定标题"
        sheet["C5"] = "SUM(C5)"
        sheet["D5"] = "WAVG(D5,C5)"
        sheet["E5"] = "固定说明"
        sheet["C5"].number_format = "#,##0.00"
        sheet["C5"].font = Font(bold=True)
        sheet["D5"].number_format = "0.0000%"
        sheet.column_dimensions["C"].width = 19
        sheet.row_dimensions[5].height = 28
        workbook.save(self.template)
        workbook.close()
        self.first = self.make_source("机构甲.xlsx", 100, .03)
        self.second = self.make_source("机构乙.xlsx", 900, .05)
        self.output = self.root / "结果.xlsx"

    def tearDown(self):
        self.temp.cleanup()

    def make_source(self, name, balance=None, rate=None, title="业务明细"):
        path = self.root / name
        workbook = Workbook()
        workbook.active.title = title
        if balance is not None:
            workbook.active["C5"] = balance
        if rate is not None:
            workbook.active["D5"] = rate
        workbook.save(path)
        workbook.close()
        return path

    def read_values(self, output=None):
        workbook = load_workbook(output or self.output, data_only=False)
        try:
            return workbook["业务明细"]["C5"].value, workbook["业务明细"]["D5"].value
        finally:
            workbook.close()

    def test_reader_leading_blanks_and_missing_cells(self):
        references = [parse_reference("C5", "业务明细"), parse_reference("A1", "业务明细")]
        with SourceWorkbook(self.first, references) as reader:
            self.assertEqual(reader.get(references[0]).value, Decimal(100))
            self.assertTrue(reader.get(references[1]).blank)

    def test_empty_sheet_is_zero(self):
        empty = self.make_source("空表.xlsx")
        reference = parse_reference("C5", "业务明细")
        with SourceWorkbook(empty, [reference]) as reader:
            self.assertTrue(reader.get(reference).blank)

    def test_sum_weighted_average_style_values_and_hash(self):
        inputs = (self.template, self.first, self.second)
        before = {path: sha256(path.read_bytes()).hexdigest() for path in inputs}
        result = aggregate(self.template, [self.first, self.second], self.output)
        total, rate = self.read_values()
        self.assertEqual(total, 1000)
        self.assertAlmostEqual(rate, .048, places=14)
        self.assertEqual(result.source_count, 2)
        self.assertEqual(result.rule_count, 2)
        self.assertEqual(result.source_sheet_count, 2)
        self.assertEqual(result.template_sheet_count, 1)
        self.assertEqual(result.sum_rule_count, 1)
        self.assertEqual(result.wavg_rule_count, 1)
        self.assertEqual(result.avg_rule_count, 0)
        workbook = load_workbook(self.output)
        try:
            sheet = workbook["业务明细"]
            self.assertEqual(sheet["E5"].value, "固定说明")
            self.assertEqual(sheet.column_dimensions["C"].width, 19)
            self.assertEqual(sheet.row_dimensions[5].height, 28)
            self.assertTrue(sheet["C5"].font.bold)
            self.assertEqual(sheet["C5"].number_format, "#,##0.00")
            self.assertIn("B2:F2", sheet.merged_cells)
            self.assertEqual(workbook.sheetnames, ["业务明细", "运行日志"])
            self.assertFalse(any(cell.data_type == "f" for ws in workbook for row in ws for cell in row))
        finally:
            workbook.close()
        self.assertEqual(before, {path: sha256(path.read_bytes()).hexdigest() for path in inputs})

    def test_blank_rate_keeps_weight_in_denominator(self):
        blank_rate = self.make_source("缺利率.xlsx", 900, None)
        result = aggregate(self.template, [self.first, blank_rate], self.output)
        self.assertAlmostEqual(self.read_values()[1], .003, places=14)
        self.assertTrue(any(issue.code == "blank_rate_with_weight" for issue in result.issues))

    def test_true_zero_rate_is_not_missing(self):
        zero_rate = self.make_source("零利率.xlsx", 900, 0)
        result = aggregate(self.template, [self.first, zero_rate], self.output)
        self.assertFalse(any(issue.code == "blank_rate_with_weight" for issue in result.issues))
        self.assertAlmostEqual(self.read_values()[1], .003, places=14)

    def test_zero_denominator_fails_with_reason(self):
        empty = self.make_source("全空.xlsx")
        with self.assertRaises(AggregationError) as caught:
            aggregate(self.template, [empty], self.output)
        self.assertFalse(self.output.exists())
        self.assertTrue(any(issue.code == "zero_weight" and issue.severity == "error"
                            for issue in caught.exception.issues))

    def test_negative_weight_cancellation_is_not_division(self):
        negative = self.make_source("冲销.xlsx", -100, .03)
        with self.assertRaises(AggregationError) as caught:
            aggregate(self.template, [self.first, negative], self.output)
        self.assertFalse(self.output.exists())
        self.assertTrue(any(issue.code == "zero_weight" and issue.severity == "error"
                            for issue in caught.exception.issues))

    def test_duplicate_source_is_counted_once(self):
        result = aggregate(self.template, [self.first, self.first], self.output)
        self.assertEqual(result.source_count, 1)
        self.assertEqual(self.read_values()[0], 100)

    def test_missing_sheet_fails_without_output(self):
        invalid = self.make_source("缺工作表.xlsx", 100, .03, "其他表")
        with self.assertRaises(AggregationError) as caught:
            aggregate(self.template, [self.first, invalid], self.output)
        self.assertFalse(self.output.exists())
        self.assertTrue(any(issue.code == "missing_sheet" for issue in caught.exception.issues))

    def test_invalid_number_fails_without_output(self):
        invalid = self.make_source("非法值.xlsx", "错误文本", .03)
        with self.assertRaises(AggregationError):
            aggregate(self.template, [invalid], self.output)
        self.assertFalse(self.output.exists())

    def test_source_formula_without_cache_is_not_blank(self):
        source = self.make_source("公式.xlsx", "=1+2", .03)
        with self.assertRaises(AggregationError) as caught:
            aggregate(self.template, [source], self.output)
        self.assertTrue(any(issue.code == "formula_cache_missing" for issue in caught.exception.issues))
        self.assertFalse(self.output.exists())

    def test_source_formula_with_cache(self):
        source = self.make_source("公式.xlsx", "=1+2", .03)
        cached = self.root / "公式有缓存.xlsx"
        namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        with ZipFile(source) as archive, ZipFile(cached, "w") as target:
            for part in archive.infolist():
                data = archive.read(part.filename)
                if part.filename == "xl/worksheets/sheet1.xml":
                    tree = ET.fromstring(data)
                    for cell in tree.iter(f"{namespace}c"):
                        if cell.get("r") == "C5":
                            cell.find(f"{namespace}v").text = "3"
                    data = ET.tostring(tree, encoding="utf-8")
                target.writestr(part, data)
        aggregate(self.template, [cached], self.output)
        self.assertEqual(self.read_values()[0], 3)

    def test_excel_error_is_not_zero(self):
        invalid = self.make_source("错误公式.xlsx", "#DIV/0!", .03)
        with self.assertRaises(AggregationError) as caught:
            aggregate(self.template, [invalid], self.output)
        self.assertTrue(any(issue.code == "excel_error" for issue in caught.exception.issues))

    def test_input_and_existing_output_cannot_be_overwritten(self):
        for target in (self.template, self.first):
            before = target.read_bytes()
            with self.assertRaises(AggregationError):
                aggregate(self.template, [self.first], target)
            self.assertEqual(target.read_bytes(), before)
        self.output.write_bytes(b"existing result")
        with self.assertRaises(AggregationError):
            aggregate(self.template, [self.first], self.output)
        self.assertEqual(self.output.read_bytes(), b"existing result")

    def test_template_cannot_be_selected_as_source(self):
        with self.assertRaises(AggregationError):
            aggregate(self.template, [self.template], self.output)

    def test_cancel_does_not_create_output(self):
        event = Event()
        event.set()
        with self.assertRaises(AggregationCancelled):
            aggregate(self.template, [self.first], self.output, cancel_event=event)
        self.assertFalse(self.output.exists())

    def test_file_appearing_during_save_is_not_overwritten(self):
        def progress(completed, total, message):
            if message.startswith("计算完成"):
                self.output.write_bytes(b"created by another process")
        with self.assertRaises(AggregationError):
            aggregate(self.template, [self.first], self.output, progress=progress)
        self.assertEqual(self.output.read_bytes(), b"created by another process")
        self.assertEqual(list(self.root.glob(".汇总临时_*")), [])

    def test_existing_run_log_is_appended_without_extra_sheets(self):
        workbook = load_workbook(self.template)
        log = workbook.create_sheet("运行日志")
        log.merge_cells("B2:D2")
        log["B2"] = "必须保留"
        log["B5"] = "SUM(A1)"
        log["B5"].font = Font(bold=True)
        workbook.create_sheet("附表")["B2"] = "自定义内容"
        alternate = self.root / "带日志的模板.xlsx"
        workbook.save(alternate)
        workbook.close()
        before = alternate.read_bytes()
        result = aggregate(alternate, [self.first], self.output)
        workbook = load_workbook(self.output)
        try:
            self.assertEqual(workbook.sheetnames, ["业务明细", "附表", "运行日志"])
            self.assertEqual(workbook["运行日志"]["B2"].value, "必须保留")
            self.assertEqual(workbook["运行日志"]["B5"].value, "SUM(A1)")
            self.assertTrue(workbook["运行日志"]["B5"].font.bold)
            self.assertIn("B2:D2", workbook["运行日志"].merged_cells)
            self.assertGreater(workbook["运行日志"].max_row, 8)
            self.assertEqual(workbook["附表"]["B2"].value, "自定义内容")
            self.assertEqual(result.template_sheet_count, 2)
            self.assertEqual(result.rule_count, 2)
        finally:
            workbook.close()
        self.assertEqual(alternate.read_bytes(), before)

    def test_template_excel_formula_is_rejected_without_output(self):
        workbook = load_workbook(self.template)
        workbook["业务明细"]["E5"] = "=SUM(C5:C6)"
        alternate = self.root / "含普通公式模板.xlsx"
        workbook.save(alternate)
        workbook.close()
        before = alternate.read_bytes()
        with self.assertRaises(AggregationError) as caught:
            aggregate(alternate, [self.first], self.output)
        self.assertFalse(self.output.exists())
        self.assertTrue(any(issue.code == "template_formula_not_supported" and issue.template_cell == "E5"
                            for issue in caught.exception.issues))
        self.assertEqual(alternate.read_bytes(), before)

    def test_formula_in_existing_log_is_not_carried_to_output(self):
        workbook = load_workbook(self.template)
        workbook.create_sheet("运行日志")["A1"] = "=1+1"
        alternate = self.root / "日志带公式模板.xlsx"
        workbook.save(alternate)
        workbook.close()
        with self.assertRaises(AggregationError) as caught:
            aggregate(alternate, [self.first], self.output)
        self.assertTrue(any(issue.code == "template_formula_not_supported" for issue in caught.exception.issues))
        self.assertFalse(self.output.exists())

    def test_unreferenced_source_sheets_are_not_counted(self):
        workbook = load_workbook(self.first)
        workbook.create_sheet("未使用的表")["A1"] = "=1/0"
        extra_source = self.root / "多表源.xlsx"
        workbook.save(extra_source)
        workbook.close()
        result = aggregate(self.template, [extra_source], self.output)
        self.assertEqual(result.source_count, 1)
        self.assertEqual(result.source_sheet_count, 1)

    def test_warnings_are_written_only_in_the_run_log(self):
        source = self.make_source("缺利率.xlsx", 900, None)
        aggregate(self.template, [self.first, source], self.output)
        workbook = load_workbook(self.output)
        try:
            self.assertEqual(workbook.sheetnames, ["业务明细", "运行日志"])
            text = "\n".join(str(cell.value) for row in workbook["运行日志"] for cell in row if cell.value is not None)
            self.assertIn("利率空白按0%", text)
            self.assertIn("业务明细!D5", text)
            self.assertIn("缺利率.xlsx", text)
        finally:
            workbook.close()

    def test_cross_sheet_weight_pairing_stays_in_same_file(self):
        template = Workbook()
        template.active.title = "不同输出表"
        template.active["F8"] = "WAVG('利率'!E7,'金额'!C5)"
        template_path = self.root / "跨表模板.xlsx"
        template.save(template_path)
        template.close()
        paths = []
        for index, (weight, rate) in enumerate(((100, .03), (900, .05))):
            workbook = Workbook()
            workbook.active.title = "金额"
            workbook.active["C5"] = weight
            workbook.create_sheet("利率")["E7"] = rate
            path = self.root / f"跨表源{index}.xlsx"
            workbook.save(path)
            workbook.close()
            paths.append(path)
        result = aggregate(template_path, paths, self.output)
        self.assertEqual(result.source_sheet_count, 4)
        self.assertEqual(result.template_sheet_count, 1)
        self.assertEqual(result.sum_rule_count, 0)
        self.assertEqual(result.wavg_rule_count, 1)
        workbook = load_workbook(self.output)
        try:
            self.assertAlmostEqual(workbook["不同输出表"]["F8"].value, .048, places=14)
        finally:
            workbook.close()

    def test_default_output_name_is_separate(self):
        path = default_output_path(self.template, self.root)
        self.assertNotEqual(path, self.template)
        self.assertEqual(path.suffix, ".xlsx")
        self.assertFalse(path.exists())


if __name__ == "__main__":
    unittest.main()
