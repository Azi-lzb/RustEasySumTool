from decimal import Decimal
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook, load_workbook

from summary_tool.configuration import DEFAULT_BINDINGS
from summary_tool.engine import aggregate
from summary_tool.functions import get_function
from summary_tool.models import AggregationError
from summary_tool.reader import decimal_value
from summary_tool.rules import RuleError, parse_rule


class ProductRatioCalculatorTests(unittest.TestCase):
    def test_decimal_products_blanks_zero_negative_and_independent_state(self):
        calculator = get_function("SUMPRODUCT").factory()
        independent = get_function("SUMPRODUCT").factory()
        for pair in (("0.1", "0.2"), (None, 10), (0, 200), (-2, 3), ("5%", "1,000")):
            self.assertEqual(calculator.add(tuple(map(decimal_value, pair))), ())
        self.assertEqual(calculator.finish(5).value, Decimal("44.02"))
        self.assertEqual(independent.finish(0).value, 0)

    def test_ratio_uses_totals_and_allows_individual_zero_denominators(self):
        calculator = get_function("RATIO").factory()
        for pair in ((2, 0), (6, 10), (None, 10)):
            self.assertEqual(calculator.add(tuple(map(decimal_value, pair))), ())
        result = calculator.finish(3)
        self.assertEqual(result.value, Decimal("0.4"))
        self.assertEqual(result.notices, ())
        signed = get_function("RATIO").factory()
        signed.add(tuple(map(decimal_value, (-2, -10))))
        self.assertEqual(signed.finish(1).value, Decimal("0.2"))
        zero = get_function("RATIO").factory()
        zero.add(tuple(map(decimal_value, (0, 10))))
        self.assertEqual(zero.finish(1).value, 0)

    def test_zero_total_denominator_returns_error_notice(self):
        for pairs in (((1, None),), ((1, 0),), ((2, 10), (6, -10))):
            with self.subTest(pairs=pairs):
                calculator = get_function("RATIO").factory()
                for pair in pairs:
                    calculator.add(tuple(map(decimal_value, pair)))
                result = calculator.finish(len(pairs))
                self.assertEqual([(notice.severity, notice.code) for notice in result.notices],
                                 [("error", "zero_denominator")])

    def test_fixed_signatures_alias_case_defaults_and_cross_sheet_references(self):
        for operation, alias in (("SUMPRODUCT", "FUN_7"), ("RATIO", "FUN_8")):
            for name in (operation.lower(), alias.lower()):
                with self.subTest(name=name):
                    rule = parse_rule(f"{name}(,'输入, O''Brien'!$B$5)", "结果", "K4", DEFAULT_BINDINGS)
                    self.assertEqual(rule.operation, operation)
                    self.assertEqual([(ref.sheet, ref.coordinate) for ref in rule.references],
                                     [("结果", "K4"), ("输入, O'Brien", "B5")])
                    for arguments, count in (("", 1), ("A1", 1), ("A1,B1,C1", 3)):
                        expression = f"{name}({arguments})"
                        with self.assertRaises(RuleError) as caught:
                            parse_rule(expression, "结果", "K4", DEFAULT_BINDINGS)
                        for text in (expression, "结果", "K4", operation, "期望 2", f"实际 {count}"):
                            self.assertIn(text, str(caught.exception))
                    with self.assertRaisesRegex(RuleError, "单个单元格"):
                        parse_rule(f"{name}(A1:A3,B1:B3)", "结果", "K4", DEFAULT_BINDINGS)


class ProductRatioFileTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory(prefix="summary_product_ratio_")
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.template = self.root / "模板.xlsx"
        self.output = self.root / "结果.xlsx"
        self.config = self.root / "config.xlsx"
        configuration = Workbook()
        configuration.active.title = "函数映射"
        configuration.active.append(["函数别名", "映射函数"])
        for row in DEFAULT_BINDINGS.items():
            configuration.active.append(row)
        configuration.save(self.config)
        configuration.close()

    def write_workbook(self, name, cells):
        path = self.root / name
        workbook = Workbook()
        try:
            workbook.active.title = "报表"
            for coordinate, value in cells.items():
                workbook.active[coordinate] = value
            workbook.active["G5"].number_format = "0.0%"
            workbook.save(path)
        finally:
            workbook.close()
        return path

    def run_batch(self, sources):
        inputs = [self.template, self.config, *sources]
        before = {path: sha256(path.read_bytes()).digest() for path in inputs}
        try:
            return aggregate(self.template, sources, self.output, config_path=self.config)
        finally:
            self.assertEqual(before, {path: sha256(path.read_bytes()).digest() for path in inputs})

    def test_multifile_products_and_ratios_match_aliases_and_preserve_result_format(self):
        self.write_workbook(self.template.name, {
            "D5": "sumproduct(A5,B5)", "E5": "fun_7(A5,B5)",
            "F5": "ratio(A5,B5)", "G5": "FUN_8(A5,B5)",
        })
        first = self.write_workbook("甲.xlsx", {"A5": 2, "B5": 10})
        second = self.write_workbook("乙.xlsx", {"A5": 6, "B5": 90})
        result = self.run_batch([first, first, second])
        self.assertEqual(result.source_count, 2)
        self.assertEqual(result.function_counts, {"SUMPRODUCT": 2, "RATIO": 2})
        self.assertEqual(result.function_bindings, {"FUN_7": "SUMPRODUCT", "FUN_8": "RATIO"})
        workbook = load_workbook(self.output)
        try:
            self.assertEqual([workbook["报表"][cell].value for cell in ("D5", "E5", "F5", "G5")],
                             [560, 560, 0.08, 0.08])
            self.assertEqual(workbook["报表"]["G5"].number_format, "0.0%")
            log = "\n".join(str(cell.value) for row in workbook["运行日志"] for cell in row)
            for text in ("SUMPRODUCT", "RATIO", "FUN_7", "FUN_8"):
                self.assertIn(text, log)
            self.assertFalse(any(cell.data_type == "f" for sheet in workbook for row in sheet for cell in row))
        finally:
            workbook.close()

    def test_omitted_parameters_and_duplicate_references_read_source_coordinates(self):
        self.write_workbook(self.template.name, {
            "K4": "FUN_7(,B5)", "L4": "FUN_8(A5,)", "K5": "FUN_7(,)", "L5": "FUN_8(,)",
        })
        sources = [self.write_workbook(f"源{index}.xlsx", {
            "K4": value, "A5": value, "B5": denominator, "L4": denominator,
            "K5": index + 3, "L5": index + 3,
        }) for index, (value, denominator) in enumerate(((2, 10), (6, 90)))]
        self.run_batch(sources)
        workbook = load_workbook(self.output, data_only=True)
        try:
            self.assertEqual([workbook["报表"][cell].value for cell in ("K4", "L4", "K5", "L5")],
                             [560, 0.08, 25, 1])
        finally:
            workbook.close()

    def test_blank_and_individual_zero_denominator_keep_all_numerators(self):
        self.write_workbook(self.template.name, {"D5": "SUMPRODUCT(A5,B5)", "G5": "RATIO(A5,B5)"})
        sources = [self.write_workbook(f"源{index}.xlsx", {"A5": numerator, "B5": denominator})
                   for index, (numerator, denominator) in enumerate(((2, 0), (6, 10), (None, 10)))]
        self.run_batch(sources)
        workbook = load_workbook(self.output, data_only=True)
        try:
            self.assertEqual(workbook["报表"]["D5"].value, 60)
            self.assertEqual(workbook["报表"]["G5"].value, 0.4)
        finally:
            workbook.close()

    def test_zero_total_denominator_aborts_entire_batch(self):
        self.write_workbook(self.template.name, {"D5": "FUN_7(A5,B5)", "G5": "FUN_8(A5,B5)"})
        for denominators in ((None, None), (0, 0), (10, -10)):
            with self.subTest(denominators=denominators):
                sources = [self.write_workbook(f"源{index}.xlsx", {"A5": index + 2, "B5": denominator})
                           for index, denominator in enumerate(denominators)]
                with self.assertRaises(AggregationError) as caught:
                    self.run_batch(sources)
                error = next(issue for issue in caught.exception.issues if issue.code == "zero_denominator")
                self.assertEqual((error.severity, error.template_sheet, error.template_cell), ("error", "报表", "G5"))
                self.assertFalse(self.output.exists())
                self.assertEqual(list(self.root.glob(".汇总临时_*")), [])

    def test_invalid_source_value_in_either_parameter_is_reported(self):
        for operation in ("SUMPRODUCT", "RATIO"):
            self.write_workbook(self.template.name, {"D5": f"{operation}(A5,B5)"})
            for coordinate in ("A5", "B5"):
                with self.subTest(operation=operation, coordinate=coordinate):
                    source = self.write_workbook("错误.xlsx", {"A5": 1, "B5": 2, coordinate: "非法数值"})
                    with self.assertRaises(AggregationError) as caught:
                        self.run_batch([source])
                    error = next(issue for issue in caught.exception.issues if issue.code == "invalid_number")
                    self.assertEqual((error.source_sheet, error.source_cell), ("报表", coordinate))
                    self.assertFalse(self.output.exists())


if __name__ == "__main__":
    unittest.main()
