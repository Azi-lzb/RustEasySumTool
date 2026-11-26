from contextlib import redirect_stdout
from decimal import Decimal
from hashlib import sha256
from io import StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from app import main
from summary_tool.configuration import DEFAULT_BINDINGS, load_configuration
from summary_tool.engine import aggregate
from summary_tool.models import AggregationError
from summary_tool.rules import RuleError, parse_rule


class ParameterRuleTests(unittest.TestCase):
    def test_omitted_arguments_keep_every_top_level_slot(self):
        cases = {
            "FUN_1": ("K4",),
            "FUN_1()": ("K4",),
            "FUN_1(   )": ("K4",),
            "FUN_1(,B5)": ("K4", "B5"),
            "FUN_1(A5,)": ("A5", "K4"),
            "FUN_1(,)": ("K4", "K4"),
            "FUN_1(A5,,C5)": ("A5", "K4", "C5"),
            "FUN_1(,,)": ("K4", "K4", "K4"),
            "FUN_1(A1:A3,)": ("A1:A3", "K4"),
            "fun_1( \t, \n B5,\t )": ("K4", "B5", "K4"),
            "F1_1": ("K4",),
            "f1_1()": ("K4",),
            "SUM(,)": ("K4", "K4"),
        }
        for expression, expected in cases.items():
            with self.subTest(expression=expression):
                rule = parse_rule(expression, "结果", "K4", DEFAULT_BINDINGS)
                self.assertEqual(tuple(ref.coordinate for ref in rule.references), expected)
                self.assertTrue(all(ref.sheet == "结果" for ref in rule.references))
                self.assertEqual(rule.expression, expression.strip())

    def test_default_address_uses_each_instruction_and_its_own_sheet(self):
        for sheet, coordinate in (("结果甲", "B7"), ("结果乙", "XFD1048576")):
            with self.subTest(sheet=sheet, coordinate=coordinate):
                rule = parse_rule("FUN_2(,'源, O''Brien(表)'!A5)", sheet, coordinate, DEFAULT_BINDINGS)
                self.assertEqual((rule.references[0].sheet, rule.references[0].coordinate), (sheet, coordinate))
                self.assertEqual(rule.references[1].sheet, "源, O'Brien(表)")
                rule = parse_rule("FUN_2('别的表'!A5,)", sheet, coordinate, DEFAULT_BINDINGS)
                self.assertEqual((rule.references[1].sheet, rule.references[1].coordinate), (sheet, coordinate))

    def test_legacy_defaults_and_explicit_overrides_are_independent_of_new_aliases(self):
        for name, target in (("f1_1", "SUM"), ("F1_2", "AVG"), ("f2_1", "WAVG"),
                             ("F1_3", "MAX"), ("F1_4", "MIN"), ("F1_5", "COUNT")):
            arguments = ",B5" if target == "WAVG" else ""
            with self.subTest(name=name):
                rule = parse_rule(f"{name}({arguments})", "表", "K4", {"FUN_1": "MAX"})
                self.assertEqual(rule.operation, target)
        rule = parse_rule("f1_1()", "表", "K4", {"f1_1": "MIN", "FUN_1": "MAX"})
        self.assertEqual(rule.operation, "MIN")

    def test_identifier_number_and_legacy_prefix_do_not_limit_signatures(self):
        for alias in ("FUN_1", "fun_999", "F1_1", "f2_1", "F12_345"):
            with self.subTest(alias=alias):
                rule = parse_rule(f"{alias}(A1,B1,C1)", "表", "K4", {alias: "sum"})
                self.assertEqual(rule.operation, "SUM")
                self.assertEqual(len(rule.references), 3)
                rule = parse_rule(f"{alias}(,B5)", "表", "K4", {alias: "WAVG"})
                self.assertEqual(rule.operation, "WAVG")
                self.assertEqual(len(rule.references), 2)

    def test_fixed_signature_errors_include_location_original_mapping_and_counts(self):
        for expression, actual in (("FUN_2()", 1), ("FUN_2", 1), ("FUN_2(A1,B1,C1)", 3)):
            with self.subTest(expression=expression), self.assertRaises(RuleError) as caught:
                parse_rule(expression, "模板页", "K4", DEFAULT_BINDINGS)
            for detail in ("模板页", "K4", expression, "FUN_2", "WAVG", "期望 2", f"实际 {actual}"):
                self.assertIn(detail, str(caught.exception))

    def test_ranges_are_one_parameter_and_keep_boundaries_and_row_major_order(self):
        rule = parse_rule("FUN_2('数值, 表'!$B$5:$C$6,'权重表'!D8:E9)", "结果", "K4", DEFAULT_BINDINGS)
        self.assertEqual(len(rule.references), 2)
        self.assertEqual([ref.shape for ref in rule.references], [(2, 2), (2, 2)])
        self.assertEqual([ref.coordinate for ref in rule.references[0].cells()], ["B5", "C5", "B6", "C6"])
        rule = parse_rule("WAVG(A1:A1,B1)", "结果", "K4")
        self.assertEqual([ref.shape for ref in rule.references], [(1, 1), (1, 1)])

    def test_ranges_with_equal_cell_counts_still_require_matching_shapes(self):
        for expression in ("WAVG(A1:A3,B1:D1)", "FUN_2(A1:B2,C1:C4)", "WAVG(A1:A3,)"):
            with self.subTest(expression=expression), self.assertRaises(RuleError) as caught:
                parse_rule(expression, "结果", "K4", DEFAULT_BINDINGS)
            self.assertIn("相同的行数和列数", str(caught.exception))

    def test_existing_scalar_algorithms_keep_their_signatures(self):
        for name in ("AVG", "MAX", "MIN", "COUNT"):
            with self.subTest(name=name):
                rule = parse_rule(f"{name}()", "表", "D9")
                self.assertEqual(rule.references[0].coordinate, "D9")
                for arguments in ("A1,B1", "A1:A3"):
                    with self.assertRaises(RuleError):
                        parse_rule(f"{name}({arguments})", "表", "D9")

    def test_unsupported_expressions_and_invalid_ranges_fail(self):
        for arguments in ("SUM(A1)", "A1+B1", "A1:B2:C3", "B2:A1", "A1:XFE2",
                          "A1:A1048577", "'[book.xlsx]表'!A1", "'未闭合!A1,B1", "1", '"A1"'):
            with self.subTest(arguments=arguments), self.assertRaises(RuleError):
                parse_rule(f"FUN_1({arguments})", "表", "K4", DEFAULT_BINDINGS)


class ParameterFileTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory(prefix="summary_parameters_")
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.template = self.root / "模板.xlsx"
        self.output = self.root / "输出.xlsx"
        self.config = self.root / "config.xlsx"
        self.write_config(DEFAULT_BINDINGS)

    def write_config(self, bindings):
        cells = {"A1": "函数别名", "B1": "映射函数"}
        for row, (alias, target) in enumerate(bindings.items(), 2):
            cells[f"A{row}"], cells[f"B{row}"] = alias, target
        self.workbook(self.config, {"函数映射": cells})

    def workbook(self, path, sheets):
        workbook = Workbook()
        try:
            for index, (title, cells) in enumerate(sheets.items()):
                sheet = workbook.active if index == 0 else workbook.create_sheet()
                sheet.title = title
                for coordinate, value in cells.items():
                    sheet[coordinate] = value
            workbook.save(path)
        finally:
            workbook.close()
        return path

    def run_batch(self, sources):
        protected = [self.template, self.config, *sources]
        before = {path: sha256(path.read_bytes()).digest() for path in protected}
        try:
            return aggregate(self.template, sources, self.output, config_path=self.config)
        finally:
            self.assertEqual(before, {path: sha256(path.read_bytes()).digest() for path in protected})

    def values(self, sheet, *coordinates):
        workbook = load_workbook(self.output)
        try:
            self.assertFalse(any(cell.data_type == "f" for ws in workbook for row in ws for cell in row))
            return tuple(workbook[sheet][coordinate].value for coordinate in coordinates)
        finally:
            workbook.close()

    def test_configuration_accepts_both_alias_formats_without_arity_constraints(self):
        self.write_config({"fun_99": "wavg", "F1_1": "WAVG", "F2_1": "SUM"})
        self.assertEqual(load_configuration(self.config).bindings,
                         {"FUN_99": "WAVG", "F1_1": "WAVG", "F2_1": "SUM"})

    def test_sum_ranges_overlap_duplicates_empty_slots_and_multiple_sources(self):
        self.workbook(self.template, {"表": {
            "K4": "FUN_1(,)", "K5": "SUM(A1:A3,A2,)",
            "K6": "fun_1(A1:B2,B2)", "K7": "F2_1(A1,B1,A2)",
        }})
        self.write_config({**DEFAULT_BINDINGS, "F2_1": "SUM"})
        sources = [self.workbook(self.root / f"源{index}.xlsx", {"表": {
            "A1": 1 * index, "A2": 2 * index, "A3": 3 * index,
            "B1": 4 * index, "B2": 5 * index, "K4": 10 * index, "K5": 20 * index,
        }}) for index in (1, 2)]
        result = self.run_batch(sources)
        self.assertEqual(self.values("表", "K4", "K5", "K6", "K7"), (60, 84, 51, 21))
        self.assertEqual(result.function_bindings, {"FUN_1": "SUM", "F2_1": "SUM"})

    def test_default_references_read_source_positions_in_each_sheet_context(self):
        self.workbook(self.template, {"甲表": {"K4": "FUN_2(,'外部, 表'!A5)", "K5": "FUN_1"},
                                       "乙表": {"B7": "FUN_3()", "B8": "FUN_4", "B9": "FUN_6()"}})
        sources = [self.workbook(self.root / f"源{index}.xlsx", {
            "外部, 表": {"A5": index}, "乙表": {"B7": 100 * index, "B8": index, "B9": 0},
            "甲表": {"K4": 10 * index, "K5": 7 * index},
        }) for index in (1, 3)]
        self.run_batch(sources)
        self.assertEqual(self.values("甲表", "K4", "K5"), (25, 28))
        self.assertEqual(self.values("乙表", "B7", "B8", "B9"), (200, 3, 2))
        self.assertEqual(load_configuration(self.config).bindings, DEFAULT_BINDINGS)

    def test_weighted_regions_accumulate_products_and_weights_across_sources(self):
        self.workbook(self.template, {"结果": {
            "K4": "FUN_2('值, O''Brien'!A1:B2,'权 重'!D4:E5)",
            "K5": "WAVG('值, O''Brien'!A1:A1,'权 重'!D4)",
        }})
        first = self.workbook(self.root / "甲.xlsx", {
            "值, O'Brien": {"A1": 1, "B1": 2, "A2": 3, "B2": 4},
            "权 重": {"D4": 1, "E4": 1, "D5": 1, "E5": 1},
        })
        second = self.workbook(self.root / "乙.xlsx", {
            "权 重": {"D4": 10, "E4": 20, "D5": 30, "E5": -5},
            "值, O'Brien": {"A1": 10, "B1": "2,000%", "A2": "  ", "B2": 0},
        })
        result = self.run_batch([first, second])
        region_value, scalar_value = self.values("结果", "K4", "K5")
        self.assertAlmostEqual(region_value, float(Decimal(510) / 59), places=13)
        self.assertAlmostEqual(scalar_value, float(Decimal(101) / 11), places=13)
        self.assertNotAlmostEqual(region_value, (2.5 + 500 / 55) / 2)
        self.assertEqual([(issue.code, issue.source_sheet, issue.source_cell) for issue in result.issues
                          if issue.severity == "warning"],
                         [("blank_rate_with_weight", "值, O'Brien", "A2"), ("negative_weight", "权 重", "E5")])

    def test_zero_total_weight_fails_for_blank_zero_and_cancelling_inputs(self):
        self.workbook(self.template, {"表": {"K4": "FUN_2(A1:A2,B1:B2)"}})
        for weights in ((None, None), (0, 0), (10, -10)):
            with self.subTest(weights=weights):
                source = self.workbook(self.root / "源.xlsx", {"表": {
                    "A1": 3, "A2": 5, "B1": weights[0], "B2": weights[1],
                }})
                with self.assertRaises(AggregationError) as caught:
                    self.run_batch([source])
                self.assertFalse(self.output.exists())
                self.assertTrue(any(issue.code == "zero_weight" and issue.severity == "error"
                                    for issue in caught.exception.issues))

    def test_zero_weight_in_one_file_is_combined_with_the_other_files(self):
        self.workbook(self.template, {"表": {"K4": "FUN_2(A1:A2,B1:B2)", "K5": "FUN_2(,)"}})
        first = self.workbook(self.root / "甲.xlsx", {"表": {"A1": 1, "A2": 3, "B1": 1, "B2": -1, "K5": 2}})
        second = self.workbook(self.root / "乙.xlsx", {"表": {"A1": 10, "A2": 20, "B1": 2, "B2": 0, "K5": 4}})
        result = self.run_batch([first, second])
        regional, repeated = self.values("表", "K4", "K5")
        self.assertEqual(regional, 9)
        self.assertAlmostEqual(repeated, 20 / 6, places=14)
        self.assertFalse(any(issue.code == "zero_weight" for issue in result.issues))

    def test_invalid_value_in_a_region_aborts_batch_with_exact_source_cell(self):
        self.workbook(self.template, {"表": {"K4": "SUM(A1:B2)"}})
        for invalid, code in (("非法数值", "invalid_number"), (True, "invalid_number"),
                              ("#DIV/0!", "excel_error"), ("=1+2", "formula_cache_missing")):
            with self.subTest(invalid=invalid):
                source = self.workbook(self.root / "源.xlsx", {"表": {"A1": 2, "B2": invalid}})
                with self.assertRaises(AggregationError) as caught:
                    self.run_batch([source])
                self.assertFalse(self.output.exists())
                issue = next(issue for issue in caught.exception.issues if issue.code == code)
                self.assertEqual((issue.source_sheet, issue.source_cell, issue.template_cell), ("表", "B2", "K4"))

    def test_all_static_errors_are_reported_before_opening_sources_and_cli_lists_all(self):
        expressions = {"K4": "FUN_2()", "K5": "FUN_2(A1,B1,C1)",
                       "K6": "WAVG(A1:A3,B1:D1)", "K7": "FUN_99(A1)"}
        self.workbook(self.template, {"表": expressions})
        source = self.workbook(self.root / "源.xlsx", {"表": {"A1": 1}})
        with patch("summary_tool.engine.SourceWorkbook", side_effect=AssertionError("静态失败不应打开源文件")):
            with self.assertRaises(AggregationError) as caught:
                self.run_batch([source])
            self.assertEqual({issue.template_cell for issue in caught.exception.issues}, set(expressions))
            output = StringIO()
            with redirect_stdout(output):
                self.assertEqual(main(["--template", str(self.template), "--sources", str(source),
                                       "--output", str(self.output), "--config", str(self.config)]), 2)
        for coordinate, expression in expressions.items():
            self.assertIn(coordinate, output.getvalue())
            self.assertIn(expression, output.getvalue())
        self.assertFalse(self.output.exists())
        self.assertEqual(list(self.root.glob(".汇总临时_*")), [])


if __name__ == "__main__":
    unittest.main()
