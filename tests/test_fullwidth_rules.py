from contextlib import redirect_stdout
from hashlib import sha256
from io import StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from summary_tool.cli import print_failure
from summary_tool.configuration import DEFAULT_BINDINGS, LEGACY_BINDINGS
from summary_tool.engine import aggregate
from summary_tool.functions import function_names, get_function
from summary_tool.models import AggregationError
from summary_tool.rules import RuleError, parse_rule


class FullwidthRuleTests(unittest.TestCase):
    def test_direct_functions_and_aliases_accept_fullwidth_parentheses(self):
        bindings = {**DEFAULT_BINDINGS, **LEGACY_BINDINGS}
        names = [(name, name) for name in function_names()] + list(bindings.items())
        for name, operation in names:
            arguments = "$A$5" if get_function(operation).arity == 1 else "$A$5,B5"
            expression = f" {name.lower()}（{arguments}） "
            with self.subTest(expression=expression):
                rule = parse_rule(expression, "报表", "K4", bindings)
                ascii_rule = parse_rule(f"{name}({arguments})", "报表", "K4", bindings)
                self.assertEqual(rule.operation, operation)
                self.assertEqual(rule.references, ascii_rule.references)
                self.assertEqual(rule.expression, expression.strip())

    def test_mixed_parentheses_and_empty_slots_follow_existing_rules(self):
        cases = {
            "SUM（A5)": ("A5",),
            "SUM(A5）": ("A5",),
            "AVG（　）": ("K4",),
            "FUN_1（，B5，）": ("K4", "B5", "K4"),
            "FUN_1（A1:A3，）": ("A1:A3", "K4"),
            "FUN_2（，B5）": ("K4", "B5"),
        }
        for expression, expected in cases.items():
            with self.subTest(expression=expression):
                rule = parse_rule(expression, "报表", "K4", DEFAULT_BINDINGS)
                self.assertEqual(tuple(ref.coordinate for ref in rule.references), expected)

    def test_sheet_names_keep_their_original_parentheses_and_quoted_commas(self):
        for reference, sheet in (("业务（含税）!A5", "业务（含税）"),
                                 ("'业务(含税)'!A5", "业务(含税)"),
                                 ("'原（甲）, O''Brien'!A5", "原（甲）, O'Brien"),
                                 ("'资料（备注），表'!A5", "资料（备注），表")):
            with self.subTest(reference=reference):
                rule = parse_rule(f"SUM（{reference}）", "报表", "K4")
                self.assertEqual(rule.references[0].sheet, sheet)

    def test_malformed_parentheses_are_reported_with_original_text_and_location(self):
        for expression in ("SUM（A5", "SUM A5）", "SUM）", "SUM（A5））", "SUM（A5） extra",
                           "FUN_1（A5", "FUN_1 A5）", "RATIO（A5）", "FUN_2（）"):
            with self.subTest(expression=expression), self.assertRaises(RuleError) as caught:
                parse_rule(expression, "报表", "K4", DEFAULT_BINDINGS)
            for text in (expression, "报表", "K4"):
                self.assertIn(text, str(caught.exception))

    def test_nested_functions_unknown_functions_and_equals_are_rejected(self):
        for expression in ("SUM（SUM（A5））", "SUM（SUM（A5，B5），C5）", "SUMM（A5）",
                           "=AVG（A5）", "=FUN_1（A5）"):
            with self.subTest(expression=expression), self.assertRaises(RuleError):
                parse_rule(expression, "报表", "K4", DEFAULT_BINDINGS)


class FullwidthFileTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory(prefix="summary_fullwidth_")
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

    def write_workbook(self, path, sheets):
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

    def test_fullwidth_rules_compute_from_sources_and_preserve_input_files(self):
        self.write_workbook(self.template, {"报表": {
            "K4": "SUM（A5）", "K5": "AVG（）", "K6": "SUM（'明细（含税）'!A1）",
            "K7": "FUN_8（A5，B5）", "K8": "FUN_7（，B5）",
        }})
        sources = [self.write_workbook(self.root / f"源{index}.xlsx", {
            "报表": {"A5": value, "B5": denominator, "K5": average, "K8": value},
            "明细（含税）": {"A1": 3}, "明细(含税)": {"A1": 999},
        }) for index, (value, denominator, average) in enumerate(((2, 10, 100), (6, 90, None)))]
        paths = [self.template, self.config, *sources]
        before = {path: sha256(path.read_bytes()).digest() for path in paths}
        aggregate(self.template, sources, self.output, config_path=self.config)
        self.assertEqual(before, {path: sha256(path.read_bytes()).digest() for path in paths})
        workbook = load_workbook(self.output, data_only=True)
        try:
            self.assertEqual([workbook["报表"][f"K{row}"].value for row in range(4, 9)],
                             [8, 50, 6, 0.08, 560])
        finally:
            workbook.close()

    def test_bad_rules_are_listed_together_before_sources_are_opened(self):
        expressions = {"A1": "SUM（A5", "B1": "WAVG（A5）", "C1": "SUMM（A5）", "D1": "SUM A5）"}
        self.write_workbook(self.template, {"报表": expressions})
        source = self.write_workbook(self.root / "源.xlsx", {"报表": {"A5": 1}})
        with patch("summary_tool.engine.SourceWorkbook", side_effect=AssertionError("静态错误应在读取来源前报出")):
            with self.assertRaises(AggregationError) as caught:
                aggregate(self.template, [source], self.output, config_path=self.config)
        self.assertEqual({issue.template_cell for issue in caught.exception.issues}, set(expressions))
        output = StringIO()
        with redirect_stdout(output):
            print_failure(caught.exception)
        for coordinate, expression in expressions.items():
            self.assertIn(coordinate, output.getvalue())
            self.assertIn(expression, output.getvalue())
        self.assertFalse(self.output.exists())


if __name__ == "__main__":
    unittest.main()
