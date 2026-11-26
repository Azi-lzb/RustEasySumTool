from decimal import Decimal
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from types import MappingProxyType
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from summary_tool import functions
from summary_tool.engine import aggregate
from summary_tool.functions import FunctionResult, FunctionSpec
from summary_tool.models import AggregationError
from summary_tool.rules import RuleError, parse_rule


_BINDINGS = MappingProxyType({"F1_1": "SUM", "F1_2": "AVG", "F2_1": "WAVG"})


class AliasRuleTests(unittest.TestCase):
    def test_alias_case_and_absolute_reference_keep_canonical_operation(self):
        for expression in ("F1_1(A5)", "f1_1(a5)", " f1_1( $a$5 ) "):
            with self.subTest(expression=expression):
                rule = parse_rule(expression, "业务明细", "C5", bindings=_BINDINGS)
                self.assertIsNotNone(rule)
                self.assertEqual(rule.operation, "SUM")
                self.assertEqual(rule.function_name, "F1_1")
                self.assertEqual((rule.sheet, rule.coordinate), ("业务明细", "C5"))
                self.assertEqual(rule.expression, expression.strip())
                self.assertEqual(len(rule.references), 1)
                self.assertEqual(
                    (rule.references[0].sheet, rule.references[0].coordinate),
                    ("业务明细", "A5"),
                )
        direct = parse_rule("sum(A5)", "业务明细", "C5", bindings={"SUM": "AVG"})
        self.assertEqual((direct.operation, direct.function_name), ("SUM", "SUM"))

    def test_aliases_accept_quoted_cross_sheet_references(self):
        cases = (
            ("f1_2('源 数据'!$a$5)", "AVG", "F1_2", (("源 数据", "A5"),)),
            (
                "f2_1('利率, 一'!$b$7,'O''Brien 表'!$A$9)", "WAVG", "F2_1",
                (("利率, 一", "B7"), ("O'Brien 表", "A9")),
            ),
        )
        for expression, operation, name, expected in cases:
            with self.subTest(expression=expression):
                rule = parse_rule(expression, "结果", "C5", bindings=_BINDINGS)
                self.assertEqual((rule.operation, rule.function_name), (operation, name))
                self.assertEqual(
                    tuple((reference.sheet, reference.coordinate) for reference in rule.references),
                    expected,
                )

    def test_alias_parameter_counts_follow_the_registered_function(self):
        expressions = (
            "F1_2(A5，A6)", "F2_1(B5)", "F2_1(B5,A5,C5)", "F2_1()",
        )
        for expression in expressions:
            with self.subTest(expression=expression), self.assertRaises(RuleError):
                parse_rule(expression, "业务明细", "C5", bindings=_BINDINGS)

    def test_unknown_aliases_and_noncanonical_targets_are_rejected(self):
        cases = (
            ("F1_99(A5)", _BINDINGS),
            ("F0_1(A5)", _BINDINGS),
            ("F1_09(A5)", _BINDINGS),
            ("F1_9(A5)", MappingProxyType({"F1_9": "NOT_REGISTERED"})),
            ("F1_9(A5)", MappingProxyType({"F1_9": "F1_1"})),
            ("F1_9(A5)", MappingProxyType({"F1_9": "WAVG"})),
        )
        for expression, bindings in cases:
            with self.subTest(expression=expression, bindings=bindings), self.assertRaises(RuleError):
                parse_rule(expression, "业务明细", "C5", bindings=bindings)

    def test_aliases_reject_excel_formulas_and_executable_expressions(self):
        expressions = (
            "=F1_1(A5)", " = f2_1(B5,A5) ", "F1_1(A5", "F1_1(A5))",
            "F1_1(A5*B5)",
            "F1_1('[源.xlsx]业务明细'!A5)", "F1_1(__import__('os'))",
        )
        for expression in expressions:
            with self.subTest(expression=expression), self.assertRaises(RuleError):
                parse_rule(expression, "业务明细", "C5", bindings=_BINDINGS)


class AliasFileTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory(prefix="summary_alias_test_")
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.template = self.root / "模板.xlsx"
        self.output = self.root / "结果.xlsx"
        self.config = self.root / "config.xlsx"
        self.write_config(dict(_BINDINGS))
        self.make_template()

    def write_config(self, bindings):
        cells = {"A1": "函数别名", "B1": "映射函数"}
        for row, (name, target) in enumerate(bindings.items(), 2):
            cells[f"A{row}"], cells[f"B{row}"] = name, target
        self.write_workbook(self.config, {"函数映射": cells})

    def write_workbook(self, path, sheets):
        workbook = Workbook()
        try:
            for index, (title, values) in enumerate(sheets.items()):
                sheet = workbook.active if index == 0 else workbook.create_sheet()
                sheet.title = title
                for coordinate, value in values.items():
                    sheet[coordinate] = value
            workbook.save(path)
        finally:
            workbook.close()
        return path

    def make_template(self, rules=None):
        cells = {"B2": "别名框架独立测试", "G5": "固定说明"}
        cells.update({"C5": "F1_1(A5)"} if rules is None else rules)
        return self.write_workbook(self.template, {"业务明细": cells})

    def make_source(self, name, value=None, rate=None, *, sheets=None, extra_sheet=False):
        contents = ({"业务明细": {"A5": value, "B5": rate}}
                    if sheets is None else dict(sheets))
        if extra_sheet:
            contents["源独有"] = {"B2": "此源工作表不能复制到结果"}
        return self.write_workbook(self.root / name, contents)

    def make_pair(self, *, extra_sheet=False):
        return [self.make_source("甲.xlsx", 100, "3%", extra_sheet=extra_sheet),
                self.make_source("乙.xlsx", 900, "5%")]

    def run_batch(self, sources, *, output=None, progress=None):
        return aggregate(
            self.template, sources, output or self.output,
            progress=progress, config_path=self.config,
        )

    def read_values(self, *coordinates, output=None):
        workbook = load_workbook(output or self.output, data_only=False)
        try:
            return tuple(workbook["业务明细"][coordinate].value for coordinate in coordinates)
        finally:
            workbook.close()

    def assert_batch_fails(self, sources, *, code=None):
        with self.assertRaises(AggregationError) as caught:
            self.run_batch(sources)
        self.assertFalse(self.output.exists())
        self.assertEqual(list(self.root.glob(".汇总临时_*")), [])
        if code is not None:
            matching = [issue for issue in caught.exception.issues if issue.code == code]
            self.assertTrue(matching, caught.exception.issues)
            self.assertTrue(all(issue.severity == "error" for issue in matching))
            self.assertEqual(
                {(issue.template_sheet, issue.template_cell) for issue in matching},
                {("业务明细", "C5")},
            )
        return caught.exception

    def test_explicit_missing_config_fails_even_for_direct_functions(self):
        self.config.unlink()
        sources = self.make_pair()
        for expression in ("F1_1(A5)", "SUM(A5)"):
            with self.subTest(expression=expression):
                self.make_template({"C5": expression})
                self.assert_batch_fails(sources)
                self.assertFalse(self.config.exists())

    def test_multiple_aliases_and_direct_names_count_real_algorithms(self):
        self.write_config({**_BINDINGS, "F1_7": "SUM", "F1_8": "AVG"})
        self.make_template({
            "C5": "F1_1(A5)", "D5": "F1_7(A5)", "E5": "SUM(A5)",
            "F5": "F1_2(A5)", "G5": "AVG(A5)",
            "H5": "F2_1(B5,A5)", "I5": "WAVG(B5,A5)",
        })
        result = self.run_batch(self.make_pair())
        self.assertEqual(self.read_values("C5", "D5", "E5", "F5", "G5"),
                         (1000, 1000, 1000, 500, 500))
        for value in self.read_values("H5", "I5"):
            self.assertAlmostEqual(value, 0.048, places=15)
        self.assertEqual(result.rule_count, 7)
        self.assertEqual(result.function_counts, {"SUM": 3, "AVG": 2, "WAVG": 2})
        self.assertEqual((result.sum_rule_count, result.avg_rule_count, result.wavg_rule_count),
                         (3, 2, 2))
        self.assertEqual(result.function_bindings,
                         {**_BINDINGS, "F1_7": "SUM"})
        self.assertNotIn("F1_8", result.function_bindings)

    def test_rebinding_sum_to_average_changes_values_without_changing_template(self):
        sources = self.make_pair()
        protected = [self.template, *sources]
        before = {path: sha256(path.read_bytes()).hexdigest() for path in protected}
        sum_output = self.root / "求和结果.xlsx"
        sum_config = self.config.read_bytes()
        sum_result = self.run_batch(sources, output=sum_output)
        self.assertEqual(self.config.read_bytes(), sum_config)
        self.write_config({"F1_1": "AVG"})
        avg_config = self.config.read_bytes()
        avg_output = self.root / "平均结果.xlsx"
        avg_result = self.run_batch(sources, output=avg_output)
        self.assertEqual(self.config.read_bytes(), avg_config)
        self.assertEqual(self.read_values("C5", output=sum_output), (1000,))
        self.assertEqual(self.read_values("C5", output=avg_output), (500,))
        self.assertEqual(sum_result.function_bindings, {"F1_1": "SUM"})
        self.assertEqual(avg_result.function_bindings, {"F1_1": "AVG"})
        self.assertEqual(sum_result.function_counts, {"SUM": 1})
        self.assertEqual(avg_result.function_counts, {"AVG": 1})
        self.assertEqual(before, {path: sha256(path.read_bytes()).hexdigest() for path in protected})

    def test_config_changed_during_progress_applies_only_to_the_next_batch(self):
        sources = self.make_pair()
        updates = []
        template_hash = sha256(self.template.read_bytes()).hexdigest()

        def change_config(done, total, message):
            if done == 1 and not updates:
                self.write_config({"F1_1": "AVG"})
                updates.append((done, total))

        current_output = self.root / "当前批.xlsx"
        current = self.run_batch(sources, output=current_output, progress=change_config)
        following_output = self.root / "下一批.xlsx"
        following = self.run_batch(sources, output=following_output)
        self.assertEqual(updates, [(1, 3)])
        self.assertEqual(self.read_values("C5", output=current_output), (1000,))
        self.assertEqual(self.read_values("C5", output=following_output), (500,))
        self.assertEqual(current.function_bindings, {"F1_1": "SUM"})
        self.assertEqual(following.function_bindings, {"F1_1": "AVG"})
        self.assertEqual(current.function_counts, {"SUM": 1})
        self.assertEqual(following.function_counts, {"AVG": 1})
        self.assertEqual(sha256(self.template.read_bytes()).hexdigest(), template_hash)

    def test_direct_sum_is_not_rebound_when_f1_1_means_average(self):
        self.write_config({"F1_1": "AVG"})
        self.make_template({"C5": "SUM(A5)", "D5": "F1_1(A5)"})
        result = self.run_batch(self.make_pair())
        self.assertEqual(self.read_values("C5", "D5"), (1000, 500))
        self.assertEqual(result.function_bindings, {"F1_1": "AVG"})
        self.assertEqual(result.function_counts, {"SUM": 1, "AVG": 1})
        self.assertEqual((result.sum_rule_count, result.avg_rule_count), (1, 1))

    def test_direct_functions_work_without_config_but_new_aliases_require_bindings(self):
        self.config.unlink()
        self.make_template({
            "C5": "SUM(A5)", "D5": "AVG(A5)", "E5": "WAVG(B5,A5)",
        })
        sources = self.make_pair()
        # Only this legacy case omits config_path; patch the root to avoid the project config.
        with patch("summary_tool.configuration.application_root", return_value=self.root):
            result = aggregate(self.template, sources, self.output)
            self.make_template({"C5": "FUN_1(A5)"})
            alias_output = self.root / "未配置别名.xlsx"
            with self.assertRaises(AggregationError) as caught:
                aggregate(self.template, sources, alias_output)
        self.assertEqual(self.read_values("C5", "D5"), (1000, 500))
        self.assertAlmostEqual(self.read_values("E5")[0], 0.048, places=15)
        self.assertEqual(result.source_count, 2)
        self.assertEqual(result.rule_count, 3)
        self.assertEqual(result.function_bindings, {})
        self.assertEqual(result.function_counts, {"SUM": 1, "AVG": 1, "WAVG": 1})
        self.assertIsNone(result.config_path)
        self.assertFalse(self.config.exists())
        self.assertFalse(alias_output.exists())
        self.assertTrue(any(issue.code == "invalid_rule" and
                            (issue.template_sheet, issue.template_cell) == ("业务明细", "C5")
                            for issue in caught.exception.issues), caught.exception.issues)

    def test_invalid_targets_and_builtin_overrides_never_publish(self):
        self.make_template({"C5": "SUM(A5)"})
        sources = self.make_pair()
        marker = self.root / "配置代码不可执行.txt"
        payload = f"__import__('pathlib').Path({str(marker)!r}).write_text('executed')"
        configurations = (
            {"F1_9": "NOT_REGISTERED"},
            {"F1_9": "F1_1"}, {"SUM": "AVG"}, {"avg": "SUM"},
            {"WAVG": "AVG"}, {"F1_9": payload},
        )
        protected = [self.template, *sources]
        before = {path: sha256(path.read_bytes()).hexdigest() for path in protected}
        for bindings in configurations:
            with self.subTest(bindings=bindings):
                self.write_config(bindings)
                config_bytes = self.config.read_bytes()
                self.assert_batch_fails(sources)
                self.assertFalse(marker.exists())
                self.assertEqual(self.config.read_bytes(), config_bytes)
                self.assertEqual(before,
                                 {path: sha256(path.read_bytes()).hexdigest() for path in protected})

    def test_weighted_alias_blank_rate_notice_locates_its_cross_sheet_parameter(self):
        self.make_template({
            "C5": "F2_1('利率, 一'!$B$7,'O''Brien 表'!$A$9)",
        })
        first = self.make_source("正常.xlsx", sheets={
            "利率, 一": {"B7": "4%"}, "O'Brien 表": {"A9": 100},
        })
        missing = self.make_source("缺利率.xlsx", sheets={
            "利率, 一": {"B7": None}, "O'Brien 表": {"A9": 900},
        })
        result = self.run_batch([first, missing])
        self.assertAlmostEqual(self.read_values("C5")[0], 0.004, places=15)
        notices = [issue for issue in result.issues if issue.code == "blank_rate_with_weight"]
        self.assertEqual(len(notices), 1)
        notice = notices[0]
        self.assertEqual(notice.severity, "warning")
        self.assertEqual(Path(notice.source_file), missing.resolve())
        self.assertEqual((notice.source_sheet, notice.source_cell), ("利率, 一", "B7"))
        self.assertEqual((notice.template_sheet, notice.template_cell), ("业务明细", "C5"))
        self.assertEqual(result.function_counts, {"WAVG": 1})

    def test_weighted_alias_negative_weight_notice_locates_second_parameter(self):
        self.make_template({"C5": "F2_1($B$5,$A$5)"})
        first = self.make_source("正权重.xlsx", 100, "3%")
        negative = self.make_source("负权重.xlsx", -20, "5%")
        result = self.run_batch([first, negative])
        self.assertAlmostEqual(self.read_values("C5")[0], 0.025, places=15)
        notices = [issue for issue in result.issues if issue.code == "negative_weight"]
        self.assertEqual(len(notices), 1)
        notice = notices[0]
        self.assertEqual(notice.severity, "warning")
        self.assertEqual(Path(notice.source_file), negative.resolve())
        self.assertEqual((notice.source_sheet, notice.source_cell), ("业务明细", "A5"))
        self.assertEqual((notice.template_sheet, notice.template_cell), ("业务明细", "C5"))
        self.assertFalse(any(issue.code == "zero_weight" for issue in result.issues))

    def test_weighted_alias_cancelled_weights_notice_is_template_level(self):
        self.make_template({"C5": "F2_1(B5,A5)"})
        first = self.make_source("正权重.xlsx", 100, "3%")
        negative = self.make_source("抵消权重.xlsx", -100, "5%")
        error = self.assert_batch_fails([first, negative])
        notices = [issue for issue in error.issues if issue.code == "zero_weight"]
        self.assertEqual(len(notices), 1)
        notice = notices[0]
        self.assertEqual(notice.severity, "error")
        self.assertEqual((notice.source_file, notice.source_sheet, notice.source_cell), ("", "", ""))
        self.assertEqual((notice.template_sheet, notice.template_cell), ("业务明细", "C5"))
        self.assertTrue(any(issue.code == "negative_weight" and issue.source_cell == "A5"
                            for issue in error.issues))

    def test_average_alias_file_count_includes_blanks_but_excludes_duplicates_and_locks(self):
        self.write_config({"F1_1": "AVG"})
        first = self.make_source("有值.xlsx", 100)
        blank = self.make_source("空白.xlsx")
        third = self.make_source("第三份.xlsx", 200)
        corrupt_lock = self.root / "~$损坏.xlsx"
        corrupt_lock.write_bytes(b"not an xlsx archive")
        missing_lock = self.root / "~$不存在.xlsx"
        base = [first, first, blank, blank, corrupt_lock, missing_lock]
        for sources, count, expected in ((base, 2, 50), (base + [third], 3, 100)):
            with self.subTest(effective_source_count=count):
                output = self.root / f"有效文件{count}.xlsx"
                result = self.run_batch(sources, output=output)
                self.assertEqual(self.read_values("C5", output=output), (expected,))
                self.assertEqual(result.source_count, count)
                self.assertEqual(result.source_sheet_count, count)
                self.assertEqual(result.function_counts, {"AVG": 1})
                self.assertEqual(result.function_bindings, {"F1_1": "AVG"})
                self.assertEqual((result.sum_rule_count, result.avg_rule_count, result.wavg_rule_count),
                                 (0, 1, 0))
                self.assertEqual(sum(issue.code == "duplicate_source" for issue in result.issues), 2)
                self.assertEqual(sum(issue.code == "lock_file" for issue in result.issues), 2)
                self.assertFalse(any(issue.severity in {"warning", "error"} for issue in result.issues))
        self.assertFalse(missing_lock.exists())

    def test_output_has_two_sheets_no_formulas_and_only_actual_binding_in_log(self):
        self.write_config({"f1_1": "avg", "F1_8": "SUM"})
        self.make_template({"C5": "f1_1($a$5)"})
        sources = self.make_pair(extra_sheet=True)
        protected = [self.template, self.config, *sources]
        before = {path: sha256(path.read_bytes()).hexdigest() for path in protected}
        result = self.run_batch(sources)
        self.assertEqual(result.function_bindings, {"F1_1": "AVG"})
        self.assertEqual(result.config_path, self.config.resolve())
        self.assertEqual((result.source_count, result.source_sheet_count, result.template_sheet_count),
                         (2, 2, 1))
        self.assertEqual(before, {path: sha256(path.read_bytes()).hexdigest() for path in protected})
        workbook = load_workbook(self.output, data_only=False)
        try:
            self.assertEqual(workbook.sheetnames, ["业务明细", "运行日志"])
            self.assertEqual(workbook["业务明细"]["C5"].value, 500)
            self.assertEqual(workbook["业务明细"]["B2"].value, "别名框架独立测试")
            self.assertEqual(workbook["业务明细"]["G5"].value, "固定说明")
            self.assertEqual(sum(cell.data_type == "f" for sheet in workbook.worksheets
                                 for row in sheet.iter_rows() for cell in row), 0)
            rows = [" ".join(str(value) for value in row if value is not None)
                    for row in workbook["运行日志"].iter_rows(values_only=True)]
            log = "\n".join(rows)
            self.assertIn(str(self.config.resolve()), log)
            self.assertTrue(any("F1_1" in row and "AVG" in row for row in rows), log)
            self.assertNotIn("F1_8", log)
        finally:
            workbook.close()

    def test_runtime_registry_extension_needs_no_parser_or_engine_branch(self):
        instances = []

        class TestDoubleCalculator:
            def __init__(self):
                self.total = Decimal(0)
                self.add_count = 0
                self.source_count = None
                instances.append(self)

            def add(self, values):
                self.total += values[0].value * 2
                self.add_count += 1
                return ()

            def finish(self, source_count):
                self.source_count = source_count
                return FunctionResult(self.total)

        spec = FunctionSpec("TESTDOUBLE", ("取数单元格",), TestDoubleCalculator)
        with patch.dict(functions.FUNCTIONS, {"TESTDOUBLE": spec}):
            self.write_config({"F1_9": "TESTDOUBLE"})
            self.make_template({"C5": "f1_9($a$5)"})
            for expression, name in (("f1_9(A5)", "F1_9"), ("testdouble(A5)", "TESTDOUBLE")):
                rule = parse_rule(expression, "业务明细", "C5", bindings={"F1_9": "TESTDOUBLE"})
                self.assertEqual((rule.operation, rule.function_name), ("TESTDOUBLE", name))
            result = self.run_batch(self.make_pair())
            self.assertEqual(self.read_values("C5"), (2000,))
            self.assertEqual(result.function_bindings, {"F1_9": "TESTDOUBLE"})
            self.assertEqual(result.function_counts, {"TESTDOUBLE": 1})
            self.assertEqual((result.sum_rule_count, result.avg_rule_count, result.wavg_rule_count),
                             (0, 0, 0))
            self.assertEqual((result.source_count, result.rule_count), (2, 1))
            self.assertEqual(len(instances), 1)
            self.assertEqual((instances[0].add_count, instances[0].source_count), (2, 2))
        self.assertIsNone(functions.get_function("TESTDOUBLE"))

    def test_invalid_alias_rules_report_template_location_without_output(self):
        sources = self.make_pair()
        cases = (
            ("F1_99(A5)", "invalid_rule"),
            ("F1_2(A5,B5)", "invalid_rule"),
            ("F2_1(B5)", "invalid_rule"),
            ("=F1_1(A5)", "template_formula_not_supported"),
        )
        for expression, code in cases:
            with self.subTest(expression=expression):
                self.make_template({"C5": expression})
                before = sha256(self.template.read_bytes()).hexdigest()
                self.assert_batch_fails(sources, code=code)
                self.assertEqual(sha256(self.template.read_bytes()).hexdigest(), before)


if __name__ == "__main__":
    unittest.main()
