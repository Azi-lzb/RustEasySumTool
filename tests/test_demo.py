"""Reusable v0.6 alias fixtures, an independent oracle, and acceptance tests.

Run generation-only tests with
``python -m unittest discover -s tests -p test_demo.py -k DemoGenerationTests``.
Every fixture uses its own temporary bindings.xlsx, never user configuration.
Normal sources and results contain values, not Excel formulas. No recalculation
is used, so the deliberately uncached error formula must stay uncached.
"""

from __future__ import annotations

from contextlib import redirect_stdout
from copy import copy
from decimal import Decimal, localcontext
import hashlib
from io import StringIO
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from create_demo import main as create_demo_main
from summary_tool.demo import create_demo_assets
from summary_tool.templates.base import FONT_NAME, FORMATS


# Test-owned defaults: do not import application defaults or read root config.
_EXPECTED_BINDINGS = {
    "FUN_1": "SUM", "FUN_3": "AVG", "FUN_2": "WAVG",
    "FUN_4": "MAX", "FUN_5": "MIN", "FUN_6": "COUNT",
}
_EXPECTED_COUNTS = {"SUM": 15, "AVG": 9, "WAVG": 10, "MAX": 9, "MIN": 9, "COUNT": 9}
_ORIGINAL_COMMENTS = {
    "C5": ("工具指令不加等号，从各源文件的同名表、指定坐标取数。默认绑定FUN_1=SUM、FUN_3=AVG、FUN_2=WAVG；"
           "运行时仅由主程序同目录config.xlsx的函数映射页配置，模板目录不另放配置副本。"
           "切换绑定后须核对表头、单位和数值格式。"),
    "D7": "FUN_2默认绑定WAVG。有金额但缺利率的机构须产生提醒；空白利率按0%，其权重仍计入该项加权分母。",
    "G5": "FUN_3默认绑定AVG：取三家机构同名表C列余额之和除以3；空白余额按0，三家全部计入分母，不按非空个数计算。切换绑定后须核对表头和单位。",
    "D10": "FUN_2默认绑定WAVG：从各机构源表第10行的数值合计加权；空白利率按0%，权重保留，不使用Excel公式。",
    "G10": "FUN_3默认绑定AVG：取各机构源表C10数值合计的算术均值；源表无需增加均值字段，不使用Excel公式。",
}


class _FixtureCase(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory(prefix="summary-demo-")
        self.addCleanup(temporary.cleanup)
        self.directory = Path(temporary.name)
        self.config_path = self.directory / "bindings.xlsx"
        configuration = Workbook()
        configuration.active.title = "函数映射"
        configuration.active.append(["函数别名", "映射函数"])
        for alias, target in _EXPECTED_BINDINGS.items():
            configuration.active.append([alias, target])
        configuration.save(self.config_path)
        configuration.close()
        self.assets = create_demo_assets(self.directory / "示例数据_v0.6")

    def workbook(self, path: Path, *, data_only: bool = False):
        workbook = load_workbook(path, data_only=data_only)
        self.addCleanup(workbook.close)
        return workbook

    @property
    def asset_paths(self):
        return (self.assets.template_path, *self.assets.source_paths, *self.assets.error_paths,
                self.assets.root / "预期结果.json")

    @staticmethod
    def digest(path: Path):
        return hashlib.sha256(path.read_bytes()).hexdigest()

    @staticmethod
    def style_digest(sheet):
        # Comments legitimately gain provenance; all original cell formatting
        # must retain the same hash, independent of appended run-log styles.
        styles = [
            (cell.coordinate, cell.number_format,
             *(repr(copy(getattr(cell, attribute)))
               for attribute in ("font", "fill", "border", "alignment", "protection")))
            for row in sheet for cell in row
        ]
        return hashlib.sha256(json.dumps(styles, ensure_ascii=False).encode("utf-8")).hexdigest()

    def assert_decimal(self, actual, expected: Decimal, *, tolerance: Decimal = Decimal("0.000000000001")):
        self.assertIsInstance(actual, (int, float, Decimal))
        self.assertNotIsInstance(actual, bool)
        self.assertLessEqual(abs(Decimal(str(actual)) - expected), tolerance)

    def assert_no_excel_formulas(self, workbook):
        for sheet in workbook:
            for row in sheet:
                for cell in row:
                    with self.subTest(sheet=sheet.title, cell=cell.coordinate):
                        self.assertNotEqual(cell.data_type, "f")
                        self.assertFalse(isinstance(cell.value, str) and cell.value.startswith("="))

    def assert_expected_values(self, workbook):
        for (sheet, coordinate), expected in self.assets.expected_values.items():
            with self.subTest(sheet=sheet, cell=coordinate):
                # Million-yuan recurring averages cross Excel's binary-float
                # boundary; keep the original 24 checks at their prior precision.
                tolerance = Decimal("0.000000001") if coordinate.startswith("G") else Decimal("0.000000000001")
                self.assert_decimal(workbook[sheet][coordinate].value, expected, tolerance=tolerance)


class DemoGenerationTests(_FixtureCase):
    def test_fixed_paths_all_61_independent_values_and_original_numeric_results(self):
        assets = self.assets
        self.assertEqual(assets.root, (self.directory / "示例数据_v0.6").resolve())
        self.assertEqual(assets.template_path, assets.root / "汇总模板.xlsx")
        self.assertEqual(tuple(path.name for path in assets.source_paths), ("机构甲.xlsx", "机构乙.xlsx", "机构丙.xlsx"))
        self.assertTrue(all(path.parent == assets.root / "正常数据" for path in assets.source_paths))
        self.assertEqual(tuple(path.name for path in assets.error_paths), ("缺工作表.xlsx", "非法数值.xlsx", "公式无缓存.xlsx"))
        self.assertTrue(all(path.parent == assets.root / "异常数据" for path in assets.error_paths))
        self.assertTrue(all(path.is_file() for path in self.asset_paths))
        self.assertEqual({path for path in assets.root.rglob("*") if path.is_file()}, set(self.asset_paths))
        self.assertEqual(len(self.asset_paths), 8)  # Seven workbooks and JSON; no configuration copy or result.
        self.assertFalse((assets.root / "示例数据_v0.6").exists())
        self.assertEqual(len(assets.expected_values), 61)
        self.assertTrue(all(isinstance(value, Decimal) for value in assets.expected_values.values()))
        # These explicit expectations do not import engine rules, generated
        # source totals, or the fixture module's raw records/calculation helpers.
        with localcontext() as context:
            context.prec = 40
            missing_flow_rate = Decimal(1) / Decimal(30)
            total_balance_rate = Decimal(23) / Decimal(560)
            total_flow_rate = Decimal(11) / Decimal(300)
            averages = tuple(Decimal(value) / Decimal(3)
                             for value in ("10000000", "10000000", "4000000", "0", "4000000", "28000000"))
            diagnostic_averages = (Decimal(-5), Decimal(17) / Decimal(3), Decimal(5) / Decimal(3))
        detail_rows = (
            ("10000000", "0.048", "10000000", "0.036"),
            ("10000000", "0.053", "10000000", "0.046"),
            ("4000000", "0.03", "6000000", missing_flow_rate),
            ("0", "0", "0", "0"),
            ("4000000", "0.005", "4000000", "0.02"),
            ("28000000", total_balance_rate, "30000000", total_flow_rate),
        )
        original_28 = {
            ("业务明细", f"{column}{row}"): Decimal(value)
            for row, expected in enumerate(detail_rows, start=5)
            for column, value in zip("CDEF", expected)
        }
        original_28.update({("业务明细", f"G{row}"): value for row, value in enumerate(averages, start=5)})
        for coordinate in ("D8", "F8"):
            del original_28["业务明细", coordinate]
        self.assertEqual(len(original_28), 28)
        self.assertEqual({key: assets.expected_values[key] for key in original_28}, original_28)
        extrema_counts = (
            (9000000, 1000000, 2), (5000000, 2000000, 3), (3000000, 1000000, 2),
            (0, 0, 0), (2000000, 1000000, 3), (16000000, 5000000, 3),
        )
        added_33 = {("业务明细", f"{column}{row}"): Decimal(value)
                    for row, expected in enumerate(extrema_counts, start=5)
                    for column, value in zip("HIJ", expected)}
        diagnostic_rows = ((-15, -5, -10, 2), (17, 7, 3, 3), (5, 5, 0, 2))
        for row, expected, average in zip(range(13, 16), diagnostic_rows, diagnostic_averages):
            added_33.update({("业务明细", f"{column}{row}"): Decimal(value)
                             for column, value in zip("CHIJ", expected)})
            added_33["业务明细", f"G{row}"] = average
        self.assertEqual(len(added_33), 33)
        independent = {**original_28, **added_33}
        self.assertEqual(assets.expected_values, independent)
        self.assertFalse(hasattr(assets, "native_formula_values"))
        manifest = json.loads((assets.root / "预期结果.json").read_text(encoding="utf-8"))
        self.assertNotIn("native_formula_values", manifest)
        self.assertEqual(manifest["version"], "0.6")
        self.assertEqual(manifest["result_count"], 61)
        self.assertIn("61格", manifest["说明"])
        self.assertIn("原28格数值保持不变", manifest["说明"])
        self.assertIn("SUM15/AVG9/WAVG10/MAX9/MIN9/COUNT9", manifest["说明"])
        self.assertIn("三家全部计入分母", manifest["算术均值约定"])
        self.assertIn("主程序同目录config.xlsx", manifest["别名约定"])
        self.assertIn("表头、单位", manifest["别名约定"])
        self.assertIn("原始记录", manifest["预期值来源"])
        self.assertIn("不读取汇总引擎", manifest["预期值来源"])
        self.assertIn("每机构原始记录先独立求和", manifest["预期值来源"])
        self.assertIn("真实0有效", manifest["极值约定"])
        self.assertIn("非空有效数值", manifest["计数约定"])
        self.assertIn("不参加第10行原项目合计", manifest["诊断约定"])
        self.assertEqual(manifest["function_bindings"], _EXPECTED_BINDINGS)
        self.assertEqual(manifest["function_counts"], _EXPECTED_COUNTS)
        self.assertEqual(sum(manifest["function_counts"].values()), 61)
        self.assertEqual(manifest["diagnostics"], {
            "13": {"label": "负数与空白", "balances": ["-5", None, "-10"]},
            "14": {"label": "并列最大", "balances": ["7", "7", "3"]},
            "15": {"label": "真实0与空白", "balances": ["0", None, "5"]},
        })
        self.assertEqual(set(manifest["expected_values"]), {"业务明细"})
        actual = {(sheet, cell): Decimal(value)
                  for sheet, cells in manifest["expected_values"].items() for cell, value in cells.items()}
        self.assertEqual(actual, independent)

    def test_template_has_one_sheet_61_no_equals_rules_and_original_style(self):
        workbook = self.workbook(self.assets.template_path)
        self.assertEqual(workbook.properties.creator, "Z.ai")
        self.assertIn("v0.6", workbook.properties.title)
        self.assertEqual(workbook.sheetnames, ["业务明细"])
        self.assert_no_excel_formulas(workbook)
        detail = workbook["业务明细"]
        self.assertEqual(detail.max_column, 10)
        self.assertTrue(all(cell.value is None for cell in detail[1]))
        self.assertTrue(all(cell.value is None for cell in detail["A"]))
        self.assertIsNotNone(detail["B2"].value)
        self.assertFalse(detail["A1"].has_style)
        self.assertFalse(detail.sheet_view.showGridLines)
        self.assertEqual(detail.freeze_panes, "C5")
        self.assertEqual(detail.print_title_rows, "$4:$4")
        self.assertEqual(detail.page_setup.fitToWidth, 1)
        self.assertIn("$B$2:$J$20", str(detail.print_area))
        self.assertTrue({"B2:J2", "B18:J18", "B20:J20"}.issubset(
            {str(region) for region in detail.merged_cells.ranges}))
        self.assertFalse(any(region.min_row <= 15 and region.max_row >= 13
                             for region in detail.merged_cells.ranges))
        self.assertEqual(detail.auto_filter.ref, "B4:J9")
        self.assertEqual(detail["B2"].font.name, FONT_NAME)
        headers = tuple(detail.cell(4, column).value for column in range(2, 11))
        self.assertEqual(headers, ("项目", "余额指标", "余额利率指标", "发生额指标", "发生额利率指标", "余额参考指标",
                                   "余额极值A", "余额极值B", "有效填报数"))
        for header in headers:
            self.assertFalse(any(algorithm in header for algorithm in _EXPECTED_BINDINGS.values()))
        for coordinate, text in _ORIGINAL_COMMENTS.items():
            self.assertEqual(detail[coordinate].comment.text, text)
        for row in range(5, 11):
            expected = (f"FUN_1(C{row})", f"FUN_2(D{row},C{row})", f"FUN_1(E{row})", f"FUN_2(F{row},E{row})",
                        f"FUN_3(C{row})", f"FUN_4(C{row})", f"FUN_5(C{row})", f"FUN_6(C{row})")
            if row == 8:
                expected = (expected[0], "无有效权重", expected[2], "无有效权重", *expected[4:])
            self.assertEqual(tuple(detail.cell(row, column).value for column in range(3, 11)), expected)
            self.assertEqual(detail[f"G{row}"].data_type, "s")
            for column in (3, 5, 7, 8, 9):
                self.assertEqual(detail.cell(row, column).number_format, FORMATS["decimal_2"])
            for column in (4, 6):
                self.assertEqual(detail.cell(row, column).number_format, "0.0000%")
            self.assertEqual(detail[f"J{row}"].number_format, "0")
        for row, label in zip(range(13, 16), ("负数与空白", "并列最大", "真实0与空白")):
            self.assertEqual(detail[f"B{row}"].value, label)
            self.assertEqual(tuple(detail[f"{column}{row}"].value for column in "CGHIJ"),
                             tuple(f"{alias}(C{row})" for alias in ("FUN_1", "FUN_3", "FUN_4", "FUN_5", "FUN_6")))
            self.assertTrue(all(detail[f"{column}{row}"].value is None for column in "DEF"))
            for column in "CGHI":
                self.assertEqual(detail[f"{column}{row}"].number_format, FORMATS["decimal_2"])
            self.assertEqual(detail[f"J{row}"].number_format, "0")
        for row in (11, 12, 16, 17, 19):
            self.assertTrue(all(cell.value is None for cell in detail[row]))
        self.assertIn("不参加第10行原项目合计", detail["B20"].value)
        prefixes = tuple(alias + "(" for alias in _EXPECTED_BINDINGS)
        rules = [cell for row in detail for cell in row
                 if isinstance(cell.value, str) and cell.value.startswith(prefixes)]
        self.assertEqual(len(rules), 61)
        self.assertEqual({target: sum(cell.value.startswith(alias + "(") for cell in rules)
                          for alias, target in _EXPECTED_BINDINGS.items()}, _EXPECTED_COUNTS)
        self.assertEqual(detail["B10"].value, "合计")
        for column in "CGHIJ":
            self.assertEqual(detail[f"{column}5"].alignment.horizontal, "right")
        self.assertTrue(detail["B5"].alignment.wrap_text)
        self.assertNotEqual(detail["B5"].fill.fgColor.rgb, detail["B6"].fill.fgColor.rgb)
        self.assertEqual(detail["B4"].border.bottom.style, "thin")
        self.assertEqual(detail["J4"].border.bottom.style, "thin")
        self.assertEqual(detail["B10"].border.top.style, "medium")
        self.assertEqual(detail["J10"].border.top.style, "medium")
        self.assertIsNone(getattr(detail["C5"].border.left, "style", None))

    def test_sources_demonstrate_blanks_missing_rates_and_true_zero_without_formulas(self):
        books = [self.workbook(path) for path in self.assets.source_paths]
        for book in books:
            self.assertEqual(book.properties.creator, "Z.ai")
            self.assertEqual(book.sheetnames, ["业务明细", "利率台账"])
            self.assert_no_excel_formulas(book)
            for sheet in book:
                self.assertTrue(all(cell.value is None for cell in sheet[1]))
                self.assertTrue(all(cell.value is None for cell in sheet["A"]))
            self.assertTrue(all(cell.value is None for cell in book["业务明细"][8][2:6]))
            self.assertEqual(book["业务明细"].max_column, 6)
            self.assertEqual(book["利率台账"].max_column, 4)
            self.assertEqual(tuple(book["业务明细"].cell(4, column).value for column in range(2, 7)),
                             ("项目", "存量余额", "存量利率", "发生额", "发生额利率"))
            self.assertEqual(tuple(book["利率台账"].cell(4, column).value for column in range(2, 5)),
                             ("项目", "存量利率", "发生额利率"))
            self.assertEqual(book["业务明细"].auto_filter.ref, "B4:F9")
            self.assertEqual(book["利率台账"].auto_filter.ref, "B4:D9")
        first, second, third = books
        self.assertEqual(first["业务明细"]["C5"].value, 1000000)
        self.assertEqual(second["业务明细"]["C5"].value, 9000000)
        self.assertIsNone(third["业务明细"]["C5"].value)
        self.assertEqual(first["业务明细"]["C7"].value, 1000000)
        self.assertIsNone(first["业务明细"]["D7"].value)
        self.assertEqual(first["业务明细"]["D9"].value, 0)
        self.assertEqual(first["利率台账"]["C9"].value, 0)
        self.assertEqual(third["业务明细"]["D9"].value, 0)

    def test_original_source_rows_are_unchanged_and_diagnostics_only_supply_balance(self):
        original_rows = (
            ((1000000, 0.03, 2000000, 0.04), (2000000, 0.04, 1000000, 0.04),
             (1000000, None, 2000000, None), (None, None, None, None), (1000000, 0, 2000000, 0)),
            ((9000000, 0.05, 3000000, 0.06), (3000000, 0.05, 3000000, 0.04),
             (3000000, 0.04, 3000000, 0.06), (None, None, None, None), (1000000, 0.02, 2000000, 0.04)),
            ((None, None, 5000000, 0.02), (5000000, 0.06, 6000000, 0.05),
             (None, None, 1000000, 0.02), (None, None, None, None), (2000000, 0, None, None)),
        )
        diagnostics = ((-5, 7, 0), (None, 7, None), (-10, 3, 5))
        labels = ("负数与空白", "并列最大", "真实0与空白")
        for path, expected_rows, balances in zip(self.assets.source_paths, original_rows, diagnostics):
            book = self.workbook(path)
            details, ledger = book["业务明细"], book["利率台账"]
            with self.subTest(source=path.name):
                self.assertEqual(tuple(tuple(details.cell(row, col).value for col in range(3, 7))
                                       for row in range(5, 10)), expected_rows)
                self.assertEqual(tuple(details[f"C{row}"].value for row in range(13, 16)), balances)
                self.assertEqual(tuple(details[f"B{row}"].value for row in range(13, 16)), labels)
                self.assertTrue(all(details.cell(row, col).value is None
                                    for row in range(13, 16) for col in range(4, 7)))
                self.assertIn("不参加第10行原项目合计", details["B20"].value)
                self.assertIn("$B$2:$F$20", str(details.print_area))
                self.assertEqual(ledger.max_row, 12)
                self.assertIn("$B$2:$D$12", str(ledger.print_area))
                self.assertEqual(details.auto_filter.ref, "B4:F9")
                self.assertEqual(ledger.auto_filter.ref, "B4:D9")

    def test_each_source_has_independently_calculated_numeric_totals(self):
        with localcontext() as context:
            context.prec = 40
            expected_totals = (
                (Decimal("5000000"), Decimal("0.022"), Decimal("7000000"), Decimal(3) / Decimal(175)),
                (Decimal("16000000"), Decimal("0.04625"), Decimal("11000000"), Decimal(14) / Decimal(275)),
                (Decimal("7000000"), Decimal(3) / Decimal(70), Decimal("12000000"), Decimal("0.035")),
            )
        for path, expected in zip(self.assets.source_paths, expected_totals):
            details = self.workbook(path)["业务明细"]
            self.assertEqual(details["B10"].value, "合计")
            for column, value in zip("CDEF", expected):
                self.assert_decimal(details[f"{column}10"].value, value)
                self.assertEqual(details[f"{column}10"].border.top.style, "medium")
            # Recompute from this institution's raw rows, retaining the weight
            # when its paired rate is blank. Never use the aggregation engine.
            with localcontext() as context:
                context.prec = 40
                for amount_column, rate_column in (("C", "D"), ("E", "F")):
                    pairs = [(Decimal(str(details[f"{rate_column}{row}"].value or 0)),
                              Decimal(str(details[f"{amount_column}{row}"].value or 0)))
                             for row in range(5, 10)]
                    amount = sum((weight for _, weight in pairs), Decimal(0))
                    rate = sum((rate * weight for rate, weight in pairs), Decimal(0)) / amount
                    self.assert_decimal(details[f"{amount_column}10"].value, amount)
                    self.assert_decimal(details[f"{rate_column}10"].value, rate)
                    self.assertEqual(details[f"{rate_column}10"].number_format, "0.0000%")
            cached = self.workbook(path, data_only=True)["业务明细"]
            self.assertEqual(tuple(cached.cell(10, column).value for column in range(3, 7)),
                             tuple(details.cell(10, column).value for column in range(3, 7)))

    def test_error_workbooks_are_distinct_and_missing_cache_stays_missing(self):
        missing, invalid, uncached = (self.workbook(path) for path in self.assets.error_paths)
        self.assertNotIn("业务明细", missing.sheetnames)
        self.assert_no_excel_formulas(missing)
        self.assertEqual(invalid["业务明细"]["C5"].value, "这不是合法金额")
        self.assert_no_excel_formulas(invalid)
        self.assertEqual(uncached["业务明细"]["C5"].value, "=SUM(500000,500000)")
        self.assertEqual(uncached["业务明细"]["C5"].data_type, "f")
        formulas = [(sheet.title, cell.coordinate) for sheet in uncached for row in sheet for cell in row
                    if cell.data_type == "f"]
        self.assertEqual(formulas, [("业务明细", "C5")])
        cached = self.workbook(self.assets.error_paths[2], data_only=True)
        self.assertIsNone(cached["业务明细"]["C5"].value)

    def test_existing_files_are_not_overwritten_and_missing_files_are_filled(self):
        for offset, path in enumerate(self.asset_paths):
            path.write_bytes(f"原文件内容，禁止默认覆盖：{offset}".encode("utf-8"))
        before = {path: (path.read_bytes(), path.stat().st_mtime_ns) for path in self.asset_paths}
        create_demo_assets(self.assets.root)
        for path, snapshot in before.items():
            self.assertEqual((path.read_bytes(), path.stat().st_mtime_ns), snapshot)
        missing = self.assets.source_paths[0]
        missing.unlink()
        create_demo_assets(self.assets.root)
        self.assertEqual(self.workbook(missing)["业务明细"]["C5"].value, 1000000)
        for path, snapshot in before.items():
            if path != missing:
                self.assertEqual((path.read_bytes(), path.stat().st_mtime_ns), snapshot)

    def test_no_force_regeneration_preserves_user_edits_and_independent_config(self):
        path = self.assets.source_paths[0]
        marker = "已编辑的机构甲原始文件".encode("utf-8")
        path.write_bytes(marker)
        configuration = load_workbook(self.config_path)
        configuration.create_sheet("用户备注")["A1"] = "保留用户批注"
        configuration.save(self.config_path)
        configuration.close()
        before = {item: (self.digest(item), item.stat().st_mtime_ns) for item in (path, self.config_path)}
        create_demo_assets(self.assets.root, force=False)
        repeated = create_demo_assets(self.assets.root)
        self.assertEqual(path.read_bytes(), marker)
        self.assertEqual({item: (self.digest(item), item.stat().st_mtime_ns) for item in before}, before)
        self.assertEqual(repeated.expected_values, self.assets.expected_values)
        self.assertFalse((self.assets.root / "config.xlsx").exists())

    def test_command_line_entry_only_creates_inputs_and_preserves_existing_files(self):
        before = {path: (self.digest(path), path.stat().st_mtime_ns) for path in self.asset_paths}
        output = StringIO()
        with redirect_stdout(output):
            self.assertEqual(create_demo_main([str(self.assets.root)]), 0)
        self.assertIn("未运行应用", output.getvalue())
        self.assertIn("未生成汇总结果", output.getvalue())
        self.assertIn("61个数值结果", output.getvalue())
        self.assertIn("FUN_1=SUM、FUN_3=AVG、FUN_2=WAVG、FUN_4=MAX、FUN_5=MIN、FUN_6=COUNT", output.getvalue())
        self.assertIn("不参加第10行原项目合计", output.getvalue())
        self.assertIn("主程序同目录config.xlsx", output.getvalue())
        self.assertEqual({path: (self.digest(path), path.stat().st_mtime_ns) for path in self.asset_paths}, before)
        self.assertEqual(set(self.assets.root.rglob("*.xlsx")), set(self.asset_paths[:-1]))

    def test_default_command_line_directory_reuses_examples_and_protects_existing_files(self):
        entry_directory = self.directory / "独立命令行项目"
        legacy = entry_directory / "示例数据" / "保留旧示例.xlsx"
        version_02 = entry_directory / "示例数据_v0.2" / "保留v0.2示例.xlsx"
        version_03 = entry_directory / "示例数据_v0.3" / "保留v0.3示例.xlsx"
        version_04 = entry_directory / "示例数据_v0.4" / "保留v0.4示例.xlsx"
        published = entry_directory / "发布" / "v0.1" / "保留旧发布.xlsx"
        published_04 = entry_directory / "发布" / "v0.4" / "保留v0.4发布.xlsx"
        configuration = entry_directory / "config.xlsx"
        protected = (legacy, version_02, version_03, version_04, published, published_04, configuration)
        for marker in protected:
            marker.parent.mkdir(parents=True, exist_ok=True)
            marker.write_bytes("已有文件不得修改".encode("utf-8"))
        before_legacy = {path: (self.digest(path), path.stat().st_mtime_ns) for path in protected}
        root = entry_directory / "示例数据"
        with patch("create_demo.__file__", str(entry_directory / "create_demo.py")):
            with redirect_stdout(StringIO()):
                self.assertEqual(create_demo_main([]), 0)
            self.assertTrue((root / "汇总模板.xlsx").is_file())
            paths = tuple(path for path in root.rglob("*") if path.is_file())
            self.assertEqual(len(paths), 9)
            before = {path: (self.digest(path), path.stat().st_mtime_ns) for path in paths}
            with redirect_stdout(StringIO()):
                self.assertEqual(create_demo_main([]), 0)
            self.assertEqual({path: (self.digest(path), path.stat().st_mtime_ns) for path in paths}, before)
        self.assertEqual({path: (self.digest(path), path.stat().st_mtime_ns) for path in protected}, before_legacy)
        self.assertEqual(len(tuple(root.rglob("*.xlsx"))), 8)
        self.assertFalse((entry_directory / "示例数据_v0.6").exists())
        self.assertEqual(self.workbook(root / "汇总模板.xlsx").sheetnames, ["业务明细"])


class DemoAggregationTests(_FixtureCase):
    def setUp(self):
        super().setUp()
        # Generation-only discovery does not import or run the engine. Full
        # acceptance always exercises its public API without dependency skips.
        from summary_tool.engine import aggregate, AggregationError
        self.aggregate = aggregate
        self.aggregation_error = AggregationError

    def aggregate_demo(self, name="汇总结果.xlsx"):
        output = self.directory / name
        result = self.aggregate(self.assets.template_path, self.assets.source_paths, output,
                                config_path=self.config_path)
        self.assertEqual(Path(result.output_path).resolve(), output.resolve())
        self.assertEqual(Path(result.config_path).resolve(), self.config_path.resolve())
        self.assertTrue(output.is_file())
        workbook = self.workbook(output)
        self.assertEqual(workbook.sheetnames, ["业务明细", "运行日志"])
        self.assert_no_excel_formulas(workbook)
        return result, workbook

    def test_aggregation_matches_all_61_independent_values_and_six_function_statistics(self):
        result, workbook = self.aggregate_demo()
        self.assertEqual(result.source_count, 3)
        self.assertEqual(result.source_sheet_count, 3)  # Six physical source tabs, only three read.
        self.assertEqual(result.template_sheet_count, 1)  # The generated log is not a template tab.
        self.assertEqual(result.rule_count, 61)
        self.assertEqual(result.sum_rule_count, 15)
        self.assertEqual(result.avg_rule_count, 9)
        self.assertEqual(result.wavg_rule_count, 10)
        self.assertEqual(result.function_bindings, _EXPECTED_BINDINGS)
        self.assertEqual(result.function_counts, _EXPECTED_COUNTS)
        self.assertEqual(sum(result.function_counts.values()), result.rule_count)
        self.assert_expected_values(workbook)
        self.assertEqual(workbook["业务明细"]["C5"].value, 10000000)
        self.assert_decimal(workbook["业务明细"]["D5"].value, Decimal("0.048"))
        self.assertEqual(workbook["业务明细"]["C10"].value, 28000000)
        self.assertEqual(workbook["业务明细"]["E10"].value, 30000000)
        with localcontext() as context:
            context.prec = 40
            self.assert_decimal(workbook["业务明细"]["D10"].value, Decimal(23) / Decimal(560))
            self.assert_decimal(workbook["业务明细"]["F10"].value, Decimal(11) / Decimal(300))
            # The third institution has no C5 balance but remains in AVG's
            # denominator, so this must not be the nonblank mean of 5,000,000.
            self.assert_decimal(workbook["业务明细"]["G5"].value, Decimal(10000000) / Decimal(3),
                                tolerance=Decimal("0.000000001"))
            self.assert_decimal(workbook["业务明细"]["G10"].value, Decimal(28000000) / Decimal(3),
                                tolerance=Decimal("0.000000001"))
        self.assertEqual(workbook["业务明细"]["G8"].value, 0)
        # All 61 values must also be available without Excel ever opening them.
        self.assert_expected_values(self.workbook(result.output_path, data_only=True))
        for row in range(5, 11):
            for column in (3, 5, 7):
                self.assertEqual(workbook["业务明细"].cell(row, column).number_format, FORMATS["decimal_2"])
            for column in (4, 6):
                self.assertEqual(workbook["业务明细"].cell(row, column).number_format, "0.0000%")

    def test_explicit_six_bindings_do_not_consult_application_root(self):
        with patch("summary_tool.configuration.application_root",
                   side_effect=AssertionError("测试不得读取root配置")):
            result, workbook = self.aggregate_demo("独立六绑定.xlsx")
        self.assertEqual(result.function_bindings, _EXPECTED_BINDINGS)
        self.assertEqual(result.function_counts, _EXPECTED_COUNTS)
        self.assert_expected_values(workbook)
        self.assertFalse((self.assets.root / "config.xlsx").exists())

    def test_extreme_sources_and_ties_are_located_in_issues_and_appended_comments(self):
        result, workbook = self.aggregate_demo()
        expected_sources = {
            "H5": ("机构乙.xlsx",), "I5": ("机构甲.xlsx",),
            "H6": ("机构丙.xlsx",), "I6": ("机构甲.xlsx",),
            "H7": ("机构乙.xlsx",), "I7": ("机构甲.xlsx",),
            "H9": ("机构丙.xlsx",), "I9": ("机构甲.xlsx", "机构乙.xlsx"),
            "H10": ("机构乙.xlsx",), "I10": ("机构甲.xlsx",),
            "H13": ("机构甲.xlsx",), "I13": ("机构丙.xlsx",),
            "H14": ("机构甲.xlsx", "机构乙.xlsx"), "I14": ("机构丙.xlsx",),
            "H15": ("机构丙.xlsx",), "I15": ("机构甲.xlsx",),
        }
        expected = {
            ("业务明细", coordinate, filename, "业务明细", "C" + coordinate[1:])
            for coordinate, filenames in expected_sources.items() for filename in filenames
        }
        source_issues = [issue for issue in result.issues if issue.code == "result_source"]
        actual = {(issue.template_sheet, issue.template_cell, Path(issue.source_file).name,
                   issue.source_sheet, issue.source_cell) for issue in source_issues}
        self.assertEqual(actual, expected)
        self.assertEqual(len(source_issues), len(expected))
        self.assertTrue(all(issue.severity == "info" for issue in source_issues))
        template = self.workbook(self.assets.template_path)["业务明细"]
        details = workbook["业务明细"]
        for coordinate, filenames in expected_sources.items():
            with self.subTest(cell=coordinate):
                comment = details[coordinate].comment
                self.assertIsNotNone(comment)
                self.assertIn("MAX" if coordinate.startswith("H") else "MIN", comment.text)
                self.assertIn(f"'业务明细'!C{coordinate[1:]}", comment.text)
                for path in self.assets.source_paths:
                    if path.name in filenames:
                        self.assertIn(path.name, comment.text)
                        self.assertIn(str(path.resolve()), comment.text)
                    else:
                        self.assertNotIn(path.name, comment.text)
                if len(filenames) > 1:
                    self.assertIn("并列", comment.text)
                original = template[coordinate].comment
                if original is not None:
                    self.assertTrue(comment.text.startswith(original.text + "\n\n"))
                    self.assertGreater(len(comment.text), len(original.text))
                    self.assertEqual(comment.author, original.author)
        self.assertEqual(details["H5"].value, 9000000)
        self.assertEqual(details["I9"].value, 1000000)
        self.assertEqual(details["H14"].value, 7)
        # Sources are attached to original source C10 totals, not project C5:C9
        # records or any of the independent diagnostics.
        self.assertEqual(details["H10"].value, 16000000)
        self.assertEqual(details["I10"].value, 5000000)

    def test_all_blank_extremes_have_zero_no_source_and_explicit_appended_notes(self):
        result, workbook = self.aggregate_demo()
        note = "无有效数值，无极值来源，按约定输出0"
        notices = [issue for issue in result.issues if issue.code == "no_valid_values"]
        self.assertEqual({(issue.template_sheet, issue.template_cell) for issue in notices},
                         {("业务明细", "H8"), ("业务明细", "I8")})
        self.assertEqual(len(notices), 2)
        self.assertTrue(all(issue.severity == "info" and issue.message == note for issue in notices))
        self.assertTrue(all(not issue.source_file and not issue.source_sheet and not issue.source_cell
                            for issue in notices))
        template = self.workbook(self.assets.template_path)["业务明细"]
        for coordinate in ("H8", "I8"):
            cell = workbook["业务明细"][coordinate]
            self.assertEqual(cell.value, 0)
            self.assertIsNotNone(cell.comment)
            self.assertTrue(cell.comment.text.startswith(template[coordinate].comment.text + "\n\n"))
            self.assertIn(note, cell.comment.text)
            self.assertFalse(any(path.name in cell.comment.text for path in self.assets.source_paths))
            self.assertFalse(any(issue.code == "result_source" and issue.template_cell == coordinate
                                 for issue in result.issues))
        self.assertEqual(workbook["业务明细"]["J8"].value, 0)

    def test_independent_diagnostics_preserve_negative_zero_and_blank_semantics_and_counts(self):
        _, workbook = self.aggregate_demo()
        details = workbook["业务明细"]
        self.assertEqual(tuple(details[f"{column}13"].value for column in "CHIJ"), (-15, -5, -10, 2))
        self.assertEqual(tuple(details[f"{column}14"].value for column in "CHIJ"), (17, 7, 3, 3))
        self.assertEqual(tuple(details[f"{column}15"].value for column in "CHIJ"), (5, 5, 0, 2))
        self.assertEqual(details["G13"].value, -5)
        self.assertEqual(details["C10"].value, 28000000)
        self.assertEqual(details["E10"].value, 30000000)
        self.assertEqual(details.auto_filter.ref, "B4:J9")
        for row in (*range(5, 11), *range(13, 16)):
            with self.subTest(row=row):
                count = details[f"J{row}"]
                self.assertIsInstance(count.value, int)
                self.assertNotIsInstance(count.value, bool)
                self.assertEqual(count.number_format, "0")
                self.assertEqual(count.alignment.horizontal, "right")
                for column in "HI":
                    self.assertEqual(details[f"{column}{row}"].number_format, FORMATS["decimal_2"])
        for row in range(13, 16):
            self.assertTrue(all(details[f"{column}{row}"].value is None for column in "DEF"))

    def test_missing_rates_warn_real_zero_retains_weight_and_all_blank_is_zero(self):
        result, workbook = self.aggregate_demo()
        warnings = [issue for issue in result.issues if issue.severity == "warning"]
        self.assertTrue(warnings)
        self.assertEqual(result.warning_count, len(warnings))
        missing_rates = [issue for issue in warnings if issue.code == "blank_rate_with_weight"]
        self.assertEqual(len(missing_rates), 2)
        self.assertEqual({(issue.template_sheet, issue.template_cell, issue.source_sheet, issue.source_cell,
                           Path(issue.source_file).name) for issue in missing_rates},
                         {("业务明细", "D7", "业务明细", "D7", "机构甲.xlsx"),
                          ("业务明细", "F7", "业务明细", "F7", "机构甲.xlsx")})
        self.assertFalse(any(issue.template_cell in {"D9", "F9"} for issue in warnings))
        self.assertFalse(any(issue.source_sheet == "利率台账" for issue in result.issues))
        self.assert_decimal(workbook["业务明细"]["D7"].value, Decimal("0.03"))
        self.assert_decimal(workbook["业务明细"]["D9"].value, Decimal("0.005"))
        self.assert_decimal(workbook["业务明细"]["F9"].value, Decimal("0.02"))
        self.assertEqual(tuple(workbook["业务明细"].cell(8, col).value for col in range(3, 11)),
                         (0, "无有效权重", 0, "无有效权重", 0, 0, 0, 0))

    def test_all_input_hashes_static_content_and_major_styles_are_preserved(self):
        tracked_paths = (*self.asset_paths, self.config_path)
        before = {path: self.digest(path) for path in tracked_paths}
        result, output = self.aggregate_demo()
        self.assertEqual({path: self.digest(path) for path in tracked_paths}, before)
        template = self.workbook(self.assets.template_path)
        cells = ("B2", "B4", "G4", "H4", "I4", "J4", "C5", "D5", "G5", "H5", "J5", "B6", "D7", "G8",
                 "H8", "I8", "I9", "B10", "C10", "D10", "E10", "F10", "G10", "H10", "I10", "J10",
                 "B13", "H13", "I13", "H14", "I15", "J15", "B18", "B20")
        for coordinate in cells:
            original, actual = template["业务明细"][coordinate], output["业务明细"][coordinate]
            with self.subTest(cell=coordinate):
                for attribute in ("font", "fill", "border", "alignment", "protection"):
                    self.assertEqual(copy(getattr(original, attribute)), copy(getattr(actual, attribute)))
                self.assertEqual(original.number_format, actual.number_format)
                if original.comment is not None:
                    self.assertIsNotNone(actual.comment)
                    self.assertTrue(actual.comment.text.startswith(original.comment.text))
                    self.assertEqual(actual.comment.author, original.comment.author)
        original, actual = template["业务明细"], output["业务明细"]
        self.assertEqual(self.style_digest(actual), self.style_digest(original))
        for coordinate, text in _ORIGINAL_COMMENTS.items():
            self.assertTrue(actual[coordinate].comment.text.startswith(text))
        self.assertEqual(str(actual.merged_cells), str(original.merged_cells))
        self.assertEqual(actual.freeze_panes, original.freeze_panes)
        self.assertEqual(actual.auto_filter.ref, original.auto_filter.ref)
        self.assertEqual(actual.print_area, original.print_area)
        self.assertEqual(actual.print_title_rows, original.print_title_rows)
        self.assertEqual(actual.sheet_view.showGridLines, original.sheet_view.showGridLines)
        self.assertEqual(actual.page_setup.orientation, original.page_setup.orientation)
        for column, dimension in original.column_dimensions.items():
            self.assertEqual(actual.column_dimensions[column].width, dimension.width)
        for row, dimension in original.row_dimensions.items():
            self.assertEqual(actual.row_dimensions[row].height, dimension.height)
        for row in original:
            for cell in row:
                if (original.title, cell.coordinate) not in self.assets.expected_values:
                    self.assertEqual(actual[cell.coordinate].value, cell.value)
        self.assertEqual(Path(result.output_path).name, "汇总结果.xlsx")

    def test_every_error_source_fails_without_publishing_or_modifying_inputs(self):
        codes = ("missing_sheet", "invalid_number", "formula_cache_missing")
        tracked_paths = (*self.asset_paths, self.config_path)
        before = {path: self.digest(path) for path in tracked_paths}
        for source, code in zip(self.assets.error_paths, codes):
            with self.subTest(source=source.name):
                output = self.directory / f"不应生成-{source.stem}.xlsx"
                with self.assertRaises(self.aggregation_error) as caught:
                    self.aggregate(self.assets.template_path, (*self.assets.source_paths, source), output,
                                   config_path=self.config_path)
                self.assertFalse(output.exists())
                self.assertFalse(tuple(self.directory.glob(".汇总临时_*.xlsx")))
                self.assertEqual({path: self.digest(path) for path in tracked_paths}, before)
                matching = [issue for issue in caught.exception.issues if issue.code == code]
                self.assertTrue(matching, caught.exception.issues)
                self.assertTrue(any(issue.severity == "error" and Path(issue.source_file).name == source.name
                                    for issue in matching))
        self.assertIsNone(self.workbook(self.assets.error_paths[2], data_only=True)["业务明细"]["C5"].value)

    def test_existing_output_is_never_replaced(self):
        output = self.directory / "已有结果必须保留.xlsx"
        output.write_bytes("用户已有结果，不得覆盖".encode("utf-8"))
        tracked_paths = (*self.asset_paths, self.config_path, output)
        before = {path: (self.digest(path), path.stat().st_mtime_ns) for path in tracked_paths}
        with self.assertRaises(self.aggregation_error):
            self.aggregate(self.assets.template_path, self.assets.source_paths, output,
                           config_path=self.config_path)
        self.assertEqual({path: (self.digest(path), path.stat().st_mtime_ns) for path in tracked_paths}, before)
        self.assertFalse(tuple(self.directory.glob(".汇总临时_*.xlsx")))

    def test_unreferenced_ledger_is_optional_and_does_not_count_as_participating(self):
        source = self.directory / "正常但无台账.xlsx"
        workbook = self.workbook(self.assets.source_paths[0])
        workbook.remove(workbook["利率台账"])
        workbook.save(source)
        output = self.directory / "未引用台账不影响汇总.xlsx"
        result = self.aggregate(self.assets.template_path, (source, *self.assets.source_paths[1:]), output,
                                config_path=self.config_path)
        self.assertEqual(result.source_count, 3)
        self.assertEqual(result.source_sheet_count, 3)
        self.assertEqual(result.template_sheet_count, 1)
        actual = self.workbook(output)
        self.assertEqual(actual.sheetnames, ["业务明细", "运行日志"])
        self.assert_no_excel_formulas(actual)
        self.assert_expected_values(actual)
        self.assertFalse(any(issue.code == "missing_sheet" for issue in result.issues))

    def test_excel_formula_in_template_is_rejected_without_publishing(self):
        template = self.directory / "误写等号的模板.xlsx"
        workbook = self.workbook(self.assets.template_path)
        workbook["业务明细"]["C10"] = "=SUM(C5:C9)"
        workbook.save(template)
        before = self.digest(template)
        output = self.directory / "普通公式不应发布.xlsx"
        with self.assertRaises(self.aggregation_error) as caught:
            self.aggregate(template, self.assets.source_paths, output, config_path=self.config_path)
        self.assertFalse(output.exists())
        self.assertEqual(self.digest(template), before)
        self.assertTrue(any(issue.code == "template_formula_not_supported"
                            and issue.template_sheet == "业务明细" and issue.template_cell == "C10"
                            for issue in caught.exception.issues))


if __name__ == "__main__":
    unittest.main()
