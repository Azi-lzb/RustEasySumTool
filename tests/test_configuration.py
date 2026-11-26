from contextlib import chdir
from dataclasses import FrozenInstanceError
from pathlib import Path
from tempfile import TemporaryDirectory
from types import MappingProxyType
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from summary_tool.configuration import (
    ConfigurationError, ConfigurationSnapshot, DEFAULT_BINDINGS,
    MAPPING_HEADERS, MAPPING_SHEET, load_configuration,
)
from summary_tool.functions import function_names


class ConfigurationTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory(prefix="summary_configuration_test_")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.path = self.root / "算法配置.xlsx"

    def write_configuration(self, rows=(), *, path=None, header=True, library=None, sheet_name=MAPPING_SHEET):
        workbook = Workbook()
        try:
            sheet = workbook.active
            sheet.title = sheet_name
            if header:
                sheet.append(MAPPING_HEADERS)
            for row in rows:
                sheet.append(row)
            if library is not None:
                reference = workbook.create_sheet("函数库")
                for row in library:
                    reference.append(row)
            workbook.save(path or self.path)
        finally:
            workbook.close()
        return path or self.path

    def assert_bad_configuration(self, rows, *, detail=None):
        self.write_configuration(rows)
        with self.assertRaises(ConfigurationError) as caught:
            load_configuration(self.path)
        message = str(caught.exception)
        self.assertIn(str(self.path.resolve()), message)
        self.assertIn(MAPPING_SHEET, message)
        self.assertIn("A2/B2", message)
        if detail is not None:
            self.assertIn(detail, message)
        return message

    def test_configuration_error_is_a_value_error(self):
        self.assertTrue(issubclass(ConfigurationError, ValueError))

    def test_loads_example_bindings_without_modifying_file(self):
        self.assertEqual(DEFAULT_BINDINGS, {
            "FUN_1": "SUM", "FUN_2": "WAVG", "FUN_3": "AVG",
            "FUN_4": "MAX", "FUN_5": "MIN", "FUN_6": "COUNT",
            "FUN_7": "SUMPRODUCT", "FUN_8": "RATIO",
        })
        self.write_configuration(DEFAULT_BINDINGS.items())
        before = self.path.read_bytes()
        snapshot = load_configuration(self.path)
        self.assertEqual(snapshot.bindings, DEFAULT_BINDINGS)
        self.assertEqual(snapshot.path, self.path.resolve())
        self.assertEqual(self.path.read_bytes(), before)

    def test_default_path_uses_application_root(self):
        path = self.root / "config.xlsx"
        self.write_configuration([("FUN_1", "AVG")], path=path)
        with patch("summary_tool.configuration.application_root", return_value=self.root) as root:
            snapshot = load_configuration()
        root.assert_called_once_with()
        self.assertEqual(snapshot.bindings, {"FUN_1": "AVG"})
        self.assertEqual(snapshot.path, path.resolve())

    def test_missing_default_returns_empty_snapshot_without_creating_file(self):
        with patch("summary_tool.configuration.application_root", return_value=self.root):
            snapshot = load_configuration()
        self.assertEqual(snapshot.bindings, {})
        self.assertIsNone(snapshot.path)
        self.assertEqual(list(self.root.iterdir()), [])

    def test_missing_explicit_path_is_an_error_without_fallback(self):
        missing = self.root / "未创建目录" / "config.xlsx"
        with patch("summary_tool.configuration.application_root", side_effect=AssertionError("不能回退")):
            for argument in (missing, str(missing)):
                with self.subTest(argument=argument), self.assertRaises(ConfigurationError) as caught:
                    load_configuration(argument)
                self.assertIn(str(missing.resolve()), str(caught.exception))
        self.assertFalse(missing.parent.exists())

    def test_toml_is_rejected_and_never_used_as_default_fallback(self):
        old = self.root / "config.toml"
        old.write_text('[functions]\nFUN_1="SUM"\n', encoding="utf-8")
        with self.assertRaisesRegex(ConfigurationError, "扩展名必须为 .xlsx"):
            load_configuration(old)
        with patch("summary_tool.configuration.application_root", return_value=self.root):
            snapshot = load_configuration()
        self.assertEqual(snapshot.bindings, {})
        self.assertIsNone(snapshot.path)

    def test_empty_mapping_does_not_use_default_bindings(self):
        for header in (True, False):
            with self.subTest(header=header):
                self.write_configuration(header=header)
                snapshot = load_configuration(self.path)
                self.assertEqual(snapshot.bindings, {})
                self.assertEqual(snapshot.path, self.path.resolve())

    def test_aliases_and_algorithm_names_are_trimmed_and_case_insensitive(self):
        self.write_configuration([(" fun_1 ", " sUm "), ("F1_2", "aVg"), ("f2_1", "wavg")])
        self.assertEqual(load_configuration(self.path).bindings,
                         {"FUN_1": "SUM", "F1_2": "AVG", "F2_1": "WAVG"})

    def test_blank_rows_extra_columns_and_reference_sheet_are_ignored(self):
        self.write_configuration(
            [(None, None), ("  ", "\t"), ("FUN_1", "SUM", "说明文字"), (None, None), ("FUN_2", "WAVG")],
            library=[("SUM", "这段查询说明可修改"), ("不存在的算法", "=1/0")],
        )
        self.assertEqual(load_configuration(self.path).bindings, {"FUN_1": "SUM", "FUN_2": "WAVG"})

    def test_mapping_sheet_is_selected_by_name_even_when_reference_sheet_is_first(self):
        self.write_configuration([("FUN_1", "SUM")], library=[("SUM", "查询说明")])
        workbook = load_workbook(self.path)
        workbook.move_sheet("函数库", offset=-1)
        workbook.save(self.path)
        workbook.close()
        self.assertEqual(load_configuration(self.path).bindings, {"FUN_1": "SUM"})

    def test_single_sheet_reference_columns_d_e_do_not_change_mapping(self):
        self.write_configuration([
            ("FUN_1", "SUM", None, "SUM", "至少1个参数"),
            (None, None, None, "RATIO", "=1/0"),
            ("FUN_8", "RATIO", None, "未知函数", "仅供查询"),
        ])
        self.assertEqual(load_configuration(self.path).bindings, {"FUN_1": "SUM", "FUN_8": "RATIO"})

    def test_header_is_optional(self):
        self.write_configuration([("FUN_1", "SUM")], header=False)
        self.assertEqual(load_configuration(self.path).bindings, {"FUN_1": "SUM"})

    def test_missing_mapping_sheet_is_an_error(self):
        self.write_configuration([("FUN_1", "SUM")], sheet_name="函数库")
        with self.assertRaisesRegex(ConfigurationError, "缺少“函数映射”工作表"):
            load_configuration(self.path)

    def test_multiple_aliases_can_use_same_algorithm_without_arity_constraints(self):
        rows = [("FUN_1", "WAVG"), ("FUN_99", "SUM"), ("F1_1", "WAVG"), ("F2_1", "SUM")]
        self.write_configuration(rows)
        self.assertEqual(load_configuration(self.path).bindings, dict(rows))

    def test_chinese_paths_relative_paths_and_uppercase_extension(self):
        directory = self.root / "中文 配置目录"
        directory.mkdir()
        path = directory / "汇总算法.XLSX"
        self.write_configuration([("fun_9", "SUM")], path=path)
        with chdir(directory):
            snapshot = load_configuration(Path(path.name))
        self.assertEqual(snapshot.bindings, {"FUN_9": "SUM"})
        self.assertEqual(snapshot.path, path.resolve())

    def test_snapshot_copies_callers_mapping_and_mapping_proxy(self):
        for proxy in (False, True):
            source = {"FUN_1": "SUM"}
            snapshot = ConfigurationSnapshot(MappingProxyType(source) if proxy else source)
            source["FUN_1"] = "AVG"
            self.assertEqual(snapshot.bindings, {"FUN_1": "SUM"})
            self.assertIsInstance(snapshot.bindings, MappingProxyType)

    def test_snapshot_fields_and_bindings_are_immutable(self):
        snapshot = ConfigurationSnapshot({"FUN_1": "SUM"}, self.path)
        with self.assertRaises(TypeError):
            snapshot.bindings["FUN_1"] = "AVG"
        with self.assertRaises(TypeError):
            del snapshot.bindings["FUN_1"]
        with self.assertRaises(FrozenInstanceError):
            snapshot.bindings = {}
        with self.assertRaises(FrozenInstanceError):
            snapshot.path = None

    def test_next_call_reloads_without_changing_existing_snapshot(self):
        self.write_configuration([("FUN_1", "SUM")])
        first = load_configuration(self.path)
        self.write_configuration([("FUN_1", "AVG"), ("FUN_2", "WAVG")])
        second = load_configuration(self.path)
        self.assertEqual(first.bindings, {"FUN_1": "SUM"})
        self.assertEqual(second.bindings, {"FUN_1": "AVG", "FUN_2": "WAVG"})
        self.path.unlink()
        self.assertEqual(first.bindings, {"FUN_1": "SUM"})
        with self.assertRaises(ConfigurationError):
            load_configuration(self.path)

    def test_reads_one_byte_snapshot_and_closes_workbook(self):
        self.write_configuration([("FUN_1", "SUM")])
        content = self.path.read_bytes()
        with (
            patch.object(Path, "read_bytes", autospec=True, return_value=content) as reader,
            patch("summary_tool.configuration.load_workbook", wraps=load_workbook) as loader,
        ):
            snapshot = load_configuration(self.path)
        reader.assert_called_once_with(self.path.resolve())
        loader.assert_called_once()
        self.assertEqual(snapshot.bindings, {"FUN_1": "SUM"})
        self.path.rename(self.root / "读取完成.xlsx")

    def test_missing_alias_or_target_and_nonstrings_are_errors(self):
        for row in ((None, "SUM"), ("", "SUM"), (123, "SUM"), (True, "SUM"),
                    ("FUN_1", None), ("FUN_1", ""), ("FUN_1", "   "),
                    ("FUN_1", 123), ("FUN_1", True)):
            with self.subTest(row=row):
                self.assert_bad_configuration([row])

    def test_mapping_cells_reject_formulas_and_excel_errors(self):
        for row in (("FUN_1", '=IF(TRUE,"SUM","AVG")'), ("=1+1", "SUM"),
                    ("FUN_1", "#DIV/0!"), ("#N/A", "SUM")):
            with self.subTest(row=row):
                self.assert_bad_configuration([row], detail="Excel 公式或错误值")

    def test_alias_numbers_must_be_positive_without_leading_zeroes(self):
        for alias in ("FUN_0", "FUN_01", "FUN_-1", "FUN_１", "FUN_1extra", "FUN_1_1",
                      "F0_1", "F1_0", "F01_1", "F1_01", "F1", "X1_1", "F１_1"):
            with self.subTest(alias=alias):
                self.assert_bad_configuration([(alias, "SUM")], detail=alias)

    def test_builtin_names_cannot_be_overridden(self):
        for alias in (*function_names(), "sum", "aVg", "wavg"):
            with self.subTest(alias=alias):
                self.assert_bad_configuration([(alias, "SUM")], detail="内置")

    def test_duplicate_alias_rows_are_case_insensitive(self):
        for aliases in (("FUN_1", "FUN_1"), ("fun_1", "FUN_1"), ("F1_1", "f1_1")):
            with self.subTest(aliases=aliases):
                self.write_configuration([(aliases[0], "SUM"), (aliases[1], "AVG")])
                with self.assertRaises(ConfigurationError) as caught:
                    load_configuration(self.path)
                self.assertIn("A3/B3", str(caught.exception))
                self.assertIn("重复", str(caught.exception))

    def test_unknown_algorithms_and_alias_chains_are_rejected(self):
        for target in ("MEDIAN", "NotRegistered", "FUN_2", "F1_99"):
            with self.subTest(target=target):
                self.assert_bad_configuration([("FUN_1", target)], detail=target)

    def test_corrupt_excel_configuration_is_reported_with_path(self):
        for content in (b"", b"not an xlsx file", b"PK\x03\x04truncated"):
            with self.subTest(content=content):
                self.path.write_bytes(content)
                with self.assertRaises(ConfigurationError) as caught:
                    load_configuration(self.path)
                self.assertIn(str(self.path.resolve()), str(caught.exception))
                self.assertIn("无法读取 Excel 配置", str(caught.exception))

    def test_read_errors_are_not_treated_as_missing_defaults(self):
        for error in (PermissionError("拒绝访问"), IsADirectoryError("不是文件"), OSError("读取失败")):
            for explicit in (False, True):
                expected_path = self.path if explicit else self.root / "config.xlsx"
                with (
                    self.subTest(error=error, explicit=explicit),
                    patch("summary_tool.configuration.application_root", return_value=self.root),
                    patch.object(Path, "read_bytes", autospec=True, side_effect=error) as reader,
                ):
                    with self.assertRaises(ConfigurationError) as caught:
                        load_configuration(self.path if explicit else None)
                reader.assert_called_once_with(expected_path.resolve())
                self.assertIn(str(expected_path.resolve()), str(caught.exception))
                self.assertIn(str(error), str(caught.exception))

    def test_algorithm_targets_are_not_executed_or_imported(self):
        marker = self.root / "不能创建.txt"
        payload = f'__import__("pathlib").Path({str(marker)!r}).write_text("executed")'
        for target in (payload, "os", "pathlib.Path", "os:system"):
            self.write_configuration([("FUN_1", target)])
            with (
                self.subTest(target=target),
                patch("builtins.eval", side_effect=AssertionError("不能执行配置")),
                patch("importlib.import_module", side_effect=AssertionError("不能导入配置目标")),
                self.assertRaises(ConfigurationError),
            ):
                load_configuration(self.path)
            self.assertFalse(marker.exists())


if __name__ == "__main__":
    unittest.main()
