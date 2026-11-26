"""用假的 Tk 根和 filedialog 验证选择及销毁，不打开真实窗口。"""

from pathlib import Path
import sys
import tempfile
from types import ModuleType
import unittest
from unittest import mock

from summary_tool import dialogs


class FileDialogsTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="summary_dialogs_")
        self.addCleanup(temporary.cleanup)
        self.app_root = Path(temporary.name).resolve()
        self.root = mock.Mock(name="hidden_root")
        self.tkinter = ModuleType("tkinter")
        self.tkinter.Tk = mock.Mock(return_value=self.root)
        self.filedialog = ModuleType("tkinter.filedialog")
        self.filedialog.askopenfilename = mock.Mock(return_value="")
        self.filedialog.askopenfilenames = mock.Mock(return_value=())
        self.filedialog.askdirectory = mock.Mock(return_value="")
        self.tkinter.filedialog = self.filedialog
        modules = mock.patch.dict(sys.modules, {
            "tkinter": self.tkinter,
            "tkinter.filedialog": self.filedialog,
        })
        modules.start()
        self.addCleanup(modules.stop)
        root_patch = mock.patch.object(dialogs, "application_root", return_value=self.app_root)
        self.application_root = root_patch.start()
        self.addCleanup(root_patch.stop)

    def _assert_no_selection(self):
        self.filedialog.askopenfilename.assert_not_called()
        self.filedialog.askopenfilenames.assert_not_called()
        self.filedialog.askdirectory.assert_not_called()

    def test_constructing_dialogs_does_not_import_or_initialize_tk(self):
        original_import = __import__

        def guarded_import(name, *args, **kwargs):
            if name == "tkinter" or name.startswith("tkinter."):
                raise AssertionError("构造时不应导入 Tk")
            return original_import(name, *args, **kwargs)

        with mock.patch("builtins.__import__", side_effect=guarded_import):
            chooser = dialogs.FileDialogs()
        self.assertEqual(chooser.app_root, self.app_root)
        self.application_root.assert_called_once_with()
        self.tkinter.Tk.assert_not_called()
        self._assert_no_selection()

    def test_context_hides_updates_and_destroys_root(self):
        chooser = dialogs.FileDialogs(self.app_root)
        with chooser as entered:
            self.assertIs(entered, chooser)
            self.assertEqual(self.root.method_calls, [mock.call.withdraw(), mock.call.update()])
            self.root.destroy.assert_not_called()
        self.tkinter.Tk.assert_called_once_with()
        self.assertEqual(self.root.method_calls, [
            mock.call.withdraw(), mock.call.update(), mock.call.destroy(),
        ])
        self._assert_no_selection()
        self.application_root.assert_not_called()

    def test_template_initial_directory_prefers_latest_then_legacy_then_app_root(self):
        legacy = self.app_root / "示例数据"
        previous = self.app_root / "示例数据_v0.2"
        previous_latest = self.app_root / "示例数据_v0.3"
        previous_current = self.app_root / "示例数据_v0.4"
        modern = self.app_root / "示例数据_v0.5"
        latest = self.app_root / "示例数据_v0.6"
        for expected, create in ((self.app_root, None), (legacy, legacy), (previous, previous),
                                 (previous_latest, previous_latest), (previous_current, previous_current),
                                 (modern, modern), (latest, latest)):
            with self.subTest(expected=expected.name):
                if create is not None:
                    create.mkdir()
                self.filedialog.askopenfilename.reset_mock()
                self.filedialog.askopenfilename.return_value = str(expected / "模板.xlsx")
                with dialogs.FileDialogs(self.app_root) as chooser:
                    selected = chooser.select_template()
                self.assertEqual(selected, str(expected / "模板.xlsx"))
                kwargs = self.filedialog.askopenfilename.call_args.kwargs
                self.assertEqual(kwargs["initialdir"], str(expected))
                self.assertIs(kwargs["parent"], self.root)
                self.assertEqual(kwargs["filetypes"], (("Excel 工作簿", "*.xlsx"),))

    def test_sources_prefer_normal_data_beside_template_and_allow_multiple(self):
        template_directory = self.app_root / "模板目录"
        normal_data = template_directory / "正常数据"
        normal_data.mkdir(parents=True)
        template = template_directory / "模板.xlsx"
        sources = [str(normal_data / "甲 有空格.xlsx"), str(normal_data / "乙.xlsx")]
        self.filedialog.askopenfilenames.return_value = sources
        with dialogs.FileDialogs(self.app_root) as chooser:
            self.assertEqual(chooser.select_sources(template), tuple(sources))
        kwargs = self.filedialog.askopenfilenames.call_args.kwargs
        self.assertEqual(kwargs["initialdir"], str(normal_data))
        self.assertIs(kwargs["parent"], self.root)
        self.assertEqual(kwargs["filetypes"], (("Excel 工作簿", "*.xlsx"),))

    def test_sources_fall_back_to_template_directory(self):
        template = self.app_root / "其他目录" / "模板.xlsx"
        with dialogs.FileDialogs(self.app_root) as chooser:
            self.assertEqual(chooser.select_sources(str(template)), ())
        self.assertEqual(
            self.filedialog.askopenfilenames.call_args.kwargs["initialdir"], str(template.parent),
        )
        self.assertFalse(template.parent.exists())

    def test_sources_compatibility_with_tcl_path_list_preserves_spaces(self):
        selected = "{C:/有 空格/甲.xlsx} C:/乙.xlsx"
        expected = ("C:/有 空格/甲.xlsx", "C:/乙.xlsx")
        self.filedialog.askopenfilenames.return_value = selected
        self.root.tk.splitlist.return_value = expected
        with dialogs.FileDialogs(self.app_root) as chooser:
            self.assertEqual(chooser.select_sources(self.app_root / "模板.xlsx"), expected)
        self.root.tk.splitlist.assert_called_once_with(selected)

    def test_output_asks_for_existing_directory_without_creating_anything(self):
        selected = str(self.app_root / "选定目录")
        self.filedialog.askdirectory.return_value = selected
        with dialogs.FileDialogs(self.app_root) as chooser:
            self.assertEqual(chooser.select_output_directory(), selected)
        kwargs = self.filedialog.askdirectory.call_args.kwargs
        self.assertIs(kwargs["parent"], self.root)
        self.assertEqual(kwargs["initialdir"], str(self.app_root))
        self.assertTrue(kwargs["mustexist"])
        self.assertEqual(list(self.app_root.iterdir()), [])

    def test_cancellation_results_are_empty_and_root_is_destroyed(self):
        with dialogs.FileDialogs(self.app_root) as chooser:
            self.assertEqual(chooser.select_template(), "")
            self.assertEqual(chooser.select_sources(self.app_root / "模板.xlsx"), ())
            self.assertEqual(chooser.select_output_directory(), "")
        self.root.destroy.assert_called_once_with()
        self.assertEqual(list(self.app_root.iterdir()), [])

    def test_dialog_exception_or_keyboard_interrupt_always_releases_root(self):
        for exception in (OSError("选择框不可用"), KeyboardInterrupt()):
            with self.subTest(exception=type(exception).__name__):
                self.root.reset_mock()
                self.filedialog.askopenfilename.side_effect = exception
                with self.assertRaises(type(exception)):
                    with dialogs.FileDialogs(self.app_root) as chooser:
                        chooser.select_template()
                self.root.destroy.assert_called_once_with()

    def test_initialization_failure_releases_root_even_before_enter_completes(self):
        for method in ("withdraw", "update"):
            with self.subTest(method=method):
                self.root.reset_mock(side_effect=True)
                getattr(self.root, method).side_effect = RuntimeError("Tk 初始化失败")
                chooser = dialogs.FileDialogs(self.app_root)
                with self.assertRaisesRegex(RuntimeError, "Tk 初始化失败"):
                    with chooser:
                        self.fail("初始化失败后不应进入上下文")
                self.root.destroy.assert_called_once_with()
                with self.assertRaises(RuntimeError):
                    chooser.select_template()
        self._assert_no_selection()

    def test_methods_require_context_and_do_not_create_an_implicit_root(self):
        chooser = dialogs.FileDialogs(self.app_root)
        for method, args in (
            (chooser.select_template, ()),
            (chooser.select_sources, (self.app_root / "模板.xlsx",)),
            (chooser.select_output_directory, ()),
        ):
            with self.subTest(method=method.__name__):
                with self.assertRaisesRegex(RuntimeError, "with FileDialogs"):
                    method(*args)
        self.tkinter.Tk.assert_not_called()
        self._assert_no_selection()

    def test_reentering_an_open_context_fails_but_outer_context_releases_root(self):
        chooser = dialogs.FileDialogs(self.app_root)
        with self.assertRaisesRegex(RuntimeError, "已打开"):
            with chooser:
                with chooser:
                    self.fail("不应重复创建根")
        self.tkinter.Tk.assert_called_once_with()
        self.root.destroy.assert_called_once_with()

    def test_check_dialogs_only_initializes_updates_and_destroys_hidden_root(self):
        self.assertIsNone(dialogs.check_dialogs())
        self.tkinter.Tk.assert_called_once_with()
        self.assertEqual(self.root.method_calls, [
            mock.call.withdraw(), mock.call.update(), mock.call.destroy(),
        ])
        self._assert_no_selection()
        self.assertEqual(list(self.app_root.iterdir()), [])

    def test_check_dialogs_failure_propagates_after_destroying_root(self):
        self.root.update.side_effect = RuntimeError("缺少 Tk 组件")
        with self.assertRaisesRegex(RuntimeError, "缺少 Tk 组件"):
            dialogs.check_dialogs()
        self.root.destroy.assert_called_once_with()
        self._assert_no_selection()


if __name__ == "__main__":
    unittest.main()
