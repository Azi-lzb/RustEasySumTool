"""两项菜单与同步汇总测试；所有文件选择和引擎调用均使用 mock。"""

from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest import mock

from summary_tool import cli


class UnreadIssues:
    def __iter__(self):
        raise AssertionError("完成摘要不应遍历问题清单")


class SummaryMenuTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="summary_cli_")
        self.addCleanup(temporary.cleanup)
        self.root_dir = Path(temporary.name).resolve()
        self.template = self.root_dir / "模板 有空格.xlsx"
        self.sources = (self.root_dir / "甲.xlsx", self.root_dir / "乙.xlsx")
        for path in (self.template, *self.sources):
            path.write_bytes(b"original workbook")
        self.output_dir = self.root_dir / "输出目录"
        self.output_path = self.output_dir / "结果.xlsx"
        self.result = SimpleNamespace(
            source_count=2,
            source_sheet_count=5,
            template_sheet_count=3,
            rule_count=9,
            sum_rule_count=4,
            wavg_rule_count=3,
            avg_rule_count=2,
            duration_seconds=0.25,
            output_path=self.output_path,
            issues=UnreadIssues(),
        )
        self.menu = cli.SummaryMenu(self.root_dir)
        self.printed = self._patch("builtins.print")
        self.factory = self._patch("summary_tool.cli.FileDialogs")
        self.context = self.factory.return_value
        self.dialogs = self.context.__enter__.return_value
        self._select_all()
        self.default_path = self._patch(
            "summary_tool.cli.default_output_path", return_value=self.output_path,
        )
        self.aggregate = self._patch("summary_tool.cli.aggregate", return_value=self.result)

    def _patch(self, target, **kwargs):
        patcher = mock.patch(target, **kwargs)
        self.addCleanup(patcher.stop)
        return patcher.start()

    def _select_all(self):
        self.dialogs.select_template.return_value = str(self.template)
        self.dialogs.select_sources.return_value = tuple(map(str, self.sources))
        self.dialogs.select_output_directory.return_value = str(self.output_dir)

    def _output(self):
        return "\n".join(" ".join(map(str, call.args)) for call in self.printed.call_args_list)

    def _run(self, inputs):
        with mock.patch("builtins.input", side_effect=inputs):
            return self.menu.run()

    def test_menu_has_exactly_two_options_and_exit_does_not_open_dialogs(self):
        with mock.patch("builtins.input", return_value="2") as read_input:
            self.assertEqual(cli.run_menu(), 0)
        read_input.assert_called_once_with()
        self.assertIn("Excel 模板汇总工具", self._output())
        self.assertNotIn("请选择", self._output())
        options = [line for line in self._output().splitlines() if line[:1].isdigit()]
        self.assertEqual(options, ["1. 按模板公式汇总", "2. 退出"])
        self.factory.assert_not_called()
        self.aggregate.assert_not_called()
        self.assertFalse(self.output_dir.exists())

    def test_old_menu_choices_and_help_are_invalid(self):
        self.assertEqual(self._run(["0", "3", "7", "帮助", "2"]), 0)
        self.assertIn("无效选项，请输入 1 或 2", self._output())
        self.assertNotIn("规则帮助", self._output())
        self.factory.assert_not_called()
        self.aggregate.assert_not_called()

    def test_selection_order_root_release_and_synchronous_engine_call(self):
        events = []

        def template():
            events.append("template")
            return str(self.template)

        def sources(template_path):
            self.assertEqual(template_path, str(self.template))
            events.append("sources")
            return tuple(map(str, self.sources))

        def output_directory():
            events.append("output")
            return str(self.output_dir)

        def release(*args):
            events.append("destroy")

        def output_path(*args):
            events.append("path")
            return self.output_path

        def aggregate(*args):
            events.append("aggregate")
            return self.result

        self.dialogs.select_template.side_effect = template
        self.dialogs.select_sources.side_effect = sources
        self.dialogs.select_output_directory.side_effect = output_directory
        self.context.__exit__.side_effect = release
        self.default_path.side_effect = output_path
        self.aggregate.side_effect = aggregate
        self.assertEqual(self._run(["1", "2"]), 0)
        self.assertEqual(events, ["template", "sources", "output", "destroy", "path", "aggregate"])
        self.factory.assert_called_once_with(self.root_dir)
        self.context.__exit__.assert_called_once_with(None, None, None)
        self.default_path.assert_called_once_with(str(self.template), str(self.output_dir))
        self.aggregate.assert_called_once_with(
            str(self.template), tuple(map(str, self.sources)), self.output_path,
        )
        for path in (self.template, *self.sources):
            self.assertEqual(path.read_bytes(), b"original workbook")
        self.assertFalse(self.output_dir.exists())

    def test_cancelling_each_dialog_returns_without_computing_or_creating_files(self):
        expected_calls = [
            [mock.call.select_template()],
            [mock.call.select_template(), mock.call.select_sources(str(self.template))],
            [mock.call.select_template(), mock.call.select_sources(str(self.template)),
             mock.call.select_output_directory()],
        ]
        for step in range(3):
            with self.subTest(step=step):
                self.factory.reset_mock()
                self._select_all()
                self.printed.reset_mock()
                if step == 0:
                    self.dialogs.select_template.return_value = ""
                elif step == 1:
                    self.dialogs.select_sources.return_value = ()
                else:
                    self.dialogs.select_output_directory.return_value = ""
                with mock.patch.object(cli, "print_selection") as selection:
                    self.assertEqual(self._run(["1", "2"]), 0)
                self.assertEqual(self.dialogs.method_calls, expected_calls[step])
                self.context.__exit__.assert_called_once_with(None, None, None)
                selection.assert_not_called()
                self.default_path.assert_not_called()
                self.aggregate.assert_not_called()
                self.assertFalse(self.output_dir.exists())
                self.assertIn("已取消，返回主菜单", self._output())
                self.assertEqual(set(self.root_dir.iterdir()), {self.template, *self.sources})

    def test_cancelled_second_job_never_reuses_first_job_selection(self):
        self.dialogs.select_template.side_effect = [str(self.template), ""]
        self.assertEqual(self._run(["1", "1", "2"]), 0)
        self.aggregate.assert_called_once()
        self.dialogs.select_sources.assert_called_once_with(str(self.template))
        self.dialogs.select_output_directory.assert_called_once_with()
        self.assertEqual(self.context.__exit__.call_count, 2)

    def test_selection_printer_outputs_only_paths_and_source_total(self):
        cli.print_selection(self.template, iter(self.sources), self.output_dir)
        self.assertEqual(self._output().splitlines(), [
            f"模板：{self.template}",
            "源文件：",
            f"  {self.sources[0]}",
            f"  {self.sources[1]}",
            "源文件总数：2",
            f"输出目录：{self.output_dir}",
        ])

    def test_success_summary_uses_all_statistics_without_reading_issues(self):
        cli.print_result(self.result)
        output = self._output()
        for expected in ("2 个源文件", "5 张源工作表", "3 张模板工作表", "9 条",
                         "SUM：4", "AVG：2", "WAVG：3", "0.25 秒", str(self.output_path)):
            self.assertIn(expected, output)
        for excluded in ("warning", "提醒", "原生公式", "日志", "读取第一份"):
            self.assertNotIn(excluded, output)
        self.assertEqual(len(output.splitlines()), 3)

    def test_generic_function_counts_do_not_need_a_new_legacy_field(self):
        self.result.function_counts = {"TESTDOUBLE": 7}
        self.result.rule_count = 7
        cli.print_result(self.result)
        self.assertIn("TESTDOUBLE：7", self._output())
        self.assertNotIn("SUM：", self._output())
        self.assertEqual(len(self._output().splitlines()), 3)

    def test_explicit_configuration_is_forwarded_without_another_dialog(self):
        config = self.root_dir / "独立配置.xlsx"
        self.menu.config_path = config
        self.assertEqual(self._run(["1", "2"]), 0)
        self.aggregate.assert_called_once_with(
            str(self.template), tuple(map(str, self.sources)), self.output_path, config_path=config,
        )
        self.assertEqual(self.dialogs.method_calls, [
            mock.call.select_template(), mock.call.select_sources(str(self.template)),
            mock.call.select_output_directory(),
        ])

    def test_failure_prints_only_brief_message_and_first_error_description(self):
        warning = SimpleNamespace(severity="warning", describe=mock.Mock())
        first = SimpleNamespace(
            severity="error", describe=mock.Mock(return_value="模板.xlsx / 明细!A1：规则无效"),
        )

        def issues():
            yield warning
            yield first
            raise AssertionError("找到首个错误后不应继续遍历")

        error = RuntimeError("模板规则有误\n不应显示的多行详情")
        error.issues = issues()
        cli.print_failure(error)
        self.assertEqual(self._output().splitlines(), [
            "汇总失败：模板规则有误",
            "首个错误位置：模板.xlsx / 明细!A1：规则无效",
        ])
        warning.describe.assert_not_called()
        first.describe.assert_called_once_with()

    def test_failure_without_error_issue_is_one_short_line(self):
        error = OSError("无写入权限" + "详情" * 200)
        warning = SimpleNamespace(severity="warning", describe=mock.Mock())
        error.issues = [warning]
        cli.print_failure(error)
        self.assertEqual(len(self._output().splitlines()), 1)
        self.assertLessEqual(len(self._output()), 245)
        warning.describe.assert_not_called()

    def test_aggregate_failure_returns_to_menu_without_warning_dump(self):
        warning = SimpleNamespace(severity="warning", describe=mock.Mock(return_value="详细提醒"))
        first = SimpleNamespace(
            severity="error", describe=mock.Mock(return_value="源.xlsx / 明细!B2：无缓存"),
        )
        later = SimpleNamespace(severity="error", describe=mock.Mock(return_value="后续错误"))
        self.aggregate.side_effect = cli.AggregationError("数据不可用", [warning, first, later])
        self.assertEqual(self._run(["1", "2"]), 0)
        self.assertIn("汇总失败：数据不可用", self._output())
        self.assertIn("明细!B2", self._output())
        self.assertNotIn("详细提醒", self._output())
        self.assertNotIn("后续错误", self._output())
        self.assertNotIn("汇总完成", self._output())
        warning.describe.assert_not_called()
        later.describe.assert_not_called()
        self.context.__exit__.assert_called_once_with(None, None, None)

    def test_dialog_failure_releases_context_and_does_not_start_engine(self):
        self.dialogs.select_sources.side_effect = RuntimeError("无法打开文件选择框")
        self.assertEqual(self._run(["1", "2"]), 0)
        self.context.__exit__.assert_called_once()
        self.dialogs.select_output_directory.assert_not_called()
        self.default_path.assert_not_called()
        self.aggregate.assert_not_called()
        self.assertIn("汇总失败：无法打开文件选择框", self._output())

    def test_existing_output_or_original_is_never_overwritten(self):
        existing = self.root_dir / "已有结果.xlsx"
        existing.write_bytes(b"keep existing")
        protected = [self.template, *self.sources, existing]
        snapshots = {path: path.read_bytes() for path in protected}
        for path in protected:
            with self.subTest(path=path.name):
                self.default_path.return_value = path
                self.assertEqual(self._run(["1", "2"]), 0)
                self.aggregate.assert_not_called()
        for path, content in snapshots.items():
            self.assertEqual(path.read_bytes(), content)
        self.assertIn("未覆盖原文件", self._output())

    def test_engine_cancellation_is_not_reported_as_failure(self):
        self.aggregate.side_effect = cli.AggregationCancelled()
        self.assertEqual(self._run(["1", "2"]), 0)
        self.assertIn("汇总已取消", self._output())
        self.assertNotIn("汇总失败", self._output())

    def test_menu_eof_and_keyboard_interrupt_exit_cleanly(self):
        self.assertEqual(self._run([EOFError]), 0)
        self.assertEqual(self._run([KeyboardInterrupt]), 130)
        self.assertIn("输入结束，已退出", self._output())
        self.factory.assert_not_called()

    def test_dialog_keyboard_interrupt_returns_to_menu_and_eof_exits(self):
        for exception, inputs in ((KeyboardInterrupt, ["1", "2"]), (EOFError, ["1"])):
            with self.subTest(exception=exception):
                self.factory.reset_mock()
                self.dialogs.select_template.side_effect = exception
                self.assertEqual(self._run(inputs), 0)
                self.context.__exit__.assert_called_once()
                self.aggregate.assert_not_called()
        self.assertIn("已取消当前操作", self._output())

    def test_interrupt_after_save_does_not_claim_no_output_or_delete_it(self):
        def interrupted(*args):
            self.output_dir.mkdir()
            self.output_path.write_bytes(b"possibly complete output")
            raise KeyboardInterrupt

        self.aggregate.side_effect = interrupted
        self.assertEqual(self._run(["1", "2"]), 0)
        self.assertEqual(self.output_path.read_bytes(), b"possibly complete output")
        self.assertIn("仍可能已生成文件", self._output())
        self.assertNotIn("没有生成", self._output())
        self.context.__exit__.assert_called_once_with(None, None, None)


if __name__ == "__main__":
    unittest.main()
