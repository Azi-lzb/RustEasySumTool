"""v0.2 两项终端菜单；文件选择完成后同步汇总，结果始终另存。"""

from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path

from . import APP_NAME
from .dialogs import FileDialogs
from .engine import AggregationCancelled, AggregationError, aggregate, default_output_path
from .models import result_function_counts


def print_selection(
    template: str | Path,
    sources: Iterable[str | Path],
    output_dir: str | Path,
) -> None:
    """仅显示本次选择的路径及源文件总数，供菜单和批处理共用。"""
    print(f"模板：{template}")
    print("源文件：")
    count = 0
    for source in sources:
        print(f"  {source}")
        count += 1
    print(f"源文件总数：{count}")
    print(f"输出目录：{output_dir}")


def print_result(result) -> None:
    """输出完成统计；不遍历问题清单或展开逐文件提醒。"""
    print(
        f"汇总完成：{result.source_count} 个源文件、"
        f"{result.source_sheet_count} 张源工作表；"
        f"{result.template_sheet_count} 张模板工作表。"
    )
    counts = "，".join(f"{name}：{count}" for name, count in result_function_counts(result).items())
    print(f"规则：{result.rule_count} 条（{counts}）；用时 {result.duration_seconds:.2f} 秒。")
    print(f"结果已保存：{result.output_path}")


def _brief(value) -> str:
    """避免多行异常消息把完整问题清单带入终端。"""
    lines = str(value).strip().splitlines()
    first_line = lines[0] if lines else ""
    return first_line if len(first_line) <= 240 else first_line[:237] + "..."


def print_failure(exc: BaseException) -> None:
    """模板静态错误一次列全，其余失败保留简短摘要。"""
    print(f"汇总失败：{_brief(exc) or type(exc).__name__}")
    issues = iter(getattr(exc, "issues", ()))
    first_error = next(
        (issue for issue in issues if issue.severity == "error"),
        None,
    )
    static_codes = {"invalid_rule", "template_formula_not_supported", "template_excel_error"}
    if first_error is not None and getattr(first_error, "code", "") in static_codes:
        print(first_error.describe())
        for issue in issues:
            if issue.severity == "error" and issue.code in static_codes:
                print(issue.describe())
        return
    if first_error is not None:
        print(f"首个错误位置：{_brief(first_error.describe())}")


class SummaryMenu:
    """只保留选择并汇总、退出；不保存上一次选择以免取消后误用。"""

    def __init__(self, app_root: str | Path | None = None, config_path: str | Path | None = None) -> None:
        self.app_root = app_root
        self.config_path = config_path

    @staticmethod
    def print_menu() -> None:
        print(f"\n{APP_NAME}")
        print("1. 按模板公式汇总")
        print("2. 退出")

    def run(self) -> int:
        while True:
            try:
                self.print_menu()
                choice = input().strip()
                if choice == "2":
                    print("已退出。")
                    return 0
                if choice != "1":
                    print("无效选项，请输入 1 或 2。")
                    continue
                try:
                    self.start_aggregation()
                except KeyboardInterrupt:
                    print("\n已取消当前操作，返回主菜单。")
                except EOFError:
                    raise
                except AggregationCancelled:
                    print("汇总已取消，返回主菜单。")
                except Exception as exc:
                    print_failure(exc)
            except EOFError:
                print("\n输入结束，已退出。")
                return 0
            except KeyboardInterrupt:
                print("\n已退出。")
                return 130

    def start_aggregation(self) -> bool:
        with FileDialogs(self.app_root) as dialogs:
            template = dialogs.select_template()
            if not template:
                print("已取消，返回主菜单。")
                return False
            sources = dialogs.select_sources(template)
            if not sources:
                print("已取消，返回主菜单。")
                return False
            output_dir = dialogs.select_output_directory()
            if not output_dir:
                print("已取消，返回主菜单。")
                return False

        # 先释放 Tk，全部选择完成前不生成路径或调用引擎。
        print_selection(template, sources, output_dir)
        output_path = default_output_path(template, output_dir)
        if Path(output_path).exists():
            raise ValueError(f"输出文件已存在，未覆盖原文件：{output_path}")
        try:
            if self.config_path is None:
                result = aggregate(template, sources, output_path)
            else:
                result = aggregate(template, sources, output_path, config_path=self.config_path)
        except KeyboardInterrupt:
            print(f"\n汇总已中断。保存阶段仍可能已生成文件，请检查：{output_path}")
            return False
        print_result(result)
        return True


def run_menu(config_path: str | Path | None = None) -> int:
    """正常退出或 EOF 返回 0；菜单处 Ctrl+C 返回 130。"""
    return SummaryMenu().run() if config_path is None else SummaryMenu(config_path=config_path).run()
