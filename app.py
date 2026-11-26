import argparse
import os
from pathlib import Path
import sys

from summary_tool import APP_NAME, __version__


def main(argv=None):
    parser = argparse.ArgumentParser(description=f"{APP_NAME}，支持函数别名配置，只输出结果值。")
    parser.add_argument("--version", action="version", version=f"{APP_NAME} {__version__}")
    parser.add_argument("--template", type=Path, help=".xlsx 模板文件")
    parser.add_argument("--sources", type=Path, nargs="+", help="一个或多个 .xlsx 源文件")
    parser.add_argument("--output", type=Path, help="不存在的 .xlsx 结果路径")
    parser.add_argument("--config", type=Path, help="外置Excel配置，默认程序同目录config.xlsx")
    parser.add_argument("--check-dialogs", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    if args.check_dialogs:
        from summary_tool.dialogs import check_dialogs
        try:
            check_dialogs()
        except Exception as exc:
            print(f"文件选择框组件不可用：{exc}", file=sys.stderr)
            return 2
        print("文件选择框组件就绪。")
        return 0
    if any((args.template, args.sources, args.output)):
        if not all((args.template, args.sources, args.output)):
            parser.error("命令行汇总需要同时提供 --template、--sources 和 --output。")
        from summary_tool.cli import print_failure, print_result, print_selection
        from summary_tool.engine import AggregationError, aggregate
        print_selection(args.template, args.sources, args.output.parent)
        try:
            if args.config is None:
                result = aggregate(args.template, args.sources, args.output)
            else:
                result = aggregate(args.template, args.sources, args.output, config_path=args.config)
        except AggregationError as exc:
            print_failure(exc)
            return 2
        print_result(result)
        return 0
    from summary_tool.cli import run_menu
    return run_menu() if args.config is None else run_menu(config_path=args.config)


if __name__ == "__main__":
    if os.name == "nt":
        import ctypes
        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        ctypes.windll.kernel32.SetConsoleCP(65001)
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\n操作已中断；若已进入保存阶段，请检查输出目录。既有文件不会被覆盖。", file=sys.stderr)
        raise SystemExit(130)
