"""Generate v0.6 alias demo inputs without launching the application or engine."""

from __future__ import annotations

import argparse
from collections.abc import Sequence
from pathlib import Path
import sys

from summary_tool.demo import create_demo_assets


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="生成v0.6纯虚构中文别名模板、正常数据、异常数据和独立预期值；不执行汇总，不生成配置副本。")
    parser.add_argument(
        "base_dir",
        nargs="?",
        type=Path,
        default=Path(__file__).resolve().parent / "示例数据",
        help="示例数据目录本身，默认使用本脚本旁的“示例数据”。",
    )
    parser.add_argument("--force", action="store_true", help="显式允许覆盖同名示例文件；默认保留所有已存在文件。")
    arguments = parser.parse_args(argv)
    try:
        assets = create_demo_assets(arguments.base_dir, force=arguments.force)
    except OSError as exc:
        print(f"生成示例失败：{exc}", file=sys.stderr)
        return 1
    print(f"示例数据目录：{assets.root}")
    print("覆盖策略：" + ("已显式允许覆盖同名文件。" if arguments.force else "已存在文件一律保留；缺少的文件已补齐。"))
    for path in (assets.template_path, *assets.source_paths, *assets.error_paths, assets.root / "预期结果.json"):
        print(f"  {path.relative_to(assets.root)}")
    print(f"独立预期值：{len(assets.expected_values)}个数值结果（已含合计），无需Excel公式重算。")
    print("别名默认绑定：FUN_1=SUM、FUN_3=AVG、FUN_2=WAVG、FUN_4=MAX、FUN_5=MIN、FUN_6=COUNT；配置仅使用主程序同目录config.xlsx，不在模板目录生成副本。")
    print("模板业务明细B:J；H/I/J5:10为余额极值及有效填报数，13:15行为独立诊断，不参加第10行原项目合计。")
    print("全部数据纯属虚构；预期值仅对应初始虚构数据和默认绑定。未运行应用，也未生成汇总结果。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
