from collections.abc import Mapping
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from types import MappingProxyType

from openpyxl import load_workbook

from .functions import function_names, get_function, is_alias
from .paths import application_root


DEFAULT_BINDINGS = {
    "FUN_1": "SUM",
    "FUN_2": "WAVG",
    "FUN_3": "AVG",
    "FUN_4": "MAX",
    "FUN_5": "MIN",
    "FUN_6": "COUNT",
    "FUN_7": "SUMPRODUCT",
    "FUN_8": "RATIO",
}

MAPPING_SHEET = "函数映射"
MAPPING_HEADERS = ("函数别名", "映射函数")

# 已有模板继续使用其原有名字和映射，配置可显式调整任一绑定。
LEGACY_BINDINGS = {
    "F1_1": "SUM",
    "F1_2": "AVG",
    "F1_3": "MAX",
    "F1_4": "MIN",
    "F1_5": "COUNT",
    "F2_1": "WAVG",
}


class ConfigurationError(ValueError):
    """配置文件无法读取、解析或通过校验。"""


@dataclass(frozen=True)
class ConfigurationSnapshot:
    bindings: Mapping[str, str]
    path: Path | None = None

    def __post_init__(self):
        object.__setattr__(self, "bindings", MappingProxyType(dict(self.bindings)))


def _error(path: Path, message: str, alias: str | None = None) -> ConfigurationError:
    location = f"配置文件“{path}”"
    if alias is not None:
        location += f"中的别名 {alias!r}"
    return ConfigurationError(f"{location}：{message}")


def _xlsx_entries(content: bytes, config_path: Path):
    try:
        workbook = load_workbook(BytesIO(content), read_only=True, data_only=False, keep_links=False)
        try:
            if MAPPING_SHEET not in workbook.sheetnames:
                raise _error(config_path, f"缺少“{MAPPING_SHEET}”工作表，请在该页的 A/B 列填写别名和实际算法。")
            entries = []
            sheet = workbook[MAPPING_SHEET]
            for row_number, cells in enumerate(sheet.iter_rows(min_col=1, max_col=2), 1):
                location = f"工作表“{MAPPING_SHEET}”第 {row_number} 行（A{row_number}/B{row_number}）"
                if any(cell.data_type in {"f", "e"} for cell in cells):
                    raise _error(config_path, f"{location}包含 Excel 公式或错误值，请直接填写文本。")
                alias, target = (cell.value.strip() if isinstance(cell.value, str) else cell.value for cell in cells)
                if row_number == 1 and (alias, target) == MAPPING_HEADERS:
                    continue
                if alias in (None, "") and target in (None, ""):
                    continue
                entries.append((alias, target, location + "，"))
            return entries
        finally:
            workbook.close()
    except ConfigurationError:
        raise
    except Exception as exc:
        raise _error(config_path, f"无法读取 Excel 配置，请检查文件是否为有效的 .xlsx 工作簿：{exc}") from exc


def load_configuration(path: str | Path | None = None) -> ConfigurationSnapshot:
    explicit_path = path is not None
    config_path = Path(path) if explicit_path else application_root() / "config.xlsx"
    try:
        config_path = config_path.resolve()
        if config_path.suffix.lower() != ".xlsx":
            raise _error(config_path, "配置文件扩展名必须为 .xlsx。")
        content = config_path.read_bytes()
    except FileNotFoundError as exc:
        if not explicit_path:
            return ConfigurationSnapshot({}, None)
        raise _error(config_path, "文件不存在。") from exc
    except ConfigurationError:
        raise
    except (OSError, ValueError) as exc:
        raise _error(config_path, f"无法读取配置：{exc}") from exc

    entries = _xlsx_entries(content, config_path)
    bindings: dict[str, str] = {}
    for source_alias, target, location in entries:
        if not isinstance(source_alias, str) or not source_alias:
            raise _error(config_path, f"{location}别名必须是非空文本。", source_alias)
        alias = source_alias.upper()
        if get_function(alias) is not None:
            raise _error(config_path, f"{location}不能用别名覆盖内置算法名。", source_alias)
        if not is_alias(alias):
            raise _error(
                config_path,
                f"{location}名称必须为 FUN_正整数（不带前导 0），也接受旧格式 F正整数_正整数。",
                source_alias,
            )
        if alias in bindings:
            raise _error(
                config_path, f"{location}与 {alias!r} 重复，别名不区分大小写。", source_alias,
            )
        if not isinstance(target, str) or not target.strip():
            raise _error(config_path, f"{location}值必须是非空的已注册算法名字符串。", source_alias)

        specification = get_function(target.upper())
        if specification is None:
            names = ", ".join(function_names())
            raise _error(
                config_path,
                f"{location}未知算法 {target!r}，目标只能是已注册算法（{names}），不支持别名链。",
                source_alias,
            )
        bindings[alias] = specification.name.upper()

    return ConfigurationSnapshot(bindings, config_path)
