from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class Issue:
    severity: str
    code: str
    message: str
    source_file: str = ""
    source_sheet: str = ""
    source_cell: str = ""
    template_sheet: str = ""
    template_cell: str = ""

    def describe(self) -> str:
        labels = {"error": "错误", "warning": "提醒", "info": "信息"}
        parts = [f"[{labels.get(self.severity, self.severity)}] {self.message}"]
        if self.source_file:
            parts.append(f"文件：{Path(self.source_file).name}")
        if self.source_sheet:
            parts.append(f"源位置：{self.source_sheet}!{self.source_cell}")
        if self.template_sheet:
            parts.append(f"模板：{self.template_sheet}!{self.template_cell}")
        return " | ".join(parts)


class AggregationError(Exception):
    def __init__(self, message: str, issues: list[Issue] | None = None):
        super().__init__(message)
        self.issues = issues or []


class AggregationCancelled(AggregationError):
    def __init__(self):
        super().__init__("已取消汇总，没有发布结果文件。")


@dataclass
class AggregationResult:
    output_path: Path
    source_count: int
    rule_count: int
    issues: list[Issue] = field(default_factory=list)
    duration_seconds: float = 0.0
    source_sheet_count: int = 0
    template_sheet_count: int = 0
    sum_rule_count: int = 0
    wavg_rule_count: int = 0
    avg_rule_count: int = 0
    function_counts: dict[str, int] = field(default_factory=dict)
    function_bindings: dict[str, str] = field(default_factory=dict)
    config_path: Path | None = None

    @property
    def warning_count(self) -> int:
        return sum(issue.severity == "warning" for issue in self.issues)


def result_function_counts(result) -> dict[str, int]:
    from .functions import function_names

    counts = getattr(result, "function_counts", None) or {
        "SUM": getattr(result, "sum_rule_count", 0),
        "AVG": getattr(result, "avg_rule_count", 0),
        "WAVG": getattr(result, "wavg_rule_count", 0),
    }
    order = dict.fromkeys((*function_names(), *counts))
    return {name: counts[name] for name in order if counts.get(name, 0)}
