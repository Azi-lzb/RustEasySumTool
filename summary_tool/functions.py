from __future__ import annotations

from collections.abc import Callable, Iterator, Sequence
from dataclasses import dataclass
from decimal import Decimal
import re
from typing import TYPE_CHECKING, Protocol

if TYPE_CHECKING:
    from .rules import CellReference


ALIAS_PATTERN = re.compile(r"(?:FUN_[1-9][0-9]*|F[1-9][0-9]*_[1-9][0-9]*)", re.IGNORECASE)


def is_alias(name: str) -> bool:
    return ALIAS_PATTERN.fullmatch(name) is not None


class NumberInput(Protocol):
    value: Decimal
    blank: bool


@dataclass(frozen=True)
class CalculationNotice:
    severity: str
    code: str
    message: str
    parameter_index: int | None = None


@dataclass(frozen=True)
class ResultSource:
    source_index: int
    parameter_index: int = 0


@dataclass(frozen=True)
class FunctionResult:
    value: Decimal
    notices: tuple[CalculationNotice, ...] = ()
    sources: tuple[ResultSource, ...] = ()
    source_note: str = ""


class Calculator(Protocol):
    def add(self, values: Sequence[NumberInput]) -> tuple[CalculationNotice, ...]: ...
    def finish(self, source_count: int) -> FunctionResult: ...


@dataclass(frozen=True)
class FunctionSpec:
    name: str
    parameter_names: tuple[str, ...]
    factory: Callable[[], Calculator]
    variadic: bool = False
    allow_ranges: bool = False
    matching_shapes: bool = False

    @property
    def arity(self) -> int:
        """固定签名的参数数，或可变签名的最少参数数。"""
        return len(self.parameter_names)

    @property
    def parameter_count(self) -> str:
        return f"至少 {self.arity}" if self.variadic else str(self.arity)

    def accepts_count(self, count: int) -> bool:
        return count >= self.arity if self.variadic else count == self.arity

    def signature(self, name: str | None = None) -> str:
        parameters = ','.join(self.parameter_names)
        if self.variadic:
            parameters += ",..."
        return f"{name or self.name}({parameters})"

    def validate_references(self, references: Sequence[CellReference]) -> None:
        if not self.allow_ranges and any(reference.shape != (1, 1) for reference in references):
            raise ValueError(f"{self.name} 的参数只接受单个单元格引用。")
        if self.matching_shapes and len({reference.shape for reference in references}) > 1:
            shapes = "、".join(f"{rows}×{columns}" for rows, columns in (ref.shape for ref in references))
            raise ValueError(f"{self.name} 的参数区域必须具有相同的行数和列数，实际为 {shapes}。")

    def reference_rows(self, references: Sequence[CellReference]) -> Iterator[tuple[CellReference, ...]]:
        """SUM 逐参数累计，固定签名按各区域内相同偏移配对。"""
        if self.variadic:
            for reference in references:
                for cell in reference.cells():
                    yield (cell,)
        else:
            yield from zip(*(reference.cells() for reference in references))


class SumCalculator:
    def __init__(self):
        self.total = Decimal(0)

    def add(self, values: Sequence[NumberInput]) -> tuple[CalculationNotice, ...]:
        self.total += sum((number.value for number in values), Decimal(0))
        return ()

    def finish(self, source_count: int) -> FunctionResult:
        return FunctionResult(self.total)


class AvgCalculator(SumCalculator):
    def finish(self, source_count: int) -> FunctionResult:
        if source_count <= 0:
            raise ValueError("AVG 至少需要一份有效源文件。")
        return FunctionResult(self.total / Decimal(source_count))


class SumProductCalculator(SumCalculator):
    def add(self, values: Sequence[NumberInput]) -> tuple[CalculationNotice, ...]:
        first, second = values
        self.total += first.value * second.value
        return ()


class RatioCalculator:
    def __init__(self):
        self.numerator = Decimal(0)
        self.denominator = Decimal(0)

    def add(self, values: Sequence[NumberInput]) -> tuple[CalculationNotice, ...]:
        numerator, denominator = values
        self.numerator += numerator.value
        self.denominator += denominator.value
        return ()

    def finish(self, source_count: int) -> FunctionResult:
        if not self.denominator:
            notice = CalculationNotice(
                "error", "zero_denominator", "RATIO 分母合计为0，无法计算合计比率，请核对空白、零值或正负抵消。",
            )
            return FunctionResult(Decimal(0), (notice,))
        return FunctionResult(self.numerator / self.denominator)


class WeightedAverageCalculator:
    def __init__(self):
        self.numerator = Decimal(0)
        self.denominator = Decimal(0)

    def add(self, values: Sequence[NumberInput]) -> tuple[CalculationNotice, ...]:
        rate, weight = values
        self.numerator += rate.value * weight.value
        self.denominator += weight.value
        notices = []
        if rate.blank and weight.value:
            notices.append(CalculationNotice(
                "warning", "blank_rate_with_weight", "利率空白按0%，权重仍进入分母。", 0,
            ))
        if weight.value < 0:
            notices.append(CalculationNotice(
                "warning", "negative_weight", "负权重按原值参与计算，请核对业务口径。", 1,
            ))
        return tuple(notices)

    def finish(self, source_count: int) -> FunctionResult:
        if not self.denominator:
            notice = CalculationNotice(
                "error", "zero_weight",
                "WAVG 权重合计为0，无法计算加权平均值，请核对空白、零权重或正负抵消。",
            )
            return FunctionResult(Decimal(0), (notice,))
        return FunctionResult(self.numerator / self.denominator)


class _ExtremeCalculator:
    def __init__(self, *, maximum: bool):
        self.maximum = maximum
        self.ordinal = -1
        self.best: Decimal | None = None
        self.sources: list[ResultSource] = []

    def add(self, values: Sequence[NumberInput]) -> tuple[CalculationNotice, ...]:
        self.ordinal += 1
        number = values[0]
        if number.blank:
            return ()
        value = number.value
        if self.best is None or (
            value > self.best if self.maximum else value < self.best
        ):
            self.best = value
            self.sources = [ResultSource(self.ordinal)]
        elif value == self.best:
            self.sources.append(ResultSource(self.ordinal))
        return ()

    def finish(self, source_count: int) -> FunctionResult:
        if self.best is not None:
            return FunctionResult(self.best, sources=tuple(self.sources))
        source_note = "无有效数值，无极值来源，按约定输出0"
        notice = CalculationNotice("info", "no_valid_values", source_note)
        return FunctionResult(Decimal(0), (notice,), source_note=source_note)


class MaxCalculator(_ExtremeCalculator):
    def __init__(self):
        super().__init__(maximum=True)


class MinCalculator(_ExtremeCalculator):
    def __init__(self):
        super().__init__(maximum=False)


class CountCalculator:
    def __init__(self):
        self.count = 0

    def add(self, values: Sequence[NumberInput]) -> tuple[CalculationNotice, ...]:
        if not values[0].blank:
            self.count += 1
        return ()

    def finish(self, source_count: int) -> FunctionResult:
        return FunctionResult(Decimal(self.count))


FUNCTIONS: dict[str, FunctionSpec] = {
    "SUM": FunctionSpec("SUM", ("取数单元格或区域",), SumCalculator, variadic=True, allow_ranges=True),
    "AVG": FunctionSpec("AVG", ("取数单元格",), AvgCalculator),
    "WAVG": FunctionSpec("WAVG", ("数值组", "权重组"), WeightedAverageCalculator,
                         allow_ranges=True, matching_shapes=True),
    "MAX": FunctionSpec("MAX", ("取数单元格",), MaxCalculator),
    "MIN": FunctionSpec("MIN", ("取数单元格",), MinCalculator),
    "COUNT": FunctionSpec("COUNT", ("取数单元格",), CountCalculator),
    "SUMPRODUCT": FunctionSpec("SUMPRODUCT", ("因子1单元格", "因子2单元格"), SumProductCalculator),
    "RATIO": FunctionSpec("RATIO", ("分子单元格", "分母单元格"), RatioCalculator),
}


def get_function(name: str) -> FunctionSpec | None:
    return FUNCTIONS.get(name.upper())


def function_names() -> tuple[str, ...]:
    return tuple(FUNCTIONS)
