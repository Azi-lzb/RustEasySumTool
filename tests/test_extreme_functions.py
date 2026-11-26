from dataclasses import FrozenInstanceError, dataclass
from decimal import Decimal
import unittest

from summary_tool import functions
from summary_tool.functions import CalculationNotice, FunctionResult, ResultSource


@dataclass(frozen=True)
class _Number:
    value: Decimal
    blank: bool = False


def _number(value="0", *, blank=False):
    return _Number(Decimal(str(value)), blank)


class ExtremeCalculatorTests(unittest.TestCase):
    def calculate(self, name, numbers):
        calculator = functions.get_function(name).factory()
        for number in numbers:
            self.assertEqual(calculator.add((number,)), ())
        result = calculator.finish(len(numbers))
        self.assertIsInstance(result.value, Decimal)
        return result

    def test_max_of_all_positive_values(self):
        result = self.calculate("MAX", [_number(v) for v in ("1.2", "3.4", "2.3")])
        self.assertEqual(result, FunctionResult(Decimal("3.4"), sources=(ResultSource(1),)))

    def test_max_of_all_negative_values_is_not_initial_zero(self):
        result = self.calculate("MAX", [_number(v) for v in (-9, -2, -5)])
        self.assertEqual(result, FunctionResult(Decimal(-2), sources=(ResultSource(1),)))

    def test_min_of_all_positive_values_is_not_initial_zero(self):
        result = self.calculate("MIN", [_number(v) for v in ("2.3", "1.2", "3.4")])
        self.assertEqual(result, FunctionResult(Decimal("1.2"), sources=(ResultSource(1),)))

    def test_min_of_all_negative_values(self):
        result = self.calculate("MIN", [_number(v) for v in (-2, -9, -5)])
        self.assertEqual(result, FunctionResult(Decimal(-9), sources=(ResultSource(1),)))

    def test_real_zero_and_mixed_sign_values_participate(self):
        cases = (
            ("MAX", (-2, 0, -4), 0, 1),
            ("MIN", (2, 0, 4), 0, 1),
            ("MAX", (-3, 0, 4), 4, 2),
            ("MIN", (-3, 0, 4), -3, 0),
        )
        for name, values, expected, source_index in cases:
            with self.subTest(name=name, values=values):
                result = self.calculate(name, [_number(v) for v in values])
                self.assertEqual(
                    result,
                    FunctionResult(Decimal(expected), sources=(ResultSource(source_index),)),
                )

    def test_blanks_before_between_and_after_values_keep_source_ordinals(self):
        for name, values, expected in (("MAX", (-3, -1), -1), ("MIN", (3, 1), 1)):
            with self.subTest(name=name):
                result = self.calculate(name, [
                    _number(blank=True), _number(values[0]), _number(blank=True),
                    _number(values[1]), _number(blank=True),
                ])
                self.assertEqual(
                    result, FunctionResult(Decimal(expected), sources=(ResultSource(3),)),
                )

    def test_blank_flag_excludes_values_from_comparison(self):
        for name in ("MAX", "MIN"):
            with self.subTest(name=name):
                result = self.calculate(name, [
                    _number(999, blank=True), _number(7), _number(-999, blank=True),
                ])
                self.assertEqual(result, FunctionResult(Decimal(7), sources=(ResultSource(1),)))

    def test_all_blank_extrema_return_zero_with_explicit_no_source_info(self):
        message = "无有效数值，无极值来源，按约定输出0"
        for name in ("MAX", "MIN"):
            with self.subTest(name=name):
                result = self.calculate(name, [_number(blank=True) for _ in range(3)])
                self.assertEqual(result.value, Decimal(0))
                self.assertEqual(result.sources, ())
                self.assertEqual(result.source_note, message)
                self.assertEqual(
                    result.notices,
                    (CalculationNotice("info", "no_valid_values", message),),
                )
                self.assertEqual(result.notices[0].message, result.source_note)

    def test_extrema_without_any_additions_also_have_no_valid_values(self):
        for name in ("MAX", "MIN"):
            with self.subTest(name=name):
                result = self.calculate(name, [])
                self.assertEqual(result.value, Decimal(0))
                self.assertEqual(result.sources, ())
                self.assertEqual(len(result.notices), 1)
                self.assertEqual(result.notices[0].code, "no_valid_values")
                self.assertEqual(result.notices[0].severity, "info")
                self.assertEqual(result.notices[0].message, result.source_note)

    def test_tied_max_keeps_every_equal_source_and_discards_old_best(self):
        result = self.calculate("MAX", [
            _number(1), _number(9), _number(-1), _number(blank=True),
            _number("9.00"), _number(4), _number("9.0"),
        ])
        self.assertEqual(
            result,
            FunctionResult(Decimal(9), sources=(ResultSource(1), ResultSource(4), ResultSource(6))),
        )

    def test_tied_min_keeps_every_equal_source_and_discards_old_best(self):
        result = self.calculate("MIN", [
            _number(9), _number(-1), _number(4), _number(blank=True),
            _number("-1.00"), _number(0), _number("-1.0"),
        ])
        self.assertEqual(
            result,
            FunctionResult(Decimal(-1), sources=(ResultSource(1), ResultSource(4), ResultSource(6))),
        )

    def test_blank_zero_is_not_a_tied_source_for_real_zero(self):
        for name in ("MAX", "MIN"):
            with self.subTest(name=name):
                result = self.calculate(name, [
                    _number(blank=True), _number(0), _number(blank=True),
                    _number("-0.00"), _number(blank=True),
                ])
                self.assertEqual(
                    result,
                    FunctionResult(Decimal(0), sources=(ResultSource(1), ResultSource(3))),
                )

    def test_max_compares_actual_decimals_not_display_rounding(self):
        values = [Decimal("1.2341"), Decimal("1.2342")]
        self.assertEqual(values[0].quantize(Decimal("0.01")), values[1].quantize(Decimal("0.01")))
        result = self.calculate("MAX", [_number(v) for v in values])
        self.assertEqual(result, FunctionResult(values[1], sources=(ResultSource(1),)))

    def test_min_compares_actual_decimals_not_display_rounding(self):
        values = [Decimal("1.2342"), Decimal("1.2341")]
        self.assertEqual(values[0].quantize(Decimal("0.01")), values[1].quantize(Decimal("0.01")))
        result = self.calculate("MIN", [_number(v) for v in values])
        self.assertEqual(result, FunctionResult(values[1], sources=(ResultSource(1),)))

    def test_finish_returns_a_snapshot_and_does_not_advance_the_ordinal(self):
        for name, better in (("MAX", 6), ("MIN", 4)):
            with self.subTest(name=name):
                calculator = functions.get_function(name).factory()
                calculator.add((_number(5),))
                first = calculator.finish(1)
                self.assertEqual(calculator.finish(1), first)
                calculator.add((_number("5.00"),))
                tied = calculator.finish(2)
                self.assertEqual(first.sources, (ResultSource(0),))
                self.assertEqual(tied.sources, (ResultSource(0), ResultSource(1)))
                calculator.add((_number(better),))
                self.assertEqual(calculator.finish(3).sources, (ResultSource(2),))
                self.assertEqual(first.sources, (ResultSource(0),))
                self.assertEqual(tied.sources, (ResultSource(0), ResultSource(1)))

    def test_new_calculators_do_not_share_numeric_or_source_state(self):
        for name in ("MAX", "MIN", "COUNT"):
            with self.subTest(name=name):
                first = functions.get_function(name).factory()
                second = functions.get_function(name).factory()
                first.add((_number(blank=True),))
                first.add((_number(7),))
                first_result = first.finish(2)
                empty_result = second.finish(0)
                self.assertEqual(empty_result.value, Decimal(0))
                self.assertEqual(empty_result.sources, ())
                second.add((_number(3),))
                second_result = second.finish(1)
                self.assertEqual(first.finish(2), first_result)
                if name == "COUNT":
                    self.assertEqual(first_result.value, Decimal(1))
                    self.assertEqual(second_result.value, Decimal(1))
                else:
                    self.assertEqual(first_result.value, Decimal(7))
                    self.assertEqual(first_result.sources, (ResultSource(1),))
                    self.assertEqual(second_result.value, Decimal(3))
                    self.assertEqual(second_result.sources, (ResultSource(0),))

    def test_count_includes_real_zero_negative_and_positive_nonblank_values(self):
        result = self.calculate("COUNT", [
            _number(blank=True), _number(0), _number(-4), _number("2.5"),
            _number(blank=True), _number("0.00"),
        ])
        self.assertEqual(result, FunctionResult(Decimal(4)))

    def test_count_all_blank_is_zero_without_a_no_value_notice(self):
        result = self.calculate("COUNT", [_number(blank=True) for _ in range(3)])
        self.assertEqual(result, FunctionResult(Decimal(0)))

    def test_count_without_any_additions_is_zero_without_a_notice(self):
        self.assertEqual(self.calculate("COUNT", []), FunctionResult(Decimal(0)))

    def test_count_uses_nonblank_additions_not_finish_source_count(self):
        calculator = functions.get_function("COUNT").factory()
        calculator.add((_number(blank=True),))
        calculator.add((_number(0),))
        for source_count in (0, 2, 99):
            with self.subTest(source_count=source_count):
                self.assertEqual(calculator.finish(source_count), FunctionResult(Decimal(1)))


class ResultSourceCompatibilityTests(unittest.TestCase):
    def test_result_source_has_a_default_parameter_and_is_frozen(self):
        source = ResultSource(3)
        self.assertEqual((source.source_index, source.parameter_index), (3, 0))
        self.assertEqual(source, ResultSource(3, 0))
        self.assertEqual(ResultSource(3, 1).parameter_index, 1)
        with self.assertRaises(FrozenInstanceError):
            source.source_index = 4
        with self.assertRaises(FrozenInstanceError):
            source.parameter_index = 1

    def test_existing_positional_function_result_construction_stays_compatible(self):
        value = Decimal("1.25")
        notices = (CalculationNotice("info", "existing", "existing notice"),)
        for result in (FunctionResult(value), FunctionResult(value, notices)):
            with self.subTest(notices=result.notices):
                self.assertEqual(result.value, value)
                self.assertEqual(result.sources, ())
                self.assertEqual(result.source_note, "")
        self.assertEqual(FunctionResult(value).notices, ())
        self.assertEqual(FunctionResult(value, notices).notices, notices)

    def test_new_source_fields_follow_existing_positional_arguments(self):
        value = Decimal(7)
        notices = (CalculationNotice("info", "custom", "custom notice"),)
        sources = (ResultSource(2), ResultSource(5))
        result = FunctionResult(value, notices, sources, "source explanation")
        self.assertEqual(
            result,
            FunctionResult(value=value, notices=notices, sources=sources, source_note="source explanation"),
        )
        self.assertEqual(result.sources, (ResultSource(2), ResultSource(5)))
        with self.assertRaises(FrozenInstanceError):
            result.source_note = "changed"


if __name__ == "__main__":
    unittest.main()
