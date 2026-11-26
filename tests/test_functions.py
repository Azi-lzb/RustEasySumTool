from dataclasses import dataclass
from decimal import Decimal
import unittest

from summary_tool import functions
from summary_tool.functions import CalculationNotice, FunctionResult, FunctionSpec


@dataclass(frozen=True)
class _Number:
    value: Decimal
    blank: bool = False


def _number(value="0", *, blank=False):
    return _Number(Decimal(str(value)), blank)


class FunctionRegistryTests(unittest.TestCase):
    def test_builtin_metadata_lookup_and_factories_are_independent(self):
        names = functions.function_names()
        self.assertIsInstance(names, tuple)
        self.assertEqual(names, tuple(functions.FUNCTIONS))
        self.assertEqual(names, ("SUM", "AVG", "WAVG", "MAX", "MIN", "COUNT", "SUMPRODUCT", "RATIO"))
        for name, arity in (
            ("SUM", 1), ("AVG", 1), ("WAVG", 2),
            ("MAX", 1), ("MIN", 1), ("COUNT", 1),
            ("SUMPRODUCT", 2), ("RATIO", 2),
        ):
            with self.subTest(name=name):
                spec = functions.get_function(name.lower())
                self.assertIs(spec, functions.FUNCTIONS[name])
                self.assertEqual(spec.name, name)
                self.assertIsInstance(spec.parameter_names, tuple)
                self.assertEqual(spec.arity, arity)
                self.assertEqual(len(spec.parameter_names), arity)
                if name == "SUM":
                    self.assertTrue(spec.variadic)
                    self.assertTrue(spec.allow_ranges)
                    self.assertEqual(spec.parameter_count, "至少 1")
                elif arity == 1:
                    self.assertEqual(spec.parameter_names, ("取数单元格",))
                self.assertIsNot(spec.factory(), spec.factory())
        self.assertIsNone(functions.get_function("NOT_REGISTERED"))

    def test_function_spec_signature_can_display_an_alias(self):
        spec = FunctionSpec(
            "PAIR", ("rate", "weight"), functions.get_function("WAVG").factory,
        )
        self.assertEqual(spec.arity, 2)
        self.assertEqual(spec.signature(), "PAIR(rate,weight)")
        self.assertEqual(spec.signature("F2_9"), "F2_9(rate,weight)")
        self.assertEqual(spec.name, "PAIR")
        self.assertEqual(spec.parameter_names, ("rate", "weight"))

    def test_alias_names_accept_positive_unpadded_identifiers(self):
        for name in ("FUN_1", "fun_999", "F1_1", "f2_9", "F12_345"):
            with self.subTest(name=name):
                self.assertTrue(functions.is_alias(name))
        for name in (
            "SUM", "F0_1", "F01_1", "F1_0", "F1_01", "F_1", "F1_-1",
            "F1_1X", " F1_1", "F1_1 ", "F1_1(A5)", "F1_1\n",
            "FUN_0", "FUN_01", "FUN_1X", "FUN_１", "FUN_1\n",
        ):
            with self.subTest(name=name):
                self.assertFalse(functions.is_alias(name))


class CalculatorTests(unittest.TestCase):
    def calculator(self, name):
        return functions.get_function(name).factory()

    def test_sum_uses_structural_decimal_inputs_without_sharing_state(self):
        first = self.calculator("SUM")
        second = self.calculator("SUM")
        for number in (_number("0.1"), _number("0.2"), _number(blank=True)):
            self.assertEqual(first.add((number,)), ())
        result = first.finish(3)
        self.assertIsInstance(result, FunctionResult)
        self.assertIsInstance(result.value, Decimal)
        self.assertEqual(result, FunctionResult(Decimal("0.3")))
        self.assertEqual(second.finish(0), FunctionResult(Decimal(0)))

    def test_average_counts_blank_inputs_in_the_file_denominator(self):
        calculator = self.calculator("AVG")
        for number in (_number(100), _number(blank=True), _number(200)):
            self.assertEqual(calculator.add((number,)), ())
        self.assertEqual(calculator.finish(3), FunctionResult(Decimal(100)))
        self.assertEqual(calculator.finish(4), FunctionResult(Decimal(75)))

    def test_average_rejects_zero_or_negative_source_count(self):
        calculator = self.calculator("AVG")
        calculator.add((_number(100),))
        for count in (0, -1):
            with self.subTest(source_count=count), self.assertRaises(ValueError):
                calculator.finish(count)

    def test_weighted_average_and_all_zero_weights_have_separate_states(self):
        calculator = self.calculator("WAVG")
        self.assertEqual(calculator.add((_number("0.03"), _number(100))), ())
        self.assertEqual(calculator.add((_number("0.05"), _number(900))), ())
        self.assertEqual(calculator.finish(2), FunctionResult(Decimal("0.048")))
        for blank in (False, True):
            with self.subTest(blank=blank):
                empty = self.calculator("WAVG")
                self.assertEqual(empty.add((_number(blank=blank), _number(blank=blank))), ())
                result = empty.finish(1)
                self.assertEqual(result.value, Decimal(0))
                self.assertEqual(len(result.notices), 1)
                notice = result.notices[0]
                self.assertIsInstance(notice, CalculationNotice)
                self.assertEqual((notice.severity, notice.code), ("error", "zero_weight"))
                self.assertIsNone(notice.parameter_index)
                self.assertTrue(notice.message)

    def test_blank_rate_notice_points_to_first_parameter_and_keeps_weight(self):
        calculator = self.calculator("WAVG")
        notices = calculator.add((_number(blank=True), _number(100)))
        self.assertIsInstance(notices, tuple)
        self.assertEqual(len(notices), 1)
        self.assertEqual(
            (notices[0].severity, notices[0].code, notices[0].parameter_index),
            ("warning", "blank_rate_with_weight", 0),
        )
        self.assertTrue(notices[0].message)
        self.assertEqual(calculator.add((_number("0.05"), _number(100))), ())
        self.assertEqual(calculator.finish(2), FunctionResult(Decimal("0.025")))
        zero_rate = self.calculator("WAVG")
        self.assertEqual(zero_rate.add((_number(0), _number(100))), ())
        self.assertEqual(zero_rate.finish(1), FunctionResult(Decimal(0)))

    def test_negative_weight_and_blank_rate_keep_distinct_parameter_indices(self):
        calculator = self.calculator("WAVG")
        notices = calculator.add((_number(blank=True), _number(-20)))
        self.assertIsInstance(notices, tuple)
        self.assertEqual(
            {(notice.code, notice.parameter_index) for notice in notices},
            {("blank_rate_with_weight", 0), ("negative_weight", 1)},
        )
        self.assertTrue(all(notice.severity == "warning" for notice in notices))
        self.assertEqual(calculator.add((_number("0.03"), _number(100))), ())
        self.assertEqual(calculator.finish(2), FunctionResult(Decimal("0.0375")))

    def test_cancelled_weights_return_a_result_level_error(self):
        calculator = self.calculator("WAVG")
        self.assertEqual(calculator.add((_number("0.03"), _number(100))), ())
        notices = calculator.add((_number("0.05"), _number(-100)))
        self.assertEqual(
            [(notice.code, notice.parameter_index) for notice in notices],
            [("negative_weight", 1)],
        )
        result = calculator.finish(2)
        self.assertEqual(result.value, Decimal(0))
        self.assertEqual(len(result.notices), 1)
        notice = result.notices[0]
        self.assertEqual((notice.severity, notice.code), ("error", "zero_weight"))
        self.assertIsNone(notice.parameter_index)


if __name__ == "__main__":
    unittest.main()
