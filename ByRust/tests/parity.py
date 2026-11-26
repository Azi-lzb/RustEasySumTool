"""Run the existing workbook contract tests against the Rust executable.

Each adapter invocation also runs the untouched Python reference and compares
business values, statistics and structured diagnostic locations. Fixtures live
only in TemporaryDirectory. Tests that patch Python internals are covered by
native Rust API tests instead and are listed explicitly below.
"""
from __future__ import annotations
import importlib
import json
import math
import os
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from openpyxl import load_workbook
from summary_tool import configuration
from summary_tool.engine import aggregate as python_aggregate
from summary_tool.models import AggregationError, AggregationResult, Issue

EXE = Path(os.environ.get("RUST_SUMMARY_EXE", ROOT / "ByRust/target/debug/excel-summary-rust.exe"))
CALLS = 0

def values(path):
    workbook = load_workbook(path)
    try:
        return {sheet.title: {cell.coordinate: cell.value for cell in sheet._cells.values()
                              if cell.value is not None} for sheet in workbook if sheet.title != "运行日志"}
    finally:
        workbook.close()

def diagnostic_keys(issues):
    return sorted((issue.severity, issue.code, issue.source_file, issue.source_sheet,
                   issue.source_cell, issue.template_sheet, issue.template_cell) for issue in issues)

def aggregate(template, sources, output, progress=None, cancel_event=None, *, config_path=None):
    global CALLS
    assert progress is None and cancel_event is None, "Use native Rust API tests for callbacks/cancellation"
    sources = list(sources)
    if config_path is None:
        default = configuration.application_root() / "config.xlsx"
        if default.exists():
            config_path = default
    command = [str(EXE), "--template", str(template), "--output", str(output), "--json", "--sources", *map(str, sources)]
    if config_path is not None:
        command += ["--config", str(config_path)]
    # Existing output/input collisions must be checked using their real paths.
    collision = Path(output).exists() or Path(output).resolve() in {Path(path).resolve() for path in [template, *sources]}
    with TemporaryDirectory(prefix="rust_python_reference_") as directory:
        reference = Path(output) if collision else Path(directory) / Path(output).name
        try:
            expected = python_aggregate(template, sources, reference, config_path=config_path)
            expected_error = None
        except AggregationError as error:
            expected, expected_error = None, error
        completed = subprocess.run(command, cwd=ROOT, capture_output=True, text=True, encoding="utf-8", timeout=90)
        try:
            payload = json.loads(completed.stdout)
        except Exception as error:
            raise AssertionError(f"Rust did not emit JSON\n{completed.stdout}\n{completed.stderr}") from error
        CALLS += 1
        if not payload["ok"]:
            error = payload["error"]
            actual_error = AggregationError(error["message"], [Issue(**issue) for issue in error["issues"]])
            assert expected_error is not None, f"Rust failed; Python succeeded: {error}"
            assert diagnostic_keys(actual_error.issues) == diagnostic_keys(expected_error.issues), (actual_error.issues, expected_error.issues)
            raise actual_error
        assert expected_error is None, f"Rust succeeded; Python failed: {expected_error}"
        actual = payload["result"]
        actual["output_path"] = Path(actual["output_path"])
        actual["config_path"] = Path(actual["config_path"]) if actual["config_path"] else None
        actual["issues"] = [Issue(**issue) for issue in actual["issues"]]
        result = AggregationResult(**actual)
        for field in ["source_count", "source_sheet_count", "template_sheet_count", "rule_count",
                      "sum_rule_count", "avg_rule_count", "wavg_rule_count", "function_counts", "function_bindings"]:
            assert getattr(result, field) == getattr(expected, field), (field, result, expected)
        assert diagnostic_keys(result.issues) == diagnostic_keys(expected.issues), (result.issues, expected.issues)
        left, right = values(reference), values(output)
        assert left.keys() == right.keys(), (left.keys(), right.keys())
        for sheet in left:
            assert left[sheet].keys() == right[sheet].keys(), (sheet, left[sheet], right[sheet])
            for cell, value in left[sheet].items():
                other = right[sheet][cell]
                if isinstance(value, (float, int)) and not isinstance(value, bool):
                    assert isinstance(other, (float, int)) and math.isclose(value, other, rel_tol=2e-15, abs_tol=0), (sheet, cell, value, other)
                else:
                    assert value == other, (sheet, cell, value, other)
        return result

CLASSES = {
    "test_core": ["FileTests"], "test_avg": ["AvgFileTests"],
    "test_aliases": ["AliasFileTests"], "test_demo": ["DemoAggregationTests"],
    "test_parameter_parsing": ["ParameterFileTests"], "test_fullwidth_rules": ["FullwidthFileTests"],
    "test_product_ratio": ["ProductRatioFileTests"], "test_template_safety": ["TemplateSafetyTests"],
    "test_run_log": ["RunLogTests"],
}
NATIVE = {
    "test_cancel_does_not_create_output", "test_file_appearing_during_save_is_not_overwritten",
    "test_config_changed_during_progress_applies_only_to_the_next_batch",
    "test_runtime_registry_extension_needs_no_parser_or_engine_branch",
    "test_configuration_accepts_both_alias_formats_without_arity_constraints",
}

if __name__ == "__main__":
    suite = unittest.TestSuite()
    for filename, names in CLASSES.items():
        module = importlib.import_module(f"tests.{filename}")
        module.aggregate = aggregate
        for name in names:
            for test in unittest.defaultTestLoader.loadTestsFromTestCase(getattr(module, name)):
                if test._testMethodName not in NATIVE:
                    suite.addTest(test)
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    print(f"Cross-language aggregation calls: {CALLS}")
    print(f"Native API tests replace Python-internal cases: {sorted(NATIVE)}")
    sys.exit(not result.wasSuccessful())
