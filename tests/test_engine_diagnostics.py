from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import Workbook

from summary_tool import functions
from summary_tool.engine import aggregate
from summary_tool.functions import CalculationNotice, FunctionResult, FunctionSpec
from summary_tool.models import AggregationError


class EngineDiagnosticsTests(unittest.TestCase):
    def test_final_error_notice_prevents_publishing(self):
        class FinalErrorCalculator:
            def add(self, values):
                return ()

            def finish(self, source_count):
                return FunctionResult(Decimal(1), (CalculationNotice("error", "final_check_failed", "最终校验失败"),))

        with TemporaryDirectory() as directory:
            root = Path(directory)
            template, source, output, config = (root / name for name in ("模板.xlsx", "源.xlsx", "结果.xlsx", "配置.xlsx"))
            workbook = Workbook()
            workbook.active["A1"] = "F1_9(A1)"
            workbook.save(template)
            workbook.active["A1"] = 2
            workbook.save(source)
            workbook.close()
            configuration = Workbook()
            configuration.active.title = "函数映射"
            configuration.active.append(["函数别名", "映射函数"])
            configuration.active.append(["F1_9", "TESTFINALERROR"])
            configuration.save(config)
            configuration.close()
            before = (template.read_bytes(), source.read_bytes())
            specification = FunctionSpec("TESTFINALERROR", ("取数单元格",), FinalErrorCalculator)
            with patch.dict(functions.FUNCTIONS, {"TESTFINALERROR": specification}):
                with self.assertRaises(AggregationError) as caught:
                    aggregate(template, [source], output, config_path=config)
            self.assertTrue(any(issue.code == "final_check_failed" for issue in caught.exception.issues))
            self.assertFalse(output.exists())
            self.assertEqual(before, (template.read_bytes(), source.read_bytes()))


if __name__ == "__main__":
    unittest.main()
