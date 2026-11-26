from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook, load_workbook

from summary_tool.models import AggregationError, AggregationResult, Issue
from summary_tool.reports import append_run_log, split_log_text


class LogCapacityTests(unittest.TestCase):
    def test_utf16_chunks_preserve_every_character(self):
        text = "开头" + "😀" * 20000 + "结尾"
        pieces = split_log_text(text)
        self.assertGreater(len(pieces), 1)
        self.assertEqual("".join(pieces), text)
        self.assertTrue(all(len(piece.encode("utf-16-le")) // 2 <= 32767 for piece in pieces))

    def test_long_source_and_message_round_trip_without_truncation(self):
        message = "开头" + "😀" * 20000 + "结尾"
        source = "来源" * 20000
        workbook = Workbook()
        issue = Issue("info", "result_source", message, source_file=source,
                      source_sheet="源表", source_cell="C5", template_sheet="业务明细", template_cell="A1")
        result = AggregationResult(Path("结果.xlsx"), 1, 1, issues=[issue], function_counts={"MAX": 1})
        append_run_log(workbook, Path("模板.xlsx"), [], datetime.now(), result)
        with TemporaryDirectory() as directory:
            path = Path(directory) / "长日志.xlsx"
            workbook.save(path)
            workbook.close()
            workbook = load_workbook(path)
            try:
                log = workbook["运行日志"]
                rows = [row for row in range(1, log.max_row + 1)
                        if "业务明细!A1" in str(log.cell(row, 4).value)]
                self.assertGreater(len(rows), 1)
                self.assertEqual("".join(log.cell(row, 5).value or "" for row in rows), source)
                self.assertEqual("".join(log.cell(row, 6).value or "" for row in rows), message)
                self.assertTrue(all((dimension.height or 0) <= 409.5 for dimension in log.row_dimensions.values()))
            finally:
                workbook.close()

    def test_full_existing_log_fails_without_destroying_history(self):
        workbook = Workbook()
        log = workbook.create_sheet("运行日志")
        log.cell(1048575, 2, "历史必须保留")
        result = AggregationResult(Path("结果.xlsx"), 1, 1)
        with self.assertRaises(AggregationError):
            append_run_log(workbook, Path("模板.xlsx"), [], datetime.now(), result)
        self.assertEqual(log.cell(1048575, 2).value, "历史必须保留")
        workbook.close()


if __name__ == "__main__":
    unittest.main()
