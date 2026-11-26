from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook, load_workbook

from summary_tool.engine import aggregate


class RunLogTests(unittest.TestCase):
    def test_empty_merged_log_layout_is_not_overwritten(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            template, source, output = (root / name for name in ("模板.xlsx", "源.xlsx", "结果.xlsx"))
            workbook = Workbook()
            workbook.active["A1"] = "SUM(A1)"
            workbook.create_sheet("运行日志").merge_cells("B2:F4")
            workbook.save(template)
            workbook.remove(workbook["运行日志"])
            workbook.active["A1"] = 5
            workbook.save(source)
            workbook.close()
            aggregate(template, [source], output)
            workbook = load_workbook(output)
            try:
                self.assertEqual(workbook.sheetnames, ["Sheet", "运行日志"])
                log = workbook["运行日志"]
                self.assertIn("B2:F4", log.merged_cells)
                self.assertIsNone(log["B2"].value)
                self.assertEqual(log["B7"].value, "本次运行日志")
            finally:
                workbook.close()

    def test_source_filename_is_logged_as_text_not_a_formula(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            template, source, output = (root / name for name in ("模板.xlsx", "=1+1.xlsx", "结果.xlsx"))
            workbook = Workbook()
            workbook.active["A1"] = "SUM(A1)"
            workbook.save(template)
            workbook.active["A1"] = 3
            workbook.save(source)
            workbook.close()
            aggregate(template, [source], output)
            workbook = load_workbook(output)
            try:
                matches = [cell for row in workbook["运行日志"] for cell in row if cell.value == "=1+1.xlsx"]
                self.assertTrue(matches)
                self.assertTrue(all(cell.data_type == "s" for cell in matches))
                self.assertFalse(any(cell.data_type == "f" for sheet in workbook for row in sheet for cell in row))
            finally:
                workbook.close()


if __name__ == "__main__":
    unittest.main()
