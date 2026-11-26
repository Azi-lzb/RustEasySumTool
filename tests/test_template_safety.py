from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from openpyxl import Workbook
from openpyxl.comments import Comment

from summary_tool.engine import AggregationError, aggregate


class TemplateSafetyTests(unittest.TestCase):
    def test_unsupported_template_objects_fail_instead_of_silently_disappearing(self):
        for part in ("xl/media/image1.png", "xl/charts/chart1.xml", "xl/drawings/drawing1.xml", "xl/activeX/activeX1.xml", "_xmlsignatures/sig1.xml"):
            with self.subTest(part=part), TemporaryDirectory() as directory:
                root = Path(directory)
                template = root / "模板.xlsx"
                source = root / "源.xlsx"
                result = root / "结果.xlsx"
                workbook = Workbook()
                workbook.active.title = "明细"
                workbook.active["A1"] = "SUM(A1)"
                workbook.save(template)
                workbook.active["A1"] = 5
                workbook.save(source)
                workbook.close()
                with ZipFile(template, "a") as archive:
                    archive.writestr(part, b"unsupported test object")
                before = template.read_bytes()
                with self.assertRaisesRegex(AggregationError, "当前版本仅支持单元格表样"):
                    aggregate(template, [source], result)
                self.assertFalse(result.exists())
                self.assertEqual(before, template.read_bytes())

    def test_empty_drawing_descriptor_is_not_a_real_object(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            workbook = Workbook()
            workbook.active["A1"] = "SUM(A1)"
            template, source, output = (root / name for name in ("模板.xlsx", "源.xlsx", "结果.xlsx"))
            workbook.save(template)
            workbook.active["A1"] = 7
            workbook.save(source)
            workbook.close()
            with ZipFile(template, "a") as archive:
                archive.writestr("xl/drawings/drawing1.xml", '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>')
            before = template.read_bytes()
            aggregate(template, [source], output)
            self.assertEqual(template.read_bytes(), before)
            self.assertTrue(output.is_file())

    def test_drawing_descriptor_with_objects_is_still_rejected(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            workbook = Workbook()
            workbook.active["A1"] = "SUM(A1)"
            template, source, output = (root / name for name in ("模板.xlsx", "源.xlsx", "结果.xlsx"))
            workbook.save(template)
            workbook.active["A1"] = 7
            workbook.save(source)
            workbook.close()
            with ZipFile(template, "a") as archive:
                archive.writestr("xl/drawings/drawing1.xml", '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:twoCellAnchor/></xdr:wsDr>')
            with self.assertRaises(AggregationError):
                aggregate(template, [source], output)
            self.assertFalse(output.exists())

    def test_ordinary_cell_comments_are_preserved(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            workbook = Workbook()
            workbook.active["A1"] = "SUM(A1)"
            workbook.active["A1"].comment = Comment("需要保留的批注", "测试")
            template = root / "模板.xlsx"
            source = root / "源.xlsx"
            workbook.save(template)
            workbook.active["A1"] = 10
            workbook.save(source)
            workbook.close()
            aggregate(template, [source], root / "结果.xlsx")
            from openpyxl import load_workbook
            result = load_workbook(root / "结果.xlsx")
            try:
                self.assertEqual(result.active["A1"].comment.text, "需要保留的批注")
            finally:
                result.close()


if __name__ == "__main__":
    unittest.main()
