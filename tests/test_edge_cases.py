import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from query_tool import load_records, query_records


class TestEdgeCases(unittest.TestCase):
    def test_query_records_empty_conditions(self):
        records = [{"编码": "A"}]
        self.assertEqual(query_records(records, {}), [])

    def test_load_records_rejects_missing_required_headers(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            excel_path = Path(tmp_dir) / "missing_required.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.append(["序号", "编码", "直径", "长度", "项目"])
            wb.save(excel_path)
            wb.close()

            with self.assertRaises(ValueError) as ctx:
                load_records(excel_path)
            self.assertIn("缺少必需列", str(ctx.exception))

    def test_load_records_rejects_header_with_gap(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            excel_path = Path(tmp_dir) / "header_gap.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.append(["序号", "编码", "", "长度", "所属项目"])
            wb.save(excel_path)
            wb.close()

            with self.assertRaises(ValueError) as ctx:
                load_records(excel_path)
            self.assertIn("标题行不完整", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()
