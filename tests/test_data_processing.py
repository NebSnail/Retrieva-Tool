import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from retrieval_tool import (
    build_query_indexes,
    dump_signed_type_mapping,
    get_icon_subsample_scale,
    get_candidate_indexes,
    get_query_headers,
    load_records,
    load_type_mapping_from_file_path,
    load_type_mapping_with_status,
    normalize_value,
    query_records,
)


class TestDataProcessing(unittest.TestCase):
    def test_get_query_headers(self):
        headers = ["序号", "编码", "直径", "长度", "所属项目"]
        self.assertEqual(get_query_headers(headers), ["直径", "长度"])

    def test_normalize_value_decimal_and_leading_zero(self):
        self.assertEqual(normalize_value("1200.50"), "1200.5")
        self.assertEqual(normalize_value("037332314020"), "037332314020")

    def test_query_with_index_candidates(self):
        records = [
            {"直径": "20", "长度": "1200", "编码": "A"},
            {"直径": "20", "长度": "1500", "编码": "B"},
            {"直径": "22", "长度": "1200", "编码": "C"},
        ]
        indexes = build_query_indexes(records, ["直径", "长度"])
        candidates = get_candidate_indexes(indexes, {"直径": "20", "长度": "1200"})
        results = query_records(records, {"直径": "20", "长度": "1200"}, candidate_indexes=candidates)
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0]["编码"], "A")

    def test_get_icon_subsample_scale_boundary(self):
        self.assertEqual(get_icon_subsample_scale(120, 120, 120), 1)
        self.assertEqual(get_icon_subsample_scale(121, 121, 120), 2)
        self.assertEqual(get_icon_subsample_scale(239, 239, 120), 2)
        self.assertEqual(get_icon_subsample_scale(240, 240, 120), 2)

    def test_load_records_preserves_leading_zero_by_cell_format(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            excel_path = Path(tmp_dir) / "sample.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.append(["序号", "编码", "直径", "长度", "所属项目"])
            ws.cell(row=2, column=1, value=1)
            code_cell = ws.cell(row=2, column=2, value=37332314020)
            code_cell.number_format = "000000000000"
            ws.cell(row=2, column=3, value=20)
            ws.cell(row=2, column=4, value=1200)
            ws.cell(row=2, column=5, value="项目A")
            wb.save(excel_path)
            wb.close()

            headers, records = load_records(excel_path)
            self.assertEqual(headers, ["序号", "编码", "直径", "长度", "所属项目"])
            self.assertEqual(records[0]["编码"], "037332314020")

    def test_config_loader_compat_and_fallback(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            config_path = Path(tmp_dir) / "types_config.json"

            # Old format compatibility
            config_path.write_text('{"默认": "C:/data/a.xlsx"}', encoding="utf-8")
            config = load_type_mapping_from_file_path(config_path)
            self.assertEqual(config["默认"]["excel"], "C:/data/a.xlsx")
            self.assertEqual(config["默认"]["icon"], "")

            # Corrupted config fallback
            config_path.write_text("{invalid_json", encoding="utf-8")
            config = load_type_mapping_from_file_path(config_path)
            self.assertEqual(config, {})

    def test_config_loader_rejects_partial_signed_fields(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            config_path = Path(tmp_dir) / "types_config.json"
            config_path.write_text('{"_data": {"默认": {"excel": "C:/a.xlsx", "icon": ""}}}', encoding="utf-8")
            config = load_type_mapping_from_file_path(config_path)
            self.assertEqual(config, {})

    def test_config_loader_accepts_signed_payload(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            config_path = Path(tmp_dir) / "types_config.json"
            mapping = {"默认": {"excel": "C:/a.xlsx", "icon": ""}}
            config_path.write_text(dump_signed_type_mapping(mapping), encoding="utf-8")
            config = load_type_mapping_from_file_path(config_path)
            self.assertEqual(config, mapping)

    def test_config_loader_rejects_unsigned_new_dict_format(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            config_path = Path(tmp_dir) / "types_config.json"
            config_path.write_text('{"默认": {"excel": "C:/a.xlsx", "icon": ""}}', encoding="utf-8")
            config = load_type_mapping_from_file_path(config_path)
            self.assertEqual(config, {})

    def test_config_loader_status_contains_warning_for_tampered_config(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            config_path = Path(tmp_dir) / "types_config.json"
            config_path.write_text('{"_data": {"默认": {"excel": "C:/a.xlsx", "icon": ""}}}', encoding="utf-8")
            mapping, warning = load_type_mapping_with_status(config_path)
            self.assertEqual(mapping, {})
            self.assertTrue(warning)
            self.assertIn("手动修改", warning)


if __name__ == "__main__":
    unittest.main()
