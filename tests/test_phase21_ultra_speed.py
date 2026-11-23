# tests/test_phase21_ultra_speed.py
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl
import xlsxwriter

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.detail_aggregator import DetailAggregator
from core.excel_utils import safe_atomic_move, sanitize_xml_text
from core.pivot_aggregator import (
    PivotAggregator,
    pad_coordinates,
    sanitize_bin_value,
)
from core.sidecar_cache import SidecarCacheReader, SidecarCacheWriter, get_cache_path


class TestPhase21UltraSpeed(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp_dir_obj = tempfile.TemporaryDirectory(prefix="wafer_test_")
        self.tmp_dir = Path(self.tmp_dir_obj.name)

    def tearDown(self) -> None:
        self.tmp_dir_obj.cleanup()

    def test_sidecar_cache_lifecycle(self) -> None:
        xlsx_path = self.tmp_dir / "mock_detail.xlsx"
        xlsx_path.write_bytes(b"DUMMY_XLSX_HEADER_CONTENT")

        writer = SidecarCacheWriter(xlsx_path)
        batch = [
            (
                "L1", "PN01", "L2_A", "LOT01", f"REC_{i % 5}",
                "file.csv", f"ID{i}", "1", "2", f"ID{i}_1_2", "1"
            )
            for i in range(10_000)
        ]
        writer.write_batch(batch)
        cache_path = writer.finalize(xlsx_path)

        self.assertTrue(cache_path.exists())
        self.assertTrue(SidecarCacheReader.is_valid(xlsx_path))
        recipes = SidecarCacheReader.get_distinct_recipes(cache_path)
        self.assertEqual(recipes, ["REC_0", "REC_1", "REC_2", "REC_3", "REC_4"])

        # 變更檔案大小後判定無效
        xlsx_path.write_bytes(b"DUMMY_XLSX_HEADER_CONTENT_MODIFIED")
        self.assertFalse(SidecarCacheReader.is_valid(xlsx_path))

    def test_calamine_fallback_and_type_sanitization(self) -> None:
        self.assertEqual(sanitize_bin_value(32), "32")
        self.assertEqual(sanitize_bin_value(1.0), "1")
        self.assertEqual(sanitize_bin_value("01"), "01")
        self.assertEqual(sanitize_bin_value(None), "")
        self.assertEqual(pad_coordinates("ID_1_2"), "ID_01_02")

        ext_xlsx = self.tmp_dir / "external_from_other_pc.xlsx"
        wb = xlsxwriter.Workbook(str(ext_xlsx))
        ws = wb.add_worksheet("工作表2")
        ws.write_row(0, 0, DetailAggregator.DETAIL_HEADERS)
        ws.write_row(1, 0, ["L1", "PN1", "L2A", "LOT1", "CP1", "f1.csv", "W1", 1, 2, "W1_1_2", 1.0])
        ws.write_row(2, 0, ["L1", "PN1", "L2A", "LOT1", "CP2", "f2.csv", "W1", 1, 2, "W1_1_2", 32])
        ws.write_row(3, 0, ["L1", "PN1", "L2A", "LOT1", "CP3", "f3.csv", "W1", 1, 2, "W1_1_2", "01"])
        ws.write_row(4, 0, ["L1", "PN1", "L2A", "LOT1", "CP4", "f4.csv", "W1", 1, 2, "W1_1_2", None])
        wb.close()

        self.assertFalse(get_cache_path(ext_xlsx).exists())

        pivot_out = self.tmp_dir / "pivot_from_calamine.xlsx"
        src_rows, det_rows, sum_rows, _ = PivotAggregator.aggregate_pivot(str(ext_xlsx), str(pivot_out))
        self.assertEqual(src_rows, 4)

        wb_check = openpyxl.load_workbook(str(pivot_out), data_only=True)
        ws_detail = wb_check["中繼明細_01"]
        self.assertEqual(ws_detail["D2"].value, "W1_01_02")
        self.assertEqual(ws_detail["E2"].value, "1")
        self.assertEqual(ws_detail["F2"].value, "32")
        self.assertEqual(ws_detail["G2"].value, "01")
        wb_check.close()

    def test_xml_sanitizer_illegal_chars(self) -> None:
        dirty = "WAFER\x00_ID\x07_TEST\x0b\x1f\tLINE1\nLINE2\r"
        cleaned = sanitize_xml_text(dirty)
        self.assertEqual(cleaned, "WAFER_ID_TEST\tLINE1\nLINE2\r")
        self.assertNotIn("\x00", cleaned)
        self.assertNotIn("\x07", cleaned)
        self.assertNotIn("\x0b", cleaned)

    def test_detail_aggregator_dual_sink_and_sheets(self) -> None:
        raw_file = self.tmp_dir / "wafer_raw_01.tsv"
        raw_file.write_text(
            "2D ID\tX\tY\tUnit BIN#\n"
            "CHIP01\t1\t2\t1\n"
            "CHIP02\t3\t4\t32\n",
            encoding="utf-8",
        )

        list_csv = self.tmp_dir / "下載清單.csv"
        list_csv.write_text(
            "第 1 層目錄,第 1 層搜尋字,第 2 層目錄,第 2 層搜尋字,處理狀態,本地完整路徑\n"
            f"AREA1-SUB-C,AREA1,SUB1,LOT88,下載成功,{raw_file}\n",
            encoding="utf-8-sig",
        )

        out_xlsx = self.tmp_dir / "下載與明細彙整_test.xlsx"
        succeeded, failed, total_written = DetailAggregator.aggregate_to_excel(str(list_csv), str(out_xlsx))

        self.assertEqual(succeeded, 1)
        self.assertEqual(total_written, 2)
        self.assertTrue(out_xlsx.exists())
        self.assertTrue(SidecarCacheReader.is_valid(out_xlsx))

        wb_check = openpyxl.load_workbook(str(out_xlsx))
        self.assertEqual(wb_check.sheetnames, ["下載清單", "工作表2"])
        ws2 = wb_check["工作表2"]
        self.assertEqual(ws2["A1"].value, "第 1 層目錄")
        self.assertEqual(ws2["A2"].number_format, "@")
        self.assertEqual(ws2["J2"].value, "CHIP01_1_2")
        self.assertEqual(ws2["K2"].value, 1)
        wb_check.close()

    def test_pivot_aggregator_mother_lot_merging_and_styles(self) -> None:
        """核心驗證：同一母批號下的兩片晶圓，在 RECIPE 彙整中必須成功合併為同一列！"""
        src_xlsx = self.tmp_dir / "src_for_pivot.xlsx"
        writer = SidecarCacheWriter(src_xlsx)

        # 晶片 A 在兩片不同晶圓 (25A7E004-01-000 與 25A7E004-01-002)，母批號皆為 25A7E004-01-00
        rows = [
            ("214-3273242A-24W", "3273242", "25A7E004-01-000", "25A7E004-01-00", "24W", "f1.csv", "CHIP_A", "1", "2", "CHIP_A_1_2", "1"),
            ("214-3273242A-SUB-C", "3273242", "25A7E004-01-002", "25A7E004-01-00", "SUB-C", "f2.csv", "CHIP_A", "01", "02", "CHIP_A_01_02", "32"),
            ("3273173A-NOREC", "3273173", "25B12001-01-00", "25B12001-01-00", "", "f3.csv", "CHIP_B", "5", "5", "CHIP_B_5_5", "32"),
        ]
        src_xlsx.write_bytes(b"MOCK_SRC")
        writer.write_batch(rows)
        writer.finalize(src_xlsx)

        pivot_out = self.tmp_dir / "pivot_paginated.xlsx"
        src_rows, det_rows, sum_rows, _ = PivotAggregator.aggregate_pivot(
            str(src_xlsx), str(pivot_out)
        )

        self.assertEqual(src_rows, 3)
        self.assertEqual(det_rows, 3)
        # 晶片 A 在兩片晶圓合併為 1 列，加上晶片 B -> 總共 2 列！
        self.assertEqual(sum_rows, 2)

        wb_check = openpyxl.load_workbook(str(pivot_out))
        self.assertIn("中繼明細_01", wb_check.sheetnames)
        self.assertIn("RECIPE彙整_01", wb_check.sheetnames)
        self.assertEqual(wb_check.active.title, "RECIPE彙整_01")

        ws_sum = wb_check["RECIPE彙整_01"]
        # 驗證 B 欄成功以逗號串接兩個 sub-lot！
        self.assertEqual(ws_sum["A2"].value, "3273242")
        self.assertIn("25A7E004-01-000", ws_sum["B2"].value)
        self.assertIn("25A7E004-01-002", ws_sum["B2"].value)
        self.assertEqual(ws_sum["C2"].value, "25A7E004-01-00")
        self.assertEqual(ws_sum["D2"].value, "CHIP_A_01_02")
        # 跨晶圓最高權重比較：24W=1, SUB-C=32 -> 綜合判定必須為 32！
        self.assertEqual(ws_sum["E2"].value, 32)

        # 驗證 ERROR_blank 欄位排除判定 (Row 3 只有空 RECIPE -> E3 為空)
        self.assertIn(ws_sum["E3"].value, ("", None))

        # 驗證表頭樣式深藍底 (#1F4E78) 與 ERROR_blank 深橘底 (#9C4A00 / #C65911)
        self.assertTrue(ws_sum["A1"].fill.fgColor.rgb.endswith("1F4E78"))
        self.assertTrue(
            ws_sum["H1"].fill.fgColor.rgb.endswith("9C4A00")
            or ws_sum["H1"].fill.fgColor.rgb.endswith("C65911")
        )
        wb_check.close()

    def test_safe_atomic_move_retry(self) -> None:
        src = self.tmp_dir / "temp_output.xlsx"
        dst = self.tmp_dir / "final_output.xlsx"
        src.write_bytes(b"EXCEL_BINARY_PAYLOAD")

        call_counter = {"count": 0}
        real_move = os.replace

        def flaky_move(s: str, d: str) -> None:
            call_counter["count"] += 1
            if call_counter["count"] <= 2:
                raise PermissionError("[WinError 32] 檔案正由防毒軟體掃描中")
            real_move(s, d)

        with patch("shutil.move", side_effect=flaky_move):
            safe_atomic_move(str(src), str(dst), max_retries=5)

        self.assertEqual(call_counter["count"], 3)
        self.assertTrue(dst.exists())
        self.assertFalse(src.exists())


if __name__ == "__main__":
    unittest.main(verbosity=2)
