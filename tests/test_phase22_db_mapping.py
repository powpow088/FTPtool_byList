# tests/test_phase22_db_mapping.py
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

import openpyxl
import xlsxwriter

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.db_config import DB_CONFIG, is_db_configured
from core.db_connector import DatabaseConnector, parse_coord_int
from core.detail_aggregator import DetailAggregator
from core.pivot_aggregator import PivotAggregator
from core.sidecar_cache import SidecarCacheWriter


class TestPhase22DbMapping(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp_dir_obj = tempfile.TemporaryDirectory(prefix="wafer_db_test_")
        self.tmp_dir = Path(self.tmp_dir_obj.name)

    def tearDown(self) -> None:
        self.tmp_dir_obj.cleanup()

    def test_parse_coord_int(self) -> None:
        """驗證座標清洗轉純整數邏輯"""
        self.assertEqual(parse_coord_int("01"), 1)
        self.assertEqual(parse_coord_int("1"), 1)
        self.assertEqual(parse_coord_int(1), 1)
        self.assertEqual(parse_coord_int(1.0), 1)
        self.assertEqual(parse_coord_int(" 05 "), 5)
        self.assertEqual(parse_coord_int(""), 0)
        self.assertEqual(parse_coord_int(None), 0)
        self.assertEqual(parse_coord_int("INVALID"), 0)

    def test_db_config_unconfigured_fallback(self) -> None:
        """驗證資料庫未配置時自動降級"""
        with patch.dict(DB_CONFIG, {"server": "YOUR_SERVER_IP", "user": "YOUR_USERNAME", "enabled": True}):
            self.assertFalse(is_db_configured())
            res = DatabaseConnector.fetch_wafer_unit_base(["LOT01"])
            self.assertEqual(res, {})

        with patch.dict(DB_CONFIG, {"server": "127.0.0.1", "user": "sa", "enabled": False}):
            self.assertFalse(is_db_configured())

    def test_mock_db_batch_mapping_and_dedup(self) -> None:
        """驗證資料庫批次查詢、座標 int 化、以及多值有序去重逗號串接"""
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_conn.cursor.return_value = mock_cursor

        # 模擬 DB 查詢回傳結果 (LotNum, MPID, MP_X, MP_Y, PanelID, Scrapped)
        # MPID_01: 有兩筆資料 (不同的 PanelID 與 Scrapped 0, 1)
        # MPID_02: 單一資料
        mock_cursor.fetchall.return_value = [
            ("23BKE002-06-00", "MPID_01", 1.0, 1.0, "PANEL_A", 0),
            ("23BKE002-06-00", "MPID_01", "01", "1", "PANEL_B", 1),
            ("23BKE002-06-00", "MPID_02", 1, 2, "PANEL_C", 0),
        ]

        with patch.dict(DB_CONFIG, {"server": "10.0.0.1", "user": "test_user", "enabled": True}):
            with patch("pymssql.connect", return_value=mock_conn):
                result = DatabaseConnector.fetch_wafer_unit_base(["23BKE002-06-00"])

        self.assertIn(("23BKE002-06-00", "MPID_01", 1, 1), result)
        panels, scrapped = result[("23BKE002-06-00", "MPID_01", 1, 1)]
        self.assertEqual(panels, "PANEL_A,PANEL_B")
        self.assertEqual(scrapped, "0,1")

        self.assertIn(("23BKE002-06-00", "MPID_02", 1, 2), result)
        panels2, scrapped2 = result[("23BKE002-06-00", "MPID_02", 1, 2)]
        self.assertEqual(panels2, "PANEL_C")
        self.assertEqual(scrapped2, "0")

    def test_detail_aggregator_integration_with_db_mapping(self) -> None:
        """驗證 DetailAggregator 在第一彙整時批次查詢 DB 並注入 PanelID, Scrapped 至工作表2"""
        raw_file = self.tmp_dir / "raw_db_wafer.csv"
        raw_file.write_text(
            "2D ID,X,Y,Unit BIN#\n"
            "MPID_10,1,2,1\n"
            "MPID_20,3,4,32\n",
            encoding="utf-8-sig",
        )

        manifest_csv = self.tmp_dir / "manifest_db.csv"
        manifest_csv.write_text(
            "第 1 層目錄,第 1 層搜尋字,第 2 層目錄,第 2 層搜尋字,處理狀態,本地完整路徑\n"
            f"PKG-SITE-24W,PKG,23BKE002-06-000,23BKE002-06-00,下載成功,{raw_file}\n",
            encoding="utf-8-sig",
        )

        out_xlsx = self.tmp_dir / "detail_out_with_db.xlsx"
        mock_db_map = {
            ("23BKE002-06-00", "MPID_10", 1, 2): ("PNL_10", "0"),
            ("23BKE002-06-00", "MPID_20", 3, 4): ("PNL_20A,PNL_20B", "1"),
        }

        with patch.object(DatabaseConnector, "fetch_wafer_unit_base", return_value=mock_db_map):
            succeeded, failed, total_written = DetailAggregator.aggregate_to_excel(
                str(manifest_csv), str(out_xlsx)
            )

        self.assertEqual(succeeded, 1)
        self.assertEqual(total_written, 2)

        wb_check = openpyxl.load_workbook(str(out_xlsx))
        ws2 = wb_check["工作表2"]
        self.assertEqual(ws2["K1"].value, "PanelID")
        self.assertEqual(ws2["L1"].value, "Scrapped")
        self.assertEqual(ws2["M1"].value, "Unit BIN#")

        self.assertEqual(ws2["K2"].value, "PNL_10")
        self.assertEqual(ws2["L2"].value, "0")
        self.assertEqual(ws2["M2"].value, 1)

        self.assertEqual(ws2["K3"].value, "PNL_20A,PNL_20B")
        self.assertEqual(ws2["L3"].value, "1")
        self.assertEqual(ws2["M3"].value, 32)
        wb_check.close()

    def test_pivot_aggregator_integration_with_panelid_scrapped(self) -> None:
        """核心整合驗證：PivotAggregator 直接消費來源明細/快取的 PanelID 與 Scrapped，免二次連線 DB"""
        src_xlsx = self.tmp_dir / "src_test_db.xlsx"
        writer = SidecarCacheWriter(src_xlsx)

        # 晶片 1: MPID_01 (X=1, Y=2, PanelID=PANEL_X1,PANEL_X2, Scrapped=0,1)
        # 晶片 2: MPID_02 (X=3, Y=4, PanelID=PANEL_Y1, Scrapped=0)
        rows = [
            (
                "214-3273242A-24W", "3273242", "23BKE002-06-000", "23BKE002-06-00", "24W",
                "f1.csv", "MPID_01", "1", "2", "MPID_01_1_2", "PANEL_X1,PANEL_X2", "0,1", "1"
            ),
            (
                "214-3273242A-SUB-C", "3273242", "23BKE002-06-000", "23BKE002-06-00", "SUB-C",
                "f2.csv", "MPID_02", "3", "4", "MPID_02_3_4", "PANEL_Y1", "0", "32"
            ),
        ]
        src_xlsx.write_bytes(b"MOCK")
        writer.write_batch(rows)
        writer.finalize(src_xlsx)

        pivot_out = self.tmp_dir / "pivot_out_db.xlsx"

        src_rows, det_rows, sum_rows, _ = PivotAggregator.aggregate_pivot(
            str(src_xlsx), str(pivot_out)
        )

        self.assertEqual(src_rows, 2)
        self.assertEqual(sum_rows, 2)

        wb_check = openpyxl.load_workbook(str(pivot_out))

        # 1. 驗證「中繼明細_01」結構未被改動 (不包含 PanelID, Scrapped)
        ws_det = wb_check["中繼明細_01"]
        det_headers = [cell.value for cell in ws_det[1]]
        self.assertEqual(det_headers[:4], ["PN", "第 2 層目錄(LOT)", "LOT", "MPID_XY"])
        self.assertNotIn("PanelID", det_headers)
        self.assertNotIn("Scrapped", det_headers)

        # 2. 驗證「RECIPE彙整_01」結構精確包含 PanelID 與 Scrapped
        ws_sum = wb_check["RECIPE彙整_01"]
        sum_headers = [cell.value for cell in ws_sum[1]]
        expected_fixed = [
            "PN", "第 2 層目錄(LOT)", "LOT", "MPID_XY", "PanelID", "Scrapped", "BinCode_sumary"
        ]
        self.assertEqual(sum_headers[:7], expected_fixed)

        # 驗證資料列 1 (MPID_01) 的 PanelID 與 Scrapped 成功注入
        self.assertEqual(ws_sum["D2"].value, "MPID_01_01_02")
        self.assertEqual(ws_sum["E2"].value, "PANEL_X1,PANEL_X2")
        self.assertEqual(ws_sum["F2"].value, "0,1")
        self.assertEqual(ws_sum["G2"].value, 1)  # BinCode_sumary

        # 驗證資料列 2 (MPID_02) 的 PanelID 與 Scrapped 成功注入
        self.assertEqual(ws_sum["D3"].value, "MPID_02_03_04")
        self.assertEqual(ws_sum["E3"].value, "PANEL_Y1")
        self.assertEqual(ws_sum["F3"].value, "0")
        self.assertEqual(ws_sum["G3"].value, 32)  # BinCode_sumary

        wb_check.close()


if __name__ == "__main__":
    unittest.main(verbosity=2)
