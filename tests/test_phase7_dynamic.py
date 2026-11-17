import os
import sys
import shutil
import unittest
from pathlib import Path

# 加入根目錄至 sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from unittest.mock import MagicMock

from models.data_models import (
    RuleItem, SearchSettings, ConflictStrategy, MatchMode, RemoteFileTask, ReportRecord, TransferStatus, TaskStats
)
from core.traverse_engine import TraverseEngine
from core.download_manager import DownloadManager
from ui.rule_table_ctk import CTkRuleTable
import customtkinter as ctk


class DummySFTPAttr:
    def __init__(self, filename, is_dir=False, size=1024, mtime=1600000000):
        self.filename = filename
        self.st_mode = 0o040755 if is_dir else 0o100644
        self.st_size = size
        self.st_mtime = mtime


class TestPhase7Dynamic(unittest.TestCase):
    """驗證微階段 7 之動態升級功能：模型相容、多層目錄解析、精準/模糊模式、Manifest CSV 動態生成與 UI 結構"""

    def setUp(self):
        self.test_dir = Path("test_phase7_tmp")
        if self.test_dir.exists():
            shutil.rmtree(self.test_dir)
        self.test_dir.mkdir(parents=True, exist_ok=True)

    def tearDown(self):
        if self.test_dir.exists():
            shutil.rmtree(self.test_dir)

    def test_7_1_models_compatibility(self):
        """7.1 驗證 MatchMode 與 RuleItem 向下相容性"""
        self.assertEqual(MatchMode.FUZZY.value, "fuzzy")
        self.assertEqual(MatchMode.EXACT.value, "exact")

        # 舊式建構方式
        item_old = RuleItem("DIR1", "DIR2", "*.csv")
        self.assertEqual(item_old.l1_keyword, "DIR1")
        self.assertEqual(item_old.l2_keyword, "DIR2")
        self.assertEqual(item_old.file_pattern, "*.csv")
        self.assertEqual(item_old.dirs, ["DIR1", "DIR2"])

        # 新式動態建構方式 (N=3 目錄層)
        item_new = RuleItem(dirs=["A", "B", "C"], file_pattern="data.csv")
        self.assertEqual(item_new.dirs, ["A", "B", "C"])
        self.assertEqual(item_new.l1_keyword, "A")
        self.assertEqual(item_new.l2_keyword, "B")
        self.assertEqual(item_new.file_pattern, "data.csv")

        # 測試 setter
        item_new.l1_keyword = "A_MODIFIED"
        self.assertEqual(item_new.dirs[0], "A_MODIFIED")

        # 測試 is_completely_empty
        empty_item = RuleItem(dirs=["", "", ""], file_pattern="")
        self.assertTrue(empty_item.is_completely_empty())

    def test_7_2_match_pattern_exact_vs_fuzzy(self):
        """7.2 驗證模糊包含 vs 完全精準相符與萬用字元優先"""
        # 1. 萬用字元優先原則 (兩模式下皆適用)
        self.assertTrue(TraverseEngine.match_pattern("RAW_SAMPLE-2026.csv", "*RAW*-*.csv", MatchMode.FUZZY))
        self.assertTrue(TraverseEngine.match_pattern("RAW_SAMPLE-2026.csv", "*RAW*-*.csv", MatchMode.EXACT))
        self.assertFalse(TraverseEngine.match_pattern("SAMPLE-2026.csv", "*RAW*-*.csv", MatchMode.EXACT))

        # 2. 模糊包含模式 (FUZZY)
        self.assertTrue(TraverseEngine.match_pattern("FOLDER_PART_A", "PART", MatchMode.FUZZY))
        self.assertTrue(TraverseEngine.match_pattern("PART", "PART", MatchMode.FUZZY))
        self.assertTrue(TraverseEngine.match_pattern("folder_part_a", "PART", MatchMode.FUZZY)) # 大小寫不敏感

        # 3. 完全精準相符模式 (EXACT)
        self.assertTrue(TraverseEngine.match_pattern("PART", "PART", MatchMode.EXACT))
        self.assertTrue(TraverseEngine.match_pattern("part", "PART", MatchMode.EXACT)) # 大小寫不敏感
        self.assertFalse(TraverseEngine.match_pattern("FOLDER_PART_A", "PART", MatchMode.EXACT)) # 包含但非完全相同
        self.assertFalse(TraverseEngine.match_pattern("PART_A", "PART", MatchMode.EXACT))

        # 4. 留空代表全選
        self.assertTrue(TraverseEngine.match_pattern("ANY_DIRECTORY", "", MatchMode.EXACT))
        self.assertTrue(TraverseEngine.match_pattern("ANY_DIRECTORY", "   ", MatchMode.FUZZY))

    def test_7_3_dynamic_n_tier_scan(self):
        """7.3 驗證 TraverseEngine 支援任意 N 層目錄遍歷"""
        mock_sftp = MagicMock()

        # 模擬遠端樹狀結構 (3 層目錄結構: / -> D1 -> D2 -> D3 -> files)
        def mock_listdir_attr(path):
            norm_p = path.rstrip("/")
            if norm_p in ("", "/"):
                return [DummySFTPAttr("DEPPT_A", is_dir=True), DummySFTPAttr("IRRELEVANT", is_dir=True)]
            elif norm_p == "/DEPPT_A":
                return [DummySFTPAttr("SUB_B", is_dir=True), DummySFTPAttr("SUB_OTHER", is_dir=True)]
            elif norm_p == "/DEPPT_A/SUB_B":
                return [DummySFTPAttr("SEC_C", is_dir=True)]
            elif norm_p == "/DEPPT_A/SUB_B/SEC_C":
                return [
                    DummySFTPAttr("target.csv", is_dir=False, size=2048, mtime=1600000000),
                    DummySFTPAttr("other.txt", is_dir=False, size=512, mtime=1600000000)
                ]
            return []

        mock_sftp.listdir_attr.side_effect = mock_listdir_attr

        cancel_ev = MagicMock()
        cancel_ev.is_set.return_value = False

        search_cfg = SearchSettings(
            global_pattern="*.csv",
            conflict_strategy=ConflictStrategy.OVERWRITE,
            local_save_path=str(self.test_dir),
            match_mode=MatchMode.EXACT
        )

        engine = TraverseEngine(
            sftp_mgr=mock_sftp,
            initial_path="/",
            search_settings=search_cfg,
            cancel_event=cancel_ev,
            log_cb=lambda lvl, msg: None,
            stats_cb=lambda s: None
        )

        # 3 層目錄規則: DEPPT_A -> SUB_B -> SEC_C, 檔案留空套用全域 *.csv
        rules = [
            RuleItem(dirs=["DEPPT_A", "SUB_B", "SEC_C"], file_pattern="")
        ]
        stats = TaskStats()
        tasks = engine.scan_all(rules, stats)

        self.assertEqual(len(tasks), 1)
        self.assertEqual(tasks[0].filename, "target.csv")
        self.assertEqual(tasks[0].dirs, ["DEPPT_A", "SUB_B", "SEC_C"])
        self.assertEqual(tasks[0].remote_full_path, "/DEPPT_A/SUB_B/SEC_C/target.csv")
        self.assertEqual(tasks[0].local_full_path, os.path.join(str(self.test_dir), "DEPPT_A", "SUB_B", "SEC_C", "target.csv"))

    def test_7_4_parse_tsv_dynamic_columns(self):
        """7.4 驗證 Excel TSV 剪貼簿動態欄位數量解析（純目錄化）"""
        # 2 欄 (2 層純目錄)
        tsv_2cols = "DIR1\tSUBDIR1\nDIR2\tSUBDIR2"
        rules_2, dir_count_2 = CTkRuleTable.parse_tsv_text(tsv_2cols, dynamic=True)
        self.assertEqual(dir_count_2, 2)
        self.assertEqual(len(rules_2), 2)
        self.assertEqual(rules_2[0].dirs, ["DIR1", "SUBDIR1"])

        # 3 欄 (3 層純目錄)
        tsv_3cols = "L1\tL2\tL3"
        rules_3, dir_count_3 = CTkRuleTable.parse_tsv_text(tsv_3cols, dynamic=True)
        self.assertEqual(dir_count_3, 3)
        self.assertEqual(rules_3[0].dirs, ["L1", "L2", "L3"])

        # 4 欄 (4 層純目錄)
        tsv_4cols = "D1\tD2\tD3\tD4"
        rules_4, dir_count_4 = CTkRuleTable.parse_tsv_text(tsv_4cols, dynamic=True)
        self.assertEqual(dir_count_4, 4)
        self.assertEqual(rules_4[0].dirs, ["D1", "D2", "D3", "D4"])

    def test_7_5_generate_dynamic_manifest_csv(self):
        """7.5 驗證動態多階層 Manifest CSV 產出"""
        records = [
            ReportRecord(
                remote_path="/D1/D2/D3/file1.csv",
                local_path=str(self.test_dir / "D1" / "D2" / "D3" / "file1.csv"),
                dirs=["D1", "D2", "D3"],
                filename="file1.csv",
                file_size=1024,
                status=TransferStatus.SUCCESS,
                message="下載成功",
                timestamp="2026-11-17 20:00:00"
            )
        ]

        csv_path = DownloadManager.generate_manifest_csv(records, str(self.test_dir))
        self.assertIsNotNone(csv_path)
        self.assertTrue(os.path.exists(csv_path))

        with open(csv_path, "r", encoding="utf-8-sig") as f:
            lines = [line.strip() for line in f.readlines()]

        header = lines[0]
        self.assertIn("第 1 層目錄", header)
        self.assertIn("第 2 層目錄", header)
        self.assertIn("第 3 層目錄", header)
        self.assertIn("檔案大小 (易讀)", header)

        data_row = lines[1]
        self.assertIn("D1", data_row)
        self.assertIn("D2", data_row)
        self.assertIn("D3", data_row)
        self.assertIn("file1.csv", data_row)

    def test_7_6_main_window_elements(self):
        """7.6 驗證主視窗元件排版精簡與新增項目"""
        from ui.main_window import MainWindow

        win = MainWindow()
        try:
            # 驗證新元件存在
            self.assertTrue(hasattr(win, "ed_user"))
            self.assertTrue(hasattr(win, "ed_pass"))
            self.assertTrue(hasattr(win, "btn_toggle_pass"))
            self.assertTrue(hasattr(win, "btn_browse_remote"))
            self.assertTrue(hasattr(win, "cb_match_mode"))
            self.assertTrue(hasattr(win, "btn_open_manifest"))

            # 驗證精簡移除之按鈕不存在
            self.assertFalse(hasattr(win, "btn_add_row"))
            self.assertFalse(hasattr(win, "btn_del_row"))
        finally:
            win.destroy()


if __name__ == "__main__":
    unittest.main()
