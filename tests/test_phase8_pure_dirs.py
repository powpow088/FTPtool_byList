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


class DummySFTPAttr:
    def __init__(self, filename, is_dir=False, size=1024, mtime=1600000000):
        self.filename = filename
        self.st_mode = 0o040755 if is_dir else 0o100644
        self.st_size = size
        self.st_mtime = mtime


class TestPhase8PureDirs(unittest.TestCase):
    """驗證微階段 8：純目錄動態階層表格與最後目錄檔案搜尋"""

    def setUp(self):
        self.test_dir = Path("test_phase8_tmp")
        if self.test_dir.exists():
            shutil.rmtree(self.test_dir)
        self.test_dir.mkdir(parents=True, exist_ok=True)

    def tearDown(self):
        if self.test_dir.exists():
            shutil.rmtree(self.test_dir)

    def test_8_1_tsv_parsing_pure_dirs(self):
        """8.1 驗證 Excel TSV 貼上之欄位 100% 全為目錄層級"""
        # 1 欄 (1 層目錄)
        tsv_1col = "PART_A\nPART_B"
        rules_1, dir_count_1 = CTkRuleTable.parse_tsv_text(tsv_1col, dynamic=True)
        self.assertEqual(dir_count_1, 1)
        self.assertEqual(len(rules_1), 2)
        self.assertEqual(rules_1[0].dirs, ["PART_A"])
        self.assertEqual(rules_1[0].file_pattern, "")

        # 2 欄 (2 層目錄: 料號 + 批號)
        tsv_2cols = "PART_A\tLOT_001\nPART_B\tLOT_002"
        rules_2, dir_count_2 = CTkRuleTable.parse_tsv_text(tsv_2cols, dynamic=True)
        self.assertEqual(dir_count_2, 2)
        self.assertEqual(len(rules_2), 2)
        self.assertEqual(rules_2[0].dirs, ["PART_A", "LOT_001"])
        self.assertEqual(rules_2[0].file_pattern, "")

        # 3 欄 (3 層目錄: 料號 + 批號 + 機台)
        tsv_3cols = "PART_A\tLOT_001\tSTATION_1\nPART_B\tLOT_002\tSTATION_2"
        rules_3, dir_count_3 = CTkRuleTable.parse_tsv_text(tsv_3cols, dynamic=True)
        self.assertEqual(dir_count_3, 3)
        self.assertEqual(rules_3[0].dirs, ["PART_A", "LOT_001", "STATION_1"])
        self.assertEqual(rules_3[0].file_pattern, "")

    def test_8_2_traverse_at_final_level(self):
        """8.2 驗證抵達最後一層目錄後，自動依據目標檔案過濾條件 (如 *RAW*) 搜尋檔案"""
        mock_sftp = MagicMock()

        # 模擬遠端 2 層資料夾樹：
        # /
        #  └── PART_A
        #        └── LOT_001 (最後一層目錄)
        #              ├── 2026_RAW_data.csv (符合 *RAW*)
        #              ├── summary.xlsx (不符合 *RAW*)
        #              └── RAW_log.txt (符合 *RAW*)
        def mock_listdir(path):
            p = path.rstrip("/")
            if p in ("", "/"):
                return [DummySFTPAttr("PART_A", is_dir=True), DummySFTPAttr("PART_OTHER", is_dir=True)]
            elif p == "/PART_A":
                return [DummySFTPAttr("LOT_001", is_dir=True), DummySFTPAttr("LOT_OTHER", is_dir=True)]
            elif p == "/PART_A/LOT_001":
                return [
                    DummySFTPAttr("2026_RAW_data.csv", is_dir=False, size=4096),
                    DummySFTPAttr("summary.xlsx", is_dir=False, size=8192),
                    DummySFTPAttr("RAW_log.txt", is_dir=False, size=1024)
                ]
            return []

        mock_sftp.listdir_attr.side_effect = mock_listdir

        cancel_ev = MagicMock()
        cancel_ev.is_set.return_value = False

        # 下方目標檔案過濾條件設為 *RAW*
        search_cfg = SearchSettings(
            global_pattern="*RAW*",
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

        # Excel 貼上的兩層目錄規則 (檔案為空，完全由 global_pattern 主導)
        rules = [
            RuleItem(dirs=["PART_A", "LOT_001"], file_pattern="")
        ]
        stats = TaskStats()
        tasks = engine.scan_all(rules, stats)

        # 應該在最後一層目錄 /PART_A/LOT_001 撈出 2 個符合 *RAW* 的檔案
        self.assertEqual(len(tasks), 2)
        filenames = {t.filename for t in tasks}
        self.assertEqual(filenames, {"2026_RAW_data.csv", "RAW_log.txt"})

        for t in tasks:
            self.assertEqual(t.dirs, ["PART_A", "LOT_001"])
            self.assertTrue(t.remote_full_path.startswith("/PART_A/LOT_001/"))
            self.assertTrue(t.local_full_path.startswith(os.path.join(str(self.test_dir), "PART_A", "LOT_001")))

    def test_8_3_rule_table_headings(self):
        """8.3 驗證 CTkRuleTable 表頭純為目錄且無括號干擾"""
        from ui.main_window import MainWindow

        win = MainWindow()
        try:
            table = win.table_rules
            # 測試設定為 3 層
            table.set_columns(3)
            self.assertEqual(table.dir_count, 3)
            self.assertEqual(len(table.tree["columns"]), 3)

            headings = [table.tree.heading(col)["text"] for col in table.tree["columns"]]
            self.assertEqual(headings, ["第 1 層目錄", "第 2 層目錄", "第 3 層目錄"])
            for h in headings:
                self.assertNotIn("(", h)
                self.assertNotIn("檔案", h)
        finally:
            win.destroy()


if __name__ == "__main__":
    unittest.main()
