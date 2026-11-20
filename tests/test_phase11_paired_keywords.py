import os
import sys
import shutil
import csv
from pathlib import Path

# 將專案根目錄加入模組搜尋路徑
project_root = Path(__file__).resolve().parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from models.data_models import (
    ReportRecord,
    RemoteFileTask,
    TransferStatus,
    RuleItem,
    SearchSettings,
    MatchMode
)
from core.download_manager import DownloadManager
from core.traverse_engine import TraverseEngine


class MockSFTPClient:
    """模擬 SFTP Client 用於測試 TraverseEngine 走訪與關鍵字成對關聯"""
    class MockStat:
        def __init__(self, filename, is_dir=False, size=1024, mtime=1700000000):
            self.filename = filename
            self.st_mode = 0o040755 if is_dir else 0o100644
            self.st_size = size
            self.st_mtime = mtime

    def __init__(self):
        self.dirs = {
            "/": [
                self.MockStat("202611_Sales_Report", is_dir=True),
                self.MockStat("202510_Old_Report", is_dir=True)
            ],
            "/202611_Sales_Report": [
                self.MockStat("ABC_Site_North", is_dir=True),
                self.MockStat("EMPTY_Site_South", is_dir=True)
            ],
            "/202611_Sales_Report/ABC_Site_North": [
                self.MockStat("data_01.csv", is_dir=False, size=2048, mtime=1700001111),
                self.MockStat("readme.txt", is_dir=False, size=512, mtime=1700002222)
            ],
            "/202611_Sales_Report/EMPTY_Site_South": [] # 空目錄
        }

    def listdir_attr(self, path):
        clean_path = path.rstrip("/")
        if not clean_path:
            clean_path = "/"
        if clean_path in self.dirs:
            return self.dirs[clean_path]
        raise FileNotFoundError(f"Path not found: {path}")


def test_paired_keywords_manifest_csv():
    """驗證方案 A：Manifest CSV 成對目錄與搜尋關鍵字生成"""
    test_dir = Path("tests/temp_paired_csv_test")
    if test_dir.exists():
        shutil.rmtree(test_dir)
    test_dir.mkdir(parents=True, exist_ok=True)

    records = [
        ReportRecord(
            remote_path="/home/data/202611_Sales_Report/ABC_Site_North/data_01.csv",
            local_path="C:\\Downloads\\202611_Sales_Report\\ABC_Site_North\\data_01.csv",
            dirs=["202611_Sales_Report", "ABC_Site_North"],
            search_dirs=["2026", "ABC"],
            filename="data_01.csv",
            file_size=2048,
            file_mtime=1700001111,
            status=TransferStatus.SUCCESS,
            message="下載成功",
            timestamp="2026-11-20 20:00:00"
        ),
        ReportRecord(
            remote_path="/home/data/202611_Sales_Report/EMPTY_Site_South",
            local_path="C:\\Downloads\\202611_Sales_Report\\EMPTY_Site_South",
            dirs=["202611_Sales_Report", "EMPTY_Site_South"],
            search_dirs=["2026", "EMPTY"],
            filename="(無相符檔案)",
            file_size=0,
            file_mtime=0,
            status=TransferStatus.EMPTY_DIR,
            message="該目錄下無相符之目標檔案",
            timestamp="2026-11-20 20:00:01"
        ),
        ReportRecord(
            remote_path="/home/data/GHOST_FOLDER",
            local_path="C:\\Downloads\\GHOST_FOLDER",
            dirs=["GHOST_FOLDER"],
            search_dirs=["GHOST_FOLDER"],
            filename="(目錄不存在)",
            file_size=0,
            file_mtime=0,
            status=TransferStatus.DIR_NOT_FOUND,
            message="遠端未找到符合此條件之資料夾",
            timestamp="2026-11-20 20:00:02"
        )
    ]

    csv_path = DownloadManager.generate_manifest_csv(records, str(test_dir), prefix="下載清單", include_local_path=True)
    assert csv_path is not None
    assert os.path.exists(csv_path)

    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = list(csv.reader(f))
        assert len(reader) == 4 # 1 header + 3 rows
        header = reader[0]
        # 驗證成對表頭
        assert header[0] == "序號"
        assert header[1] == "第 1 層目錄"
        assert header[2] == "第 1 層搜尋字"
        assert header[3] == "第 2 層目錄"
        assert header[4] == "第 2 層搜尋字"
        assert header[5] == "檔案名稱"
        assert header[6] == "副檔名"
        assert header[7] == "檔案大小 (Bytes)"
        assert header[8] == "檔案大小 (易讀)"
        assert header[9] == "檔案修改時間"
        assert header[10] == "處理狀態"
        assert header[11] == "處理說明"
        assert header[12] == "遠端完整路徑"
        assert header[13] == "本地完整路徑"
        assert header[14] == "完成時間"

        # 驗證 row 1: 正常檔案
        r1 = reader[1]
        assert r1[1] == "202611_Sales_Report"
        assert r1[2] == "2026"
        assert r1[3] == "ABC_Site_North"
        assert r1[4] == "ABC"
        assert r1[5] == "data_01.csv"
        assert r1[6] == ".csv"
        assert r1[10] == "下載成功"

        # 驗證 row 2: 空目錄
        r2 = reader[2]
        assert r2[1] == "202611_Sales_Report"
        assert r2[2] == "2026"
        assert r2[3] == "EMPTY_Site_South"
        assert r2[4] == "EMPTY"
        assert r2[10] == "空目錄 (無相符檔案)"

        # 驗證 row 3: 目錄不存在
        r3 = reader[3]
        assert r3[1] == "GHOST_FOLDER"
        assert r3[2] == "GHOST_FOLDER"
        assert r3[10] == "目錄不存在"

    shutil.rmtree(test_dir, ignore_errors=True)
    print("  -> test_paired_keywords_manifest_csv 驗證通過！")


def test_traverse_engine_passes_search_dirs():
    """驗證 TraverseEngine 在走訪時能夠正確綁定並向下傳遞搜尋字元組"""
    import threading
    from unittest.mock import MagicMock
    mock_sftp = MockSFTPClient()
    settings = SearchSettings(
        global_pattern="*.csv",
        match_mode=MatchMode.FUZZY,
        local_save_path="C:\\Downloads"
    )
    cancel_ev = threading.Event()
    engine = TraverseEngine(
        sftp_mgr=mock_sftp,
        initial_path="/",
        search_settings=settings,
        cancel_event=cancel_ev,
        log_cb=lambda lvl, msg: None,
        stats_cb=lambda s: None
    )

    from models.data_models import TaskStats
    stats = TaskStats()

    # 規則 1: 模糊匹配第 1 層 2026，第 2 層 ABC
    # 規則 2: 模糊匹配第 1 層 2026，第 2 層 EMPTY
    # 規則 3: 目錄不存在 GHOST
    rules = [
        RuleItem(dirs=["2026", "ABC"], file_pattern="*.csv"),
        RuleItem(dirs=["2026", "EMPTY"], file_pattern="*.csv"),
        RuleItem(dirs=["GHOST_DIR", "SUB"], file_pattern="*.csv")
    ]

    tasks = engine.scan_all(rules, stats)

    # 驗證任務中的 search_dirs
    assert len(tasks) == 1
    t = tasks[0]
    assert t.filename == "data_01.csv"
    assert t.dirs == ["202611_Sales_Report", "ABC_Site_North"]
    assert t.search_dirs == ["2026", "ABC"]

    # 驗證 empty_records 中的 search_dirs
    assert len(engine.empty_records) == 2
    empty_dirs = [r for r in engine.empty_records if r.status == TransferStatus.EMPTY_DIR]
    not_found_dirs = [r for r in engine.empty_records if r.status == TransferStatus.DIR_NOT_FOUND]

    assert len(empty_dirs) == 1
    assert empty_dirs[0].dirs == ["202611_Sales_Report", "EMPTY_Site_South"]
    assert empty_dirs[0].search_dirs == ["2026", "EMPTY"]

    assert len(not_found_dirs) == 1
    assert not_found_dirs[0].search_dirs == ["GHOST_DIR", "SUB"]

    print("  -> test_traverse_engine_passes_search_dirs 驗證通過！")


if __name__ == "__main__":
    print("=== 執行微階段 15 (方案 A 成對搜尋關鍵字) 單元測試 ===")
    test_paired_keywords_manifest_csv()
    test_traverse_engine_passes_search_dirs()
    print("=== 微階段 15 所有測試 100% 通過！ ===")
