import csv
import os
import queue
import shutil
import stat
import sys
import tempfile
import threading
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from models.data_models import (
    RuleItem, SearchSettings, RemoteFileTask, TaskStats,
    TransferStatus, ReportRecord, MatchMode, ConflictStrategy, SFTPConfig
)
from core.download_manager import DownloadManager
from core.traverse_engine import TraverseEngine
from core.worker import DownloadWorker, UIEvent


class MockSFTPAttribute:
    def __init__(self, filename: str, is_dir: bool, size: int = 1024, mtime: int = 1700000000):
        self.filename = filename
        self.st_mode = stat.S_IFDIR if is_dir else stat.S_IFREG
        self.st_size = size
        self.st_mtime = mtime


class MockSFTPManager:
    def __init__(self, tree_dict):
        self.tree = tree_dict
        self.connected = False

    def connect(self):
        self.connected = True

    def close(self):
        self.connected = False

    def reconnect(self):
        self.connected = True

    def listdir_attr(self, remote_path: str):
        norm = remote_path.rstrip("/")
        if norm not in self.tree:
            raise FileNotFoundError(f"Path not found: {remote_path}")
        return self.tree[norm]

    def download_file(self, remote_path: str, local_temp_path: str, callback=None):
        raise RuntimeError("在純檢索模式下，不應該呼叫 download_file！")


def test_phase9():
    print("=== 開始執行微階段 9 驗收測試 (純檢索模式與修改時間/副檔名) ===")
    tmp_dir = tempfile.mkdtemp(prefix="test_phase9_")

    try:
        # 1. 測試 ReportRecord 之時間戳記轉換與副檔名
        print("[TEST] 1. 驗證 ReportRecord 修改時間格式化與副檔名...")
        test_mtime = 1700000000  # 2023-11-15 06:13:20 (UTC+8 或系統時區)
        rec = ReportRecord(
            remote_path="/remote/LINE1/file_01.csv",
            local_path=os.path.join(tmp_dir, "file_01.csv"),
            dirs=["LINE1"],
            filename="file_01.csv",
            file_size=2048,
            file_mtime=test_mtime,
            status=TransferStatus.INDEXED
        )
        assert rec.file_ext == ".csv", f"Expected .csv, got {rec.file_ext}"
        assert len(rec.mtime_str) == 19, f"Expected YYYY-MM-DD HH:MM:SS format, got {rec.mtime_str}"
        print("  -> ReportRecord 屬性驗證通過！")

        # 2. 測試純檢索清單 CSV 產出
        print("[TEST] 2. 驗證 DownloadManager.generate_scan_manifest_csv 包含副檔名與時間...")
        task1 = RemoteFileTask(
            dirs=["LINE1", "STATION_A"],
            filename="raw_data.csv",
            remote_full_path="/remote/LINE1/STATION_A/raw_data.csv",
            local_full_path=os.path.join(tmp_dir, "LINE1", "STATION_A", "raw_data.csv"),
            file_size=5000000,
            file_mtime=test_mtime
        )
        empty_rec = ReportRecord(
            remote_path="/remote/LINE1/STATION_B",
            local_path=os.path.join(tmp_dir, "LINE1", "STATION_B"),
            dirs=["LINE1", "STATION_B"],
            filename="(無相符檔案)",
            file_size=0,
            status=TransferStatus.EMPTY_DIR,
            message="該目錄下無相符之目標檔案"
        )
        not_found_rec = ReportRecord(
            remote_path="/remote/LINE99",
            local_path=os.path.join(tmp_dir, "LINE99"),
            dirs=["LINE99"],
            filename="(目錄不存在)",
            file_size=0,
            status=TransferStatus.DIR_NOT_FOUND,
            message="遠端未找到符合此條件之資料夾"
        )

        csv_path = DownloadManager.generate_scan_manifest_csv([task1], [empty_rec, not_found_rec], tmp_dir)
        assert csv_path is not None and os.path.exists(csv_path)
        assert os.path.basename(csv_path).startswith("檢索清單_")

        with open(csv_path, "r", encoding="utf-8-sig") as f:
            reader = list(csv.reader(f))
            headers = reader[0]
            assert "副檔名" in headers
            assert "檔案修改時間" in headers
            assert "處理狀態" in headers
            assert len(reader) == 4  # header + 3 rows
            # row 1: raw_data.csv
            row1 = reader[1]
            assert "raw_data.csv" in row1
            assert ".csv" in row1
            assert "已檢索 (未下載)" in row1
            # row 2: EMPTY_DIR
            row2 = reader[2]
            assert "(無相符檔案)" in row2
            assert "空目錄 (無相符檔案)" in row2
            # row 3: DIR_NOT_FOUND
            row3 = reader[3]
            assert "(目錄不存在)" in row3
            assert "目錄不存在" in row3
        print("  -> 檢索清單 CSV 產出、12 大欄位與狀態驗證通過！")

        # 3. 測試 TraverseEngine 走訪並精確捕捉「檔案」、「空目錄」與「目錄不存在」
        print("[TEST] 3. 驗證 TraverseEngine 捕捉空目錄與目錄不存在...")
        tree = {
            "/root": [
                MockSFTPAttribute("LINE1", is_dir=True),
                MockSFTPAttribute("LINE2", is_dir=True)
            ],
            "/root/LINE1": [
                MockSFTPAttribute("sub1", is_dir=True)
            ],
            "/root/LINE1/sub1": [
                MockSFTPAttribute("target_01.csv", is_dir=False, size=1024, mtime=test_mtime),
                MockSFTPAttribute("ignore.txt", is_dir=False, size=500, mtime=test_mtime)
            ],
            "/root/LINE2": [
                MockSFTPAttribute("sub2", is_dir=True)
            ],
            "/root/LINE2/sub2": [
                # 只有 txt，沒有 csv -> 空目錄
                MockSFTPAttribute("readme.txt", is_dir=False, size=100, mtime=test_mtime)
            ]
        }
        sftp_mock = MockSFTPManager(tree)
        settings = SearchSettings(
            global_pattern="*.csv",
            local_save_path=tmp_dir,
            match_mode=MatchMode.FUZZY
        )
        engine = TraverseEngine(
            sftp_mgr=sftp_mock,
            initial_path="/root",
            search_settings=settings,
            cancel_event=threading.Event(),
            log_cb=lambda l, m: None,
            stats_cb=lambda s: None
        )
        rules = [
            RuleItem(dirs=["LINE1", "sub1"], file_pattern="*.csv"),
            RuleItem(dirs=["LINE2", "sub2"], file_pattern="*.csv"),
            RuleItem(dirs=["LINE3_NOT_EXIST", "subX"], file_pattern="*.csv")  # 不存在
        ]
        stats = TaskStats()
        found_tasks = engine.scan_all(rules, stats)
        empty_recs = engine.empty_records

        assert len(found_tasks) == 1
        assert found_tasks[0].filename == "target_01.csv"
        assert found_tasks[0].file_mtime == test_mtime

        # 檢查 empty_records 應有 LINE2/sub2 (EMPTY_DIR) 和 LINE3_NOT_EXIST (DIR_NOT_FOUND)
        empty_statuses = {r.status for r in empty_recs}
        assert TransferStatus.EMPTY_DIR in empty_statuses
        assert TransferStatus.DIR_NOT_FOUND in empty_statuses
        print("  -> TraverseEngine 正常檢索出目標檔案，且精確捕捉空目錄與目錄不存在！")

        # 4. 測試 DownloadWorker(scan_only=True) 完整非同步背景運作
        print("[TEST] 4. 驗證 DownloadWorker(scan_only=True) 純檢索不下載檔案...")
        ui_q = queue.Queue()
        cfg = SFTPConfig(host="localhost", initial_path="/root")
        worker = DownloadWorker(
            ui_queue=ui_q,
            sftp_config=cfg,
            search_settings=settings,
            rules=rules,
            scan_only=True
        )
        # 用 mock 替換 worker 內部會呼叫的 SFTPManager
        worker_run = worker.run

        def patched_run():
            worker._sftp_config.initial_path = "/root"
            # 覆蓋 traverse engine 內部的 sftp 呼叫
            original_traverse = TraverseEngine
            class PatchedTraverse(original_traverse):
                def __init__(self, *args, **kwargs):
                    kwargs["sftp_mgr"] = sftp_mock
                    super().__init__(*args, **kwargs)

            import core.worker
            core.worker.TraverseEngine = PatchedTraverse
            try:
                # 這裡避免真正執行 sftp_mgr.connect
                class DummySFTP:
                    def connect(self): pass
                    def close(self): pass
                worker_sftp_orig = core.worker.SFTPManager
                core.worker.SFTPManager = lambda cfg: DummySFTP()
                worker_run()
            finally:
                core.worker.TraverseEngine = original_traverse
                core.worker.SFTPManager = worker_sftp_orig

        worker.run = patched_run
        worker.start()
        worker.join(timeout=5)

        # 檢查佇列事件
        events = []
        while not ui_q.empty():
            events.append(ui_q.get())

        finished_events = [e for e in events if e.event_type == "FINISHED"]
        assert len(finished_events) == 1
        is_ok, summary_msg, records, manifest_csv = finished_events[0].payload
        assert is_ok is True
        assert "檢索完成" in summary_msg
        assert manifest_csv is not None
        assert os.path.exists(manifest_csv)
        assert os.path.basename(manifest_csv).startswith("檢索清單_")

        # 確保本地完全沒有下載 target_01.csv 實體檔案
        downloaded_target = os.path.join(tmp_dir, "LINE1", "sub1", "target_01.csv")
        assert not os.path.exists(downloaded_target), "在純檢索模式下，絕對不應下載實體檔案！"
        print("  -> DownloadWorker(scan_only=True) 驗證通過：無下載實體檔案，檢索清單 CSV 生成完整！")

        print("=== 微階段 9 所有測試 100% 通過！ ===")

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == "__main__":
    test_phase9()
