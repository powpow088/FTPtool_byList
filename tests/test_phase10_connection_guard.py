import os
import shutil
import sys
import tempfile
import threading
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from models.data_models import (
    SFTPConfig, SearchSettings, RemoteFileTask, ConflictStrategy, TransferStatus
)
from core.sftp_manager import SFTPManager
from core.download_manager import DownloadManager


class DummyChannel:
    def __init__(self, closed=False):
        self.closed = closed


class DummyTransport:
    def __init__(self, active=True):
        self._active = active

    def is_active(self):
        return self._active


class MockGuardedSFTPManager:
    """模擬帶有連線狀態探測與計數的 SFTPManager"""

    def __init__(self):
        self.reconnect_called = 0
        self.download_called = 0
        self._active = True
        self.SOCKET_TIMEOUT = 30.0

    def is_active(self) -> bool:
        return self._active

    def reconnect(self) -> None:
        self.reconnect_called += 1
        self._active = True

    def ensure_connected(self) -> None:
        if not self.is_active():
            self.reconnect()

    def download_file(self, remote_path: str, local_temp_path: str, callback=None) -> None:
        if not self._active:
            raise RuntimeError("SFTP 已斷線！")
        self.download_called += 1
        # 寫入假檔案
        with open(local_temp_path, "wb") as f:
            f.write(b"mock data 12345")


def test_connection_guard():
    print("=== 開始執行微階段 10 驗收測試 (高可用連線防護與防封鎖微延遲) ===")
    tmp_dir = tempfile.mkdtemp(prefix="test_phase10_")

    try:
        # 1. 測試 is_active 與 ensure_connected 狀態探測
        print("[TEST] 1. 驗證 SFTPManager.is_active 與 ensure_connected 機制...")
        cfg = SFTPConfig(host="127.0.0.1")
        mgr = SFTPManager(cfg)
        assert mgr.is_active() is False, "未連線時 is_active 應回傳 False"

        # 模擬 mock 物件
        mock_mgr = MockGuardedSFTPManager()
        assert mock_mgr.is_active() is True
        mock_mgr._active = False
        assert mock_mgr.is_active() is False
        mock_mgr.ensure_connected()
        assert mock_mgr.is_active() is True
        assert mock_mgr.reconnect_called == 1
        print("  -> 連線健康探測與自動恢復驗證通過！")

        # 2. 測試接關續傳略過機制 (SIZE_MTIME 略過)
        print("[TEST] 2. 驗證接關續傳防護 (本地已存在同大小同時間檔案自動略過)...")
        local_file = os.path.join(tmp_dir, "existing.csv")
        with open(local_file, "wb") as f:
            f.write(b"already_downloaded_data")
        stat_info = os.stat(local_file)
        test_mtime = int(stat_info.st_mtime)
        test_size = stat_info.st_size

        task_existing = RemoteFileTask(
            dirs=["LINE1"],
            filename="existing.csv",
            remote_full_path="/remote/LINE1/existing.csv",
            local_full_path=local_file,
            file_size=test_size,
            file_mtime=test_mtime
        )

        logs = []
        dl_mgr = DownloadManager(
            sftp_mgr=mock_mgr,
            strategy=ConflictStrategy.SIZE_MTIME,
            cancel_event=threading.Event(),
            log_cb=lambda l, msg: logs.append((l, msg)),
            file_progress_cb=lambda f, t, total: None,
            request_delay_ms=0,
            session_refresh_interval=3
        )

        rec = dl_mgr.process_task(task_existing)
        assert rec.status == TransferStatus.SKIPPED
        assert any("[續傳略過]" in msg for _, msg in logs)
        print("  -> 續傳略過防護與 [續傳略過] 日誌標籤驗證通過！")

        # 3. 測試防封鎖微延遲 (Gentle Pacing Delay)
        print("[TEST] 3. 驗證防封鎖微延遲 (request_delay_ms)...")
        delay_ms = 80
        dl_mgr_delay = DownloadManager(
            sftp_mgr=mock_mgr,
            strategy=ConflictStrategy.OVERWRITE,
            cancel_event=threading.Event(),
            log_cb=lambda l, msg: None,
            file_progress_cb=lambda f, t, total: None,
            request_delay_ms=delay_ms,
            session_refresh_interval=100
        )
        task_new = RemoteFileTask(
            dirs=["LINE1"],
            filename="new1.csv",
            remote_full_path="/remote/LINE1/new1.csv",
            local_full_path=os.path.join(tmp_dir, "new1.csv"),
            file_size=15,
            file_mtime=int(time.time())
        )
        t0 = time.perf_counter()
        dl_mgr_delay.process_task(task_new)
        t1 = time.perf_counter()
        elapsed_ms = (t1 - t0) * 1000
        assert elapsed_ms >= delay_ms * 0.8, f"耗時 {elapsed_ms:.1f}ms 小於預期微延遲 {delay_ms}ms"
        print(f"  -> 微延遲實測耗時: {elapsed_ms:.2f} ms (設定: {delay_ms} ms)，防封鎖微延遲驗證通過！")

        # 4. 測試主動連線換氣 (Proactive Session Refresh)
        print("[TEST] 4. 驗證主動換氣 Session Refresh (每 N 個檔案自動重連)...")
        refresh_logs = []
        mock_mgr.reconnect_called = 0
        dl_mgr_refresh = DownloadManager(
            sftp_mgr=mock_mgr,
            strategy=ConflictStrategy.OVERWRITE,
            cancel_event=threading.Event(),
            log_cb=lambda l, msg: refresh_logs.append((l, msg)),
            file_progress_cb=lambda f, t, total: None,
            request_delay_ms=0,
            session_refresh_interval=3  # 設定每 3 個檔案主動換氣一次
        )

        for i in range(1, 8):
            t = RemoteFileTask(
                dirs=["LINE1"],
                filename=f"batch_{i}.csv",
                remote_full_path=f"/remote/LINE1/batch_{i}.csv",
                local_full_path=os.path.join(tmp_dir, f"batch_{i}.csv"),
                file_size=15,
                file_mtime=int(time.time())
            )
            dl_mgr_refresh.process_task(t)

        # 7 個檔案，每 3 個換氣一次 -> 在第 3, 6 個檔案換氣，共換氣 2 次
        assert mock_mgr.reconnect_called == 2, f"Expected 2 reconnects, got {mock_mgr.reconnect_called}"
        assert any("主動連線換氣" in msg for _, msg in refresh_logs)
        print("  -> 主動連線換氣 (Proactive Session Refresh) 觸發次數與機制驗證通過！")

        print("=== 微階段 10 所有測試 100% 通過！ ===")

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == "__main__":
    test_connection_guard()
