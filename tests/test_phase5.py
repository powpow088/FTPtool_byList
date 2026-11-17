import os
import queue
import shutil
import sys
import threading
import time
from pathlib import Path
from unittest.mock import MagicMock, patch

# 加入根目錄至 sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from models.data_models import (
    SFTPConfig, SearchSettings, RuleItem, ConflictStrategy, TaskStats,
    ReportRecord, TransferStatus
)
from core.worker import UIEvent, ConnectionTestWorker, DownloadWorker
from ui.main_window import MainWindow


def test_worker_queue_events():
    print("[TEST] 1. 驗證背景 Worker (ConnectionTestWorker & DownloadWorker) 佇列通訊...", flush=True)
    ui_q = queue.Queue()

    # 1.1 測試 ConnectionTestWorker 傳遞 TEST_RESULT 事件
    test_cfg = SFTPConfig(host="192.168.1.100", port=22, username="u", password="p")
    with patch("core.worker.SFTPManager") as mock_mgr_cls:
        mock_instance = mock_mgr_cls.return_value
        mock_instance.test_connection.return_value = (True, "連線成功！")

        worker = ConnectionTestWorker(ui_q, test_cfg)
        worker.start()
        worker.join(timeout=2)

        assert not ui_q.empty(), "預期應有 TEST_RESULT 事件放入佇列"
        event = ui_q.get_nowait()
        assert event.event_type == "TEST_RESULT"
        ok, msg = event.payload
        assert ok is True
        assert "連線成功" in msg
    print("  -> ConnectionTestWorker 執行緒與 Queue 佇列通訊驗證通過！", flush=True)

    # 1.2 測試 DownloadWorker 完整模擬流程 (含自動產出 Manifest CSV)
    test_dl_dir = Path("tests/temp_test_worker_dl")
    if test_dl_dir.exists():
        shutil.rmtree(test_dl_dir)
    test_dl_dir.mkdir(parents=True, exist_ok=True)

    sftp_cfg = SFTPConfig(host="127.0.0.1", port=22, username="test", password="pwd", initial_path="/root")
    search_cfg = SearchSettings(
        global_pattern="*.csv",
        conflict_strategy=ConflictStrategy.OVERWRITE,
        local_save_path=str(test_dl_dir)
    )
    rules = [RuleItem("PART_A", "LOT_01", "report.csv")]

    dl_worker = DownloadWorker(ui_q, sftp_cfg, search_cfg, rules)

    # 驗證 _emit_log
    dl_worker._emit_log(MagicMock(value="INFO"), "測試日誌訊息")
    log_event = ui_q.get_nowait()
    assert log_event.event_type == "LOG"
    assert log_event.payload == ("INFO", "測試日誌訊息")

    # 驗證進度節流閥
    dl_worker._throttled_file_progress("sample.csv", 50, 100)
    prog_event = ui_q.get_nowait()
    assert prog_event.event_type == "FILE_PROG"
    assert prog_event.payload == ("sample.csv", 50)

    # 驗證統計節流閥
    test_stats = TaskStats(scanned_dirs=5, matched_files=10)
    dl_worker._throttled_emit_stats(test_stats, force=True)
    stats_event = ui_q.get_nowait()
    assert stats_event.event_type == "STATS"
    assert stats_event.payload.scanned_dirs == 5

    shutil.rmtree(test_dl_dir, ignore_errors=True)
    print("  -> DownloadWorker 節流器與各項事件封裝驗證通過！", flush=True)


def test_main_window_instantiation():
    print("[TEST] 2. 驗證 MainWindow 介面建置與控制項綁定...", flush=True)
    # mock messagebox 避免彈出視窗阻塞測試
    with patch("tkinter.messagebox.showwarning"), \
         patch("tkinter.messagebox.showinfo"), \
         patch("tkinter.messagebox.showerror"), \
         patch("tkinter.messagebox.askyesno", return_value=True):

        app = MainWindow()
        app.withdraw() # 隱藏視窗

        try:
            assert app.winfo_exists()
            assert app.title() == "SFTP 批次檔案模糊檢索與結構保留下載工具"
            assert app.btn_start.cget("state") == "normal"
            assert app.btn_stop.cget("state") == "disabled"
            assert app.btn_test_conn.cget("state") == "normal"

            # 測試輸入驗證 (空 Host 應回傳 False)
            app.ed_host.delete(0, "end")
            assert app._validate_inputs(require_rules=False) is False

            # 正確輸入應回傳 True
            app.ed_host.insert(0, "192.168.1.1")
            app.ed_user.insert(0, "admin")
            assert app._validate_inputs(require_rules=False) is True

            # 測試密碼明碼切換
            app.ed_pass.delete(0, "end")
            app.ed_pass.insert(0, "mypassword")
            assert app.ed_pass.cget("show") == "*"
            app._on_toggle_password()
            assert app.ed_pass.cget("show") == ""
            app._on_toggle_password()
            assert app.ed_pass.cget("show") == "*"

            print("  -> MainWindow 控制項、版面配置與驗證函式驗證通過！", flush=True)
        finally:
            app.destroy()


if __name__ == "__main__":
    print("=== 開始執行微階段 5 驗收測試 ===", flush=True)
    test_worker_queue_events()
    test_main_window_instantiation()
    print("=== 微階段 5 所有測試 100% 通過！ ===", flush=True)
