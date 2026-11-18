import csv
import os
import shutil
import sys
import threading
import time
from pathlib import Path
from unittest.mock import MagicMock

# 加入根目錄至 sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from models.data_models import (
    ConflictStrategy, RemoteFileTask, TransferStatus, ReportRecord, TaskCancelledError
)
from core.download_manager import DownloadManager


def test_conflict_strategy():
    print("[TEST] 1. 驗證重複檔案略過策略 (ConflictStrategy)...")
    test_dir = Path("tests/temp_test_dl")
    if test_dir.exists():
        shutil.rmtree(test_dir)
    test_dir.mkdir(parents=True, exist_ok=True)

    local_file = test_dir / "sample.csv"
    content = b"header1,header2\nval1,val2\n"
    local_file.write_bytes(content)
    now_ts = int(time.time())
    os.utime(local_file, (now_ts, now_ts))
    file_size = len(content)

    mock_sftp = MagicMock()
    cancel_event = threading.Event()

    task_match = RemoteFileTask(
        l1_dir="L1", l2_dir="L2", filename="sample.csv",
        remote_full_path="/remote/sample.csv", local_full_path=str(local_file),
        file_size=file_size, file_mtime=now_ts
    )

    # 1.1 SIZE_MTIME 策略：大小與時間相同 -> 略過
    dm_size_mtime = DownloadManager(mock_sftp, ConflictStrategy.SIZE_MTIME, cancel_event, lambda l, m: None, lambda f, tr, tot: None)
    skip, reason = dm_size_mtime.should_skip_file(task_match)
    assert skip is True
    assert "大小與修改時間相同" in reason

    # 1.2 SIZE_MTIME 策略：大小或時間不同 -> 不略過
    task_diff_size = RemoteFileTask(
        l1_dir="L1", l2_dir="L2", filename="sample.csv",
        remote_full_path="/remote/sample.csv", local_full_path=str(local_file),
        file_size=file_size + 999, file_mtime=now_ts
    )
    skip, _ = dm_size_mtime.should_skip_file(task_diff_size)
    assert skip is False

    # 1.3 OVERWRITE 策略：一律不略過
    dm_overwrite = DownloadManager(mock_sftp, ConflictStrategy.OVERWRITE, cancel_event, lambda l, m: None, lambda f, tr, tot: None)
    skip, _ = dm_overwrite.should_skip_file(task_match)
    assert skip is False

    # 1.4 SKIP_EXISTING 策略：只要存在一律略過
    dm_skip = DownloadManager(mock_sftp, ConflictStrategy.SKIP_EXISTING, cancel_event, lambda l, m: None, lambda f, tr, tot: None)
    skip, _ = dm_skip.should_skip_file(task_match)
    assert skip is True

    print("  -> 重複檔案處理策略驗證通過！")


def test_atomic_download_and_cancellation():
    print("[TEST] 2. 驗證原子下載 (.part) 與取消中斷暫存檔清理...")
    test_dir = Path("tests/temp_test_atomic")
    if test_dir.exists():
        shutil.rmtree(test_dir)
    test_dir.mkdir(parents=True, exist_ok=True)

    target_file = test_dir / "subdir" / "data.csv"
    remote_ts = 1700000000

    task = RemoteFileTask(
        l1_dir="DIR1", l2_dir="DIR2", filename="data.csv",
        remote_full_path="/remote/data.csv", local_full_path=str(target_file),
        file_size=1024, file_mtime=remote_ts
    )

    mock_sftp = MagicMock()
    cancel_event = threading.Event()

    def fake_download_ok(remote_path, temp_path, callback=None):
        with open(temp_path, "wb") as f:
            f.write(b"data content " * 50)
        if callback:
            callback(100, 100)

    mock_sftp.download_file.side_effect = fake_download_ok

    dm = DownloadManager(mock_sftp, ConflictStrategy.OVERWRITE, cancel_event, lambda l, m: None, lambda f, tr, tot: None)
    record = dm.process_task(task)

    assert record.status == TransferStatus.SUCCESS
    assert target_file.exists()
    assert not (Path(str(target_file) + ".part")).exists()
    assert abs(int(os.stat(target_file).st_mtime) - remote_ts) <= 2
    print("  -> 原子下載與時間戳同步驗證通過！")

    # 測試傳輸途中取消中斷並自動清理 .part
    def fake_download_cancel(remote_path, temp_path, callback=None):
        with open(temp_path, "wb") as f:
            f.write(b"partial dirty data")
        cancel_event.set() # 模擬使用者按下停止
        if callback:
            callback(10, 100) # 回呼時觸發 TaskCancelledError

    mock_sftp.download_file.side_effect = fake_download_cancel
    cancel_task = RemoteFileTask(
        l1_dir="DIR1", l2_dir="DIR2", filename="cancelled.csv",
        remote_full_path="/remote/cancelled.csv", local_full_path=str(test_dir / "cancelled.csv"),
        file_size=2048, file_mtime=remote_ts
    )

    try:
        dm.process_task(cancel_task)
        assert False, "預期應拋出 TaskCancelledError"
    except TaskCancelledError:
        # 驗證 .part 檔已被安全刪除
        assert not (Path(str(test_dir / "cancelled.csv") + ".part")).exists()
        print("  -> 取消中斷與 .part 髒檔案自動清理驗證通過！")

    # 清理測試目錄
    shutil.rmtree(test_dir, ignore_errors=True)


def test_generate_manifest_csv():
    print("[TEST] 3. 驗證本地自動生成完整下載清單 CSV (Manifest CSV)...")
    manifest_dir = Path("tests/temp_test_manifest")
    if manifest_dir.exists():
        shutil.rmtree(manifest_dir)
    manifest_dir.mkdir(parents=True, exist_ok=True)

    records = [
        ReportRecord(
            remote_path="/home/data/PART_A/LOT_01/report1.csv",
            local_path="C:\\Downloads\\PART_A\\LOT_01\\report1.csv",
            l1_dir="PART_A",
            l2_dir="LOT_01",
            filename="report1.csv",
            file_size=1048576,
            status=TransferStatus.SUCCESS,
            message="下載成功",
            timestamp="2026-11-17 18:50:00"
        ),
        ReportRecord(
            remote_path="/home/data/PART_A/LOT_01/report2.csv",
            local_path="C:\\Downloads\\PART_A\\LOT_01\\report2.csv",
            l1_dir="PART_A",
            l2_dir="LOT_01",
            filename="report2.csv",
            file_size=512,
            status=TransferStatus.SKIPPED,
            message="本地檔案大小與修改時間相同，略過下載",
            timestamp="2026-11-17 18:50:01"
        )
    ]

    csv_path = DownloadManager.generate_manifest_csv(records, str(manifest_dir))
    assert csv_path is not None
    assert os.path.exists(csv_path)
    assert "下載清單_" in os.path.basename(csv_path)

    # 驗證以 utf-8-sig 讀回內容
    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = list(csv.reader(f))
        assert len(reader) == 3 # 1 row header + 2 rows records
        headers = reader[0]
        assert headers == [
            "序號", "第 1 層目錄", "第 2 層目錄", "檔案名稱", "副檔名",
            "檔案大小 (Bytes)", "檔案大小 (易讀)", "檔案修改時間",
            "處理狀態", "處理說明", "遠端完整路徑", "本地完整路徑", "完成時間"
        ]
        row1 = reader[1]
        assert row1[0] == "1"
        assert row1[1] == "PART_A"
        assert row1[2] == "LOT_01"
        assert row1[3] == "report1.csv"
        assert row1[4] == ".csv"
        assert row1[5] == "1048576"
        assert row1[6] == "1.00 MB"
        assert row1[8] == "下載成功"

        row2 = reader[2]
        assert row2[0] == "2"
        assert row2[4] == ".csv"
        assert row2[8] == "已略過"

    # 清理測試目錄
    shutil.rmtree(manifest_dir, ignore_errors=True)
    print("  -> 本地 Manifest CSV 10 大欄位與編碼驗證通過！")


if __name__ == "__main__":
    print("=== 開始執行微階段 3 驗收測試 ===")
    test_conflict_strategy()
    test_atomic_download_and_cancellation()
    test_generate_manifest_csv()
    print("=== 微階段 3 所有測試 100% 通過！ ===")
