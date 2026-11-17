import csv
import datetime
import os
import socket
import threading
import time
from typing import Callable, List, Optional, Tuple
import paramiko

from models.data_models import (
    RemoteFileTask, ConflictStrategy, TransferStatus, ReportRecord, LogLevel, TaskCancelledError
)
from core.sftp_manager import SFTPManager


class DownloadManager:
    """
    負責建立本地鏡像目錄、檢查衝突略過策略、執行 .part 原子下載、
    自動重試，以及在下載完成後自動生成本地完整下載清單 CSV (Manifest CSV)。
    """

    MAX_RETRY = 3

    def __init__(
        self,
        sftp_mgr: SFTPManager,
        strategy: ConflictStrategy,
        cancel_event: threading.Event,
        log_cb: Callable[[LogLevel, str], None],
        file_progress_cb: Callable[[str, int, int], None]
    ) -> None:
        self._sftp = sftp_mgr
        self._strategy = strategy
        self._cancel_event = cancel_event
        self._log = log_cb
        self._progress_cb = file_progress_cb

    def _check_cancel(self) -> None:
        """檢查是否收到中止訊號"""
        if self._cancel_event.is_set():
            raise TaskCancelledError("使用者主動中止下載")

    def should_skip_file(self, task: RemoteFileTask) -> Tuple[bool, str]:
        """根據使用者設定的 ConflictStrategy 判定是否略過下載"""
        if not os.path.exists(task.local_full_path):
            return False, ""

        if self._strategy == ConflictStrategy.OVERWRITE:
            return False, ""

        if self._strategy == ConflictStrategy.SKIP_EXISTING:
            return True, "本地已存在同名檔案 (策略：已存在則略過)"

        if self._strategy == ConflictStrategy.SIZE_MTIME:
            try:
                stat_res = os.stat(task.local_full_path)
                local_size = stat_res.st_size
                local_mtime = int(stat_res.st_mtime)
                # 容許 1 秒以內的檔案系統時間精度誤差
                if local_size == task.file_size and abs(local_mtime - task.file_mtime) <= 1:
                    return True, "本地檔案大小與修改時間相同，略過下載"
            except OSError:
                return False, ""

        return False, ""

    def process_task(self, task: RemoteFileTask) -> ReportRecord:
        """執行單一檔案檢查與下載（含 .part 暫存保護與 3 次斷線重連）"""
        self._check_cancel()

        # 1. 檢查是否符合略過條件
        skip, reason = self.should_skip_file(task)
        if skip:
            self._log(LogLevel.WARNING, f"[略過] {task.remote_full_path} ({reason})")
            return self._make_record(task, TransferStatus.SKIPPED, reason)

        # 2. 建立本地鏡像目錄
        local_dir = os.path.dirname(task.local_full_path)
        try:
            os.makedirs(local_dir, exist_ok=True)
        except OSError as e:
            msg = f"無法建立本地目錄 {local_dir}: {e}"
            self._log(LogLevel.ERROR, msg)
            return self._make_record(task, TransferStatus.FAILED, msg)

        temp_path = task.local_full_path + ".part"

        def _paramiko_cb(transferred: int, total: int) -> None:
            if self._cancel_event.is_set():
                raise TaskCancelledError("使用者於檔案傳輸途中按下停止")
            self._progress_cb(task.filename, transferred, total)

        # 3. 帶 Exponential Backoff 重試機制的下載迴圈
        for attempt in range(1, self.MAX_RETRY + 1):
            self._check_cancel()
            try:
                self._sftp.download_file(task.remote_full_path, temp_path, callback=_paramiko_cb)

                # 原子性替換為正式檔案，並同步遠端修改時間
                os.replace(temp_path, task.local_full_path)
                os.utime(task.local_full_path, (task.file_mtime, task.file_mtime))

                self._log(LogLevel.SUCCESS, f"[完成] {task.remote_full_path} -> {task.local_full_path}")
                return self._make_record(task, TransferStatus.SUCCESS, "下載成功")

            except TaskCancelledError:
                self._cleanup_temp(temp_path)
                raise

            except (PermissionError, FileNotFoundError) as e:
                self._cleanup_temp(temp_path)
                msg = f"檔案存取失敗 ({type(e).__name__}): {e}"
                self._log(LogLevel.ERROR, f"[失敗] {task.remote_full_path} - {msg}")
                return self._make_record(task, TransferStatus.FAILED, msg)

            except (socket.timeout, paramiko.SSHException, EOFError, OSError) as e:
                self._cleanup_temp(temp_path)
                if attempt < self.MAX_RETRY:
                    wait_sec = 2 ** (attempt - 1)
                    self._log(
                        LogLevel.WARNING,
                        f"下載中斷 ({task.filename}): {e}，{wait_sec} 秒後重連並重試 ({attempt}/{self.MAX_RETRY})..."
                    )
                    time.sleep(wait_sec)
                    try:
                        self._sftp.reconnect()
                    except Exception as reconn_err:
                        self._log(LogLevel.WARNING, f"重連嘗試失敗: {reconn_err}")
                else:
                    msg = f"網路重試 {self.MAX_RETRY} 次後仍失敗: {e}"
                    self._log(LogLevel.ERROR, f"[失敗] {task.remote_full_path} - {msg}")
                    return self._make_record(task, TransferStatus.FAILED, msg)

        return self._make_record(task, TransferStatus.FAILED, "未知錯誤")

    @classmethod
    def generate_manifest_csv(cls, records: List[ReportRecord], local_root_dir: str) -> Optional[str]:
        """
        在本地儲存目標目錄根目錄下自動產出完整下載清單 CSV (Manifest CSV)。
        包含詳細欄位，支援動態多層目錄，採用 utf-8-sig 編碼確保 Excel 雙擊不亂碼。
        回傳生成的 CSV 完整絕對路徑。
        """
        if not records or not local_root_dir:
            return None

        try:
            os.makedirs(local_root_dir, exist_ok=True)
            timestamp_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            csv_filename = f"下載清單_{timestamp_str}.csv"
            csv_full_path = os.path.join(local_root_dir, csv_filename)

            max_dir_count = max((len(r.dirs) for r in records), default=2)
            max_dir_count = max(2, max_dir_count)

            headers = ["序號"]
            for i in range(max_dir_count):
                headers.append(f"第 {i + 1} 層目錄")
            headers.extend([
                "檔案名稱",
                "本地完整路徑",
                "遠端完整路徑",
                "檔案大小 (Bytes)",
                "檔案大小 (易讀)",
                "處理狀態",
                "處理說明",
                "完成時間"
            ])

            with open(csv_full_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(headers)

                for idx, r in enumerate(records, start=1):
                    row_dirs = list(r.dirs) + [""] * (max_dir_count - len(r.dirs))
                    row = [str(idx)]
                    row.extend(row_dirs)
                    row.extend([
                        r.filename,
                        r.local_path,
                        r.remote_path,
                        str(r.file_size),
                        r.readable_size,
                        r.status.value,
                        r.message,
                        r.timestamp
                    ])
                    writer.writerow(row)

            return csv_full_path

        except Exception as e:
            return None

    def _make_record(self, task: RemoteFileTask, status: TransferStatus, message: str) -> ReportRecord:
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        return ReportRecord(
            remote_path=task.remote_full_path,
            local_path=task.local_full_path,
            dirs=task.dirs,
            filename=task.filename,
            file_size=task.file_size,
            status=status,
            message=message,
            timestamp=now_str
        )

    @staticmethod
    def _cleanup_temp(temp_path: str) -> None:
        try:
            if os.path.exists(temp_path):
                os.remove(temp_path)
        except OSError:
            pass
