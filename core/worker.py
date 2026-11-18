import queue
import threading
import time
from dataclasses import dataclass
from typing import List, Any, Optional

from models.data_models import (
    SFTPConfig, SearchSettings, RuleItem, TaskStats, ReportRecord,
    LogLevel, TransferStatus, TaskCancelledError
)
from core.sftp_manager import SFTPManager
from core.traverse_engine import TraverseEngine
from core.download_manager import DownloadManager


def format_bytes(size: int) -> str:
    """將 Bytes 數值格式化為易讀容量字串"""
    s = float(size)
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if s < 1024.0 or unit == "TB":
            if unit == "B":
                return f"{int(s)} B"
            return f"{s:.2f} {unit}"
        s /= 1024.0
    return f"{size} B"


@dataclass
class UIEvent:
    """跨執行緒傳遞給 CustomTkinter 主迴圈的事件封包"""
    event_type: str  # "LOG" | "STATS" | "OVERALL_PROG" | "FILE_PROG" | "FINISHED" | "TEST_RESULT"
    payload: Any


class ConnectionTestWorker(threading.Thread):
    """獨立負責『測試連線』按鈕的背景執行緒，完全不卡死 UI"""

    def __init__(self, ui_queue: queue.Queue[UIEvent], config: SFTPConfig) -> None:
        super().__init__(daemon=True)
        self._q = ui_queue
        self._config = config

    def run(self) -> None:
        mgr = SFTPManager(self._config)
        ok, msg = mgr.test_connection()
        self._q.put(UIEvent("TEST_RESULT", (ok, msg)))


class DownloadWorker(threading.Thread):
    """純 Python 背景工作執行緒 (100% 相容 Windows SAC，零 PyQt6 依賴)"""

    THROTTLE_INTERVAL_SEC = 0.08  # 80ms 節流閘門

    def __init__(
        self,
        ui_queue: queue.Queue[UIEvent],
        sftp_config: SFTPConfig,
        search_settings: SearchSettings,
        rules: List[RuleItem],
        scan_only: bool = False
    ) -> None:
        super().__init__(daemon=True)
        self._q = ui_queue
        self._sftp_config = sftp_config
        self._search_settings = search_settings
        self._rules = rules
        self._scan_only = scan_only
        self._cancel_event = threading.Event()

        self._last_file_emit_time: float = 0.0
        self._last_file_pct: int = -1
        self._last_stats_emit_time: float = 0.0

    def request_cancel(self) -> None:
        """主動觸發取消旗標"""
        self._cancel_event.set()

    def run(self) -> None:
        stats = TaskStats()
        records: List[ReportRecord] = []
        sftp_mgr = SFTPManager(self._sftp_config)
        manifest_csv_path: Optional[str] = None

        try:
            self._emit_log(LogLevel.INFO, f"正在連線至 SFTP {self._sftp_config.host}:{self._sftp_config.port} ...")
            sftp_mgr.connect()
            self._emit_log(LogLevel.SUCCESS, "SFTP 連線建立成功。")

            # 階段一：三層目錄快取遍歷
            traverse_engine = TraverseEngine(
                sftp_mgr=sftp_mgr,
                initial_path=self._sftp_config.initial_path,
                search_settings=self._search_settings,
                cancel_event=self._cancel_event,
                log_cb=self._emit_log,
                stats_cb=lambda s: self._throttled_emit_stats(s, force=False)
            )
            tasks = traverse_engine.scan_all(self._rules, stats)
            empty_records = getattr(traverse_engine, "empty_records", [])
            self._throttled_emit_stats(stats, force=True)

            # 純檢索模式 (Scan-Only)
            if self._scan_only:
                self._q.put(UIEvent("OVERALL_PROG", 1.0))
                manifest_csv_path = DownloadManager.generate_scan_manifest_csv(
                    tasks, empty_records, self._search_settings.local_save_path
                )
                total_bytes = sum(t.file_size for t in tasks)
                readable_total = format_bytes(total_bytes)
                if manifest_csv_path:
                    self._emit_log(LogLevel.SUCCESS, f"已自動生成檢索清單：{manifest_csv_path}")

                summary = (
                    f"檢索完成！掃描目錄: {stats.scanned_dirs} | "
                    f"命中檔案: {len(tasks)} (共 {readable_total}) | "
                    f"空/無效目錄: {len(empty_records)}"
                )
                self._emit_log(LogLevel.SUCCESS, summary)
                self._q.put(UIEvent("FINISHED", (True, summary, [], manifest_csv_path)))
                return

            total_files = len(tasks)
            if total_files == 0:
                self._q.put(UIEvent("OVERALL_PROG", 1.0))
                self._emit_log(LogLevel.WARNING, "掃描完成，未找到任何符合條件的檔案。")
                if empty_records:
                    manifest_csv_path = DownloadManager.generate_manifest_csv(
                        empty_records, self._search_settings.local_save_path
                    )
                self._q.put(UIEvent("FINISHED", (True, "掃描完成，無符合條件之檔案。", records, manifest_csv_path)))
                return

            self._emit_log(LogLevel.INFO, f"共找到 {total_files} 個目標檔案，開始執行鏡像下載...")

            # 階段二：依序下載與鏡像保留
            dl_mgr = DownloadManager(
                sftp_mgr=sftp_mgr,
                strategy=self._search_settings.conflict_strategy,
                cancel_event=self._cancel_event,
                log_cb=self._emit_log,
                file_progress_cb=self._throttled_file_progress
            )

            for idx, task in enumerate(tasks, start=1):
                record = dl_mgr.process_task(task)
                records.append(record)

                if record.status == TransferStatus.SUCCESS:
                    stats.downloaded_ok += 1
                elif record.status == TransferStatus.SKIPPED:
                    stats.skipped_files += 1
                else:
                    stats.failed_files += 1

                self._throttled_emit_stats(stats, force=False)
                self._q.put(UIEvent("OVERALL_PROG", idx / total_files))

            self._throttled_emit_stats(stats, force=True)

            # 階段三：自動產出本地完整下載清單 Manifest CSV (合併下載結果與空目錄紀錄)
            all_records = records + empty_records
            manifest_csv_path = DownloadManager.generate_manifest_csv(
                all_records, self._search_settings.local_save_path
            )
            if manifest_csv_path:
                self._emit_log(LogLevel.SUCCESS, f"已自動生成完整下載清單：{manifest_csv_path}")

            summary = (
                f"任務完成！成功: {stats.downloaded_ok} | "
                f"略過: {stats.skipped_files} | 失敗: {stats.failed_files}"
            )
            self._emit_log(LogLevel.SUCCESS, summary)
            self._q.put(UIEvent("FINISHED", (True, summary, records, manifest_csv_path)))

        except TaskCancelledError:
            self._throttled_emit_stats(stats, force=True)
            all_records = records + getattr(traverse_engine, "empty_records", []) if 'traverse_engine' in locals() else records
            if all_records:
                manifest_csv_path = DownloadManager.generate_manifest_csv(
                    all_records, self._search_settings.local_save_path
                )
            msg = "任務已由使用者安全中止（已下載檔案與清單皆已保留）。"
            self._emit_log(LogLevel.WARNING, msg)
            self._q.put(UIEvent("FINISHED", (False, msg, records, manifest_csv_path)))

        except FileNotFoundError:
            msg = f"遠端初始路徑不存在：{self._sftp_config.initial_path}"
            self._emit_log(LogLevel.ERROR, msg)
            self._q.put(UIEvent("FINISHED", (False, msg, records, None)))

        except Exception as e:
            self._throttled_emit_stats(stats, force=True)
            msg = f"任務因未預期異常中止：{type(e).__name__} - {e}"
            self._emit_log(LogLevel.ERROR, msg)
            self._q.put(UIEvent("FINISHED", (False, msg, records, None)))

        finally:
            sftp_mgr.close()

    def _emit_log(self, level: LogLevel, msg: str) -> None:
        self._q.put(UIEvent("LOG", (level.value, msg)))

    def _throttled_file_progress(self, filename: str, transferred: int, total: int) -> None:
        if total <= 0:
            return
        pct = int((transferred / total) * 100)
        now = time.monotonic()
        if pct == 100 or (
            pct != self._last_file_pct and (now - self._last_file_emit_time) >= self.THROTTLE_INTERVAL_SEC
        ):
            self._last_file_emit_time = now
            self._last_file_pct = pct
            self._q.put(UIEvent("FILE_PROG", (filename, pct)))

    def _throttled_emit_stats(self, stats: TaskStats, force: bool = False) -> None:
        now = time.monotonic()
        if force or (now - self._last_stats_emit_time) >= self.THROTTLE_INTERVAL_SEC:
            self._last_stats_emit_time = now
            self._q.put(
                UIEvent(
                    "STATS",
                    TaskStats(
                        scanned_dirs=stats.scanned_dirs,
                        matched_files=stats.matched_files,
                        downloaded_ok=stats.downloaded_ok,
                        skipped_files=stats.skipped_files,
                        failed_files=stats.failed_files
                    )
                )
            )
