import datetime
import os
from dataclasses import dataclass, field
from enum import Enum
from typing import List, Dict, Set, Optional


class ConflictStrategy(Enum):
    """檔案衝突/重複處理策略"""
    SIZE_MTIME = "size_mtime"   # 比對大小與修改時間略過 (預設推薦)
    OVERWRITE = "overwrite"     # 直接覆蓋
    SKIP_EXISTING = "skip"      # 已存在則略過


class MatchMode(Enum):
    """目錄比對模式"""
    FUZZY = "fuzzy"             # 模糊包含 (Contains) - 預設
    EXACT = "exact"             # 完全精準相符 (Exact Match)


class LogLevel(Enum):
    """即時日誌等級及其 HEX 顯示色碼"""
    INFO = "INFO"         # 藍色
    SUCCESS = "SUCCESS"   # 綠色
    WARNING = "WARNING"   # 橘黃色
    ERROR = "ERROR"       # 紅色

    @property
    def color_hex(self) -> str:
        color_map = {
            LogLevel.INFO: "#0066CC",
            LogLevel.SUCCESS: "#28A745",
            LogLevel.WARNING: "#D39E00",
            LogLevel.ERROR: "#DC3545",
        }
        return color_map.get(self, "#333333")


class TransferStatus(Enum):
    """單一檔案傳輸狀態"""
    SUCCESS = "下載成功"
    SKIPPED = "已略過"
    FAILED = "失敗"
    CANCELLED = "已取消"
    INDEXED = "已檢索 (未下載)"
    EMPTY_DIR = "空目錄 (無相符檔案)"
    DIR_NOT_FOUND = "目錄不存在"


class TaskCancelledError(Exception):
    """使用者主動點擊停止時拋出的內部控制例外（毫秒級中斷）"""
    pass


@dataclass
class SFTPConfig:
    """SFTP 連線設定"""
    host: str = ""
    port: int = 22
    username: str = ""
    password: str = ""
    save_password: bool = False
    initial_path: str = "/"

    def normalized_initial_path(self) -> str:
        """確保初始路徑以 / 開頭"""
        path = self.initial_path.strip()
        if not path:
            return "/"
        return path if path.startswith("/") else "/" + path


@dataclass
class RuleItem:
    """單一檢索規則列（支援動態 N 層目錄與檔案名稱）"""
    dirs: List[str] = field(default_factory=list)
    file_pattern: str = ""

    def __init__(
        self,
        l1_keyword: str = "",
        l2_keyword: str = "",
        file_pattern: str = "",
        dirs: Optional[List[str]] = None
    ):
        if dirs is not None:
            self.dirs = list(dirs)
        elif l1_keyword or l2_keyword:
            self.dirs = [l1_keyword, l2_keyword]
        else:
            self.dirs = []
        self.file_pattern = file_pattern

    @property
    def l1_keyword(self) -> str:
        return self.dirs[0] if len(self.dirs) > 0 else ""

    @l1_keyword.setter
    def l1_keyword(self, val: str):
        while len(self.dirs) < 1:
            self.dirs.append("")
        self.dirs[0] = val

    @property
    def l2_keyword(self) -> str:
        return self.dirs[1] if len(self.dirs) > 1 else ""

    @l2_keyword.setter
    def l2_keyword(self, val: str):
        while len(self.dirs) < 2:
            self.dirs.append("")
        self.dirs[1] = val

    def is_completely_empty(self) -> bool:
        """檢查該列是否所有目錄與檔案欄位完全為空白"""
        all_empty_dirs = all(not d.strip() for d in self.dirs) if self.dirs else True
        return all_empty_dirs and not self.file_pattern.strip()


@dataclass
class SearchSettings:
    """全域搜尋與儲存設定"""
    global_pattern: str = ""
    conflict_strategy: ConflictStrategy = ConflictStrategy.SIZE_MTIME
    local_save_path: str = ""
    match_mode: MatchMode = MatchMode.FUZZY
    request_delay_ms: int = 50          # 下載間隔微延遲 (毫秒)，預設 50ms 溫和防封
    session_refresh_interval: int = 150 # 主動換氣間隔 (每 N 個檔案主動重連一次，避開防火牆限制)


@dataclass
class RemoteFileTask:
    """代表經過多層比對後，確認要處理的遠端檔案任務"""
    dirs: List[str] = field(default_factory=list)
    filename: str = ""
    remote_full_path: str = ""
    local_full_path: str = ""
    file_size: int = 0
    file_mtime: int = 0

    def __init__(
        self,
        l1_dir: str = "",
        l2_dir: str = "",
        filename: str = "",
        remote_full_path: str = "",
        local_full_path: str = "",
        file_size: int = 0,
        file_mtime: int = 0,
        dirs: Optional[List[str]] = None
    ):
        if dirs is not None:
            self.dirs = list(dirs)
        elif l1_dir or l2_dir:
            self.dirs = [l1_dir, l2_dir]
        else:
            self.dirs = []
        self.filename = filename
        self.remote_full_path = remote_full_path
        self.local_full_path = local_full_path
        self.file_size = file_size
        self.file_mtime = file_mtime

    @property
    def l1_dir(self) -> str:
        return self.dirs[0] if len(self.dirs) > 0 else ""

    @property
    def l2_dir(self) -> str:
        return self.dirs[1] if len(self.dirs) > 1 else ""


@dataclass
class TaskStats:
    """任務即時統計數據"""
    scanned_dirs: int = 0
    matched_files: int = 0
    downloaded_ok: int = 0
    skipped_files: int = 0
    failed_files: int = 0


@dataclass
class ReportRecord:
    """單筆檔案執行紀錄，包含完整欄位資訊供產生 Manifest CSV"""
    remote_path: str
    local_path: str
    dirs: List[str] = field(default_factory=list)
    filename: str = ""
    file_size: int = 0
    file_mtime: int = 0
    status: TransferStatus = TransferStatus.SUCCESS
    message: str = ""
    timestamp: str = ""

    def __init__(
        self,
        remote_path: str,
        local_path: str,
        l1_dir: str = "",
        l2_dir: str = "",
        filename: str = "",
        file_size: int = 0,
        status: TransferStatus = TransferStatus.SUCCESS,
        message: str = "",
        timestamp: str = "",
        dirs: Optional[List[str]] = None,
        file_mtime: int = 0
    ):
        self.remote_path = remote_path
        self.local_path = local_path
        if dirs is not None:
            self.dirs = list(dirs)
        elif l1_dir or l2_dir:
            self.dirs = [l1_dir, l2_dir]
        else:
            self.dirs = []
        self.filename = filename
        self.file_size = file_size
        self.file_mtime = file_mtime
        self.status = status
        self.message = message
        self.timestamp = timestamp

    @property
    def l1_dir(self) -> str:
        return self.dirs[0] if len(self.dirs) > 0 else ""

    @property
    def l2_dir(self) -> str:
        return self.dirs[1] if len(self.dirs) > 1 else ""

    @property
    def readable_size(self) -> str:
        """將 Bytes 轉換為人類易讀格式 (B, KB, MB, GB)"""
        size = float(self.file_size)
        for unit in ["B", "KB", "MB", "GB", "TB"]:
            if size < 1024.0 or unit == "TB":
                if unit == "B":
                    return f"{int(size)} B"
                return f"{size:.2f} {unit}"
            size /= 1024.0
        return f"{self.file_size} B"

    @property
    def mtime_str(self) -> str:
        """將 Unix Timestamp 轉換為 YYYY-MM-DD HH:MM:SS 易讀格式"""
        if not self.file_mtime or self.file_mtime <= 0:
            return ""
        try:
            return datetime.datetime.fromtimestamp(self.file_mtime).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return ""

    @property
    def file_ext(self) -> str:
        """提取檔案小寫副檔名（含小數點，例如 .csv），若無檔名則為空字串"""
        if not self.filename:
            return ""
        _, ext = os.path.splitext(self.filename)
        return ext.lower()
