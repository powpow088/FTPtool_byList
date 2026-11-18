import fnmatch
import os
import posixpath
import socket
import stat
import threading
import time
from collections import defaultdict
from typing import List, Dict, Set, Callable, Tuple, Optional
import paramiko

from models.data_models import (
    RuleItem, SearchSettings, RemoteFileTask, TaskStats, LogLevel, TaskCancelledError, MatchMode,
    ReportRecord, TransferStatus
)
from core.sftp_manager import SFTPManager


class TraverseEngine:
    """
    負責將純目錄規則編譯並透過實體目錄快取機制執行動態 N 階層遍歷。
    核心保證：
    1. 遠端任何一個實體目錄最多只會被呼叫 1 次 listdir_attr()。
    2. 嚴格鎖定列對列（Row-by-Row）成對目錄關聯性，避免笛卡兒積交叉污染。
    3. 抵達最後一層目錄後，自動依據目標檔案條件 (如 *RAW*) 搜尋檔案。
    4. 支援萬用字元 (*, ?) 與模糊包含 / 完全精準模式切換。
    """

    MAX_RETRY = 3

    def __init__(
        self,
        sftp_mgr: SFTPManager,
        initial_path: str,
        search_settings: SearchSettings,
        cancel_event: threading.Event,
        log_cb: Callable[[LogLevel, str], None],
        stats_cb: Callable[[TaskStats], None]
    ) -> None:
        self._sftp = sftp_mgr
        self._initial_path = initial_path.rstrip("/") or "/"
        self._settings = search_settings
        self._cancel_event = cancel_event
        self._log = log_cb
        self._stats_cb = stats_cb
        self.empty_records: List[ReportRecord] = []

    @staticmethod
    def compile_rule_tree(
        rules: List[RuleItem],
        global_file_pattern: str
    ) -> Dict[str, Dict[str, Set[str]]]:
        """
        [相容性方法] 將表格列預先整合為傳統三層規則樹：
        Tree[l1_pattern][l2_pattern] = Set[file_pattern]
        """
        tree: Dict[str, Dict[str, Set[str]]] = defaultdict(lambda: defaultdict(set))
        global_pat = global_file_pattern.strip()

        valid_rules = [r for r in rules if not r.is_completely_empty()]
        if not valid_rules:
            valid_rules = [RuleItem(dirs=["", ""], file_pattern=global_pat)]

        for r in valid_rules:
            l1 = r.l1_keyword.strip()
            l2 = r.l2_keyword.strip()
            f_pat = r.file_pattern.strip() if r.file_pattern.strip() else global_pat
            tree[l1][l2].add(f_pat)

        return tree

    @staticmethod
    def match_pattern(
        target_name: str,
        rule_expr: str,
        match_mode: MatchMode = MatchMode.FUZZY
    ) -> bool:
        """
        核心萬用字元與比對演算法：
        - 留空代表全選 (True)
        - 支援 ';' 或 ',' 分隔多組條件 (OR 邏輯)
        - 含 '*' 或 '?' 一律優先呼叫 fnmatchcase 進行萬用字元比對
        - 無萬用字元時：
            - MatchMode.EXACT: target == expr (完全精準相符)
            - MatchMode.FUZZY: expr in target (模糊包含)
        - 全程大小寫不敏感 (Case-Insensitive)
        """
        expr = rule_expr.strip()
        if not expr:
            return True

        target_lower = target_name.lower()
        sub_patterns = [p.strip() for p in expr.replace(";", ",").split(",") if p.strip()]
        if not sub_patterns:
            return True

        for pat in sub_patterns:
            pat_lower = pat.lower()
            if "*" in pat_lower or "?" in pat_lower:
                if fnmatch.fnmatchcase(target_lower, pat_lower):
                    return True
            else:
                if match_mode == MatchMode.EXACT:
                    if target_lower == pat_lower:
                        return True
                else:
                    if pat_lower in target_lower:
                        return True
        return False

    def scan_all(self, rules: List[RuleItem], stats: TaskStats) -> List[RemoteFileTask]:
        """執行動態純目錄多層遍歷，並於最後一層目錄檢索目標檔案"""
        self._check_cancel()
        self.empty_records = []
        global_pat = self._settings.global_pattern.strip()
        match_mode = getattr(self._settings, "match_mode", MatchMode.FUZZY)

        valid_rules = [r for r in rules if not r.is_completely_empty()]
        if not valid_rules:
            # 預設兩層全選
            valid_rules = [RuleItem(dirs=["", ""], file_pattern=global_pat)]

        # 正規化每筆規則為 (rule_idx: int, dirs: tuple, file_pattern: str)
        norm_rules: List[Tuple[int, Tuple[str, ...], str]] = []
        for idx, r in enumerate(valid_rules):
            d_tuple = tuple(d.strip() for d in r.dirs)
            f_pat = r.file_pattern.strip() if r.file_pattern.strip() else global_pat
            norm_rules.append((idx, d_tuple, f_pat))

        self._log(LogLevel.INFO, f"開始掃描遠端根目錄: {self._initial_path}")
        mode_desc = "完全精準相符" if match_mode == MatchMode.EXACT else "模糊包含"
        self._log(LogLevel.INFO, f"目錄比對模式: {mode_desc}")
        self._log(LogLevel.INFO, f"目標檔案過濾條件: [{global_pat or '全部檔案'}]")

        active_paths: Dict[Tuple[str, ...], Set[Tuple[int, Tuple[str, ...], str]]] = {
            (): set(norm_rules)
        }

        # 實體目錄 listing 快取，保證任何一個實體目錄最多只會呼叫一次 SFTP
        dir_cache: Dict[str, List[paramiko.SFTPAttributes]] = {}

        def get_entries(rel_phys_dirs: Tuple[str, ...], is_root: bool = False) -> List[paramiko.SFTPAttributes]:
            full_remote = (
                posixpath.join(self._initial_path, *rel_phys_dirs)
                if rel_phys_dirs else self._initial_path
            )
            if full_remote not in dir_cache:
                self._check_cancel()
                entries = self._safe_listdir_attr(full_remote, is_root=is_root)
                stats.scanned_dirs += 1
                self._stats_cb(stats)
                dir_cache[full_remote] = entries
            return dir_cache[full_remote]

        # 逐層向下推進資料夾
        while True:
            self._check_cancel()
            has_more_dirs = any(
                len(rem_dirs) > 0
                for rule_set in active_paths.values()
                for _, rem_dirs, _ in rule_set
            )
            if not has_more_dirs:
                break

            next_active_paths: Dict[Tuple[str, ...], Set[Tuple[int, Tuple[str, ...], str]]] = defaultdict(set)

            for phys_tuple, rule_set in active_paths.items():
                to_advance: Dict[str, Set[Tuple[int, Tuple[str, ...], str]]] = defaultdict(set)
                for r_idx, rem_dirs, f_pat in rule_set:
                    if len(rem_dirs) > 0:
                        next_dir_rule = rem_dirs[0]
                        rem_rest = rem_dirs[1:]
                        to_advance[next_dir_rule].add((r_idx, rem_rest, f_pat))
                    else:
                        next_active_paths[phys_tuple].add((r_idx, (), f_pat))

                if not to_advance:
                    continue

                entries = get_entries(phys_tuple, is_root=(len(phys_tuple) == 0))
                sub_dirs = [e.filename for e in entries if stat.S_ISDIR(e.st_mode or 0)]

                for phys_sub in sub_dirs:
                    for rule_dir_pat, rule_tails in to_advance.items():
                        if self.match_pattern(phys_sub, rule_dir_pat, match_mode):
                            next_tuple = phys_tuple + (phys_sub,)
                            next_active_paths[next_tuple].update(rule_tails)

            active_paths = next_active_paths
            current_depth = max((len(p) for p in active_paths.keys()), default=0)
            self._log(LogLevel.INFO, f"第 {current_depth} 層目錄比對完成，命中 {len(active_paths)} 個有效路徑。")

        # 檢查是否有未匹配到任何實體目錄的規則 (目錄不存在)
        reached_rule_indices = {r_idx for rule_set in active_paths.values() for r_idx, _, _ in rule_set}
        for r_idx, d_tuple, f_pat in norm_rules:
            if r_idx not in reached_rule_indices:
                remote_dir = posixpath.join(self._initial_path, *d_tuple) if d_tuple else self._initial_path
                local_dir = os.path.join(self._settings.local_save_path, *d_tuple) if d_tuple else self._settings.local_save_path
                self.empty_records.append(
                    ReportRecord(
                        remote_path=remote_dir,
                        local_path=local_dir,
                        dirs=list(d_tuple),
                        filename="(目錄不存在)",
                        file_size=0,
                        file_mtime=0,
                        status=TransferStatus.DIR_NOT_FOUND,
                        message="遠端未找到符合此條件之資料夾"
                    )
                )

        # 最終階段：在所有終端實體目錄（最後一層目錄）中，檢索符合條件之目標檔案
        tasks: List[RemoteFileTask] = []
        seen_remote_files: Set[str] = set()

        for phys_tuple, rule_set in active_paths.items():
            self._check_cancel()
            file_patterns = {f_pat for _, _, f_pat in rule_set}
            if not file_patterns:
                file_patterns = {global_pat}

            matched_in_dir = 0
            entries = get_entries(phys_tuple, is_root=(len(phys_tuple) == 0))
            for entry in entries:
                if not stat.S_ISREG(entry.st_mode or 0):
                    continue
                fname = entry.filename
                # 在最後一層目錄中檢索檔案
                if any(self.match_pattern(fname, fp, match_mode=MatchMode.FUZZY) for fp in file_patterns):
                    full_remote = (
                        posixpath.join(self._initial_path, *phys_tuple, fname)
                        if phys_tuple else posixpath.join(self._initial_path, fname)
                    )
                    if full_remote in seen_remote_files:
                        continue
                    seen_remote_files.add(full_remote)
                    matched_in_dir += 1

                    local_file_path = (
                        os.path.join(self._settings.local_save_path, *phys_tuple, fname)
                        if phys_tuple else os.path.join(self._settings.local_save_path, fname)
                    )

                    tasks.append(
                        RemoteFileTask(
                            dirs=list(phys_tuple),
                            filename=fname,
                            remote_full_path=full_remote,
                            local_full_path=local_file_path,
                            file_size=int(entry.st_size or 0),
                            file_mtime=int(entry.st_mtime or int(time.time()))
                        )
                    )
                    stats.matched_files += 1

            if matched_in_dir == 0:
                # 該終端實體目錄底下無相符檔案
                remote_dir = posixpath.join(self._initial_path, *phys_tuple) if phys_tuple else self._initial_path
                local_dir = os.path.join(self._settings.local_save_path, *phys_tuple) if phys_tuple else self._settings.local_save_path
                self.empty_records.append(
                    ReportRecord(
                        remote_path=remote_dir,
                        local_path=local_dir,
                        dirs=list(phys_tuple),
                        filename="(無相符檔案)",
                        file_size=0,
                        file_mtime=0,
                        status=TransferStatus.EMPTY_DIR,
                        message="該目錄下無相符之目標檔案"
                    )
                )

            self._stats_cb(stats)

        self._log(LogLevel.SUCCESS, f"目錄與檔案檢索完成，共篩選出 {len(tasks)} 個目標檔案。")
        return tasks

    def _safe_listdir_attr(
        self, remote_path: str, is_root: bool = False
    ) -> List[paramiko.SFTPAttributes]:
        """帶有權限攔截與斷線重試的目錄列舉封裝"""
        for attempt in range(1, self.MAX_RETRY + 1):
            self._check_cancel()
            try:
                return self._sftp.listdir_attr(remote_path)
            except PermissionError:
                self._log(LogLevel.ERROR, f"權限不足 (Permission Denied)，略過目錄: {remote_path}")
                return []
            except FileNotFoundError:
                if is_root:
                    raise
                self._log(LogLevel.WARNING, f"遠端目錄不存在，略過: {remote_path}")
                return []
            except (socket.timeout, paramiko.SSHException, EOFError) as e:
                if attempt < self.MAX_RETRY:
                    wait_sec = 2 ** (attempt - 1)
                    self._log(
                        LogLevel.WARNING,
                        f"讀取目錄網路異常 ({remote_path}): {e}，{wait_sec} 秒後重試 ({attempt}/{self.MAX_RETRY})..."
                    )
                    time.sleep(wait_sec)
                    try:
                        self._sftp.reconnect()
                    except Exception as reconn_err:
                        self._log(LogLevel.WARNING, f"重連嘗試失敗: {reconn_err}")
                else:
                    self._log(LogLevel.ERROR, f"目錄重試 {self.MAX_RETRY} 次仍失敗，略过: {remote_path}")
                    return []
        return []

    def _check_cancel(self) -> None:
        """主動檢查使用者是否已發起中斷"""
        if self._cancel_event.is_set():
            raise TaskCancelledError("使用者主動中止掃描")
