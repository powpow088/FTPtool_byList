# core/sidecar_cache.py
import os
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterator, List, Optional, Sequence, Tuple, Union


def get_cache_path(xlsx_path: Union[str, Path]) -> Path:
    """取得對應 Excel 檔案的隱藏旁路快取路徑 (.檔名.xlsx.cache.db)"""
    p = Path(xlsx_path)
    return p.with_name(f".{p.name}.cache.db")


class SidecarCacheWriter:
    """
    彙整 1 專用：極速批次寫入 SQLite 旁路快取，並在完成後寫入主檔指紋。
    若寫入過程發生任何異常，自動 rollback、關閉連線並清除未完成的殘留快取檔。
    """

    DDL_STATEMENTS = (
        """
        CREATE TABLE IF NOT EXISTS meta (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            file_size INTEGER NOT NULL,
            mtime_ns INTEGER NOT NULL,
            row_count INTEGER NOT NULL,
            created_at TEXT NOT NULL
        );
        """,
        """
        CREATE TABLE IF NOT EXISTS wafer_details (
            layer1 TEXT,
            pn TEXT,
            layer2 TEXT,
            lot TEXT,
            recipe TEXT,
            file_name TEXT,
            mpid TEXT,
            mp_x TEXT,
            mp_y TEXT,
            mpid_x_y TEXT,
            unit_bin TEXT
        );
        """,
    )

    INSERT_SQL = (
        "INSERT INTO wafer_details "
        "(layer1, pn, layer2, lot, recipe, file_name, mpid, mp_x, mp_y, mpid_x_y, unit_bin) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    )

    def __init__(self, target_xlsx_path: Union[str, Path]):
        self.cache_path = get_cache_path(target_xlsx_path)
        self._cleanup_files()

        if self.cache_path.parent:
            os.makedirs(self.cache_path.parent, exist_ok=True)

        self.conn: Optional[sqlite3.Connection] = sqlite3.connect(
            str(self.cache_path), isolation_level=None
        )
        self.total_rows = 0
        self._closed = False

        try:
            self._configure_fast_pragmas()
            self.conn.execute("BEGIN TRANSACTION;")
            for ddl in self.DDL_STATEMENTS:
                self.conn.execute(ddl)
            self.conn.execute("COMMIT;")
        except Exception:
            self.abort()
            raise

    def _configure_fast_pragmas(self) -> None:
        assert self.conn is not None
        self.conn.execute("PRAGMA journal_mode = WAL;")
        self.conn.execute("PRAGMA synchronous = OFF;")
        self.conn.execute("PRAGMA cache_size = 100000;")
        self.conn.execute("PRAGMA temp_store = MEMORY;")
        self.conn.execute("PRAGMA locking_mode = EXCLUSIVE;")

    def write_batch(self, records: Sequence[Tuple[Any, ...]]) -> None:
        """批次寫入 11 欄晶片明細資料（建議每批 50,000 列）"""
        if not records:
            return
        if self._closed or self.conn is None:
            raise RuntimeError("SidecarCacheWriter 連線已關閉，無法繼續寫入。")

        try:
            self.conn.execute("BEGIN TRANSACTION;")
            self.conn.executemany(self.INSERT_SQL, records)
            self.conn.execute("COMMIT;")
            self.total_rows += len(records)
        except Exception:
            self.abort()
            raise

    def finalize(self, final_xlsx_path: Union[str, Path]) -> Path:
        """
        寫入最終 .xlsx 的 file_size 與 st_mtime_ns 指紋，並將 WAL 合併封口後關閉連線。
        注意：必須在 final_xlsx_path 已完成 wb.close() 並搬移到位後呼叫。
        """
        if self._closed or self.conn is None:
            raise RuntimeError("SidecarCacheWriter 連線已關閉，無法執行 finalize。")

        try:
            stat = os.stat(final_xlsx_path)
            created_at = datetime.now(timezone.utc).isoformat()

            self.conn.execute("BEGIN TRANSACTION;")
            self.conn.execute(
                """
                INSERT OR REPLACE INTO meta (id, file_size, mtime_ns, row_count, created_at)
                VALUES (1, ?, ?, ?, ?)
                """,
                (stat.st_size, stat.st_mtime_ns, self.total_rows, created_at),
            )
            self.conn.execute("COMMIT;")

            # 將 WAL 日誌合併回主資料庫並切回 DELETE 模式，避免 Windows 目錄殘留 -wal/-shm 檔
            self.conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
            self.conn.execute("PRAGMA journal_mode = DELETE;")
            self.conn.close()
            self._closed = True
            return self.cache_path
        except Exception:
            self.abort()
            raise

    def abort(self) -> None:
        """異常發生時安全回滾、釋放控制代碼並刪除殘留的髒快取檔案"""
        if self.conn is not None:
            try:
                if self.conn.in_transaction:
                    self.conn.rollback()
            except sqlite3.Error:
                pass
            try:
                self.conn.close()
            except sqlite3.Error:
                pass
            self.conn = None
        self._closed = True
        self._cleanup_files()

    def _cleanup_files(self) -> None:
        for suffix in ("", "-wal", "-shm", "-journal"):
            p = Path(f"{self.cache_path}{suffix}")
            try:
                if p.exists():
                    p.unlink()
            except OSError:
                pass

    def __enter__(self) -> "SidecarCacheWriter":
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        if exc_type is not None:
            self.abort()
        elif not self._closed:
            self.abort()


class SidecarCacheReader:
    """彙整 2 專用：校驗旁路快取有效性、秒撈不重複 RECIPE 與串流讀取晶片資料"""

    @staticmethod
    def is_valid(xlsx_path: Union[str, Path], cache_path: Optional[Union[str, Path]] = None) -> bool:
        xlsx_p = Path(xlsx_path)
        cache_p = Path(cache_path) if cache_path else get_cache_path(xlsx_p)

        if not xlsx_p.exists() or not cache_p.exists():
            return False

        try:
            stat = os.stat(xlsx_p)
            conn = sqlite3.connect(f"file:{cache_p}?mode=ro", uri=True)
            try:
                row = conn.execute(
                    "SELECT file_size, mtime_ns FROM meta WHERE id = 1"
                ).fetchone()
                return bool(
                    row is not None
                    and int(row[0]) == stat.st_size
                    and int(row[1]) == stat.st_mtime_ns
                )
            finally:
                conn.close()
        except (OSError, sqlite3.Error):
            return False

    @staticmethod
    def get_distinct_recipes(cache_path: Union[str, Path]) -> List[str]:
        """在 0.05 秒內透過 SQL 撈出所有唯一的非空 RECIPE 清單（排序確保決定性）"""
        conn = sqlite3.connect(f"file:{cache_path}?mode=ro", uri=True)
        try:
            rows = conn.execute("SELECT DISTINCT recipe FROM wafer_details").fetchall()
            recipes = sorted({str(r[0]).strip() for r in rows if r and r[0] is not None and str(r[0]).strip()})
            return recipes
        finally:
            conn.close()

    @staticmethod
    def has_blank_recipe(cache_path: Union[str, Path]) -> bool:
        """快速檢查是否存在空白 RECIPE（決定是否需要建立 ERROR_blank 欄位）"""
        conn = sqlite3.connect(f"file:{cache_path}?mode=ro", uri=True)
        try:
            row = conn.execute(
                "SELECT 1 FROM wafer_details WHERE recipe IS NULL OR TRIM(recipe) = '' LIMIT 1"
            ).fetchone()
            return row is not None
        finally:
            conn.close()

    @staticmethod
    def stream_records(
        cache_path: Union[str, Path], batch_size: int = 50_000
    ) -> Iterator[Tuple[Any, ...]]:
        """以固定批次大小串流讀取 11 欄明細，記憶體維持恆定"""
        conn = sqlite3.connect(f"file:{cache_path}?mode=ro", uri=True)
        try:
            conn.execute("PRAGMA cache_size = 50000;")
            cursor = conn.execute(
                "SELECT layer1, pn, layer2, lot, recipe, file_name, "
                "mpid, mp_x, mp_y, mpid_x_y, unit_bin FROM wafer_details"
            )
            while True:
                batch = cursor.fetchmany(batch_size)
                if not batch:
                    break
                yield from batch
        finally:
            conn.close()
