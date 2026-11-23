# core/db_connector.py
from typing import Any, Callable, Dict, List, Optional, Sequence, Set, Tuple
import pymssql

from core.db_config import DB_CONFIG, is_db_configured


def parse_coord_int(val: Any) -> int:
    """座標安全轉為純整數 int，消除前導零或小數位 (例如 '01', 1.0, 1 -> 1)"""
    if val is None:
        return 0
    s = str(val).strip()
    if not s:
        return 0
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return 0


class DatabaseConnector:
    """負責與 MS SQL Server 進行高效率批次查詢與晶片屬性 (PanelID, Scrapped) 記憶體映射"""

    @classmethod
    def fetch_wafer_unit_base(
        cls,
        mother_lots: Sequence[str],
        log_cb: Optional[Callable[[str], None]] = None
    ) -> Dict[Tuple[str, str, int, int], Tuple[str, str]]:
        """
        以乾淨母批號清單為條件，一次性向 [YM_ULT_UnitBase] 批次撈取晶片資訊。
        回傳: {(LotNum, MPID, MP_X_int, MP_Y_int): (PanelID_串接字串, Scrapped_串接字串)}
        若資料庫未設定或連線失敗，自動降級回傳空字典，確保報表正常產出。
        """
        # 1. 檢查是否配置資料庫
        if not is_db_configured():
            if log_cb:
                log_cb("[提示] 資料庫連線未配置或未啟用，略過 PanelID 與 Scrapped 欄位查詢。")
            return {}

        # 2. 過濾有效母批號清單 (去重)
        valid_lots = sorted({str(lot).strip() for lot in mother_lots if str(lot).strip()})
        if not valid_lots:
            return {}

        if log_cb:
            log_cb(f"正在連線 MS SQL Server 資料庫，批次查詢 {len(valid_lots)} 個母批號之 PanelID 與 Scrapped...")

        # 暫存結構: key -> {"panels": list, "panel_set": set, "scraps": list, "scrap_set": set}
        raw_map: Dict[Tuple[str, str, int, int], Dict[str, Any]] = {}

        # 3. 分批查詢 (每批 500 個 LOT，防止超過 SQL Server 2,100 個參數上限)
        batch_size = 500
        table_name = DB_CONFIG.get("table_name", "[dc].[dbo].[YM_ULT_UnitBase]")

        conn = None
        try:
            conn = pymssql.connect(
                server=DB_CONFIG["server"],
                port=DB_CONFIG["port"],
                user=DB_CONFIG["user"],
                password=DB_CONFIG["password"],
                database=DB_CONFIG["database"],
                login_timeout=DB_CONFIG.get("timeout", 15),
                timeout=DB_CONFIG.get("timeout", 30),
                charset="UTF-8",
            )
            cursor = conn.cursor()

            total_queried_rows = 0

            for i in range(0, len(valid_lots), batch_size):
                lot_chunk = valid_lots[i : i + batch_size]
                placeholders = ",".join(["%s"] * len(lot_chunk))
                sql = (
                    f"SELECT [LotNum], [MPID], [MP_X], [MP_Y], [PanelID], [Scrapped] "
                    f"FROM {table_name} "
                    f"WHERE [LotNum] IN ({placeholders})"
                )

                cursor.execute(sql, tuple(lot_chunk))
                rows = cursor.fetchall()
                total_queried_rows += len(rows)

                for r in rows:
                    lot_num = str(r[0]).strip() if r[0] is not None else ""
                    mpid = str(r[1]).strip() if r[1] is not None else ""
                    x_int = parse_coord_int(r[2])
                    y_int = parse_coord_int(r[3])

                    panel_id = str(r[4]).strip() if r[4] is not None else ""
                    scrapped = str(r[5]).strip() if r[5] is not None else ""

                    if not lot_num or not mpid:
                        continue

                    key = (lot_num, mpid, x_int, y_int)
                    if key not in raw_map:
                        raw_map[key] = {
                            "panels": [],
                            "panel_set": set(),
                            "scraps": [],
                            "scrap_set": set(),
                        }
                    item = raw_map[key]

                    # 有序去重累積 PanelID
                    if panel_id and panel_id not in item["panel_set"]:
                        item["panels"].append(panel_id)
                        item["panel_set"].add(panel_id)

                    # 有序去重累積 Scrapped
                    if scrapped and scrapped not in item["scrap_set"]:
                        item["scraps"].append(scrapped)
                        item["scrap_set"].add(scrapped)

            cursor.close()

            if log_cb:
                log_cb(f"資料庫查詢完成！共讀取 {total_queried_rows:,} 筆記錄，成功聚合 {len(raw_map):,} 個晶片鍵值。")

        except Exception as e:
            if log_cb:
                log_cb(f"[警告] 資料庫查詢異常: {e}，跳過 PanelID 與 Scrapped 比對。")
            return {}
        finally:
            if conn:
                try:
                    conn.close()
                except Exception:
                    pass

        # 4. 轉換為最終字串元組字典: key -> (panel_id_joined, scrapped_joined)
        final_result: Dict[Tuple[str, str, int, int], Tuple[str, str]] = {}
        for key, item in raw_map.items():
            panel_str = ",".join(item["panels"])
            scrap_str = ",".join(item["scraps"])
            final_result[key] = (panel_str, scrap_str)

        return final_result
