# core/detail_aggregator.py
import csv
import os
import tempfile
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple
import xlsxwriter

from core.db_connector import DatabaseConnector, parse_coord_int
from core.excel_utils import safe_atomic_move, sanitize_xml_text
from core.sidecar_cache import SidecarCacheWriter


class DetailAggregator:
    """
    彙整資料 1（極速雙軌串流版）：
    - 管道 A：xlsxwriter(constant_memory=True) 寫入 %TEMP% 暫存檔，結束後安全原子搬移
    - 管道 B：同步每 50,000 列寫入 SQLite 旁路快取 (.cache.db)，供彙整 2 秒級讀取
    - 整合 MS SQL 資料庫撈取 PanelID 與 Scrapped 注入明細第 K, L 欄
    - 完美相容既有介面簽名：aggregate_to_excel(...) -> (succeeded, failed, total_written)
    """

    OUT_SHEET_NAME = "工作表2"
    MAX_ROWS_PER_SHEET = 1_000_000  # 含表頭列
    CACHE_BATCH_SIZE = 50_000

    DETAIL_HEADERS: List[str] = [
        "第 1 層目錄",
        "PN",
        "第 2 層目錄",
        "LOT",
        "RECIPE",
        "檔案名稱",
        "2D ID (MPID)",
        "X (MP_X)",
        "Y (MP_Y)",
        "MPID_X_Y",
        "PanelID",
        "Scrapped",
        "Unit BIN#",
    ]

    # A~L 欄寬度 (索引 0~11 鎖定純文字 '@')，M 欄 (索引 12) 寬度 13 保留 General/數值格式
    COLUMN_WIDTHS: Dict[str, int] = {
        "A": 24, "B": 13, "C": 22, "D": 22, "E": 12,
        "F": 62, "G": 21, "H": 13, "I": 13, "J": 28,
        "K": 24, "L": 14, "M": 13
    }
    DETAIL_COL_WIDTHS_LIST: Tuple[int, ...] = (24, 13, 22, 22, 12, 62, 21, 13, 13, 28, 24, 14, 13)

    @staticmethod
    def parse_recipe(level1_dir: str, pn: str = "") -> str:
        """
        RECIPE 解析演算法 (方案 2: 以 PN 為錨點自動定位 + 連字號自適應退避):
        1. 優先方案: 若有提供 pn 且存在於 level1_dir:
           - 找到 pn 在 level1_dir 中的位置
           - 從 pn 之後尋找第一個連字號 '-'
           - 取該連字號之後的全部文字作為 RECIPE (例如 '214-3273242A-SUB-C' 配 '3273242' -> 'SUB-C')
           - 亦涵蓋無廠區前綴結構 (例如 '3273173A-DUMMY' 配 '3273173' -> 'DUMMY')
        2. 退避方案: 若無 pn 或 pn 未在目錄名中命中:
           - 若有 >= 2 個連字號: 取第 2 個連字號後的全部文字
           - 若只有 1 個連字號: 取第 1 個連字號後的全部文字
           - 若無連字號: 回傳空字串
        """
        s = level1_dir.strip()
        p = pn.strip()

        # 優先方案: 以 PN 為錨點
        if p and p.lower() in s.lower():
            pn_pos = s.lower().find(p.lower())
            dash_pos = s.find("-", pn_pos + len(p))
            if dash_pos != -1:
                recipe = s[dash_pos + 1:].strip()
                if recipe:
                    return recipe

        # 退避方案: 連字號數量自適應
        dashes = [i for i, ch in enumerate(s) if ch == "-"]
        if len(dashes) >= 2:
            return s[dashes[1] + 1:].strip()
        elif len(dashes) == 1:
            return s[dashes[0] + 1:].strip()
        return ""

    @staticmethod
    def read_local_text(file_path: str) -> str:
        """支援 UTF-8 BOM、標準 UTF-8 及 Windows CP950/ANSI 編碼自動辨識讀取"""
        with open(file_path, "rb") as f:
            raw_bytes = f.read()

        if not raw_bytes:
            return ""

        # 1. 檢查 UTF-8 BOM
        if raw_bytes.startswith(b"\xef\xbb\xbf"):
            return raw_bytes[3:].decode("utf-8", errors="replace")

        # 2. 嘗試標準 UTF-8
        try:
            return raw_bytes.decode("utf-8")
        except UnicodeDecodeError:
            pass

        # 3. 嘗試 Windows 繁體中文 CP950 (ANSI)
        try:
            return raw_bytes.decode("cp950")
        except UnicodeDecodeError:
            pass

        # 4. 保底 Latin-1
        return raw_bytes.decode("latin-1", errors="replace")

    @classmethod
    def parse_single_file(
        cls,
        file_path: str,
        level1: str,
        pn: str,
        level2: str,
        lot: str,
        db_map: Optional[Dict[Tuple[str, str, int, int], Tuple[str, str]]] = None,
    ) -> Tuple[List[List[Any]], Optional[str]]:
        """
        解析單一晶片測試檔案：
        - 搜尋前 10 列中包含 2D ID, X, Y, UNIT BIN# 之標題列（大小寫不敏感，相容逗號與 Tab）
        - 第一欄為 Hard Limit 之列略過不計
        - RECIPE 依 PN 錨點解析
        - 若提供 db_map，依 (clean_lot, unit_id, x_int, y_int) 對齊填入 PanelID 與 Scrapped
        """
        if not os.path.exists(file_path):
            return [], f"實體檔案不存在: {file_path}"

        raw_text = cls.read_local_text(file_path)
        if not raw_text.strip():
            return [], "空檔案"

        lines = [line.strip("\r\n") for line in raw_text.splitlines() if line.strip("\r\n")]
        if not lines:
            return [], "無有效內容"

        # 尋找前 10 列中的標題列
        header_idx = -1
        sep = ","
        id_col = -1
        x_col = -1
        y_col = -1
        bin_col = -1

        for i, line in enumerate(lines[:10]):
            line_upper = line.upper()
            if "2D ID" in line_upper:
                sep = "\t" if "\t" in line else ","
                reader = csv.reader([line], delimiter=sep)
                fields = next(reader, [])
                t_id, t_x, t_y, t_bin = -1, -1, -1, -1
                for k, col in enumerate(fields):
                    c_clean = col.strip().upper()
                    if c_clean == "2D ID":
                        t_id = k
                    elif c_clean == "X":
                        t_x = k
                    elif c_clean == "Y":
                        t_y = k
                    elif c_clean in ("UNIT BIN#", "UNIT BIN #", "BIN#"):
                        t_bin = k
                if t_id >= 0 and t_x >= 0 and t_y >= 0 and t_bin >= 0:
                    header_idx = i
                    id_col, x_col, y_col, bin_col = t_id, t_x, t_y, t_bin
                    break

        if header_idx < 0:
            return [], "前 10 列找不到 2D ID、X、Y、Unit BIN# 欄位"

        file_name = os.path.basename(file_path)
        sub_code = cls.parse_recipe(level1, pn)
        max_col_needed = max(id_col, x_col, y_col, bin_col)

        clean_lot = lot.strip()
        if not clean_lot and level2.strip():
            l2_val = level2.strip()
            clean_lot = l2_val[:-1] if len(l2_val) > 1 else l2_val

        rows: List[List[Any]] = []
        for line in lines[header_idx + 1:]:
            if not line.strip():
                continue
            reader = csv.reader([line], delimiter=sep)
            cols = next(reader, [])
            if not cols:
                continue

            # 略過 Hard Limit 列
            first_field = cols[0].strip()
            if first_field.lower() == "hard limit":
                continue

            if len(cols) > max_col_needed:
                unit_id = cols[id_col].strip()
                x_val = cols[x_col].strip()
                y_val = cols[y_col].strip()
                bin_val = cols[bin_col].strip()

                if unit_id or bin_val:
                    mpid_x_y = f"{unit_id}_{x_val}_{y_val}" if (unit_id or x_val or y_val) else ""
                    final_bin: Any = bin_val
                    if bin_val.isdigit():
                        try:
                            final_bin = int(bin_val)
                        except ValueError:
                            final_bin = bin_val

                    panel_id = ""
                    scrapped = ""
                    if db_map:
                        x_int = parse_coord_int(x_val)
                        y_int = parse_coord_int(y_val)
                        match_val = db_map.get((clean_lot, unit_id, x_int, y_int))
                        if match_val:
                            panel_id, scrapped = match_val

                    row = [
                        level1,
                        pn,
                        level2,
                        lot,
                        sub_code,
                        file_name,
                        unit_id,
                        x_val,
                        y_val,
                        mpid_x_y,
                        panel_id,
                        scrapped,
                        final_bin
                    ]
                    rows.append(row)

        return rows, None

    @classmethod
    def _setup_detail_sheet(
        cls,
        ws: xlsxwriter.worksheet.Worksheet,
        fmt_text: xlsxwriter.format.Format,
        fmt_header_text: xlsxwriter.format.Format,
        fmt_header_general: xlsxwriter.format.Format,
    ) -> None:
        """在寫入資料列前鎖定欄寬、純文字格式 (@) 與表頭列（Row 0）"""
        for col_idx, width in enumerate(cls.DETAIL_COL_WIDTHS_LIST):
            if col_idx <= 11:  # A ~ L 欄 (含 PanelID, Scrapped)：純文字格式 '@'
                ws.set_column(col_idx, col_idx, width, fmt_text)
            else:  # M 欄 (Unit BIN#)：保留 General / 數值格式
                ws.set_column(col_idx, col_idx, width)

        for col_idx, header_text in enumerate(cls.DETAIL_HEADERS):
            hdr_fmt = fmt_header_text if col_idx <= 11 else fmt_header_general
            ws.write_string(0, col_idx, header_text, hdr_fmt)

    @staticmethod
    def _write_bin_cell(
        ws: xlsxwriter.worksheet.Worksheet, row_idx: int, col_idx: int, bin_val: Any
    ) -> None:
        """K 欄 (Unit BIN#) 寫入：數值以 write_number 輸出，字串以 write_string 輸出"""
        if bin_val is None or bin_val == "":
            return
        if isinstance(bin_val, int):
            ws.write_number(row_idx, col_idx, bin_val)
        elif isinstance(bin_val, float):
            if bin_val.is_integer():
                ws.write_number(row_idx, col_idx, int(bin_val))
            else:
                ws.write_number(row_idx, col_idx, bin_val)
        elif isinstance(bin_val, str) and bin_val.isdigit() and (len(bin_val) == 1 or not bin_val.startswith("0")):
            ws.write_number(row_idx, col_idx, int(bin_val))
        else:
            ws.write_string(row_idx, col_idx, str(bin_val))

    @classmethod
    def aggregate_to_excel(
        cls,
        manifest_csv_path: str,
        output_xlsx_path: str,
        log_cb: Optional[Callable[[str], None]] = None,
        progress_cb: Optional[Callable[[int, int, int, str], None]] = None
    ) -> Tuple[int, int, int]:
        """
        讀取下載清單 CSV，解析其所有「下載成功」的檔案，透過雙軌串流產出：
        - 管道 A：xlsxwriter(constant_memory=True) 輸出雙工作表 Excel（Sheet 1 下載清單、Sheet 2 工作表2...）
        - 管道 B：同步產生 SQLite 旁路快取 (.cache.db)，供彙整 2 秒級載入
        回傳: (成功解析檔案數, 略過/失敗檔案數, 總寫入明細筆數)
        """
        if not os.path.exists(manifest_csv_path):
            raise FileNotFoundError(f"找不到下載清單: {manifest_csv_path}")

        # 1. 讀取下載清單 CSV
        with open(manifest_csv_path, "r", encoding="utf-8-sig") as f:
            reader = list(csv.reader(f))

        if not reader:
            raise ValueError("下載清單為空檔案")

        manifest_headers = reader[0]
        manifest_rows = reader[1:]

        col_map = {name.strip(): idx for idx, name in enumerate(manifest_headers)}
        path_col = col_map.get("本地完整路徑")
        status_col = col_map.get("處理狀態")
        l1_col = col_map.get("第 1 層目錄")
        pn_col = col_map.get("第 1 層搜尋字")
        l2_col = col_map.get("第 2 層目錄")
        lot_col = col_map.get("第 2 層搜尋字")

        if path_col is None or status_col is None:
            raise ValueError("下載清單缺少必要欄位：'本地完整路徑' 或 '處理狀態'")

        valid_items = []
        for r_idx, r in enumerate(manifest_rows, start=2):
            local_path = r[path_col].strip() if path_col < len(r) else ""
            status = r[status_col].strip() if status_col < len(r) else ""
            if status == "下載成功" and local_path:
                valid_items.append((r_idx, r, local_path))

        total_files = len(valid_items)
        tmp_dir = tempfile.gettempdir()
        fd, temp_xlsx_path = tempfile.mkstemp(prefix="wafer_detail_", suffix=".xlsx", dir=tmp_dir)
        os.close(fd)

        cache_writer = SidecarCacheWriter(output_xlsx_path)
        cache_buffer: List[Tuple[Any, ...]] = []

        wb = xlsxwriter.Workbook(
            temp_xlsx_path,
            {"constant_memory": True, "tmpdir": tmp_dir},
        )
        wb_closed = False

        try:
            fmt_list_header = wb.add_format({"bold": True})
            fmt_text = wb.add_format({"num_format": "@"})
            fmt_header_text = wb.add_format({"bold": True, "num_format": "@"})
            fmt_header_general = wb.add_format({"bold": True})

            # Sheet 1: 下載清單（表頭加粗、欄寬 18）
            ws_list = wb.add_worksheet("下載清單")
            num_list_cols = max(len(manifest_headers), 1)
            ws_list.set_column(0, num_list_cols - 1, 18)

            for c_idx, h_val in enumerate(manifest_headers):
                ws_list.write_string(0, c_idx, sanitize_xml_text(h_val), fmt_list_header)

            for r_offset, row_vals in enumerate(manifest_rows, start=1):
                for c_idx, val in enumerate(row_vals):
                    if val is not None and val != "":
                        ws_list.write_string(r_offset, c_idx, sanitize_xml_text(val))

            # Sheet 2+: 工作表2 (首頁無後綴，超過 100 萬列切 工作表2_2, 工作表2_3...)
            page_no = 1
            current_ws = wb.add_worksheet(cls.OUT_SHEET_NAME)
            cls._setup_detail_sheet(current_ws, fmt_text, fmt_header_text, fmt_header_general)

            row_in_sheet = 1
            total_written = 0
            succeeded_files = 0
            failed_files = 0

            # 提前收集所有有效母批號，批次向 MS SQL 查詢 UnitBase (PanelID, Scrapped)
            all_mother_lots = set()
            for _, r, _ in valid_items:
                l_val = r[lot_col].strip() if lot_col is not None and lot_col < len(r) else ""
                if not l_val and l2_col is not None and l2_col < len(r):
                    l2_v = r[l2_col].strip()
                    l_val = l2_v[:-1] if len(l2_v) > 1 else l2_v
                if l_val:
                    all_mother_lots.add(l_val)

            db_map = DatabaseConnector.fetch_wafer_unit_base(sorted(all_mother_lots), log_cb=log_cb)

            for file_idx, (r_idx, r, local_path) in enumerate(valid_items, start=1):
                l1 = r[l1_col].strip() if l1_col is not None and l1_col < len(r) else ""
                pn = r[pn_col].strip() if pn_col is not None and pn_col < len(r) else ""
                l2 = r[l2_col].strip() if l2_col is not None and l2_col < len(r) else ""
                lot = r[lot_col].strip() if lot_col is not None and lot_col < len(r) else ""

                fname = os.path.basename(local_path)
                parsed_rows, err = cls.parse_single_file(local_path, l1, pn, l2, lot, db_map=db_map)
                if err:
                    failed_files += 1
                    if log_cb:
                        log_cb(f"[略過] [{file_idx}/{total_files}] 第 {r_idx} 列 {fname}: {err}")
                    if progress_cb:
                        progress_cb(file_idx, total_files, total_written, fname)
                    continue

                succeeded_files += 1

                for row_data in parsed_rows:
                    if row_in_sheet >= cls.MAX_ROWS_PER_SHEET:
                        page_no += 1
                        sheet_name = f"{cls.OUT_SHEET_NAME}_{page_no}"
                        current_ws = wb.add_worksheet(sheet_name)
                        cls._setup_detail_sheet(
                            current_ws, fmt_text, fmt_header_text, fmt_header_general
                        )
                        row_in_sheet = 1

                    # A ~ L 欄 (索引 0~11，含 PanelID, Scrapped) 寫入純文字
                    for c_idx in range(12):
                        v_str = sanitize_xml_text(row_data[c_idx])
                        if v_str:
                            current_ws.write_string(row_in_sheet, c_idx, v_str)

                    # M 欄 (Unit BIN#，索引 12)
                    cls._write_bin_cell(current_ws, row_in_sheet, 12, row_data[12])

                    row_in_sheet += 1
                    total_written += 1

                    # 旁路快取累積寫入
                    cache_buffer.append(tuple(sanitize_xml_text(col) for col in row_data))
                    if len(cache_buffer) >= cls.CACHE_BATCH_SIZE:
                        cache_writer.write_batch(cache_buffer)
                        cache_buffer.clear()

                if progress_cb:
                    progress_cb(file_idx, total_files, total_written, fname)

                if log_cb and (file_idx == 1 or file_idx % 5 == 0 or file_idx == total_files):
                    pct = (file_idx / total_files) * 100 if total_files > 0 else 100.0
                    log_cb(f"[{file_idx}/{total_files}] ({pct:.1f}%) 正在解析: {fname} (累計寫入 {total_written:,} 筆晶片)...")

            # 沖刷剩餘快取
            if cache_buffer:
                cache_writer.write_batch(cache_buffer)
                cache_buffer.clear()

            if log_cb:
                log_cb(f"所有檔案解析完成 (共 {succeeded_files} 個成功檔，寫入 {total_written:,} 筆晶片)，正在封裝 Excel 活頁簿至硬碟 (請稍候)...")

            wb.close()
            wb_closed = True

            safe_atomic_move(temp_xlsx_path, output_xlsx_path)
            cache_writer.finalize(output_xlsx_path)

            return succeeded_files, failed_files, total_written

        except Exception:
            if not wb_closed:
                try:
                    wb.close()
                except Exception:
                    pass
            if os.path.exists(temp_xlsx_path):
                try:
                    os.remove(temp_xlsx_path)
                except OSError:
                    pass
            cache_writer.abort()
            raise
