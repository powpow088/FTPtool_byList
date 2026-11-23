import os
import csv
import datetime
from pathlib import Path
from typing import List, Tuple, Optional, Callable, Dict, Any
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


class DetailAggregator:
    """
    負責穿透已下載之晶片/晶圓測試檔案 (CSV/TSV)，
    依據原 VBA 巨集規範萃取 11 大欄位，並整合下載清單產出為單一雙工作表 Excel 活頁簿 (.xlsx)。
    """

    MAX_ROWS_PER_SHEET = 1_000_000 # 單一工作表上限，超過自動分頁至 工作表2_2, 工作表2_3
    OUT_SHEET_NAME = "工作表2"

    COLUMN_WIDTHS: Dict[str, int] = {
        "A": 24, # 第 1 層目錄
        "B": 13, # PN
        "C": 22, # 第 2 層目錄
        "D": 22, # LOT
        "E": 12, # RECIPE
        "F": 62, # 檔案名稱
        "G": 21, # 2D ID (MPID)
        "H": 13, # X (MP_X)
        "I": 13, # Y (MP_Y)
        "J": 28, # MPID_X_Y
        "K": 13, # Unit BIN#
    }

    HEADERS_DETAIL: List[str] = [
        "第 1 層目錄", "PN", "第 2 層目錄", "LOT", "RECIPE",
        "檔案名稱", "2D ID (MPID)", "X (MP_X)", "Y (MP_Y)",
        "MPID_X_Y", "Unit BIN#"
    ]

    @staticmethod
    def parse_recipe(level1_dir: str, pn: str = "") -> str:
        """
        RECIPE 解析演算法 (方案 2: 以 PN 為錨點自動定位 + 連字號自適應退避):
        1. 優先方案: 若有提供 pn 且存在於 level1_dir:
           - 找到 pn 在 level1_dir 中的位置
           - 從 pn 之後尋找第一個連字號 '-'
           - 若找到，則取該連字號之後的全部文字作為 RECIPE (例如 '214-3273242A-SUB-C' 配 '3273242' -> 'SUB-C')
           - 亦涵蓋無廠區前綴結構 (例如 '3273173A-DUMMY' 配 '3273173' -> 'DUMMY')
        2. 退避方案: 若無 pn 或 pn 未在目錄名中命中:
           - 若有 >= 2 個連字號: 取第 2 個連字號後的全部文字 (例如 'AAA-BBB-SUB-C' -> 'SUB-C')
           - 若只有 1 個連字號: 取第 1 個連字號後的全部文字 (例如 '3273173A-DUMMY' -> 'DUMMY')
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
        """
        支援 UTF-8 BOM、標準 UTF-8 及 Windows CP950/ANSI 編碼自動辨識讀取。
        """
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
        lot: str
    ) -> Tuple[List[List[Any]], Optional[str]]:
        """
        解析單一晶片測試檔案：
        - 搜尋前 10 列中包含 2D ID, X, Y, UNIT BIN# 之標題列（大小寫不敏感，自動相容逗號或 Tab）
        - 略過第一欄為 Hard Limit 之列
        - 組合 MPID_X_Y
        回傳 (明細列清單, 錯誤訊息/略過原因)
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
                # 簡單依分隔符號分割 (相容引號)
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

            if len(cols) > id_col:
                unit_id = cols[id_col].strip()
                if unit_id:
                    if len(cols) <= max_col_needed:
                        # 欄位不足跳過該列
                        continue

                    x_val = cols[x_col].strip()
                    y_val = cols[y_col].strip()
                    bin_val = cols[bin_col].strip()
                    mpid_x_y = f"{unit_id}_{x_val}_{y_val}"

                    # Unit BIN# 數值/文字轉換
                    final_bin: Any = bin_val
                    if bin_val.isdigit() and len(bin_val) <= 9:
                        try:
                            final_bin = int(bin_val)
                        except ValueError:
                            final_bin = bin_val

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
                        final_bin
                    ]
                    rows.append(row)

        return rows, None

    @classmethod
    def aggregate_to_excel(
        cls,
        manifest_csv_path: str,
        output_xlsx_path: str,
        log_cb: Optional[Callable[[str], None]] = None,
        progress_cb: Optional[Callable[[int, int, int, str], None]] = None
    ) -> Tuple[int, int, int]:
        """
        讀取下載清單 CSV，解析其所有「下載成功」的檔案，產出雙工作表 Excel 活頁簿：
        - Sheet 1: 下載清單
        - Sheet 2: 工作表2 (晶片明細，達 100 萬列自動切頁 工作表2_2...)
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

        # 動態定位下載清單中的必要欄位
        col_map = {name.strip(): idx for idx, name in enumerate(manifest_headers)}
        path_col = col_map.get("本地完整路徑")
        status_col = col_map.get("處理狀態")
        l1_col = col_map.get("第 1 層目錄")
        pn_col = col_map.get("第 1 層搜尋字")
        l2_col = col_map.get("第 2 層目錄")
        lot_col = col_map.get("第 2 層搜尋字")

        if path_col is None or status_col is None:
            raise ValueError("下載清單缺少必要欄位：'本地完整路徑' 或 '處理狀態'")

        # 預先篩選所有「下載成功」之實體檔案
        valid_items = []
        for r_idx, r in enumerate(manifest_rows, start=2):
            local_path = r[path_col].strip() if path_col < len(r) else ""
            status = r[status_col].strip() if status_col < len(r) else ""
            if status == "下載成功" and local_path:
                valid_items.append((r_idx, r, local_path))

        total_files = len(valid_items)

        # 2. 建立 OpenPyXL 活頁簿
        wb = openpyxl.Workbook()

        # 工作表 1: 下載清單
        ws_list = wb.active
        ws_list.title = "下載清單"
        ws_list.append(manifest_headers)
        for c in ws_list[1]:
            c.font = Font(bold=True)

        for r in manifest_rows:
            ws_list.append(r)

        # 自動調整 Sheet 1 欄寬
        for col_idx in range(1, len(manifest_headers) + 1):
            col_letter = get_column_letter(col_idx)
            ws_list.column_dimensions[col_letter].width = 18

        # 工作表 2: 明細 (支援 100 萬列自動切頁)
        page_no = 1
        ws_detail = wb.create_sheet(title=cls.OUT_SHEET_NAME)
        cls._setup_detail_sheet(ws_detail)

        current_ws = ws_detail
        row_in_sheet = 1 # 表頭已佔第 1 列
        total_written = 0
        succeeded_files = 0
        failed_files = 0

        for file_idx, (r_idx, r, local_path) in enumerate(valid_items, start=1):
            l1 = r[l1_col].strip() if l1_col is not None and l1_col < len(r) else ""
            pn = r[pn_col].strip() if pn_col is not None and pn_col < len(r) else ""
            l2 = r[l2_col].strip() if l2_col is not None and l2_col < len(r) else ""
            lot = r[lot_col].strip() if lot_col is not None and lot_col < len(r) else ""

            fname = os.path.basename(local_path)
            rows, err = cls.parse_single_file(local_path, l1, pn, l2, lot)
            if err:
                failed_files += 1
                if log_cb:
                    log_cb(f"[略過] [{file_idx}/{total_files}] 第 {r_idx} 列 {fname}: {err}")
                if progress_cb:
                    progress_cb(file_idx, total_files, total_written, fname)
                continue

            succeeded_files += 1

            # 寫入工作表 (直接 append，並以 O(1) 座標快速綁定純文字 @ 格式，兼具極致效能與純文字規格)
            for row_data in rows:
                if row_in_sheet >= cls.MAX_ROWS_PER_SHEET:
                    page_no += 1
                    sheet_name = f"{cls.OUT_SHEET_NAME}_{page_no}"
                    current_ws = wb.create_sheet(title=sheet_name)
                    cls._setup_detail_sheet(current_ws)
                    row_in_sheet = 1

                current_ws.append(row_data)
                row_in_sheet += 1
                total_written += 1
                for c_idx in range(1, len(row_data)): # 前 10 欄 (A~J) 設為純文字 @
                    current_ws.cell(row=row_in_sheet, column=c_idx).number_format = "@"

            if progress_cb:
                progress_cb(file_idx, total_files, total_written, fname)

            # 每 5 個檔案或首尾檔案輸出即時日誌
            if log_cb and (file_idx == 1 or file_idx % 5 == 0 or file_idx == total_files):
                pct = (file_idx / total_files) * 100 if total_files > 0 else 100.0
                log_cb(f"[{file_idx}/{total_files}] ({pct:.1f}%) 正在解析: {fname} (累計寫入 {total_written:,} 筆晶片)...")

        # 3. 儲存檔案
        if log_cb:
            log_cb(f"所有檔案解析完成 (共 {succeeded_files} 個成功檔，寫入 {total_written:,} 筆晶片)，正在儲存 Excel 活頁簿至硬碟 (請稍候)...")
        os.makedirs(os.path.dirname(os.path.abspath(output_xlsx_path)), exist_ok=True)
        wb.save(output_xlsx_path)

        return succeeded_files, failed_files, total_written

    @classmethod
    def _setup_detail_sheet(cls, ws) -> None:
        """設定明細工作表的表頭、粗體與指定欄寬"""
        ws.append(cls.HEADERS_DETAIL)
        for cell in ws[1]:
            cell.font = Font(bold=True)

        for col_letter, width in cls.COLUMN_WIDTHS.items():
            ws.column_dimensions[col_letter].width = width
