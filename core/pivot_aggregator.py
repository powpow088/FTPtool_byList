# -*- coding: utf-8 -*-
"""
核心模組: 晶圓測試資料樞紐分析與 BIN 權重彙整引擎 (PivotAggregator)
完整實現 000.md 之 ExportPNLOTRecipeWithBlankWarning 業務規則：
1. 自動讀取下載與明細彙整活頁簿中所有 A1 為「第 1 層目錄」之明細工作表
2. 座標補零正規化 (RBE_PadCoordinates: 末兩碼座標補零為兩位數)
3. LOT 智慧提取規則 (優先取 D 欄 LOT 乾淨母批號，若空退避 C 欄去尾一碼)
4. PN 絕對防呆 (優先取 B 欄，若空退避 A 欄，保證絕不為空)
5. 橫向 RECIPE 動態展開與 BIN 清單逗號去重串接
6. BIN 優先權判定演算法 (32 > 1 > others) 與整列 BinCode_sumary 計算
7. ERROR_blank 異常欄獨立置於最右側，排除不參與整列 BinCode_sumary 判定
8. 中繼明細與 RECIPE 彙整雙活頁簿建置、群組保護與 100 萬列自動切頁
9. 商業級深藍白字/橘白字表頭、列高 30、欄寬自適應與 AutoFilter
"""

import os
import datetime
from typing import List, Dict, Tuple, Any, Optional, Callable
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


class PivotAggregator:
    """晶圓樞紐與 BIN 權重統計彙整引擎"""

    MAX_ROWS_PER_SHEET = 1000000  # 100 萬列自動切頁
    BLANK_HEADER = "ERROR_blank"   # RECIPE 空白保留欄名

    # 色彩定義 (對齊 000.md VBA)
    # RGB(31, 78, 120) -> #1F4E78 (深藍)
    HEADER_FILL_NORMAL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    # RGB(156, 74, 0) -> #9C4A00 (深橘)
    HEADER_FILL_BLANK = PatternFill(start_color="9C4A00", end_color="9C4A00", fill_type="solid")
    HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 欄寬標準
    WIDTH_PN = 14
    WIDTH_LOT = 20
    WIDTH_MPID_XY = 30
    WIDTH_RECIPE_COL = 18
    WIDTH_SUMMARY_BINCODE = 21

    @staticmethod
    def pad_coordinates(mpid: str) -> str:
        """
        座標正規化演算法 (對齊 VBA RBE_PadCoordinates):
        以底線 '_' 分割，若長度 >= 3 且最後兩段為 1 位數純數字，自動補零為 2 位數。
        例如: '25A7E_1_2' -> '25A7E_01_02'
              'MPID_01_02' -> 'MPID_01_02'
        """
        parts = mpid.split("_")
        if len(parts) >= 3:
            if len(parts[-2]) == 1 and parts[-2].isdigit():
                parts[-2] = f"0{parts[-2]}"
            if len(parts[-1]) == 1 and parts[-1].isdigit():
                parts[-1] = f"0{parts[-1]}"
        return "_".join(parts)

    @staticmethod
    def calc_bin_rank(codes: str) -> int:
        """
        BIN 優先權判定演算法 (對齊 VBA RBE_BinRank):
        - 代碼包含 '32' -> Priority 3 (最高)
        - 代碼包含 '1'  -> Priority 2
        - 代碼包含其他非空數值 -> Priority 1 (others)
        - 空白或無代碼 -> Priority 0
        """
        if not codes:
            return 0
        rank = 0
        for token in codes.split(","):
            t = token.strip()
            if t == "32":
                return 3
            elif t == "1":
                if rank < 2:
                    rank = 2
            elif t != "":
                if rank == 0:
                    rank = 1
        return rank

    @staticmethod
    def display_rank(priority: int, whole_row: bool = False) -> Any:
        """
        權重代碼顯示轉換 (對齊 VBA RBE_DisplayRank):
        - 3 -> 32
        - 2 -> 1
        - 1 -> 'Others' (若為整列綜合) 或 'others' (若為單一 RECIPE 欄位)
        - 0 -> 空字串
        """
        if priority == 3:
            return 32
        elif priority == 2:
            return 1
        elif priority == 1:
            return "Others" if whole_row else "others"
        return ""

    @staticmethod
    def extract_lot(clean_lot: str, full_lot: str) -> str:
        """
        LOT 智慧提取規則 (優化原 VBA):
        - 優先直接採用 clean_lot (使用者於介面貼入之母批號，絕不誤切有效尾字)
        - 若 clean_lot 為空，退避取 full_lot 去除最後一碼 (長度 > 1) 或 full_lot 本身
        """
        c = clean_lot.strip()
        if c:
            return c
        f = full_lot.strip()
        if len(f) > 1:
            return f[:-1]
        return f

    @staticmethod
    def extract_pn(clean_pn: str, level1_dir: str) -> str:
        """
        PN 安全退避原則:
        - 優先採用 clean_pn
        - 若為空，退避採用 level1_dir，保證產出報表中 PN 永不為空
        """
        p = clean_pn.strip()
        if p:
            return p
        return level1_dir.strip()

    @classmethod
    def aggregate_pivot(
        cls,
        source_xlsx_path: str,
        output_xlsx_path: Optional[str] = None,
        log_cb: Optional[Callable[[str], None]] = None,
        progress_cb: Optional[Callable[[float, str], None]] = None
    ) -> Tuple[int, int, int, str]:
        """
        核心執行進入點:
        讀取 source_xlsx_path 中所有 A1 為「第 1 層目錄」的工作表，
        輸出包含「中繼明細_01...」與「RECIPE彙整_01...」的活頁簿。
        回傳: (來源有效列數, 中繼明細寫入列數, RECIPE彙整寫入列數, 產出檔案路徑)
        """
        if not os.path.exists(source_xlsx_path):
            raise FileNotFoundError(f"找不到來源活頁簿: {source_xlsx_path}")

        if log_cb:
            log_cb(f"正在載入來源活頁簿: {os.path.basename(source_xlsx_path)} ...")
        if progress_cb:
            progress_cb(0.05, f"載入來源檔案: {os.path.basename(source_xlsx_path)}")

        # 1. 載入來源活頁簿，掃描符合標頭的工作表
        src_wb = openpyxl.load_workbook(source_xlsx_path, read_only=True, data_only=True)
        source_sheets = []

        for sheet_name in src_wb.sheetnames:
            ws = src_wb[sheet_name]
            # 讀取第 1 列表頭 (前 15 欄)
            header_row = next(ws.iter_rows(min_row=1, max_row=1, max_col=15, values_only=True), None)
            if header_row and len(header_row) >= 1 and str(header_row[0]).strip() == "第 1 層目錄":
                source_sheets.append((sheet_name, ws))

        if not source_sheets:
            src_wb.close()
            raise ValueError("在來源檔案中找不到任何 A1 為「第 1 層目錄」的明細工作表！")

        if log_cb:
            log_cb(f"共辨識到 {len(source_sheets)} 張明細工作表: {', '.join([s[0] for s in source_sheets])}")

        # 2. 記憶體聚合結構
        # details: Dict[pn, Dict[full_lot, {
        #     "lot": clean_lot,
        #     "ids": List[mpid_xy],
        #     "id_set": Dict[mpid_xy, int], # mpid_xy -> item_idx
        #     "bins": Dict[recipe, Dict[item_idx, List[str]]]
        # }]]
        details: Dict[str, Dict[str, Dict[str, Any]]] = {}
        recipes_ordered: List[str] = []
        recipes_set: set = set()
        blank_rows_count = 0
        total_source_rows = 0

        # 3. 讀取並聚合來源資料
        for s_idx, (sheet_name, ws) in enumerate(source_sheets, start=1):
            if log_cb:
                log_cb(f"正在讀取 [{s_idx}/{len(source_sheets)}] {sheet_name} ...")
            if progress_cb:
                progress_cb(0.1 + (s_idx / len(source_sheets)) * 0.35, f"讀取來源表: {sheet_name} ({s_idx}/{len(source_sheets)})")

            rows_iter = ws.iter_rows(min_row=1, values_only=True)
            headers = next(rows_iter, None)
            if not headers:
                continue

            # 動態定位欄位索引
            col_map = {str(h).strip(): idx for idx, h in enumerate(headers) if h is not None}
            col_l1 = col_map.get("第 1 層目錄", 0)
            col_pn = col_map.get("PN", 1)
            col_l2 = col_map.get("第 2 層目錄", 2)
            col_lot = col_map.get("LOT", 3)
            col_recipe = col_map.get("RECIPE", 4)
            col_mpid_xy = col_map.get("MPID_X_Y", 9)
            col_bin = col_map.get("Unit BIN#", 10)

            for row in rows_iter:
                if not row or all(v is None or str(v).strip() == "" for v in row):
                    continue

                total_source_rows += 1

                raw_l1 = str(row[col_l1]).strip() if col_l1 < len(row) and row[col_l1] is not None else ""
                raw_pn = str(row[col_pn]).strip() if col_pn < len(row) and row[col_pn] is not None else ""
                raw_l2 = str(row[col_l2]).strip() if col_l2 < len(row) and row[col_l2] is not None else ""
                raw_lot = str(row[col_lot]).strip() if col_lot < len(row) and row[col_lot] is not None else ""
                raw_recipe = str(row[col_recipe]).strip() if col_recipe < len(row) and row[col_recipe] is not None else ""
                raw_mpid = str(row[col_mpid_xy]).strip() if col_mpid_xy < len(row) and row[col_mpid_xy] is not None else ""
                raw_bin = row[col_bin] if col_bin < len(row) and row[col_bin] is not None else ""

                # 五個關鍵欄位皆為空時略過
                if not raw_pn and not raw_l2 and not raw_recipe and not raw_mpid and raw_bin == "":
                    continue

                pn = cls.extract_pn(raw_pn, raw_l1)
                full_lot = raw_l2
                lot = cls.extract_lot(raw_lot, full_lot)
                mpid = cls.pad_coordinates(raw_mpid)

                # BIN 值數值正規化 (例如 1.0 -> 1)
                if isinstance(raw_bin, float) and raw_bin.is_integer():
                    bin_code = str(int(raw_bin))
                else:
                    bin_code = str(raw_bin).strip()

                # RECIPE 處理 (空值歸入 ERROR_blank)
                if not raw_recipe:
                    recipe = cls.BLANK_HEADER
                    blank_rows_count += 1
                else:
                    recipe = raw_recipe
                    if recipe not in recipes_set and recipe != cls.BLANK_HEADER:
                        recipes_set.add(recipe)
                        recipes_ordered.append(recipe)

                # 建立 pn / full_lot 巢狀群組
                if pn not in details:
                    details[pn] = {}
                lots_map = details[pn]

                if full_lot not in lots_map:
                    lots_map[full_lot] = {
                        "lot": lot,
                        "ids": [],
                        "id_set": {},
                        "bins": {}
                    }
                grp = lots_map[full_lot]

                # 記錄晶片 ID
                if mpid not in grp["id_set"]:
                    item_idx = len(grp["ids"])
                    grp["ids"].append(mpid)
                    grp["id_set"][mpid] = item_idx
                else:
                    item_idx = grp["id_set"][mpid]

                # 記錄 BIN 代碼 (同晶片多個代碼以逗號串接且不重複)
                if recipe not in grp["bins"]:
                    grp["bins"][recipe] = {}
                recipe_bins = grp["bins"][recipe]

                if item_idx not in recipe_bins:
                    recipe_bins[item_idx] = [bin_code]
                else:
                    if bin_code and bin_code not in recipe_bins[item_idx]:
                        recipe_bins[item_idx].append(bin_code)

        src_wb.close()

        if total_source_rows == 0 or not details:
            raise ValueError("來源工作表中無任何有效明細資料可供彙整！")

        # 若有空白 RECIPE，將 ERROR_blank 固定置於最右側
        if blank_rows_count > 0:
            if cls.BLANK_HEADER not in recipes_ordered:
                recipes_ordered.append(cls.BLANK_HEADER)

        if log_cb:
            log_cb(f"來源讀取完成：共 {total_source_rows:,} 筆明細，辨識出 {len(recipes_ordered)} 個 RECIPE 欄位")

        # 4. 建立摘要資料庫 (summaries)
        # summaries: Dict[pn, Dict[lot, {
        #     "ids": List[mpid_xy],
        #     "id_set": Dict[mpid_xy, int],
        #     "full_lists": Dict[item_idx, List[str]], # all full_lots
        #     "bins": Dict[recipe, Dict[item_idx, int]], # recipe -> item_idx -> max_priority
        #     "row_ranks": Dict[item_idx, int] # item_idx -> row_max_priority (excluding blank)
        # }]]
        summaries: Dict[str, Dict[str, Dict[str, Any]]] = {}

        for pn, lots_map in details.items():
            if pn not in summaries:
                summaries[pn] = {}
            sum_lots = summaries[pn]

            for full_lot, grp in lots_map.items():
                lot = grp["lot"]
                if lot not in sum_lots:
                    sum_lots[lot] = {
                        "ids": [],
                        "id_set": {},
                        "full_lists": {},
                        "bins": {},
                        "row_ranks": {}
                    }
                s_grp = sum_lots[lot]

                # 映射 ID
                id_map = {}
                for local_idx, mpid in enumerate(grp["ids"]):
                    if mpid not in s_grp["id_set"]:
                        s_idx = len(s_grp["ids"])
                        s_grp["ids"].append(mpid)
                        s_grp["id_set"][mpid] = s_idx
                        s_grp["full_lists"][s_idx] = [full_lot]
                    else:
                        s_idx = s_grp["id_set"][mpid]
                        if full_lot not in s_grp["full_lists"][s_idx]:
                            s_grp["full_lists"][s_idx].append(full_lot)
                    id_map[local_idx] = s_idx

                # 計算 BIN 優先權
                for recipe, hits in grp["bins"].items():
                    if recipe not in s_grp["bins"]:
                        s_grp["bins"][recipe] = {}
                    dst_hits = s_grp["bins"][recipe]

                    for local_idx, code_list in hits.items():
                        s_idx = id_map[local_idx]
                        codes_str = ",".join(code_list)
                        priority = cls.calc_bin_rank(codes_str)
                        if priority > 0:
                            # 保留最高優先權
                            if s_idx not in dst_hits or priority > dst_hits[s_idx]:
                                dst_hits[s_idx] = priority
                            # 正式 RECIPE 才參與整列判定 (ERROR_blank 排除)
                            if recipe != cls.BLANK_HEADER:
                                if s_idx not in s_grp["row_ranks"] or priority > s_grp["row_ranks"][s_idx]:
                                    s_grp["row_ranks"][s_idx] = priority

        # 5. 決定產出檔案路徑
        if not output_xlsx_path:
            dir_name = os.path.dirname(os.path.abspath(source_xlsx_path))
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_xlsx_path = os.path.join(dir_name, f"PN_LOT_RECIPE_明細與彙整_{timestamp}.xlsx")

        # 6. 建立目的活頁簿
        dst_wb = openpyxl.Workbook()
        # 移除預設工作表
        default_sheet = dst_wb.active

        # 7. 輸出「中繼明細」系列分頁
        if log_cb:
            log_cb("正在產出 [中繼明細] 工作表...")
        if progress_cb:
            progress_cb(0.5, "正在產出中繼明細工作表...")
        detail_written = cls._write_detail_sheets(dst_wb, details, recipes_ordered, log_cb)

        # 8. 輸出「RECIPE彙整」系列分頁
        if log_cb:
            log_cb("正在產出 [RECIPE彙整] 交叉樞紐表...")
        if progress_cb:
            progress_cb(0.75, "正在產出 RECIPE 彙整交叉表...")
        summary_written = cls._write_summary_sheets(dst_wb, summaries, recipes_ordered, log_cb)

        # 移除預設的空白 Sheet
        if default_sheet in dst_wb.worksheets:
            dst_wb.remove(default_sheet)

        # 預設選中 RECIPE彙整_01
        if "RECIPE彙整_01" in dst_wb.sheetnames:
            dst_wb.active = dst_wb["RECIPE彙整_01"]

        # 9. 儲存檔案
        if log_cb:
            log_cb(f"所有樞紐運算完成 (中繼明細 {detail_written:,} 列，RECIPE彙整 {summary_written:,} 列)，正在儲存活頁簿至硬碟 (請稍候)...")
        if progress_cb:
            progress_cb(0.9, "正在儲存 Excel 活頁簿至硬碟...")
        os.makedirs(os.path.dirname(os.path.abspath(output_xlsx_path)), exist_ok=True)
        dst_wb.save(output_xlsx_path)

        if progress_cb:
            progress_cb(1.0, "樞紐彙整完成！")
        if log_cb:
            log_cb(f"🎉 樞紐彙整完成！中繼明細: {detail_written:,} 列，RECIPE彙整: {summary_written:,} 列")

        return total_source_rows, detail_written, summary_written, output_xlsx_path

    @classmethod
    def _write_detail_sheets(
        cls,
        wb: openpyxl.Workbook,
        details: Dict[str, Dict[str, Dict[str, Any]]],
        recipes_ordered: List[str],
        log_cb: Optional[Callable[[str], None]] = None
    ) -> int:
        """寫入中繼明細分頁 (支援群組完整性保護與 100 萬列自動切頁)"""
        page_no = 1
        ws = wb.create_sheet(title=f"中繼明細_{page_no:02d}")
        col_count = 4 + len(recipes_ordered)
        cls._setup_sheet_headers(ws, is_detail=True, recipes=recipes_ordered)

        next_row = 2
        total_written = 0

        for pn, lots_map in details.items():
            for full_lot, grp in lots_map.items():
                lot = grp["lot"]
                ids = grp["ids"]
                n = len(ids)
                if n == 0:
                    continue

                # 群組保護：若整包可放入新頁，不在舊頁尾端拆碎
                room = cls.MAX_ROWS_PER_SHEET - next_row + 1
                if n <= cls.MAX_ROWS_PER_SHEET - 1 and n > room:
                    cls._finish_page(ws, next_row, col_count)
                    page_no += 1
                    ws = wb.create_sheet(title=f"中繼明細_{page_no:02d}")
                    cls._setup_sheet_headers(ws, is_detail=True, recipes=recipes_ordered)
                    next_row = 2

                # 分批寫入當前群組
                first_item = 0
                while first_item < n:
                    if next_row > cls.MAX_ROWS_PER_SHEET:
                        cls._finish_page(ws, next_row, col_count)
                        page_no += 1
                        ws = wb.create_sheet(title=f"中繼明細_{page_no:02d}")
                        cls._setup_sheet_headers(ws, is_detail=True, recipes=recipes_ordered)
                        next_row = 2

                    room = cls.MAX_ROWS_PER_SHEET - next_row + 1
                    take = min(n - first_item, room)

                    for idx in range(first_item, first_item + take):
                        mpid = ids[idx]
                        row_data = [pn, full_lot, lot, mpid]

                        # 動態填入各 RECIPE 的 BIN 代碼字串
                        for recipe in recipes_ordered:
                            code_list = grp["bins"].get(recipe, {}).get(idx, [])
                            row_data.append(",".join(code_list) if code_list else "")

                        ws.append(row_data)
                        next_row += 1
                        total_written += 1

                    first_item += take

        cls._finish_page(ws, next_row, col_count)
        return total_written

    @classmethod
    def _write_summary_sheets(
        cls,
        wb: openpyxl.Workbook,
        summaries: Dict[str, Dict[str, Dict[str, Any]]],
        recipes_ordered: List[str],
        log_cb: Optional[Callable[[str], None]] = None
    ) -> int:
        """寫入 RECIPE 彙整分頁 (支援群組完整性保護與 100 萬列自動切頁)"""
        page_no = 1
        ws = wb.create_sheet(title=f"RECIPE彙整_{page_no:02d}")
        col_count = 5 + len(recipes_ordered)
        cls._setup_sheet_headers(ws, is_detail=False, recipes=recipes_ordered)

        next_row = 2
        total_written = 0

        for pn, sum_lots in summaries.items():
            for lot, s_grp in sum_lots.items():
                ids = s_grp["ids"]
                n = len(ids)
                if n == 0:
                    continue

                # 群組保護：若整包可放入新頁，不在舊頁尾端拆碎
                room = cls.MAX_ROWS_PER_SHEET - next_row + 1
                if n <= cls.MAX_ROWS_PER_SHEET - 1 and n > room:
                    cls._finish_page(ws, next_row, col_count)
                    page_no += 1
                    ws = wb.create_sheet(title=f"RECIPE彙整_{page_no:02d}")
                    cls._setup_sheet_headers(ws, is_detail=False, recipes=recipes_ordered)
                    next_row = 2

                first_item = 0
                while first_item < n:
                    if next_row > cls.MAX_ROWS_PER_SHEET:
                        cls._finish_page(ws, next_row, col_count)
                        page_no += 1
                        ws = wb.create_sheet(title=f"RECIPE彙整_{page_no:02d}")
                        cls._setup_sheet_headers(ws, is_detail=False, recipes=recipes_ordered)
                        next_row = 2

                    room = cls.MAX_ROWS_PER_SHEET - next_row + 1
                    take = min(n - first_item, room)

                    for idx in range(first_item, first_item + take):
                        mpid = ids[idx]
                        full_lot_str = ",".join(s_grp["full_lists"].get(idx, []))
                        row_rank = s_grp["row_ranks"].get(idx, 0)
                        summary_bincode = cls.display_rank(row_rank, whole_row=True)

                        row_data = [pn, full_lot_str, lot, mpid, summary_bincode]

                        # 各 RECIPE 權重顯示
                        for recipe in recipes_ordered:
                            priority = s_grp["bins"].get(recipe, {}).get(idx, 0)
                            row_data.append(cls.display_rank(priority, whole_row=False))

                        ws.append(row_data)
                        next_row += 1
                        total_written += 1

                    first_item += take

        cls._finish_page(ws, next_row, col_count)
        return total_written

    @classmethod
    def _setup_sheet_headers(
        cls,
        ws: openpyxl.worksheet.worksheet.Worksheet,
        is_detail: bool,
        recipes: List[str]
    ) -> None:
        """設定表頭欄位、深藍/深橘白字樣式、列高與自適應欄寬"""
        if is_detail:
            fixed_headers = ["PN", "第 2 層目錄(LOT)", "LOT", "MPID_XY"]
        else:
            fixed_headers = ["PN", "第 2 層目錄(LOT)", "LOT", "MPID_XY", "BinCode_sumary"]

        headers = fixed_headers + recipes
        ws.append(headers)

        ws.row_dimensions[1].height = 30

        # 套用樣式
        for col_idx, h_text in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = cls.HEADER_FONT
            cell.alignment = cls.HEADER_ALIGNMENT
            if h_text == cls.BLANK_HEADER:
                cell.fill = cls.HEADER_FILL_BLANK
            else:
                cell.fill = cls.HEADER_FILL_NORMAL

        # 設定欄寬
        ws.column_dimensions["A"].width = cls.WIDTH_PN
        ws.column_dimensions["B"].width = 24 if is_detail else 46
        ws.column_dimensions["C"].width = cls.WIDTH_LOT
        ws.column_dimensions["D"].width = cls.WIDTH_MPID_XY

        if not is_detail:
            ws.column_dimensions["E"].width = cls.WIDTH_SUMMARY_BINCODE

        start_recipe_col = 5 if is_detail else 6
        for c_idx in range(start_recipe_col, start_recipe_col + len(recipes)):
            col_letter = get_column_letter(c_idx)
            ws.column_dimensions[col_letter].width = cls.WIDTH_RECIPE_COL

    @classmethod
    def _finish_page(
        cls,
        ws: openpyxl.worksheet.worksheet.Worksheet,
        next_row: int,
        col_count: int
    ) -> None:
        """設定 AutoFilter 自動篩選範圍"""
        if next_row > 2:
            last_col_letter = get_column_letter(col_count)
            ws.auto_filter.ref = f"A1:{last_col_letter}{next_row - 1}"
