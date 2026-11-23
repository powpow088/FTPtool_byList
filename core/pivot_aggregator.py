# core/pivot_aggregator.py
import gc
import math
import os
import sys
import tempfile
from typing import Any, Callable, Dict, Iterator, List, Optional, Sequence, Tuple
import xlsxwriter
from python_calamine import CalamineWorkbook

from core.excel_utils import safe_atomic_move, sanitize_xml_text
from core.sidecar_cache import SidecarCacheReader, get_cache_path


def sanitize_bin_value(val: Any) -> str:
    """清洗 Unit BIN# 值，確保輸出為乾淨純字串"""
    if val is None or isinstance(val, bool):
        return ""
    if isinstance(val, int):
        return str(val)
    if isinstance(val, float):
        if math.isnan(val) or math.isinf(val):
            return ""
        return str(int(val)) if val.is_integer() else str(val).strip()
    return sanitize_xml_text(val).strip()


def sanitize_text_cell(val: Any) -> str:
    """一般欄位安全轉為乾淨字串"""
    if val is None:
        return ""
    if isinstance(val, float):
        if math.isnan(val) or math.isinf(val):
            return ""
        return str(int(val)) if val.is_integer() else str(val).strip()
    return sanitize_xml_text(val).strip()


def pad_coordinates(mpid_xy: Any) -> str:
    """
    晶圓座標補零正規化：
    將 MPID_X_Y 尾端兩碼純數字自動補零為至少兩位數（例如 'ID_1_2' -> 'ID_01_02'）。
    """
    s = sanitize_text_cell(mpid_xy)
    if not s:
        return ""
    parts = s.split("_")
    if len(parts) >= 3 and parts[-1].isdigit() and parts[-2].isdigit():
        parts[-2] = parts[-2].zfill(2)
        parts[-1] = parts[-1].zfill(2)
        return "_".join(parts)
    if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
        return f"{parts[0].zfill(2)}_{parts[1].zfill(2)}"
    return s


class PaginatedStreamSheetWriter:
    """
    基於 xlsxwriter(constant_memory=True) 的百萬列自動切頁串流寫入器：
    - 支援群組完整性保護（整包 LOT 若可放入新頁則不在舊頁尾端拆碎）
    - 切頁時自動建立新工作表 (_01, _02...)、設定欄寬與純文字格式 (@)、重繪深藍/深橘表頭
    - 關閉分頁前自動設定 AutoFilter
    """

    def __init__(
        self,
        workbook: xlsxwriter.Workbook,
        base_sheet_name: str,
        fixed_headers: Sequence[str],
        dynamic_recipes: Sequence[str],
        has_error_blank: bool,
        col_widths: Sequence[float],
        max_rows_per_sheet: int = 1_000_000,
    ):
        self.wb = workbook
        self.base_name = base_sheet_name
        self.fixed_headers = list(fixed_headers)
        self.dynamic_recipes = list(dynamic_recipes)
        self.has_error_blank = has_error_blank
        self.col_widths = list(col_widths)
        self.max_rows_per_sheet = max_rows_per_sheet

        self.total_cols = (
            len(self.fixed_headers)
            + len(self.dynamic_recipes)
            + (1 if self.has_error_blank else 0)
        )

        self.fmt_text = self.wb.add_format({"num_format": "@"})
        self.fmt_header_blue = self.wb.add_format({
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": "#1F4E78",  # RGB(31, 78, 120) 深藍底白字
            "num_format": "@",
            "align": "center",
            "valign": "vcenter",
        })
        self.fmt_header_orange = self.wb.add_format({
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": "#9C4A00",  # RGB(156, 74, 0) 深橘底白字 (ERROR_blank)
            "num_format": "@",
            "align": "center",
            "valign": "vcenter",
        })

        self.page_index = 0
        self.current_row = 0
        self.ws: Optional[xlsxwriter.worksheet.Worksheet] = None
        self.first_worksheet: Optional[xlsxwriter.worksheet.Worksheet] = None
        self.total_written_rows = 0
        self._open_new_sheet()

    def _open_new_sheet(self) -> None:
        if self.ws is not None and self.current_row > 1:
            try:
                self.ws.autofilter(0, 0, self.current_row - 1, self.total_cols - 1)
            except Exception:
                pass

        self.page_index += 1
        sheet_title = f"{self.base_name}_{self.page_index:02d}"
        self.ws = self.wb.add_worksheet(sheet_title)
        if self.first_worksheet is None:
            self.first_worksheet = self.ws
        self.current_row = 0

        self.ws.set_row(0, 30)

        for col_idx in range(self.total_cols):
            width = self.col_widths[col_idx] if col_idx < len(self.col_widths) else 14
            self.ws.set_column(col_idx, col_idx, width, self.fmt_text)

        col_cursor = 0
        for h in self.fixed_headers:
            self.ws.write_string(0, col_cursor, h, self.fmt_header_blue)
            col_cursor += 1

        for recipe in self.dynamic_recipes:
            self.ws.write_string(0, col_cursor, recipe, self.fmt_header_blue)
            col_cursor += 1

        if self.has_error_blank:
            self.ws.write_string(0, col_cursor, "ERROR_blank", self.fmt_header_orange)

        self.current_row = 1

    def write_lot_group(self, group_rows: Sequence[Sequence[Any]]) -> None:
        """寫入單一 LOT 晶片群組，落實群組完整性保護與自動切頁"""
        group_len = len(group_rows)
        if group_len == 0:
            return

        fits_in_fresh_sheet = (1 + group_len) <= self.max_rows_per_sheet
        if (
            (self.current_row + group_len > self.max_rows_per_sheet)
            and self.current_row > 1
            and fits_in_fresh_sheet
        ):
            self._open_new_sheet()

        for row_vals in group_rows:
            if self.current_row >= self.max_rows_per_sheet:
                self._open_new_sheet()

            assert self.ws is not None
            for col_idx, val in enumerate(row_vals):
                if val is None or val == "":
                    continue
                if isinstance(val, int):
                    self.ws.write_number(self.current_row, col_idx, val)
                elif isinstance(val, float):
                    if val.is_integer():
                        self.ws.write_number(self.current_row, col_idx, int(val))
                    else:
                        self.ws.write_number(self.current_row, col_idx, val)
                else:
                    self.ws.write_string(self.current_row, col_idx, str(val))

            self.current_row += 1
            self.total_written_rows += 1

    def finish(self) -> None:
        """完成所有寫入，對當前分頁追加 AutoFilter"""
        if self.ws is not None and self.current_row > 1:
            try:
                self.ws.autofilter(0, 0, self.current_row - 1, self.total_cols - 1)
            except Exception:
                pass


class PivotAggregator:
    """
    彙整資料 2（PN/LOT/RECIPE 交叉樞紐分析與權重評級）：
    - 優先命中 SQLite 旁路快取 (~1.8 秒)；跨機僅有單一 .xlsx 時自動退避至 Rust calamine (~6 秒)
    - 嚴格落實正統雙層分群：
      1. details: (pn, full_lot) 群組 -> 中繼明細
      2. summaries: (pn, clean_lot) 母批號群組 -> RECIPE彙整 (多晶圓合併同一列、B欄逗號串接、跨晶圓最高權重比較)
    - 完美相容既有介面簽名：aggregate_pivot(...) -> (total_source_rows, detail_written, summary_written, output_xlsx_path)
    """

    MAX_ROWS_PER_SHEET = 1_000_000
    BLANK_HEADER = "ERROR_blank"

    REQUIRED_COLS: Tuple[str, ...] = (
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
        "Unit BIN#",
    )

    pad_coordinates = staticmethod(pad_coordinates)

    @staticmethod
    def calc_bin_rank(codes: str) -> int:
        """BIN 優先權判定演算法: 包含 32 -> 3; 包含 1 -> 2; 其他數值 -> 1; 空 -> 0"""
        if not codes:
            return 0
        tokens = [c.strip() for c in codes.split(",") if c.strip()]
        if not tokens:
            return 0
        token_set = set(tokens)
        if "32" in token_set:
            return 3
        if "1" in token_set:
            return 2
        return 1

    @staticmethod
    def display_rank(priority: int, whole_row: bool = False) -> Any:
        if priority == 3:
            return 32
        elif priority == 2:
            return 1
        elif priority == 1:
            return "Others" if whole_row else "others"
        return ""

    @staticmethod
    def extract_lot(clean_lot: str, full_lot: str) -> str:
        c = clean_lot.strip()
        if c:
            return c
        f = full_lot.strip()
        if len(f) > 1:
            return f[:-1]
        return f

    @staticmethod
    def extract_pn(clean_pn: str, level1_dir: str) -> str:
        p = clean_pn.strip()
        if p:
            return p
        return level1_dir.strip()

    @classmethod
    def stream_source_records(
        cls, source_xlsx_path: str, log_cb: Optional[Callable[[str], None]] = None
    ) -> Tuple[str, Iterator[Tuple[Any, ...]]]:
        """智慧路由資料載入器：回傳 ('sqlite_cache' | 'calamine_fallback', iterator)"""
        cache_path = get_cache_path(source_xlsx_path)
        if SidecarCacheReader.is_valid(source_xlsx_path, cache_path):
            if log_cb:
                log_cb("偵測到有效 SQLite 旁路快取，啟動極速二進位串流模式...")
            return "sqlite_cache", SidecarCacheReader.stream_records(cache_path)

        if log_cb:
            log_cb("未偵測到旁路快取（或主檔已變更），自動啟動 Rust calamine 退避解析引擎...")

        def _calamine_generator() -> Iterator[Tuple[Any, ...]]:
            wb = CalamineWorkbook.from_path(source_xlsx_path)
            for sheet_name in wb.sheet_names:
                rows_iter = iter(
                    wb.get_sheet_by_name(sheet_name).to_python(skip_empty_area=True)
                )
                header = next(rows_iter, None)
                if not header or sanitize_text_cell(header[0]) != "第 1 層目錄":
                    continue

                col_map = {
                    sanitize_text_cell(h): idx
                    for idx, h in enumerate(header)
                    if h is not None and sanitize_text_cell(h)
                }
                indices = [col_map.get(col, -1) for col in cls.REQUIRED_COLS]

                for row in rows_iter:
                    row_len = len(row)
                    yield tuple(
                        row[idx] if (0 <= idx < row_len) else None
                        for idx in indices
                    )

        return "calamine_fallback", _calamine_generator()

    @classmethod
    def aggregate_pivot(
        cls,
        source_xlsx_path: str,
        output_xlsx_path: Optional[str] = None,
        log_cb: Optional[Callable[[str], None]] = None,
        progress_cb: Optional[Callable[[float, str], None]] = None,
        max_rows_per_sheet: int = MAX_ROWS_PER_SHEET,
    ) -> Tuple[int, int, int, str]:
        """
        核心執行進入點 (完美相容既有回傳規格):
        回傳: (來源有效列數, 中繼明細寫入列數, RECIPE彙整寫入列數, 產出檔案路徑)
        """
        if not os.path.exists(source_xlsx_path):
            raise FileNotFoundError(f"找不到來源活頁簿: {source_xlsx_path}")

        if not output_xlsx_path:
            dir_name = os.path.dirname(os.path.abspath(source_xlsx_path))
            import datetime
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_xlsx_path = os.path.join(dir_name, f"PN_LOT_RECIPE_明細與彙整_{ts}.xlsx")

        source_mode, records_iter = cls.stream_source_records(source_xlsx_path, log_cb=log_cb)
        cache_path = get_cache_path(source_xlsx_path)

        # -------------------------------------------------------------
        # 階段一：RECIPE 欄位預探索
        # -------------------------------------------------------------
        recipe_to_id: Dict[str, int] = {}
        id_to_recipe: List[str] = []
        has_error_blank = False

        if source_mode == "sqlite_cache":
            for rec in SidecarCacheReader.get_distinct_recipes(cache_path):
                interned_rec = sys.intern(rec)
                recipe_to_id[interned_rec] = len(id_to_recipe)
                id_to_recipe.append(interned_rec)
            has_error_blank = SidecarCacheReader.has_blank_recipe(cache_path)

        ERROR_BLANK_ID = -1

        # 正統雙層分群資料結構：
        # details[pn][full_lot] = {"lot": lot, "ids": [mpid], "id_set": {mpid: idx}, "bins": {rec_id: {idx: list[bin]}}}
        # summaries[pn][lot] = {"ids": [mpid], "id_set": {mpid: idx}, "full_lists": {idx: [full_lot...]}, "bins": {rec_id: {idx: max_rank}}, "row_ranks": {idx: max_row_rank}}
        details: Dict[str, Dict[str, Dict[str, Any]]] = {}
        summaries: Dict[str, Dict[str, Dict[str, Any]]] = {}

        total_source_rows = 0
        was_gc_enabled = gc.isenabled()
        if was_gc_enabled:
            gc.disable()

        try:
            for row in records_iter:
                total_source_rows += 1
                if progress_cb and total_source_rows % 25_000 == 0:
                    progress_cb(0.1, f"讀取與分群聚合: {total_source_rows:,} 列")

                l1 = sanitize_text_cell(row[0])
                pn_raw = sanitize_text_cell(row[1])
                l2 = sanitize_text_cell(row[2])
                lot_raw = sanitize_text_cell(row[3])
                recipe_raw = sanitize_text_cell(row[4])
                mpid = sanitize_text_cell(row[6])
                mp_x = sanitize_text_cell(row[7])
                mp_y = sanitize_text_cell(row[8])
                mpid_xy_raw = sanitize_text_cell(row[9])
                unit_bin = sanitize_bin_value(row[10])

                # 關鍵欄位全為空時略過
                if not pn_raw and not l2 and not recipe_raw and not mpid and not mpid_xy_raw and not unit_bin:
                    continue

                pn = sys.intern(cls.extract_pn(pn_raw, l1))
                full_lot = sys.intern(l2)
                lot = sys.intern(cls.extract_lot(lot_raw, full_lot))

                if not mpid_xy_raw and (mpid or mp_x or mp_y):
                    mpid_xy_raw = f"{mpid}_{mp_x}_{mp_y}"
                mpid_xy = pad_coordinates(mpid_xy_raw)
                if not mpid_xy:
                    continue

                if not recipe_raw:
                    has_error_blank = True
                    rec_id = ERROR_BLANK_ID
                else:
                    rec_intern = sys.intern(recipe_raw)
                    rec_id = recipe_to_id.get(rec_intern)  # type: ignore[assignment]
                    if rec_id is None:
                        rec_id = len(id_to_recipe)
                        recipe_to_id[rec_intern] = rec_id
                        id_to_recipe.append(rec_intern)

                # 1. 填入 details 群組（以 full_lot 為鍵）
                if pn not in details:
                    details[pn] = {}
                lots_map = details[pn]
                if full_lot not in lots_map:
                    lots_map[full_lot] = {"lot": lot, "ids": [], "id_set": {}, "bins": {}}
                grp = lots_map[full_lot]

                if mpid_xy not in grp["id_set"]:
                    item_idx = len(grp["ids"])
                    grp["ids"].append(mpid_xy)
                    grp["id_set"][mpid_xy] = item_idx
                else:
                    item_idx = grp["id_set"][mpid_xy]

                if rec_id not in grp["bins"]:
                    grp["bins"][rec_id] = {}
                recipe_bins = grp["bins"][rec_id]

                if item_idx not in recipe_bins:
                    recipe_bins[item_idx] = [unit_bin] if unit_bin else []
                else:
                    if unit_bin and unit_bin not in recipe_bins[item_idx]:
                        recipe_bins[item_idx].append(unit_bin)

                # 2. 填入 summaries 群組（以母批號 clean_lot 為鍵，跨晶圓合併！）
                if pn not in summaries:
                    summaries[pn] = {}
                sum_lots = summaries[pn]
                if lot not in sum_lots:
                    sum_lots[lot] = {
                        "ids": [],
                        "id_set": {},
                        "full_lists": {},
                        "bins": {},
                        "row_ranks": {}
                    }
                s_grp = sum_lots[lot]

                if mpid_xy not in s_grp["id_set"]:
                    s_idx = len(s_grp["ids"])
                    s_grp["ids"].append(mpid_xy)
                    s_grp["id_set"][mpid_xy] = s_idx
                    s_grp["full_lists"][s_idx] = [full_lot]
                else:
                    s_idx = s_grp["id_set"][mpid_xy]
                    if full_lot not in s_grp["full_lists"][s_idx]:
                        s_grp["full_lists"][s_idx].append(full_lot)

                # 權重評級
                if unit_bin:
                    priority = cls.calc_bin_rank(unit_bin)
                    if priority > 0:
                        if rec_id not in s_grp["bins"]:
                            s_grp["bins"][rec_id] = {}
                        dst_hits = s_grp["bins"][rec_id]
                        if s_idx not in dst_hits or priority > dst_hits[s_idx]:
                            dst_hits[s_idx] = priority

                        if rec_id != ERROR_BLANK_ID:
                            if s_idx not in s_grp["row_ranks"] or priority > s_grp["row_ranks"][s_idx]:
                                s_grp["row_ranks"][s_idx] = priority
        finally:
            if was_gc_enabled:
                gc.enable()
            gc.collect()

        if total_source_rows == 0 or not details:
            raise ValueError("來源工作表中無任何有效明細資料可供彙整！")

        # 排序 RECIPE 欄位
        if source_mode == "calamine_fallback":
            recipes_ordered = sorted(id_to_recipe)
            old_to_new_col = {recipe_to_id[rec]: new_idx for new_idx, rec in enumerate(recipes_ordered)}
        else:
            recipes_ordered = id_to_recipe
            old_to_new_col = {idx: idx for idx in range(len(recipes_ordered))}

        num_recipes = len(recipes_ordered)
        if log_cb:
            log_cb(f"來源讀取完成：共 {total_source_rows:,} 列，辨識出 {num_recipes} 個 RECIPE" + ("（含 ERROR_blank）" if has_error_blank else ""))

        # -------------------------------------------------------------
        # 階段二：串流分頁寫出（中繼明細 & RECIPE彙整）
        # -------------------------------------------------------------
        tmp_dir = tempfile.gettempdir()
        fd, temp_xlsx_path = tempfile.mkstemp(prefix="wafer_pivot_", suffix=".xlsx", dir=tmp_dir)
        os.close(fd)

        wb = xlsxwriter.Workbook(
            temp_xlsx_path,
            {"constant_memory": True, "tmpdir": tmp_dir},
        )
        wb_closed = False

        try:
            # 1. 輸出中繼明細系列分頁
            if log_cb:
                log_cb("正在產出 [中繼明細] 工作表...")
            if progress_cb:
                progress_cb(0.4, "正在產出中繼明細工作表...")

            detail_fixed_headers = ["PN", "第 2 層目錄(LOT)", "LOT", "MPID_XY"]
            detail_col_widths = [14, 24, 20, 30] + [18] * (num_recipes + (1 if has_error_blank else 0))
            detail_writer = PaginatedStreamSheetWriter(
                workbook=wb,
                base_sheet_name="中繼明細",
                fixed_headers=detail_fixed_headers,
                dynamic_recipes=recipes_ordered,
                has_error_blank=has_error_blank,
                col_widths=detail_col_widths,
                max_rows_per_sheet=max_rows_per_sheet,
            )

            tail_cols = num_recipes + (1 if has_error_blank else 0)

            for pn, lots_map in details.items():
                for full_lot, grp in lots_map.items():
                    lot = grp["lot"]
                    ids = grp["ids"]
                    if not ids:
                        continue
                    group_rows: List[List[Any]] = []
                    for idx, mpid_xy in enumerate(ids):
                        tail: List[Any] = [""] * tail_cols
                        for rec_id, code_dict in grp["bins"].items():
                            c_list = code_dict.get(idx, [])
                            joined = ",".join(c_list) if c_list else ""
                            if rec_id == ERROR_BLANK_ID:
                                if has_error_blank:
                                    tail[-1] = joined
                            else:
                                offset = old_to_new_col[rec_id]
                                tail[offset] = joined
                        group_rows.append([pn, full_lot, lot, mpid_xy] + tail)
                    detail_writer.write_lot_group(group_rows)

            detail_writer.finish()
            detail_written = detail_writer.total_written_rows

            # 2. 輸出 RECIPE彙整 系列分頁
            if log_cb:
                log_cb("正在產出 [RECIPE彙整] 交叉樞紐表...")
            if progress_cb:
                progress_cb(0.7, "正在產出 RECIPE 彙整交叉表...")

            summary_fixed_headers = ["PN", "第 2 層目錄(LOT)", "LOT", "MPID_XY", "BinCode_sumary"]
            summary_col_widths = [14, 46, 20, 30, 21] + [18] * (num_recipes + (1 if has_error_blank else 0))
            summary_writer = PaginatedStreamSheetWriter(
                workbook=wb,
                base_sheet_name="RECIPE彙整",
                fixed_headers=summary_fixed_headers,
                dynamic_recipes=recipes_ordered,
                has_error_blank=has_error_blank,
                col_widths=summary_col_widths,
                max_rows_per_sheet=max_rows_per_sheet,
            )

            for pn, sum_lots in summaries.items():
                for lot, s_grp in sum_lots.items():
                    ids = s_grp["ids"]
                    if not ids:
                        continue
                    group_rows = []
                    for idx, mpid_xy in enumerate(ids):
                        full_lot_str = ",".join(s_grp["full_lists"].get(idx, []))
                        row_rank = s_grp["row_ranks"].get(idx, 0)
                        summary_bincode = cls.display_rank(row_rank, whole_row=True)

                        tail = [""] * tail_cols
                        for rec_id, hit_dict in s_grp["bins"].items():
                            p_val = hit_dict.get(idx, 0)
                            code_val = cls.display_rank(p_val, whole_row=False)
                            if rec_id == ERROR_BLANK_ID:
                                if has_error_blank:
                                    tail[-1] = code_val
                            else:
                                offset = old_to_new_col[rec_id]
                                tail[offset] = code_val

                        group_rows.append([pn, full_lot_str, lot, mpid_xy, summary_bincode] + tail)
                    summary_writer.write_lot_group(group_rows)

            summary_writer.finish()
            summary_written = summary_writer.total_written_rows

            # 預設聚焦於 RECIPE彙整_01
            if summary_writer.first_worksheet is not None:
                summary_writer.first_worksheet.activate()

            if log_cb:
                log_cb(f"所有樞紐運算完成 (中繼明細 {detail_written:,} 列，RECIPE彙整 {summary_written:,} 列)，正在儲存活頁簿至硬碟 (請稍候)...")
            if progress_cb:
                progress_cb(0.9, "正在儲存 Excel 活頁簿至硬碟...")

            wb.close()
            wb_closed = True

            safe_atomic_move(temp_xlsx_path, output_xlsx_path)

            if progress_cb:
                progress_cb(1.0, "樞紐彙整完成！")
            if log_cb:
                log_cb(f"🎉 樞紐彙整完成！中繼明細: {detail_written:,} 列，RECIPE彙整: {summary_written:,} 列")

            return total_source_rows, detail_written, summary_written, output_xlsx_path

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
            raise
