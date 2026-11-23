# -*- coding: utf-8 -*-
"""
測試模組: 微階段 19 (晶圓測試資料樞紐分析與 BIN 權重彙整引擎)
驗證:
1. 座標補零演算法 (pad_coordinates)
2. BIN 優先權判定演算法 (calc_bin_rank & display_rank: 32 > 1 > others)
3. LOT 與 PN 智慧提取與安全退避 (extract_lot & extract_pn)
4. 多工作表掃描、二維動態轉置、ERROR_blank 最右側排除與雙工作表產出驗證
"""

import os
import sys
import shutil
import tempfile

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import openpyxl
from core.pivot_aggregator import PivotAggregator


def test_pad_coordinates():
    """驗證座標補零演算法"""
    # 典型兩段坐標皆為 1 位數
    assert PivotAggregator.pad_coordinates("WAFER1_1_2") == "WAFER1_01_02"
    # 已為 2 位數不變
    assert PivotAggregator.pad_coordinates("WAFER1_01_02") == "WAFER1_01_02"
    assert PivotAggregator.pad_coordinates("WAFER1_10_25") == "WAFER1_10_25"
    # 前方帶多段底線
    assert PivotAggregator.pad_coordinates("FAB_214_LOT_X_1_9") == "FAB_214_LOT_X_01_09"
    # 非純數字座標不破壞
    assert PivotAggregator.pad_coordinates("WAFER_A_B") == "WAFER_A_B"
    # 底線少於 2 個
    assert PivotAggregator.pad_coordinates("ONLY_ONE") == "ONLY_ONE"
    print("  -> test_pad_coordinates 驗證通過！")


def test_calc_bin_rank_and_display():
    """驗證 BIN 優先權 (32 > 1 > others) 與顯示轉換"""
    # 優先權 3: 包含 32
    assert PivotAggregator.calc_bin_rank("32") == 3
    assert PivotAggregator.calc_bin_rank("1,32") == 3
    assert PivotAggregator.calc_bin_rank("2,32,1") == 3
    assert PivotAggregator.display_rank(3, whole_row=False) == 32
    assert PivotAggregator.display_rank(3, whole_row=True) == 32

    # 優先權 2: 包含 1 但無 32
    assert PivotAggregator.calc_bin_rank("1") == 2
    assert PivotAggregator.calc_bin_rank("1,2") == 2
    assert PivotAggregator.calc_bin_rank("2,1,5") == 2
    assert PivotAggregator.display_rank(2, whole_row=False) == 1
    assert PivotAggregator.display_rank(2, whole_row=True) == 1

    # 優先權 1: 包含其他數值
    assert PivotAggregator.calc_bin_rank("2") == 1
    assert PivotAggregator.calc_bin_rank("2,5,7") == 1
    assert PivotAggregator.display_rank(1, whole_row=False) == "others"
    assert PivotAggregator.display_rank(1, whole_row=True) == "Others"

    # 優先權 0: 空值
    assert PivotAggregator.calc_bin_rank("") == 0
    assert PivotAggregator.display_rank(0) == ""

    print("  -> test_calc_bin_rank_and_display 驗證通過！")


def test_extract_lot_and_pn():
    """驗證 LOT 與 PN 提取與退避邏輯"""
    # LOT: D 欄有乾淨母批號時優先使用
    assert PivotAggregator.extract_lot("25A7E004-01-00", "25A7E004-01-002") == "25A7E004-01-00"
    # LOT: D 欄為空時退避 C 欄去尾一碼
    assert PivotAggregator.extract_lot("", "25A7E004-01-002") == "25A7E004-01-00"
    assert PivotAggregator.extract_lot("", "A") == "A"

    # PN: B 欄有值時優先使用
    assert PivotAggregator.extract_pn("3273242", "214-3273242A-SUB-C") == "3273242"
    # PN: B 欄為空時退避使用 A 欄，確保永不為空
    assert PivotAggregator.extract_pn("", "214-3273242A-SUB-C") == "214-3273242A-SUB-C"

    print("  -> test_extract_lot_and_pn 驗證通過！")


def test_aggregate_pivot_workflow():
    """驗證完整樞紐分析流程：多工作表掃描、轉置、樣式與 ERROR_blank 隔離"""
    temp_dir = tempfile.mkdtemp()
    try:
        source_xlsx = os.path.join(temp_dir, "下載與明細彙整_20261123_120000.xlsx")
        wb = openpyxl.Workbook()

        # 工作表 1: 下載清單 (應被自動跳過)
        ws_manifest = wb.active
        ws_manifest.title = "下載清單"
        ws_manifest.append(["序號", "第 1 層目錄", "處理狀態"])
        ws_manifest.append([1, "214-3273242A-SUB-C", "下載成功"])

        # 工作表 2: 明細第 1 頁
        ws_detail1 = wb.create_sheet(title="工作表2")
        headers = [
            "第 1 層目錄", "PN", "第 2 層目錄", "LOT", "RECIPE",
            "檔案名稱", "2D ID (MPID)", "X (MP_X)", "Y (MP_Y)", "MPID_X_Y", "Unit BIN#"
        ]
        ws_detail1.append(headers)

        # 晶片 1 (RECIPE 1): 24W -> BIN 1
        ws_detail1.append([
            "214-3273242A-24W", "3273242", "25A7E004-01-000", "25A7E004-01-00", "24W",
            "file1.csv", "CHIP_A", "1", "2", "CHIP_A_1_2", 1
        ])
        # 晶片 1 (RECIPE 2): SUB-C -> BIN 32 (同一母批號的不同子批資料夾 25A7E004-01-002)
        ws_detail1.append([
            "214-3273242A-SUB-C", "3273242", "25A7E004-01-002", "25A7E004-01-00", "SUB-C",
            "file2.csv", "CHIP_A", "01", "02", "CHIP_A_01_02", 32
        ])
        # 晶片 2: 只有 24W -> BIN 2 (others)
        ws_detail1.append([
            "214-3273242A-24W", "3273242", "25A7E004-01-000", "25A7E004-01-00", "24W",
            "file1.csv", "CHIP_B", "2", "3", "CHIP_B_2_3", 2
        ])

        # 工作表 3: 明細第 2 頁 (模擬 100 萬列切頁的情境)
        ws_detail2 = wb.create_sheet(title="工作表2_2")
        ws_detail2.append(headers)
        # 晶片 3: RECIPE 為空 (測試 ERROR_blank 異常欄)
        ws_detail2.append([
            "3273173A-NORECIPE", "3273173", "25B12001-01-00", "25B12001-01-00", "",
            "file3.csv", "CHIP_C", "5", "5", "CHIP_C_5_5", 32
        ])

        wb.save(source_xlsx)
        wb.close()

        # 執行樞紐分析
        output_xlsx = os.path.join(temp_dir, "PN_LOT_RECIPE_明細與彙整_TEST.xlsx")
        logs = []
        src_rows, det_rows, sum_rows, out_path = PivotAggregator.aggregate_pivot(
            source_xlsx_path=source_xlsx,
            output_xlsx_path=output_xlsx,
            log_cb=lambda m: logs.append(m)
        )

        assert src_rows == 4
        assert det_rows == 4  # 晶片 A 在 2 個 fullLot 各 1 次 = 2，晶片 B = 1，晶片 C = 1 -> 4
        assert sum_rows == 3  # 晶片 A, 晶片 B, 晶片 C -> 3
        assert os.path.exists(output_xlsx)

        # 檢驗產出的 Excel
        res_wb = openpyxl.load_workbook(output_xlsx)

        # 1. 檢查工作表清單
        assert "中繼明細_01" in res_wb.sheetnames
        assert "RECIPE彙整_01" in res_wb.sheetnames
        assert res_wb.active == res_wb["RECIPE彙整_01"]

        # 2. 檢驗中繼明細
        ws_d = res_wb["中繼明細_01"]
        d_headers = [cell.value for cell in ws_d[1]]
        assert d_headers[:4] == ["PN", "第 2 層目錄(LOT)", "LOT", "MPID_XY"]
        assert "24W" in d_headers
        assert "SUB-C" in d_headers
        assert "ERROR_blank" in d_headers
        # ERROR_blank 必須在最右側
        assert d_headers[-1] == "ERROR_blank"
        # 驗證座標已被補零
        d_row2 = [cell.value for cell in ws_d[2]]
        assert d_row2[3] == "CHIP_A_01_02"

        # 3. 檢驗 RECIPE 彙整
        ws_s = res_wb["RECIPE彙整_01"]
        s_headers = [cell.value for cell in ws_s[1]]
        assert s_headers[:5] == ["PN", "第 2 層目錄(LOT)", "LOT", "MPID_XY", "BinCode_sumary"]
        assert s_headers[-1] == "ERROR_blank"

        # 檢查晶片 A (CHIP_A_01_02)
        # 關聯兩個 fullLot: 25A7E004-01-000 與 25A7E004-01-002
        s_row2 = [cell.value for cell in ws_s[2]]
        assert s_row2[0] == "3273242"
        assert "25A7E004-01-000" in s_row2[1]
        assert "25A7E004-01-002" in s_row2[1]
        assert s_row2[2] == "25A7E004-01-00"
        assert s_row2[3] == "CHIP_A_01_02"
        # BinCode_sumary: 晶片 A 在 24W 為 1，SUB-C 為 32 -> 整列最高權重必須是 32！
        assert s_row2[4] == 32

        # 檢查晶片 B (CHIP_B_02_03)
        # 只有 24W (BIN 2) -> BinCode_sumary 必須是 "Others"！
        s_row3 = [cell.value for cell in ws_s[3]]
        assert s_row3[3] == "CHIP_B_02_03"
        assert s_row3[4] == "Others"

        # 檢查晶片 C (CHIP_C_05_05): 只有空 RECIPE (ERROR_blank)
        # 關鍵驗證：ERROR_blank 絕對不能影響 BinCode_sumary，所以 BinCode_sumary 必須為空 (None 或 "")！
        s_row4 = [cell.value for cell in ws_s[4]]
        assert s_row4[3] == "CHIP_C_05_05"
        assert s_row4[4] in ("", None)  # 不參與判定，顯示空值！
        # 但在 ERROR_blank 欄位應有權重代碼 (BIN 32 -> 32)
        blank_col_idx = s_headers.index("ERROR_blank")
        assert s_row4[blank_col_idx] == 32

        # 4. 檢驗表頭色彩
        blank_header_cell = ws_s.cell(row=1, column=blank_col_idx + 1)
        assert blank_header_cell.fill.start_color.rgb == "009C4A00" or blank_header_cell.fill.start_color.rgb == "9C4A00"

        # 5. 檢驗 AutoFilter 存在
        assert ws_s.auto_filter.ref is not None

        res_wb.close()
        print("  -> test_aggregate_pivot_workflow 驗證通過！")

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_ui_aggregate2_button_and_states():
    """驗證 UI 上 btn_aggregate2 按鈕存在、順序正確、字型與三態變化邏輯"""
    from ui.main_window import MainWindow

    win = MainWindow()
    temp_dir = tempfile.mkdtemp()
    try:
        assert hasattr(win, "btn_aggregate2")
        btn = win.btn_aggregate2
        assert "彙整資料2" in btn.cget("text")

        # 驗證字型為 12pt 粗體
        font_cfg = btn.cget("font")
        if hasattr(font_cfg, "cget"):
            assert font_cfg.cget("size") == 12
            assert font_cfg.cget("weight") == "bold"
        else:
            assert font_cfg[1] == 12
            assert font_cfg[2] == "bold"

        # 狀態 1: 本地無明細檔案 -> 停用 (disabled, 灰色)
        win.ed_local_path.delete(0, "end")
        win.ed_local_path.insert(0, temp_dir)
        win._update_aggregate2_btn_state()
        assert btn.cget("state") == "disabled"
        assert btn.cget("fg_color") == "#6C757D"

        # 狀態 2: 本地有「下載與明細彙整_*.xlsx」，但尚未產出樞紐報表 -> 啟用高亮 (normal, 紫色)
        detail_xlsx = os.path.join(temp_dir, "下載與明細彙整_20261123_100000.xlsx")
        open(detail_xlsx, "w").close()
        win._update_aggregate2_btn_state()
        assert btn.cget("state") == "normal"
        assert btn.cget("fg_color") == "#6F42C1"

        # 狀態 3: 本地已有「PN_LOT_RECIPE_明細與彙整_*.xlsx」 -> 完成態 (normal, 綠色)
        pivot_xlsx = os.path.join(temp_dir, "PN_LOT_RECIPE_明細與彙整_20261123_100500.xlsx")
        open(pivot_xlsx, "w").close()
        win._update_aggregate2_btn_state()
        assert btn.cget("state") == "normal"
        assert btn.cget("fg_color") == "#20C997"

        print("  -> test_ui_aggregate2_button_and_states 驗證通過！")
    finally:
        win.destroy()
        shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    print("=== 執行微階段 19 (晶圓測試資料樞紐分析與 BIN 權重彙整引擎) 單元測試 ===")
    test_pad_coordinates()
    test_calc_bin_rank_and_display()
    test_extract_lot_and_pn()
    test_aggregate_pivot_workflow()
    test_ui_aggregate2_button_and_states()
    print("=== 微階段 19 所有測試 100% 通過！ ===")
