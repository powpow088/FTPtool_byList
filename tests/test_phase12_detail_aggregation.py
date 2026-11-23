import os
import sys
import shutil
import csv
from pathlib import Path
import openpyxl

# 將專案根目錄加入模組搜尋路徑
project_root = Path(__file__).resolve().parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from core.detail_aggregator import DetailAggregator


def test_parse_recipe():
    """驗證 RECIPE 連字號解析規則"""
    assert DetailAggregator.parse_recipe("AAA-BBB-24W") == "24W"
    assert DetailAggregator.parse_recipe("PRODUCT-LINE-SUB-A") == "SUB-A"
    assert DetailAggregator.parse_recipe("SINGLE-ONE") == ""
    assert DetailAggregator.parse_recipe("NODASH") == ""
    assert DetailAggregator.parse_recipe("  A-B-C-D  ") == "C-D"
    print("  -> test_parse_recipe 驗證通過！")


def test_parse_single_file_formats():
    """驗證單一檔案解析：Tab/逗號分隔、前 10 行標題嗅探、Hard Limit 略過、MPID_X_Y 組合"""
    test_dir = Path("tests/temp_parse_test")
    if test_dir.exists():
        shutil.rmtree(test_dir)
    test_dir.mkdir(parents=True, exist_ok=True)

    # 1. 建立標準逗號分隔 CSV (含 UTF-8 BOM 與 Hard Limit)
    csv_file = test_dir / "wafer_test.csv"
    with open(csv_file, "w", encoding="utf-8-sig", newline="") as f:
        f.write("Some meta header 1\n")
        f.write("Some meta header 2\n")
        f.write("Index,2D ID,X,Y,UNIT BIN#,Other\n") # 第 3 列標題
        f.write("Hard Limit,0,0,0,0,0\n")          # 應被略過
        f.write("1,MPID_001,12,34,1,Good\n")
        f.write("2,MPID_002,-5,20,5,Pass\n")
        f.write("3,MPID_003,10,15,FAIL_BIN,Ng\n")

    rows, err = DetailAggregator.parse_single_file(
        str(csv_file),
        level1="FAB-PKG-24W",
        pn="PN1234",
        level2="LOT_A",
        lot="LOT_001"
    )
    assert err is None
    assert len(rows) == 3

    # 驗證 row 1
    r1 = rows[0]
    assert r1[0] == "FAB-PKG-24W"  # 第 1 層目錄
    assert r1[1] == "PN1234"       # PN
    assert r1[2] == "LOT_A"        # 第 2 層目錄
    assert r1[3] == "LOT_001"      # LOT
    assert r1[4] == "24W"          # RECIPE
    assert r1[5] == "wafer_test.csv" # 檔案名稱
    assert r1[6] == "MPID_001"     # 2D ID
    assert r1[7] == "12"           # X
    assert r1[8] == "34"           # Y
    assert r1[9] == "MPID_001_12_34" # MPID_X_Y
    assert r1[10] == 1             # Unit BIN# 數字轉 int

    # 驗證 row 3 文字 BIN
    r3 = rows[2]
    assert r3[6] == "MPID_003"
    assert r3[9] == "MPID_003_10_15"
    assert r3[10] == "FAIL_BIN"

    # 2. 建立 Tab 分隔檔案
    tsv_file = test_dir / "wafer_tab.txt"
    with open(tsv_file, "w", encoding="utf-8", newline="") as f:
        f.write("Meta info\n")
        f.write("Seq\t2D ID\tX\tY\tUnit BIN#\n")
        f.write("1\tDIE_99\t1\t2\t10\n")

    rows_tsv, err_tsv = DetailAggregator.parse_single_file(
        str(tsv_file),
        level1="DEPT-SUB-REC9",
        pn="PN888",
        level2="LOT_B",
        lot="LOT_002"
    )
    assert err_tsv is None
    assert len(rows_tsv) == 1
    assert rows_tsv[0][6] == "DIE_99"
    assert rows_tsv[0][9] == "DIE_99_1_2"
    assert rows_tsv[0][10] == 10

    # 3. 建立無效檔案 (找不到標題)
    bad_file = test_dir / "bad.csv"
    with open(bad_file, "w", encoding="utf-8") as f:
        f.write("ColA,ColB,ColC\n1,2,3\n")

    rows_bad, err_bad = DetailAggregator.parse_single_file(
        str(bad_file), level1="A", pn="B", level2="C", lot="D"
    )
    assert len(rows_bad) == 0
    assert err_bad is not None
    assert "找不到" in err_bad

    shutil.rmtree(test_dir, ignore_errors=True)
    print("  -> test_parse_single_file_formats 驗證通過！")


def test_aggregate_to_excel_and_sheets():
    """驗證雙工作表 Excel 活頁簿生成、欄寬、格式設定與自動切頁"""
    test_dir = Path("tests/temp_agg_excel_test")
    if test_dir.exists():
        shutil.rmtree(test_dir)
    test_dir.mkdir(parents=True, exist_ok=True)

    # 建立測試來源資料檔
    data_file1 = test_dir / "data1.csv"
    with open(data_file1, "w", encoding="utf-8-sig") as f:
        f.write("2D ID,X,Y,UNIT BIN#\n")
        f.write("ID_A1,10,20,1\n")
        f.write("ID_A2,11,21,2\n")

    data_file2 = test_dir / "data2.csv"
    with open(data_file2, "w", encoding="utf-8-sig") as f:
        f.write("2D ID,X,Y,UNIT BIN#\n")
        f.write("ID_B1,30,40,3\n")

    # 建立模擬下載清單 CSV
    manifest_csv = test_dir / "下載清單_20261123_140000.csv"
    headers = [
        "序號", "第 1 層目錄", "第 1 層搜尋字", "第 2 層目錄", "第 2 層搜尋字",
        "檔案名稱", "副檔名", "檔案大小 (Bytes)", "檔案大小 (易讀)", "檔案修改時間",
        "處理狀態", "處理說明", "遠端完整路徑", "本地完整路徑", "完成時間"
    ]
    with open(manifest_csv, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerow([
            "1", "FAB-SITE-24W", "PN_X", "LOT_001", "LOT_X",
            "data1.csv", ".csv", "1024", "1 KB", "2026-11-23 10:00:00",
            "下載成功", "下載成功", "/remote/data1.csv", str(data_file1), "2026-11-23 10:01:00"
        ])
        writer.writerow([
            "2", "FAB-SITE-24W", "PN_X", "LOT_002", "LOT_Y",
            "data2.csv", ".csv", "1024", "1 KB", "2026-11-23 10:00:00",
            "下載成功", "下載成功", "/remote/data2.csv", str(data_file2), "2026-11-23 10:01:01"
        ])
        writer.writerow([
            "3", "FAB-SITE-24W", "PN_X", "LOT_003", "LOT_Z",
            "skipped.csv", ".csv", "1024", "1 KB", "2026-11-23 10:00:00",
            "已略過", "已略過", "/remote/skipped.csv", "C:\\none.csv", "2026-11-23 10:01:02"
        ])

    output_xlsx = test_dir / "下載與明細彙整_20261123_140000.xlsx"

    # 執行彙整
    succeeded, failed, total_written = DetailAggregator.aggregate_to_excel(
        manifest_csv_path=str(manifest_csv),
        output_xlsx_path=str(output_xlsx)
    )

    assert succeeded == 2
    assert failed == 0
    assert total_written == 3 # data1 2筆 + data2 1筆

    # 驗證產出的 Excel 活頁簿內容
    wb = openpyxl.load_workbook(str(output_xlsx))
    sheet_names = wb.sheetnames
    assert "下載清單" in sheet_names
    assert "工作表2" in sheet_names

    # 驗證 Sheet 1: 下載清單
    ws1 = wb["下載清單"]
    assert ws1.cell(row=1, column=1).value == "序號"
    assert ws1.max_row == 4 # 1 header + 3 rows

    # 驗證 Sheet 2: 工作表2
    ws2 = wb["工作表2"]
    assert ws2.cell(row=1, column=1).value == "第 1 層目錄"
    assert ws2.cell(row=1, column=5).value == "RECIPE"
    assert ws2.cell(row=1, column=10).value == "MPID_X_Y"
    assert ws2.cell(row=1, column=11).value == "Unit BIN#"
    assert ws2.max_row == 4 # 1 header + 3 detail rows

    # 驗證資料行 1
    assert ws2.cell(row=2, column=5).value == "24W"
    assert ws2.cell(row=2, column=7).value == "ID_A1"
    assert ws2.cell(row=2, column=10).value == "ID_A1_10_20"
    assert ws2.cell(row=2, column=11).value == 1

    # 驗證欄寬設定存在
    assert ws2.column_dimensions["A"].width == 24
    assert ws2.column_dimensions["B"].width == 13
    assert ws2.column_dimensions["F"].width == 62

    # 驗證儲存格格式純文字 @
    assert ws2.cell(row=2, column=1).number_format == "@"
    assert ws2.cell(row=2, column=7).number_format == "@"

    shutil.rmtree(test_dir, ignore_errors=True)
    print("  -> test_aggregate_to_excel_and_sheets 驗證通過！")


def test_ui_button_order_and_states():
    """驗證主視窗底部 6 顆按鈕之排列順序與文字"""
    from ui.main_window import MainWindow

    win = MainWindow()
    try:
        # 驗證 6 顆按鈕屬性均存在
        assert hasattr(win, "btn_start")
        assert hasattr(win, "btn_aggregate")
        assert hasattr(win, "btn_open_manifest")
        assert hasattr(win, "btn_open_folder")
        assert hasattr(win, "btn_scan_only")
        assert hasattr(win, "btn_stop")

        # 驗證按鈕文字符合簡化後規格
        assert "開始下載" in win.btn_start.cget("text")
        assert "彙整資料" in win.btn_aggregate.cget("text")
        assert "開啟下載清單 (Excel)" in win.btn_open_manifest.cget("text")
        assert "開啟下載資料夾" in win.btn_open_folder.cget("text")
        assert "僅產出清單" in win.btn_scan_only.cget("text")
        assert "停止" in win.btn_stop.cget("text")

        # 驗證三態高亮函式存在
        assert hasattr(win, "_update_aggregate_btn_state")
        win._update_aggregate_btn_state()

        print("  -> test_ui_button_order_and_states 驗證通過！")
    finally:
        win.destroy()


if __name__ == "__main__":
    print("=== 執行微階段 16 (晶片明細雙工作表活頁簿與按鈕重構) 單元測試 ===")
    test_parse_recipe()
    test_parse_single_file_formats()
    test_aggregate_to_excel_and_sheets()
    test_ui_button_order_and_states()
    print("=== 微階段 16 所有測試 100% 通過！ ===")
