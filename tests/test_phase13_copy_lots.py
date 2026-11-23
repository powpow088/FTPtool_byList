import sys
from pathlib import Path

# 將專案根目錄加入模組搜尋路徑
project_root = Path(__file__).resolve().parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from models.data_models import RuleItem
from ui.main_window import MainWindow


def test_extract_distinct_lots_two_levels():
    """驗證兩層目錄下提取第 2 層批號、排除標題列與有序去重"""
    rules = [
        RuleItem(dirs=["PN_A", "批號"]),             # 標題列，應排除
        RuleItem(dirs=["PN_A", "25CNE001-04-00"]),     # 第 1 筆
        RuleItem(dirs=["PN_A", "25C2E001-03-00"]),     # 第 2 筆
        RuleItem(dirs=["PN_B", "25CNE001-04-00"]),     # 重複批號，應去重
        RuleItem(dirs=["PN_B", "lot no."]),           # 英文標題，應排除
        RuleItem(dirs=["PN_B", ""]),                  # 空值，應排除
        RuleItem(dirs=["PN_C", "261HE006-01-00"]),     # 第 3 筆
        RuleItem(dirs=["PN_C", "25C2E001-03-00"]),     # 重複批號，應去重
    ]

    lots = MainWindow.extract_distinct_lots(rules)

    # 驗證筆數與內容（有序去重且無標題）
    assert lots == [
        "25CNE001-04-00",
        "25C2E001-03-00",
        "261HE006-01-00"
    ]

    # 驗證純逗號串接
    joined = ",".join(lots)
    assert joined == "25CNE001-04-00,25C2E001-03-00,261HE006-01-00"
    assert " " not in joined
    print("  -> test_extract_distinct_lots_two_levels 驗證通過！")


def test_extract_distinct_lots_single_level():
    """驗證單層目錄下提取第 1 層批號"""
    rules = [
        RuleItem(dirs=["LOT"]),               # 標題列
        RuleItem(dirs=["LOT_A01"]),
        RuleItem(dirs=["LOT_B02"]),
        RuleItem(dirs=["LOT_A01"]),           # 重複
    ]

    lots = MainWindow.extract_distinct_lots(rules)
    assert lots == ["LOT_A01", "LOT_B02"]
    assert ",".join(lots) == "LOT_A01,LOT_B02"
    print("  -> test_extract_distinct_lots_single_level 驗證通過！")


def test_ui_copy_button_and_font_unification():
    """驗證複製按鈕綠色系外觀，以及底部 6 顆按鈕字型統一標準化"""
    win = MainWindow()
    try:
        assert hasattr(win, "btn_copy_lots")
        btn = win.btn_copy_lots
        assert "複製批號 (,)" in btn.cget("text")
        # 驗證綠色系
        assert btn.cget("fg_color") == "#198754"

        # 驗證底部 7 顆按鈕字型嚴格一致為 ("Microsoft JhengHei UI", 12, "bold")
        for b in [
            win.btn_start,
            win.btn_aggregate,
            win.btn_aggregate2,
            win.btn_open_manifest,
            win.btn_open_folder,
            win.btn_scan_only,
            win.btn_stop
        ]:
            font_cfg = b.cget("font")
            # 支援 CTkFont 或 tuple
            if hasattr(font_cfg, "cget"):
                assert font_cfg.cget("size") == 12
                assert font_cfg.cget("weight") == "bold"
            else:
                assert font_cfg[1] == 12
                assert font_cfg[2] == "bold"

        print("  -> test_ui_copy_button_and_font_unification 驗證通過！")
    finally:
        win.destroy()


if __name__ == "__main__":
    print("=== 執行微階段 17 & 18 (一鍵複製批號與字型統一) 單元測試 ===")
    test_extract_distinct_lots_two_levels()
    test_extract_distinct_lots_single_level()
    test_ui_copy_button_and_font_unification()
    print("=== 所有測試 100% 通過！ ===")
