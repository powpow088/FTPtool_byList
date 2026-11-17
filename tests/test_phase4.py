import sys
from pathlib import Path

# 加入根目錄至 sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import customtkinter as ctk
from ui.rule_table_ctk import CTkRuleTable
from models.data_models import RuleItem


def test_tsv_parsing():
    print("[TEST] 1. 驗證 Excel TSV 剪貼簿解析引擎 (parse_tsv_text)...")
    parser = CTkRuleTable.parse_tsv_text

    # 1.1 測試 1 欄剪貼簿資料
    tsv_1col = "PART_A\nPART_B\nPART_C\n"
    rules, truncated = parser(tsv_1col)
    assert len(rules) == 3
    assert truncated is False
    assert rules[0] == RuleItem(l1_keyword="PART_A", l2_keyword="", file_pattern="")
    assert rules[1].l1_keyword == "PART_B"
    assert rules[2].l1_keyword == "PART_C"

    # 1.2 測試 2 欄剪貼簿資料 (料號 + 批號)
    tsv_2col = "PART_A\tLOT_001\nPART_A\tLOT_002\r\nPART_B\tLOT_003\r\n"
    rules, truncated = parser(tsv_2col)
    assert len(rules) == 3
    assert truncated is False
    assert rules[0] == RuleItem(l1_keyword="PART_A", l2_keyword="LOT_001", file_pattern="")
    assert rules[1] == RuleItem(l1_keyword="PART_A", l2_keyword="LOT_002", file_pattern="")
    assert rules[2] == RuleItem(l1_keyword="PART_B", l2_keyword="LOT_003", file_pattern="")

    # 1.3 測試 3 欄剪貼簿資料
    tsv_3col = "PART_A\tLOT_001\t*RAW*-*.csv\nPART_B\tLOT_002\tsummary.xlsx\n"
    rules, truncated = parser(tsv_3col)
    assert len(rules) == 2
    assert truncated is False
    assert rules[0] == RuleItem(l1_keyword="PART_A", l2_keyword="LOT_001", file_pattern="*RAW*-*.csv")
    assert rules[1] == RuleItem(l1_keyword="PART_B", l2_keyword="LOT_002", file_pattern="summary.xlsx")

    # 1.4 測試 4 欄以上剪貼簿資料 (超過 3 欄自動截斷並回傳 True)
    tsv_4col = "PART_A\tLOT_001\tfile.csv\tEXTRA_COL1\tEXTRA_COL2\n"
    rules, truncated = parser(tsv_4col)
    assert len(rules) == 1
    assert truncated is True
    assert rules[0] == RuleItem(l1_keyword="PART_A", l2_keyword="LOT_001", file_pattern="file.csv")

    # 1.5 測試 Excel 特殊引號與尾端空換行
    tsv_quotes = '"PART_A, SPECIAL"\t"LOT 01"\t"summary.csv"\n\r\n   \t  \t  \n'
    rules, _ = parser(tsv_quotes)
    assert len(rules) == 1
    assert rules[0].l1_keyword == "PART_A, SPECIAL"
    assert rules[0].l2_keyword == "LOT 01"
    assert rules[0].file_pattern == "summary.csv"
    print("  -> Excel TSV 解析與 1/2/3/4+ 欄自動適配驗證通過！")


def test_rule_table_widget():
    print("[TEST] 2. 驗證 CTkRuleTable 介面控制項資料操作...")
    root = ctk.CTk()
    root.withdraw() # 隱藏視窗

    try:
        table = CTkRuleTable(root)
        table.pack()

        # 2.1 測試初始狀態
        assert len(table.tree.get_children()) == 0

        # 2.2 測試 load_rules
        sample_rules = [
            RuleItem("PART_A", "LOT_01", "*.csv"),
            RuleItem("PART_B", "LOT_02", "*RAW*.xlsx")
        ]
        table.load_rules(sample_rules)
        assert len(table.tree.get_children()) == 2

        # 2.3 測試 get_all_rules
        all_rules = table.get_all_rules()
        assert len(all_rules) == 2
        assert all_rules[0].l1_keyword == "PART_A"
        assert all_rules[0].l2_keyword == "LOT_01"
        assert all_rules[0].file_pattern == "*.csv"
        assert all_rules[1].l1_keyword == "PART_B"

        # 2.4 測試 append_empty_row
        table.append_empty_row()
        assert len(table.tree.get_children()) == 3
        # 空白列在 get_all_rules 應被自動忽略
        assert len(table.get_all_rules()) == 2

        # 2.5 測試選取與刪除
        children = table.tree.get_children()
        table.tree.selection_set(children[0])
        table.remove_selected_rows()
        assert len(table.tree.get_children()) == 2

        print("  -> CTkRuleTable 控制項載入、新增、刪除與過濾驗證通過！")
    finally:
        root.destroy()


if __name__ == "__main__":
    print("=== 開始執行微階段 4 驗收測試 ===")
    test_tsv_parsing()
    test_rule_table_widget()
    print("=== 微階段 4 所有測試 100% 通過！ ===")
