import stat
import sys
import threading
from pathlib import Path
from unittest.mock import MagicMock

# 加入根目錄至 sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from models.data_models import (
    RuleItem, SearchSettings, TaskStats, ConflictStrategy, TaskCancelledError
)
from core.traverse_engine import TraverseEngine


class MockSFTPAttribute:
    def __init__(self, filename: str, is_dir: bool = False, size: int = 1024, mtime: int = 1700000000):
        self.filename = filename
        self.st_mode = stat.S_IFDIR if is_dir else stat.S_IFREG
        self.st_size = size
        self.st_mtime = mtime


def test_match_pattern():
    print("[TEST] 1. 驗證萬用字元與模糊比對演算法...")
    mp = TraverseEngine.match_pattern

    # 1.1 留空全選
    assert mp("any_file.csv", "") is True
    assert mp("any_file.csv", "   ") is True

    # 1.2 純文字子字串包含 (大小寫不敏感)
    assert mp("PART_A_2026", "part_a") is True
    assert mp("my_part_01", "PART") is True
    assert mp("OTHER_FOLDER", "PART") is False

    # 1.3 萬用字元 *RAW*-*.csv
    assert mp("2026_RAW_BATCH01-data.csv", "*RAW*-*.csv") is True
    assert mp("LOT5_RAW-final.csv", "*RAW*-*.csv") is True
    assert mp("2026_RAW_LOT01.txt", "*RAW*-*.csv") is False  # 副檔名非 csv
    assert mp("RAW_summary.csv", "*RAW*-*.csv") is False     # 缺少中線 '-'
    assert mp("2026_raw_batch01-data.CSV", "*RAW*-*.csv") is True # 大小寫不敏感

    # 1.4 多條件分號/逗號分隔 (OR 邏輯)
    multi_rule = "*RAW*-*.csv; *.xlsx, test*"
    assert mp("sample_RAW-run.csv", multi_rule) is True
    assert mp("data.xlsx", multi_rule) is True
    assert mp("test_results.log", multi_rule) is True
    assert mp("unrelated.txt", multi_rule) is False
    print("  -> 萬用字元與模糊比對演算法驗證通過！")


def test_compile_rule_tree():
    print("[TEST] 2. 驗證規則樹 (Rule Tree) 預編譯與聚合...")
    rules = [
        RuleItem(l1_keyword="PART_A", l2_keyword="LOT_01", file_pattern="summary"),
        RuleItem(l1_keyword="PART_A", l2_keyword="LOT_02", file_pattern="*RAW*-*.csv"),
        RuleItem(l1_keyword="PART_A", l2_keyword="LOT_01", file_pattern="report.csv"),
        RuleItem(l1_keyword="PART_B", l2_keyword="", file_pattern="") # 留空
    ]

    tree = TraverseEngine.compile_rule_tree(rules, global_file_pattern="*default*")

    # 驗證 PART_A / LOT_01 聚合了兩個 file_pattern
    assert tree["PART_A"]["LOT_01"] == {"summary", "report.csv"}
    # 驗證 PART_A / LOT_02 保持獨立
    assert tree["PART_A"]["LOT_02"] == {"*RAW*-*.csv"}
    # 驗證 PART_B / "" 自動套用全域 pattern
    assert tree["PART_B"][""] == {"*default*"}
    print("  -> 規則樹預編譯與聚合驗證通過！")


def test_scan_all_simulation():
    print("[TEST] 3. 模擬三層目錄快取遍歷與列對列關聯鎖定...")
    mock_sftp = MagicMock()

    # 模擬遠端目錄結構:
    # /root
    # ├── PART_A_DIR (dir)
    # │   ├── LOT_01_TEST (dir)
    # │   │   ├── summary_v1.csv (file)
    # │   │   ├── raw_data.csv (file)
    # │   └── LOT_02_TEST (dir)
    # │       └── 2026_RAW_BATCH01-data.csv (file)
    # └── OTHER_DIR (dir)

    def mock_listdir(path):
        if path == "/root":
            return [
                MockSFTPAttribute("PART_A_DIR", is_dir=True),
                MockSFTPAttribute("OTHER_DIR", is_dir=True)
            ]
        elif path == "/root/PART_A_DIR":
            return [
                MockSFTPAttribute("LOT_01_TEST", is_dir=True),
                MockSFTPAttribute("LOT_02_TEST", is_dir=True)
            ]
        elif path == "/root/PART_A_DIR/LOT_01_TEST":
            return [
                MockSFTPAttribute("summary_v1.csv", is_dir=False, size=2048),
                MockSFTPAttribute("raw_data.csv", is_dir=False, size=4096)
            ]
        elif path == "/root/PART_A_DIR/LOT_02_TEST":
            return [
                MockSFTPAttribute("2026_RAW_BATCH01-data.csv", is_dir=False, size=8192),
                MockSFTPAttribute("ignore_me.txt", is_dir=False, size=100)
            ]
        return []

    mock_sftp.listdir_attr.side_effect = mock_listdir

    search_settings = SearchSettings(
        global_pattern="*.csv",
        conflict_strategy=ConflictStrategy.SIZE_MTIME,
        local_save_path="C:\\LocalDownload"
    )

    rules = [
        RuleItem(l1_keyword="PART_A", l2_keyword="LOT_01", file_pattern="summary"),
        RuleItem(l1_keyword="PART_A", l2_keyword="LOT_02", file_pattern="*RAW*-*.csv")
    ]

    cancel_event = threading.Event()
    logs = []
    stats = TaskStats()

    engine = TraverseEngine(
        sftp_mgr=mock_sftp,
        initial_path="/root",
        search_settings=search_settings,
        cancel_event=cancel_event,
        log_cb=lambda lvl, msg: logs.append((lvl, msg)),
        stats_cb=lambda s: None
    )

    tasks = engine.scan_all(rules, stats)

    # 驗證結果:
    # 1. 命中 LOT_01 底下的 summary_v1.csv (因為匹配 summary)
    #    但不包含 LOT_01 底下的 raw_data.csv (因為 LOT_01 只抓 summary)
    # 2. 命中 LOT_02 底下的 2026_RAW_BATCH01-data.csv (匹配 *RAW*-*.csv)
    #    但不包含 ignore_me.txt
    assert len(tasks) == 2, f"預期命中 2 個檔案，實際命中 {len(tasks)}"
    filenames = [t.filename for t in tasks]
    assert "summary_v1.csv" in filenames
    assert "2026_RAW_BATCH01-data.csv" in filenames

    # 驗證本地路徑鏡像組合
    for t in tasks:
        if t.filename == "summary_v1.csv":
            assert t.local_full_path.endswith("PART_A_DIR\\LOT_01_TEST\\summary_v1.csv") or \
                   t.local_full_path.endswith("PART_A_DIR/LOT_01_TEST/summary_v1.csv")
            assert t.remote_full_path == "/root/PART_A_DIR/LOT_01_TEST/summary_v1.csv"

    # 驗證每層目錄僅被呼叫 1 次 listdir_attr
    called_paths = [args[0] for args, _ in mock_sftp.listdir_attr.call_args_list]
    assert called_paths.count("/root") == 1
    assert called_paths.count("/root/PART_A_DIR") == 1
    assert called_paths.count("/root/PART_A_DIR/LOT_01_TEST") == 1
    assert called_paths.count("/root/PART_A_DIR/LOT_02_TEST") == 1

    print("  -> 模擬三層目錄遍歷與路徑鏡像驗證通過！")

    # 4. 驗證取消中斷
    cancel_event.set()
    try:
        engine.scan_all(rules, stats)
        assert False, "預期應拋出 TaskCancelledError"
    except TaskCancelledError:
        print("  -> 中斷控制 (TaskCancelledError) 驗證通過！")


if __name__ == "__main__":
    print("=== 開始執行微階段 2 驗收測試 ===")
    test_match_pattern()
    test_compile_rule_tree()
    test_scan_all_simulation()
    print("=== 微階段 2 所有測試 100% 通過！ ===")
