import os
import sys
from pathlib import Path

# 加入根目錄至 sys.path 確保模組可直接匯入
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from models.data_models import (
    ConflictStrategy, LogLevel, TransferStatus, TaskCancelledError,
    SFTPConfig, RuleItem, SearchSettings, RemoteFileTask, TaskStats, ReportRecord
)
from core.config_manager import ConfigManager


def test_data_models():
    print("[TEST] 1. 驗證 ConflictStrategy 與 LogLevel...")
    assert ConflictStrategy.SIZE_MTIME.value == "size_mtime"
    assert ConflictStrategy.OVERWRITE.value == "overwrite"
    assert ConflictStrategy.SKIP_EXISTING.value == "skip"

    assert LogLevel.INFO.color_hex == "#0066CC"
    assert LogLevel.SUCCESS.color_hex == "#28A745"
    assert LogLevel.WARNING.color_hex == "#D39E00"
    assert LogLevel.ERROR.color_hex == "#DC3545"

    print("[TEST] 2. 驗證 SFTPConfig 初始路徑正規化...")
    cfg1 = SFTPConfig(initial_path="home/data")
    assert cfg1.normalized_initial_path() == "/home/data"
    cfg2 = SFTPConfig(initial_path="/var/sftp")
    assert cfg2.normalized_initial_path() == "/var/sftp"
    cfg3 = SFTPConfig(initial_path="")
    assert cfg3.normalized_initial_path() == "/"

    print("[TEST] 3. 驗證 RuleItem 空白判定...")
    empty_rule = RuleItem("   ", "", "  ")
    assert empty_rule.is_completely_empty() is True
    valid_rule = RuleItem("PART_A", "", "")
    assert valid_rule.is_completely_empty() is False

    print("[TEST] 4. 驗證 ReportRecord 易讀容量計算...")
    rec1 = ReportRecord(remote_path="/a", local_path="/b", file_size=500)
    assert rec1.readable_size == "500 B"
    rec2 = ReportRecord(remote_path="/a", local_path="/b", file_size=1024 * 1024 * 2)
    assert rec2.readable_size == "2.00 MB"
    rec3 = ReportRecord(remote_path="/a", local_path="/b", file_size=1024 * 512)
    assert rec3.readable_size == "512.00 KB"
    print("  -> Data Models 驗證通過！")


def test_config_manager():
    print("[TEST] 5. 驗證密碼混淆與還原演算法...")
    passwords = ["Secret123!", "密碼@2026#", "p@$$w0rd_with_special!*()", ""]
    for pwd in passwords:
        obf = ConfigManager._obfuscate(pwd)
        deobf = ConfigManager._deobfuscate(obf)
        assert deobf == pwd, f"密碼還原失敗: 原密碼 '{pwd}', 還原 '{deobf}'"
    print("  -> 密碼混淆演算法對稱性驗證通過！")

    print("[TEST] 6. 驗證 config.json 讀取與寫入...")
    test_cfg_path = Path("tests/temp_test_config.json")
    if test_cfg_path.exists():
        test_cfg_path.unlink()

    # 1. 測試讀取不存在的檔案應回傳預設值
    sftp_cfg, search_cfg, rules = ConfigManager.load_config(test_cfg_path)
    assert sftp_cfg.host == ""
    assert sftp_cfg.port == 22
    assert len(rules) == 0

    # 2. 測試寫入自訂設定
    sftp_in = SFTPConfig(
        host="192.168.1.50",
        port=2222,
        username="test_user",
        password="my_secure_password",
        save_password=True,
        initial_path="remote/data/path"
    )
    search_in = SearchSettings(
        global_pattern="*RAW*-*.csv; *.xlsx",
        conflict_strategy=ConflictStrategy.OVERWRITE,
        local_save_path="C:\\TestDownloads"
    )
    rules_in = [
        RuleItem(l1_keyword="PART_A", l2_keyword="LOT_01", file_pattern="summary"),
        RuleItem(l1_keyword="   ", l2_keyword="", file_pattern=""), # 全空列，應被過濾
        RuleItem(l1_keyword="PART_B", l2_keyword="", file_pattern="*test*.csv")
    ]

    ConfigManager.save_config(sftp_in, search_in, rules_in, test_cfg_path)
    assert test_cfg_path.exists()

    # 3. 測試讀回並比對
    sftp_out, search_out, rules_out = ConfigManager.load_config(test_cfg_path)
    assert sftp_out.host == "192.168.1.50"
    assert sftp_out.port == 2222
    assert sftp_out.username == "test_user"
    assert sftp_out.password == "my_secure_password"
    assert sftp_out.save_password is True
    assert sftp_out.initial_path == "/remote/data/path" # 自動補上 /

    assert search_out.global_pattern == "*RAW*-*.csv; *.xlsx"
    assert search_out.conflict_strategy == ConflictStrategy.OVERWRITE
    assert search_out.local_save_path == "C:\\TestDownloads"

    # 全空列被過濾，剩下 2 列
    assert len(rules_out) == 2
    assert rules_out[0].l1_keyword == "PART_A"
    assert rules_out[0].l2_keyword == "LOT_01"
    assert rules_out[0].file_pattern == "summary"
    assert rules_out[1].l1_keyword == "PART_B"

    # 清理暫存測試檔
    if test_cfg_path.exists():
        test_cfg_path.unlink()
    print("  -> ConfigManager 序列化與反序列化驗證通過！")


if __name__ == "__main__":
    print("=== 開始執行微階段 1 驗收測試 ===")
    test_data_models()
    test_config_manager()
    print("=== 微階段 1 所有測試 100% 通過！ ===")
