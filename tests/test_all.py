import sys
import subprocess
from pathlib import Path

# 加入根目錄
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


def run_all_tests():
    test_files = [
        "tests/test_phase1.py",
        "tests/test_phase2.py",
        "tests/test_phase3.py",
        "tests/test_phase4.py",
        "tests/test_phase5.py",
        "tests/test_phase7_dynamic.py",
        "tests/test_phase8_pure_dirs.py",
        "tests/test_phase9_scan_only.py",
        "tests/test_phase10_connection_guard.py",
    ]

    print("==================================================")
    print("      SFTP 批次下載工具 — 全模組自動化驗收測試      ")
    print("==================================================")

    all_passed = True
    for test_file in test_files:
        print(f"\n>>> 正在執行: {test_file} ...")
        res = subprocess.run([sys.executable, "-u", test_file], capture_output=True, text=True, encoding="utf-8")
        print(res.stdout)
        if res.returncode != 0:
            print(f"❌ 測試失敗: {test_file}")
            print(res.stderr)
            all_passed = False
            break
        else:
            print(f"✅ {test_file} 通過！")

    print("\n==================================================")
    if all_passed:
        print("  🎉🎉🎉 全數 9 大單元與整合測試 100% 通過！ 🎉🎉🎉")
    else:
        print("  ❌ 存在未通過之測試項目，請檢修！")
    print("==================================================")
    return all_passed


if __name__ == "__main__":
    success = run_all_tests()
    sys.exit(0 if success else 1)
