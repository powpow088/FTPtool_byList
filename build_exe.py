import os
import shutil
import subprocess
import sys
from pathlib import Path


def build_exe():
    root_dir = Path(__file__).resolve().parent
    os.chdir(root_dir)

    print("==================================================")
    print("  正在編譯 SFTP 批次下載工具 — 單檔可執行程式 (EXE)")
    print("==================================================")

    # 1. 清理舊快取
    print("[1/3] 清理舊 build 與 dist 目錄...")
    build_dir = root_dir / "build"
    dist_exe = root_dir / "dist" / "SFTP批次下載工具.exe"
    if build_dir.exists():
        shutil.rmtree(build_dir, ignore_errors=True)
    if dist_exe.exists():
        try:
            dist_exe.unlink()
        except OSError:
            pass

    # 2. 執行 PyInstaller
    print("[2/3] 正在執行 PyInstaller 編譯 SFTPTool.spec ...")
    cmd = [sys.executable, "-m", "PyInstaller", "--clean", "SFTPTool.spec"]
    res = subprocess.run(cmd)
    if res.returncode != 0:
        print("\n[ERROR] PyInstaller 編譯失敗！請檢查上方訊息。")
        sys.exit(res.returncode)

    # 3. 驗證產物
    print("[3/3] 驗證輸出目標...")
    if dist_exe.exists():
        size_mb = dist_exe.stat().st_size / (1024 * 1024)
        print("==================================================")
        print("  🎉 單檔 EXE 編譯成功！")
        print(f"  📁 輸出檔案: {dist_exe}")
        print(f"  📦 檔案大小: {size_mb:.2f} MB")
        print("==================================================")
    else:
        print("[ERROR] 未找到生成的 EXE 檔案！")
        sys.exit(1)


if __name__ == "__main__":
    build_exe()
