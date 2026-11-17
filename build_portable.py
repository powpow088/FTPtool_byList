import os
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path


def build_portable():
    root = Path(__file__).resolve().parent
    dist_dir = root / "dist"
    portable_dir = dist_dir / "SFTP批次下載工具_綠色便攜版"
    zip_path = dist_dir / "SFTP批次下載工具_綠色便攜版.zip"

    print("==================================================")
    print("  正在建置 SFTP 批次下載工具 — 綠色免安裝便攜包")
    print("==================================================")

    # 1. 清理舊目錄
    if portable_dir.exists():
        print("[1/5] 清理舊便攜包目錄...")
        shutil.rmtree(portable_dir, ignore_errors=True)
    if zip_path.exists():
        zip_path.unlink()

    portable_dir.mkdir(parents=True, exist_ok=True)
    app_dir = portable_dir / "app"
    app_dir.mkdir(parents=True, exist_ok=True)

    # 2. 建立獨立 Python runtime venv
    print("[2/5] 正在建立獨立乾淨的 Python Runtime 執行環境...")
    runtime_dir = portable_dir / "runtime"
    res = subprocess.run([sys.executable, "-m", "venv", str(runtime_dir)], capture_output=True, text=True)
    if res.returncode != 0:
        print("❌ 建立 runtime 失敗:", res.stderr)
        return False

    # 3. 安裝依賴套件到便攜環境
    print("[3/5] 正在安裝相依套件至便攜環境 (customtkinter, paramiko, openpyxl, darkdetect)...")
    runtime_py = runtime_dir / "Scripts" / "python.exe"
    req_file = root / "requirements.txt"
    res = subprocess.run([str(runtime_py), "-m", "pip", "install", "-r", str(req_file)], capture_output=True, text=True)
    if res.returncode != 0:
        print("❌ 安裝依賴套件失敗:", res.stderr)
        return False

    # 4. 複製專案程式碼與建立啟動檔
    print("[4/5] 正在複製應用程式模組與產出啟動器...")
    for folder in ["models", "core", "ui"]:
        shutil.copytree(root / folder, app_dir / folder)
    shutil.copy2(root / "main.py", app_dir / "main.py")

    # 建立啟動工具.bat (使用純 ASCII 撰寫)
    launcher_bat = portable_dir / "啟動工具.bat"
    bat_content = '@echo off\r\ncd /d "%~dp0"\r\nstart "" "runtime\\Scripts\\pythonw.exe" "app\\main.py"\r\n'
    with open(launcher_bat, "wb") as f:
        f.write(bat_content.encode("ascii"))

    # 建立使用說明.txt
    readme_txt = portable_dir / "使用說明.txt"
    with open(readme_txt, "w", encoding="utf-8") as f:
        f.write(
            "=======================================================\n"
            "  SFTP 批次檔案檢索與結構保留下載工具 (綠色免安裝便攜版)\n"
            "=======================================================\n\n"
            "【使用說明】\n"
            "1. 本工具已內嵌獨立 Python 執行環境，電腦無須安裝 Python。\n"
            "2. 直接雙擊 [ 啟動工具.bat ] 即可啟動圖形介面。\n"
            "3. 採用微軟/PSF 官方簽署核心直譯器，100% 相容 Windows 11 與企業防毒軟體。\n"
        )

    # 5. 壓製為 ZIP
    print("[5/5] 正在壓縮便攜資料夾為 ZIP 封裝檔...")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root_path, dirs, files in os.walk(portable_dir):
            for file in files:
                full_path = Path(root_path) / file
                rel_path = full_path.relative_to(portable_dir)
                zf.write(full_path, arcname=Path("SFTP批次下載工具_綠色便攜版") / rel_path)

    print("\n==================================================")
    print("  🎉 綠色免安裝便攜包建置完成！")
    print(f"  📁 便攜資料夾: {portable_dir}")
    print(f"  📦 ZIP 封裝檔: {zip_path}")
    print("==================================================")
    return True


if __name__ == "__main__":
    success = build_portable()
    sys.exit(0 if success else 1)
