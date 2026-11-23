# SFTP 批次檔案檢索與結構保留下載工具 — 微階段實作任務清單 (Micro-Phase Task Checklist)

## 微階段 1：環境建立、資料模型與設定持久化 (Foundation & Models)
- [x] 1.1 建立依賴環境
  - [x] 1.1.1 建立 `requirements.txt`（包含 `customtkinter>=6.0.0`, `paramiko>=3.4.0`, `openpyxl>=3.1.2`, `darkdetect>=0.8.0`）
  - [x] 1.1.2 執行 pip 安裝並驗證核心套件能正常 import
- [x] 1.2 實作核心資料模型 `models/data_models.py`
  - [x] 1.2.1 實作 `ConflictStrategy(Enum)`（SIZE_MTIME, OVERWRITE, SKIP_EXISTING）
  - [x] 1.2.2 實作 `LogLevel(Enum)`（INFO, SUCCESS, WARNING, ERROR 及其對應 HEX 色碼）
  - [x] 1.2.3 實作 `TransferStatus(Enum)`（下載成功, 已略過, 失敗, 已取消）
  - [x] 1.2.4 實作自訂中斷例外 `TaskCancelledError(Exception)`
  - [x] 1.2.5 實作資料類別 `SFTPConfig`（host, port, username, password, save_password, initial_path）
  - [x] 1.2.6 實作資料類別 `RuleItem`（is_completely_empty 方法）
  - [x] 1.2.7 實作資料類別 `SearchSettings`（global_pattern, conflict_strategy, local_save_path）
  - [x] 1.2.8 實作資料類別 `RemoteFileTask`
  - [x] 1.2.9 實作資料類別 `TaskStats`
  - [x] 1.2.10 實作資料類別 `ReportRecord`（包含完整 10 大欄位資訊與格式化易讀容量計算）
- [x] 1.3 實作設定檔管理員 `core/config_manager.py`
  - [x] 1.3.1 實作 XOR + URLSafe-Base64 密碼混淆演算法 `_obfuscate()` 與還原演算法 `_deobfuscate()`
  - [x] 1.3.2 實作 `save_config()`：序列化寫入 `config.json`（過濾空規則列）
  - [x] 1.3.3 實作 `load_config()`：從 `config.json` 安全反序列化，防範檔案不存在或格式異常
- [x] 1.4 【微階段 1 驗收測試】
  - [x] 1.4.1 撰寫並執行測試腳本 `tests/test_phase1.py`：驗證資料模型結構、密碼混淆加解密與 `config.json` 讀寫

---

## 微階段 2：SFTP 連線管理與三層規則樹遍歷 (SFTP & Traversal Engine)
- [x] 2.1 實作 SFTP 客戶端封裝 `core/sftp_manager.py`
  - [x] 2.1.1 實作 `connect()`：建立 Paramiko SSHClient 與 SFTPClient，配置 15 秒 Socket Timeout 與 30 秒 KeepAlive
  - [x] 2.1.2 實作 `test_connection()`：驗證主機連線、帳密認證，並檢查 `initial_path` 是否存在且為資料夾
  - [x] 2.1.3 實作 `reconnect()`：安全斷開並重建 SFTP Session
  - [x] 2.1.4 實作 `listdir_attr(remote_path)`：列出目錄所有檔案與資料夾屬性
  - [x] 2.1.5 實作 `download_file(remote_path, local_temp_path, callback)`：底層分塊傳輸
  - [x] 2.1.6 實作 `close()`：安全關閉與釋放 SFTP 與 SSH 資源
- [x] 2.2 實作通用三層遍歷引擎 `core/traverse_engine.py`
  - [x] 2.2.1 實作規則樹預編譯方法 `compile_rule_tree()`
  - [x] 2.2.2 實作萬用字元與模糊比對演算法 `match_pattern()`
  - [x] 2.2.3 實作實體目錄快取遍歷 `scan_all()`
  - [x] 2.2.4 實作權限不足跳過與中斷檢查
- [x] 2.3 【微階段 2 驗收測試】
  - [x] 2.3.1 撰寫並執行測試腳本 `tests/test_phase2.py`

---

## 微階段 3：本地鏡像下載與自動 Manifest CSV 產出 (Download & Manifest)
- [x] 3.1 實作下載與本地目錄管理器 `core/download_manager.py`
  - [x] 3.1.1 實作衝突判斷 `should_skip_file()`
  - [x] 3.1.2 實作本地鏡像目錄建立
  - [x] 3.1.3 實作 `.part` 暫存檔原子性下載
  - [x] 3.1.4 實作毫秒級中斷響應
  - [x] 3.1.5 實作網路中斷 Exponential Backoff 自動重連
- [x] 3.2 實作本地完整下載清單 CSV 自動生成器
  - [x] 3.2.1 實作 `generate_manifest_csv()` 方法
  - [x] 3.2.2 格式化輸出檔名為 `下載清單_YYYYMMDD_HHMMSS.csv`，寫入本地根目錄
  - [x] 3.2.3 採用 `utf-8-sig` 編碼輸出完整 10 大欄位
- [x] 3.3 【微階段 3 驗收測試】
  - [x] 3.3.1 撰寫並執行測試腳本 `tests/test_phase3.py`

---

## 微階段 4：Excel 剪貼簿深度解析與規則表格 (CustomTkinter Table Widget)
- [x] 4.1 實作 Excel 剪貼簿表格控制項 `ui/rule_table_ctk.py`
  - [x] 4.1.1 封裝 `CTkRuleTable(ctk.CTkFrame)`，內嵌 `ttk.Treeview` 並配置 Dark 深色現代化樣式表
  - [x] 4.1.2 覆寫 `<Control-v>` 與 `<Delete>`：攔截剪貼簿貼上與刪除選取列
  - [x] 4.1.3 實作靜態方法 `parse_tsv_text()`
  - [x] 4.1.4 實作 `paste_from_clipboard()`
  - [x] 4.1.5 實作滑鼠雙擊 `<Double-1>` 儲存格浮動原地編輯
  - [x] 4.1.6 實作輔助方法
- [x] 4.2 【微階段 4 驗收測試】
  - [x] 4.2.1 撰寫並執行測試腳本 `tests/test_phase4.py`

---

## 微階段 5：非同步背景執行緒、訊號節流閥與主視窗整合 (Worker & MainWindow)
- [x] 5.1 實作背景執行緒 `core/worker.py`
  - [x] 5.1.1 定義 `UIEvent` 跨執行緒封包
  - [x] 5.1.2 實作獨立背景執行緒 `ConnectionTestWorker(threading.Thread)`
  - [x] 5.1.3 實作主工作執行緒 `DownloadWorker(threading.Thread)`
  - [x] 5.1.4 實作任務結束自動呼叫 `DownloadManager.generate_manifest_csv`
- [x] 5.2 實作主視窗 `ui/main_window.py`
  - [x] 5.2.1 建置四大區塊佈局
  - [x] 5.2.2 實作連線設定區
  - [x] 5.2.3 實作規則操作按鈕
  - [x] 5.2.4 實作下載設定區
  - [x] 5.2.5 實作監控區
  - [x] 5.2.6 實作 80ms 跨執行緒 UI 事件輪詢器
  - [x] 5.2.7 實作控制按鈕
  - [x] 5.2.8 實作設定持久化
- [x] 5.3 實作應用程式進入點 `main.py`
  - [x] 5.3.1 初始化 CustomTkinter，套用暗黑模式與深藍色調

---

## 微階段 7（動態升級）：動態多階層、遠端路徑瀏覽器與排版精簡 (Dynamic Upgrades)
- [x] 7.1 資料模型升級 (`models/data_models.py`)
  - [x] 7.1.1 新增 `MatchMode(Enum)`：`FUZZY = "fuzzy"` (模糊包含，預設) 與 `EXACT = "exact"` (完全精準相符)
  - [x] 7.1.2 擴充 `RuleItem`：支援任意 $N$ 層目錄結構清單 `dirs: List[str]` 與 `file_pattern: str`，並保留向下相容屬性 `l1_keyword`、`l2_keyword`
  - [x] 7.1.3 擴充 `SearchSettings`：納入 `match_mode: MatchMode`
  - [x] 7.1.4 擴充 `ReportRecord`：支援動態目錄階層清單與 CSV 欄位對應
- [x] 7.2 遍歷引擎支援動態 $N$ 階層與比對模式 (`core/traverse_engine.py`)
  - [x] 7.2.1 擴充 `match_pattern()`：在 `MatchMode.EXACT` 下進行完全相符比較（大小寫不敏感），萬用字元（包含 `*` 或 `?`）維持 `fnmatchcase` 萬用字元比對，留空維持代表匹配全選
  - [x] 7.2.2 實作泛型遞迴目錄遍歷演算法：支援任意深度 $N$ 之動態目錄階層展開與實體目錄快取
- [x] 7.3 動態自適應規則表格 (`ui/rule_table_ctk.py`)
  - [x] 7.3.1 擴充 `parse_tsv_text()`：動態識別 $N$ 欄（前 $N-1$ 欄為目錄層級，最後 1 欄為檔案名稱）
  - [x] 7.3.2 實作 `set_columns(dir_count: int)`：動態刷新 `tree["columns"]` 與表頭文字為 `[第 1 層目錄]`, `[第 2 層目錄]`, ..., `[檔案名稱]`
  - [x] 7.3.3 維持儲存格雙擊浮動編輯能力
- [x] 7.4 實作遠端路徑瀏覽器視窗 (`ui/remote_browser_dialog.py`)
  - [x] 7.4.1 建立 `RemoteBrowserDialog(ctk.CTkToplevel)` 獨立視窗，支援 Modal 模式與深色主題
  - [x] 7.4.2 實作路徑導航列：當前路徑文字顯示、`[⬆ 回上一層]` 按鈕、`[🔄 重新整理]` 按鈕
  - [x] 7.4.3 實作遠端目錄列表：以 `ttk.Treeview` 僅呈現資料夾名稱、修改時間，雙擊資料夾進入下一層
  - [x] 7.4.4 實作背景非同步載入目錄清單：避免網路延遲造成視窗無回應，載入期間顯示提示
  - [x] 7.4.5 實作確認與取消按鈕：點選 `[選擇此目錄]` 將當前路徑回填至主視窗 `ed_init_path`，點選 `[取消]` 關閉對話框
- [x] 7.5 主視窗版面與按鈕精簡落實 (`ui/main_window.py`)
  - [x] 7.5.1 第一列 (Row 1) 重新排列為：`主機 (Host)`、`Port`、`帳號`、`密碼 (+ 👁)` 相鄰相接
  - [x] 7.5.2 第二列 (Row 2) 重新排列為：`遠端初始路徑`、`[瀏覽遠端...]` 按鈕、`記住連線資訊` 核取方塊、`[測試連線]` 按鈕
  - [x] 7.5.3 規則區移除 `[➕ 新增一列]` 與 `[➖ 刪除選取]` 按鈕
  - [x] 7.5.4 規則區新增 `比對模式` 下拉選單（選項：`模糊包含 (Contains)`、`完全精準相符 (Exact)`，預設模糊包含）
  - [x] 7.5.5 執行控制列精簡為單一醒目按鈕 `[📑 開啟下載清單 (Excel)]`，移除手動匯出按鈕
- [x] 7.6 下載管理器動態報表輸出 (`core/download_manager.py`)
  - [x] 7.6.1 擴充 `generate_manifest_csv`：動態識別多階層目錄，寫入完整 10+ 欄位 UTF-8-SIG CSV 檔案
- [x] 7.7 【動態升級全項驗收測試】
  - [x] 7.7.1 撰寫並執行測試腳本 `tests/test_phase7_dynamic.py`：驗證動態多層欄位適配、精準/模糊比對、遠端瀏覽器邏輯與介面配置
  - [x] 7.7.2 執行完整回歸測試 `tests/test_all.py`，確保舊功能無任何破壞

---

## 微階段 8：純目錄動態階層表格與最後目錄檔案過濾 (Pure Directory Hierarchy & Final Level Filtering)
- [x] 8.1 規則表格全欄位純目錄化 (`ui/rule_table_ctk.py`)
  - [x] 8.1.1 改造 `parse_tsv_text(raw_text, dynamic=True)`：Excel 貼上幾欄即完全對應幾層資料夾，不再將最後一欄抽成檔案名稱
  - [x] 8.1.2 改造 `set_columns(dir_count)`：所有欄位標題一律為 `第 1 層目錄`、`第 2 層目錄`、...、`第 N 層目錄`，完全移除檔案名稱欄位
  - [x] 8.1.3 表頭標題精簡化，移除干擾閱讀的括號文字，呈現純淨的 `第 N 層目錄`
  - [x] 8.1.4 改造 `get_all_rules()` 與 `load_rules()`：整行所有數值純粹作為目錄清單 `dirs: List[str]`
- [x] 8.2 遍歷引擎調整為抵達最後一層目錄執行檔案搜尋 (`core/traverse_engine.py`)
  - [x] 8.2.1 遍歷推進直至該規則之最後一層目錄（包含動態 1 層、2 層、3 層... N 層目錄）
  - [x] 8.2.2 在抵達最後一層目錄後，統一讀取 `SearchSettings.global_pattern`（如 `*RAW*`、`*RAW*-*.csv`）並匹配該目錄下的檔案
  - [x] 8.2.3 保持實體目錄快取機制，相同目錄在任何深度僅呼叫 1 次 `listdir_attr`
- [x] 8.3 主視窗設定標籤與輸入驗證調整 (`ui/main_window.py`)
  - [x] 8.3.1 將「全域檔案過濾」標籤更名為「目標檔案過濾:」，placeholder 標示為 `例: *RAW*; *RAW*-*.csv (在最後一層目錄搜尋，留空代表全選)`
  - [x] 8.3.2 調整 `_validate_inputs`：允許使用者輸入任意純目錄階層，只要目錄有填寫即放行
- [x] 8.4 【純目錄架構全項驗收測試】
  - [x] 8.4.1 撰寫並執行測試腳本 `tests/test_phase8_pure_dirs.py`：驗證貼上 2 欄即為 2 層純目錄、貼上 3 欄即為 3 層純目錄，並在最後目錄成功檢索目標檔案
  - [x] 8.4.2 執行完整回歸測試 `tests/test_all.py`，確保 7 大測試全部通過

---

## 微階段 9：PyInstaller 打包發布與一鍵編譯 (Packaging & Distribution)
- [x] 9.1 建立 PyInstaller 資源設定檔 (`SFTPTool.spec`)
  - [x] 9.1.1 收集 CustomTkinter 所有 JSON 主題與字型資產 (`collect_all('customtkinter')`)
  - [x] 9.1.2 收集 Paramiko 通訊協定與暗號學依賴 (`collect_all('paramiko')`)
  - [x] 9.1.3 設定無主控台模式 (`console=False`) 隱藏黑色 CMD 視窗
  - [x] 9.1.4 設定輸出目標名稱為 `SFTP批次下載工具.exe`
- [x] 9.2 建立一鍵編譯批次檔 (`build_exe.bat`)
  - [x] 9.2.1 純 ASCII 腳本設計，防範 Windows CP950 編碼問題
  - [x] 9.2.2 自動檢測與安裝 `pyinstaller`
  - [x] 9.2.3 自動清理舊 `build/` 與 `dist/` 快取目錄
  - [x] 9.2.4 執行打包並驗證 `dist/SFTP批次下載工具.exe` 產出成功
- [x] 9.3 執行打包並進行 EXE 執行期驗證
  - [x] 9.3.1 執行編譯程序產出 EXE
  - [x] 9.3.2 驗證 EXE 啟動與資產載入正常

---

## 微階段 10：綠色免安裝便攜包一鍵建置與壓縮 (Portable Distribution)
- [x] 10.1 建立便攜包建置腳本 (`build_portable.bat`, `build_portable.py`)
  - [x] 10.1.1 純 ASCII 腳本設計，防範 Windows CP950 編碼問題
  - [x] 10.1.2 自動建立獨立 Python runtime 執行環境（包含官方微軟/PSF 簽署之 `_tkinter.pyd`、`tcl/tk` 與標準庫）
  - [x] 10.1.3 自動安裝依賴套件（`customtkinter`, `paramiko`, `openpyxl`, `darkdetect`）至便攜環境
  - [x] 10.1.4 自動複製專案模組程式碼 (`main.py`, `models/`, `core/`, `ui/`)
- [x] 10.2 建立便攜包啟動檔 (`啟動工具.bat`)
  - [x] 10.2.1 採用 `pythonw.exe` 執行 `main.py`，達成無黑視窗背景啟動
  - [x] 10.2.2 附帶使用說明文件
- [x] 10.3 自動壓製發布用 ZIP 包
  - [x] 10.3.1 透過 Python/PowerShell 自動將便攜資料夾壓縮為 `dist/SFTP批次下載工具_綠色便攜版.zip`
  - [x] 10.3.2 驗證便攜包結構完整與啟動配置正確

---

## 微階段 11：GitHub 專案發布與說明文件建置 (Git & GitHub Deployment)
- [x] 11.1 撰寫專案完整說明文件 (`README.md`)
  - [x] 11.1.1 撰寫專案簡介與核心特點（動態純目錄階層、最後目錄檔案檢索、遠端路徑瀏覽器、Excel 貼上自適應、Windows 11 SAC 免疫）
  - [x] 11.1.2 撰寫系統架構與相依套件說明（Python 3.10+、CustomTkinter、Paramiko、OpenPyXL）
  - [x] 11.1.3 撰寫快速開始與開發環境安裝教學（`start.bat` 一鍵啟動）
  - [x] 11.1.4 撰寫打包發布教學（`build_exe.bat` 單檔 EXE 編譯與 `build_portable.bat` 綠色便攜包製作）
  - [x] 11.1.5 撰寫操作說明與 Excel 貼上對照範例（多層目錄貼上格式、比對模式切換、目標檔案萬用字元過濾）
- [x] 11.2 Git 版本控制初始化與配置
  - [x] 11.2.1 檢查 `.gitignore` 確保排除了 `build/`、`dist/`、`__pycache__/`、`config.json` 等暫存與隱私檔案
  - [x] 11.2.2 執行 `git init` 初始化本地儲存庫
  - [x] 11.2.3 建立/切換至 `main` 主要分支 (`git branch -M main`)
  - [x] 11.2.4 設定遠端儲存庫位址 `https://github.com/powpow088/FTPtool_byList.git`
- [x] 11.3 提交與推送代碼至 GitHub
  - [x] 11.3.1 暫存所有原始碼、測試、說明文件與建置腳本 (`git add .`)
  - [x] 11.3.2 建立初始 Commit (`git commit -m "..."`)
  - [x] 11.3.3 推送代碼至遠端儲存庫 `origin main`
  - [x] 11.3.4 驗證推送結果與遠端同步狀態

---

## 微階段 12：純檢索模式 (Scan-Only) 與檔案修改時間/副檔名實作 (Scan-Only & File Metadata)
- [x] 12.1 擴充資料模型層 (`models/data_models.py`)
  - [x] 12.1.1 擴充 `TransferStatus(Enum)`：新增 `INDEXED = "已檢索 (未下載)"`、`EMPTY_DIR = "空目錄 (無相符檔案)"`、`DIR_NOT_FOUND = "目錄不存在"`
  - [x] 12.1.2 擴充 `ReportRecord` 資料類別：新增 `file_mtime: int = 0` 屬性儲存遠端 Unix 時間戳記
  - [x] 12.1.3 實作 `ReportRecord.mtime_str` 計算屬性：將 `file_mtime` 格式化為 `YYYY-MM-DD HH:MM:SS` 字串（若無時間則為空字串）
  - [x] 12.1.4 實作 `ReportRecord.file_ext` 計算屬性：提取小寫副檔名（如 `.csv`、`.txt`）
- [x] 12.2 擴充報表生成器與純檢索清單封裝 (`core/download_manager.py`)
  - [x] 12.2.1 升級 `generate_manifest_csv` 表頭與欄位：納入 `副檔名`、`檔案修改時間`、`處理狀態`、`處理說明`、`遠端完整路徑`、`本地完整路徑`、`完成時間` 等 12 大欄位
  - [x] 12.2.2 支援純檢索清冊命名規則：純檢索模式下檔名產生為 `檢索清單_YYYYMMDD_HHMMSS.csv`，一般下載維持 `下載清單_YYYYMMDD_HHMMSS.csv`
  - [x] 12.2.3 實作 `generate_scan_manifest_csv()`：接收檢索出的檔案任務清單與空目錄清單，直接批次寫入 CSV 報表
- [x] 12.3 走訪引擎記錄空目錄與未匹配目錄 (`core/traverse_engine.py`)
  - [x] 12.3.1 在 `scan_all()` 中追蹤每條目錄規則之終端目錄走訪結果
  - [x] 12.3.2 若終端目錄無相符檔案，產生標記為 `EMPTY_DIR` 之紀錄物件
  - [x] 12.3.3 若遠端無法抵達該目錄，產生標記為 `DIR_NOT_FOUND` 之紀錄物件
  - [x] 12.3.4 保證所有檢索出的 `RemoteFileTask` 均帶有遠端 `entry.st_mtime` 數值
- [x] 12.4 背景工作執行緒支援純檢索模式 (`core/worker.py`)
  - [x] 12.4.1 `DownloadWorker` 構造函式新增 `scan_only: bool = False` 參數
  - [x] 12.4.2 在 `scan_only=True` 模式下：走訪目錄後直接產出 `檢索清單_*.csv`，跳過實際下載傳輸迴圈
  - [x] 12.4.3 統計純檢索總檔案數與總容量大小，並透過 `UIEvent` 傳送給介面
- [x] 12.5 主視窗 UI 整合與按鈕狀態連動 (`ui/main_window.py`)
  - [x] 12.5.1 在底部控制按鈕列新增獨立按鈕 `[🔍 僅搜尋與產出清單]`
  - [x] 12.5.2 實作 `on_btn_scan_only_click`：檢查主機與規則設定，啟動 `scan_only=True` 背景執行緒
  - [x] 12.5.3 檢索期間停用輸入欄位與下載按鈕，啟用停止按鈕
  - [x] 12.5.4 檢索完成後於狀態列更新「檢索完成！共發現 X 個檔案 (總計 Y)，清單已生成」
  - [x] 12.5.5 自動記錄產出的最新清單絕對路徑，並解鎖 `[📑 開啟下載清單 (Excel)]` 一鍵以 Excel 開啟
- [x] 12.6 單元測試與回歸驗證
  - [x] 12.6.1 撰寫並執行測試腳本 `tests/test_phase9_scan_only.py`：驗證純檢索模式下無實體檔案下載、CSV 包含修改時間與副檔名、空目錄狀態正確呈現
  - [x] 12.6.2 執行全套回歸測試 `tests/test_all.py`，確保 8 大模組測試 100% 通過
  - [x] 12.6.3 重新執行 `build_exe.bat` 與 `build_portable.bat` 更新發布包
  - [x] 12.6.4 提交代碼至 Git 並推送到 GitHub 遠端儲存庫

---

## 微階段 13：純檢索清單移除本地路徑欄位優化 (Scan-Only Manifest Optimization)
- [x] 13.1 下載管理器 CSV 產出方法支援欄位條件切換 (`core/download_manager.py`)
  - [x] 13.1.1 `generate_manifest_csv` 新增 `include_local_path: bool = True` 參數
  - [x] 13.1.2 當 `include_local_path=False` 時，表頭排除「本地完整路徑」欄位
  - [x] 13.1.3 當 `include_local_path=False` 時，資料列排除 `r.local_path` 數值
  - [x] 13.1.4 `generate_scan_manifest_csv` 呼叫 `generate_manifest_csv` 時明確傳入 `include_local_path=False`
- [x] 13.2 規格書與說明文件同步更新 (`docs/spec.md`, `README.md`)
  - [x] 13.2.1 更新 `docs/spec.md` 3.3 章節說明純檢索清單不含本地路徑欄位
  - [x] 13.2.2 更新 `README.md` 清單欄位對照說明
- [x] 13.3 單元測試與回歸驗證
  - [x] 13.3.1 更新 `tests/test_phase9_scan_only.py`：驗證 `檢索清單_*.csv` 表頭無「本地完整路徑」欄位，純輸出遠端資訊
  - [x] 13.3.2 執行全套回歸測試 `tests/test_all.py`，確保全數 8 大模組測試 100% 通過
  - [x] 13.3.3 重新執行 `build_portable.bat` 與 `build_exe.py` 更新發布包
  - [x] 13.3.4 提交代碼至 Git 並推送到 GitHub 遠端儲存庫

---

## 微階段 14：生產環境高可用連線防護、防封鎖微延遲與續傳接關強化 (High-Availability & Anti-Ban Protection)
- [x] 14.1 擴充資料模型與設定項 (`models/data_models.py`)
  - [x] 14.1.1 在 `SearchSettings` 新增 `request_delay_ms: int = 50` 屬性（預設 50ms 溫和下載間隔）
  - [x] 14.1.2 在 `SearchSettings` 新增 `session_refresh_interval: int = 150` 屬性（每 150 個檔案主動換氣刷新連線）
- [x] 14.2 強化 SFTP 連線管理器 (`core/sftp_manager.py`)
  - [x] 14.2.1 將 `SOCKET_TIMEOUT` 從 15.0 秒調升為 30.0 秒，適應尖峰時段網路波動
  - [x] 14.2.2 實作 `is_active() -> bool` 方法：即時檢測 SSH Transport 是否存活且 SFTP Channel 未關閉
  - [x] 14.2.3 實作 `ensure_connected()` 方法：若連線已失效自動在背景呼叫 `reconnect()` 重建連線
- [x] 14.3 強化下載管理器與防崩潰機制 (`core/download_manager.py`)
  - [x] 14.3.1 在 `process_task` 進入前呼叫 `ensure_connected()`，杜絕「單一斷線導致後續所有檔案秒失敗」的雪崩效應
  - [x] 14.3.2 實作檔案下載完成後的防封鎖微延遲（支援中斷訊號感知）
  - [x] 14.3.3 實作「主動連線換氣 (Proactive Session Refresh)」機制：累積下載達 150 個檔案時主動優雅重置 Session，徹底避開伺服器/防火牆 10~15 分鐘強制中斷與 Rekey 異常
  - [x] 14.3.4 升級斷線重試退避演算法（等待時間從 1s 改為 3s、6s、12s），避免短時間密集重連觸發伺服器 Fail2ban 臨時封鎖 IP
  - [x] 14.3.5 強化略過訊息提示：當檔案命中 `SIZE_MTIME` 時，日誌明確標示 `[續傳略過]`
- [x] 14.4 主視窗 UI 新增下載頻率選單與續傳日誌 (`ui/main_window.py`)
  - [x] 14.4.1 在下載設定區新增「傳輸間隔:」下拉選單（`極速 (0ms)`、`穩定推薦 (50ms)`、`溫和防封 (150ms)`）
  - [x] 14.4.2 將選擇的間隔數值連動至 `SearchSettings.request_delay_ms`
  - [x] 14.4.3 持久化儲存與載入傳輸間隔設定至 `config.json`
- [x] 14.5 單元測試與回歸驗證
  - [x] 14.5.1 撰寫並執行測試腳本 `tests/test_phase10_connection_guard.py`：驗證連線健康預檢、防封鎖微延遲、主動換氣計數器、以及退避時間增長
  - [x] 14.5.2 執行全套回歸測試 `python tests/test_all.py`，確保 9 大模組測試 100% 通過
  - [x] 14.5.3 重新執行 `build_portable.bat` 與 `build_exe.py` 更新發布包
  - [x] 14.5.4 提交代碼至 Git 並推送到 GitHub 遠端儲存庫

---

## 微階段 14：生產環境高可用連線防護、防封鎖微延遲與續傳接關強化 (High-Availability & Anti-Ban Protection)
- [x] 14.1 擴充資料模型與設定項 (`models/data_models.py`)
  - [x] 14.1.1 在 `SearchSettings` 新增 `request_delay_ms: int = 50` 屬性（預設 50ms 溫和下載間隔）
  - [x] 14.1.2 在 `SearchSettings` 新增 `session_refresh_interval: int = 150` 屬性（每 150 個檔案主動換氣刷新連線）
- [x] 14.2 強化 SFTP 連線管理器 (`core/sftp_manager.py`)
  - [x] 14.2.1 將 `SOCKET_TIMEOUT` 從 15.0 秒調升為 30.0 秒，適應尖峰時段網路波動
  - [x] 14.2.2 實作 `is_active() -> bool` 方法：即時檢測 SSH Transport 是否存活且 SFTP Channel 未關閉
  - [x] 14.2.3 實作 `ensure_connected()` 方法：若連線已失效自動在背景呼叫 `reconnect()` 重建連線
- [x] 14.3 強化下載管理器與防崩潰機制 (`core/download_manager.py`)
  - [x] 14.3.1 在 `process_task` 進入前呼叫 `ensure_connected()`，杜絕「單一斷線導致後續所有檔案秒失敗」的雪崩效應
  - [x] 14.3.2 實作檔案下載完成後的防封鎖微延遲（支援中斷訊號感知）
  - [x] 14.3.3 實作「主動連線換氣 (Proactive Session Refresh)」機制：累積下載達 150 個檔案時主動優雅重置 Session，徹底避開伺服器/防火牆 10~15 分鐘強制中斷與 Rekey 異常
  - [x] 14.3.4 升級斷線重試退避演算法（等待時間從 1s 改為 3s、6s、12s），避免短時間密集重連觸發伺服器 Fail2ban 臨時封鎖 IP
  - [x] 14.3.5 強化略過訊息提示：當檔案命中 `SIZE_MTIME` 時，日誌明確標示 `[續傳略過]`
- [x] 14.4 主視窗 UI 新增下載頻率選單與續傳日誌 (`ui/main_window.py`)
  - [x] 14.4.1 在下載設定區新增「傳輸頻率:」下拉選單（`極速 (0ms)`、`穩定推薦 (50ms)`、`溫和防封 (150ms)`）
  - [x] 14.4.2 將選擇的間隔數值連動至 `SearchSettings.request_delay_ms`
  - [x] 14.4.3 持久化儲存與載入傳輸間隔設定至 `config.json`
- [x] 14.5 單元測試與回歸驗證
  - [x] 14.5.1 撰寫並執行測試腳本 `tests/test_phase10_connection_guard.py`：驗證連線健康預檢、防封鎖微延遲、主動換氣計數器、以及退避時間增長
  - [x] 14.5.2 執行全套回歸測試 `python tests/test_all.py`，確保 9 大模組測試 100% 通過
  - [x] 14.5.3 重新執行 `build_portable.bat` 與 `build_exe.py` 更新發布包
  - [x] 14.5.4 提交代碼至 Git 並推送到 GitHub 遠端儲存庫

---

## 微階段 15：CSV 報表成對搜尋關鍵字欄位實作 (Matched Search Keywords in Manifest CSV - Option A)
- [x] 15.1 擴充資料模型層 (`models/data_models.py`)
  - [x] 15.1.1 在 `RemoteFileTask` 類別新增 `search_dirs: List[str] = field(default_factory=list)` 屬性
  - [x] 15.1.2 在 `RemoteFileTask.__init__` 新增 `search_dirs: Optional[List[str]] = None` 參數並正確初始化
  - [x] 15.1.3 在 `ReportRecord` 類別新增 `search_dirs: List[str] = field(default_factory=list)` 屬性
  - [x] 15.1.4 在 `ReportRecord.__init__` 新增 `search_dirs: Optional[List[str]] = None` 參數並正確初始化
- [x] 15.2 走訪引擎記錄匹配之目錄搜尋字 (`core/traverse_engine.py`)
  - [x] 15.2.1 在 `scan_all()` 的路徑追蹤中保存規則原始搜尋字元組 `orig_dirs_tuple`
  - [x] 15.2.2 檢索出檔案建立 `RemoteFileTask` 時，將命中規則的 `orig_dirs_tuple` 轉為列表賦值給 `task.search_dirs`
  - [x] 15.2.3 產生目錄不存在紀錄 (`DIR_NOT_FOUND`) 時，將原始規則搜尋字寫入 `ReportRecord.search_dirs`
  - [x] 15.2.4 產生終端空目錄紀錄 (`EMPTY_DIR`) 時，將命中路徑對應之規則搜尋字寫入 `ReportRecord.search_dirs`
- [x] 15.3 擴充報表生成器成對輸出欄位 (`core/download_manager.py`)
  - [x] 15.3.1 在 `_make_record()` 中將 `task.search_dirs` 傳遞賦值給 `ReportRecord.search_dirs`
  - [x] 15.3.2 升級 `generate_manifest_csv()` 動態計算最大層數：取實際目錄數與搜尋字數之最大值（至少 2 層）
  - [x] 15.3.3 升級 CSV 表頭結構：改為成對排列 `第 1 層目錄`, `第 1 層搜尋字`, `第 2 層目錄`, `第 2 層搜尋字` ...
  - [x] 15.3.4 升級 CSV 資料列結構：依照目錄層級成對寫入實際目錄名稱與對應搜尋字
  - [x] 15.3.5 升級 `generate_scan_manifest_csv()`：將 `task.search_dirs` 正確傳入純檢索清單紀錄
- [x] 15.4 單元測試與發布回歸驗證
  - [x] 15.4.1 更新既有測試案例（`tests/test_phase3.py`, `tests/test_phase9_scan_only.py`, `tests/test_phase10_connection_guard.py`）之 CSV 表頭斷言
  - [x] 15.4.2 撰寫全新測試模組 `tests/test_phase11_paired_keywords.py`：驗證成對欄位輸出、模糊包含對照、空目錄與未匹配目錄搜尋字保留
  - [x] 15.4.3 執行全套回歸測試 `python tests/test_all.py`，確保 10 大模組 100% 通過
  - [x] 15.4.4 重新執行 `build_exe.py` 與 `build_portable.py` 編譯單檔 EXE 與打包便攜綠色版 ZIP
  - [x] 15.4.5 提交代碼至 Git 並推送到 GitHub 遠端儲存庫

---

## 微階段 16：晶片明細雙工作表活頁簿產出與底部操作按鈕重構 (Wafer Detail Aggregation & UI Button Restructure)
- [x] 16.1 建立明細穿透解析與 Excel 彙整核心模組 (`core/detail_aggregator.py`)
  - [x] 16.1.1 實作 `parse_recipe(level1_dir: str) -> str`：依據第二個連字號 `-` 提取字串，未滿兩個 `-` 則回傳空字串
  - [x] 16.1.2 實作 `read_text_auto_encoding(file_path: str) -> str`：自動辨識 UTF-8 BOM、標準 UTF-8 及 Windows CP950/ANSI 編碼
  - [x] 16.1.3 實作 `find_header_line(lines: List[str]) -> Tuple[int, int, int, int, int, str]`：前 10 列中動態定位含有 `2D ID`, `X`, `Y`, `UNIT BIN#` 之標題列，自動辨識逗號或 Tab 分隔符號
  - [x] 16.1.4 實作單一檔案明細資料逐列剖析器：排除第一欄為 `Hard Limit` 之列，組合出 `MPID_X_Y`，擷取 `Unit BIN#`
  - [x] 16.1.5 實作檔案容錯與略過機制：遇到空檔或未含必要標題之檔案，記錄錯誤說明並跳過，不中斷整批處理
  - [x] 16.1.6 實作雙工作表活頁簿建置器：
    - Sheet 1 (`下載清單`)：讀入當前下載清單 CSV，複製所有欄位與資料列，設定表頭加粗與自動欄寬
    - Sheet 2 (`工作表2`)：寫入 11 欄明細資料，A~J 欄強制設定純文字格式 `@`，K 欄設定為數值/文字格式，表頭加粗並嚴格按照原 VBA 指定欄寬（A:24, B:13, C:22, D:22, E:12, F:62, G:21, H:13, I:13, J:28, K:13）
  - [x] 16.1.7 實作 100 萬列自動切頁機制：明細資料列達到 1,000,000 列時自動接續建立 `工作表2_2`、`工作表2_3`...，表頭與格式相同
  - [x] 16.1.8 檔案命名固定為 `下載與明細彙整_YYYYMMDD_HHMMSS.xlsx` 並自動儲存於本地目標儲存路徑根目錄
- [x] 16.2 介面按鈕順序重構與三態狀態引擎 (`ui/main_window.py`)
  - [x] 16.2.1 重新排列底部操作按鈕，由左至右嚴格順序為：`[▶ 開始下載]`、`[📊 彙整資料]`、`[📑 開啟下載清單 (Excel)]`、`[📁 開啟下載資料夾]`、`[🔍 僅產出清單]`、`[⏹ 停止]`
  - [x] 16.2.2 簡化按鈕文字：將原「開始搜尋與下載」簡化為「▶ 開始下載」，原「僅搜尋與產出清單」簡化為「🔍 僅產出清單」
  - [x] 16.2.3 將「⏹ 停止」按鈕移至最右側並維持紅色醒目配色
  - [x] 16.2.4 新增「[📊 彙整資料]」按鈕，並實作三態狀態引擎 `_update_aggregate_btn_state()`：
    - 狀態 1 (不可執行)：本地儲存路徑下無 `下載清單_*.csv` 且無當前完成任務時，按鈕設為停用 (Disabled, 灰色)
    - 狀態 2 (可執行)：任務剛下載完成或本地存在清單且尚未產出對應明細檔時，按鈕設為啟用並呈現醒目亮橘/紫色 (Highlight)
    - 狀態 3 (已有檔案)：本地儲存路徑下已存在對應之 `下載與明細彙整_*.xlsx` 時，按鈕呈現綠色完成態並提示可點擊重新產出或開啟
  - [x] 16.2.5 實作非同步彙整事件處理 `_on_aggregate_data_task()`：以背景執行緒解析，不卡死 UI 介面，並同步輸出解析進度至即時日誌
  - [x] 16.2.6 彙整完成彈跳對話框：「彙整完成！共寫入 X 筆晶片資料，是否立即開啟 Excel？」，點「是」以 `os.startfile` 開啟 `.xlsx`，點「否」留在原畫面
- [x] 16.3 單元測試與自動化驗證
  - [x] 16.3.1 撰寫全新測試模組 `tests/test_phase12_detail_aggregation.py`：驗證 RECIPE 萃取、前 10 行標題辨識、逗號/Tab 分隔、Hard Limit 略過、MPID_X_Y 合成、雙 Sheet 生成、100 萬列切頁與無效檔案容錯
  - [x] 16.3.2 驗證 `ui/main_window.py` 6 顆按鈕順序與三態狀態變化邏輯
  - [x] 16.3.3 將測試 12 納入 `tests/test_all.py`，確保 11 大模組 100% 通過
  - [x] 16.3.4 重新執行 `build_exe.py` 與 `build_portable.py` 編譯單檔 EXE 與綠色免安裝便攜 ZIP 包
  - [x] 16.3.5 提交代碼至 Git 並推送到 GitHub 遠端儲存庫

---

## 微階段 17：規則表格一鍵複製批號字串功能實作 (One-Click Distinct Lot Copying to Clipboard)
- [x] 17.1 核心批號提取與過濾演算法實作 (`ui/main_window.py`)
  - [x] 17.1.1 實作 `extract_distinct_lots(rules: List[RuleItem]) -> List[str]` 靜態方法
  - [x] 17.1.2 支援動態層級批號定位：目錄層級 >= 2 優先提取第 2 層目錄 (`dirs[1]`)，單層目錄提取第 1 層目錄 (`dirs[0]`)
  - [x] 17.1.3 智慧過濾標題列：排除空值與包含 `批號`, `lot`, `lot no`, `lot no.`, `lot id`, `lot_id`, `lot number` 等不區分大小寫之標題文字
  - [x] 17.1.4 實作有序去重 (Distinct)：使用 `dict.fromkeys()` 保留原始資料首次出現順序，去除重複批號
  - [x] 17.1.5 格式化串接：使用純逗點 `,` 無空格進行串接 (`','.join(...)`)
- [x] 17.2 介面按鈕整合與剪貼簿操作 (`ui/main_window.py`)
  - [x] 17.2.1 在規則操作列之比對模式選單右側（使用者圈選位置）新增 `btn_copy_lots` 按鈕，文字標示為 `[📋 複製批號 (,)]`
  - [x] 17.2.2 按鈕配色採用紫色高質感樣式 (`fg_color="#6F42C1"`, `hover_color="#59359A"`)，與相鄰按鈕視覺區隔
  - [x] 17.2.3 實作 `_on_copy_lots_to_clipboard()` 事件處理：安全寫入 Windows 系統剪貼簿 (`clipboard_clear` / `clipboard_append` / `update`)
  - [x] 17.2.4 實作操作回饋：成功時於即時日誌視窗以 `SUCCESS` 標籤輸出「已成功複製 X 筆批號至剪貼簿（純逗號分隔，已自動去重）！」，若無有效資料則彈出提示
  - [x] 17.2.5 在 `_set_running_state(self, running: bool)` 中將 `btn_copy_lots` 納入執行狀態連動（任務執行中設為停用）
- [x] 17.3 單元測試與回歸驗證
  - [x] 17.3.1 撰寫全新測試模組 `tests/test_phase13_copy_lots.py`：驗證單層與多層批號提取、標題列自動過濾、有序去重、純逗號串接與按鈕屬性
  - [x] 17.3.2 將測試 13 納入 `tests/test_all.py`，確保 12 大單元與整合測試 100% 通過
  - [x] 17.3.3 重新執行 `build_exe.py` 與 `build_portable.py` 編譯最新單檔 EXE 與綠色便攜 ZIP 包
  - [x] 17.3.4 提交代碼至 Git 並推送到 GitHub 遠端儲存庫
