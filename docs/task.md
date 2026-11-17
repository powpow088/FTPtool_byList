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
- [ ] 11.1 撰寫專案完整說明文件 (`README.md`)
  - [ ] 11.1.1 撰寫專案簡介與核心特點（動態純目錄階層、最後目錄檔案檢索、遠端路徑瀏覽器、Excel 貼上自適應、Windows 11 SAC 免疫）
  - [ ] 11.1.2 撰寫系統架構與相依套件說明（Python 3.10+、CustomTkinter、Paramiko、OpenPyXL）
  - [ ] 11.1.3 撰寫快速開始與開發環境安裝教學（`start.bat` 一鍵啟動）
  - [ ] 11.1.4 撰寫打包發布教學（`build_exe.bat` 單檔 EXE 編譯與 `build_portable.bat` 綠色便攜包製作）
  - [ ] 11.1.5 撰寫操作說明與 Excel 貼上對照範例（多層目錄貼上格式、比對模式切換、目標檔案萬用字元過濾）
- [ ] 11.2 Git 版本控制初始化與配置
  - [ ] 11.2.1 檢查 `.gitignore` 確保排除了 `build/`、`dist/`、`__pycache__/`、`config.json` 等暫存與隱私檔案
  - [ ] 11.2.2 執行 `git init` 初始化本地儲存庫
  - [ ] 11.2.3 建立/切換至 `main` 主要分支 (`git branch -M main`)
  - [ ] 11.2.4 設定遠端儲存庫位址 `https://github.com/powpow088/FTPtool_byList.git`
- [ ] 11.3 提交與推送代碼至 GitHub
  - [ ] 11.3.1 暫存所有原始碼、測試、說明文件與建置腳本 (`git add .`)
  - [ ] 11.3.2 建立初始 Commit (`git commit -m "..."`)
  - [ ] 11.3.3 推送代碼至遠端儲存庫 `origin main`
  - [ ] 11.3.4 驗證推送結果與遠端同步狀態
