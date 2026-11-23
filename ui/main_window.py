import csv
import datetime
import os
import queue
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import List, Optional

import customtkinter as ctk

from models.data_models import (
    SFTPConfig, SearchSettings, ConflictStrategy, TaskStats, ReportRecord, RuleItem, MatchMode
)
from core.config_manager import ConfigManager
from core.worker import ConnectionTestWorker, DownloadWorker, UIEvent
from core.detail_aggregator import DetailAggregator
from ui.rule_table_ctk import CTkRuleTable
from ui.remote_browser_dialog import RemoteBrowserDialog


class MainWindow(ctk.CTk):
    """主視窗 GUI (CustomTkinter 版)，四大功能區塊、彩色 Tag 日誌視窗與 80ms 輪詢器"""

    LOG_COLORS = {
        "INFO": "#3399FF",
        "SUCCESS": "#28A745",
        "WARNING": "#FFC107",
        "ERROR": "#FF4D4D"
    }

    def __init__(self) -> None:
        super().__init__()
        self.title("SFTP 批次檔案模糊檢索與結構保留下載工具")
        self.geometry("1080x820")
        self.minsize(980, 720)

        self._ui_queue: queue.Queue[UIEvent] = queue.Queue()
        self._test_worker: Optional[ConnectionTestWorker] = None
        self._dl_worker: Optional[DownloadWorker] = None
        self._last_records: List[ReportRecord] = []
        self._last_manifest_csv: Optional[str] = None

        self._init_ui()
        self._load_saved_state()

        # 啟動 80ms 跨執行緒 UI 事件輪詢器 (12.5 FPS 節流)
        self._after_id: Optional[str] = self.after(80, self._process_ui_queue)
        self.protocol("WM_DELETE_WINDOW", self._on_close_window)

    def _init_ui(self) -> None:
        self.grid_rowconfigure(1, weight=3) # 表格區
        self.grid_rowconfigure(3, weight=2) # 日誌區
        self.grid_columnconfigure(0, weight=1)

        # ---------------- 1. SFTP 連線設定區 ----------------
        conn_frame = ctk.CTkFrame(self)
        conn_frame.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 5))
        conn_frame.grid_columnconfigure(1, weight=3)
        conn_frame.grid_columnconfigure(5, weight=2)
        conn_frame.grid_columnconfigure(7, weight=2)

        ctk.CTkLabel(conn_frame, text="1. SFTP 連線設定", font=("Microsoft JhengHei UI", 13, "bold")).grid(
            row=0, column=0, columnspan=8, sticky="w", padx=10, pady=(5, 5)
        )

        # 第一列 (Row 1): 主機、Port、帳號、密碼相鄰
        ctk.CTkLabel(conn_frame, text="主機 (Host):").grid(row=1, column=0, sticky="e", padx=(8, 2), pady=4)
        self.ed_host = ctk.CTkEntry(conn_frame, placeholder_text="例: 192.168.1.100")
        self.ed_host.grid(row=1, column=1, sticky="ew", padx=3, pady=4)

        ctk.CTkLabel(conn_frame, text="Port:").grid(row=1, column=2, sticky="e", padx=(6, 2), pady=4)
        self.ed_port = ctk.CTkEntry(conn_frame, width=55)
        self.ed_port.insert(0, "22")
        self.ed_port.grid(row=1, column=3, sticky="w", padx=3, pady=4)

        ctk.CTkLabel(conn_frame, text="帳號:").grid(row=1, column=4, sticky="e", padx=(8, 2), pady=4)
        self.ed_user = ctk.CTkEntry(conn_frame)
        self.ed_user.grid(row=1, column=5, sticky="ew", padx=3, pady=4)

        ctk.CTkLabel(conn_frame, text="密碼:").grid(row=1, column=6, sticky="e", padx=(8, 2), pady=4)
        pass_box = ctk.CTkFrame(conn_frame, fg_color="transparent")
        pass_box.grid(row=1, column=7, sticky="ew", padx=(3, 8), pady=4)
        pass_box.grid_columnconfigure(0, weight=1)

        self.ed_pass = ctk.CTkEntry(pass_box, show="*")
        self.ed_pass.grid(row=0, column=0, sticky="ew")
        self.btn_toggle_pass = ctk.CTkButton(pass_box, text="👁", width=34, height=28, command=self._on_toggle_password)
        self.btn_toggle_pass.grid(row=0, column=1, padx=(3, 0))

        # 第二列 (Row 2): 遠端初始路徑、[瀏覽遠端...]、記住連線資訊、測試連線
        ctk.CTkLabel(conn_frame, text="遠端初始路徑:").grid(row=2, column=0, sticky="e", padx=(8, 2), pady=4)
        self.ed_init_path = ctk.CTkEntry(conn_frame, placeholder_text="/")
        self.ed_init_path.insert(0, "/")
        self.ed_init_path.grid(row=2, column=1, sticky="ew", padx=3, pady=4)

        self.btn_browse_remote = ctk.CTkButton(
            conn_frame, text="瀏覽遠端...", width=95, height=28, fg_color="#17A2B8", hover_color="#138496",
            command=self._on_browse_remote
        )
        self.btn_browse_remote.grid(row=2, column=2, columnspan=2, sticky="w", padx=3, pady=4)

        self.chk_save_pass = ctk.CTkCheckBox(conn_frame, text="記住連線資訊")
        self.chk_save_pass.grid(row=2, column=5, sticky="w", padx=8, pady=4)

        self.btn_test_conn = ctk.CTkButton(conn_frame, text="測試連線", width=100, height=28, command=self._on_test_connection)
        self.btn_test_conn.grid(row=2, column=7, sticky="ew", padx=(3, 8), pady=4)

        # ---------------- 2. 動態過濾規則表格區 ----------------
        rule_frame = ctk.CTkFrame(self)
        rule_frame.grid(row=1, column=0, sticky="nsew", padx=12, pady=5)
        rule_frame.grid_rowconfigure(1, weight=1)
        rule_frame.grid_columnconfigure(0, weight=1)

        bar_frame = ctk.CTkFrame(rule_frame, fg_color="transparent")
        bar_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=(5, 2))

        ctk.CTkLabel(
            bar_frame,
            text="2. 階層過濾規則 (支援 Excel Ctrl+V 貼上多層資料夾)",
            font=("Microsoft JhengHei UI", 13, "bold")
        ).pack(side="left", padx=5)

        self.btn_paste_excel = ctk.CTkButton(
            bar_frame, text="📋 貼上 Excel 清單", width=130, command=lambda: self.table_rules.paste_from_clipboard()
        )
        self.btn_paste_excel.pack(side="left", padx=4)

        self.btn_import_file = ctk.CTkButton(
            bar_frame, text="📂 從檔案匯入", width=100, command=self._on_import_rules_file
        )
        self.btn_import_file.pack(side="left", padx=4)

        self.btn_clear_table = ctk.CTkButton(
            bar_frame, text="🗑️ 清空表格", width=90, fg_color="#7F8C8D", hover_color="#616A6B", command=self._on_clear_table
        )
        self.btn_clear_table.pack(side="left", padx=4)

        # 比對模式下拉選單
        ctk.CTkLabel(bar_frame, text="目錄比對模式:", font=("Microsoft JhengHei UI", 11)).pack(side="left", padx=(12, 4))
        self.cb_match_mode = ctk.CTkOptionMenu(
            bar_frame,
            values=["模糊包含 (Contains)", "完全精準相符 (Exact)"],
            width=150
        )
        self.cb_match_mode.set("模糊包含 (Contains)")
        self.cb_match_mode.pack(side="left", padx=4)

        # 複製批號按鈕 (純逗號串接 + 有序去重，位於目錄比對模式右側)
        self.btn_copy_lots = ctk.CTkButton(
            bar_frame,
            text="📋 複製批號 (,)",
            width=115,
            fg_color="#6F42C1",
            hover_color="#59359A",
            command=self._on_copy_lots_to_clipboard
        )
        self.btn_copy_lots.pack(side="left", padx=(10, 4))

        self.table_rules = CTkRuleTable(rule_frame)
        self.table_rules.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)

        # ---------------- 3. 下載與檔案設定區 ----------------
        cfg_frame = ctk.CTkFrame(self)
        cfg_frame.grid(row=2, column=0, sticky="ew", padx=12, pady=5)
        cfg_frame.grid_columnconfigure(1, weight=2)
        cfg_frame.grid_columnconfigure(3, weight=1)
        cfg_frame.grid_columnconfigure(5, weight=1)

        ctk.CTkLabel(cfg_frame, text="3. 下載與檔案設定", font=("Microsoft JhengHei UI", 13, "bold")).grid(
            row=0, column=0, columnspan=6, sticky="w", padx=10, pady=(5, 3)
        )

        ctk.CTkLabel(cfg_frame, text="目標檔案過濾:").grid(row=1, column=0, sticky="e", padx=5, pady=4)
        self.ed_global_pat = ctk.CTkEntry(
            cfg_frame,
            placeholder_text="例: *RAW*; *.csv (在最後一層目錄搜尋，留空代表全選)"
        )
        self.ed_global_pat.grid(row=1, column=1, sticky="ew", padx=5, pady=4)

        ctk.CTkLabel(cfg_frame, text="重複檔案處理:").grid(row=1, column=2, sticky="e", padx=5, pady=4)
        self.cb_strategy = ctk.CTkOptionMenu(
            cfg_frame,
            values=["比對大小與修改時間略過 (推薦)", "直接覆蓋 (Overwrite)", "已存在則略過 (Skip)"]
        )
        self.cb_strategy.grid(row=1, column=3, sticky="ew", padx=5, pady=4)

        ctk.CTkLabel(cfg_frame, text="傳輸頻率:").grid(row=1, column=4, sticky="e", padx=5, pady=4)
        self.cb_delay = ctk.CTkOptionMenu(
            cfg_frame,
            values=["穩定推薦 (50ms)", "溫和防封 (150ms)", "極速 (0ms)"],
            width=130
        )
        self.cb_delay.set("穩定推薦 (50ms)")
        self.cb_delay.grid(row=1, column=5, sticky="ew", padx=5, pady=4)

        ctk.CTkLabel(cfg_frame, text="本地儲存路徑:").grid(row=2, column=0, sticky="e", padx=5, pady=4)
        self.ed_local_path = ctk.CTkEntry(cfg_frame)
        self.ed_local_path.grid(row=2, column=1, columnspan=4, sticky="ew", padx=5, pady=4)

        self.btn_browse_local = ctk.CTkButton(cfg_frame, text="瀏覽...", width=90, command=self._on_browse_local)
        self.btn_browse_local.grid(row=2, column=5, padx=5, pady=4)

        # ---------------- 4. 執行狀態與日誌反饋區 ----------------
        mon_frame = ctk.CTkFrame(self)
        mon_frame.grid(row=3, column=0, sticky="nsew", padx=12, pady=(5, 10))
        mon_frame.grid_rowconfigure(2, weight=1)
        mon_frame.grid_columnconfigure(0, weight=1)

        prog_bar_frame = ctk.CTkFrame(mon_frame, fg_color="transparent")
        prog_bar_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=(5, 2))
        prog_bar_frame.grid_columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(prog_bar_frame)
        self.progress_bar.set(0.0)
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=(5, 10))

        self.lbl_file_prog = ctk.CTkLabel(prog_bar_frame, text="目前檔案: 待命中", width=260, anchor="w")
        self.lbl_file_prog.grid(row=0, column=1, sticky="e", padx=5)

        self.lbl_stats = ctk.CTkLabel(mon_frame, text="已掃描目錄: 0 | 找到檔案: 0 | 下載成功: 0 | 略過: 0 | 失敗: 0", anchor="w")
        self.lbl_stats.grid(row=1, column=0, sticky="w", padx=10, pady=2)

        self.txt_log = ctk.CTkTextbox(mon_frame, font=("Consolas", 10))
        self.txt_log.grid(row=2, column=0, sticky="nsew", padx=5, pady=5)
        self._setup_log_tags()

        ctrl_frame = ctk.CTkFrame(mon_frame, fg_color="transparent")
        ctrl_frame.grid(row=3, column=0, sticky="ew", padx=5, pady=5)

        # 1. 開始下載 (核心主動作，最左側)
        self.btn_start = ctk.CTkButton(
            ctrl_frame, text="▶ 開始下載", font=("Microsoft JhengHei UI", 12, "bold"),
            fg_color="#28A745", hover_color="#218838", command=self._on_start_task
        )
        self.btn_start.pack(side="left", padx=5)

        # 2. 彙整資料 (本次新增，具備三態高亮防呆)
        self.btn_aggregate = ctk.CTkButton(
            ctrl_frame, text="📊 彙整資料", font=("Microsoft JhengHei UI", 12, "bold"),
            fg_color="#FD7E14", hover_color="#E8590C", command=self._on_aggregate_data_task
        )
        self.btn_aggregate.pack(side="left", padx=5)

        # 3. 開啟下載清單 (Excel)
        self.btn_open_manifest = ctk.CTkButton(
            ctrl_frame, text="📑 開啟下載清單 (Excel)", font=("Microsoft JhengHei UI", 11, "bold"),
            fg_color="#1F538D", hover_color="#14375E", width=180, command=self._on_open_manifest
        )
        self.btn_open_manifest.pack(side="left", padx=5)

        # 4. 開啟下載資料夾
        self.btn_open_folder = ctk.CTkButton(
            ctrl_frame, text="📁 開啟下載資料夾", command=self._on_open_local_folder
        )
        self.btn_open_folder.pack(side="left", padx=5)

        # 5. 僅產出清單 (純盤點模式)
        self.btn_scan_only = ctk.CTkButton(
            ctrl_frame, text="🔍 僅產出清單", font=("Microsoft JhengHei UI", 12, "bold"),
            fg_color="#17A2B8", hover_color="#138496", command=self._on_scan_only_task
        )
        self.btn_scan_only.pack(side="left", padx=5)

        # 6. 停止 (移至最右側防誤觸)
        self.btn_stop = ctk.CTkButton(
            ctrl_frame, text="⏹ 停止", fg_color="#DC3545", hover_color="#C82333", state="disabled", command=self._on_stop_task
        )
        self.btn_stop.pack(side="left", padx=5)

    def _setup_log_tags(self) -> None:
        """設定 Textbox 的彩色字元 Tag"""
        for level, color in self.LOG_COLORS.items():
            self.txt_log.tag_config(level, foreground=color)
        self.txt_log.tag_config("TIMESTAMP", foreground="#888888")

    # ---------------- 跨執行緒 80ms 輪詢器 ----------------

    def _process_ui_queue(self) -> None:
        """每 80ms 批次取出所有累積的 UIEvent，徹底解決多執行緒阻塞與日誌重繪卡死"""
        try:
            if not self.winfo_exists():
                return
            while not self._ui_queue.empty():
                event = self._ui_queue.get_nowait()
                if event.event_type == "LOG":
                    level, msg = event.payload
                    now_str = datetime.datetime.now().strftime("%H:%M:%S")
                    self.txt_log.insert("end", f"[{now_str}] ", "TIMESTAMP")
                    self.txt_log.insert("end", f"[{level}] {msg}\n", level)
                    self.txt_log.see("end")
                elif event.event_type == "STATS":
                    stats: TaskStats = event.payload
                    self.lbl_stats.configure(
                        text=f"已掃描目錄: {stats.scanned_dirs} | "
                             f"找到檔案: {stats.matched_files} | "
                             f"下載成功: {stats.downloaded_ok} | "
                             f"略過: {stats.skipped_files} | "
                             f"失敗: {stats.failed_files}"
                    )
                elif event.event_type == "OVERALL_PROG":
                    self.progress_bar.set(float(event.payload))
                elif event.event_type == "FILE_PROG":
                    fname, pct = event.payload
                    self.lbl_file_prog.configure(text=f"目前檔案: {fname} ({pct}%)")
                elif event.event_type == "FINISHED":
                    is_ok, summary_msg, records, manifest_path = event.payload
                    self._last_records = records
                    self._last_manifest_csv = manifest_path
                    self._set_running_state(False)
                    self._update_aggregate_btn_state()
                    self.lbl_file_prog.configure(text="目前檔案: 任務結束")

                    if manifest_path and os.path.exists(manifest_path):
                        csv_name = os.path.basename(manifest_path)
                        is_scan_report = "檢索清單" in csv_name
                        dialog_title = "檢索任務完成" if is_scan_report else "下載任務完成"
                        desc_title = "檢索清單 CSV" if is_scan_report else "下載清單 CSV"
                        ans = messagebox.askyesno(
                            dialog_title,
                            f"{summary_msg}\n\n已為您自動產出完整{desc_title}：\n{csv_name}\n\n是否立即以 Excel 開啟清單？"
                        )
                        if ans:
                            os.startfile(manifest_path)
                    else:
                        messagebox.showinfo("任務結束", summary_msg)
                elif event.event_type == "TEST_RESULT":
                    ok, msg = event.payload
                    self.btn_test_conn.configure(state="normal", text="測試連線")
                    if ok:
                        messagebox.showinfo("連線測試成功", msg)
                        self._enqueue_log("SUCCESS", msg)
                    else:
                        messagebox.showerror("連線測試失敗", msg)
                        self._enqueue_log("ERROR", msg)
        except Exception:
            pass
        finally:
            try:
                if self.winfo_exists():
                    self._after_id = self.after(80, self._process_ui_queue)
            except Exception:
                pass

    def destroy(self) -> None:
        """安全銷毀視窗並取消計時器"""
        if getattr(self, "_after_id", None):
            try:
                self.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None
        super().destroy()

    def _enqueue_log(self, level: str, msg: str) -> None:
        self._ui_queue.put(UIEvent("LOG", (level, msg)))

    # ---------------- 使用者互動與驗證 ----------------

    def _on_toggle_password(self) -> None:
        show_char = self.ed_pass.cget("show")
        self.ed_pass.configure(show="" if show_char == "*" else "*")

    def _on_browse_remote(self) -> None:
        """開啟遠端路徑瀏覽器對話框"""
        host = self.ed_host.get().strip()
        user = self.ed_user.get().strip()
        if not host or not user:
            messagebox.showwarning("連線資訊不足", "請先填寫 SFTP 主機與帳號，以便連線瀏覽目錄！")
            return

        sftp_cfg, _, _ = self._collect_current_config()
        current_init_path = self.ed_init_path.get().strip() or "/"

        def _on_path_chosen(chosen_path: str) -> None:
            self.ed_init_path.delete(0, "end")
            self.ed_init_path.insert(0, chosen_path)
            self._enqueue_log("INFO", f"已設定遠端初始路徑為: {chosen_path}")

        RemoteBrowserDialog(
            master=self,
            sftp_config=sftp_cfg,
            initial_path=current_init_path,
            on_path_selected=_on_path_chosen
        )

    def _validate_inputs(self, require_rules: bool = True) -> bool:
        host = self.ed_host.get().strip()
        user = self.ed_user.get().strip()
        port_str = self.ed_port.get().strip()

        if not host:
            messagebox.showwarning("輸入驗證錯誤", "主機名稱 (Host) 不得為空！")
            return False
        if not user:
            messagebox.showwarning("輸入驗證錯誤", "登入帳號 (Username) 不得為空！")
            return False
        try:
            port = int(port_str)
            if port < 1 or port > 65535:
                raise ValueError()
        except ValueError:
            messagebox.showwarning("輸入驗證錯誤", "Port 必須為 1 ~ 65535 的合法整數！")
            return False

        if not require_rules:
            return True

        rules = self.table_rules.get_all_rules()
        global_pat = self.ed_global_pat.get().strip()
        if not rules and not global_pat:
            messagebox.showwarning("規則驗證錯誤", "請至少輸入 1 列資料夾階層規則，或填寫目標檔案過濾條件！")
            return False

        local_dir = self.ed_local_path.get().strip()
        if not local_dir:
            messagebox.showwarning("路徑驗證錯誤", "本地儲存路徑不得為空！")
            return False

        if not os.path.exists(local_dir):
            ans = messagebox.askyesno("建立目錄確認", f"本地儲存路徑不存在：\n{local_dir}\n\n是否自動建立？")
            if not ans:
                return False
            try:
                os.makedirs(local_dir, exist_ok=True)
            except OSError as e:
                messagebox.showerror("錯誤", f"無法建立本地目錄：{e}")
                return False

        # 測試本地寫入權限
        test_file = os.path.join(local_dir, ".sftp_probe.tmp")
        try:
            with open(test_file, "w") as f:
                f.write("ok")
            os.remove(test_file)
        except OSError as e:
            messagebox.showerror("權限錯誤", f"本地路徑不具備寫入權限：\n{e}")
            return False

        return True

    def _collect_current_config(self) -> tuple[SFTPConfig, SearchSettings, List[RuleItem]]:
        port = int(self.ed_port.get().strip() or "22")
        strategy_text = self.cb_strategy.get()
        if "覆蓋" in strategy_text:
            strategy = ConflictStrategy.OVERWRITE
        elif "已存在" in strategy_text:
            strategy = ConflictStrategy.SKIP_EXISTING
        else:
            strategy = ConflictStrategy.SIZE_MTIME

        mode_text = self.cb_match_mode.get()
        if "精準" in mode_text:
            match_mode = MatchMode.EXACT
        else:
            match_mode = MatchMode.FUZZY

        delay_text = self.cb_delay.get()
        if "極速" in delay_text:
            delay_ms = 0
        elif "150ms" in delay_text:
            delay_ms = 150
        else:
            delay_ms = 50

        sftp_cfg = SFTPConfig(
            host=self.ed_host.get().strip(),
            port=port,
            username=self.ed_user.get().strip(),
            password=self.ed_pass.get(),
            save_password=bool(self.chk_save_pass.get()),
            initial_path=self.ed_init_path.get().strip()
        )
        search_cfg = SearchSettings(
            global_pattern=self.ed_global_pat.get().strip(),
            conflict_strategy=strategy,
            local_save_path=self.ed_local_path.get().strip(),
            match_mode=match_mode,
            request_delay_ms=delay_ms,
            session_refresh_interval=150
        )
        rules = self.table_rules.get_all_rules()
        return sftp_cfg, search_cfg, rules

    def _on_test_connection(self) -> None:
        if not self._validate_inputs(require_rules=False):
            return
        sftp_cfg, _, _ = self._collect_current_config()
        self.btn_test_conn.configure(state="disabled", text="測試中...")
        self._test_worker = ConnectionTestWorker(self._ui_queue, sftp_cfg)
        self._test_worker.start()

    def _on_scan_only_task(self) -> None:
        """僅搜尋遠端目錄並產出清單 (不下載二進位實體檔案)"""
        if not self._validate_inputs(require_rules=True):
            return

        sftp_cfg, search_cfg, rules = self._collect_current_config()
        ConfigManager.save_config(sftp_cfg, search_cfg, rules)

        self._last_records.clear()
        self._last_manifest_csv = None
        self.progress_bar.set(0.0)
        self.lbl_file_prog.configure(text="目前檔案: 僅搜尋目錄與檔案清單中...")
        self._set_running_state(True)

        self._dl_worker = DownloadWorker(self._ui_queue, sftp_cfg, search_cfg, rules, scan_only=True)
        self._dl_worker.start()

    def _on_start_task(self) -> None:
        if not self._validate_inputs(require_rules=True):
            return

        sftp_cfg, search_cfg, rules = self._collect_current_config()
        ConfigManager.save_config(sftp_cfg, search_cfg, rules)

        self._last_records.clear()
        self._last_manifest_csv = None
        self.progress_bar.set(0.0)
        self.lbl_file_prog.configure(text="目前檔案: 掃描遠端目錄中...")
        self._set_running_state(True)

        self._dl_worker = DownloadWorker(self._ui_queue, sftp_cfg, search_cfg, rules)
        self._dl_worker.start()

    def _on_stop_task(self) -> None:
        if self._dl_worker and self._dl_worker.is_alive():
            self.btn_stop.configure(state="disabled")
            self._enqueue_log("WARNING", "正在發送中斷訊號，等待安全停止...")
            self._dl_worker.request_cancel()

    def _set_running_state(self, running: bool) -> None:
        self.btn_scan_only.configure(state="disabled" if running else "normal")
        self.btn_start.configure(state="disabled" if running else "normal")
        self.btn_test_conn.configure(state="disabled" if running else "normal")
        self.btn_browse_remote.configure(state="disabled" if running else "normal")
        self.btn_stop.configure(state="normal" if running else "disabled")
        self.btn_paste_excel.configure(state="disabled" if running else "normal")
        self.btn_clear_table.configure(state="disabled" if running else "normal")
        self.btn_import_file.configure(state="disabled" if running else "normal")
        self.cb_match_mode.configure(state="disabled" if running else "normal")
        self.cb_delay.configure(state="disabled" if running else "normal")
        self.btn_copy_lots.configure(state="disabled" if running else "normal")
        if running:
            self.btn_aggregate.configure(state="disabled")
        else:
            self._update_aggregate_btn_state()

    @staticmethod
    def extract_distinct_lots(rules: List[RuleItem]) -> List[str]:
        """
        從規則清單中提取有效批號 (Lot No.)：
        - 若規則有 >= 2 層目錄，優先取第 2 層目錄 (dirs[1])；若只有 1 層目錄則取第 1 層 (dirs[0])
        - 智慧過濾空行與標題列 (例如 '批號', 'lot', 'lot no', 'lot no.', 'lot id', 'lot_id')
        - 有序去重 (Distinct)
        """
        excluded_headers = {
            "批號", "lot", "lot no", "lot no.", "lot id", "lot_id",
            "lot_no", "lotid", "lotnumber", "批次", "lot號", "lot編號"
        }
        raw_lots: List[str] = []
        for r in rules:
            if not r.dirs:
                continue
            # 若目錄有兩層或更多，批號通常位於第二層 (dirs[1])，否則取第一層 (dirs[0])
            target_lot = r.dirs[1].strip() if len(r.dirs) >= 2 else r.dirs[0].strip()
            if not target_lot:
                continue
            if target_lot.lower() in excluded_headers:
                continue
            raw_lots.append(target_lot)

        # 有序去重 (保留先出現的順序)
        return list(dict.fromkeys(raw_lots))

    def _on_copy_lots_to_clipboard(self) -> None:
        """一鍵提取表格中所有批號，以純逗號串接並有序去重後複製至剪貼簿"""
        rules = self.table_rules.get_rules()
        unique_lots = self.extract_distinct_lots(rules)

        if not unique_lots:
            messagebox.showinfo("提示", "表格內未找到有效的批號資料，請先貼上或匯入清單。")
            return

        joined_str = ",".join(unique_lots)
        try:
            self.clipboard_clear()
            self.clipboard_append(joined_str)
            self.update()  # 確保寫入 Windows 剪貼簿生效
            msg = f"已成功複製 {len(unique_lots)} 筆批號至剪貼簿（純逗號分隔，已自動去重）！"
            self._enqueue_log("SUCCESS", msg)
        except Exception as e:
            messagebox.showerror("複製失敗", f"寫入剪貼簿時發生錯誤：{e}")

    def _find_latest_download_manifest(self, local_dir: str) -> Optional[str]:
        """尋找本地目錄中最新的一份下載清單 CSV"""
        if not local_dir or not os.path.exists(local_dir):
            return None
        csvs = sorted(Path(local_dir).glob("下載清單_*.csv"), key=os.path.getmtime, reverse=True)
        return str(csvs[0]) if csvs else None

    def _update_aggregate_btn_state(self) -> None:
        """
        三態智慧按鈕偵測：
        1. 不可執行 (Disabled, 灰色): 無下載清單且無當前完成任務
        2. 可執行 (Highlight, 亮橘): 剛下載完成或本地有清單，但尚未產出對應的 .xlsx
        3. 已有檔案 (Completed, 綠色/完成態): 本地已有對應的「下載與明細彙整_*.xlsx」
        """
        if not hasattr(self, "btn_aggregate"):
            return

        local_dir = self.ed_local_path.get().strip()
        latest_csv = self._last_manifest_csv if (
            self._last_manifest_csv and "下載清單" in os.path.basename(self._last_manifest_csv) and os.path.exists(self._last_manifest_csv)
        ) else self._find_latest_download_manifest(local_dir)

        if not latest_csv:
            # 狀態 1: 不可執行 (灰色 Disabled)
            self.btn_aggregate.configure(
                state="disabled",
                text="📊 彙整資料",
                fg_color="#6C757D"
            )
            return

        # 檢查是否已存在對應的「下載與明細彙整_*.xlsx」
        has_xlsx = False
        if local_dir and os.path.exists(local_dir):
            xlsxs = list(Path(local_dir).glob("下載與明細彙整_*.xlsx"))
            if xlsxs:
                has_xlsx = True

        if has_xlsx:
            # 狀態 3: 已有檔案 (正常綠色完成態)
            self.btn_aggregate.configure(
                state="normal",
                text="📊 彙整資料 (已產出)",
                fg_color="#20C997",
                hover_color="#17A2B8"
            )
        else:
            # 狀態 2: 可執行 (醒目亮橘高亮)
            self.btn_aggregate.configure(
                state="normal",
                text="📊 彙整資料",
                fg_color="#FD7E14",
                hover_color="#E8590C"
            )

    def _on_aggregate_data_task(self) -> None:
        """點擊「📊 彙整資料」時，以非同步執行緒解析晶片檔案並產出雙工作表 .xlsx"""
        import threading
        local_dir = self.ed_local_path.get().strip()
        manifest_csv = self._last_manifest_csv if (
            self._last_manifest_csv and "下載清單" in os.path.basename(self._last_manifest_csv) and os.path.exists(self._last_manifest_csv)
        ) else self._find_latest_download_manifest(local_dir)

        if not manifest_csv or not os.path.exists(manifest_csv):
            messagebox.showwarning("提示", "未找到有效的下載清單 CSV，請先執行下載任務。")
            return

        timestamp_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_xlsx = os.path.join(local_dir, f"下載與明細彙整_{timestamp_str}.xlsx")

        self.btn_aggregate.configure(state="disabled", text="彙整中...")
        self._enqueue_log("INFO", f"開始進行晶片資料穿透解析與彙整，來源清單: {os.path.basename(manifest_csv)} ...")

        def run_worker():
            try:
                succeeded, failed, total_rows = DetailAggregator.aggregate_to_excel(
                    manifest_csv_path=manifest_csv,
                    output_xlsx_path=output_xlsx,
                    log_cb=lambda msg: self._enqueue_log("WARNING", msg)
                )
                self._enqueue_log("SUCCESS", f"晶片明細彙整完成！共解析 {succeeded} 個檔案 (失敗/略過 {failed} 個)，寫入 {total_rows} 筆明細。")
                self._enqueue_log("SUCCESS", f"匯出檔案: {output_xlsx}")

                def on_done():
                    self._update_aggregate_btn_state()
                    ans = messagebox.askyesno(
                        "彙整完成",
                        f"晶片資料彙整完成！\n共解析 {succeeded} 個檔案，寫入 {total_rows} 筆晶片明細。\n\n已成功產出：\n{os.path.basename(output_xlsx)}\n\n是否立即以 Excel 開啟活頁簿？"
                    )
                    if ans:
                        os.startfile(output_xlsx)

                self.after(0, on_done)

            except Exception as ex:
                self._enqueue_log("ERROR", f"彙整失敗: {ex}")
                self.after(0, lambda: [
                    self._update_aggregate_btn_state(),
                    messagebox.showerror("彙整失敗", f"產出明細活頁簿時發生錯誤：\n{ex}")
                ])

        threading.Thread(target=run_worker, daemon=True).start()

    def _on_clear_table(self) -> None:
        if messagebox.askyesno("確認清空", "確定要清空表格內的所有搜尋規則嗎？"):
            self.table_rules.load_rules([])

    def _on_import_rules_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="匯入規則檔案",
            filetypes=[("支援格式", "*.csv *.txt *.xlsx"), ("所有檔案", "*.*")]
        )
        if not file_path:
            return
        try:
            ext = Path(file_path).suffix.lower()
            if ext == ".xlsx":
                import openpyxl
                wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
                ws = wb.active
                lines = []
                for row in ws.iter_rows(values_only=True):
                    cells = [str(c) if c is not None else "" for c in row]
                    lines.append("\t".join(cells))
                raw_text = "\n".join(lines)
            else:
                with open(file_path, "r", encoding="utf-8-sig") as f:
                    raw_text = f.read().replace(",", "\t") if ext == ".csv" else f.read()

            rules, dir_count = CTkRuleTable.parse_tsv_text(raw_text, dynamic=True)
            if rules:
                self.table_rules.load_rules(rules)
                self._enqueue_log("INFO", f"已成功從檔案載入 {len(rules)} 列規則（適配 {dir_count} 層目錄）：{file_path}")
        except Exception as e:
            messagebox.showerror("匯入失敗", f"讀取規則檔案失敗：{e}")

    def _on_browse_local(self) -> None:
        folder = filedialog.askdirectory(title="選擇本地儲存目錄", initialdir=self.ed_local_path.get())
        if folder:
            self.ed_local_path.delete(0, "end")
            self.ed_local_path.insert(0, os.path.normpath(folder))
            self._update_aggregate_btn_state()

    def _on_open_manifest(self) -> None:
        """一鍵開啟當次任務產出的清單 CSV (支援下載清單與檢索清單)"""
        if self._last_manifest_csv and os.path.exists(self._last_manifest_csv):
            os.startfile(self._last_manifest_csv)
        else:
            # 嘗試到本地儲存路徑尋找最新的一份清單 (下載清單或檢索清單)
            local_dir = self.ed_local_path.get().strip()
            if local_dir and os.path.exists(local_dir):
                all_csvs = list(Path(local_dir).glob("下載清單_*.csv")) + list(Path(local_dir).glob("檢索清單_*.csv"))
                csvs = sorted(all_csvs, key=os.path.getmtime, reverse=True)
                if csvs:
                    os.startfile(str(csvs[0]))
                    return
            messagebox.showinfo("提示", "目前尚未找到產出的清單 CSV，請先執行搜尋或下載任務。")

    def _on_open_local_folder(self) -> None:
        local_dir = self.ed_local_path.get().strip()
        if local_dir and os.path.exists(local_dir):
            os.startfile(local_dir)
        else:
            messagebox.showwarning("提示", "本地目錄尚不存在，請先指定有效路徑。")

    def _load_saved_state(self) -> None:
        sftp_cfg, search_cfg, rules = ConfigManager.load_config()
        self.ed_host.insert(0, sftp_cfg.host)
        self.ed_port.delete(0, "end")
        self.ed_port.insert(0, str(sftp_cfg.port))
        self.ed_user.insert(0, sftp_cfg.username)
        if sftp_cfg.save_password:
            self.chk_save_pass.select()
            self.ed_pass.insert(0, sftp_cfg.password)
        self.ed_init_path.delete(0, "end")
        self.ed_init_path.insert(0, sftp_cfg.initial_path or "/")

        self.ed_global_pat.insert(0, search_cfg.global_pattern)
        if search_cfg.conflict_strategy == ConflictStrategy.OVERWRITE:
            self.cb_strategy.set("直接覆蓋 (Overwrite)")
        elif search_cfg.conflict_strategy == ConflictStrategy.SKIP_EXISTING:
            self.cb_strategy.set("已存在則略過 (Skip)")
        else:
            self.cb_strategy.set("比對大小與修改時間略過 (推薦)")

        if search_cfg.match_mode == MatchMode.EXACT:
            self.cb_match_mode.set("完全精準相符 (Exact)")
        else:
            self.cb_match_mode.set("模糊包含 (Contains)")

        if search_cfg.request_delay_ms == 0:
            self.cb_delay.set("極速 (0ms)")
        elif search_cfg.request_delay_ms >= 150:
            self.cb_delay.set("溫和防封 (150ms)")
        else:
            self.cb_delay.set("穩定推薦 (50ms)")

        self.ed_local_path.insert(0, search_cfg.local_save_path)
        self.table_rules.load_rules(rules)
        self._update_aggregate_btn_state()

    def _on_close_window(self) -> None:
        if self._dl_worker and self._dl_worker.is_alive():
            if not messagebox.askyesno("任務執行中", "目前仍有下載任務正在進行，確定要中止並退出嗎？"):
                return
            self._dl_worker.request_cancel()

        sftp_cfg, search_cfg, rules = self._collect_current_config()
        ConfigManager.save_config(sftp_cfg, search_cfg, rules)
        self.destroy()
