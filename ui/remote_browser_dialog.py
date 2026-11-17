import datetime
import posixpath
import stat
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Callable, List, Optional
import customtkinter as ctk

from models.data_models import SFTPConfig
from core.sftp_manager import SFTPManager


class RemoteBrowserDialog(ctk.CTkToplevel):
    """
    遠端路徑瀏覽器彈出對話框：
    - 連線至遠端 SFTP 伺服器
    - 提供層層點選資料夾導航、回上一層、重新整理
    - 點選「選擇此目錄」自動回填主視窗之遠端初始路徑
    - 100% 非同步載入，介面不卡頓
    """

    def __init__(
        self,
        master,
        sftp_config: SFTPConfig,
        initial_path: str,
        on_path_selected: Callable[[str], None]
    ) -> None:
        super().__init__(master)
        self.title("遠端路徑瀏覽器")
        self.geometry("680x480")
        self.minsize(550, 380)

        self._sftp_config = sftp_config
        self._current_path = initial_path.strip() if initial_path.strip() else "/"
        if not self._current_path.startswith("/"):
            self._current_path = "/" + self._current_path
        self._on_path_selected = on_path_selected

        self._sftp_mgr: Optional[SFTPManager] = None
        self._is_loading = False

        self._init_ui()
        self._center_window()

        self.transient(master)
        self.grab_set()

        # 視窗開啟後非同步初始化連線並載入
        self.after(100, self._start_load_directory)

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _center_window(self) -> None:
        """使對話框相對於父視窗置中"""
        self.update_idletasks()
        try:
            pw = self.master.winfo_width()
            ph = self.master.winfo_height()
            px = self.master.winfo_rootx()
            py = self.master.winfo_rooty()
            w = self.winfo_width()
            h = self.winfo_height()
            x = px + max(0, (pw - w) // 2)
            y = py + max(0, (ph - h) // 2)
            self.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            pass

    def _init_ui(self) -> None:
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # 頂部路徑與導航列
        top_bar = ctk.CTkFrame(self, fg_color="transparent")
        top_bar.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))
        top_bar.grid_columnconfigure(1, weight=1)

        btn_up = ctk.CTkButton(
            top_bar, text="⬆ 回上一層", width=90, height=30, command=self._navigate_up
        )
        btn_up.grid(row=0, column=0, padx=(0, 6), sticky="w")

        self.txt_path = ctk.CTkEntry(top_bar, height=30, font=("Consolas", 11))
        self.txt_path.insert(0, self._current_path)
        self.txt_path.grid(row=0, column=1, sticky="ew", padx=(0, 6))
        self.txt_path.bind("<Return>", lambda e: self._on_enter_path())

        btn_go = ctk.CTkButton(
            top_bar, text="前往", width=60, height=30, command=self._on_enter_path
        )
        btn_go.grid(row=0, column=2, padx=(0, 6), sticky="e")

        btn_refresh = ctk.CTkButton(
            top_bar, text="🔄 重新整理", width=80, height=30, command=self._start_load_directory
        )
        btn_refresh.grid(row=0, column=3, sticky="e")

        # 中部 Treeview 列表
        center_frame = ctk.CTkFrame(self)
        center_frame.grid(row=1, column=0, sticky="nsew", padx=12, pady=6)
        center_frame.grid_rowconfigure(0, weight=1)
        center_frame.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Browser.Treeview",
            background="#2B2B2B",
            foreground="#DCE4EE",
            fieldbackground="#2B2B2B",
            rowheight=26,
            borderwidth=0,
            font=("Microsoft JhengHei UI", 10)
        )
        style.configure(
            "Browser.Treeview.Heading",
            background="#1F538D",
            foreground="#FFFFFF",
            relief="flat",
            font=("Microsoft JhengHei UI", 10, "bold")
        )

        columns = ("col_name", "col_mtime")
        self.tree = ttk.Treeview(
            center_frame,
            columns=columns,
            show="headings",
            selectmode="browse",
            style="Browser.Treeview"
        )
        self.tree.heading("col_name", text="📁 資料夾名稱", anchor="w")
        self.tree.heading("col_mtime", text="修改時間", anchor="w")
        self.tree.column("col_name", width=420, anchor="w", stretch=True)
        self.tree.column("col_mtime", width=180, anchor="w", stretch=False)

        vsb = ctk.CTkScrollbar(center_frame, orientation="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew", padx=(2, 0), pady=2)
        vsb.grid(row=0, column=1, sticky="ns", padx=2, pady=2)

        self.tree.bind("<Double-1>", self._on_double_click_item)

        # 底部狀態列與按鈕
        bottom_bar = ctk.CTkFrame(self, fg_color="transparent")
        bottom_bar.grid(row=2, column=0, sticky="ew", padx=12, pady=(6, 12))
        bottom_bar.grid_columnconfigure(0, weight=1)

        self.lbl_status = ctk.CTkLabel(
            bottom_bar, text="準備連線...", anchor="w", text_color="#A0A0A0"
        )
        self.lbl_status.grid(row=0, column=0, sticky="w")

        btn_select = ctk.CTkButton(
            bottom_bar,
            text="✔ 選擇此目錄",
            width=110,
            height=32,
            fg_color="#28A745",
            hover_color="#218838",
            command=self._on_select_current
        )
        btn_select.grid(row=0, column=1, padx=(0, 8), sticky="e")

        btn_cancel = ctk.CTkButton(
            bottom_bar,
            text="取消",
            width=80,
            height=32,
            fg_color="#6C757D",
            hover_color="#5A6268",
            command=self._on_close
        )
        btn_cancel.grid(row=0, column=2, sticky="e")

    def _navigate_up(self) -> None:
        """回上一層目錄"""
        if self._is_loading:
            return
        if self._current_path in ("/", ""):
            return
        parent = posixpath.dirname(self._current_path.rstrip("/"))
        self._current_path = parent if parent else "/"
        self._update_path_entry()
        self._start_load_directory()

    def _on_enter_path(self) -> None:
        """使用者手動輸入路徑並按前往"""
        if self._is_loading:
            return
        typed = self.txt_path.get().strip()
        if not typed:
            typed = "/"
        if not typed.startswith("/"):
            typed = "/" + typed
        self._current_path = typed
        self._update_path_entry()
        self._start_load_directory()

    def _update_path_entry(self) -> None:
        self.txt_path.delete(0, "end")
        self.txt_path.insert(0, self._current_path)

    def _on_double_click_item(self, event) -> None:
        """雙擊資料夾進入下一層"""
        if self._is_loading:
            return
        selected = self.tree.selection()
        if not selected:
            return
        folder_name = self.tree.item(selected[0], "values")[0]
        # 去除圖示前綴（若有的話）
        clean_name = folder_name.replace("📁 ", "").strip()
        new_path = posixpath.join(self._current_path, clean_name)
        self._current_path = new_path
        self._update_path_entry()
        self._start_load_directory()

    def _start_load_directory(self) -> None:
        """發起背景執行緒載入遠端目錄列表"""
        if self._is_loading:
            return
        self._is_loading = True
        self.lbl_status.configure(text=f"正在載入遠端目錄: {self._current_path} ...")
        self.tree.delete(*self.tree.get_children())

        threading.Thread(target=self._worker_load, daemon=True).start()

    def _worker_load(self) -> None:
        try:
            if not self._sftp_mgr:
                self._sftp_mgr = SFTPManager(self._sftp_config)
                self._sftp_mgr.connect()

            entries = self._sftp_mgr.listdir_attr(self._current_path)
            folders = []
            for e in entries:
                if stat.S_ISDIR(e.st_mode or 0):
                    mtime_str = ""
                    if e.st_mtime:
                        mtime_str = datetime.datetime.fromtimestamp(e.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                    folders.append((e.filename, mtime_str))

            # 依名稱字母順序排序
            folders.sort(key=lambda x: x[0].lower())

            self.after(0, self._on_load_success, folders)
        except Exception as e:
            self.after(0, self._on_load_error, str(e))

    def _on_load_success(self, folders: List[tuple]) -> None:
        self._is_loading = False
        self.tree.delete(*self.tree.get_children())
        for fname, mtime in folders:
            self.tree.insert("", "end", values=(f"📁 {fname}", mtime))
        self.lbl_status.configure(text=f"載入完成，共 {len(folders)} 個資料夾")

    def _on_load_error(self, err_msg: str) -> None:
        self._is_loading = False
        self.lbl_status.configure(text=f"載入失敗: {err_msg}", text_color="#DC3545")
        messagebox.showerror("遠端瀏覽失敗", f"無法載入目錄 [{self._current_path}]：\n{err_msg}", parent=self)

    def _on_select_current(self) -> None:
        """確認選擇當前目錄"""
        # 若使用者在清單中選中某個子資料夾，以該子資料夾為主；否則以當前路徑為主
        selected = self.tree.selection()
        if selected:
            folder_name = self.tree.item(selected[0], "values")[0]
            clean_name = folder_name.replace("📁 ", "").strip()
            chosen_path = posixpath.join(self._current_path, clean_name)
        else:
            chosen_path = self._current_path

        if self._on_path_selected:
            self._on_path_selected(chosen_path)
        self._on_close()

    def _on_close(self) -> None:
        """關閉視窗並斷開 SFTP 連線"""
        if self._sftp_mgr:
            try:
                self._sftp_mgr.close()
            except Exception:
                pass
            self._sftp_mgr = None
        self.grab_release()
        self.destroy()
