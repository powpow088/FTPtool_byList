import csv
import io
import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Tuple, Optional, Union
import customtkinter as ctk

from models.data_models import RuleItem


class CTkRuleTable(ctk.CTkFrame):
    """
    專為 CustomTkinter 打造的高效能動態純目錄規則表格：
    - 底層採用 ttk.Treeview 確保數千列貼上不卡頓
    - 支援 Ctrl+V 攔截解析 Excel TSV：貼入幾欄即動態展開為幾層純資料夾
    - 支援滑鼠雙擊儲存格浮現 CTkEntry 原地編輯
    """

    def __init__(self, master, dir_count: int = 2, has_file_col: bool = False, **kwargs) -> None:
        super().__init__(master, **kwargs)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.dir_count = max(1, dir_count)
        self.has_file_col = has_file_col
        self._setup_style()

        self.tree = ttk.Treeview(
            self, columns=(), show="headings", selectmode="extended"
        )
        self.set_columns(self.dir_count, self.has_file_col)

        vsb = ctk.CTkScrollbar(self, orientation="vertical", command=self.tree.yview)
        hsb = ctk.CTkScrollbar(self, orientation="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew", padx=(2, 0), pady=(2, 0))
        vsb.grid(row=0, column=1, sticky="ns", padx=(0, 2), pady=(2, 0))
        hsb.grid(row=1, column=0, sticky="ew", padx=(2, 0), pady=(0, 2))

        # 鍵盤與滑鼠事件綁定
        self.tree.bind("<Control-v>", self._on_ctrl_v)
        self.tree.bind("<Control-V>", self._on_ctrl_v)
        self.tree.bind("<Delete>", lambda e: self.remove_selected_rows())
        self.tree.bind("<Double-1>", self._on_double_click_cell)

        self._active_editor: Optional[ctk.CTkEntry] = None

    def _setup_style(self) -> None:
        """配置符合 CustomTkinter Dark Mode 的 Treeview 樣式"""
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Treeview",
            background="#2B2B2B",
            foreground="#DCE4EE",
            fieldbackground="#2B2B2B",
            rowheight=28,
            borderwidth=0,
            font=("Microsoft JhengHei UI", 10)
        )
        style.configure(
            "Treeview.Heading",
            background="#1F538D",
            foreground="#FFFFFF",
            relief="flat",
            font=("Microsoft JhengHei UI", 10, "bold")
        )
        style.map("Treeview", background=[("selected", "#14375E")])

    def set_columns(self, dir_count: int = 2, has_file_col: bool = False) -> None:
        """根據目錄層數動態配置 Treeview 欄位與簡潔表頭"""
        dir_count = max(1, dir_count)
        self.dir_count = dir_count
        self.has_file_col = has_file_col

        cols = [f"col_dir_{i+1}" for i in range(dir_count)]
        headers = [f"第 {i+1} 層目錄" for i in range(dir_count)]

        if self.has_file_col:
            cols.append("col_file")
            headers.append("檔案名稱")

        self.tree["columns"] = cols
        for col_id, text in zip(cols, headers):
            self.tree.heading(col_id, text=text, anchor="w")
            self.tree.column(col_id, width=240, minwidth=140, anchor="w", stretch=True)

    def _on_ctrl_v(self, event=None) -> str:
        self.paste_from_clipboard()
        return "break"

    def paste_from_clipboard(self) -> None:
        """讀取系統剪貼簿並依 Excel 欄位數動態展開適配為純目錄"""
        try:
            raw_text = self.winfo_toplevel().clipboard_get()
        except tk.TclError:
            return
        if not raw_text or not raw_text.strip():
            return

        rules, detected_dir_count = self.parse_tsv_text(raw_text, dynamic=True)
        if not rules:
            return

        # 貼上時一律適配為純目錄
        self.set_columns(detected_dir_count, has_file_col=False)

        # 清除現存資料，載入新剪貼簿內容
        self.tree.delete(*self.tree.get_children())

        for r in rules:
            row_dirs = list(r.dirs) + [""] * (self.dir_count - len(r.dirs))
            self.tree.insert("", "end", values=tuple(row_dirs))

    @staticmethod
    def parse_tsv_text(raw_text: str, dynamic: bool = False):
        """
        核心剪貼簿解析演算法：
        - dynamic=False (相容微階段 4 舊測試)：強制適配為 3 欄 (l1, l2, file_pattern)
        - dynamic=True (微階段 8 純目錄架構)：貼入幾欄即 100% 解析為幾層純資料夾
        """
        if not raw_text or not raw_text.strip():
            return ([], False) if not dynamic else ([], 2)

        reader = csv.reader(io.StringIO(raw_text), dialect="excel-tab")
        raw_rows = []
        max_cols = 0
        for raw_row in reader:
            cleaned = [c.strip().strip('"').strip("'") for c in raw_row]
            if not any(cleaned):
                continue
            raw_rows.append(cleaned)
            if len(cleaned) > max_cols:
                max_cols = len(cleaned)

        if not raw_rows:
            return ([], False) if not dynamic else ([], 2)

        if not dynamic:
            parsed = []
            truncated = False
            for cleaned in raw_rows:
                n = len(cleaned)
                if n == 1:
                    parsed.append(RuleItem(cleaned[0], "", ""))
                elif n == 2:
                    parsed.append(RuleItem(cleaned[0], cleaned[1], ""))
                elif n == 3:
                    parsed.append(RuleItem(cleaned[0], cleaned[1], cleaned[2]))
                else:
                    truncated = True
                    parsed.append(RuleItem(cleaned[0], cleaned[1], cleaned[2]))
            return parsed, truncated
        else:
            dir_count = max(1, max_cols)
            parsed = []
            for row in raw_rows:
                padded = row + [""] * (dir_count - len(row))
                dirs = padded[:dir_count]
                parsed.append(RuleItem(dirs=dirs, file_pattern=""))
            return parsed, dir_count

    def _on_double_click_cell(self, event) -> None:
        """雙擊儲存格浮現動態 CTkEntry 原地編輯"""
        if self._active_editor:
            self._active_editor.destroy()

        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)  # '#1', '#2', ...
        if not row_id or not col_id:
            return

        bbox = self.tree.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, width, height = bbox
        col_idx = int(col_id.replace("#", "")) - 1

        current_vals = list(self.tree.item(row_id, "values"))
        current_text = current_vals[col_idx] if col_idx < len(current_vals) else ""

        editor = ctk.CTkEntry(self.tree, width=width, height=height, corner_radius=0)
        editor.insert(0, current_text)
        editor.select_range(0, "end")
        editor.place(x=x, y=y, width=width, height=height)
        editor.focus_set()
        self._active_editor = editor

        def _commit_edit(e=None) -> None:
            if not self._active_editor:
                return
            new_val = editor.get().strip()
            while len(current_vals) <= col_idx:
                current_vals.append("")
            current_vals[col_idx] = new_val
            if self.tree.exists(row_id):
                self.tree.item(row_id, values=current_vals)
            editor.destroy()
            self._active_editor = None

        editor.bind("<Return>", _commit_edit)
        editor.bind("<FocusOut>", _commit_edit)
        editor.bind("<Escape>", lambda e: editor.destroy())

    def append_empty_row(self) -> None:
        """新增空白列"""
        total_cols = self.dir_count + (1 if self.has_file_col else 0)
        empty_vals = [""] * total_cols
        item_id = self.tree.insert("", "end", values=tuple(empty_vals))
        self.tree.see(item_id)

    def remove_selected_rows(self) -> None:
        """刪除選取列（支援 Delete 鍵）"""
        for item_id in self.tree.selection():
            self.tree.delete(item_id)

    def load_rules(self, rules: List[RuleItem]) -> None:
        """載入規則清單至表格並自動調整欄數"""
        self.tree.delete(*self.tree.get_children())
        if not rules:
            self.set_columns(2, has_file_col=False)
            self.append_empty_row()
            return

        max_dirs = max((len(r.dirs) for r in rules), default=2)
        max_dirs = max(1, max_dirs)
        has_file = any(bool(r.file_pattern.strip()) for r in rules)

        self.set_columns(max_dirs, has_file_col=has_file)

        for r in rules:
            row_dirs = list(r.dirs) + [""] * (self.dir_count - len(r.dirs))
            if has_file:
                self.tree.insert("", "end", values=(*row_dirs, r.file_pattern))
            else:
                self.tree.insert("", "end", values=tuple(row_dirs))

    def get_all_rules(self) -> List[RuleItem]:
        """取得所有非空規則列"""
        rules: List[RuleItem] = []
        for item_id in self.tree.get_children():
            vals = list(self.tree.item(item_id, "values"))
            if not vals:
                continue
            if self.has_file_col:
                dirs = [str(v).strip() for v in vals[:-1]]
                file_pat = str(vals[-1]).strip() if len(vals) > len(dirs) else ""
            else:
                dirs = [str(v).strip() for v in vals]
                file_pat = ""
            item = RuleItem(dirs=dirs, file_pattern=file_pat)
            if not item.is_completely_empty():
                rules.append(item)
        return rules
