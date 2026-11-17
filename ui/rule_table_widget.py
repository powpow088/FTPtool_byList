import csv
import io
from typing import List, Tuple, Optional
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QKeyEvent, QKeySequence, QGuiApplication
from PyQt6.QtWidgets import (
    QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox, QAbstractItemView
)
from models.data_models import RuleItem


class RuleTableWidget(QTableWidget):
    """
    自訂三層規則表格：
    - 攔截 Ctrl+V 鍵盤事件與 Delete 刪除列
    - 使用 csv.reader(dialect='excel-tab') 解析 Excel 剪貼簿 TSV
    - 自動適配 1 欄 / 2 欄 / 3 欄 / 4+ 欄，動態擴增列數與批次高效填入
    """

    HEADERS = [
        "第 1 層目錄 (模糊包含，留空=全選)",
        "第 2 層目錄 (模糊包含，留空=全選)",
        "檔案名稱 (支援 *RAW*-*.csv，留空=套用全域)"
    ]

    def __init__(self, parent=None) -> None:
        super().__init__(0, 3, parent)
        self.setHorizontalHeaderLabels(self.HEADERS)
        self.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.setAlternatingRowColors(True)

    def keyPressEvent(self, event: QKeyEvent) -> None:
        """攔截 Ctrl+V / Cmd+V 快捷鍵與 Delete 鍵"""
        if event.matches(QKeySequence.StandardKey.Paste):
            self.paste_from_clipboard()
            event.accept()
            return
        if event.key() == Qt.Key.Key_Delete:
            self.remove_selected_rows()
            event.accept()
            return
        super().keyPressEvent(event)

    def paste_from_clipboard(self) -> None:
        """從系統剪貼簿讀取 Excel TSV 文字並執行智慧欄位適配填入"""
        clipboard = QGuiApplication.clipboard()
        raw_text = clipboard.text() if clipboard else ""
        if not raw_text or not raw_text.strip():
            return

        rows_data, has_truncated_cols = self.parse_tsv_text(raw_text)
        if not rows_data:
            return

        # 決定貼上起始列：
        # 1. 若使用者有選取列，從最小選取列開始覆寫
        # 2. 若無選取，檢查最後一列是否全空，是則覆寫最後一列，否則附加於末尾
        selected_indexes = self.selectedIndexes()
        if selected_indexes:
            start_row = min(idx.row() for idx in selected_indexes)
        else:
            last_row = self.rowCount() - 1
            if last_row >= 0 and self._is_row_empty(last_row):
                start_row = last_row
            else:
                start_row = self.rowCount()

        required_rows = start_row + len(rows_data)

        # 鎖定 UI 更新與訊號，即使一次貼上數千列也能在 10ms 內完成且不卡死
        self.setUpdatesEnabled(False)
        self.blockSignals(True)
        try:
            if required_rows > self.rowCount():
                self.setRowCount(required_rows)

            for offset, rule in enumerate(rows_data):
                target_row = start_row + offset
                self._set_row_rule(target_row, rule)
        finally:
            self.blockSignals(False)
            self.setUpdatesEnabled(True)

        if has_truncated_cols:
            QMessageBox.information(
                self,
                "欄位自動適配提示",
                "偵測到剪貼簿資料超過 3 欄，系統已自動擷取前 3 欄並忽略多餘欄位。"
            )

    @staticmethod
    def parse_tsv_text(raw_text: str) -> Tuple[List[RuleItem], bool]:
        """
        深水區 2 核心演算法：
        解析 TSV 字串，清理前後空白與多餘引號，並自動適配 1/2/3/4+ 欄。
        回傳 (規則清單, 是否曾截斷第 4 欄以後資料)
        """
        if not raw_text or not raw_text.strip():
            return [], False

        reader = csv.reader(io.StringIO(raw_text), dialect="excel-tab")
        parsed_rules: List[RuleItem] = []
        has_truncated = False

        for raw_row in reader:
            cleaned = [cell.strip().strip('"').strip("'") for cell in raw_row]
            # 忽略全空列 (例如 Excel 複製區塊尾端常見的空換行)
            if not any(cleaned):
                continue

            col_count = len(cleaned)
            if col_count == 1:
                rule = RuleItem(l1_keyword=cleaned[0], l2_keyword="", file_pattern="")
            elif col_count == 2:
                rule = RuleItem(l1_keyword=cleaned[0], l2_keyword=cleaned[1], file_pattern="")
            elif col_count == 3:
                rule = RuleItem(l1_keyword=cleaned[0], l2_keyword=cleaned[1], file_pattern=cleaned[2])
            else:
                has_truncated = True
                rule = RuleItem(l1_keyword=cleaned[0], l2_keyword=cleaned[1], file_pattern=cleaned[2])

            parsed_rules.append(rule)

        return parsed_rules, has_truncated

    def append_empty_row(self) -> None:
        """在表格底部插入空白列"""
        row_idx = self.rowCount()
        self.insertRow(row_idx)
        self._set_row_rule(row_idx, RuleItem())
        self.scrollToBottom()

    def remove_selected_rows(self) -> None:
        """移除使用者反白選取的所有列"""
        selected_rows = sorted({idx.row() for idx in self.selectedIndexes()}, reverse=True)
        for r in selected_rows:
            self.removeRow(r)

    def get_all_rules(self) -> List[RuleItem]:
        """匯出表格內所有非空規則列"""
        rules: List[RuleItem] = []
        for r in range(self.rowCount()):
            l1 = self._cell_text(r, 0)
            l2 = self._cell_text(r, 1)
            fp = self._cell_text(r, 2)
            item = RuleItem(l1_keyword=l1, l2_keyword=l2, file_pattern=fp)
            if not item.is_completely_empty():
                rules.append(item)
        return rules

    def load_rules(self, rules: List[RuleItem]) -> None:
        """批次載入規則清單至表格"""
        self.setUpdatesEnabled(False)
        try:
            self.setRowCount(0)
            if not rules:
                self.append_empty_row()
                return
            self.setRowCount(len(rules))
            for i, rule in enumerate(rules):
                self._set_row_rule(i, rule)
        finally:
            self.setUpdatesEnabled(True)

    def _set_row_rule(self, row: int, rule: RuleItem) -> None:
        self.setItem(row, 0, QTableWidgetItem(rule.l1_keyword))
        self.setItem(row, 1, QTableWidgetItem(rule.l2_keyword))
        self.setItem(row, 2, QTableWidgetItem(rule.file_pattern))

    def _cell_text(self, row: int, col: int) -> str:
        item = self.item(row, col)
        return item.text().strip() if item else ""

    def _is_row_empty(self, row: int) -> bool:
        return not (self._cell_text(row, 0) or self._cell_text(row, 1) or self._cell_text(row, 2))
