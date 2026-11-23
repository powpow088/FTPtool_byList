# core/excel_utils.py
import os
import re
import shutil
import time
from typing import Any

# XML 1.0 規範禁止的所有控制字元 (0x00-0x08, 0x0B-0x0C, 0x0E-0x1F, 0x7F-0x84, 0x86-0x9F)
# 保留合法空白字元：\x09 (Tab), \x0A (LF), \x0D (CR)
_ILLEGAL_XML_CHARS_RE = re.compile(
    r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x84\x86-\x9f]"
)


def sanitize_xml_text(val: Any) -> str:
    """
    剝除 XML 1.0 禁止的控制字元，防止 Excel 開啟大型報表時出現「部分內容有問題」損毀警告。
    採用快速路徑檢查：絕大多數正常晶圓字串不會觸發正則替換，效能損耗趨近於零。
    """
    if val is None:
        return ""
    s = str(val)
    if any(ord(ch) < 32 and ch not in ("\t", "\n", "\r") for ch in s):
        return _ILLEGAL_XML_CHARS_RE.sub("", s)
    return s


def safe_atomic_move(src_temp_path: str, dst_final_path: str, max_retries: int = 8) -> None:
    """
    將本機 %TEMP% 產生的完整檔案安全移動至最終目標路徑。
    內建指數退避重試機制，防止 Windows Defender / 企業防毒軟體即時掃描造成的 WinError 32 鎖死。
    """
    dst_dir = os.path.dirname(os.path.abspath(dst_final_path))
    if dst_dir:
        os.makedirs(dst_dir, exist_ok=True)

    for attempt in range(max_retries):
        try:
            # 同磁碟區底層自動執行原子級 os.replace；跨磁碟區則安全複製並清理來源檔
            shutil.move(src_temp_path, dst_final_path)
            return
        except PermissionError:
            if attempt == max_retries - 1:
                raise
            # 指數級漸進等待 (0.1s, 0.2s, 0.4s, 0.8s...) 讓防毒軟體釋放 File Handle
            time.sleep(0.1 * (2 ** min(attempt, 4)))
