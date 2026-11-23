# core/db_config.py
import os
from typing import Dict, Any

# ==============================================================================
# MS SQL Server 資料庫連線設定
# 使用者可在此直接填入連線資訊；未來部署至後端伺服器時，亦可透過環境變數注入
# ==============================================================================

DB_CONFIG: Dict[str, Any] = {
    # 伺服器主機名稱或 IP 位址 (例如: "192.168.1.100" 或 "sqlserver.corp.local")
    "server": os.getenv("MSSQL_HOST", "YOUR_SERVER_IP"),

    # 連接埠 (預設 1433)
    "port": int(os.getenv("MSSQL_PORT", "1433")),

    # 目標資料庫名稱
    "database": os.getenv("MSSQL_DB", "dc"),

    # 登入使用者帳號
    "user": os.getenv("MSSQL_USER", "YOUR_USERNAME"),

    # 登入密碼
    "password": os.getenv("MSSQL_PASSWORD", "YOUR_PASSWORD"),

    # 資料表完整路徑
    "table_name": "[dc].[dbo].[YM_ULT_UnitBase]",

    # 查詢逾時時間 (秒)
    "timeout": int(os.getenv("MSSQL_TIMEOUT", "30")),

    # 是否啟用資料庫查詢 (設為 False 則完全跳過 DB 連線)
    "enabled": True,
}


def is_db_configured() -> bool:
    """判斷使用者是否已填入有效之資料庫連線資訊"""
    if not DB_CONFIG.get("enabled", True):
        return False
    server = str(DB_CONFIG.get("server", "")).strip()
    user = str(DB_CONFIG.get("user", "")).strip()
    if not server or server == "YOUR_SERVER_IP":
        return False
    if not user or user == "YOUR_USERNAME":
        return False
    return True
