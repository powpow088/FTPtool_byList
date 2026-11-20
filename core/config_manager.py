import base64
import json
import os
from pathlib import Path
from typing import Tuple, List

from models.data_models import SFTPConfig, SearchSettings, RuleItem, ConflictStrategy, MatchMode


class ConfigManager:
    """負責讀寫 config.json，並對儲存的密碼進行混淆保護"""

    DEFAULT_CONFIG_PATH = Path("config.json")
    _OBFUSCATION_SALT = b"SFTP_BATCH_DOWNLOADER_2026_KEY"

    @classmethod
    def load_config(cls, path: Path = DEFAULT_CONFIG_PATH) -> Tuple[SFTPConfig, SearchSettings, List[RuleItem]]:
        """從指定路徑讀取設定檔，若檔案不存在或格式損毀則安全回傳預設值"""
        if not path.exists():
            return SFTPConfig(), SearchSettings(), []
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)

            conn = data.get("connection", {})
            save_pwd = bool(conn.get("save_password", False))
            raw_pwd = cls._deobfuscate(conn.get("saved_password", "")) if save_pwd else ""

            sftp_cfg = SFTPConfig(
                host=str(conn.get("host", "")).strip(),
                port=int(conn.get("port", 22)),
                username=str(conn.get("username", "")).strip(),
                password=raw_pwd,
                save_password=save_pwd,
                initial_path=str(conn.get("initial_path", "/")).strip()
            )

            ss = data.get("search_settings", {})
            strategy_val = ss.get("conflict_strategy", ConflictStrategy.SIZE_MTIME.value)
            try:
                strategy = ConflictStrategy(strategy_val)
            except ValueError:
                strategy = ConflictStrategy.SIZE_MTIME

            match_mode_val = ss.get("match_mode", MatchMode.FUZZY.value)
            try:
                match_mode = MatchMode(match_mode_val)
            except ValueError:
                match_mode = MatchMode.FUZZY

            search_cfg = SearchSettings(
                global_pattern=str(ss.get("global_pattern", "")).strip(),
                conflict_strategy=strategy,
                local_save_path=str(ss.get("local_save_path", "")).strip(),
                match_mode=match_mode,
                request_delay_ms=int(ss.get("request_delay_ms", 50)),
                session_refresh_interval=int(ss.get("session_refresh_interval", 150))
            )

            rules = []
            for r in data.get("rules", []):
                dirs_data = r.get("dirs")
                if isinstance(dirs_data, list):
                    item = RuleItem(
                        dirs=[str(d) for d in dirs_data],
                        file_pattern=str(r.get("file_pattern", ""))
                    )
                else:
                    item = RuleItem(
                        l1_keyword=str(r.get("l1_keyword", "")),
                        l2_keyword=str(r.get("l2_keyword", "")),
                        file_pattern=str(r.get("file_pattern", ""))
                    )
                rules.append(item)

            return sftp_cfg, search_cfg, rules
        except Exception:
            # 發生非預期格式損毀時，回傳預設值防止崩潰
            return SFTPConfig(), SearchSettings(), []

    @classmethod
    def save_config(
        cls,
        sftp_cfg: SFTPConfig,
        search_cfg: SearchSettings,
        rules: List[RuleItem],
        path: Path = DEFAULT_CONFIG_PATH
    ) -> None:
        """將設定序列化儲存至 config.json，自動過濾全空規則列與混淆密碼"""
        encoded_pwd = cls._obfuscate(sftp_cfg.password) if sftp_cfg.save_password else ""
        payload = {
            "connection": {
                "host": sftp_cfg.host.strip(),
                "port": sftp_cfg.port,
                "username": sftp_cfg.username.strip(),
                "save_password": sftp_cfg.save_password,
                "saved_password": encoded_pwd,
                "initial_path": sftp_cfg.normalized_initial_path()
            },
            "search_settings": {
                "global_pattern": search_cfg.global_pattern.strip(),
                "conflict_strategy": search_cfg.conflict_strategy.value,
                "local_save_path": search_cfg.local_save_path.strip(),
                "match_mode": search_cfg.match_mode.value,
                "request_delay_ms": search_cfg.request_delay_ms,
                "session_refresh_interval": search_cfg.session_refresh_interval
            },
            "rules": [
                {
                    "dirs": r.dirs,
                    "l1_keyword": r.l1_keyword,
                    "l2_keyword": r.l2_keyword,
                    "file_pattern": r.file_pattern
                }
                for r in rules if not r.is_completely_empty()
            ]
        }
        # 確保父目錄存在
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

    @classmethod
    def _obfuscate(cls, plain_text: str) -> str:
        """使用 XOR 與 URLSafe-Base64 進行密碼混淆"""
        if not plain_text:
            return ""
        data = plain_text.encode("utf-8")
        salt_len = len(cls._OBFUSCATION_SALT)
        xored = bytes(b ^ cls._OBFUSCATION_SALT[i % salt_len] for i, b in enumerate(data))
        return base64.urlsafe_b64encode(xored).decode("ascii")

    @classmethod
    def _deobfuscate(cls, encoded_text: str) -> str:
        """還原混淆之密碼"""
        if not encoded_text:
            return ""
        try:
            xored = base64.urlsafe_b64decode(encoded_text.encode("ascii"))
            salt_len = len(cls._OBFUSCATION_SALT)
            data = bytes(b ^ cls._OBFUSCATION_SALT[i % salt_len] for i, b in enumerate(xored))
            return data.decode("utf-8")
        except Exception:
            return ""
