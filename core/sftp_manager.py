import socket
import stat
from typing import List, Tuple, Callable, Optional
import paramiko
from models.data_models import SFTPConfig


class SFTPManager:
    """封裝 Paramiko SSH/SFTP 連線、KeepAlive 心跳與安全重連機制"""

    SOCKET_TIMEOUT = 15.0
    KEEPALIVE_INTERVAL = 30

    def __init__(self, config: SFTPConfig) -> None:
        self._config = config
        self._ssh: Optional[paramiko.SSHClient] = None
        self._sftp: Optional[paramiko.SFTPClient] = None

    def connect(self) -> None:
        """建立 SSH 與 SFTP 連線，並配置逾時與 KeepAlive 心跳"""
        self.close()
        self._ssh = paramiko.SSHClient()
        self._ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

        self._ssh.connect(
            hostname=self._config.host.strip(),
            port=self._config.port,
            username=self._config.username.strip(),
            password=self._config.password,
            timeout=self.SOCKET_TIMEOUT,
            banner_timeout=self.SOCKET_TIMEOUT,
            auth_timeout=self.SOCKET_TIMEOUT,
            look_for_keys=False,
            allow_agent=False
        )

        transport = self._ssh.get_transport()
        if transport:
            transport.set_keepalive(self.KEEPALIVE_INTERVAL)

        self._sftp = self._ssh.open_sftp()
        if self._sftp.get_channel():
            self._sftp.get_channel().settimeout(self.SOCKET_TIMEOUT)

    def test_connection(self) -> Tuple[bool, str]:
        """供 GUI『測試連線』按鈕呼叫，驗證認證與 Initial Path 是否存在且為資料夾"""
        try:
            self.connect()
            init_path = self._config.normalized_initial_path()
            attr = self._sftp.stat(init_path)
            if not stat.S_ISDIR(attr.st_mode):
                return False, f"遠端路徑存在但非目錄：{init_path}"
            return True, f"連線成功！已驗證遠端目錄存在：{init_path}"
        except paramiko.AuthenticationException:
            return False, "認證失敗：帳號或密碼錯誤。"
        except FileNotFoundError:
            return False, f"遠端初始路徑不存在：{self._config.normalized_initial_path()}"
        except PermissionError:
            return False, f"權限不足：無法存取遠端初始路徑 {self._config.normalized_initial_path()}"
        except (socket.timeout, TimeoutError):
            return False, f"連線逾時（超過 {int(self.SOCKET_TIMEOUT)} 秒），請檢查 Host 與 Port。"
        except Exception as e:
            return False, f"連線發生異常：{type(e).__name__} - {str(e)}"
        finally:
            self.close()

    def reconnect(self) -> None:
        """安全斷線並重建連線"""
        self.close()
        self.connect()

    def listdir_attr(self, remote_path: str) -> List[paramiko.SFTPAttributes]:
        """取得遠端目錄內容屬性清單"""
        if not self._sftp:
            raise paramiko.SSHException("SFTP 尚未連線")
        return self._sftp.listdir_attr(remote_path)

    def download_file(
        self,
        remote_path: str,
        local_temp_path: str,
        callback: Optional[Callable[[int, int], None]] = None
    ) -> None:
        """執行分塊下載至本地暫存檔"""
        if not self._sftp:
            raise paramiko.SSHException("SFTP 尚未連線")
        self._sftp.get(remote_path, local_temp_path, callback=callback)

    def close(self) -> None:
        """安全釋放 SFTP 與 SSH 資源"""
        if self._sftp:
            try:
                self._sftp.close()
            except Exception:
                pass
            self._sftp = None
        if self._ssh:
            try:
                self._ssh.close()
            except Exception:
                pass
            self._ssh = None
