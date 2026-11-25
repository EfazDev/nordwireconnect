import os
import sys
import time
import platform
import subprocess
import PyKits

def shell_run(cmd: str | list, timeout: int=30) -> subprocess.CompletedProcess[str]:
    no_win = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    return subprocess.run(
        cmd,
        shell=not isinstance(cmd, list),
        stdin=subprocess.DEVNULL,
        capture_output=True,
        text=True,
        timeout=timeout,
        creationflags=no_win
    )
def systemMessage(message: str): raise NotImplementedError
def mainMessage(message: str): raise NotImplementedError
def warnMessage(message: str): raise NotImplementedError
def successMessage(message: str): raise NotImplementedError
def errorMessage(message: str, title: str="Uh oh!"): raise NotImplementedError
def successMessageBox(message: str, title: str="Success!"): raise NotImplementedError

pip_class = PyKits.pip()
requests = PyKits.request(throw_exceptions=False)
nord_socket = PyKits.Socket(port=61245)

class SystemInterface:
    program_files: str
    app_data_path: str
    wireguard_location: str
    def __init__(self, program_files: str, app_data_path: str, wireguard_location: str):
        self.program_files = program_files
        self.app_data_path = app_data_path
        self.wireguard_location = wireguard_location
    def encrypt_data(self, key_name: str, data: str): raise NotImplementedError
    def decrypt_data(self, key_name: str, encrypted: str): raise NotImplementedError
    def setup_shortcuts(self, button: bool=False): raise NotImplementedError
    def run_sudo_cmd(self, cmd: list): raise NotImplementedError
    def run_sudo_cmd_with_code(self, cmd: list): raise NotImplementedError
    def setup_startup_service(self): raise NotImplementedError
    def service_exists(self): raise NotImplementedError
    def start_service(self): raise NotImplementedError
    def run_wireguard_command(self, command: list): raise NotImplementedError
    def install_wireguard_service(self, path: str): raise NotImplementedError
    def uninstall_wireguard_service(self, name: str): raise NotImplementedError
    def _end_wireguard(self, service): raise NotImplementedError
    def _end_wireguard_installer(self): raise NotImplementedError
    def get_current_wireguard_version(self): raise NotImplementedError
    def format_for_latest_wireguard(self, wireguard_json: dict): raise NotImplementedError
    def unarchive_asn_database(self, path: str): raise NotImplementedError
    def delete_shortcuts(self): raise NotImplementedError
    def delete_registry(self): raise NotImplementedError
    def relaunch_as_admin(self): raise NotImplementedError
    def uninstall_app(self): raise NotImplementedError
    def reset_dns_cache(self): raise NotImplementedError
    def reset_networking(self): raise NotImplementedError
    def clear_wireguard_conf_files(self): raise NotImplementedError
    def install_wireguard(self, send_command): raise NotImplementedError
    def needing_state(self): raise NotImplementedError
    def unneeding_state(self): raise NotImplementedError
    def enable_debug_console(self, title: str="NordWireConnect", on_streams_changed=None): raise NotImplementedError
    def disable_debug_console(self, on_streams_changed=None): raise NotImplementedError

class WindowsInterface(SystemInterface):
    def __init__(self):
        import win32api # type: ignore
        import win32con # type: ignore
        import win32crypt # type: ignore
        import win32event # type: ignore
        import win32process # type: ignore
        import win32service # type: ignore
        import binascii # type: ignore
        import win32com.client # type: ignore
        import win32com.shell.shell as winshell # type: ignore
        import pythoncom # type: ignore
        import ctypes # type: ignore
        import gzip # type: ignore
        pf_dir = os.getenv("ProgramFiles") or r"C:\Program Files"
        appdata_dir = os.getenv("LOCALAPPDATA") or os.path.join(os.getenv("USERPROFILE") or r"C:\Windows\System32\config\systemprofile", "AppData", "Local")
        super().__init__(pf_dir, os.path.join(appdata_dir, "NordWireConnect"), os.path.join(pf_dir, "WireGuard"))
        self._win32crypt = win32crypt
        self._win32com_client = win32com.client
        self._pythoncom = pythoncom
        self._ctypes = ctypes
        self._binascii = binascii
        self._win32service = win32service
        self._win32event = win32event
        self._win32process = win32process
        self._win32api = win32api
        self._win32con = win32con
        self._winshell = winshell
        self._gzip = gzip
        self.console_allocated = False
        self.console_streams = []
        self.null_streams = []
        self.app_install_path = os.path.join(self.program_files, "NordWireConnect")
        self.wireguard_exe = os.path.join(self.program_files, "WireGuard", "wireguard.exe")
        self.wg_exe = os.path.join(self.program_files, "WireGuard", "wg.exe")
    def encrypt_data(self, key_name: str, data: str): 
        return self._binascii.hexlify(self._win32crypt.CryptProtectData(
            data.encode("utf-16-le"), None, None, None, None, 0
        )).decode()
    def decrypt_data(self, key_name: str, encrypted: str): 
        _, decrypted_password_string = self._win32crypt.CryptUnprotectData(
            self._binascii.unhexlify(encrypted), None, None, None, 0
        )
        return decrypted_password_string.decode("utf-16-le")
    def get_shortcut_folder(self, name: str) -> str:
        try:
            self._pythoncom.CoInitialize()
            try:
                p = self._win32com_client.Dispatch("WScript.Shell").SpecialFolders(name)
                if p and os.path.isdir(p): return p
            finally: self._pythoncom.CoUninitialize()
        except Exception: pass
        if name == "Desktop": return os.path.join(os.environ.get("USERPROFILE", ""), "Desktop")
        return os.path.join(os.getenv("APPDATA", ""), "Microsoft", "Windows", "Start Menu", "Programs")
    def setup_shortcuts(self, button: bool=False):
        try:
            self._pythoncom.CoInitialize()
            try:
                shell = self._win32com_client.Dispatch("WScript.Shell")
                installed_exe = os.path.join(self.app_install_path, "NordWireConnect.exe")
                target_exe = installed_exe if os.path.exists(installed_exe) else os.path.abspath(sys.argv[0] or sys.executable)
                icon_path = os.path.join(self.app_install_path, "Resources", "app_icon.ico")
                desk_short = shell.CreateShortcut(os.path.join(self.get_shortcut_folder("Desktop"), "NordWireConnect.lnk"))
                desk_short.TargetPath = target_exe
                desk_short.IconLocation = icon_path
                desk_short.Description = "Open NordWireConnect for VPN!"
                desk_short.Save()
                startm_short = shell.CreateShortcut(os.path.join(self.get_shortcut_folder("Programs"), "NordWireConnect.lnk"))
                startm_short.TargetPath = target_exe
                startm_short.IconLocation = icon_path
                startm_short.Description = "Open NordWireConnect for VPN!"
                startm_short.Save()
                if button == True: successMessageBox(f"Successfully set up shortcuts for NordWireConnect!", "NordWireConnect")
            finally: self._pythoncom.CoUninitialize()
        except Exception as e: errorMessage(f"Unable to save shortcuts: {str(e)}")
    def run_sudo_cmd(self, cmd: list):
        self._ctypes.windll.shell32.ShellExecuteW(
            None,
            "runas",
            cmd[0],
            " ".join(cmd[1:]),
            None,
            1
        )
    def run_sudo_cmd_with_code(self, cmd: list):
        proc = self._winshell.ShellExecuteEx(
            lpVerb="runas",
            lpFile=cmd[0],
            lpParameters=" ".join(cmd[1:]),
            nShow=self._win32con.SW_SHOWNORMAL,
            fMask=0x00000040
        )
        handle = proc["hProcess"]
        try:
            self._win32event.WaitForSingleObject(handle, self._win32event.INFINITE)
            exit_code = self._win32process.GetExitCodeProcess(handle)
            return exit_code
        finally:
            try: self._win32api.CloseHandle(handle)
            except Exception: pass
    def setup_startup_service(self):
        name = "NordWireConnectService"
        service_exe = os.path.join(self.app_install_path, "NordWireConnectService.exe")
        subprocess.run(["sc", "stop", name], capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW, stdin=subprocess.DEVNULL)
        subprocess.run(["sc", "delete", name], capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW, stdin=subprocess.DEVNULL)
        subprocess.run(["sc", "create", name, f"binPath={service_exe}", "start=auto"], check=True, creationflags=subprocess.CREATE_NO_WINDOW, stdin=subprocess.DEVNULL)
        subprocess.run(["sc", "start", name], check=True, creationflags=subprocess.CREATE_NO_WINDOW, stdin=subprocess.DEVNULL)
    def service_exists(self): return nord_socket.exists() and pip_class.getIfProcessIsOpened("NordWireConnectService")
    def start_service(self):
        service_exe = os.path.join(self.app_install_path, "NordWireConnectService.exe")
        bat_path = os.path.join(self.app_data_path, "StartService.bat")
        if self._service_exists("NordWireConnectService"):
            with open(bat_path, "w") as f: f.write("@echo off\nsc start NordWireConnectService\n")
        else:
            with open(bat_path, "w") as f: f.write(f"@echo off\nsc delete NordWireConnectService\nsc create NordWireConnectService binPath= \"{service_exe}\" start=auto\nsc start NordWireConnectService\n")
        self.run_sudo_cmd([f"\"{bat_path}\""])
    def run_wireguard_command(self, command: list, timeout: int=10): return subprocess.run([self.wg_exe] + command, capture_output=True, text=True, timeout=timeout, creationflags=subprocess.CREATE_NO_WINDOW, stdin=subprocess.DEVNULL)
    def get_service_state(self, service_name: str):
        scm = None
        svc = None
        try:
            scm = self._win32service.OpenSCManager(None, None, self._win32service.SC_MANAGER_CONNECT)
            svc = self._win32service.OpenService(scm, service_name, self._win32service.SERVICE_QUERY_STATUS)
            return self._win32service.QueryServiceStatus(svc)[1]
        except Exception: return None
        finally:
            if svc:
                try: self._win32service.CloseServiceHandle(svc)
                except Exception: pass
            if scm:
                try: self._win32service.CloseServiceHandle(scm)
                except Exception: pass
    def _service_exists(self, service_name: str) -> bool: return self.get_service_state(service_name) is not None
    def wait_for_service_deleted(self, service_name: str, timeout: float=12.0) -> bool:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if not self._service_exists(service_name): return True
            time.sleep(0.1)
        return not self._service_exists(service_name)
    def install_wireguard_service(self, path: str):
        tunnel_name = os.path.splitext(os.path.basename(path))[0]
        if self._service_exists("WireGuardTunnel$" + tunnel_name): self.uninstall_wireguard_service(tunnel_name)
        return subprocess.run([self.wireguard_exe, "/installtunnelservice", path], capture_output=True, text=True, timeout=25, check=False, creationflags=subprocess.CREATE_NO_WINDOW, stdin=subprocess.DEVNULL)
    def reinstall_wireguard_service(self, name: str, path: str): return subprocess.run([self.wg_exe, "syncconf", name, path], capture_output=True, text=True, timeout=15, check=False, creationflags=subprocess.CREATE_NO_WINDOW, stdin=subprocess.DEVNULL)
    def uninstall_wireguard_service(self, name: str):
        svc_name = "WireGuardTunnel$" + name
        if not self._service_exists(svc_name): return subprocess.CompletedProcess(args=[self.wireguard_exe, "/uninstalltunnelservice", name], returncode=0, stdout="", stderr="")
        res = subprocess.run([self.wireguard_exe, "/uninstalltunnelservice", name], capture_output=True, text=True, timeout=20, check=False, creationflags=subprocess.CREATE_NO_WINDOW, stdin=subprocess.DEVNULL)
        if res.returncode != 0 and self._service_exists(svc_name):
            subprocess.run(["sc", "stop", svc_name], capture_output=True, timeout=10, creationflags=subprocess.CREATE_NO_WINDOW, stdin=subprocess.DEVNULL)
            subprocess.run(["sc", "delete", svc_name], capture_output=True, timeout=10, creationflags=subprocess.CREATE_NO_WINDOW, stdin=subprocess.DEVNULL)
        self.wait_for_service_deleted(svc_name, timeout=12.0)
        return res
    def _end_wireguard(self, service):
        scm = self._win32service.OpenSCManager(None, None, self._win32service.SC_MANAGER_ENUMERATE_SERVICE)
        try: statuses = self._win32service.EnumServicesStatus(scm)
        finally:
            try: self._win32service.CloseServiceHandle(scm)
            except Exception: pass
        for name, _, _ in statuses:
            name: str = name
            if name.startswith("WireGuardTunnel$"): service.handle_command({"cmd": "uninstall-wire-tunnel", "server": name.replace("WireGuardTunnel$", "", 1)})
        self.connected_tunnel = None
        return {"suc": True, "res": 0}
    def _end_wireguard_installer(self):
        end = shell_run("taskkill /IM wireguard-installer.exe /F")
        return {"suc": end.returncode == 0, "res": end.returncode}
    def get_current_wireguard_version(self):
        for p in [
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        ]:
            root_key = None
            try:
                root_key = self._win32api.RegOpenKey(self._win32con.HKEY_LOCAL_MACHINE, p)
                count = self._win32api.RegQueryInfoKey(root_key)[0]
                for i in range(count):
                    subkey = None
                    try:
                        subkey_name = self._win32api.RegEnumKey(root_key, i)
                        subkey = self._win32api.RegOpenKey(root_key, subkey_name)
                        name = self._win32api.RegQueryValueEx(subkey, "DisplayName")[0]
                        if "WireGuard" in name: return self._win32api.RegQueryValueEx(subkey, "DisplayVersion")[0]
                    except Exception: continue
                    finally:
                        if subkey:
                            try: self._win32api.RegCloseKey(subkey)
                            except Exception: pass
            except Exception: continue
            finally:
                if root_key:
                    try: self._win32api.RegCloseKey(root_key)
                    except Exception: pass
        return None
    def format_for_latest_wireguard(self, wireguard_json: dict):
        if wireguard_json.get("windowsdl-win"):
            v = wireguard_json.get("windowsdl-win", {})
            if v.get("uptodate") and v.get("version"): return v
    def unarchive_asn_database(self, path: str):
        results = []
        with self._gzip.open(path, "rt", encoding="utf-8") as f:
            for line in f:
                parts = line.strip().split("\t")
                if len(parts) >= 5:
                    start_ip = parts[0]
                    end_ip = parts[1]
                    asn = parts[2]
                    asn_name = parts[4]
                    if asn != "0":
                        results.append({
                            "s_ip": start_ip,
                            "e_ip": end_ip,
                            "asn": asn,
                            "asn_name": asn_name
                        })
        return results
    def delete_shortcuts(self):
        try:
            for folder in ("Desktop", "Programs"):
                p = os.path.join(self.get_shortcut_folder(folder), "NordWireConnect.lnk")
                if os.path.exists(p): os.remove(p)
        except Exception as e: errorMessage(f"Unable to delete shortcuts: {str(e)}")
    def delete_registry(self):
        for reg_path in ("Software\\NordWireConnect", r"Software\Microsoft\Windows\CurrentVersion\Uninstall\NordWireConnect"):
            try: self._win32api.RegDeleteKey(self._win32con.HKEY_LOCAL_MACHINE, reg_path)
            except Exception as e:
                if getattr(e, "winerror", None) != 2 and getattr(e, "args", [None])[0] != 2: errorMessage(f"Unable to delete registry: {str(e)}")
    def relaunch_as_admin(self):
        is_compiled = bool(getattr(sys, "frozen", False) or "__compiled__" in globals())
        exe = os.path.abspath(sys.argv[0]) if is_compiled else sys.executable
        args = sys.argv[1:] if is_compiled else sys.argv
        self.run_sudo_cmd([exe] + [f"\"{arg}\"" for arg in args])
        sys.exit(0)
    def uninstall_app(self):
        if not pip_class.getIfRunningWindowsAdmin():
            mainMessage("Please relaunch this program as Admin to uninstall!")
            self.relaunch_as_admin()
            return
        mainMessage("Stopping active WireGuard tunnel..")
        try: self.uninstall_wireguard_service("NordWireConnect")
        except Exception: pass
        mainMessage("Uninstalling Service..")
        subprocess.run(["sc", "stop", "NordWireConnectService"], capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW, stdin=subprocess.DEVNULL)
        subprocess.run(["sc", "delete", "NordWireConnectService"], capture_output=True, check=False, creationflags=subprocess.CREATE_NO_WINDOW, stdin=subprocess.DEVNULL)
        subprocess.run(["taskkill", "/IM", "NordWireConnectService.exe", "/F"], capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW, stdin=subprocess.DEVNULL)
        time.sleep(1)
        if os.path.exists(os.path.join(self.app_install_path, "Service", "NordWireConnectService.exe")):
            try: os.remove(os.path.join(self.app_install_path, "Service", "NordWireConnectService.exe"))
            except Exception: pass
        mainMessage("Clearing shortcuts on current user..")
        self.delete_shortcuts()
        mainMessage("Clearing Registry..")
        self.delete_registry()
        mainMessage("Removing app..")
        target_dir = os.path.normpath(self.app_install_path)
        if os.path.basename(target_dir).lower() != "nordwireconnect":
            errorMessage(f"Refusing to remove unexpected install path: {target_dir}")
            return
        cmd = f"timeout /t 2 >nul & taskkill /IM NordWireConnect.exe /F >nul 2>&1 & rmdir /s /q \"{target_dir}\""
        subprocess.Popen(["cmd.exe", "/c", cmd], creationflags=subprocess.CREATE_NO_WINDOW, stdin=subprocess.DEVNULL)
        sys.exit(0)
    def reset_dns_cache(self):
        try: subprocess.run(["ipconfig", "/flushdns"], check=True, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW, stdin=subprocess.DEVNULL)
        except Exception as e: errorMessage(f"Unable to reset DNS Cache: {str(e)}")
    def reset_networking(self):
        try:
            with open(os.path.join(self.app_data_path, "ResetNetwork.bat"), "w") as f: f.write("@echo off\nnetsh winsock reset\nnetsh int ip reset\nipconfig /release\nipconfig /renew\nipconfig /flushdns\ntimeout 3\nshutdown /f /r /t 0\ntimeout 5")
            self.run_sudo_cmd(["cmd.exe", "/c", f"\"{os.path.join(self.app_data_path, 'ResetNetwork.bat')}\""])
        except Exception as e: errorMessage(f"Unable to reset Networking: {str(e)}")
    def clear_wireguard_conf_files(self):
        try:
            mainMessage("Clearing Wireguard .conf files..")
            for filename in os.listdir(self.app_data_path):
                if filename.endswith(".conf"): os.remove(os.path.join(self.app_data_path, filename))
        except Exception as e: errorMessage(f"Unable to clear Wireguard .conf files: {str(e)}")
    def install_wireguard(self, send_command):
        download_url = "https://download.wireguard.com/windows-client/wireguard-installer.exe"
        installer_path = os.path.join(self.app_data_path, "wireguard-installer.exe")
        mainMessage("Downloading WireGuard installer...")
        download = requests.download(download_url, installer_path)
        if download.ok:
            mainMessage("Installing WireGuard...")
            wireguard_exit_code = self.run_sudo_cmd_with_code([installer_path])
            send_command("combined-end-wireguard")
            time.sleep(1)
            if os.path.exists(installer_path): os.remove(installer_path)
            if wireguard_exit_code == 0:
                mainMessage("WireGuard installed successfully.")
                time.sleep(3)
                return True
            else: return False
        else: return False
    def needing_state(self): self._ctypes.windll.kernel32.SetThreadExecutionState(0x80000001 | 0x00000040)
    def unneeding_state(self): self._ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)
    def _enable_ansi_on_console(self):
        kernel32 = self._ctypes.windll.kernel32
        ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004
        for handle_id in (-11, -12):
            handle = kernel32.GetStdHandle(handle_id)
            if not handle or handle == -1: continue
            mode = self._ctypes.c_uint32()
            if not kernel32.GetConsoleMode(handle, self._ctypes.byref(mode)): continue
            kernel32.SetConsoleMode(handle, mode.value | ENABLE_VIRTUAL_TERMINAL_PROCESSING)
    def _disable_console_close_button(self):
        hwnd = self._ctypes.windll.kernel32.GetConsoleWindow()
        if not hwnd: return
        hmenu = self._ctypes.windll.user32.GetSystemMenu(hwnd, False)
        if not hmenu: return
        self._ctypes.windll.user32.DeleteMenu(hmenu, 0xF060, 0x0000)
    def enable_debug_console(self, title: str="NordWireConnect", on_streams_changed=None) -> bool:
        kernel32 = self._ctypes.windll.kernel32
        if kernel32.GetConsoleWindow() != 0:
            self._enable_ansi_on_console()
            return False
        if not kernel32.AllocConsole(): return False
        self.console_allocated = True
        kernel32.SetConsoleTitleW(title)
        self._disable_console_close_button()
        self._enable_ansi_on_console()
        old_nulls = self.null_streams
        self.null_streams = []
        out_f = open("CONOUT$", "w", encoding="utf-8", buffering=1)
        err_f = open("CONOUT$", "w", encoding="utf-8", buffering=1)
        in_f = open("CONIN$", "r", encoding="utf-8")
        self.console_streams = [out_f, err_f, in_f]
        sys.stdout, sys.stderr, sys.stdin = out_f, err_f, in_f
        if on_streams_changed: on_streams_changed()
        for s in old_nulls:
            try: s.close()
            except Exception: pass
        return True
    def disable_debug_console(self, on_streams_changed=None) -> bool:
        if not self.console_allocated: return False
        old_streams = self.console_streams
        self.console_streams = []
        null_out = open(os.devnull, "w", encoding="utf-8")
        null_err = open(os.devnull, "w", encoding="utf-8")
        null_in = open(os.devnull, "r", encoding="utf-8")
        self.null_streams = [null_out, null_err, null_in]
        sys.stdout, sys.stderr, sys.stdin = null_out, null_err, null_in
        if on_streams_changed: on_streams_changed()
        for s in old_streams:
            try: s.close()
            except Exception: pass
        self._ctypes.windll.kernel32.FreeConsole()
        self.console_allocated = False
        return True
supported_systems = ["Windows"]
not_supported_txt = "NordWireConnect is only supported on Windows and macOS."
if platform.system() == "Windows": system = WindowsInterface()