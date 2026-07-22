import os
import sys
import time
import platform
import subprocess
import PyKits

def shell_run(cmd: str | list) -> subprocess.CompletedProcess[str]:
    if type(cmd) is list:
        s = subprocess.run(
            cmd,
            capture_output=True,
            text=True
        )
    else:
        s = subprocess.run(
            cmd,
            shell=True,
            capture_output=True,
            text=True
        )
    return s
def systemMessage(message: str): raise NotImplementedError
def mainMessage(message: str): raise NotImplementedError
def warnMessage(message: str): raise NotImplementedError
def successMessage(message: str): raise NotImplementedError
def errorMessage(message: str, title: str="Uh oh!"): raise NotImplementedError
def successMessageBox(message: str, title: str="Success!"): raise NotImplementedError

pip_class = PyKits.pip()
requests = PyKits.request()
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
        super().__init__(os.getenv("ProgramFiles"), os.path.join(os.getenv("LOCALAPPDATA"), "NordWireConnect"), os.path.join(os.getenv("ProgramFiles"), "WireGuard"))
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
    def setup_shortcuts(self, button: bool=False):
        try:
            self._pythoncom.CoInitialize()
            try:
                shell = self._win32com_client.Dispatch("WScript.Shell")
                desk_shortpath = os.path.join(os.environ["USERPROFILE"], "Desktop", f"NordWireConnect.lnk")
                desk_short = shell.CreateShortcut(desk_shortpath)
                desk_short.TargetPath = sys.executable
                desk_short.IconLocation = os.path.join(self.program_files, "Resources", "app_icon.ico")
                desk_short.Description = "Open NordWireConnect for VPN!"
                desk_short.Save()
                startm_shortpath = os.path.join(os.getenv("APPDATA"), "Microsoft", "Windows", "Start Menu", "Programs", f"NordWireConnect.lnk")
                startm_short = shell.CreateShortcut(startm_shortpath)
                startm_short.TargetPath = sys.executable
                startm_short.IconLocation = os.path.join(self.program_files, "Resources", "app_icon.ico")
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
        self._win32event.WaitForSingleObject(handle, self._win32event.INFINITE)
        exit_code = self._win32process.GetExitCodeProcess(handle)
        return exit_code
    def setup_startup_service(self):
        name = "NordWireConnectService"
        service_exe = os.path.join(self.program_files, "NordWireConnectService.exe")
        subprocess.run(["sc", "stop", name], capture_output=True)
        subprocess.run(["sc", "delete", name], capture_output=True)
        subprocess.run(["sc", "create", name, f"binPath={service_exe}", "start=auto"], check=True)
        subprocess.run(["sc", "start", name], check=True)
    def service_exists(self): return nord_socket.exists() and pip_class.getIfProcessIsOpened("NordWireConnectService.")
    def start_service(self):
        with open(os.path.join(self.app_data_path, "StartService.bat"), "w") as f: f.write(f"@echo off\nsc delete NordWireConnectService\nsc create NordWireConnectService binPath= \"{os.path.join(self.program_files, 'NordWireConnectService.exe')}\" start=auto\nsc start NordWireConnectService")
        self.run_sudo_cmd([f"\"{os.path.join(self.app_data_path, 'StartService.bat')}\""])
    def run_wireguard_command(self, command: list): return subprocess.run([self.wg_exe] + command, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
    def install_wireguard_service(self, path: str): return subprocess.run([self.wireguard_exe, "/installtunnelservice", path], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
    def reinstall_wireguard_service(self, name: str, path: str): return subprocess.run([self.wg_exe, "syncconf", name, path], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
    def uninstall_wireguard_service(self, name: str): return subprocess.run([self.wireguard_exe, "/uninstalltunnelservice", name], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
    def _end_wireguard(self, service):
        scm = self._win32service.OpenSCManager(None, None, self._win32service.SC_MANAGER_ENUMERATE_SERVICE)
        statuses = self._win32service.EnumServicesStatus(scm)
        for name, _, _ in statuses:
            name: str = name
            if name.startswith("WireGuardTunnel$"): 
                service.handle_command(f"uninstall-wire-tunnel", {"server": name.replace('WireGuardTunnel$', '', 1)})
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
            try:
                root_key = self._win32api.RegOpenKey(self._win32con.HKEY_LOCAL_MACHINE, p)
                for i in range(self._win32api.RegQueryInfoKey(root_key)[0]):
                    subkey_name = self._win32api.RegEnumKey(root_key, i)
                    subkey = self._win32api.RegOpenKey(root_key, subkey_name)
                    try:
                        name = self._win32api.RegQueryValueEx(subkey, "DisplayName")[0]
                        if "WireGuard" in name: return self._win32api.RegQueryValueEx(subkey, "DisplayVersion")[0]
                    except Exception: continue
                    self._win32api.RegCloseKey(subkey)
                self._win32api.RegCloseKey(root_key)
            except Exception: continue
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
            desk_shortpath = os.path.join(os.environ["USERPROFILE"], "Desktop", f"NordWireConnect.lnk")
            if os.path.exists(desk_shortpath): os.remove(desk_shortpath)
            startm_shortpath = os.path.join(os.getenv("APPDATA"), "Microsoft", "Windows", "Start Menu", "Programs", f"NordWireConnect.lnk")
            if os.path.exists(startm_shortpath): os.remove(startm_shortpath)
        except Exception as e: errorMessage(f"Unable to delete shortcuts: {str(e)}")
    def delete_registry(self):
        try:
            app_reg_path = "Software\\NordWireConnect"
            self._win32api.RegDeleteKey(self._win32con.HKEY_LOCAL_MACHINE, app_reg_path)
            registry_path = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\NordWireConnect"
            self._win32api.RegDeleteKey(self._win32con.HKEY_LOCAL_MACHINE, registry_path)
        except Exception as e: errorMessage(f"Unable to delete registry: {str(e)}")
    def relaunch_as_admin(self):
        self.run_sudo_cmd([sys.executable] + [f'"{arg}"' for arg in sys.argv])
        sys.exit(0)
    def uninstall_app(self):
        if not pip_class.getIfRunningWindowsAdmin():
            mainMessage("Please relaunch this program as Admin to uninstall!")
            self.relaunch_as_admin()
            return
        mainMessage("Uninstalling Service..")
        subprocess.run("sc stop NordWireConnectService", shell=True)
        subprocess.run(
            "sc delete NordWireConnectService",
            shell=True,
            check=False
        )
        subprocess.run("taskkill /IM NordWireConnectService.exe /F", shell=True)
        time.sleep(1)
        if os.path.exists(os.path.join(self.program_files, "Service", "NordWireConnectService.exe")): os.remove(os.path.join(self.program_files, "Service", "NordWireConnectService.exe"))
        mainMessage("Clearing shortcuts on current user..")
        self.delete_shortcuts()
        mainMessage("Clearing Registry..")
        self.delete_registry()
        mainMessage("Removing app..")
        cmd = f'''
        timeout /t 2 >nul
        taskkill /IM NordWireConnect.exe /F
        del "{os.path.join(self.program_files, "NordWireConnect.exe")}"
        rmdir /s /q "{self.program_files}"
        '''
        subprocess.Popen(
            ["cmd.exe", "/c", cmd],
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        sys.exit(0)
    def reset_dns_cache(self):
        try: subprocess.run("ipconfig /flushdns", shell=True, check=True)
        except Exception as e: errorMessage(f"Unable to reset DNS Cache: {str(e)}")
    def reset_networking(self):
        try:
            with open(os.path.join(self.app_data_path, "ResetNetwork.bat"), "w") as f: f.write("@echo off\nnetsh winsock reset\nnetsh int ip reset\nipconfig /release\nipconfig /renew\nipconfig /flushdns\ntimeout 3\nshutdown /f /r /t 0\ntimeout 5")
            self.run_sudo_cmd(["cmd.exe", "/c", f'"{os.path.join(self.app_data_path, "ResetNetwork.bat")}"'])
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
supported_systems = ["Windows"]
not_supported_txt = "NordWireConnect is only supported on Windows and macOS."
if platform.system() == "Windows":
    system = WindowsInterface()