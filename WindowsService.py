"""
NordWireConnect

Made by @EfazDev
https://www.efaz.dev
"""

# Modules
import os
import sys
import time
import PyKits
import System
import logging
import win32ts # type: ignore
import datetime
import win32con # type: ignore
import subprocess
import win32event # type: ignore
import win32service # type: ignore
import win32process # type: ignore
import win32profile # type: ignore
import win32security # type: ignore
import ntsecuritycon # type: ignore
import servicemanager # type: ignore
import win32serviceutil # type: ignore

# Variables
service_pipe = r"\\.\pipe\NordWireConnect"
program_files = os.path.join(os.getenv("ProgramFiles"), "NordWireConnect")
nordwireconnect_location = os.path.join(program_files, "NordWireConnect.exe")
wireguard_location = os.path.join(os.getenv("ProgramFiles"), "WireGuard")
version = "1.3.5j"
windows_service = System.WindowsInterface()
nord_socket = PyKits.Socket(port=61245)
colors_class = PyKits.Colors()
pip_class = PyKits.pip()

# Logging
def info(message: str): servicemanager.LogInfoMsg(message); mainMessage(message)
def error(message: str): servicemanager.LogErrorMsg(message); errorMessage(message)
def warn(message: str): servicemanager.LogWarningMsg(message); warnMessage(message)
def systemMessage(message: str): colors_class.print(message, colors_class.hex_to_ansi2("#3E5FFF"))
def mainMessage(message: str): colors_class.print(message, 15)
def errorMessage(message: str, title: str="Uh oh!"):  colors_class.print(message, 9)
def warnMessage(message: str): colors_class.print(message, 11)
def successMessage(message: str): colors_class.print(message, 10)
def setup_logging():
    handler_name = "Service"
    log_path = os.path.join(program_files, "Logs")
    if not os.path.exists(log_path): os.makedirs(log_path,mode=511)
    generated_file_name = f'NordWireConnect_{handler_name}_{datetime.datetime.now().strftime("%B_%d_%Y_%H_%M_%S_%f")}.log' 
    if hasattr(sys.stdout, "reconfigure"): sys.stdout.reconfigure(encoding='utf-8')
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    file_handler = logging.FileHandler(os.path.join(log_path, generated_file_name), encoding="utf-8")
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    if sys.stdout:
        try:
            stdout_stream = logging.StreamHandler(sys.stdout)
            stdout_stream.setLevel(logging.INFO)
            stdout_stream.setFormatter(logging.Formatter("%(message)s"))
            logger.addHandler(stdout_stream)
        except Exception: pass
    logger.addHandler(file_handler)
    sys.stdout = PyKits.stdout(logger, logging.INFO)
    sys.stderr = PyKits.stdout(logger, logging.ERROR)
    System.systemMessage = systemMessage
    System.mainMessage = mainMessage
    System.successMessage = successMessage
    System.warnMessage = warnMessage
    System.errorMessage = errorMessage

# Pipe Administration
def create_pipe_sa():
    sd = win32security.SECURITY_DESCRIPTOR()
    sd.Initialize()
    dacl = win32security.ACL()
    dacl.AddAccessAllowedAce(
        win32security.ACL_REVISION,
        ntsecuritycon.FILE_ALL_ACCESS,
        win32security.LookupAccountName(None, "SYSTEM")[0]
    )
    dacl.AddAccessAllowedAce(
        win32security.ACL_REVISION,
        ntsecuritycon.FILE_GENERIC_READ | ntsecuritycon.FILE_GENERIC_WRITE,
        win32security.LookupAccountName(None, "Authenticated Users")[0]
    )
    sd.SetSecurityDescriptorDacl(True, dacl, False)
    sa = win32security.SECURITY_ATTRIBUTES()
    sa.SECURITY_DESCRIPTOR = sd
    return sa

# Unbricking
def shell_run(cmd: str | list) -> subprocess.CompletedProcess[str]: 
    s = System.shell_run(cmd)
    if s.returncode != 0: error(f"Command failed ({cmd}): {s.stderr.strip()}")
    return s
def check_for_unbricks(idx: str) -> bool:
    out = shell_run(f"netsh interface ipv4 show interface {idx}").stdout.lower()
    for ln in out.splitlines():
        if "weak host sends" in ln and "enabled" in ln: return True
        if "weak host receives" in ln and "enabled" in ln: return True
        if "forwarding" in ln and "enabled" in ln: return True
    out = shell_run(f"netsh interface ipv6 show interface {idx}").stdout.lower()
    for ln in out.splitlines():
        if "weak host sends" in ln and "enabled" in ln: return True
        if "weak host receives" in ln and "enabled" in ln: return True
        if "forwarding" in ln and "enabled" in ln: return True
    return False
def wait_for_service(name: str, timeout: int=30, status: int=win32service.SERVICE_RUNNING) -> bool:
    start = time.time()
    while time.time() - start < timeout:
        try:
            current_status = win32serviceutil.QueryServiceStatus(name)[1]
            if current_status == status: return True
        except Exception: pass
        time.sleep(0.5)
    return False

# Service Class
class NordWireConnectService(win32serviceutil.ServiceFramework):
    _svc_name_ = "NordWireConnectService"
    _svc_display_name_ = "NordWireConnectService"
    _svc_description_ = "Handles NordWireConnect with UI Handling, Background Tasks and WireGuard Management!"
    _svc_controls_accepted_ = win32service.SERVICE_ACCEPT_POWEREVENT | win32service.SERVICE_ACCEPT_SESSIONCHANGE
    def __init__(self, args):
        super().__init__(args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.ui_running = False
        self.last_session_id = None
        self.connected_tunnel = None
        self.cached_routing = []
        self.cleared_older = False
        self.prevent_opening = False
    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)
    def SvcDoRun(self):
        info("NordWireConnect service starting")
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        if nord_socket.exists(): nord_socket.wait_till_free()
        nord_socket.subscribe("NordWireConnectCommands", self.handle_command)
        nord_socket.listen()
        while True:
            if win32event.WaitForSingleObject(self.stop_event, 2000) == win32event.WAIT_OBJECT_0: break
            self.ensure_ui_running()
    def SvcOtherEx(self, control: int, event_type: int, data):
        if control == win32service.SERVICE_CONTROL_POWEREVENT:
            if event_type == win32con.PBT_APMSUSPEND:
                info("System suspending")
                self.ui_running = False
                self.last_session_id = None
            elif event_type in (win32con.PBT_APMRESUMEAUTOMATIC, win32con.PBT_APMRESUMECRITICAL):
                info("System resumed")
                time.sleep(5)
                self.ui_running = False
                self.last_session_id = None
    def handle_command(self, cmd_inf: dict) -> dict:
        cmd = "<unknown>"
        try:
            cmd = cmd_inf.get("cmd", "")
            if cmd != "wireguard-check": info(f"Received command: {cmd}")
            if cmd == "end-wireguard": return windows_service._end_wireguard(self)
            elif cmd == "ui-opening": self.prevent_opening = False; return {"suc": True}
            elif cmd == "ui-closing": self.prevent_opening = True; return {"suc": True}
            elif cmd == "service-version": return {"suc": True, "version": version}
            elif cmd == "end-wireguard-installer": return windows_service._end_wireguard_installer()
            elif cmd == "uninstall-wire-tunnel":
                service_name = "NordWireConnect"
                if not service_name: return {"suc": False, "res": "No service name provided"}
                remove = windows_service.uninstall_wireguard_service(service_name)
                if self.connected_tunnel == service_name: self.connected_tunnel = None
                return {"suc": remove.returncode == 0, "res": remove.returncode}
            elif cmd == "install-wire-tunnel":
                self.handle_command({"cmd": "unbrick-adapter"})
                config_path = cmd_inf.get("config_path", None)
                if not config_path: return {"suc": False, "res": "No config path provided"}
                add = windows_service.install_wireguard_service(config_path)
                self.connected_tunnel = "NordWireConnect"
                wait_for_service("WireGuardTunnel$NordWireConnect")
                return {"suc": add.returncode == 0, "res": add.returncode}
            elif cmd == "combined-disconnect":
                server_name = cmd_inf.get("known_server", None)
                if server_name: self.handle_command({"cmd": "uninstall-wire-tunnel", "server": server_name})
                return self.handle_command({"cmd": "end-wireguard"})
            elif cmd == "reinstall-wire-tunnel":
                config_path = cmd_inf.get("config_path", None)
                self.handle_command({"cmd": "unbrick-adapter"})
                if not config_path: return {"suc": False, "res": "No config path provided"}
                add = windows_service.reinstall_wireguard_service("NordWireConnect", config_path)
                self.connected_tunnel = "NordWireConnect"
                wait_for_service("WireGuardTunnel$NordWireConnect")
                return {"suc": add.returncode == 0, "res": add.returncode}
            elif cmd == "combined-end-wireguard":
                self.handle_command({"cmd": "end-wireguard"})
                return self.handle_command({"cmd": "end-wireguard-installer"})
            elif cmd == "data-usage":
                if self.connected_tunnel:
                    wg_check = windows_service.run_wireguard_command(["show"])
                    output = wg_check.stdout.strip()
                    lines = output.splitlines()
                    for l in lines:
                        l = l.strip()
                        if "transfer:" in l: 
                            l = l.replace("transfer: ", "").replace(" received", "").replace(" sent", "").replace(", ", ",").split(",")
                            if len(l) == 2:
                                return {"suc": True, "download": l[0], "upload": l[1]}
                return {"suc": False, "upload": "0", "download": "0"}
            elif cmd == "unbrick-adapter":
                get_interfaces_req = shell_run("netsh interface ipv4 show route").stdout + shell_run("netsh interface ipv6 show route").stdout
                interface_indexes = {
                    line.split()[-2] for line in get_interfaces_req.splitlines()
                    if "0.0.0.0/0" in line or "::/0" in line
                }
                for idx in interface_indexes:
                    if not idx.isdigit(): continue
                    if not check_for_unbricks(idx): continue
                    shell_run(f"netsh interface ipv4 set interface {idx} forwarding=disabled weakhostsend=disabled weakhostreceive=disabled")
                    shell_run(f"netsh interface ipv6 set interface {idx} forwarding=disabled weakhostsend=disabled weakhostreceive=disabled")
                return {"suc": True}
            elif cmd == "wireguard-check":
                wg_check = windows_service.run_wireguard_command(["show"])
                output = wg_check.stdout.strip()
                if not output: return {"suc": True, "active": False}
                elif "latest handshake: never" in output.lower(): return {"suc": True, "active": False}
                elif "latest handshake" in output.lower(): return {"suc": True, "active": True}
                else: return {"suc": True, "active": False}
            elif cmd == "connection-status":
                if self.connected_tunnel: return {"suc": True, "active": True, "tunnel": self.connected_tunnel}
                else: return {"suc": True, "active": False, "tunnel": "NotConnected"}
        except Exception as e: error(f"Failed running pipe command ({cmd}): {repr(e)}")
        return {"suc": False, "res": "1"}
    def ensure_ui_running(self):
        try:
            session_id = win32ts.WTSGetActiveConsoleSessionId()
            if session_id == 0xFFFFFFFF:
                self.ui_running = False
                self.last_session_id = None
                return
            if self.prevent_opening == True: return
            if not pip_class.getIfProcessIsOpened("explorer.exe"): return
            if self.ui_running and session_id == self.last_session_id and pip_class.getIfProcessIsOpened("NordWireConnect.exe"): return
            user_token = win32ts.WTSQueryUserToken(session_id)
            env = win32profile.CreateEnvironmentBlock(user_token, False)
            startup = win32process.STARTUPINFO()
            startup.lpDesktop = "winsta0\\default"
            win32process.CreateProcessAsUser(
                user_token,
                None,
                f'"{nordwireconnect_location}"',
                None,
                None,
                False,
                win32con.CREATE_NEW_CONSOLE | win32con.CREATE_UNICODE_ENVIRONMENT,
                env,
                None,
                startup
            )
            self.ui_running = True
            self.last_session_id = session_id
            info("NordWireConnect UI launched")
        except Exception as e:
            error(f"UI launch failed: {repr(e)}")
            self.ui_running = False
            self.last_session_id = None
            time.sleep(5)
    
# Service Runtime
if __name__ == "__main__":
    colors_class.fix_windows_ansi()
    setup_logging()
    systemMessage(f"{'-'*5:^5} NordWireConnectService v{version} {'-'*5:^5}")
    servicemanager.Initialize()
    servicemanager.PrepareToHostSingle(NordWireConnectService)
    mainMessage("Starting service..")
    servicemanager.StartServiceCtrlDispatcher()