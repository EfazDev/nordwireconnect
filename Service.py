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
import logging
import win32ts # type: ignore
import datetime
import win32con # type: ignore
import win32file # type: ignore
import win32pipe # type: ignore
import threading # type: ignore
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
version = "1.3.0e"
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
def shell_run(cmd: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        cmd,
        shell=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )
def check_for_unbricks(idx: str) -> bool:
    out = subprocess.run(f"netsh interface ipv4 show interface {idx}", shell=True, capture_output=True, text=True).stdout.lower()
    for ln in out.splitlines():
        if "weak host sends" in ln and "enabled" in ln: return True
        if "weak host receives" in ln and "enabled" in ln: return True
        if "forwarding" in ln and "enabled" in ln: return True
    out = subprocess.run(f"netsh interface ipv6 show interface {idx}", shell=True, capture_output=True, text=True).stdout.lower()
    for ln in out.splitlines():
        if "weak host sends" in ln and "enabled" in ln: return True
        if "weak host receives" in ln and "enabled" in ln: return True
        if "forwarding" in ln and "enabled" in ln: return True
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
        threading.Thread(target=self.pipe_server, daemon=True).start()
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
    def handle_command(self, command: str) -> str:
        try:
            if command != "wireguard-check": info(f"Received command: {command}")
            if command == "end-wireguard":
                end = shell_run("taskkill /IM wireguard.exe /F")
                time.sleep(0.5)
                scm = win32service.OpenSCManager(None, None, win32service.SC_MANAGER_ENUMERATE_SERVICE)
                statuses = win32service.EnumServicesStatus(scm)
                for (name, display, status) in statuses:
                    if name.startswith("WireGuardTunnel$"): win32serviceutil.RemoveService(name)
                self.connected_tunnel = None
                return str(end.returncode)
            elif command == "ui-opening": 
                self.prevent_opening = False
                return "0"
            elif command == "ui-closing": 
                self.prevent_opening = True
                return "0"
            elif command == "service-version": return version
            elif command == "end-wireguard-installer":
                end = shell_run("taskkill /IM wireguard-installer.exe /F")
                self.connected_tunnel = None
                return str(end.returncode)
            elif command.startswith("uninstall-wire-tunnel"):
                command = command.split(" ")
                remove = subprocess.run([os.path.join(wireguard_location, "wireguard.exe"), "/uninstalltunnelservice", command[1]], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                if self.connected_tunnel == command[1]: self.connected_tunnel = None
                return str(remove.returncode)
            elif command.startswith("install-wire-tunnel"):
                self.handle_command("unbrick-adapter")
                command = command.split(" ")
                config_path = " ".join(command[1:])
                add = subprocess.run([os.path.join(wireguard_location, "wireguard.exe"), "/installtunnelservice", config_path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                self.connected_tunnel = ".".join(os.path.basename(config_path).split(".")[:-1])
                return str(add.returncode)
            elif command.startswith("combined-disconnect"):
                command = command.split(" ")
                if len(command) > 1:
                    server_name = " ".join(command[1:])
                    self.handle_command(f"uninstall-wire-tunnel {server_name}")
                return self.handle_command("end-wireguard")
            elif command.startswith("reinstall-wire-tunnel"):
                command = command.split(" ")
                shortened_name = command[1]
                config_path = " ".join(command[2:])
                self.handle_command(f"uninstall-wire-tunnel {shortened_name}")
                return self.handle_command(f"install-wire-tunnel {config_path}")
            elif command.startswith("combined-end-wireguard"):
                self.handle_command("end-wireguard")
                return self.handle_command("end-wireguard-installer")
            elif command.startswith("data-usage"):
                if self.connected_tunnel:
                    wg_check = subprocess.run([os.path.join(wireguard_location, "wg.exe"), "show"], capture_output=True, text=True)
                    output = wg_check.stdout.strip()
                    lines = output.splitlines()
                    for l in lines:
                        l = l.strip()
                        if "transfer:" in l: return l.replace("transfer: ", "").replace(" received", "").replace(" sent", "").replace(", ", ",")
                return "0,0"
            elif command.startswith("router-ip-info"):
                route_out = subprocess.run("route print 0.0.0.0", shell=True, capture_output=True, text=True).stdout.replace("[SEPARATE]", "") + "\n[SEPARATE]\n" + subprocess.run("route print ::/0", shell=True, capture_output=True, text=True).stdout.replace("[SEPARATE]", "")
                parsing_mode = None
                active = False
                ipv4 = None
                ipv6 = None
                for l in route_out.splitlines():
                    if "IPv4 Route Table" in l: parsing_mode = 1
                    elif "IPv6 Route Table" in l: parsing_mode = 2
                    elif "[SEPARATE]" in l: parsing_mode = None
                    elif "Active Routes:" in l: active = True
                    elif "Persistent Routes:" in l: active = False
                    elif parsing_mode and active:
                        if "0.0.0.0" in l and parsing_mode == 1:
                            for p in l.split(" "):
                                p = p.strip()
                                if not p: continue
                                if p == "0.0.0.0": continue
                                if p.count(".") == 3:
                                    ipv4 = p
                                    parsing_mode = None
                                    active = False
                        elif "::/0" in l and parsing_mode == 2:
                            for p in l.split(" "):
                                p = p.strip()
                                if not p: continue
                                if p == "::/0": continue
                                if ":" in p:
                                    ipv6 = p
                                    parsing_mode = None
                                    active = False
                return f"{ipv4},{ipv6}"
            elif command.startswith("unbrick-adapter"):
                get_interfaces_req = subprocess.run(
                    "netsh interface ipv4 show route",
                    shell=True,
                    capture_output=True,
                    text=True
                ).stdout + subprocess.run(
                    "netsh interface ipv6 show route",
                    shell=True,
                    capture_output=True,
                    text=True
                ).stdout
                interface_indexes = {
                    line.split()[-2] for line in get_interfaces_req.splitlines()
                    if "0.0.0.0/0" in line or "::/0" in line
                }
                for idx in interface_indexes:
                    if not idx.isdigit(): continue
                    if not check_for_unbricks(idx): continue
                    shell_run(f"netsh interface ipv4 set interface {idx} forwarding=disabled weakhostsend=disabled weakhostreceive=disabled")
                    shell_run(f"netsh interface ipv6 set interface {idx} forwarding=disabled weakhostsend=disabled weakhostreceive=disabled")
                return "0"
            elif command == "wireguard-check":
                wg_check = subprocess.run([os.path.join(wireguard_location, "wg.exe"), "show"], capture_output=True, text=True)
                output = wg_check.stdout.strip()
                if not output: return "1"
                elif "latest handshake: never" in output.lower(): return "1"
                elif "latest handshake" in output.lower(): return "0"
                else: return "1"
            elif command == "connection-status":
                if self.connected_tunnel: return self.connected_tunnel
                else: return "NotConnected"
        except Exception as e: error(f"Failed running pipe command ({command}): {repr(e)}")
        return "1"
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
    def pipe_server(self):
        info("Pipe server starting")
        while win32event.WaitForSingleObject(self.stop_event, 0) != win32event.WAIT_OBJECT_0:
            pipe = None
            try:
                sa = create_pipe_sa()
                pipe = win32pipe.CreateNamedPipe(
                    service_pipe,
                    win32pipe.PIPE_ACCESS_DUPLEX,
                    win32pipe.PIPE_TYPE_MESSAGE |
                    win32pipe.PIPE_READMODE_MESSAGE |
                    win32pipe.PIPE_WAIT,
                    win32pipe.PIPE_UNLIMITED_INSTANCES,
                    4096,
                    4096,
                    0,
                    sa
                )
                win32pipe.ConnectNamedPipe(pipe, None)
                while True:
                    try: _, data = win32file.ReadFile(pipe, 4096)
                    except Exception: break
                    command = data.decode().strip()
                    response = self.handle_command(command)
                    win32file.WriteFile(pipe, response.encode())
            except Exception as e: error(f"Pipe error: {repr(e)}")
            finally:
                if pipe:
                    try: win32file.CloseHandle(pipe)
                    except Exception: pass

# Service Runtime
if __name__ == "__main__":
    colors_class.fix_windows_ansi()
    setup_logging()
    systemMessage(f"{'-'*5:^5} NordWireConnectService v{version} {'-'*5:^5}")
    servicemanager.Initialize()
    servicemanager.PrepareToHostSingle(NordWireConnectService)
    mainMessage("Starting service..")
    servicemanager.StartServiceCtrlDispatcher()