"""
NordWireConnect

Made by @EfazDev
https://www.efaz.dev
"""

# Modules
import win32serviceutil # type: ignore
import win32service # type: ignore
import win32event # type: ignore
import win32process # type: ignore
import win32profile # type: ignore
import win32security # type: ignore
import ntsecuritycon # type: ignore
import win32pipe # type: ignore
import win32file # type: ignore
import threading # type: ignore
import win32con # type: ignore
import win32ts # type: ignore
import subprocess
import servicemanager # type: ignore
import time
import sys
import os

# Variables
PIPE_NAME = r"\\.\pipe\NordWireConnect"
nordwireconnect_location = os.path.join(os.getenv("ProgramFiles"), "NordWireConnect", "Main", "NordWireConnect.exe")
wireguard_location = os.path.join(os.getenv("ProgramFiles"), "WireGuard")
if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'): cur_path = os.path.dirname(sys.executable)
else: cur_path = os.path.dirname(sys.argv[0])

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

# Service Class
class NordWireService(win32serviceutil.ServiceFramework):
    _svc_name_ = "NordWireConnectService"
    _svc_display_name_ = "NordWireConnectService"
    _svc_description_ = "Start NordWireConnect on boot and handle administrative commands"
    def __init__(self, args):
        super().__init__(args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.ui_running = False
        self.last_session_id = None
        self.connected_tunnel = None
        self.cached_routing = []
    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)
    def SvcDoRun(self):
        servicemanager.LogInfoMsg("NordWireConnect service starting")
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        threading.Thread(target=self.pipe_server).start()
        while True:
            if win32event.WaitForSingleObject(self.stop_event, 2000) == win32event.WAIT_OBJECT_0: break
            self.ensure_ui_running()
    def handle_command(self, command: str) -> str:
        servicemanager.LogInfoMsg(f"Received command: {command}")
        try:
            if command == "end-wireguard-tunnels":
                kill = subprocess.run("powershell -Command \"Get-Service WireGuardTunnel* | Stop-Service -Force\"", stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, shell=True)
                self.connected_tunnel = None
                return str(kill.returncode)
            elif command == "end-wireguard":
                end = subprocess.run("taskkill /IM wireguard.exe /F", stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, shell=True)
                self.connected_tunnel = None
                return str(end.returncode)
            elif command == "end-wireguard-installer":
                end = subprocess.run("taskkill /IM wireguard-installer.exe /F", stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, shell=True)
                self.connected_tunnel = None
                return str(end.returncode)
            elif command.startswith("uninstall-wire-tunnel"):
                command = command.split(" ")
                remove = subprocess.run([os.path.join(wireguard_location, "wireguard.exe"), "/uninstalltunnelservice", command[1]], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                if self.connected_tunnel == command[1]: self.connected_tunnel = None
                return str(remove.returncode)
            elif command.startswith("install-wire-tunnel"):
                command = command.split(" ")
                config_path = " ".join(command[1:])
                add = subprocess.run([os.path.join(wireguard_location, "wireguard.exe"), "/installtunnelservice", config_path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                self.connected_tunnel = ".".join(os.path.basename(config_path).split(".")[:-1])
                return str(add.returncode)
            elif command == "connection-status":
                if self.connected_tunnel: return self.connected_tunnel
                else: return "NotConnected"
        except Exception as e: servicemanager.LogErrorMsg(f"Failed running pipe command ({command}): {repr(e)}")
        return "1"
    def ensure_ui_running(self):
        try:
            session_id = win32ts.WTSGetActiveConsoleSessionId()
            if session_id == 0xFFFFFFFF:
                self.ui_running = False
                self.last_session_id = None
                return
            if self.ui_running and session_id == self.last_session_id: return
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
            servicemanager.LogInfoMsg("NordWireConnect UI launched")
        except Exception as e:
            servicemanager.LogErrorMsg(f"UI launch failed: {repr(e)}")
            self.ui_running = False
            self.last_session_id = None
    def pipe_server(self):
        servicemanager.LogInfoMsg("Pipe server starting")
        while True:
            pipe = None
            try:
                sa = create_pipe_sa()
                pipe = win32pipe.CreateNamedPipe(
                    PIPE_NAME,
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
                    except: break
                    command = data.decode().strip()
                    response = self.handle_command(command)
                    win32file.WriteFile(pipe, response.encode())
            except Exception as e: servicemanager.LogErrorMsg(f"Pipe error: {repr(e)}")
            finally:
                if pipe:
                    try: win32file.CloseHandle(pipe)
                    except: pass

# Service Runtime
if __name__ == "__main__":
    servicemanager.Initialize()
    servicemanager.PrepareToHostSingle(NordWireService)
    servicemanager.StartServiceCtrlDispatcher()