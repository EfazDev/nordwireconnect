"""
NordWireConnect

Made by @EfazDev
https://www.efaz.dev
"""

# Modules
import os
import time
import win32ts # type: ignore
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
nordwireconnect_location = os.path.join(program_files, "Main", "NordWireConnect.exe")
wireguard_location = os.path.join(os.getenv("ProgramFiles"), "WireGuard")

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
        self.cleared_older = False
    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)
    def SvcDoRun(self):
        servicemanager.LogInfoMsg("NordWireConnect service starting")
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        threading.Thread(target=self.pipe_server, daemon=True).start()
        while True:
            if win32event.WaitForSingleObject(self.stop_event, 2000) == win32event.WAIT_OBJECT_0: break
            self.ensure_ui_running()
    def SvcOtherEx(self, control: int, event_type: int, data):
        if control == win32service.SERVICE_CONTROL_POWEREVENT:
            if event_type == win32con.PBT_APMSUSPEND:
                servicemanager.LogInfoMsg("System suspending")
                self.ui_running = False
                self.last_session_id = None
            elif event_type in (win32con.PBT_APMRESUMEAUTOMATIC, win32con.PBT_APMRESUMECRITICAL):
                servicemanager.LogInfoMsg("System resumed")
                time.sleep(5)
    def handle_command(self, command: str) -> str:
        servicemanager.LogInfoMsg(f"Received command: {command}")
        try:
            if command == "end-wireguard-tunnels":
                kill = shell_run("powershell -Command \"Get-Service WireGuardTunnel* | Stop-Service -Force\"")
                self.connected_tunnel = None
                return str(kill.returncode)
            elif command == "end-wireguard":
                end = shell_run("taskkill /IM wireguard.exe /F")
                self.connected_tunnel = None
                return str(end.returncode)
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
                self.handle_command("end-wireguard-tunnels")
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
            elif command.startswith("generate-preshared"):
                command = command.split(" ")
                add = subprocess.run([os.path.join(wireguard_location, "wg.exe"), "genpsk"], text=True, capture_output=True)
                return add.stdout.strip()
            elif command.startswith("generate-public-key"):
                command = command.split(" ")
                private_key = " ".join(command[1:])
                add = subprocess.run([os.path.join(wireguard_location, "wg.exe"), "pubkey"], text=True, capture_output=True, input=private_key)
                return add.stdout.strip()
            elif command.startswith("unbrick-adapter"):
                get_interfaces_req = subprocess.run(
                    "powershell -NoProfile -Command \"Get-NetRoute -DestinationPrefix 0.0.0.0/0 | Select-Object -ExpandProperty InterfaceIndex\"",
                    shell=True,
                    capture_output=True,
                    text=True
                )
                interface_indexes = {
                    int(line.strip())
                    for line in get_interfaces_req.stdout.splitlines()
                    if line.strip().isdigit()
                }
                for idx in interface_indexes:
                    # Should Unbrick?
                    if not check_for_unbricks(idx): continue
                    
                    # Unbrick
                    shell_run(f"netsh interface ipv4 set interface {idx} forwarding=disabled weakhostsend=disabled weakhostreceive=disabled")
                    shell_run(f"netsh interface ipv6 set interface {idx} forwarding=disabled weakhostsend=disabled weakhostreceive=disabled")
                return "0"
            elif command.startswith("block-public-ipv6"):
                cmds = [
                    'netsh advfirewall firewall add rule name="NordWireConnect Allow IPv6 LinkLocal" dir=out action=allow protocol=IPv6 remoteip=fe80::/10',
                    'netsh advfirewall firewall add rule name="NordWireConnect Allow IPv6 ULA" dir=out action=allow protocol=IPv6 remoteip=fd00::/8',
                    'netsh advfirewall firewall add rule name="NordWireConnect Block IPv6 Public" dir=out action=block protocol=IPv6 remoteip=2000::/3'
                ]
                for cmd in cmds: add = shell_run(cmd)
                return "0"
            elif command.startswith("unlock-public-ipv6"):
                rules = [
                    "NordWireConnect Allow IPv6 LinkLocal",
                    "NordWireConnect Allow IPv6 ULA",
                    "NordWireConnect Block IPv6 Public"
                ]
                for n in rules:
                    subprocess.run(
                        ["cmd.exe", "/c", f'netsh advfirewall firewall delete rule name="{n}"'],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                        check=False
                    )
                return "0"
            elif command.startswith("cleared-older-version"):
                if self.cleared_older: return "0"
                else: return "1"
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
            try:
                if os.path.exists(os.path.join(program_files, "NordWireConnect.exe")):
                    os.remove(os.path.join(program_files, "NordWireConnect.exe"))
                    self.cleared_older = True
                if os.path.exists(os.path.join(program_files, "NordWireConnectService.exe")):
                    os.remove(os.path.join(program_files, "NordWireConnectService.exe"))
                    self.cleared_older = True
            except Exception as e: servicemanager.LogInfoMsg("Clearing Older Versions of NordWireConnect..")
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