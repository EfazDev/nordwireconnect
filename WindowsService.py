"""
NordWireConnect

Made by @EfazDev
https://www.efaz.dev
"""

# Modules
import os
import re
import sys
import time
import PyKits
import System
import logging
import win32ts # type: ignore
import win32api # type: ignore
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
program_files = os.path.join(os.getenv("ProgramFiles"), "NordWireConnect")
nordwireconnect_location = os.path.join(program_files, "NordWireConnect.exe")
wireguard_location = os.path.join(os.getenv("ProgramFiles"), "WireGuard")
version = "1.3.5m"
windows_service = System.WindowsInterface()
nord_socket = PyKits.Socket(port=61245)
terminal = PyKits.Terminal()
pip_class = PyKits.pip()
safe_tunnel_re = re.compile(r"^[A-Za-z0-9_.\-]{1,64}$")
quiet_commands = {"wireguard-check", "tunnel-stats", "data-usage", "connection-status"}

# Logging
def info(message: str): servicemanager.LogInfoMsg(message); mainMessage(message)
def error(message: str): servicemanager.LogErrorMsg(message); errorMessage(message)
def warn(message: str): servicemanager.LogWarningMsg(message); warnMessage(message)
def systemMessage(message: str): terminal.print_ansi(message, PyKits.Terminal.hex_to_ansi("#3E5FFF"))
def mainMessage(message: str): terminal.print_ansi(message, 15)
def errorMessage(message: str, title: str="Uh oh!"): terminal.print_ansi(message, 9)
def warnMessage(message: str): terminal.print_ansi(message, 11)
def successMessage(message: str): terminal.print_ansi(message, 10)
def prune_old_logs(log_path: str, prefix: str, keep: int=15):
    try:
        entries = [os.path.join(log_path, f) for f in os.listdir(log_path) if f.startswith(prefix) and f.endswith(".log")]
        if len(entries) > keep:
            entries.sort(key=lambda p: os.path.getmtime(p), reverse=True)
            for stale in entries[keep:]:
                try: os.remove(stale)
                except Exception: pass
    except Exception: pass
def setup_logging():
    global terminal
    handler_name = "Service"
    log_path = os.path.join(program_files, "Logs")
    if not os.path.exists(log_path): os.makedirs(log_path, mode=511)
    prune_old_logs(log_path, f"NordWireConnect_{handler_name}_", keep=15)
    generated_file_name = f"NordWireConnect_{handler_name}_{datetime.datetime.now().strftime('%B_%d_%Y_%H_%M_%S_%f')}.log"
    if hasattr(sys.stdout, "reconfigure"):
        try: sys.stdout.reconfigure(encoding="utf-8")
        except Exception: pass
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    file_handler = logging.FileHandler(os.path.join(log_path, generated_file_name), encoding="utf-8")
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    logger.addHandler(file_handler)
    terminal = PyKits.Terminal(logger=logger, log_level=logging.INFO)
    err_terminal = PyKits.Terminal(logger=logger, log_level=logging.ERROR)
    terminal.inject_stdout()
    err_terminal.inject_stderr()
    System.systemMessage = systemMessage
    System.mainMessage = mainMessage
    System.successMessage = successMessage
    System.warnMessage = warnMessage
    System.errorMessage = errorMessage

# Unbricking
def shell_run(cmd: str | list) -> subprocess.CompletedProcess[str]: 
    s = System.shell_run(cmd)
    if s.returncode != 0: error(f"Command failed ({cmd}): {s.stderr.strip()}")
    return s
def format_wg_bytes(n: int) -> str:
    val = float(n)
    for unit in ("B", "KiB", "MiB", "GiB", "TiB", "PiB"):
        if abs(val) < 1024.0: return f"{val:.2f} {unit}"
        val /= 1024.0
    return f"{val:.2f} PiB"
def sanitize_tunnel_name(name: str) -> str:
    if not name or not isinstance(name, str): return "NordWireConnect"
    name = name.strip()
    if safe_tunnel_re.match(name): return name
    return "NordWireConnect"
def get_tunnel_stats(tunnel_name: str="NordWireConnect") -> dict:
    svc_state = windows_service.get_service_state(f"WireGuardTunnel${tunnel_name}")
    service_running = svc_state == win32service.SERVICE_RUNNING
    if not service_running: return {"suc": True, "service_running": False, "active": False, "has_handshake": False, "handshake_age": None, "rx_bytes": 0, "tx_bytes": 0}
    try:
        wg_dump = windows_service.run_wireguard_command(["show", tunnel_name, "dump"])
        if wg_dump.returncode == 0 and wg_dump.stdout.strip():
            lines = wg_dump.stdout.strip().splitlines()
            if len(lines) >= 2:
                peer_parts = lines[1].split("\t")
                if len(peer_parts) >= 7:
                    latest_hs = int(peer_parts[4])
                    rx_bytes = int(peer_parts[5])
                    tx_bytes = int(peer_parts[6])
                    hs_age = max(0.0, time.time() - latest_hs) if latest_hs > 0 else None
                    has_hs = latest_hs > 0
                    active = has_hs and (hs_age is not None and hs_age < 210.0)
                    return {"suc": True, "service_running": True, "active": active, "has_handshake": has_hs, "handshake_age": round(hs_age, 1) if hs_age is not None else None, "rx_bytes": rx_bytes, "tx_bytes": tx_bytes}
    except Exception: pass
    wg_check = windows_service.run_wireguard_command(["show", tunnel_name])
    output = wg_check.stdout.strip().lower()
    has_hs = ("latest handshake" in output) and ("latest handshake: never" not in output)
    return {"suc": True, "service_running": True, "active": has_hs, "has_handshake": has_hs, "handshake_age": 0.0 if has_hs else None, "rx_bytes": 0, "tx_bytes": 0}
def extract_default_route_indexes(route_output: str) -> set[str]:
    indexes = set()
    for line in route_output.splitlines():
        lower_line = line.lower()
        if "nordwireconnect" in lower_line or "wireguard" in lower_line: continue
        if "0.0.0.0/0" not in line and "::/0" not in line: continue
        parts = line.split()
        for i, tok in enumerate(parts):
            if tok in ("0.0.0.0/0", "::/0") and i + 1 < len(parts):
                candidate = parts[i + 1]
                if candidate.isdigit(): indexes.add(candidate)
                break
    return indexes
def check_for_unbricks(idx: str) -> tuple[bool, bool]:
    fix_v4 = False
    fix_v6 = False
    out4 = System.shell_run(["netsh", "interface", "ipv4", "show", "interface", idx]).stdout.lower()
    for ln in out4.splitlines():
        if ("weak host sends" in ln or "weak host receives" in ln or "forwarding" in ln) and "enabled" in ln:
            fix_v4 = True
            break
    out6 = System.shell_run(["netsh", "interface", "ipv6", "show", "interface", idx]).stdout.lower()
    for ln in out6.splitlines():
        if ("weak host sends" in ln or "weak host receives" in ln or "forwarding" in ln) and "enabled" in ln:
            fix_v6 = True
            break
    return fix_v4, fix_v6
def wait_for_service(name: str, timeout: int=15, status: int=win32service.SERVICE_RUNNING) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if windows_service.get_service_state(name) == status: return True
        time.sleep(0.2)
    return windows_service.get_service_state(name) == status

# Service Class
class NordWireConnectService(win32serviceutil.ServiceFramework):
    _svc_name_ = "NordWireConnectService"
    _svc_display_name_ = "NordWireConnectService"
    _svc_description_ = "Handles NordWireConnect with UI Handling, Background Tasks and WireGuard Management!"
    _svc_controls_accepted_ = win32service.SERVICE_ACCEPT_STOP | win32service.SERVICE_ACCEPT_POWEREVENT | win32service.SERVICE_ACCEPT_SESSIONCHANGE
    def __init__(self, args):
        super().__init__(args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.ui_running = False
        self.ui_process_handle = None
        self.last_session_id = None
        self.connected_tunnel = None
        self.prevent_opening = False
        self.last_unbrick_time = 0.0
        self.last_resume_time = 0.0
        self.resume_ui_after = 0.0
    def close_ui_handle(self):
        if self.ui_process_handle:
            try: win32api.CloseHandle(self.ui_process_handle)
            except Exception: pass
            self.ui_process_handle = None
    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.close_ui_handle()
        win32event.SetEvent(self.stop_event)
    def SvcDoRun(self):
        info("NordWireConnect service starting")
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        if nord_socket.exists(): nord_socket.wait_till_free(timeout=15)
        nord_socket.subscribe("NordWireConnectCommands", self.handle_command)
        nord_socket.listen()
        while True:
            if win32event.WaitForSingleObject(self.stop_event, 2000) == win32event.WAIT_OBJECT_0: break
            self.ensure_ui_running()
        nord_socket.close()
    def SvcOtherEx(self, control: int, event_type: int, data):
        if control == win32service.SERVICE_CONTROL_POWEREVENT:
            if event_type == win32con.PBT_APMSUSPEND:
                info("System suspending")
                self.ui_running = False
            elif event_type in (win32con.PBT_APMRESUMEAUTOMATIC, win32con.PBT_APMRESUMECRITICAL):
                info("System resumed")
                now = time.monotonic()
                self.last_resume_time = now
                self.resume_ui_after = now + 5.0
                self.ui_running = False
    def handle_command(self, cmd_inf: dict) -> dict:
        cmd = "<unknown>"
        try:
            cmd = cmd_inf.get("cmd", "")
            if cmd not in quiet_commands: info(f"Received command: {cmd}")
            if cmd == "end-wireguard": return windows_service._end_wireguard(self)
            elif cmd == "ui-opening": self.prevent_opening = False; return {"suc": True}
            elif cmd == "ui-closing": self.prevent_opening = True; return {"suc": True}
            elif cmd == "service-version": return {"suc": True, "version": version}
            elif cmd == "end-wireguard-installer": return windows_service._end_wireguard_installer()
            elif cmd == "uninstall-wire-tunnel":
                service_name = sanitize_tunnel_name(cmd_inf.get("server") or "NordWireConnect")
                remove = windows_service.uninstall_wireguard_service(service_name)
                if self.connected_tunnel == service_name: self.connected_tunnel = None
                return {"suc": remove.returncode == 0, "res": remove.returncode}
            elif cmd == "install-wire-tunnel":
                self.handle_command({"cmd": "unbrick-adapter"})
                config_path = cmd_inf.get("config_path", None)
                if not config_path or not os.path.isfile(config_path): return {"suc": False, "res": "Invalid config path provided"}
                add = windows_service.install_wireguard_service(config_path)
                if add.returncode != 0:
                    err_detail = (add.stderr or add.stdout or "").strip()
                    error(f"WireGuard installtunnelservice failed (code {add.returncode}): {err_detail}")
                    return {"suc": False, "res": add.returncode, "err": err_detail}
                self.connected_tunnel = "NordWireConnect"
                started = wait_for_service("WireGuardTunnel$NordWireConnect", timeout=15)
                return {"suc": started, "res": add.returncode}
            elif cmd == "combined-disconnect":
                server_name = cmd_inf.get("known_server", None)
                if server_name: self.handle_command({"cmd": "uninstall-wire-tunnel", "server": server_name})
                self.handle_command({"cmd": "uninstall-wire-tunnel", "server": "NordWireConnect"})
                return self.handle_command({"cmd": "end-wireguard"})
            elif cmd == "reinstall-wire-tunnel":
                config_path = cmd_inf.get("config_path", None)
                self.handle_command({"cmd": "unbrick-adapter"})
                if not config_path or not os.path.isfile(config_path): return {"suc": False, "res": "Invalid config path provided"}
                add = windows_service.reinstall_wireguard_service("NordWireConnect", config_path)
                if add.returncode != 0: return {"suc": False, "res": add.returncode}
                self.connected_tunnel = "NordWireConnect"
                started = wait_for_service("WireGuardTunnel$NordWireConnect", timeout=15)
                return {"suc": started, "res": add.returncode}
            elif cmd == "combined-end-wireguard":
                self.handle_command({"cmd": "end-wireguard"})
                return self.handle_command({"cmd": "end-wireguard-installer"})
            elif cmd == "tunnel-stats":
                stats = get_tunnel_stats("NordWireConnect")
                stats["seconds_since_resume"] = round(max(0.0, time.monotonic() - self.last_resume_time), 1) if self.last_resume_time > 0 else None
                return stats
            elif cmd == "data-usage":
                stats = get_tunnel_stats("NordWireConnect")
                if stats.get("service_running"): return {"suc": True, "download": format_wg_bytes(stats["rx_bytes"]), "upload": format_wg_bytes(stats["tx_bytes"]), "rx_bytes": stats["rx_bytes"], "tx_bytes": stats["tx_bytes"]}
                return {"suc": False, "upload": "0 B", "download": "0 B", "rx_bytes": 0, "tx_bytes": 0}
            elif cmd == "unbrick-adapter":
                now = time.monotonic()
                if not cmd_inf.get("force") and (now - self.last_unbrick_time) < 60.0: return {"suc": True, "cached": True}
                r4 = System.shell_run(["netsh", "interface", "ipv4", "show", "route"]).stdout
                r6 = System.shell_run(["netsh", "interface", "ipv6", "show", "route"]).stdout
                interface_indexes = extract_default_route_indexes(r4 + "\n" + r6)
                for idx in interface_indexes:
                    fix_v4, fix_v6 = check_for_unbricks(idx)
                    if fix_v4: System.shell_run(["netsh", "interface", "ipv4", "set", "interface", idx, "forwarding=disabled", "weakhostsend=disabled", "weakhostreceive=disabled"])
                    if fix_v6: System.shell_run(["netsh", "interface", "ipv6", "set", "interface", idx, "forwarding=disabled", "weakhostsend=disabled", "weakhostreceive=disabled"])
                self.last_unbrick_time = now
                return {"suc": True}
            elif cmd == "wireguard-check":
                stats = get_tunnel_stats("NordWireConnect")
                return {"suc": True, "active": stats["active"], "service_running": stats["service_running"], "has_handshake": stats["has_handshake"], "handshake_age": stats["handshake_age"]}
            elif cmd == "connection-status":
                svc_running = windows_service.get_service_state("WireGuardTunnel$NordWireConnect") == win32service.SERVICE_RUNNING
                if svc_running:
                    self.connected_tunnel = "NordWireConnect"
                    return {"suc": True, "active": True, "tunnel": "NordWireConnect"}
                self.connected_tunnel = None
                return {"suc": True, "active": False, "tunnel": "NotConnected"}
        except Exception as e: error(f"Failed running command ({cmd}): {repr(e)}")
        return {"suc": False, "res": "1"}
    def ensure_ui_running(self):
        if time.monotonic() < self.resume_ui_after: return
        user_token = None
        env = None
        try:
            session_id = win32ts.WTSGetActiveConsoleSessionId()
            if session_id == 0xFFFFFFFF:
                self.ui_running = False
                self.last_session_id = None
                self.close_ui_handle()
                return
            if self.prevent_opening == True: return
            if self.ui_process_handle and session_id == self.last_session_id:
                if win32event.WaitForSingleObject(self.ui_process_handle, 0) == win32event.WAIT_TIMEOUT: return
                self.close_ui_handle()
            if pip_class.getIfProcessIsOpened("NordWireConnect.exe"):
                self.ui_running = True
                self.last_session_id = session_id
                return
            if not pip_class.getIfProcessIsOpened("explorer.exe"): return
            user_token = win32ts.WTSQueryUserToken(session_id)
            env = win32profile.CreateEnvironmentBlock(user_token, False)
            startup = win32process.STARTUPINFO()
            startup.lpDesktop = "winsta0\\default"
            h_proc, h_thread, _, _ = win32process.CreateProcessAsUser(
                user_token,
                None,
                f"\"{nordwireconnect_location}\"",
                None,
                None,
                False,
                win32con.CREATE_NEW_CONSOLE | win32con.CREATE_UNICODE_ENVIRONMENT,
                env,
                None,
                startup
            )
            try: win32api.CloseHandle(h_thread)
            except Exception: pass
            self.close_ui_handle()
            self.ui_process_handle = h_proc
            self.ui_running = True
            self.last_session_id = session_id
            info("NordWireConnect UI launched")
        except Exception as e:
            error(f"UI launch failed: {repr(e)}")
            self.ui_running = False
            self.last_session_id = None
            self.close_ui_handle()
            self.resume_ui_after = time.monotonic() + 5.0
        finally:
            if env:
                try: win32profile.DestroyEnvironmentBlock(env)
                except Exception: pass
            if user_token:
                try: win32api.CloseHandle(user_token)
                except Exception: pass

# Service Runtime
def main():
    setup_logging()
    systemMessage(f"{'-'*5:^5} NordWireConnectService v{version} {'-'*5:^5}")
    cli_args = [a for a in sys.argv[1:] if a != "--service"]
    if cli_args:
        win32serviceutil.HandleCommandLine(NordWireConnectService, argv=[sys.argv[0]] + cli_args)
        return
    try:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(NordWireConnectService)
        mainMessage("Starting service..")
        servicemanager.StartServiceCtrlDispatcher()
    except Exception as e: error(f"Service dispatcher failed to start: {repr(e)}"); raise
if __name__ == "__main__": main()