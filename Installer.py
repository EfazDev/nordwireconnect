"""
NordWireConnect

Made by @EfazDev
https://www.efaz.dev
"""

# Modules
import os
import sys
import time
import shutil
import typing
import ctypes
import PyKits
import logging
import datetime
import platform
import win32api # type: ignore
import win32con # type: ignore
import threading
import subprocess

# Variables
pip_class = PyKits.pip()
terminal = PyKits.Terminal()
main_os = platform.system()
pre_program_files = os.getenv("ProgramFiles") or r"C:\Program Files"
app_data_path = os.path.join(pip_class.getLocalAppData(), "NordWireConnect")
program_files = os.path.join(pre_program_files, "NordWireConnect")
cur_path = os.path.dirname(os.path.abspath(__file__))
version = "1.3.5p"

# Logging and Messages
def systemMessage(message: str): terminal.print_ansi(message, PyKits.Terminal.hex_to_ansi("#3E5FFF"))
def mainMessage(message: str): terminal.print_ansi(message, 15)
def errorMessage(message: str, title: str="Uh oh!"): 
    terminal.print_ansi(message, 9)
    if title == "": title = "NordWireConnect"
    elif title != "NordWireConnect": title = f"NordWireConnect: {title}"
    def a(): ctypes.windll.user32.MessageBoxW(0, message, title, 0x0 | 0x10)
    threading.Thread(target=a, daemon=True).start()
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
    handler_name = "Installer"
    log_path = os.path.join(app_data_path, "Logs")
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

# Installer
def get_if_compiled() -> bool: return bool(getattr(sys, "frozen", False) or "__compiled__" in globals())
def get_outer_executable() -> str: return os.path.abspath(sys.argv[0]) if get_if_compiled() else os.path.abspath(sys.executable)
def relaunch_as_admin():
    is_comp = get_if_compiled()
    target_exe = get_outer_executable()
    args = sys.argv[1:] if is_comp else sys.argv
    params = " ".join(f"\"{arg}\"" for arg in args)
    ctypes.windll.shell32.ShellExecuteW(
        None,
        "runas",
        target_exe,
        params,
        None,
        1
    )
    sys.exit(0)
def install():
    outer_exe = get_outer_executable()
    if not outer_exe.lower().startswith(pre_program_files.lower()):
        if not pip_class.getIfRunningWindowsAdmin():
            mainMessage("Please relaunch this program as Admin to continue installation!")
            relaunch_as_admin()
            return
        
        # Unzip Install Files
        mainMessage("Exporting Installation Files..")
        bundle_dir = os.path.dirname(os.path.abspath(__file__))
        zip_file = os.path.join(bundle_dir, "NordWireConnect.zip")
        staging_dir = os.path.join(app_data_path, "StagingInstaller")
        if os.path.exists(staging_dir): shutil.rmtree(staging_dir, ignore_errors=True)
        res = pip_class.unzipFile(zip_file, staging_dir, ["NordWireConnect.exe", "NordWireConnectService.exe"])
        if res.returncode != 0:
            errorMessage("Unable to unzip NordWireConnect installation zip.")
            return
        
        # Check for Installation and Prepare
        mainMessage("Preparing for installation..")
        os.makedirs(program_files, exist_ok=True)
        no_win = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        subprocess.run(["sc", "stop", "NordWireConnectService"], capture_output=True, creationflags=no_win)
        subprocess.run(["taskkill", "/IM", "NordWireConnectService.exe", "/F"], capture_output=True, creationflags=no_win)
        subprocess.run(["taskkill", "/IM", "NordWireConnect.exe", "/F"], capture_output=True, creationflags=no_win)
        time.sleep(3)
        for entry in os.listdir(program_files):
            if entry.lower() == "logs": continue
            entry_path = os.path.join(program_files, entry)
            try:
                if os.path.isdir(entry_path): shutil.rmtree(entry_path, ignore_errors=True)
                else: os.remove(entry_path)
            except Exception: pass
        
        # Copy Install Files
        mainMessage("Copying Installation Files..")
        shutil.copytree(os.path.join(bundle_dir, "Resources"), os.path.join(program_files, "Resources"), dirs_exist_ok=True) 
        shutil.copytree(staging_dir, program_files, dirs_exist_ok=True)
        shutil.rmtree(staging_dir, ignore_errors=True)

        # Create Service
        mainMessage("Creating NordWireConnect Service..")
        subprocess.run(["sc", "delete", "NordWireConnectService"], capture_output=True, check=False)
        subprocess.run(
            f"sc create NordWireConnectService binPath= \"{os.path.join(program_files, 'NordWireConnectService.exe')}\" start=auto",
            shell=True,
            check=True
        )

        # Start Service
        mainMessage("Starting NordWireConnect Service..")
        subprocess.run(["sc", "start", "NordWireConnectService"], capture_output=True)

        # Finalize and Setup Registry
        mainMessage("Setting up registry..")
        setup_registry()
        successMessage("DONE! This window will close in 3 seconds..")
        time.sleep(3)
        sys.exit(0)
def setup_registry():
    try:
        mainMessage("Marking Program Installation into Windows..")
        app_reg_path = "Software\\NordWireConnect"
        app_key = win32api.RegCreateKey(win32con.HKEY_LOCAL_MACHINE, app_reg_path)
        win32api.RegSetValueEx(app_key, "InstallPath", 0, win32con.REG_SZ, program_files)
        win32api.RegSetValueEx(app_key, "Installed", 0, win32con.REG_DWORD, 1)
        win32api.RegCloseKey(app_key)
        registry_path = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\NordWireConnect"
        registry_key = win32api.RegCreateKey(win32con.HKEY_LOCAL_MACHINE, registry_path)
        win32api.RegSetValueEx(registry_key, "UninstallString", 0, win32con.REG_SZ, f"\"{os.path.join(program_files, 'NordWireConnect.exe')}\" -uninstall-nord-wire-connect")
        win32api.RegSetValueEx(registry_key, "DisplayName", 0, win32con.REG_SZ, "NordWireConnect")
        win32api.RegSetValueEx(registry_key, "DisplayVersion", 0, win32con.REG_SZ, version)
        win32api.RegSetValueEx(registry_key, "DisplayIcon", 0, win32con.REG_SZ, os.path.join(program_files, "Resources", "app_icon.ico"))
        win32api.RegSetValueEx(registry_key, "InstallLocation", 0, win32con.REG_SZ, program_files)
        win32api.RegSetValueEx(registry_key, "Publisher", 0, win32con.REG_SZ, "EfazDev")
        win32api.RegSetValueEx(registry_key, "EstimatedSize", 0, win32con.REG_DWORD, min(get_folder_size(program_files, formatWithAbbreviation=False) // 1024, 0xFFFFFFFF))
        win32api.RegCloseKey(registry_key)
    except Exception as e: errorMessage(f"Unable to setup registry: {str(e)}")
def format_size(size_bytes: int) -> str:
    if size_bytes == 0: return "0 Bytes"
    size_units = ["Bytes", "KB", "MB", "GB", "TB"]
    unit_index = 0
    while size_bytes >= 1024 and unit_index < len(size_units) - 1: size_bytes /= 1024; unit_index += 1
    return f"{size_bytes:.2f} {size_units[unit_index]}"
def get_folder_size(folder_path: str, formatWithAbbreviation: bool=True) -> typing.Union[str, int]:
    total_size = 0
    stack = [folder_path]
    while stack:
        current = stack.pop()
        try:
            with os.scandir(current) as it:
                for entry in it:
                    try:
                        if entry.is_file(follow_symlinks=False): total_size += entry.stat(follow_symlinks=False).st_size
                        elif entry.is_dir(follow_symlinks=False): stack.append(entry.path)
                    except Exception: pass
        except Exception: pass
    if formatWithAbbreviation == True: return format_size(total_size)
    else: return total_size

# Main Runtime
def app():
    global cur_path
    os.makedirs(app_data_path, exist_ok=True)
    setup_logging()
    systemMessage(f"{'-'*5:^5} NordWireConnect v{version} {'-'*5:^5}")

    # Only for Windows
    if main_os != "Windows":
        errorMessage("NordWireConnect is only supported on Windows.")
        sys.exit(1)
        return
    
    # Prepare NordWireConnect
    cur_path = os.path.dirname(get_outer_executable())

    try: install()
    except Exception as e: errorMessage(f"There was an error during Installation Handler: {str(e)}"); sys.exit(1); return
if __name__ == "__main__": app()