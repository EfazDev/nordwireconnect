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
colors_class = PyKits.Colors()
main_os = platform.system()
pre_program_files = os.getenv("ProgramFiles")
app_data_path = os.path.join(pip_class.getLocalAppData(), "NordWireConnect")
program_files = os.path.join(pre_program_files, "NordWireConnect")
cur_path = os.path.dirname(os.path.abspath(__file__))
version = "1.2.0"

# Logging and Messages
def systemMessage(message: str): colors_class.print(message, colors_class.hex_to_ansi2("#3E5FFF"))
def mainMessage(message: str): colors_class.print(message, 15)
def errorMessage(message: str, title: str="Uh oh!"): 
    colors_class.print(message, 9)
    if title == "": title = "NordWireConnect"
    elif title != "NordWireConnect": title = f"NordWireConnect: {title}"
    def a(): ctypes.windll.user32.MessageBoxW(0, message, title, 0x0 | 0x10)
    threading.Thread(target=a, daemon=True).start()
def warnMessage(message: str): colors_class.print(message, 11)
def successMessage(message: str): colors_class.print(message, 10)
def setup_logging():
    handler_name = "Installer"
    log_path = os.path.join(app_data_path, "Logs")
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

# Installer
def relaunch_as_admin():
    params = " ".join(f'"{arg}"' for arg in sys.argv)
    ctypes.windll.shell32.ShellExecuteW(
        None,
        "runas",
        sys.executable,
        params,
        None,
        1
    )
    sys.exit(0)
def install():
    if not sys.executable.startswith(pre_program_files):
        if not pip_class.getIfRunningWindowsAdmin():
            mainMessage("Please relaunch this program as Admin to continue installation!")
            relaunch_as_admin()
            return
        mainMessage("Exporting Files..")
        zip_file = os.path.join(os.path.dirname(__file__), "NordWireConnect.zip")
        res = pip_class.unzipFile(zip_file, os.path.join(os.path.dirname(__file__), "installer"), ["NordWireConnect", "NordWireConnectService"])
        if res.returncode != 0:
            errorMessage("Unable to unzip NordWireConnect installation zip.")
            return
        mainMessage("Installing NordWireConnect..")
        os.makedirs(program_files, exist_ok=True)
        if os.path.exists(os.path.join(cur_path, "NordSessionData.json")): shutil.copy(os.path.join(cur_path, "NordSessionData.json"), os.path.join(app_data_path, "NordSessionData.json"))
        if os.path.exists(os.path.join(cur_path, "ConnectConfig.json")): shutil.copy(os.path.join(cur_path, "ConnectConfig.json"), os.path.join(app_data_path, "ConnectConfig.json"))
        shutil.copytree(os.path.join(os.path.dirname(__file__), "resources"), os.path.join(program_files, "resources"), dirs_exist_ok=True) 
        subprocess.run("taskkill /IM NordWireConnect.exe /F", shell=True)
        time.sleep(3)
        shutil.copytree(os.path.join(os.path.dirname(__file__), "installer", "NordWireConnect"), os.path.join(program_files, "Main"), dirs_exist_ok=True)
        subprocess.run("sc stop NordWireConnectService", shell=True)
        subprocess.run("taskkill /IM NordWireConnectService.exe /F", shell=True)
        time.sleep(3)
        shutil.copytree(os.path.join(os.path.dirname(__file__), "installer", "NordWireConnectService"), os.path.join(program_files, "Service"), dirs_exist_ok=True)
        mainMessage("Starting NordWireConnect Service..")
        subprocess.run(
            "sc delete NordWireConnectService",
            shell=True,
            check=False
        )
        subprocess.run(
            f'sc create NordWireConnectService binPath= "{os.path.join(program_files, "Service", "NordWireConnectService.exe")}" start=auto',
            shell=True,
            check=True
        )
        subprocess.run("sc start NordWireConnectService", shell=True)
        mainMessage("Setting up registry..")
        setup_registry()
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
        win32api.RegSetValueEx(registry_key, "UninstallString", 0, win32con.REG_SZ, f"\"{os.path.join(program_files, 'Main', 'NordWireConnect.exe')}\" -uninstall-nord-wire-connect")
        win32api.RegSetValueEx(registry_key, "DisplayName", 0, win32con.REG_SZ, "NordWireConnect")
        win32api.RegSetValueEx(registry_key, "DisplayVersion", 0, win32con.REG_SZ, version)
        win32api.RegSetValueEx(registry_key, "DisplayIcon", 0, win32con.REG_SZ, os.path.join(program_files, "resources", "app_icon.ico"))
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
    colors_class.fix_windows_ansi()
    setup_logging()
    systemMessage(f"{'-'*5:^5} NordWireConnect v{version} {'-'*5:^5}")
    os.makedirs(app_data_path, exist_ok=True)

    # Only for Windows
    if main_os != "Windows":
        errorMessage("NordWireConnect is only supported on Windows.")
        sys.exit(1)
        return
    
    # Prepare NordWireConnect
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'): cur_path = os.path.dirname(sys.executable)
    else: cur_path = os.path.dirname(sys.argv[0])

    try: install()
    except Exception as e: errorMessage(f"There was an error during Installation Handler: {str(e)}"); sys.exit(1); return
if __name__ == "__main__": app()