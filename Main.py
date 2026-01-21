"""
NordWireConnect

Made by @EfazDev
https://www.efaz.dev
"""

# Modules
import os
import gc
import sys
import json
import time
import copy
import ctypes
import PyKits
import pystray # type: ignore
import threading
import platform
import subprocess
import pythoncom # type: ignore
import win32file # type: ignore
import win32api # type: ignore
import win32con # type: ignore
import win32pipe # type: ignore
import win32com.shell.shell as winshell # type: ignore
import win32event # type: ignore
import win32process # type: ignore
import tkinter as tk
from PIL import Image # type: ignore
import win32com.client # type: ignore
from queue import Queue
from tkinter import simpledialog

# Variables
pip_class = PyKits.pip()
colors_class = PyKits.Colors()
requests = PyKits.request()
main_os = platform.system()
work_queue = Queue()
pre_program_files = os.getenv("ProgramFiles")
app_data_path = os.path.join(pip_class.getLocalAppData(), "NordWireConnect")
wireguard_location = os.path.join(pre_program_files, "WireGuard")
program_files = os.path.join(pre_program_files, "NordWireConnect")
cur_path = os.path.dirname(os.path.abspath(__file__))
private_key = None
config_data = {
    "access_token": "",
    "dns": "1.1.1.1, 1.0.0.1",
    "server": "auto",
    "auto_connect": False
}
session_data = {
    "connected": False,
    "connection_text": "Not Connected.",
    "server": None,
    "current_load": None,
    "session_time": 0,
    "exact_connect": None
}
server_dictionary = {
    "Auto": [("Auto", "auto", None)]
}
country_list = None
city_list = None
full_files = False
pystray_icon = None
stop_app = False
version = "1.0.9"
PIPE_NAME = r"\\.\pipe\NordWireConnect"
ROOT = tk.Tk()

# Logging and Messages
def systemMessage(message): colors_class.print(message, colors_class.hex_to_ansi2("#3E5FFF"))
def mainMessage(message): colors_class.print(message, 15)
def errorMessage(message, title="Uh oh!"): 
    colors_class.print(message, 9)
    message = message[:256]
    if config_data.get("notifications"): notification(title, message, img=os.path.join(program_files, "resources", "error.ico"))
    else:
        if title == "": title = "NordWireConnect"
        elif title != "NordWireConnect": title = f"NordWireConnect: {title}"
        def a(): ctypes.windll.user32.MessageBoxW(0, message, title, 0x0 | 0x10)
        work_queue.put((a, ()))
def warnMessage(message): colors_class.print(message, 11)
def successMessage(message): colors_class.print(message, 10)
def successMessageBox(message, title="Success!"):
    message = message[:256]
    if config_data.get("notifications"): notification(title, message, img=os.path.join(program_files, "resources", "connected.ico"))
    else:
        if title == "": title = "NordWireConnect"
        elif title != "NordWireConnect": title = f"NordWireConnect: {title}"
        def a(): ctypes.windll.user32.MessageBoxW(0, message, title, 0x0 | 0x00000040)
        work_queue.put((a, ()))

# UI App
def getImageObj(img=None):
    try: return Image.open(img)
    except Exception: return Image.new("RGB", (64, 64), color = 'blue')
def notification(title, message, img=None):
    def a(img): 
        if not img: img = os.path.join(program_files, "resources", "app_icon.ico")
        if pystray_icon:
            pre = pystray_icon.icon
            pystray_icon.icon = getImageObj(img); pystray_icon.update_menu()
            pystray_icon.notify(message, title)
            pystray_icon.icon = pre; pystray_icon.update_menu()
    work_queue.put((a, (img,)))
def setup(icon): pystray_icon.visible = True
def quit_app(icon, item=None):
    global stop_app
    try:
        disconnect(pystray_icon)
        sess = os.path.join(app_data_path, "NordSessionData.json")
        if os.path.exists(sess): os.remove(sess)
    except: pass
    if pystray_icon: pystray_icon.stop()
    stop_app = True
def update_tray():
    if pystray_icon: pystray_icon.update_menu()
def about(icon):
    def a(): ctypes.windll.user32.MessageBoxW(0, f"Version: v{version}\nMade by @EfazDev\nhttps://www.efaz.dev", "About NordWireConnect", 0x0 | 0x00000040)
    work_queue.put((a, ()))
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
def select_server_list(message="Please select a server!"):
    win = tk.Toplevel()
    win.title("Select NordVPN Server")
    win.geometry("320x120")
    win.resizable(False, False)
    win.attributes("-topmost", True)
    win.iconphoto(False, tk.PhotoImage(file=os.path.join(program_files, "resources", "app_icon.png")))
    selected = {
        "title": None,
        "tunnel_id": None,
        "exact_mode": None
    }
    result = {"value": None}
    message_var = tk.StringVar(value=message)
    label = tk.Label(
        win,
        textvariable=message_var,
        anchor="w",
        padx=10,
        pady=8
    )
    label.pack(fill="x")
    btn_frame = tk.Frame(win)
    btn_frame.pack(side="bottom", pady=6)
    dropdown_btn = tk.Button(win, text="Select server ▼", width=28)
    dropdown_btn.pack(pady=6)
    def on_select(title, tunnel_id, exact_mode):
        selected["title"] = title
        selected["tunnel_id"] = tunnel_id
        selected["exact_mode"] = exact_mode
        if title == "Auto": dropdown_btn.config(text=tunnel_id)
        else: dropdown_btn.config(text=title)
    def build(parent_menu, data):
        if isinstance(data, dict):
            for title, value in sorted(data.items()):
                submenu = tk.Menu(parent_menu, tearoff=0)
                parent_menu.add_cascade(label=title, menu=submenu)
                build(submenu, value)
        elif isinstance(data, list):
            for title, tunnel_id, exact_mode in data:
                def b(title, tunnel_id, exact_mode):
                    parent_menu.add_command(
                        label=title,
                        command=lambda t=title, tid=tunnel_id, em=exact_mode:
                            on_select(t, tid, em)
                    )
                b(title, tunnel_id, exact_mode)
        elif isinstance(data, tuple):
            title, tunnel_id, exact_mode = data
            def b(title, tunnel_id, exact_mode):
                parent_menu.add_command(
                    label=title,
                    command=lambda t=title, tid=tunnel_id, em=exact_mode:
                        on_select(t, tid, em)
                )
            b(title, tunnel_id, exact_mode)
        else: raise TypeError(f"Unsupported menu data: {type(data)}")
    menu = tk.Menu(win, tearoff=0)
    build(menu, server_dictionary)
    def show_menu():
        x = dropdown_btn.winfo_rootx()
        y = dropdown_btn.winfo_rooty() + dropdown_btn.winfo_height()
        menu.tk_popup(x, y)
    dropdown_btn.config(command=show_menu)
    def confirm():
        if selected["tunnel_id"] is not None: result["value"] = selected
        win.destroy()
    def cancel():
        result["value"] = None
        win.destroy()
    tk.Button(btn_frame, text="Confirm", width=10, command=confirm).pack(side="left", padx=6)
    tk.Button(btn_frame, text="Cancel", width=10, command=cancel).pack(side="right", padx=6)
    win.protocol("WM_DELETE_WINDOW", cancel)
    win.grab_set()
    win.wait_window()
    return result["value"]
def connect_selected_server_button(icon=None):
    selected_server = select_server_list("Please select a server to connect to!")
    if not selected_server:
        return
    work_queue.put((connect_server, (selected_server["tunnel_id"], pystray_icon, selected_server["exact_mode"])))
    update_tray()
def save_auto_selected_server_button(icon=None):
    selected_server = select_server_list("Please select a server to save as a default location!")
    if not selected_server:
        return
    default_location(selected_server["tunnel_id"], pystray_icon, selected_server["exact_mode"])
    update_tray()
def setup_shortcuts(icon=None):
    try:
        pythoncom.CoInitialize()
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            desk_shortpath = os.path.join(os.environ["USERPROFILE"], "Desktop", f"NordWireConnect.lnk")
            desk_short = shell.CreateShortcut(desk_shortpath)
            desk_short.TargetPath = sys.executable
            desk_short.IconLocation = os.path.join(program_files, "resources", "app_icon.ico")
            desk_short.Description = "Open NordWireConnect for VPN!"
            desk_short.Save()
            startm_shortpath = os.path.join(os.getenv("APPDATA"), "Microsoft", "Windows", "Start Menu", "Programs", f"NordWireConnect.lnk")
            startm_short = shell.CreateShortcut(startm_shortpath)
            startm_short.TargetPath = sys.executable
            startm_short.IconLocation = os.path.join(program_files, "resources", "app_icon.ico")
            startm_short.Description = "Open NordWireConnect for VPN!"
            startm_short.Save()
            successMessageBox(f"Successfully set up shortcuts for NordWireConnect!", "NordWireConnect")
        finally: pythoncom.CoUninitialize()
    except Exception as e: errorMessage(f"Unable to save shortcuts: {str(e)}")
def delete_shortcuts(icon=None):
    try:
        desk_shortpath = os.path.join(os.environ["USERPROFILE"], "Desktop", f"NordWireConnect.lnk")
        if os.path.exists(desk_shortpath): os.remove(desk_shortpath)
        startm_shortpath = os.path.join(os.getenv("APPDATA"), "Microsoft", "Windows", "Start Menu", "Programs", f"NordWireConnect.lnk")
        if os.path.exists(startm_shortpath): os.remove(startm_shortpath)
    except Exception as e: errorMessage(f"Unable to delete shortcuts: {str(e)}")
def setup_registry(icon=None):
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
def delete_registry(icon=None):
    try:
        app_reg_path = "Software\\NordWireConnect"
        win32api.RegDeleteKey(win32con.HKEY_LOCAL_MACHINE, app_reg_path)
        registry_path = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\NordWireConnect"
        win32api.RegDeleteKey(win32con.HKEY_LOCAL_MACHINE, registry_path)
    except Exception as e: errorMessage(f"Unable to delete registry: {str(e)}")
def uninstall_app():
    if not pip_class.getIfRunningWindowsAdmin():
        mainMessage("Please relaunch this program as Admin to uninstall!")
        relaunch_as_admin()
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
    if os.path.exists(os.path.join(program_files, "Service", "NordWireConnectService.exe")): os.remove(os.path.join(program_files, "Service", "NordWireConnectService.exe"))
    mainMessage("Clearing shortcuts on current user..")
    delete_shortcuts()
    mainMessage("Clearing Registry..")
    delete_registry()
    mainMessage("Removing app..")
    cmd = f'''
    timeout /t 2 >nul
    taskkill /IM NordWireConnect.exe /F
    del "{os.path.join(program_files, "Main", "NordWireConnect.exe")}"
    rmdir /s /q "{program_files}"
    '''
    subprocess.Popen(
        ["cmd.exe", "/c", cmd],
        creationflags=subprocess.CREATE_NO_WINDOW
    )
    sys.exit(0)
def check_stop_flag():
    if stop_app:
        mainMessage("Exiting app..")
        ROOT.quit()
        sys.exit(0)
        return
    ROOT.after(100, check_stop_flag)
def worker():
    while True:
        func, args = work_queue.get()
        try: func(*args)
        except Exception as e: errorMessage(f"Error while running worker ({getattr(func, "__name__")}): {str(e)}")
        finally: work_queue.task_done()

# Data Handling
def session_data_modified():
    with open(os.path.join(app_data_path, "NordSessionData.json"), "w") as f: json.dump(session_data, f, indent=4)
def run_cache():
    global country_list
    global city_list
    mainMessage("Updating NordVPN server cache..")
    recommended_servers = requests.get("https://api.nordvpn.com/v1/servers/recommendations?&filters\\[servers_technologies\\]\\[identifier\\]=wireguard_udp&limit=300")
    all_servers = requests.get("https://api.nordvpn.com/v1/servers?&filters\\[servers_technologies\\]\\[identifier\\]=wireguard_udp&limit=10000")
    server_list = recommended_servers if (config_data.get("optimize_server_list", False) == True) else all_servers
    raw_servers = server_list.json if server_list.ok else []
    raw_all_servers = all_servers.json if all_servers.ok else []
    all_cities = {}
    all_countries = requests.get("https://api.nordvpn.com/v1/countries")
    all_countries = all_countries.json if all_countries.ok else []
    for s in raw_all_servers:
        for l in s["locations"]:
            c = l["country"]["city"]
            all_cities[c["id"]] = c["name"]
    all_cities = [{"id": i, "name": n} for i, n in all_cities.items()]
    with open(os.path.join(app_data_path, "NordServerCache.json"), "w") as f: json.dump(raw_servers, f)
    with open(os.path.join(app_data_path, "NordCountryCache.json"), "w") as f: json.dump(all_countries, f)
    with open(os.path.join(app_data_path, "NordCityCache.json"), "w") as f: json.dump(all_cities, f)
    country_list = list(all_countries)
    city_list = list(all_cities)
    recommended_servers.json.clear()
    del recommended_servers.text
    del recommended_servers.raw_data
    all_servers.json.clear()
    del all_servers.text
    del all_servers.raw_data
    raw_servers.clear()
    del all_servers
    all_countries.clear()
    raw_all_servers.clear()
    all_cities.clear()
    del server_list
    del recommended_servers
    gc.collect()
def send_command(cmd: str) -> str:
    while True:
        try:
            win32pipe.WaitNamedPipe(PIPE_NAME, 5000)
            handle = win32file.CreateFile(
                PIPE_NAME,
                win32file.GENERIC_READ | win32file.GENERIC_WRITE,
                0,
                None,
                win32file.OPEN_EXISTING,
                0,
                None
            ); break
        except Exception as e: time.sleep(0.2)
    win32pipe.SetNamedPipeHandleState( 
        handle, 
        win32pipe.PIPE_READMODE_MESSAGE, 
        None, 
        None 
    ) 
    win32file.WriteFile(handle, cmd.encode("utf-8"))
    _, data = win32file.ReadFile(handle, 4096) 
    win32file.CloseHandle(handle) 
    return data.decode("utf-8")
def format_seconds(seconds):
    seconds = int(seconds)
    days, remaining_seconds = divmod(seconds, 86400)
    hours, remaining_seconds = divmod(remaining_seconds, 3600)
    minutes, seconds = divmod(remaining_seconds, 60)
    return f"{days:02d}:{hours:02d}:{minutes:02d}:{seconds:02d}"
def get_if_virtual_location(server: dict):
    for sp in server.get("specifications", []):
        if sp["identifier"] == "virtual_location":
            for val in sp["values"]:
                if val["value"] == "true": return True
    return False
def change_dns(icon=None):
    inputted = simpledialog.askstring(title="DNS Input", prompt="Enter IPv4 DNS servers to use (comma separated):")
    ips = inputted.split(", ")
    added = [ip for ip in ips if requests.get_if_ip(ip)]
    if added:
        config_data["dns"] = ", ".join(added)
        save_configuration()
        successMessageBox(f"Successfully changed DNS to: {', '.join(added)}", "NordWireConnect")
        update_tray()
    else: errorMessage("No valid IPv4 DNS servers were provided.")
def change_access_token(icon=None):
    global private_key
    inputted = simpledialog.askstring(title="Access Token Input", prompt="Enter your generated NordVPN Access Token from the NordVPN website (https://my.nordaccount.com/dashboard/nordvpn/):")
    if len(inputted) > 50:
        test_request = requests.get(f"https://api.nordvpn.com/v1/users/services/credentials", auth=["token", inputted])
        if test_request.ok:
            test_request = test_request.json
            config_data["username"] = test_request.get("username", "")
            config_data["access_token"] = inputted
            private_key = test_request.get("nordlynx_private_key", "")
            successMessageBox("Successfully changed NordVPN Access Token!", "NordWireConnect")
            save_configuration()
            update_tray()
        else: errorMessage("The provided NordVPN Access Token is invalid.")
    else: errorMessage("The provided NordVPN Access Token is invalid.")
def default_location(server_name: str, icon=None, exact_mode=None): 
    if exact_mode: config_data["default_location"] = f"{exact_mode}_{server_name}"
    else: config_data["default_location"] = server_name
    save_configuration()
def mark_not_connected(icon=None):
    session_data["session_time"] = 0
    session_data["connection_text"] = "Not Connected."
    session_data["connected"] = False
    session_data["server"] = None
    session_data_modified()
    update_tray()
def save_configuration(icon=None):
    mainMessage("Saving configuration..")
    with open(os.path.join(app_data_path, "ConnectConfig.json"), "w") as f: json.dump(config_data, f, indent=4)
def export_configuration(icon=None):
    save_configuration(icon)
    subprocess.run(f"explorer \"{app_data_path}\"", shell=True)
    successMessageBox("Successfully saved configuration! Opened the folder with the configuration.", "NordWireConnect")
def format_size(size_bytes):
    if size_bytes == 0: return "0 Bytes"
    size_units = ["Bytes", "KB", "MB", "GB", "TB"]
    unit_index = 0
    while size_bytes >= 1024 and unit_index < len(size_units) - 1: size_bytes /= 1024; unit_index += 1
    return f"{size_bytes:.2f} {size_units[unit_index]}"
def get_folder_size(folder_path, formatWithAbbreviation=True):
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
    
# Toggle Functions
def change_load_swapping(icon=None, item=None):
    config_data["load_swapping"] = not config_data.get("load_swapping", False)
    save_configuration()
    update_tray()
def check_load_swapping(icon=None, item=None): return config_data.get("load_swapping", False) == True
def change_auto_connect(icon=None, item=None):
    config_data["auto_connect"] = not config_data.get("auto_connect", True)
    save_configuration()
    update_tray()
def check_auto_connect(icon=None, item=None): return config_data.get("auto_connect", True) == True
def change_notifications(icon=None, item=None):
    config_data["notifications"] = not config_data.get("notifications", False)
    save_configuration()
    update_tray()
def check_notifications(icon=None, item=None): return config_data.get("notifications", False) == True
def change_server_list(icon=None, item=None):
    config_data["optimize_server_list"] = not config_data.get("optimize_server_list", False)
    work_queue.put((run_cache, ()))
    save_configuration()
    update_tray()
def check_server_list(icon=None, item=None): return config_data.get("optimize_server_list", False) == True
def change_split_lan_routing(icon=None, item=None):
    config_data["split_lan_routing"] = not config_data.get("split_lan_routing", False)
    save_configuration()
    update_tray()
def check_split_lan_routing(icon=None, item=None): return config_data.get("split_lan_routing", False) == True

# Differenting Status
def differ_from_status_text(icon):
    if session_data["connected"] == True: return "Disconnect from Server"
    elif session_data["connection_text"].startswith("Connecting.."): return "Cancel Connection"
    else: return "Connect to Server"
def differ_from_status_action(icon):
    if session_data["connection_text"].startswith("Connecting"):
        session_data["session_time"] = 0
        session_data["connection_text"] = "Not Connected."
        session_data["connected"] = False
        session_data["server"] = None
        update_tray()
    elif session_data["connected"] == True: 
        s = disconnect(pystray_icon)
        if s == "0": successMessageBox("Successfully disconnected from NordVPN server!", "NordWireConnect")
    else: 
        if config_data["default_location"] == "auto":
            config_data["server"] = "auto"
            work_queue.put((connect, (pystray_icon,)))
        else:
            mode, server_loc = config_data["default_location"].split("_")
            config_data["server"] = server_loc
            work_queue.put((connect, (pystray_icon, mode)))
def differ_from_status_action2(icon):
    if session_data["connection_text"].startswith("Connecting") or session_data["connection_text"].startswith("Not"): return
    work_queue.put((connect, (pystray_icon,)))
def differ_to_status(icon): return session_data["connection_text"]
def differ_to_status1(icon):
    if session_data["connected"] == True: return f"City: {session_data['server']['locations'][0]['country']['city']['name']}"
    else: return "City: N/A"
def differ_to_status2(icon): 
    if session_data["connected"] == True: return f"IP Address: {session_data['server']['station']}"
    else: return "IP Address: N/A"
def differ_to_status3(icon): 
    if session_data["connected"] == True: return f"Hostname: {session_data['server']['hostname']}"
    else: return "Hostname: N/A"
def differ_to_status4(icon): 
    if session_data["connected"] == True: return f"Load: {session_data['current_load']}%"
    else: return "Load: N/A"
def differ_to_status5(icon): 
    if session_data["connected"] == True: return f"Session Time: {format_seconds(session_data['session_time'])}"
    else: return "Session Time: N/A"
def differ_to_config1(icon): return f"DNS Servers: {config_data['dns']}"
def differ_to_config2(icon): return f"Account: {config_data.get('username', 'N/A')} (Token {len(config_data['access_token']) > 60 and 'Given' or 'Ungiven'})"
def differ_to_config3(icon): 
    if config_data.get("default_location") and config_data["default_location"] != "auto":
        _, server_loc = config_data["default_location"].split("_")
        return f"Default Location: {server_loc}"
    else:
        return "Default Location: Auto"

# Connections
def get_allowed_ips(): 
    if config_data.get("split_lan_routing"): return "0.0.0.0/5, 8.0.0.0/7, 11.0.0.0/8, 12.0.0.0/6, 16.0.0.0/4, 32.0.0.0/3, 64.0.0.0/2, 128.0.0.0/3, 160.0.0.0/5, 168.0.0.0/6, 172.0.0.0/12, 172.32.0.0/11, 172.64.0.0/10, 172.128.0.0/9, 173.0.0.0/8, 174.0.0.0/7, 176.0.0.0/4, 192.0.0.0/9, 192.128.0.0/11, 192.160.0.0/13, 192.169.0.0/16, 192.170.0.0/15, 192.172.0.0/14, 192.176.0.0/12, 192.192.0.0/10, 193.0.0.0/8, 194.0.0.0/7, 196.0.0.0/6, 200.0.0.0/5, 208.0.0.0/4, ::/1, 8000::/2, c000::/3, e000::/4, f000::/5, f800::/6, fe00::/9, ff00::/8"
    else: return "0.0.0.0/0, ::/1, 8000::/2, c000::/3, e000::/4, f000::/5, f800::/6, fe00::/9, ff00::/8"
def connect_server(server_name: str, icon=None, exact_mode=None):
    global config_data
    previous_data = config_data
    config_data["server"] = server_name
    try: connect(icon, exact_mode=exact_mode)
    except Exception: pass
    config_data = previous_data
def connect_session():
    if session_data["connected"] == True and config_data.get("auto_connect", True) == True:
        mainMessage("Restoring previous session...")
        connect_server(session_data["server"]["name"].replace(" ", "").replace("#", ""), pystray_icon, exact_mode="server")
    elif send_command("connection-status") != "NotConnected": disconnect()
def connect(icon=None, exact_mode=None):
    global session_data
    try:
        session_data["session_time"] = 0
        session_data["connected"] = False
        session_data["connection_text"] = "Connecting..."
        session_data_modified()
        update_tray()

        if not private_key:
            errorMessage("Please provide your Access Token before trying to connect!")
            mark_not_connected(icon)
            return

        # Load WireGuard Installation
        if not os.path.exists(wireguard_location):
            successMessageBox("Wireguard was not found in your computer. Installing now from wireguard.com.", "NordWireConnect")
            download_url = "https://download.wireguard.com/windows-client/wireguard-installer.exe"
            installer_path = os.path.join(app_data_path, "wireguard-installer.exe")
            mainMessage("Downloading WireGuard installer...")
            download = requests.download(download_url, installer_path)
            if download.ok:
                mainMessage("Installing WireGuard...")
                proc = winshell.ShellExecuteEx(
                    lpVerb="runas",
                    lpFile=installer_path,
                    lpParameters="",
                    nShow=win32con.SW_SHOWNORMAL,
                    fMask=0x00000040
                )
                handle = proc["hProcess"]
                win32event.WaitForSingleObject(handle, win32event.INFINITE)
                wireguard_exit_code = win32process.GetExitCodeProcess(handle)
                send_command("end-wireguard")
                send_command("end-wireguard-installer")
                time.sleep(1)
                if os.path.exists(installer_path): os.remove(installer_path)
                if wireguard_exit_code == 0:
                    mainMessage("WireGuard installed successfully.")
                    time.sleep(3)
                else:
                    errorMessage("Failed to install WireGuard.")
                    mark_not_connected(icon)
                    return
            else:
                errorMessage("Failed to download WireGuard installer.")
                mark_not_connected(icon)
                return

        # Load NordVPN Server Recommendations
        if not requests.get_if_connected(): 
            mainMessage("Stopping WireGuard services...")
            prev_stats = copy.deepcopy(session_data)
            disconnect(icon)
            session_data = prev_stats
        if not requests.get_if_connected():
            errorMessage("There was an issue trying to connect to your internet. Press Cancel Connection if you want to disconnect.")
            while not requests.get_if_connected():
                if session_data["connection_text"] == "Not Connected." and session_data["connected"] == False: break
                time.sleep(1)
            if session_data["connection_text"] == "Not Connected." and session_data["connected"] == False:
                mainMessage("Canceled connection.")
                mark_not_connected(icon)
                return
        if config_data["server"] == "auto" or config_data["server"] == None or exact_mode == None:
            all_server_recommendations = requests.get("https://api.nordvpn.com/v1/servers/recommendations?&filters\\[servers_technologies\\]\\[identifier\\]=wireguard_udp&limit=100")
            if not all_server_recommendations.ok:
                errorMessage("Unable to get NordVPN server recommendations.")
                mark_not_connected(icon)
                return
            all_server_recommendations = all_server_recommendations.json
        elif exact_mode == "country":
            if not country_list:
                errorMessage("Unable to get NordVPN country list.")
                mark_not_connected(icon)
                return
            country_id = None
            for c in country_list:
                if c.get("name") == config_data["server"]:
                    country_id = c.get("id")
                    break
            if country_id == None:
                errorMessage("Unable to get NordVPN country ID.")
                mark_not_connected(icon)
                return
            all_server_recommendations = requests.get(f"https://api.nordvpn.com/v1/servers/recommendations?&filters[country_id]={country_id}&filters\\[servers_technologies\\]\\[identifier\\]=wireguard_udp&limit=100")
            if not all_server_recommendations.ok:
                errorMessage("Unable to get NordVPN server recommendations.")
                mark_not_connected(icon)
                return
            all_server_recommendations = all_server_recommendations.json
        elif exact_mode == "city":
            if not city_list:
                errorMessage("Unable to get NordVPN city list.")
                mark_not_connected(icon)
                return
            city_id = None
            for c in city_list:
                if c.get("name") == config_data["server"]:
                    city_id = c.get("id")
                    break
            if city_id == None:
                errorMessage("Unable to get NordVPN city ID.")
                mark_not_connected(icon)
                return
            all_server_recommendations = requests.get(f"https://api.nordvpn.com/v1/servers/recommendations?&filters[country_city_id]={city_id}&filters\\[servers_technologies\\]\\[identifier\\]=wireguard_udp&limit=100")
            if not all_server_recommendations.ok:
                errorMessage("Unable to get NordVPN server recommendations.")
                mark_not_connected(icon)
                return
            all_server_recommendations = all_server_recommendations.json
        else:
            all_server_recommendations = requests.get("https://api.nordvpn.com/v1/servers?&filters\\[servers_technologies\\]\\[identifier\\]=wireguard_udp&limit=10000")
            if not all_server_recommendations.ok:
                errorMessage("Unable to get NordVPN server recommendations.")
                mark_not_connected(icon)
                return
            all_server_recommendations = all_server_recommendations.json
            all_server_recommendations.sort(key=lambda x: x["load"])
        mainMessage(f"Loaded {len(all_server_recommendations)} NordVPN server recommendation(s).")

        # Stop Wireguard Service
        mainMessage("Stopping WireGuard services...")
        prev_stats = copy.deepcopy(session_data)
        disconnect(icon)
        session_data = prev_stats

        # Block Instances of IPv6
        mainMessage("Blocking Public IPv6 Addresses...")
        send_command("block-public-ipv6")

        # Start Finding Servers
        mainMessage(f"Connecting to NordVPN Servers in {config_data['server']}..")
        connected_server = None
        for s in all_server_recommendations:
            shortened_name = s["name"].replace(" ", "").replace("#", "")
            country_name = s["locations"][0]["country"]["name"].replace(" ", "")
            city_name = s["locations"][0]["country"]["city"]["name"].replace(" ", "")
            if not (config_data["server"] == "auto" or config_data["server"] == None) and not (shortened_name == config_data["server"].replace(" ", "").replace("#", "") or city_name == config_data["server"].replace(" ", "") or country_name == config_data["server"].replace(" ", "")):
                continue
            if session_data["connection_text"] == "Not Connected." and session_data["connected"] == False: break
            if s["status"] != "online" or s["load"] == 0:
                mainMessage(f"Server offline: {s['name']}")
                continue

            mainMessage(f"Attempting connection for: {s['name']}")
            wireguard_metadata = None
            for m in s.get("technologies", []):
                if m.get("identifier") == "wireguard_udp":
                    wireguard_metadata = m.get("metadata", {})
                    break
            if not wireguard_metadata: continue
            configuration = f"""# {shortened_name}.{city_name}.conf
[Interface]
PrivateKey = {private_key}
Address = 10.5.0.2/32
DNS = {config_data.get('dns')}

[Peer]
PublicKey = {wireguard_metadata[0].get("value")}
AllowedIPs = {get_allowed_ips()}
Endpoint = {s.get('station')}:51820
PersistentKeepalive = 25"""
            config_path = os.path.join(app_data_path, f"{shortened_name}.{city_name}.conf")
            with open(config_path, "w") as f: f.write(configuration)
            if session_data["connection_text"] == "Not Connected." and session_data["connected"] == False: break
            send_command(f"uninstall-wire-tunnel {shortened_name}.{city_name}")
            add = send_command(f"install-wire-tunnel {config_path}")
            if add == "0":
                mainMessage(f"Added new WireGuard configuration for server: {s['name']}")
                tunnel_check()
                time.sleep(1)
                if tunnel_check():
                    connected_server = s
                    break
                else: 
                    send_command(f"uninstall-wire-tunnel {shortened_name}.{city_name}")
                    continue
            else:
                errorMessage(f"Failed to add WireGuard configuration for server: {s['name']}")
                continue
        all_server_recommendations.clear()
        gc.collect()
        if session_data["connection_text"] == "Not Connected." and session_data["connected"] == False: return
        elif connected_server:
            successMessage(f"Successfully connected to server: {connected_server['name']}!")
            successMessage(f"Load: {connected_server['load']}%")
            successMessage(f"City: {connected_server['locations'][0]['country']['city']['name']}")
            successMessage(f"IP Address: {connected_server['station']}")
            successMessage(f"Hostname: {connected_server['hostname']}")
            session_data["connected"] = True
            session_data["server"] = connected_server
            session_data["current_load"] = connected_server['load']
            session_data["connection_text"] = f"Connected to {connected_server['name']}"
            session_data["session_time"] = 0
            session_data["exact_connect"] = exact_mode
            session_data_modified()
            successMessageBox(f"Successfully connected to server: {connected_server['name']}!", "Yay!")
            image = getImageObj(os.path.join(program_files, "resources", "connected.ico"))
            if pystray_icon:
                pystray_icon.icon = image
                update_tray()
        else:
            errorMessage("Unable to connect to a NordVPN server.")
            mark_not_connected(icon)
            return
    except Exception as e:
        errorMessage(f"Unable to connect to a NordVPN server due to an error: {str(e)}.")
        mark_not_connected(icon)
        return
def disconnect(icon=None):
    mainMessage("Disconnecting Wireguard..")
    if session_data["connected"] == True:
        server_name = session_data["server"]["name"].replace(" ", "").replace("#", "")
        send_command(f"uninstall-wire-tunnel {server_name}")
    session_data["session_time"] = 0
    session_data["connection_text"] = "Not Connected."
    session_data["connected"] = False
    session_data["server"] = None
    session_data_modified()
    image = getImageObj(os.path.join(program_files, "resources", "app_icon.ico"))
    if pystray_icon:
        pystray_icon.icon = image
        pystray_icon.update_menu()
    send_command("end-wireguard-tunnels")
    send_command("unlock-public-ipv6")
    return send_command("end-wireguard")
def brute_end_wireguard(icon=None):
    disconnect(icon)
    mark_not_connected(icon)
def handle_stat_thread(icon):
    count = 0
    since_connected = 0
    while True:
        time.sleep(1)
        try:
            if session_data["connected"] == True:
                count += 1
                since_connected += 1
                if not requests.get_if_connected() and not session_data["connection_text"].startswith("Connecting"):
                    mainMessage("Reconnecting due to loss of connection.")
                    differ_from_status_action2(icon)
                    continue
                if count % 20 == 0:
                    server_id = session_data["server"]["id"]
                    servers = requests.get(f"https://api.nordvpn.com/v1/servers?&filters[servers.id]={server_id}&limit=1")
                    servers = servers.json if servers.ok else []
                    found_server = None
                    for s in servers:
                        if s["id"] == server_id:
                            found_server = s
                            break
                    session_data["current_load"] = found_server["load"] if found_server else session_data["current_load"]
                    if check_load_swapping() and session_data["current_load"] > 50:
                        mainMessage("Reconnecting due to load > 50%..")
                        differ_from_status_action2(icon)
                session_data["session_time"] = since_connected
                update_tray()
            else:
                count = 0
                since_connected = 0
        except Exception: pass
def tunnel_check():
    return requests.get_if_connected()

# Main Runtime
def app():
    global config_data
    global pystray_icon
    global private_key
    global session_data
    global cur_path
    colors_class.fix_windows_ansi()
    systemMessage(f"{'-'*5:^5} NordWireConnect v{version} {'-'*5:^5}")
    ROOT.withdraw()

    os.makedirs(app_data_path, exist_ok=True)

    # Only for Windows
    if main_os != "Windows":
        errorMessage("NordWireConnect is only supported on Windows.")
        sys.exit(1)
        return

    # Load Uninstaller
    if "-uninstall-nord-wire-connect" in sys.argv: 
        try: uninstall_app()
        except Exception as e: errorMessage(f"There was an error during Installation Handler: {str(e)}"); sys.exit(1); return

    # Prepare NordWireConnect
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'): cur_path = os.path.dirname(sys.executable)
    else: cur_path = os.path.dirname(sys.argv[0])
    if pip_class.getAmountOfProcesses("NordWireConnect.exe") > 2: 
        errorMessage("NordWireConnect is already running right now!")
        sys.exit(1)
        return
    
    # Load Configuration
    mainMessage("Loading Configuration..")
    try:
        if os.path.exists(os.path.join(app_data_path, "ConnectConfig.json")):
            with open(os.path.join(app_data_path, "ConnectConfig.json"), "r") as f: config_data = json.load(f)
        if os.path.exists(os.path.join(app_data_path, "NordSessionData.json")): 
            with open(os.path.join(app_data_path, "NordSessionData.json"), "r") as f: session_data = json.load(f)
    except Exception as e: errorMessage(f"There was an error trying to load configuration: {str(e)}"); sys.exit(1); return

    # Break Hanging Connections
    try:
        if not requests.get_if_connected():
            mainMessage("Quick Disconnecting Wireguard..")
            pre = copy.deepcopy(config_data)
            pre2 = copy.deepcopy(session_data)
            disconnect()
            config_data = pre
            session_data = pre2
            session_data_modified()
            while not requests.get_if_connected(): time.sleep(0.1)
    except Exception as e: errorMessage(f"There was an error trying to break hanging connections: {str(e)}"); sys.exit(1); return

    # Fetch NordVPN Credentials
    if config_data.get("access_token"):
        mainMessage("Getting NordVPN credentials..")
        try:
            credentials = requests.get(f"https://api.nordvpn.com/v1/users/services/credentials", auth=["token", config_data.get("access_token")])
            if not credentials.ok:
                errorMessage("Unable to get NordVPN credentials. Is the access token valid?")
                sys.exit(1)
                return
            credentials = credentials.json
            private_key = credentials.get("nordlynx_private_key", "")
        except Exception as e:
            errorMessage(f"Unable to get NordVPN credentials. Exception: {str(e)}")
            sys.exit(1)
            return
    
    # Load Servers
    try:
        mainMessage("Fetching Servers..")
        if os.path.exists(os.path.join(app_data_path, "NordServerCache.json")):
            with open(os.path.join(app_data_path, "NordServerCache.json"), "r") as f: raw_servers = json.load(f)
            work_queue.put((run_cache, ()))
        else:
            run_cache()
            with open(os.path.join(app_data_path, "NordServerCache.json"), "r") as f: raw_servers = json.load(f)
        for s in raw_servers:
            if s["status"] != "online" or s["load"] == 0:
                continue
            virutal_location = get_if_virtual_location(s)
            wireguard_metadata = None
            for m in s.get("technologies", []):
                if m.get("identifier") == "wireguard_udp":
                    wireguard_metadata = m.get("metadata", {})
                    break
            if not wireguard_metadata:
                continue
            country_name = s["locations"][0]["country"]["name"]
            city_name = s["locations"][0]["country"]["city"]["name"]
            city_with_virtual_text = city_name + (" - Virtual" if virutal_location else "")
            server_name = s["name"]
            if country_name not in server_dictionary:
                server_dictionary[country_name] = {
                    "Auto": [("Auto", country_name, "country")]
                }
            if city_with_virtual_text not in server_dictionary.get(country_name):
                server_dictionary[country_name][city_with_virtual_text] = [("Auto", city_name, "city")]
            server_dictionary[country_name][city_with_virtual_text].append((server_name, server_name.replace(" ", "").replace("#", ""), "server"))
        raw_servers.clear()
        gc.collect()
    except Exception as e: errorMessage(f"Unable to load server list. Exception: {str(e)}"); sys.exit(1); return

    # Create Tray App
    try:
        mainMessage("Creating Tray App..")
        def create_mini_func(func):
            def a(icon, item): work_queue.put((func, (pystray_icon,)))
            return a
        image = getImageObj(os.path.join(program_files, "resources", "app_icon.ico"))
        pystray_icon = pystray.Icon(
            "NordWireConnect",
            image,
            "NordWireConnect",
            menu=pystray.Menu(
                pystray.MenuItem("NordWireConnect", lambda icon, item: None, enabled=False),
                pystray.MenuItem(differ_to_status, lambda icon, item: None, enabled=False),
                pystray.MenuItem(differ_to_status1, lambda icon, item: None, enabled=False, visible=lambda icon: session_data["connected"] == True),
                pystray.MenuItem(differ_to_status2, lambda icon, item: None, enabled=False, visible=lambda icon: session_data["connected"] == True),
                pystray.MenuItem(differ_to_status3, lambda icon, item: None, enabled=False, visible=lambda icon: session_data["connected"] == True),
                pystray.MenuItem(differ_to_status4, lambda icon, item: None, enabled=False, visible=lambda icon: session_data["connected"] == True),
                pystray.MenuItem(differ_to_status5, lambda icon, item: None, enabled=False, visible=lambda icon: session_data["connected"] == True),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem(differ_from_status_text, create_mini_func(differ_from_status_action)),
                pystray.MenuItem("Reconnect Servers", create_mini_func(differ_from_status_action2), visible=lambda icon: session_data["connected"] == True),
                pystray.MenuItem("Servers", lambda icon, item: ROOT.after(0, connect_selected_server_button, icon)),
                pystray.MenuItem("Configuration", pystray.Menu(
                    pystray.MenuItem(differ_to_config1, lambda icon, item: None, enabled=False),
                    pystray.MenuItem(differ_to_config2, lambda icon, item: None, enabled=False),
                    pystray.Menu.SEPARATOR,
                    pystray.MenuItem("Change DNS Servers", lambda icon, item: ROOT.after(0, change_dns, icon)),
                    pystray.MenuItem("Change Access Token", lambda icon, item: ROOT.after(0, change_access_token, icon)),
                    pystray.MenuItem(differ_to_config3, lambda icon, item: ROOT.after(0, save_auto_selected_server_button, icon)),
                    pystray.MenuItem("Set Shortcuts", create_mini_func(setup_shortcuts)),
                    pystray.MenuItem("Brute End Wireguard", create_mini_func(brute_end_wireguard)),
                    pystray.MenuItem("Load Swapping", change_load_swapping, checked=check_load_swapping, radio=True),
                    pystray.MenuItem("Auto Reconnect", change_auto_connect, checked=check_auto_connect, radio=True),
                    pystray.MenuItem("Notifications", change_notifications, checked=check_notifications, radio=True),
                    pystray.MenuItem("Optimize Server List", change_server_list, checked=check_server_list, radio=True),
                    pystray.MenuItem("Split Private LAN Routing", change_split_lan_routing, checked=check_split_lan_routing, radio=True),
                )),
                pystray.MenuItem("About NordWireConnect", about),
                pystray.MenuItem("Quit", quit_app)
            )
        )
    except Exception as e: errorMessage(f"Unable to run app. Exception: {str(e)}"); sys.exit(1); return

    # Run Main Loop
    try:
        mainMessage("Finishing app load!")
        pystray_icon.run_detached(setup)
        threading.Thread(target=handle_stat_thread, args=[pystray_icon], daemon=True).start()
        threading.Thread(target=worker, daemon=True).start()
        work_queue.put((connect_session, ()))
        ROOT.after(100, check_stop_flag)
        ROOT.iconphoto(True, tk.PhotoImage(file=os.path.join(program_files, "resources", "app_icon.png"))) 
        tk.mainloop()
    except Exception as e: errorMessage(f"Unable to run app. Exception: {str(e)}"); sys.exit(1); return
if __name__ == "__main__": app()