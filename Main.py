"""
NordWireConnect

Made by @EfazDev
https://www.efaz.dev
"""

# Modules
import os
import sys
import json
import time
import ctypes
import typing
import PyKits
import pystray # type: ignore
import logging
import datetime
import platform
import win32api # type: ignore
import win32con # type: ignore
import threading
import pythoncom # type: ignore
import win32file # type: ignore
import win32pipe # type: ignore
import subprocess
import win32event # type: ignore
import win32process # type: ignore
import tkinter as tk
from PIL import Image # type: ignore
import win32com.client # type: ignore
from queue import Queue
from tkinter import simpledialog
import win32com.shell.shell as winshell # type: ignore

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
private_key = None
config_data = {
    "access_token": "",
    "username": "",
    "dns": "103.86.96.100, 103.86.99.100", # NordVPN DNS
    "server": "auto",
    "auto_connect": True,
    "notifications": False,
    "optimize_server_list": False,
    "default_location": "auto",
    "load_swapping": False,
    "split_lan_routing": False,
    "loss_connect_protection": True
}
session_data = {
    "connected": False,
    "connection_text": "Not Connected",
    "server": None,
    "current_load": None,
    "session_time": 0,
    "exact_connect": None
}
full_files = False
pystray_icon = None
stop_app = False
version = "1.1.5"
service_pipe = r"\\.\pipe\NordWireConnect"
tk_root = tk.Tk()
Icon = pystray.Icon
MenuItem = pystray.MenuItem

# Logging and Messages
def systemMessage(message: str): colors_class.print(message, colors_class.hex_to_ansi2("#3E5FFF"))
def mainMessage(message: str): colors_class.print(message, 15)
def errorMessage(message: str, title: str="Uh oh!"): 
    colors_class.print(message, 9)
    message = message[:256]
    if config_data.get("notifications"): notification(title, message, img=os.path.join(program_files, "resources", "error.ico"))
    else:
        if title == "": title = "NordWireConnect"
        elif title != "NordWireConnect": title = f"NordWireConnect: {title}"
        def a(): ctypes.windll.user32.MessageBoxW(0, message, title, 0x0 | 0x10)
        addToThread(a, ())
def warnMessage(message: str): colors_class.print(message, 11)
def successMessage(message: str): colors_class.print(message, 10)
def successMessageBox(message: str, title: str="Success!"):
    message = message[:256]
    if config_data.get("notifications"): notification(title, message, img=os.path.join(program_files, "resources", "connected.ico"))
    else:
        if title == "": title = "NordWireConnect"
        elif title != "NordWireConnect": title = f"NordWireConnect: {title}"
        def a(): ctypes.windll.user32.MessageBoxW(0, message, title, 0x0 | 0x00000040)
        addToThread(a, ())
def setup_logging():
    handler_name = "Main"
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

# UI App
def addToThread(function: typing.Callable, args: typing.Iterable, connection: bool=False): 
    if connection == False: threading.Thread(target=function, args=args, daemon=True).start()
    else: work_queue.put((function, args))
def unicon(function: typing.Callable, args: typing.Iterable=[]) -> typing.Callable:
    def new_func(icon: Icon=None, item: MenuItem=None): return function(*args)
    return new_func
def getImageObj(img: str=None) -> Image.Image:
    try: return Image.open(img)
    except Exception: return Image.new("RGB", (64, 64), color = 'blue')
def notification(title: str, message: str, img: str=None):
    def a():
        if pystray_icon: pystray_icon.notify(message, title)
    addToThread(a, ())
def set_icon(img: str):
    image = getImageObj(os.path.join(program_files, "resources", img))
    if pystray_icon:
        pystray_icon.icon = image
        update_tray()
    if img != "connected.ico": unneeding_state()
def setup(): pystray_icon.visible = True
def needing_state(): ctypes.windll.kernel32.SetThreadExecutionState(0x80000001 | 0x00000040)
def unneeding_state(): ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)
def quit_app():
    global stop_app
    try:
        if config_data.get("auto_connect", True): disconnect(only_disconnect=True)
        else:
            disconnect()
            sess = os.path.join(app_data_path, "NordSessionData.json")
            if os.path.exists(sess): os.remove(sess)
    except: pass
    if pystray_icon: pystray_icon.stop()
    stop_app = True
def update_tray():
    if pystray_icon: pystray_icon.update_menu()
def about():
    def a(): ctypes.windll.user32.MessageBoxW(0, f"Version: v{version}\nMade by @EfazDev\nhttps://www.efaz.dev", "About NordWireConnect", 0x0 | 0x00000040)
    addToThread(a, ())
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
def select_server_list(message: str="Please select a server!") -> typing.Union[None, dict]:
    if not os.path.exists(os.path.join(app_data_path, "NordServerCache.json")):
        return
    with open(os.path.join(app_data_path, "NordServerCache.json"), "r") as f: raw_servers = json.load(f)
    server_dictionary = {
        "Auto": [("Auto", "auto", None)]
    }
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
    def on_select(title: str, tunnel_id: str, exact_mode: str):
        selected["title"] = title
        selected["tunnel_id"] = tunnel_id
        selected["exact_mode"] = exact_mode
        if title.lower() == "auto": dropdown_btn.config(text=tunnel_id)
        else: dropdown_btn.config(text=title)
    def build(parent_menu, data: typing.Union[dict, list, tuple]):
        if isinstance(data, dict):
            for title, value in sorted(data.items(), key=lambda x: ("\0" if x[0].lower() == "auto" else x[0])):
                submenu = tk.Menu(parent_menu, tearoff=0)
                parent_menu.add_cascade(label=title, menu=submenu)
                build(submenu, value)
        elif isinstance(data, list):
            data.sort(key=lambda x: ("\0" if x[0].lower() == "auto" else x[0]))
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
def connect_selected_server_button():
    selected_server = select_server_list("Please select a server to connect to!")
    if not selected_server:
        return
    addToThread(connect_server, (selected_server["tunnel_id"], selected_server["exact_mode"]), connection=True)
    update_tray()
def save_auto_selected_server_button():
    selected_server = select_server_list("Please select a server to save as a default location!")
    if not selected_server:
        return
    default_location(selected_server["tunnel_id"], selected_server["exact_mode"])
    update_tray()
def setup_shortcuts():
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
def delete_shortcuts():
    try:
        desk_shortpath = os.path.join(os.environ["USERPROFILE"], "Desktop", f"NordWireConnect.lnk")
        if os.path.exists(desk_shortpath): os.remove(desk_shortpath)
        startm_shortpath = os.path.join(os.getenv("APPDATA"), "Microsoft", "Windows", "Start Menu", "Programs", f"NordWireConnect.lnk")
        if os.path.exists(startm_shortpath): os.remove(startm_shortpath)
    except Exception as e: errorMessage(f"Unable to delete shortcuts: {str(e)}")
def delete_registry():
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
        tk_root.quit()
        sys.exit(0)
        return
    tk_root.after(100, check_stop_flag)
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
    mainMessage("Updating NordVPN server cache..")
    while not tunnel_check(timeout=1): time.sleep(0.1)
    all_servers = requests.get("https://api.nordvpn.com/v1/servers/recommendations?&filters\\[servers_technologies\\]\\[identifier\\]=wireguard_udp&limit=10000")
    raw_servers = (all_servers.json[:500] if (config_data.get("optimize_server_list", False) == True) else all_servers.json) if all_servers.ok else []
    raw_all_servers = all_servers.json if all_servers.ok else []
    all_cities = {}
    while not tunnel_check(timeout=1): time.sleep(0.1)
    all_countries = requests.get("https://api.nordvpn.com/v1/countries")
    all_countries = all_countries.json if all_countries.ok else []
    all_servers_by_id = {}
    for s in raw_all_servers:
        for l in s["locations"]:
            c = l["country"]["city"]
            all_cities[c["id"]] = c["name"]
        all_servers_by_id[s["name"].replace(" ", "").replace("#", "")] = s["id"]
    all_cities = [{"id": i, "name": n} for i, n in all_cities.items()]
    with open(os.path.join(app_data_path, "NordServerCache.json"), "w") as f: json.dump(raw_servers, f)
    with open(os.path.join(app_data_path, "NordCountryCache.json"), "w") as f: json.dump(all_countries, f)
    with open(os.path.join(app_data_path, "NordCityCache.json"), "w") as f: json.dump(all_cities, f)
    with open(os.path.join(app_data_path, "NordServerIDCache.json"), "w") as f: json.dump(all_servers_by_id, f)
def send_command(cmd: str) -> str:
    while True:
        try:
            win32pipe.WaitNamedPipe(service_pipe, 5000)
            handle = win32file.CreateFile(
                service_pipe,
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
def format_seconds(seconds: int) -> str:
    seconds = int(seconds)
    days, remaining_seconds = divmod(seconds, 86400)
    hours, remaining_seconds = divmod(remaining_seconds, 3600)
    minutes, seconds = divmod(remaining_seconds, 60)
    return f"{days:02d}:{hours:02d}:{minutes:02d}:{seconds:02d}"
def get_if_virtual_location(server: dict) -> bool:
    for sp in server.get("specifications", []):
        if sp["identifier"] == "virtual_location":
            for val in sp["values"]:
                if val["value"] == "true": return True
    return False
def change_dns():
    inputted = simpledialog.askstring(title="DNS Input", prompt="Enter IPv4 DNS servers to use (comma separated):")
    ips = inputted.split(", ")
    added = [ip for ip in ips if requests.get_if_ip(ip)]
    if added:
        config_data["dns"] = ", ".join(added)
        save_configuration()
        successMessageBox(f"Successfully changed DNS to: {', '.join(added)}", "NordWireConnect")
        update_tray()
    else: errorMessage("No valid IPv4 DNS servers were provided.")
def change_access_token():
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
def default_location(server_name: str, exact_mode: str=None): 
    if exact_mode: config_data["default_location"] = f"{exact_mode}_{server_name}"
    else: config_data["default_location"] = server_name
    save_configuration()
def mark_not_connected():
    session_data["session_time"] = 0
    session_data["connection_text"] = "Not Connected"
    session_data["connected"] = False
    session_data["server"] = None
    session_data_modified()
    update_tray()
def save_configuration():
    mainMessage("Saving configuration..")
    with open(os.path.join(app_data_path, "ConnectConfig.json"), "w") as f: json.dump(config_data, f, indent=4)

# Toggle Functions
def change_load_swapping():
    config_data["load_swapping"] = not config_data.get("load_swapping", False)
    save_configuration()
    update_tray()
def check_load_swapping(): return config_data.get("load_swapping", False) == True
def change_auto_connect():
    config_data["auto_connect"] = not config_data.get("auto_connect", True)
    save_configuration()
    update_tray()
def check_auto_connect(): return config_data.get("auto_connect", True) == True
def change_notifications():
    config_data["notifications"] = not config_data.get("notifications", False)
    save_configuration()
    update_tray()
def check_notifications(): return config_data.get("notifications", False) == True
def change_server_list():
    config_data["optimize_server_list"] = not config_data.get("optimize_server_list", False)
    addToThread(run_cache, ())
    save_configuration()
    update_tray()
def check_server_list(): return config_data.get("optimize_server_list", False) == True
def change_split_lan_routing():
    config_data["split_lan_routing"] = not config_data.get("split_lan_routing", False)
    save_configuration()
    if session_data["connected"] == True: differ_from_status_action2()
    update_tray()
def check_split_lan_routing(): return config_data.get("split_lan_routing", False) == True
def change_loss_connect_protection():
    config_data["loss_connect_protection"] = not config_data.get("loss_connect_protection", True)
    save_configuration()
    update_tray()
def check_loss_connect_protection(): return config_data.get("loss_connect_protection", True) == True

# Differenting Status
def differ_from_status_text() -> str:
    if session_data["connected"] == True: return "Disconnect from Server"
    elif session_data["connection_text"].startswith("Connecting"): return "Cancel Connection"
    else: return "Connect to Server"
def differ_from_status_action():
    if session_data["connection_text"].startswith("Connecting"):
        session_data["session_time"] = 0
        session_data["connection_text"] = "Not Connected"
        session_data["connected"] = False
        session_data["server"] = None
        set_icon("app_icon.ico")
    elif session_data["connected"] == True: 
        s = disconnect()
        if s == "0": successMessageBox("Successfully disconnected from NordVPN server!", "NordWireConnect")
    else: 
        if config_data.get("default_location", "auto") == "auto":
            config_data["server"] = "auto"
            addToThread(connect, (), connection=True)
        else:
            mode, server_loc = config_data.get("default_location").split("_")
            config_data["server"] = server_loc
            addToThread(connect, (mode,), connection=True)
def differ_from_status_action2():
    if session_data["connection_text"].startswith("Connecting") or session_data["connection_text"].startswith("Not"): return
    def pre_connect(): connect(exact_mode=session_data.get("exact_connect"))
    addToThread(pre_connect, (), connection=True)
def differ_to_status(): return session_data["connection_text"]
def differ_to_status1() -> str:
    if session_data["connected"] == True: return f"City: {session_data['server']['locations'][0]['country']['city']['name']}"
    else: return "City: N/A"
def differ_to_status2() -> str: 
    if session_data["connected"] == True: return f"IP Address: {session_data['server']['station']}"
    else: return "IP Address: N/A"
def differ_to_status3() -> str: 
    if session_data["connected"] == True: return f"Hostname: {session_data['server']['hostname']}"
    else: return "Hostname: N/A"
def differ_to_status4() -> str: 
    if session_data["connected"] == True: return f"Load: {session_data['current_load']}%"
    else: return "Load: N/A"
def differ_to_status5() -> str: 
    if session_data["connected"] == True: return f"Session Time: {format_seconds(session_data['session_time'])}"
    else: return "Session Time: N/A"
def differ_to_config1() -> str: return f"DNS Servers: {config_data['dns']}"
def differ_to_config2() -> str: return f"Account: {config_data.get('username', 'N/A')} (Token {len(config_data['access_token']) > 60 and 'Given' or 'Ungiven'})"
def differ_to_config3() -> str: 
    if config_data.get("default_location") and config_data["default_location"].lower() != "auto":
        _, server_loc = config_data["default_location"].split("_")
        return f"Default Location: {server_loc}"
    else: return "Default Location: Auto"

# Connections
def get_allowed_ips() -> str: 
    if config_data.get("split_lan_routing"): return "0.0.0.0/5, 8.0.0.0/7, 11.0.0.0/8, 12.0.0.0/6, 16.0.0.0/4, 32.0.0.0/3, 64.0.0.0/2, 128.0.0.0/3, 160.0.0.0/5, 168.0.0.0/6, 172.0.0.0/12, 172.32.0.0/11, 172.64.0.0/10, 172.128.0.0/9, 173.0.0.0/8, 174.0.0.0/7, 176.0.0.0/4, 192.0.0.0/9, 192.128.0.0/11, 192.160.0.0/13, 192.169.0.0/16, 192.170.0.0/15, 192.172.0.0/14, 192.176.0.0/12, 192.192.0.0/10, 193.0.0.0/8, 194.0.0.0/7, 196.0.0.0/6, 200.0.0.0/5, 208.0.0.0/4, ::/1, 8000::/2, c000::/3, e000::/4, f000::/5, f800::/6, fe00::/9, ff00::/8"
    else: return "0.0.0.0/0, ::/0"
def connect_server(server_name: str, exact_mode: str=None):
    config_data["server"] = server_name
    try: connect(exact_mode=exact_mode)
    except Exception: pass
def get_client_public_key(private_key: str) -> typing.Tuple[str, str]:
    pub_key_path = os.path.join(app_data_path, "NordPublicKey.json")
    if not os.path.exists(pub_key_path):
        pub_key = send_command(f"generate-public-key {private_key}")
        preshared = send_command(f"generate-preshared")
        with open(pub_key_path, "w") as f: json.dump({"pub": pub_key, "pre": preshared}, f)
    else:
        with open(pub_key_path, "r") as f: public_key_json = json.load(f)
        pub_key = public_key_json.get("pub")
        preshared = public_key_json.get("pre")
    return pub_key, preshared
def connect_session():
    global session_data
    if session_data["connected"] == True and config_data.get("auto_connect", True) == True:
        mainMessage("Restoring previous session...")
        org = session_data.get("org_server_id", "auto")
        org_exact_connect = session_data.get("exact_connect")
        connect_server(session_data["server"]["name"].replace(" ", "").replace("#", ""), exact_mode="server")
        config_data["server"] = org
        session_data["exact_connect"] = org_exact_connect
    elif session_data["connection_text"].startswith("Connecting"):
        mainMessage("Reconnecting..")
        session_data["connected"] = False
        session_data["connection_text"] = "Not Connected"
        config_data["server"] = None
        session_data["server"] = None
        session_data["current_load"] = None
        session_data["exact_connect"] = None
        differ_from_status_action()
    elif send_command("connection-status") != "NotConnected": disconnect()
def connect(exact_mode: str=None):
    global session_data
    try:
        session_data["session_time"] = 0
        session_data["connected"] = False
        session_data["connection_text"] = "Connecting..."
        session_data_modified()
        set_icon("awaiting.ico")

        if not private_key:
            errorMessage("Please provide your Access Token before trying to connect!")
            mark_not_connected()
            set_icon("error.ico")
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
                send_command("combined-end-wireguard")
                time.sleep(1)
                if os.path.exists(installer_path): os.remove(installer_path)
                if wireguard_exit_code == 0:
                    mainMessage("WireGuard installed successfully.")
                    time.sleep(3)
                else:
                    errorMessage("Failed to install WireGuard.")
                    mark_not_connected()
                    set_icon("error.ico")
                    return
            else:
                errorMessage("Failed to download WireGuard installer.")
                mark_not_connected()
                set_icon("error.ico")
                return

        # Load NordVPN Server Recommendations
        disconnect(only_disconnect=True)
        if not tunnel_check():
            if config_data.get("auto_connect", True) == True:
                errorMessage("There was an issue trying to connect to your internet. Press Cancel Connection if you want to disconnect.")
                set_icon("warning.ico")
                while not tunnel_check(timeout=1):
                    if session_data["connection_text"] == "Not Connected" and session_data["connected"] == False: break
                    time.sleep(1)
                if session_data["connection_text"] == "Not Connected" and session_data["connected"] == False:
                    mainMessage("Canceled connection.")
                    return
                set_icon("awaiting.ico")
            else: 
                errorMessage("There was an issue trying to connect to your internet.")
                set_icon("error.ico")
                return
        if config_data["server"] == "auto" or config_data["server"] == None or exact_mode == None:
            all_server_recommendations = requests.get("https://api.nordvpn.com/v1/servers/recommendations?&filters\\[servers_technologies\\]\\[identifier\\]=wireguard_udp&limit=100")
            if not all_server_recommendations.ok:
                errorMessage("Unable to get NordVPN server recommendations.")
                mark_not_connected()
                set_icon("error.ico")
                return
            all_server_recommendations = all_server_recommendations.json
        elif exact_mode == "country":
            country_list_path = os.path.join(app_data_path, "NordCountryCache.json")
            if not os.path.exists(country_list_path):
                errorMessage("Unable to find NordVPN country list.")
                mark_not_connected()
                set_icon("error.ico")
                return
            with open(country_list_path, "r") as f: country_list = json.load(f)
            country_id = None
            for c in country_list:
                if c.get("name") == config_data["server"]:
                    country_id = c.get("id")
                    break
            if country_id == None:
                errorMessage("Unable to get NordVPN country ID.")
                mark_not_connected()
                set_icon("error.ico")
                return
            all_server_recommendations = requests.get(f"https://api.nordvpn.com/v1/servers/recommendations?&filters[country_id]={country_id}&filters\\[servers_technologies\\]\\[identifier\\]=wireguard_udp&limit=100")
            if not all_server_recommendations.ok:
                errorMessage("Unable to get NordVPN server recommendations.")
                mark_not_connected()
                set_icon("error.ico")
                return
            all_server_recommendations = all_server_recommendations.json
        elif exact_mode == "city":
            city_list_path = os.path.join(app_data_path, "NordCityCache.json")
            if not os.path.exists(city_list_path):
                errorMessage("Unable to find NordVPN city list.")
                mark_not_connected()
                set_icon("error.ico")
                return
            with open(city_list_path, "r") as f: city_list = json.load(f)
            city_id = None
            for c in city_list:
                if c.get("name") == config_data["server"]:
                    city_id = c.get("id")
                    break
            if city_id == None:
                errorMessage("Unable to get NordVPN city ID.")
                mark_not_connected()
                set_icon("error.ico")
                return
            all_server_recommendations = requests.get(f"https://api.nordvpn.com/v1/servers/recommendations?&filters[country_city_id]={city_id}&filters\\[servers_technologies\\]\\[identifier\\]=wireguard_udp&limit=100")
            if not all_server_recommendations.ok:
                errorMessage("Unable to get NordVPN server recommendations.")
                mark_not_connected()
                set_icon("error.ico")
                return
            all_server_recommendations = all_server_recommendations.json
        elif exact_mode == "server":
            server_list_path = os.path.join(app_data_path, "NordServerIDCache.json")
            if not os.path.exists(server_list_path):
                errorMessage("Unable to find NordVPN server list.")
                mark_not_connected()
                set_icon("error.ico")
                return
            with open(server_list_path, "r") as f: server_id_list = json.load(f)
            if server_id_list.get(config_data["server"]):
                server_id = server_id_list.get(config_data["server"])
                all_server_recommendations = requests.get(f"https://api.nordvpn.com/v1/servers?&filters[servers.id]={server_id}&limit=1")
                if not all_server_recommendations.ok:
                    errorMessage("Unable to get NordVPN server recommendations.")
                    mark_not_connected()
                    set_icon("error.ico")
                    return
                all_server_recommendations = all_server_recommendations.json
            else:
                errorMessage("Unable to find NordVPN server.")
                mark_not_connected()
                set_icon("error.ico")
                return
        else:
            all_server_recommendations = requests.get("https://api.nordvpn.com/v1/servers/recommendations?&filters\\[servers_technologies\\]\\[identifier\\]=wireguard_udp&limit=100")
            if not all_server_recommendations.ok:
                errorMessage("Unable to get NordVPN server recommendations.")
                mark_not_connected()
                set_icon("error.ico")
                return
            all_server_recommendations = all_server_recommendations.json
        mainMessage(f"Loaded {len(all_server_recommendations)} NordVPN server recommendation(s).")

        # Start Finding Servers
        mainMessage(f"Connecting to NordVPN Servers in {config_data['server']}..")
        connected_server = None
        attempts = 0
        for s in all_server_recommendations:
            shortened_name = s["name"].replace(" ", "").replace("#", "")
            country_name = s["locations"][0]["country"]["name"].replace(" ", "")
            city_name = s["locations"][0]["country"]["city"]["name"].replace(" ", "")
            if not (config_data["server"] == "auto" or config_data["server"] == None) and not (shortened_name == config_data["server"].replace(" ", "").replace("#", "") or city_name == config_data["server"].replace(" ", "") or country_name == config_data["server"].replace(" ", "")):
                continue
            if session_data["connection_text"] == "Not Connected" and session_data["connected"] == False: break
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
            # client_public, preshared = get_client_public_key(private_key)
            configuration = f"""# {shortened_name}.{city_name}.conf
[Interface]
PrivateKey = {private_key}
Address = 10.5.0.2/32
DNS = {config_data.get('dns')}
MTU = 1380

[Peer]
PublicKey = {wireguard_metadata[0].get("value")}
AllowedIPs = {get_allowed_ips()}
Endpoint = {s.get('hostname')}:51820
PersistentKeepalive = 25"""
            config_path = os.path.join(app_data_path, f"{shortened_name}.{city_name}.conf")
            with open(config_path, "w") as f: f.write(configuration)
            if session_data["connection_text"] == "Not Connected" and session_data["connected"] == False: break
            add = send_command(f"reinstall-wire-tunnel {shortened_name}.{city_name} {config_path}")
            attempts += 1
            if add == "0":
                mainMessage(f"Added new WireGuard configuration for server: {s['name']}")
                tunnel_check(timeout=3)
                time.sleep(1)
                if tunnel_check(timeout=3):
                    connected_server = s
                    break
                else: 
                    send_command(f"uninstall-wire-tunnel {shortened_name}.{city_name}")
                    disconnect(only_disconnect=True)
                    continue
            else:
                errorMessage(f"Failed to add WireGuard configuration for server: {s['name']}")
                continue
        if session_data["connection_text"] == "Not Connected" and session_data["connected"] == False: return
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
            session_data["org_server_id"] = config_data["server"]
            session_data_modified()
            successMessageBox(f"Successfully connected to server: {connected_server['name']}!", "Yay!")
            set_icon("connected.ico")
            needing_state()
        else:
            errorMessage("Unable to connect to a NordVPN server.")
            mark_not_connected()
            set_icon("error.ico")
            return
    except Exception as e:
        errorMessage(f"Unable to connect to a NordVPN server due to an error: {str(e)}.")
        mark_not_connected()
        set_icon("error.ico")
        return
def disconnect(only_disconnect: bool=False) -> str:
    mainMessage("Disconnecting Wireguard..")
    awared = None
    if session_data["connected"] == True:
        server_name = session_data["server"]["name"].replace(" ", "").replace("#", "")
        awared = server_name
    if only_disconnect == False:
        session_data["session_time"] = 0
        session_data["connection_text"] = "Not Connected"
        session_data["connected"] = False
        session_data["server"] = None
        session_data_modified()
        set_icon("app_icon.ico")
        update_tray()
    if awared: return send_command(f"combined-disconnect {awared}")
    else: return send_command("combined-disconnect")
def brute_end_wireguard():
    disconnect()
    mark_not_connected()
def handle_stat_thread():
    count = 0
    since_connected = 0
    while True:
        time.sleep(1)
        try:
            if session_data["connected"] == True:
                count += 1
                since_connected += 1
                if count % 10 == 0 and config_data.get("loss_connect_protection", True) == True and not tunnel_check() and not session_data["connection_text"].startswith("Connecting"):
                    mainMessage("Reconnecting due to loss of connection.")
                    differ_from_status_action2()
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
                        differ_from_status_action2()
                session_data["session_time"] = since_connected
                update_tray()
            else:
                count = 0
                since_connected = 0
        except Exception: pass
def tunnel_check(timeout: int=5): return requests.get_if_connected(server="1.1.1.1", timeout=timeout)

# Main Runtime
def app():
    global config_data
    global pystray_icon
    global private_key
    global session_data
    colors_class.fix_windows_ansi()
    setup_logging()
    systemMessage(f"{'-'*5:^5} NordWireConnect v{version} {'-'*5:^5}")
    tk_root.withdraw()

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

    # Clear Old Versions
    was_old = send_command("cleared-older-version") == "0"
    if was_old == True: setup_shortcuts()

    # Prepare NordWireConnect
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
        if not tunnel_check():
            disconnect(only_disconnect=True)
            while not tunnel_check(timeout=1): time.sleep(0.1)
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
        addToThread(run_cache, ())
    except Exception as e: errorMessage(f"Unable to load server list. Exception: {str(e)}"); sys.exit(1); return

    # Create Tray App
    try:
        mainMessage("Creating Tray App..")
        def create_mini_func(func: typing.Callable, connected: bool=False):
            def a(icon: Icon, item: MenuItem): addToThread(func, (), connection=False)
            return a
        image = getImageObj(os.path.join(program_files, "resources", "app_icon.ico"))
        pystray_icon = pystray.Icon(
            "NordWireConnect",
            image,
            "NordWireConnect",
            menu=pystray.Menu(
                pystray.MenuItem("NordWireConnect", lambda icon, item: None, enabled=False),
                pystray.MenuItem(unicon(differ_to_status), lambda icon, item: None, enabled=False),
                pystray.MenuItem(unicon(differ_to_status1), lambda icon, item: None, enabled=False, visible=lambda icon: session_data["connected"] == True),
                pystray.MenuItem(unicon(differ_to_status2), lambda icon, item: None, enabled=False, visible=lambda icon: session_data["connected"] == True),
                pystray.MenuItem(unicon(differ_to_status3), lambda icon, item: None, enabled=False, visible=lambda icon: session_data["connected"] == True),
                pystray.MenuItem(unicon(differ_to_status4), lambda icon, item: None, enabled=False, visible=lambda icon: session_data["connected"] == True),
                pystray.MenuItem(unicon(differ_to_status5), lambda icon, item: None, enabled=False, visible=lambda icon: session_data["connected"] == True),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem(unicon(differ_from_status_text), create_mini_func(differ_from_status_action, connected=True)),
                pystray.MenuItem("Reconnect Servers", create_mini_func(differ_from_status_action2, connected=True), visible=lambda icon: session_data["connected"] == True),
                pystray.MenuItem("Servers", lambda icon, item: tk_root.after(0, connect_selected_server_button)),
                pystray.MenuItem("Configuration", pystray.Menu(
                    pystray.MenuItem(unicon(differ_to_config1), lambda icon, item: None, enabled=False),
                    pystray.MenuItem(unicon(differ_to_config2), lambda icon, item: None, enabled=False),
                    pystray.Menu.SEPARATOR,
                    pystray.MenuItem("Change DNS Servers", lambda icon, item: tk_root.after(0, change_dns)),
                    pystray.MenuItem("Change Access Token", lambda icon, item: tk_root.after(0, change_access_token)),
                    pystray.MenuItem(unicon(differ_to_config3), lambda icon, item: tk_root.after(0, save_auto_selected_server_button)),
                    pystray.MenuItem("Set Shortcuts", create_mini_func(setup_shortcuts)),
                    pystray.MenuItem("Brute End Wireguard", create_mini_func(brute_end_wireguard)),
                    pystray.MenuItem("Load Swapping", unicon(change_load_swapping), checked=unicon(check_load_swapping), radio=True),
                    pystray.MenuItem("Auto Reconnect", unicon(change_auto_connect), checked=unicon(check_auto_connect), radio=True),
                    pystray.MenuItem("Notifications", unicon(change_notifications), checked=unicon(check_notifications), radio=True),
                    pystray.MenuItem("Optimize Server List", unicon(change_server_list), checked=unicon(check_server_list), radio=True),
                    pystray.MenuItem("Split Private LAN Routing", unicon(change_split_lan_routing), checked=unicon(check_split_lan_routing), radio=True),
                    pystray.MenuItem("Loss Connect Protection", unicon(change_loss_connect_protection), checked=unicon(check_loss_connect_protection), radio=True),
                )),
                pystray.MenuItem("About NordWireConnect", unicon(about)),
                pystray.MenuItem("Quit", unicon(quit_app))
            )
        )
    except Exception as e: errorMessage(f"Unable to run app. Exception: {str(e)}"); sys.exit(1); return

    # Run Main Loop
    try:
        mainMessage("Finishing app load!")
        addToThread(connect_session, (), connection=True)
        threading.Thread(target=worker, daemon=True).start()
        threading.Thread(target=handle_stat_thread, daemon=True).start()
        tk_root.after(100, check_stop_flag)
        tk_root.iconphoto(True, tk.PhotoImage(file=os.path.join(program_files, "resources", "app_icon.png"))) 
        pystray_icon.run_detached(unicon(setup))
        tk.mainloop()
    except Exception as e: errorMessage(f"Unable to run app. Exception: {str(e)}"); sys.exit(1); return
if __name__ == "__main__": app()