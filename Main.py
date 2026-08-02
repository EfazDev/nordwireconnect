"""
NordWireConnect

Made by @EfazDev
https://www.efaz.dev
"""

# Modules
import os
import sys
import copy
import json
import time
import zlib
import bisect
import typing
import PyKits
import System
import pystray # type: ignore
import logging
import datetime
import platform
import threading
import traceback
import ipaddress
import functools
import webbrowser
import tkinter as tk
from PIL import Image # type: ignore
from queue import Queue
from tkinter import simpledialog, messagebox

# Variables
pip_class = PyKits.pip()
colors_class = PyKits.Colors()
nord_socket = PyKits.Socket(port=61245)
requests = PyKits.request(throw_exceptions=False)
main_os = platform.system()
work_queue = Queue()
ui_queue = Queue()
system_class = System.system
plan_timer = 0
stat_count = 0
since_connected = 0
pre_program_files = system_class.program_files
app_data_path = system_class.app_data_path
wireguard_location = os.path.join(pre_program_files, "WireGuard")
program_files = os.path.join(pre_program_files, "NordWireConnect")
one_instance_lock = None
private_key = None
plan_reminded = False
data_reminded = False
debug_mode = False
base_config = {
    "access_token": "",
    "username": "",
    "openvpn_username": "",
    "openvpn_password": "",
    "nordvpn_email": "",
    "persistent_keepalive": 10,
    "last_version": None,
    "dns": "103.86.96.100, 103.86.99.100", # NordVPN DNS
    "server": "auto",
    "auto_connect": True,
    "notifications": True,
    "optimize_server_list": False,
    "default_location": "auto",
    "load_swapping": False,
    "load_swapping_threshold": 50,
    "split_lan_routing": False,
    "loss_connect_protection": True,
    "threat_protection_lite": False,
    "track_servers": True,
    "server_history": None,
    "asn_server_list": False,
    "plan_renewal_reminder": False,
    "plan_renewal_data": None,
    "plan_reminder_level": 5,
    "monthly_data_usage": None,
    "total_data_usage": None,
    "data_usage_reminder": False,
    "usage_reminder_level": 3,
    "connection_channel": None,
    "check_for_updates": True,
    "data_encrypted": False,
    "beta_updates": False,
    "debug_updates": False
}
config_data = copy.deepcopy(base_config)
config_data_type_allowed = {
    "access_token": str,
    "username": str,
    "openvpn_username": str,
    "openvpn_password": str,
    "nordvpn_email": str,
    "persistent_keepalive": int,
    "dns": str,
    "last_version": typing.Union[str, None],
    "server": str,
    "auto_connect": bool,
    "notifications": bool,
    "optimize_server_list": bool,
    "default_location": str,
    "load_swapping": bool,
    "load_swapping_threshold": int,
    "split_lan_routing": bool,
    "asn_server_list": bool,
    "loss_connect_protection": bool,
    "threat_protection_lite": bool,
    "track_servers": bool,
    "server_history": typing.Union[list, None],
    "plan_renewal_reminder": bool,
    "plan_renewal_data": typing.Union[dict, None],
    "plan_reminder_level": int,
    "monthly_data_usage": typing.Union[dict, None],
    "total_data_usage": typing.Union[dict, None],
    "data_usage_reminder": bool,
    "usage_reminder_level": int,
    "connection_channel": typing.Union[int, None],
    "data_encrypted": bool,
    "check_for_updates": bool,
    "beta_updates": bool,
    "debug_updates": bool
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
add_data_count = 0
version = "1.3.5j"
cached_asn_data = {"loaded": False, "names": [], "starts": [], "ends": []}
service_pipe = r"\\.\pipe\NordWireConnect"
tk_root = None
Icon = pystray.Icon
MenuItem = pystray.MenuItem
server_filter = "&filters[servers.status]=online&fields[servers.id]&fields[servers.name]&fields[servers.hostname]&fields[servers.station]&fields[servers.status]&fields[servers.load]&fields[servers.created_at]&fields[servers.groups.id]&fields[servers.groups.title]&fields[servers.technologies.id]&fields[servers.technologies.metadata]&fields[servers.technologies.pivot.status]&fields[servers.specifications.identifier]&fields[servers.specifications.values.value]&fields[servers.locations.country.name]&fields[servers.locations.country.code]&fields[servers.locations.country.city.name]&fields[servers.locations.country.city.latitude]&fields[servers.locations.country.city.longitude]&fields[servers.locations.country.city.hub_score]&fields[servers.ips]"

# Logging and Messages
def systemMessage(message: str): colors_class.print(message, colors_class.hex_to_ansi2("#3E5FFF"))
def mainMessage(message: str): colors_class.print(message, 15)
def errorMessage(message: str, title: str="Uh oh!"): 
    colors_class.print(message, 9)
    message = message[:256]
    if config_data.get("notifications", True): notification(title, message)
    else:
        if title == "": title = "NordWireConnect"
        elif title != "NordWireConnect": title = f"NordWireConnect: {title}"
        def a(): messagebox.showerror(title, message)
        tk_root.after(0, a)
def warnMessage(message: str): colors_class.print(message, 11)
def successMessage(message: str): colors_class.print(message, 10)
def successMessageBox(message: str, title: str="Success!"):
    message = message[:256]
    if config_data.get("notifications", True): notification(title, message)
    else:
        if title == "": title = "NordWireConnect"
        elif title != "NordWireConnect": title = f"NordWireConnect: {title}"
        def a(): messagebox.showinfo(title, message)
        tk_root.after(0, a)
def trace(e: Exception):
    if not debug_mode: return str(e)
    _, tb_v, tb_b = sys.exc_info()
    tb_lines = traceback.extract_tb(tb_b)
    lines = []
    lines.append(colors_class.foreground("Traceback (most recent call last):", color="Magenta", bright=True))
    for fn, ln, f, tx in tb_lines:
        lines.append(f'  File {colors_class.foreground(fn, color="Magenta", bright=True)}, line {colors_class.foreground(ln, color="Magenta", bright=True)}, in {colors_class.foreground(f, color="Magenta", bright=True)}')
        if tx: lines.append(f'    {tx}')
    exc_t = type(tb_v).__name__
    exc_m = str(tb_v)
    lines.append(f'{colors_class.foreground(colors_class.bold(f"{exc_t}:"), color="Magenta", bright=True)} {colors_class.foreground(exc_m, color="Magenta", bright=False)}')
    return "\n".join(lines)
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
    System.systemMessage = systemMessage
    System.mainMessage = mainMessage
    System.successMessage = successMessage
    System.warnMessage = warnMessage 
    System.successMessageBox = successMessageBox
    System.errorMessage = errorMessage

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
def notification(title: str, message: str):
    def a():
        if pystray_icon: pystray_icon.notify(message, title)
    addToThread(a, ())
def set_icon(img: str):
    image = getImageObj(os.path.join(program_files, "Resources", img))
    if pystray_icon:
        pystray_icon.icon = image
        update_tray()
    if img != "connected.ico": unneeding_state()
def setup(): pystray_icon.visible = True
def needing_state(): system_class.needing_state()
def unneeding_state(): system_class.unneeding_state()
def quit_app():
    global stop_app
    try:
        if config_data.get("auto_connect", True): 
            session_data["connected"] = False
            disconnect(only_disconnect=True)
        else:
            disconnect()
            sess = os.path.join(app_data_path, "NordSessionData.json")
            if os.path.exists(sess): os.remove(sess)
    except: pass
    if pystray_icon: pystray_icon.stop()
    stop_app = True
def update_tray():
    if pystray_icon and main_os == "Windows": pystray_icon.update_menu()
def about():
    def a(): messagebox.showinfo("About NordWireConnect", f"Version: v{version}\nMade by @EfazDev\nhttps://www.efaz.dev")
    addToThread(a, ())
def relaunch_as_admin(): system_class.relaunch_as_admin()
def center_tkinter_window(window: tk.Toplevel):
    window.update_idletasks() 
    x = (window.winfo_screenwidth() // 2) - (window.winfo_width() // 2)
    y = (window.winfo_screenheight() // 2) - (window.winfo_height() // 2)
    window.geometry(f'+{x}+{y}')
def encrypt_data(key_name: str, data: str) -> str: return system_class.encrypt_data(key_name, data)
def decode_data(key_name: str, encrypted: str) -> str: return system_class.decrypt_data(key_name, encrypted)
def select_server_list(message: str="Please select a server!", servers: list=None) -> typing.Union[None, dict]:
    if not os.path.exists(os.path.join(app_data_path, "NordServerCache.json")):
        return
    with open(os.path.join(app_data_path, "NordServerCache.json"), "r") as f: raw_servers = json.load(f)
    if servers == None:
        server_dictionary = {
            "Auto": [("Auto", "auto", None)]
        }
        for s in raw_servers:
            if s["status"] != "online" or s["load"] == 0:
                continue
            virutal_location = get_if_virtual_location(s)
            wireguard_metadata = None
            for m in s.get("technologies", []):
                if m.get("id") == 35:
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
            if check_asn_server_list():
                if city_with_virtual_text not in server_dictionary[country_name]:
                    server_dictionary[country_name][city_with_virtual_text] = {
                        "Auto": [("Auto", city_name, "city")]
                    }
                server_ip = s.get("station")
                asn_name = get_asn_name_from_ip(server_ip)
                if not asn_name: asn_name = "Unknown ASN"
                if asn_name not in server_dictionary[country_name][city_with_virtual_text]:
                    server_dictionary[country_name][city_with_virtual_text][asn_name] = [
                        ("Auto", f"{city_name}|||{asn_name}", "asn")
                    ]
                server_dictionary[country_name][city_with_virtual_text][asn_name].append((server_name, server_name.replace(" ", "").replace("#", ""), "server"))
            else:
                if city_with_virtual_text not in server_dictionary[country_name]:
                    server_dictionary[country_name][city_with_virtual_text] = [
                        ("Auto", city_name, "city")
                    ]
                server_dictionary[country_name][city_with_virtual_text].append(
                    (server_name, server_name.replace(" ", "").replace("#", ""), "server")
                )
    else: server_dictionary = servers
    win = tk.Toplevel()
    win.title("Select NordVPN Server")
    win.geometry("320x120")
    win.resizable(False, False)
    win.attributes("-topmost", True)
    center_tkinter_window(win)
    win.iconphoto(False, tk.PhotoImage(file=os.path.join(program_files, "Resources", "app_icon.png")))
    selected = {
        "title": None,
        "tunnel_id": None,
        "exact_mode": None
    }
    result = {"value": None}
    label = tk.Label(win, textvariable=tk.StringVar(value=message), anchor="w", padx=10, pady=8)
    label.pack(fill="x")
    rs_f = tk.Frame(win, padx=10)
    rs_f.pack(fill="x")
    rows = []
    def reflow(): win.update_idletasks(); win.geometry("")
    def add_row(c):
        frame = tk.Frame(rs_f)
        frame.pack(fill="x", pady=2)
        num_of_rows = len(rows)
        btn = tk.Button(frame, text="Select Option", anchor="w", relief="groove", width=40, fg="gray")
        btn.pack(side="left")
        def_fg = btn.cget("fg")
        def reset_btn(i: int=num_of_rows, children: list=c):
            for row in rows[i:]: row["frame"].destroy()
            del rows[i:]
            if i > 0 and rows[i-1].get("committed"): selected.update(rows[i-1]["committed"])
            else: selected.update({"title": None, "tunnel_id": None, "exact_mode": None})
            add_row(children)
        rst = tk.Button(frame, text="✕", width=2, relief="flat", fg="gray", command=reset_btn)
        rst.pack(side="left", padx=(2, 0))
        def pick_handler(node: dict, i: int=num_of_rows):
            for row in rows[i+1:]: row["frame"].destroy()
            del rows[i+1:]
            rows[i]["btn"].config(text=node["label"], fg=def_fg, relief="sunken")
            committed = {"title": node["label"], "tunnel_id": node["tunnel_id"], "exact_mode": node["exact_mode"]}
            rows[i]["committed"] = committed
            selected.update(committed)
            if node["children"]: add_row(node["children"])
            else: reflow()
        def show_menu(b: tk.Button=btn, ch: list=c):
            m = tk.Menu(win, tearoff=0)
            for node in ch: m.add_command(label=node["label"], command=lambda n=node: pick_handler(n))
            x = b.winfo_rootx()
            y = b.winfo_rooty() + b.winfo_height()
            m.tk_popup(x, y)
        btn.config(command=show_menu)
        rows.append({"frame": frame, "btn": btn, "committed": None})
        reflow()
    def format_items(data: typing.Union[list, dict], root: bool=True):
        nodes = []
        if type(data) is list:
            for lbl, tid, em in data:
                nodes.append({"label": lbl, "tunnel_id": tid, "exact_mode": em, "children": None})
        elif type(data) is dict:
            for key in sorted(data.keys(), key=lambda x: ("\0" if x.lower() == "auto" else x)):
                val = data[key]
                if key.lower() == "auto":
                    for label, tunnel, exact in val: nodes.append({"label": label, "tunnel_id": tunnel if tunnel else label, "exact_mode": exact, "children": None})
                    continue
                if type(val) is list: info = val[0]
                elif type(val) is dict: info = val.get("Auto", [("Auto", key, "country")])[0]
                chi = format_items(val, root=False)
                nodes.append({"label": key, "tunnel_id": info[1], "exact_mode": info[2], "children": chi if chi else None})
        return nodes
    add_row(format_items(server_dictionary))
    btn_frame = tk.Frame(win)
    btn_frame.pack(side="bottom", pady=6)
    def confirm():
        if selected["tunnel_id"] != None: result["value"] = dict(selected)
        win.destroy()
    def cancel(): win.destroy()
    tk.Button(btn_frame, text="Confirm", width=10, command=confirm).pack(side="left", padx=6)
    tk.Button(btn_frame, text="Cancel",  width=10, command=cancel).pack(side="left", padx=6)
    win.protocol("WM_DELETE_WINDOW", cancel)
    reflow()
    win.grab_set()
    win.wait_window()
    return result["value"]
def connect_selected_server_button():
    selected_server = select_server_list("Please select a server to connect to!")
    if not selected_server:
        return
    addToThread(connect_server, (selected_server["tunnel_id"], selected_server["exact_mode"]), connection=True)
    update_tray()
def connect_history_server_button():
    generated_servers = {}
    if not config_data.get("server_history"): config_data["server_history"] = []
    for s in config_data.get("server_history", []):
        country = s.get("country")
        city = s.get("city")
        server_name = s.get("name")
        date = s.get("time")
        dt = datetime.datetime.fromtimestamp(date).astimezone()
        if not country or not city or not server_name: continue
        day = convert_to_local(dt).strftime("%Y-%m-%d")
        time = convert_to_local(dt).strftime("%H:%M")
        if day not in generated_servers: generated_servers[day] = {}
        if time not in generated_servers[day]: generated_servers[day][time] = []
        generated_servers[day][time].insert(0, (f"{server_name} ({city}, {country})", server_name.replace(" ", "").replace("#", ""), "server"))
    selected_server = select_server_list("Please select a server from history to connect to!", servers=generated_servers)
    if not selected_server: return
    addToThread(connect_server, (selected_server["tunnel_id"], selected_server["exact_mode"]), connection=True)
    update_tray()
def save_auto_selected_server_button():
    selected_server = select_server_list("Please select a server to save as a default location!")
    if not selected_server: return
    default_location(selected_server["tunnel_id"], selected_server["exact_mode"])
    update_tray()
def setup_shortcuts(button: bool=False): return system_class.setup_shortcuts(button=button)
def tkinter_clear_servers_history():
    try:
        confirmed = messagebox.askyesno("Clear Server Connected History", "Are you sure you want to clear the server connected history? This cannot be undone.")
        if confirmed:
            config_data["server_history"] = None
            save_configuration()
            update_tray()
            successMessageBox("Successfully cleared server connected history!", "NordWireConnect")
    except Exception as e: errorMessage(f"Unable to clear server connected history: {trace(e)}")
def clear_server_connected_history(): tk_root.after(0, tkinter_clear_servers_history)
def delete_shortcuts(): system_class.delete_shortcuts()
def delete_registry(): system_class.delete_registry()
def uninstall_app(): system_class.uninstall_app()
def check_stop_flag():
    if stop_app:
        mainMessage("Exiting app..")
        send_command("ui-closing")
        if one_instance_lock and one_instance_lock.exists(): one_instance_lock.release()
        tk_root.quit()
        sys.exit(0)
        return
    tk_root.after(100, check_stop_flag)
def worker():
    while True:
        func, args = work_queue.get()
        try: func(*args)
        except Exception as e: errorMessage(f"Error while running worker ({getattr(func, "__name__")}): {trace(e)}")
        finally: work_queue.task_done()
def get_latest_version() -> typing.Union[dict, None]:
    try:
        if tunnel_check(timeout=1):
            latest_version = requests.get("https://raw.githubusercontent.com/EfazDev/nordwireconnect/refs/heads/main/Version.json")
            if latest_version.ok: 
                res = latest_version.json
                for v in res.get("versions", []):
                    if v.get("version") and v.get("beta", False) == config_data.get("beta_updates", False) and v.get("debug", False) == config_data.get("debug_updates", False): return v
    except Exception as e: errorMessage(f"Error while trying to check for updates: {trace(e)}")
def get_latest_wireguard_version() -> dict:
    try:
        if tunnel_check(timeout=1):
            latest_version = requests.get("https://build.wireguard.com/distros.json")
            if latest_version.ok: 
                res = latest_version.json
                return system_class.format_for_latest_wireguard(res)
    except Exception as e: errorMessage(f"Error while trying to check for updates: {trace(e)}")
def get_current_wireguard_version() -> typing.Union[str, None]: return system_class.get_current_wireguard_version()
def auto_check_for_updates():
    if session_data.get("connection_text", "").startswith("Connecting"):
        while session_data.get("connection_text", "").startswith("Connecting"): time.sleep(0.1)
    wait_until_connected()
    latest_version = get_latest_version()
    if latest_version and latest_version["version"] > version: 
        notification("Update Available!", f"NordWireConnect v{latest_version['version']} is now available for download!")
    def wireguard_update():
        latest_wireguard_version = get_latest_wireguard_version()
        current_wireguard_version = get_current_wireguard_version()
        if not latest_wireguard_version: return
        if not current_wireguard_version: return
        if latest_wireguard_version["version"] > current_wireguard_version: 
            notification("Wireguard Update Available!", f"Wireguard v{latest_wireguard_version['version']} is now available for download! It is suggested to update for security.")
    wireguard_update()
def tkinter_install_wireguard_updates(latest_ver: str):
    try:
        confirmed = messagebox.askyesno("Wireguard Update Available!", f"Wireguard v{latest_ver['version']} is now available to be downloaded!\nDo you wish to proceed?", icon="info")
        if confirmed:
            response = system_class.install_wireguard(send_command)
            if not response:
                errorMessage("Failed to download WireGuard installer.")
                mark_not_connected()
                set_icon("error.ico")
                return
    except Exception as e: errorMessage(f"Unable to check for new updates due to an exception: {trace(e)}")
def tkinter_install_updates(latest_ver: str):
    try:
        mainMessage(f"Update v{latest_ver} is now available!")
        confirmed = messagebox.askyesno("Update Available!", f"NordWireConnect v{latest_ver} is now available to be downloaded from GitHub!\nDo you wish to proceed?", icon="info")
        if confirmed:
            installer_path = os.path.join(app_data_path, "NordWireConnectInstaller.exe")
            download_progress = requests.download(f"https://github.com/EfazDev/nordwireconnect/releases/download/v{latest_ver}/{'NordWireConnectDebugInstaller.exe' if config_data.get('debug_updates', False) else 'NordWireConnectInstaller.exe'}", installer_path)
            if download_progress.ok:
                if config_data.get("auto_connect", True): 
                    session_data["connected"] = False
                    disconnect(only_disconnect=True)
                else:
                    disconnect()
                    sess = os.path.join(app_data_path, "NordSessionData.json")
                    if os.path.exists(sess): os.remove(sess)
                if pystray_icon: pystray_icon.stop()
                system_class.run_sudo_cmd([installer_path])
            else: 
                if os.path.exists(installer_path): os.remove(installer_path)
                errorMessage("Unable to download NordWireConnectInstaller from GitHub!")
    except Exception as e: errorMessage(f"Unable to check for new updates due to an exception: {trace(e)}")
def check_for_updates():
    try:
        latest_version = get_latest_version()
        if not latest_version:
            errorMessage("Unable to check for new updates! Check your internet connection..")
            return
        if latest_version["version"] > version or debug_mode != latest_version.get("debug", False): 
            latest_ver = latest_version["version"]
            test_available = requests.head(f"https://github.com/EfazDev/nordwireconnect/releases/download/v{latest_ver}/{'NordWireConnectDebugInstaller.exe' if config_data.get('debug_updates', False) else 'NordWireConnectInstaller.exe'}")
            if not test_available.ok: 
                errorMessage("I'm sorry! The latest version is currently unavailable to be downloaded. Please check back later!")
                return
            tk_root.after(100, tkinter_install_updates, latest_ver)
        else: errorMessage("You're already on the latest version of NordWireConnect!")
    except Exception as e: errorMessage(f"Unable to check for new updates due to an exception: {trace(e)}")
def check_for_wireguard_updates():
    try:
        latest_version = get_latest_wireguard_version()
        current_version = get_current_wireguard_version()
        if not latest_version:
            errorMessage("Unable to check for new updates! Check your internet connection..")
            return
        if not current_version:
            tk_root.after(100, tkinter_install_wireguard_updates, latest_version)
            return
        if latest_version["version"] > current_version: 
            tk_root.after(100, tkinter_install_wireguard_updates, latest_version)
        else: errorMessage("You're already on the latest version of Wireguard!")
    except Exception as e: errorMessage(f"Unable to check for new updates due to an exception: {trace(e)}")
def reset_dns_cache():
    try:
        mainMessage("Resetting DNS Cache..")
        system_class.reset_dns_cache()
        successMessageBox("Successfully reset DNS Cache!", "NordWireConnect")
    except Exception as e: errorMessage(f"Unable to reset DNS Cache: {trace(e)}")
def reset_networking():
    try:
        confirmed = messagebox.askyesno("Reset Networking", f"Are you sure you want to reset Networking?\nThis will require administrator permissions and rebooting.", icon="question")
        if confirmed:
            mainMessage("Resetting Networking..")
            system_class.reset_networking()
    except Exception as e: errorMessage(f"Unable to reset Networking: {trace(e)}")
def clear_wireguard_conf_files():
    try:
        confirmed = messagebox.askyesno("Clear Wireguard .conf Files", f"Are you sure you want to clear all Wireguard .conf files?", icon="question")
        if confirmed:
            mainMessage("Clearing Wireguard .conf files..")
            system_class.clear_wireguard_conf_files()
    except Exception as e: errorMessage(f"Unable to clear Wireguard .conf files: {trace(e)}")
def clear_configuration():
    try:
        global config_data
        confirmed = messagebox.askyesno("Clear Configuration", f"Are you sure you want to clear all NordWireConnect configuration data?", icon="question")
        if confirmed:
            disconnect()
            mainMessage("Clearing configuration data..")
            config_data = copy.deepcopy(base_config)
            save_configuration()
            update_tray()
    except Exception as e: errorMessage(f"Unable to clear configuration data: {trace(e)}")
def clear_configuration_except_account():
    try:
        confirmed = messagebox.askyesno("Clear Configuration (except Account)", f"Are you sure you want to clear all NordWireConnect configuration data except your NordVPN Account Authentication Data?", icon="question")
        if confirmed:
            disconnect()
            mainMessage("Clearing configuration data (without account)..")
            config_path = os.path.join(app_data_path, "ConnectConfig.json")
            for k in config_data.keys():
                if k not in ["access_token","username","openvpn_username","openvpn_password","nordvpn_email"]:
                    config_data[k] = base_config.get(k)
            with open(config_path, "w") as f: json.dump(config_data, f)
            load_configuration()
            update_tray()
    except Exception as e: errorMessage(f"Unable to clear configuration data: {trace(e)}")
def get_current_time() -> datetime.datetime: return datetime.datetime.now(tz=datetime.timezone.utc)
def convert_to_local(datetime: datetime.datetime) -> datetime.datetime: return datetime.astimezone()
def get_asn_name_from_ip(ip: str) -> typing.Union[str, None]:
    global cached_asn_data
    try:
        if not cached_asn_data["loaded"]: return None
        ip_int = int(ipaddress.ip_address(ip))
        i = bisect.bisect_right(cached_asn_data["starts"], ip_int) - 1
        if i >= 0 and ip_int <= cached_asn_data["ends"][i]: return cached_asn_data["names"][i]
    except Exception as e: errorMessage(f"Unable to get ASN name from IP: {trace(e)}")
    return None
def get_vpn_renewal_data() -> dict:
    try:
        if tunnel_check(timeout=1):
            if not config_data.get("access_token", ""): return {"active": False, "expires_at": None, "raw_expiry": "0000-00-00 00:00:00"}
            decrypted_token = decode_data("access_token", config_data.get("access_token", ""))
            service_data = requests.get(f"https://api.nordvpn.com/v1/users/services", auth=["token", decrypted_token])
            if service_data.ok:
                for s in service_data.json:
                    if s.get("service", {}).get("identifier") == "vpn":
                        expiry = convert_nordtime_to_datetime(s.get("expires_at", "0000-00-00 00:00:00"))
                        if expiry:
                            return {
                                "active": True,
                                "expires_at": expiry,
                                "raw_expiry": s.get("expires_at", "0000-00-00 00:00:00")
                            }
    except Exception as e: errorMessage(f"Unable to get VPN plan renewal data: {trace(e)}")
    return {"active": False, "expires_at": None, "raw_expiry": "0000-00-00 00:00:00"}
def get_if_within_reminder_period(expiry) -> typing.Union[int, None]:
    cur_time = get_current_time()
    diff = convert_nordtime_to_datetime(expiry) - cur_time
    if diff.days <= 1 and config_data.get("plan_reminder_level", 5) == 6: return 6
    elif diff.days <= 3 and config_data.get("plan_reminder_level", 5) == 5: return 5
    elif diff.days <= 7 and config_data.get("plan_reminder_level", 5) == 4: return 4
    elif diff.days <= 14 and config_data.get("plan_reminder_level", 5) == 3: return 3
    elif diff.days <= 30 and config_data.get("plan_reminder_level", 5) == 2: return 2
    elif diff.days <= 60 and config_data.get("plan_reminder_level", 5) == 1: return 1
    elif diff.days <= 120 and config_data.get("plan_reminder_level", 5) == 0: return 0
    else: return -1
def get_if_within_data_usage() -> typing.Union[int, None]:
    if config_data.get("monthly_data_usage"):
        cur_date = get_current_time()
        usage = config_data["monthly_data_usage"]
        total = usage['download'] + usage['upload']
        if usage["month"] != cur_date.month or usage["year"] != cur_date.year: total = 0
        if total >= 1_000_000_000 and config_data.get("usage_reminder_level", 3) == 8: return 8
        elif total >= 5_000_000_000 and config_data.get("usage_reminder_level", 3) == 7: return 7
        elif total >= 10_000_000_000 and config_data.get("usage_reminder_level", 3) == 6: return 6
        elif total >= 50_000_000_000 and config_data.get("usage_reminder_level", 3) == 5: return 5
        elif total >= 100_000_000_000 and config_data.get("usage_reminder_level", 3) == 4: return 4
        elif total >= 250_000_000_000 and config_data.get("usage_reminder_level", 3) == 3: return 3
        elif total >= 500_000_000_000 and config_data.get("usage_reminder_level", 3) == 2: return 2
        elif total >= 1_024_000_000_000 and config_data.get("usage_reminder_level", 3) == 1: return 1
        elif total >= 2_048_000_000_000 and config_data.get("usage_reminder_level", 3) == 0: return 0
        else: return -1
    else: return -1 
def refresh_data_usage():
    global data_reminded
    try:
        data_reminded = False
        if config_data.get("monthly_data_usage", {}) and config_data.get("data_usage_reminder") == True:
            res = get_if_within_data_usage()
            selector = {
                8: "1 GB",
                7: "5 GB",
                6: "10 GB",
                5: "50 GB",
                4: "100 GB", 
                3: "250 GB",
                2: "500 GB",
                1: "1 TB",
                0: "2 TB"
            }
            if res != -1: 
                successMessageBox(f"Your data usage has exceeded the limit of {selector[res]}!", "NordVPN Data Usage Reminder")
                data_reminded = True
            else: data_reminded = False
    except Exception as e: errorMessage(f"Unable to load data usage information. Exception: {trace(e)}")
def refresh_plan_info():
    global plan_reminded
    try:
        if config_data.get("access_token"):
            mainMessage("Fetching Plan Renewal Information..")
            plan_renewal_info = get_vpn_renewal_data()
            config_data["plan_renewal_data"] = {"expiry": plan_renewal_info.get("raw_expiry"), "active": plan_renewal_info.get("active")}
        else: config_data["plan_renewal_data"] = {"expiry": "0000-00-00 00:00:00", "active": False}
        plan_reminded = False
        if config_data.get("plan_renewal_data", {}).get("active") and config_data.get("plan_renewal_reminder") == True:
            res = get_if_within_reminder_period(config_data["plan_renewal_data"]["expiry"])
            selector = {
                6: "a day",
                5: "3 days",
                4: "a week", 
                3: "2 weeks",
                2: "a month",
                1: "2 months",
                0: "4 months"
            }
            if res != -1: 
                successMessageBox(f"Your NordVPN plan is expiring in less than {selector[res]}!", "NordVPN Plan Expiry Reminder")
                plan_reminded = True
            else: plan_reminded = False
        save_configuration()
    except Exception as e: errorMessage(f"Unable to load plan information. Exception: {trace(e)}")

# Data Handling
def session_data_modified():
    with open(os.path.join(app_data_path, "NordSessionData.json"), "w") as f: json.dump(session_data, f, indent=4)
@functools.lru_cache(maxsize=3)
def convert_nordtime_to_datetime(time_str: str) -> datetime.datetime: return datetime.datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S").replace(tzinfo=datetime.timezone.utc)
def run_cache():
    mainMessage("Updating server cache..")
    while not tunnel_check(timeout=1): time.sleep(0.1)
    all_servers = requests.get(f"https://api.nordvpn.com/v1/servers/recommendations?&filters\\[servers_technologies\\]\\[identifier\\]=wireguard_udp&limit=10000")
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
    all_asns = {}
    while not tunnel_check(timeout=1): time.sleep(0.1)
    all_asns_req = requests.download("https://iptoasn.com/data/ip2asn-v4.tsv.gz", os.path.join(app_data_path, "ip2asn-v4.tsv.gz"))
    all_asns = system_class.unarchive_asn_database(os.path.join(app_data_path, "ip2asn-v4.tsv.gz")) if all_asns_req.ok else []
    parsed_data = []
    for d in all_asns:
        s_ip = int(ipaddress.ip_address(d.get("s_ip")))
        e_ip = int(ipaddress.ip_address(d.get("e_ip")))
        parsed_data.append((s_ip, e_ip, d.get("asn_name")))
    parsed_data.sort(key=lambda x: x[0])
    cached_asn_data["starts"] = [x[0] for x in parsed_data]
    cached_asn_data["ends"] = [x[1] for x in parsed_data]
    cached_asn_data["names"] = [x[2] for x in parsed_data]
    cached_asn_data["loaded"] = True
    with open(os.path.join(app_data_path, "NordServerCache.json"), "w") as f: json.dump(raw_servers, f)
    with open(os.path.join(app_data_path, "NordCountryCache.json"), "w") as f: json.dump(all_countries, f)
    with open(os.path.join(app_data_path, "NordCityCache.json"), "w") as f: json.dump(all_cities, f)
    with open(os.path.join(app_data_path, "NordServerIDCache.json"), "w") as f: json.dump(all_servers_by_id, f)
    with open(os.path.join(app_data_path, "NordASNCache.json"), "w") as f: json.dump(all_asns, f)
def send_command(cmd: str, args: dict={}) -> dict: 
    res = nord_socket.request("NordWireConnectCommands", {"cmd": cmd, **args}, timeout=60)
    if res is None: return {}
    return res

@functools.lru_cache(maxsize=5)
def format_seconds(seconds: int) -> str:
    seconds = int(seconds)
    days, remaining_seconds = divmod(seconds, 86400)
    hours, remaining_seconds = divmod(remaining_seconds, 3600)
    minutes, seconds = divmod(remaining_seconds, 60)
    return f"{days:02d}:{hours:02d}:{minutes:02d}:{seconds:02d}"
@functools.lru_cache(maxsize=5)
def format_wireguard_bytes(n: int) -> str:
    for unit in ("B", "KiB", "MiB", "GiB", "TiB", "PiB"):
        if abs(n) < 1024:
            return f"{n:.2f} {unit}"
        n /= 1024
    return f"{n:.2f} PiB"
@functools.lru_cache(maxsize=5)
def parse_wireguard_bytes(s: str) -> int:
    units = {"B": 0, "KiB": 1, "MiB": 2, "GiB": 3, "TiB": 4, "PiB": 5}
    value, unit = s.split()
    return int(float(value) * (1024 ** units[unit]))
def get_if_virtual_location(server: dict) -> bool:
    for sp in server.get("specifications", []):
        if sp["identifier"] == "virtual_location":
            for val in sp["values"]:
                if val["value"] == "true": return True
    return False
def change_dns():
    inputted = simpledialog.askstring(title="DNS Input", prompt="Enter IPv4 DNS servers to use (comma separated):", initialvalue="103.86.96.100, 103.86.99.100")
    ips = inputted.replace(" ", "").split(",")
    added = [ip for ip in ips if requests.get_if_ip(ip)]
    if added:
        config_data["dns"] = ", ".join(added)
        save_configuration()
        successMessageBox(f"Successfully changed DNS to: {', '.join(added)}", "NordWireConnect")
        update_tray()
    else: errorMessage("No valid IPv4 DNS servers were provided.")
def change_access_token():
    global private_key
    inputted = simpledialog.askstring(title="Access Token Input", prompt="Enter your generated NordVPN Access Token from the NordVPN website (https://my.nordaccount.com/dashboard/nordvpn/):", show="*")
    if len(inputted) > 50:
        test_request = requests.get(f"https://api.nordvpn.com/v1/users/services/credentials", auth=["token", inputted])
        if test_request.ok:
            test_request2 = requests.get(f"https://api.nordvpn.com/v1/users/current", auth=["token", inputted])
            if test_request2.ok:
                test_request = test_request.json
                test_request2 = test_request2.json
                config_data["openvpn_username"] = encrypt_data("openvpn_username", test_request.get("username", ""))
                config_data["openvpn_password"] = encrypt_data("openvpn_password", test_request.get("password", ""))
                config_data["nordvpn_email"] = test_request2.get("email", "")
                config_data["username"] = test_request2.get("username", "")
                config_data["access_token"] = encrypt_data("access_token", inputted)
                private_key = test_request.get("nordlynx_private_key", "")
                successMessageBox(f"Successfully logged into {test_request2.get('email')}!", "NordWireConnect")
                save_configuration()
                update_tray()
            else: errorMessage("The provided NordVPN Access Token is invalid.")
        else: errorMessage("The provided NordVPN Access Token is invalid.")
    else: errorMessage("The provided NordVPN Access Token is invalid.")
def change_connection_channel():
    inputted = simpledialog.askstring(title="Connection Channel Input", prompt="Enter the Connection Channel you would like to assign this computer (1-10) (enter \"none\" for no channels):", initialvalue="none")
    try:
        if inputted.lower() == "none":
            config_data["connection_channel"] = None
            successMessageBox(f"Successfully cleared Connection Channel!", "NordWireConnect")
            save_configuration()
            update_tray()
            return
        inputted = int(inputted)
        if inputted >= 1 and inputted <= 10:
            config_data["connection_channel"] = inputted
            successMessageBox(f"Successfully changed Connection Channel to {inputted}!", "NordWireConnect")
            save_configuration()
            update_tray()
        else: errorMessage("The given channel number is not in range from 1 to 10.")
    except Exception as e: errorMessage("The given channel number is invalid.")
def change_keepalive():
    inputted = simpledialog.askinteger(title="Persistent Keepalive Input", prompt="Enter the amount of keepalive (seconds) you would like to set (1-999):", initialvalue=25, minvalue=1, maxvalue=999)
    try:
        if inputted >= 1 and inputted <= 999:
            config_data["persistent_keepalive"] = inputted
            successMessageBox(f"Successfully changed Persistent Keepalive to {inputted}!", "NordWireConnect")
            save_configuration()
            update_tray()
        else: errorMessage("The given channel number is not in range from 1 to 999.")
    except Exception as e: errorMessage("The given channel number is invalid.")
def change_load_swap_threshold():
    inputted = simpledialog.askinteger(title="Load Swapping Threshold Input", prompt="Enter the load percentage threshold for swapping servers (0-100):", initialvalue=50, minvalue=0, maxvalue=100)
    try:
        if inputted >= 0 and inputted <= 100:
            config_data["load_swap_threshold"] = inputted
            successMessageBox(f"Successfully changed Load Swapping Threshold to {inputted}%!", "NordWireConnect")
            save_configuration()
            update_tray()
        else: errorMessage("The given load percentage is not in range from 0 to 100.")
    except Exception as e: errorMessage("The given load percentage is invalid.")
def open_nord_dashboard():
    try: webbrowser.open("https://my.nordaccount.com/dashboard/nordvpn/")
    except Exception as e: errorMessage(f"Unable to open NordVPN Dashboard: {trace(e)}")
def add_server_to_history(info: dict):
    if not config_data.get("server_history"): config_data["server_history"] = []
    if not check_track_servers(): return
    history = config_data["server_history"]
    history.insert(0, info)
    if len(history) > 25: history = history[:25]
    config_data["server_history"] = history
    save_configuration()
def add_data_usage_to_history(download: int, upload: int):
    global add_data_count
    if download < 0 or upload < 0: return
    cur_date = get_current_time()
    if not config_data.get("monthly_data_usage"): config_data["monthly_data_usage"] = {"month": cur_date.month, "year": cur_date.year, "upload": 0, "download": 0}
    if not config_data.get("total_data_usage"): config_data["total_data_usage"] = {"upload": 0, "download": 0}
    usage = config_data["monthly_data_usage"]
    total_usage = config_data["total_data_usage"]
    if usage["month"] != cur_date.month or usage["year"] != cur_date.year: config_data["monthly_data_usage"] = {"month": cur_date.month, "year": cur_date.year, "upload": 0, "download": 0}
    usage["upload"] += upload
    usage["download"] += download
    config_data["monthly_data_usage"] = usage
    total_usage["upload"] += upload
    total_usage["download"] += download
    config_data["total_data_usage"] = total_usage
    if data_reminded == False: refresh_data_usage()
    add_data_count += 1
    if add_data_count % 10 == 0: save_configuration()
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
def load_configuration():
    global config_data
    if os.path.exists(os.path.join(app_data_path, "ConnectConfig.json")):
        with open(os.path.join(app_data_path, "ConnectConfig.json"), "r") as f: temp_conf = json.load(f)
        for k, v in temp_conf.items():
            if isinstance(v, config_data_type_allowed.get(k)): config_data[k] = v
        if config_data.get("data_encrypted") == False:
            for k in ["access_token", "openvpn_username", "openvpn_password"]:
                if type(config_data.get(k)) is str: config_data[k] = encrypt_data(k, config_data[k])
            config_data["data_encrypted"] = True
            save_configuration()
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
    config_data["notifications"] = not config_data.get("notifications", True)
    save_configuration()
    update_tray()
def check_notifications(): return config_data.get("notifications", True) == True
def change_server_list():
    config_data["optimize_server_list"] = not config_data.get("optimize_server_list", False)
    addToThread(run_cache, ())
    save_configuration()
    update_tray()
def check_server_list(): return config_data.get("optimize_server_list", False) == True
def change_asn_server_list():
    config_data["asn_server_list"] = not config_data.get("asn_server_list", False)
    addToThread(run_cache, ())
    save_configuration()
    update_tray()
def check_asn_server_list(): return config_data.get("asn_server_list", False) == True
def change_autoupdates():
    config_data["check_for_updates"] = not config_data.get("check_for_updates", True)
    save_configuration()
    update_tray()
def check_autoupdates(): return config_data.get("check_for_updates", True) == True
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
def change_plan_renewal_reminder():
    config_data["plan_renewal_reminder"] = not config_data.get("plan_renewal_reminder", True)
    save_configuration()
    update_tray()
def check_plan_renewal_reminder(): return config_data.get("plan_renewal_reminder", True) == True
def change_data_usage_reminder():
    config_data["data_usage_reminder"] = not config_data.get("data_usage_reminder", False)
    save_configuration()
    update_tray()
def check_data_usage_reminder(): return config_data.get("data_usage_reminder", False) == True
def change_beta_updates():
    config_data["beta_updates"] = not config_data.get("beta_updates", False)
    save_configuration()
    update_tray()
def check_beta_updates(): return config_data.get("beta_updates", False) == True
def change_debug_updates():
    config_data["debug_updates"] = not config_data.get("debug_updates", False)
    save_configuration()
    update_tray()
def check_debug_updates(): return config_data.get("debug_updates", False) == True
def change_track_servers():
    config_data["track_servers"] = not config_data.get("track_servers", False)
    save_configuration()
    update_tray()
def check_track_servers(): return config_data.get("track_servers", False) == True
def change_reminder_level(level_index):
    config_data["plan_reminder_level"] = level_index
    save_configuration()
    update_tray()
def check_reminder_level(level_index): return config_data.get("plan_reminder_level", 5) == level_index
def change_usage_level(level_index):
    config_data["usage_reminder_level"] = level_index
    save_configuration()
    update_tray()
def check_usage_level(level_index): return config_data.get("usage_reminder_level", 3) == level_index
def check_threat_protection_lite(): return config_data.get("threat_protection_lite", False) == True
def change_threat_protection_lite():
    config_data["threat_protection_lite"] = not config_data.get("threat_protection_lite", False)
    save_configuration()
    if session_data["connected"] == True: differ_from_status_action2()
    update_tray()

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
        if s.get("suc") == True: successMessageBox("Successfully disconnected from NordVPN server!", "NordWireConnect")
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
def differ_to_city() -> str:
    if session_data["connected"] == True: return f"City: {session_data['server']['locations'][0]['country']['city']['name']}"
    else: return "City: N/A"
def differ_to_ip_add() -> str: 
    if session_data["connected"] == True: return f"IP Address: {session_data['server']['station']}"
    else: return "IP Address: N/A"
def differ_to_hostname() -> str: 
    if session_data["connected"] == True: return f"Hostname: {session_data['server']['hostname']}"
    else: return "Hostname: N/A"
def differ_to_load() -> str: 
    if session_data["connected"] == True: return f"Load: {session_data['current_load']}%"
    else: return "Load: N/A"
def differ_to_session_time() -> str: 
    if session_data["connected"] == True: return f"Session Time: {format_seconds(session_data['session_time'])}"
    else: return "Session Time: N/A"
def differ_to_data_usage() -> str: 
    if session_data["connected"] == True: return f"Data Usage: {session_data.get('download_data_usage', 'N/A')} ⬇️, {session_data.get('upload_data_usage', 'N/A')} ⬆️"
    else: return "Data Usage: N/A"
def differ_to_monthly_data_usage() -> str:
    if config_data.get("monthly_data_usage"):
        cur_date = get_current_time()
        usage = config_data["monthly_data_usage"]
        if usage["month"] != cur_date.month or usage["year"] != cur_date.year: return "Monthly Data Usage: N/A"
        return f"Monthly Data Usage: {format_wireguard_bytes(usage['download'])} ⬇️, {format_wireguard_bytes(usage['upload'])} ⬆️"
    else: return "Monthly Data Usage: N/A"
def differ_to_total_data_usage() -> str:
    if config_data.get("total_data_usage"):
        usage = config_data["total_data_usage"]
        return f"Total Data Usage: {format_wireguard_bytes(usage['download'])} ⬇️, {format_wireguard_bytes(usage['upload'])} ⬆️"
    else: return "Total Data Usage: N/A"
def differ_to_dns() -> str: 
    if check_threat_protection_lite(): return "DNS Servers: Threat Protection Lite"
    return f"DNS Servers: {config_data['dns']}"
def differ_to_nord_acc() -> str: return f"Account: {config_data.get('username', 'N/A')} (Token {len(config_data.get('access_token', '')) > 60 and 'Given' or 'Ungiven'})"
def differ_to_channel() -> str: return f"Connection Channel: {config_data.get('connection_channel', 'N/A')}"
def differ_to_keepalive() -> str: return f"Persistent Keepalive: {config_data.get('persistent_keepalive', 10)}"
def differ_to_def_loc() -> str: 
    if config_data.get("default_location") and config_data["default_location"].lower() != "auto":
        _, server_loc = config_data["default_location"].split("_")
        return f"Default Location: {server_loc}"
    else: return "Default Location: Auto"
def differ_to_load_swap_usage() -> str: return f"Load Swapping Threshold: {config_data.get('load_swapping_threshold', 50)}%"
def differ_to_renewal_active() -> str:
    if config_data.get("plan_renewal_data", {}).get("active"): return "Plan Active: True"
    else: return "Plan Active: False"
def differ_to_renewal_date() -> str:
    if config_data.get("plan_renewal_data", {}).get("active") and config_data["plan_renewal_data"].get("expiry"):
        expiry_date = convert_nordtime_to_datetime(config_data["plan_renewal_data"]["expiry"])
        return f"Plan Renewal Date: {expiry_date.astimezone().strftime('%a %B %d, %Y at %H:%M:%S')}"
    else: return "Plan Renewal Date: N/A"
def differ_to_renewal_remaining() -> str:
    if config_data.get("plan_renewal_data", {}).get("active") and config_data["plan_renewal_data"].get("expiry"):
        current_date = get_current_time()
        expiry_date = convert_nordtime_to_datetime(config_data["plan_renewal_data"]["expiry"])
        time_diff = expiry_date - current_date
        str_list = []
        if time_diff.days > 0: str_list.append(f"{time_diff.days} days")
        hours = time_diff.seconds // 3600
        if hours > 0: str_list.append(f"{hours} hours")
        minutes = (time_diff.seconds % 3600) // 60
        if minutes > 0: str_list.append(f"{minutes} minutes")
        seconds = time_diff.seconds % 60
        if seconds > 0: str_list.append(f"{seconds} seconds")
        return f"Plan Renewal Remaining: {', '.join(str_list)}"
    else: return "Plan Renewal Remaining: N/A"

# Connections
def calculate_allowed_ips(include_ranges: list[str], exclude_ranges: list[str]) -> list[str]:
    inc = [ipaddress.ip_network(n) for n in include_ranges]
    ex = [ipaddress.ip_network(n) for n in exclude_ranges]
    res = []
    for ver in [4, 6]:
        v_included = [i for i in inc if i.version == ver]
        v_excluded = [e for e in ex if e.version == ver]
        ver_res = []
        for i in v_included:
            cur_net = [i]
            for e in v_excluded:
                new_net = []
                for net in cur_net:
                    if net.overlaps(e):
                        intersection = max(net.network_address, e.network_address)
                        mask = max(net.prefixlen, e.prefixlen)
                        actual_exclusion = ipaddress.ip_network(f"{intersection}/{mask}")
                        new_net.extend(list(net.address_exclude(actual_exclusion)))
                    else: new_net.append(net)
                cur_net = new_net
            ver_res.extend(cur_net)
        res.extend(ipaddress.collapse_addresses(ver_res))
    return [str(net) for net in res]
def get_allowed_ips() -> str: 
    if config_data.get("split_lan_routing"): allowed = calculate_allowed_ips(["0.0.0.0/0", "::/0"], ["10.0.0.0/14", "10.4.0.0/16", "10.6.0.0/15", "10.8.0.0/13", "10.16.0.0/12", "10.32.0.0/11", "10.64.0.0/10", "10.128.0.0/9", "172.16.0.0/12", "192.168.0.0/16", "169.254.0.0/16", "fc00::/7", "fe80::/10", "::1/128"])
    else: allowed = calculate_allowed_ips(["0.0.0.0/0", "::/0"] , [])
    return ", ".join(allowed)
@functools.lru_cache(maxsize=100)
def convert_to_channel(hostname: str) -> int: return ((zlib.crc32(hostname.encode("utf-8")) & 0xffffffff) % 10) + 1
def connect_server(server_name: str, exact_mode: str=None):
    config_data["server"] = server_name
    try: connect(exact_mode=exact_mode)
    except Exception: pass
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
    elif send_command("connection-status").get("active") == True: disconnect()
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
        if not system_class.get_current_wireguard_version():
            successMessageBox("Wireguard was not found in your computer. Installing now.", "NordWireConnect")
            res = system_class.install_wireguard(send_command)
            if not res:
                errorMessage("Failed to install WireGuard.")
                mark_not_connected()
                set_icon("error.ico")
                return

        # Load NordVPN Server Recommendations
        if not tunnel_check(): disconnect(only_disconnect=True)
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
        should_filter = False
        asn_city_name = None
        if config_data["server"] == "auto" or config_data["server"] == None or exact_mode == None:
            all_server_recommendations = requests.get(f"https://api.nordvpn.com/v1/servers/recommendations?{server_filter}&filters\\[servers_technologies\\]\\[identifier\\]=wireguard_udp&limit=100")
            if not all_server_recommendations.ok:
                errorMessage("Unable to get NordVPN server recommendations.")
                mark_not_connected()
                set_icon("error.ico")
                return
            all_server_recommendations = all_server_recommendations.json
            should_filter = True
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
            all_server_recommendations = requests.get(f"https://api.nordvpn.com/v1/servers/recommendations?{server_filter}&filters[country_id]={country_id}&filters\\[servers_technologies\\]\\[identifier\\]=wireguard_udp&limit=100")
            if not all_server_recommendations.ok:
                errorMessage("Unable to get NordVPN server recommendations.")
                mark_not_connected()
                set_icon("error.ico")
                return
            all_server_recommendations = all_server_recommendations.json
            should_filter = True
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
            all_server_recommendations = requests.get(f"https://api.nordvpn.com/v1/servers/recommendations?{server_filter}&filters[country_city_id]={city_id}&filters\\[servers_technologies\\]\\[identifier\\]=wireguard_udp&limit=100")
            if not all_server_recommendations.ok:
                errorMessage("Unable to get NordVPN server recommendations.")
                mark_not_connected()
                set_icon("error.ico")
                return
            all_server_recommendations = all_server_recommendations.json
            should_filter = True
        elif exact_mode == "asn":
            asn_city_name = None
            try: city_name, target_asn = config_data["server"].rsplit("|||", 1)
            except Exception:
                errorMessage("Invalid ASN format given.")
                mark_not_connected()
                set_icon("error.ico")
                return
            asn_city_name = city_name
            city_list_path = os.path.join(app_data_path, "NordCityCache.json")
            if not os.path.exists(city_list_path):
                errorMessage("Unable to find NordVPN city list.")
                mark_not_connected()
                set_icon("error.ico")
                return
            with open(city_list_path, "r") as f:  city_list = json.load(f)
            city_id = None
            for c in city_list:
                if c.get("name") == city_name:
                    city_id = c.get("id")
                    break
            if city_id is None:
                errorMessage("Unable to get NordVPN city ID.")
                mark_not_connected()
                set_icon("error.ico")
                return
            req = requests.get(f"https://api.nordvpn.com/v1/servers/recommendations?{server_filter}&filters[country_city_id]={city_id}&filters\\[servers_technologies\\]\\[identifier\\]=wireguard_udp&limit=20000")
            if not req.ok:
                errorMessage("Unable to get NordVPN server recommendations.")
                mark_not_connected()
                set_icon("error.ico")
                return
            city_servers = req.json
            all_server_recommendations = []
            for server in city_servers:
                if len(all_server_recommendations) >= 100: break
                server_ip = server.get("station") 
                if server_ip:
                    asn_name = get_asn_name_from_ip(server_ip)
                    if asn_name == target_asn: all_server_recommendations.append(server)
            if not all_server_recommendations:
                errorMessage(f"No available servers found for {target_asn} in {city_name}.")
                mark_not_connected()
                set_icon("error.ico")
                return
            should_filter = True
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
                all_server_recommendations = requests.get(f"https://api.nordvpn.com/v1/servers?{server_filter}&filters[servers.id]={server_id}&limit=1")
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
            all_server_recommendations = requests.get(f"https://api.nordvpn.com/v1/servers/recommendations?{server_filter}&filters\\[servers_technologies\\]\\[identifier\\]=wireguard_udp&limit=100")
            if not all_server_recommendations.ok:
                errorMessage("Unable to get NordVPN server recommendations.")
                mark_not_connected()
                set_icon("error.ico")
                return
            all_server_recommendations = all_server_recommendations.json
            should_filter = True
        if should_filter and config_data.get("connection_channel"):
            ind = 0
            channeled_list = []
            for s in all_server_recommendations:
                if convert_to_channel(s["hostname"]) == config_data.get("connection_channel"): channeled_list.append(s)
                ind += 1
            all_server_recommendations = channeled_list
        mainMessage(f"Loaded {len(all_server_recommendations)} NordVPN server recommendation(s).")

        # Start Finding Servers
        mainMessage(f"Connecting to NordVPN Servers in {config_data['server']}..")
        connected_server = None
        attempts = 0
        for s in all_server_recommendations:
            shortened_name = s["name"].replace(" ", "").replace("#", "")
            country_name = s["locations"][0]["country"]["name"].replace(" ", "")
            city_name = s["locations"][0]["country"]["city"]["name"].replace(" ", "")
            if not (config_data["server"] == "auto" or config_data["server"] == None) and not (shortened_name == config_data["server"].replace(" ", "").replace("#", "") or city_name == config_data["server"].replace(" ", "") or country_name == config_data["server"].replace(" ", "") or city_name == asn_city_name.replace(" ", "")):
                continue
            if session_data["connection_text"] == "Not Connected" and session_data["connected"] == False: break
            if s["status"] != "online" or s["load"] == 0:
                mainMessage(f"Server offline: {s['name']}")
                continue

            mainMessage(f"Attempting connection for: {s['name']}")
            wireguard_metadata = None
            for m in s.get("technologies", []):
                if m.get("id") == 35:
                    wireguard_metadata = m.get("metadata", {})
                    break
            if not wireguard_metadata: continue
            wait_until_connected()
            dns = config_data.get("dns", "103.86.96.100, 103.86.99.100")
            if check_threat_protection_lite(): dns = "103.86.99.99, 103.86.96.96"
            configuration = f"""# {shortened_name}.{city_name}.conf
[Interface]
PrivateKey = {private_key}
ListenPort = 51820
DNS = {dns}
Address = 10.5.0.2/16
PostUp = echo Connected > "{os.path.join(app_data_path, 'ConnectionStatus')}"
PostDown = echo Disconnected > "{os.path.join(app_data_path, 'ConnectionStatus')}"

[Peer]
PublicKey = {wireguard_metadata[0].get("value")}
Endpoint = {s.get('station')}:51820
AllowedIPs = {get_allowed_ips()}
PersistentKeepalive = {config_data.get('persistent_keepalive', 10)}"""
            if not server_check(server=s.get('station')): continue
            config_path = os.path.join(app_data_path, f"NordWireConnect.conf")
            with open(config_path, "w") as f: f.write(configuration)
            if session_data["connection_text"] == "Not Connected" and session_data["connected"] == False: break
            if send_command("wireguard-check").get("active") == True: send_command("uninstall-wire-tunnel", {"server": f"{shortened_name}.{city_name}"})
            vpn_res = send_command("install-wire-tunnel", {"config_path": config_path})
            attempts += 1
            #if os.path.exists(config_path): os.remove(config_path)
            if vpn_res and vpn_res.get("suc") == True:
                mainMessage(f"Added new WireGuard configuration for server: {s['name']}")
                tunnel_check(timeout=3)
                time.sleep(1)
                if tunnel_check(timeout=3):
                    connected_server = s
                    break
                else: 
                    send_command("uninstall-wire-tunnel", {"server": f"{shortened_name}.{city_name}"})
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
            add_server_to_history({
                "city": connected_server['locations'][0]['country']['city']['name'],
                "country": connected_server['locations'][0]['country']['name'],
                "time": get_current_time().timestamp(),
                "name": connected_server['name'],
            })
            successMessageBox(f"Successfully connected to server: {connected_server['name']}!", "Yay!")
            set_icon("connected.ico")
            needing_state()
        else:
            errorMessage("Unable to connect to a NordVPN server.")
            mark_not_connected()
            set_icon("error.ico")
            return
    except Exception as e:
        errorMessage(f"Unable to connect to a NordVPN server due to an error: {trace(e)}.")
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
    if awared: return send_command("combined-disconnect", {"known_server": awared})
    else: return send_command("combined-disconnect")
def brute_end_wireguard():
    disconnect()
    mark_not_connected()
def handle_stat_thread():
    global stat_count
    global since_connected
    try:
        if session_data["connected"] == True:
            stat_count += 1
            since_connected += 1
            if config_data.get("loss_connect_protection", True) == True:
                if stat_count % 10 == 0 and not tunnel_check() and not tunnel_check(timeout=2) and not tunnel_check(timeout=2) and not session_data["connection_text"].startswith("Connecting"):
                    mainMessage("Reconnecting due to loss of connection.")
                    differ_from_status_action2()
                    return tk_root.after(1000, handle_stat_thread)
                if send_command("wireguard-check").get("active") == False and not session_data["connection_text"].startswith("Connecting"):
                    mainMessage("Reconnecting due to loss of VPN.")
                    differ_from_status_action2()
                    return tk_root.after(1000, handle_stat_thread)
                if os.path.exists(os.path.join(app_data_path, "ConnectionStatus")):
                    with open(os.path.join(app_data_path, "ConnectionStatus"), "r") as f: status = f.read().strip()
                    if status == "Disconnected" and not session_data["connection_text"].startswith("Connecting"):
                        mainMessage("Reconnecting due to loss of VPN.")
                        differ_from_status_action2()
                        return tk_root.after(1000, handle_stat_thread)

            if stat_count % 20 == 0:
                server_id = session_data["server"]["id"]
                servers = requests.get(f"https://api.nordvpn.com/v1/servers?&filters[servers.id]={server_id}&limit=1")
                servers = servers.json if servers.ok else []
                found_server = None
                for s in servers:
                    if s["id"] == server_id:
                        found_server = s
                        break
                session_data["current_load"] = found_server["load"] if found_server else session_data["current_load"]
                if check_load_swapping() and session_data["current_load"] > config_data.get("load_swapping_threshold", 50):
                    mainMessage("Reconnecting due to load > 50%..")
                    differ_from_status_action2()
                data_usage = send_command("data-usage")
                if data_usage.get("suc") == True and len(data_usage) == 2:
                    add_data_usage_to_history(parse_wireguard_bytes(data_usage["download"])-parse_wireguard_bytes(session_data["download_data_usage"]), parse_wireguard_bytes(data_usage["upload"])-parse_wireguard_bytes(session_data["upload_data_usage"]))
                    session_data["download_data_usage"] = data_usage["download"]
                    session_data["upload_data_usage"] = data_usage["upload"]
            session_data["session_time"] = since_connected
            pystray_icon.title = f"NordWireConnect v{version}\nConnected to {session_data['server']['name']}!\nCity: {session_data['server']['locations'][0]['country']['city']['name']}\nSession Time: {format_seconds(session_data['session_time'])}"
            update_tray()
        else:
            stat_count = 0
            since_connected = 0
            pystray_icon.title = f"NordWireConnect v{version}"
    except Exception: pass
    tk_root.after(1000, handle_stat_thread)
def poll_ui_queue():
    while not ui_queue.empty():
        try:
            func = ui_queue.get_nowait()
            if func: func()
        except Exception as e: errorMessage(f"Error handling UI event: {trace(e)}")
    tk_root.after(50, poll_ui_queue)
def handle_plan_renewal_thread():
    global plan_reminded
    global plan_timer
    plan_timer += 1
    if plan_timer % 300 == 0:
        try: 
            if config_data.get("plan_renewal_reminder", True) == True and config_data.get("plan_renewal_data", {}).get("active") == True and plan_reminded == False: refresh_plan_info()
        except Exception: pass
    if config_data.get("plan_renewal_reminder", True) == True and config_data.get("plan_renewal_data", {}).get("active") == True and plan_reminded == False and config_data.get("plan_renewal_data", {}).get("expiry"):
        try:
            res = get_if_within_reminder_period(config_data["plan_renewal_data"]["expiry"])
            if res != -1: refresh_plan_info()
        except Exception: pass
    tk_root.after(1000, handle_plan_renewal_thread)
def tunnel_check(timeout: int=5): return requests.get_if_connected(server="1.1.1.1", timeout=timeout)
def server_check(server: str="1.1.1.1", timeout: int=5): 
    c = 0
    t = time.time()
    while c < 3 and time.time() - t < timeout*timeout:
        if requests.get_if_connected(server=server, timeout=timeout): c += 1
        time.sleep(0.1)
    return c == 3
def wait_until_connected(): 
    while not requests.get_if_connected(): time.sleep(0.1)

# Main Runtime
def app():
    global one_instance_lock
    global config_data
    global pystray_icon
    global private_key
    global session_data
    global debug_mode
    global tk_root
    if main_os == "Windows": colors_class.fix_windows_ansi()
    setup_logging()
    systemMessage(f"{'-'*5:^5} NordWireConnect v{version} {'-'*5:^5}")

    # Setup Tkinter
    mainMessage("Starting Tkinter Library..")
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        os.environ["TCL_LIBRARY"] = os.path.join(sys._MEIPASS, "_tcl_data")
        os.environ["TK_LIBRARY"] = os.path.join(sys._MEIPASS, '_tk_data')
    tk_root = tk.Tk()
    tk_root.withdraw()

    # Ensure App Data Path
    mainMessage("Ensuring App Data Paths..")
    os.makedirs(app_data_path, exist_ok=True)

    # Check Debugging Mode
    try: 
        debug_mode = os.path.exists(os.path.join(os.path.dirname(__file__), "debug_flag"))
        if debug_mode:
            with open(os.path.join(os.path.dirname(__file__), "debug_flag"), "r") as f: debug_mode = f.read().strip() == "true"
    except Exception: debug_mode = False
    if debug_mode: warnMessage("Debug Mode is enabled and console window should be visible if you can see this!")

    # Only for Windows
    if main_os not in System.supported_systems:
        errorMessage(System.not_supported_txt)
        sys.exit(1)
        return

    # Load Uninstaller
    if "-uninstall-nord-wire-connect" in sys.argv: 
        try: uninstall_app()
        except Exception as e: errorMessage(f"There was an error during Installation Handler: {trace(e)}"); sys.exit(1); return

    # NordWireConnect Instance Check
    mainMessage("Checking for Other Existing NordWireConnect Instances..")
    lock_path = os.path.join(app_data_path, "OneInstanceLock")
    one_instance_lock = PyKits.Lock(lock_path)
    if one_instance_lock.exists():
        errorMessage("Another instance of NordWireConnect is already running. Exiting now.")
        sys.exit(1)
        return
    else:
        try: one_instance_lock.acquire()
        except Exception as e: errorMessage(f"There was an error trying to acquire instance lock: {trace(e)}"); sys.exit(1); return
    
    # Load Configuration
    mainMessage("Loading Configuration..")
    try:
        load_configuration()
        if os.path.exists(os.path.join(app_data_path, "NordSessionData.json")): 
            with open(os.path.join(app_data_path, "NordSessionData.json"), "r") as f: session_data = json.load(f)
    except Exception as e: errorMessage(f"There was an error trying to load configuration: {trace(e)}"); sys.exit(1); return

    # Handle Starting with Service
    mainMessage("Verifying Service Status..")
    if not system_class.service_exists():
        mainMessage("NordWireConnectService is not running. Attempting to start it now..")
        system_class.start_service()
    was_old = config_data.get("last_version") != version
    if was_old == True: setup_shortcuts()
    send_command("ui-opening")
    config_data["last_version"] = version

    # Break Hanging Connections
    mainMessage("Breaking Hanging Connections..")
    try:
        if not tunnel_check():
            disconnect(only_disconnect=True)
            while not tunnel_check(timeout=1): time.sleep(0.1)
    except Exception as e: errorMessage(f"There was an error trying to break hanging connections: {trace(e)}"); sys.exit(1); return

    # Fetch NordVPN Credentials
    if config_data.get("access_token"):
        mainMessage("Getting NordVPN credentials..")
        try:
            decrypted_token = decode_data("access_token", config_data["access_token"])
            credentials = requests.get(f"https://api.nordvpn.com/v1/users/services/credentials", auth=["token", decrypted_token])
            if not credentials.ok:
                errorMessage("Unable to get NordVPN credentials. Is the access token valid?")
                sys.exit(1)
                return
            credentials = credentials.json
            private_key = credentials.get("nordlynx_private_key", "")
        except Exception as e:
            errorMessage(f"Unable to get NordVPN credentials. Exception: {trace(e)}")
            sys.exit(1)
            return
    
    # Load Servers
    try:
        mainMessage("Fetching Servers..")
        addToThread(run_cache, ())
    except Exception as e: errorMessage(f"Unable to load server list. Exception: {trace(e)}"); sys.exit(1); return
    
    # Load Plan Information
    try: refresh_plan_info()
    except Exception as e: errorMessage(f"Unable to load plan information. Exception: {trace(e)}")

    # Create Tray App
    try:
        mainMessage("Creating Tray App..")
        def create_mini_func(func: typing.Callable, connected: bool=False):
            def a(icon: Icon, item: MenuItem): addToThread(func, (), connection=False)
            return a
        shared_app = None
        if main_os == "Darwin":
            from Cocoa import NSApplication
            shared_app = NSApplication.sharedApplication()
        image = getImageObj(os.path.join(program_files, "Resources", "app_icon.ico"))
        pystray_icon = pystray.Icon(
            "NordWireConnect",
            image,
            "NordWireConnect",
            menu=pystray.Menu(
                pystray.MenuItem("NordWireConnect", lambda icon, item: None, enabled=False),
                pystray.MenuItem(unicon(differ_to_status), lambda icon, item: None, enabled=False),
                pystray.MenuItem(unicon(differ_to_city), lambda icon, item: None, enabled=False, visible=lambda icon: session_data["connected"] == True),
                pystray.MenuItem(unicon(differ_to_ip_add), lambda icon, item: None, enabled=False, visible=lambda icon: session_data["connected"] == True),
                pystray.MenuItem(unicon(differ_to_hostname), lambda icon, item: None, enabled=False, visible=lambda icon: session_data["connected"] == True),
                pystray.MenuItem(unicon(differ_to_load), lambda icon, item: None, enabled=False, visible=lambda icon: session_data["connected"] == True),
                pystray.MenuItem(unicon(differ_to_session_time), lambda icon, item: None, enabled=False, visible=lambda icon: session_data["connected"] == True),
                pystray.MenuItem(unicon(differ_to_data_usage), lambda icon, item: None, enabled=False, visible=lambda icon: session_data["connected"] == True),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem(unicon(differ_from_status_text), create_mini_func(differ_from_status_action, connected=True)),
                pystray.MenuItem("Reconnect Servers", create_mini_func(differ_from_status_action2, connected=True), visible=lambda icon: session_data["connected"] == True),
                pystray.MenuItem("Servers", lambda icon, item: ui_queue.put(connect_selected_server_button)),
                pystray.MenuItem("History", lambda icon, item: ui_queue.put(connect_history_server_button)),
                pystray.MenuItem("Configuration", pystray.Menu(
                    pystray.MenuItem("Connection", pystray.Menu(
                        pystray.MenuItem(unicon(differ_to_dns), lambda icon, item: None, enabled=False),
                        pystray.MenuItem(unicon(differ_to_nord_acc), lambda icon, item: None, enabled=False),
                        pystray.MenuItem(unicon(differ_to_channel), lambda icon, item: None, enabled=False),
                        pystray.MenuItem(unicon(differ_to_keepalive), lambda icon, item: None, enabled=False),
                        pystray.Menu.SEPARATOR,
                        pystray.MenuItem("Change DNS Servers", lambda icon, item: ui_queue.put(change_dns), visible=lambda icon: check_threat_protection_lite() == False),
                        pystray.MenuItem("Change Access Token", lambda icon, item: ui_queue.put(change_access_token)),
                        pystray.MenuItem("Change Connection Channel", lambda icon, item: ui_queue.put(change_connection_channel)),
                        pystray.MenuItem("Change Persistent Keepalive", lambda icon, item: ui_queue.put(change_keepalive)),
                        pystray.MenuItem(unicon(differ_to_def_loc), lambda icon, item: ui_queue.put(save_auto_selected_server_button)),
                        pystray.Menu.SEPARATOR,
                        pystray.MenuItem("Load Swapping", unicon(change_load_swapping), checked=unicon(check_load_swapping), radio=True),
                        pystray.MenuItem(unicon(differ_to_load_swap_usage), lambda icon, item: ui_queue.put(change_load_swap_threshold), visible=unicon(check_load_swapping)),
                        pystray.MenuItem("Auto Reconnect", unicon(change_auto_connect), checked=unicon(check_auto_connect), radio=True),
                        pystray.MenuItem("Split Private LAN Routing", unicon(change_split_lan_routing), checked=unicon(check_split_lan_routing), radio=True),
                        pystray.MenuItem("Loss Connect Protection", unicon(change_loss_connect_protection), checked=unicon(check_loss_connect_protection), radio=True),
                        pystray.MenuItem("Threat Protection Lite", unicon(change_threat_protection_lite), checked=unicon(check_threat_protection_lite), radio=True)
                    )),
                    pystray.MenuItem("Plan", pystray.Menu(
                        pystray.MenuItem(unicon(differ_to_renewal_active), lambda icon, item: None, enabled=False, radio=True),
                        pystray.MenuItem(unicon(differ_to_renewal_date), lambda icon, item: None, enabled=False, radio=True),
                        pystray.MenuItem(unicon(differ_to_renewal_remaining), lambda icon, item: None, enabled=False, radio=True),
                        pystray.MenuItem("Plan Renewal Reminder", unicon(change_plan_renewal_reminder), checked=unicon(check_plan_renewal_reminder), radio=True),
                        pystray.MenuItem("Reminder Level", pystray.Menu(
                            pystray.MenuItem("Level 6 (1 day)", unicon(change_reminder_level, [6]), checked=unicon(check_reminder_level, [6]), radio=True),
                            pystray.MenuItem("Level 5 (3 days)", unicon(change_reminder_level, [5]), checked=unicon(check_reminder_level, [5]), radio=True),
                            pystray.MenuItem("Level 4 (1 week)", unicon(change_reminder_level, [4]), checked=unicon(check_reminder_level, [4]), radio=True),
                            pystray.MenuItem("Level 3 (2 weeks)", unicon(change_reminder_level, [3]), checked=unicon(check_reminder_level, [3]), radio=True),
                            pystray.MenuItem("Level 2 (1 month)", unicon(change_reminder_level, [2]), checked=unicon(check_reminder_level, [2]), radio=True),
                            pystray.MenuItem("Level 1 (2 months)", unicon(change_reminder_level, [1]), checked=unicon(check_reminder_level, [1]), radio=True),
                            pystray.MenuItem("Level 0 (4 months)", unicon(change_reminder_level, [0]), checked=unicon(check_reminder_level, [0]), radio=True),
                        )),
                        pystray.MenuItem("Refresh Plan Information", create_mini_func(refresh_plan_info)),
                        pystray.MenuItem("Open NordVPN Dashboard", unicon(open_nord_dashboard))
                    )),
                    pystray.MenuItem("Data Usage", pystray.Menu(
                        pystray.MenuItem(unicon(differ_to_monthly_data_usage), lambda icon, item: None, enabled=False, radio=True),
                        pystray.MenuItem(unicon(differ_to_total_data_usage), lambda icon, item: None, enabled=False, radio=True),
                        pystray.MenuItem("Data Usage Reminder (Monthly)", unicon(change_data_usage_reminder), checked=unicon(check_data_usage_reminder), radio=True),
                        pystray.MenuItem("Usage Level", pystray.Menu(
                            pystray.MenuItem("Level 8 (1 GB)", unicon(change_usage_level, [8]), checked=unicon(check_usage_level, [8]), radio=True),
                            pystray.MenuItem("Level 7 (5 GB)", unicon(change_usage_level, [7]), checked=unicon(check_usage_level, [7]), radio=True),
                            pystray.MenuItem("Level 6 (10 GB)", unicon(change_usage_level, [6]), checked=unicon(check_usage_level, [6]), radio=True),
                            pystray.MenuItem("Level 5 (50 GB)", unicon(change_usage_level, [5]), checked=unicon(check_usage_level, [5]), radio=True),
                            pystray.MenuItem("Level 4 (100 GB)", unicon(change_usage_level, [4]), checked=unicon(check_usage_level, [4]), radio=True),
                            pystray.MenuItem("Level 3 (250 GB)", unicon(change_usage_level, [3]), checked=unicon(check_usage_level, [3]), radio=True),
                            pystray.MenuItem("Level 2 (500 GB)", unicon(change_usage_level, [2]), checked=unicon(check_usage_level, [2]), radio=True),
                            pystray.MenuItem("Level 1 (1 TB)", unicon(change_usage_level, [1]), checked=unicon(check_usage_level, [1]), radio=True),
                            pystray.MenuItem("Level 0 (2 TB)", unicon(change_usage_level, [0]), checked=unicon(check_usage_level, [0]), radio=True),
                        ))
                    )),
                    pystray.MenuItem("App", pystray.Menu(
                        pystray.MenuItem("Notifications", unicon(change_notifications), checked=unicon(check_notifications), radio=True),
                        pystray.MenuItem("Optimize Server List", unicon(change_server_list), checked=unicon(check_server_list), radio=True),
                        pystray.MenuItem("Add ASNs to Server List", unicon(change_asn_server_list), checked=unicon(check_asn_server_list), radio=True),
                        pystray.MenuItem("Auto Check for Updates", unicon(change_autoupdates), checked=unicon(check_autoupdates), radio=True),
                        pystray.MenuItem("Beta Updates", unicon(change_beta_updates), checked=unicon(check_beta_updates), radio=True),
                        pystray.MenuItem("Debug Mode", unicon(change_debug_updates), checked=unicon(check_debug_updates), radio=True),
                        pystray.MenuItem("Track Servers Connected", unicon(change_track_servers), checked=unicon(check_track_servers), radio=True),
                    )),
                    pystray.MenuItem("Tools", pystray.Menu(
                        pystray.MenuItem("Set Shortcuts", create_mini_func(unicon(setup_shortcuts, [True]))),
                        pystray.MenuItem("Clear Server Connected History", create_mini_func(unicon(clear_server_connected_history))),
                        pystray.MenuItem("Brute End Wireguard", create_mini_func(brute_end_wireguard)),
                        pystray.MenuItem("Reset DNS Cache", lambda icon, item: unicon(reset_dns_cache)),
                        pystray.MenuItem("Reset Networking", lambda icon, item: ui_queue.put(reset_networking)),
                        pystray.MenuItem("Clear Wireguard .conf Files", lambda icon, item: ui_queue.put(clear_wireguard_conf_files)),
                        pystray.MenuItem("Clear Configuration", lambda icon, item: ui_queue.put(clear_configuration)),
                        pystray.MenuItem("Clear Configuration (except Account)", lambda icon, item: ui_queue.put(clear_configuration_except_account)),
                    )),
                )),
                pystray.MenuItem("About NordWireConnect", unicon(about)),
                pystray.MenuItem("Check for Updates", unicon(check_for_updates)),
                pystray.MenuItem("Check for Wireguard Updates", unicon(check_for_wireguard_updates)),
                pystray.MenuItem("Quit", unicon(quit_app))
            ),
            darwin_nsapplication=shared_app if main_os == "Darwin" else None
        )
    except Exception as e: errorMessage(f"Unable to run app. Exception: {trace(e)}"); sys.exit(1); return

    # Run Main Loop
    try:
        mainMessage("Finishing app load!")
        addToThread(connect_session, (), connection=True)
        threading.Thread(target=worker, daemon=True).start()
        tk_root.after(0, handle_stat_thread)
        tk_root.after(0, handle_plan_renewal_thread)
        if config_data.get("check_for_updates", True) == True: threading.Thread(target=auto_check_for_updates, daemon=True).start()
        tk_root.after(100, check_stop_flag)
        tk_root.iconphoto(True, tk.PhotoImage(file=os.path.join(program_files, "Resources", "app_icon.png"))) 
        tk_root.after(50, poll_ui_queue)
        pystray_icon.run_detached(unicon(setup))
        tk.mainloop()
    except Exception as e: errorMessage(f"Unable to run app. Exception: {trace(e)}"); sys.exit(1); return
if __name__ == "__main__": app()