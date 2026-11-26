"""
NordWireConnect

Made by @EfazDev
https://www.efaz.dev
"""

# Modules
import os
import re
import sys
import copy
import json
import math
import time
import zlib
import array
import socket
import bisect
import random
import struct
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
terminal = PyKits.Terminal()
nord_socket = PyKits.Socket(port=61245)
requests = PyKits.request(throw_exceptions=False)
main_os = platform.system()
work_queue = Queue()
ui_queue = Queue()
asn_lock = threading.Lock()
cache_lock = threading.Lock()
system_class = System.system
plan_timer = 0
stat_count = 0
since_connected = 0
connect_monotonic_time = 0.0
app_logger = None
reconnect_queued = threading.Event()
last_tunnel_rx = 0
last_tunnel_tx = 0
last_swapped_server_id = None
last_tray_signature = None
pause_until_monotonic = 0.0
paused_target = None
active_tunnel_mtu = None
last_physical_ips = frozenset()
excluded_maintenance_hosts = {}
cached_cybersec_ips = {"ips": [], "updated_at": 0.0}
cached_user_insights = {"loaded": False, "latitude": None, "longitude": None, "country_code": "", "updated_at": 0.0}
nord_server_groups = {
    "all": ("All Servers", None),
    "standard": ("Standard VPN", 11),
    "p2p": ("P2P", 15),
    "double_vpn": ("Double VPN", 1),
    "onion_over_vpn": ("Onion Over VPN", 3),
    "dedicated_ip": ("Dedicated IP", 9)
}
pause_intervals = ((300, "5 Minutes"), (900, "15 Minutes"), (1800, "30 Minutes"), (3600, "1 Hour"), (86400, "24 Hours"))
wireguard_tech_filter = "&filters[servers_technologies][identifier]=wireguard_udp"
probe_endpoints = (("1.1.1.1", 443), ("8.8.8.8", 443), ("9.9.9.9", 53))
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
    "encrypted_private_key": "",
    "username": "",
    "openvpn_username": "",
    "openvpn_password": "",
    "nordvpn_email": "",
    "persistent_keepalive": 25,
    "mtu": None,
    "last_version": None,
    "dns": "103.86.96.100, 103.86.99.100", # NordVPN DNS
    "server": "auto",
    "server_group": "all",
    "auto_connect": True,
    "notifications": True,
    "optimize_server_list": False,
    "default_location": "auto",
    "load_swapping": False,
    "load_swapping_threshold": 50,
    "split_lan_routing": False,
    "routing_leak_protection": True,
    "instant_reject_ipv6": True,
    "allowlisted_subnets": [],
    "loss_connect_protection": True,
    "server_health_monitor": True,
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
    "debug_mode": False
}
config_data = copy.deepcopy(base_config)
config_data_type_allowed = {
    "access_token": str,
    "encrypted_private_key": str,
    "username": str,
    "openvpn_username": str,
    "openvpn_password": str,
    "nordvpn_email": str,
    "persistent_keepalive": int,
    "mtu": typing.Union[int, None],
    "dns": str,
    "last_version": typing.Union[str, None],
    "server": str,
    "server_group": str,
    "auto_connect": bool,
    "notifications": bool,
    "optimize_server_list": bool,
    "default_location": str,
    "load_swapping": bool,
    "load_swapping_threshold": int,
    "split_lan_routing": bool,
    "routing_leak_protection": bool,
    "instant_reject_ipv6": bool,
    "allowlisted_subnets": list,
    "asn_server_list": bool,
    "loss_connect_protection": bool,
    "server_health_monitor": bool,
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
    "debug_mode": bool
}
session_data = {
    "connected": False,
    "connection_text": "Not Connected",
    "server": None,
    "current_load": None,
    "server_health": "Good",
    "session_time": 0,
    "exact_connect": None,
    "download_data_usage": "0.00 B",
    "upload_data_usage": "0.00 B"
}
full_files = False
pystray_icon = None
stop_app = False
add_data_count = 0
version = "1.5.0r"
current_icon_name = "app_icon.ico"
tray_ready = threading.Event()
icon_lock = threading.Lock()
cached_asn_data = {"loaded": False, "starts": array.array("I"), "ends": array.array("I"), "name_ids": array.array("I"), "unique_names": []}
server_asn_map = {}
server_asn_map_loaded = False
server_cache_mem = {"mtime": 0.0, "servers": []}
service_pipe = r"\\.\pipe\NordWireConnect"
tk_root = None
Icon = pystray.Icon
MenuItem = pystray.MenuItem
server_filter = "&filters[servers.status]=online&fields[servers.id]&fields[servers.name]&fields[servers.hostname]&fields[servers.station]&fields[servers.status]&fields[servers.load]&fields[servers.created_at]&fields[servers.groups.id]&fields[servers.groups.title]&fields[servers.technologies.id]&fields[servers.technologies.metadata]&fields[servers.technologies.pivot.status]&fields[servers.specifications.identifier]&fields[servers.specifications.values.value]&fields[servers.locations.country.id]&fields[servers.locations.country.name]&fields[servers.locations.country.code]&fields[servers.locations.country.city.id]&fields[servers.locations.country.city.name]&fields[servers.locations.country.city.latitude]&fields[servers.locations.country.city.longitude]&fields[servers.locations.country.city.hub_score]&fields[servers.ips]"

# Logging and Messages
def systemMessage(message: str): terminal.print_ansi(message, PyKits.Terminal.hex_to_ansi("#3E5FFF"))
def mainMessage(message: str): terminal.print_ansi(message, 15)
def errorMessage(message: str, title: str="Uh oh!"): 
    terminal.print_ansi(message, 9)
    message = message[:256]
    if config_data.get("notifications", True) and pystray_icon: notification(title, message)
    elif tk_root:
        if title == "": title = "NordWireConnect"
        elif title != "NordWireConnect": title = f"NordWireConnect: {title}"
        def a(): messagebox.showerror(title, message)
        tk_root.after(0, a)
def warnMessage(message: str): terminal.print_ansi(message, 11)
def successMessage(message: str): terminal.print_ansi(message, 10)
def successMessageBox(message: str, title: str="Success!"):
    message = message[:256]
    if config_data.get("notifications", True) and pystray_icon: notification(title, message)
    elif tk_root:
        if title == "": title = "NordWireConnect"
        elif title != "NordWireConnect": title = f"NordWireConnect: {title}"
        def a(): messagebox.showinfo(title, message)
        tk_root.after(0, a)
def prune_old_logs(log_path: str, prefix: str, keep: int=15):
    try:
        entries = [os.path.join(log_path, f) for f in os.listdir(log_path) if f.startswith(prefix) and f.endswith(".log")]
        if len(entries) > keep:
            entries.sort(key=lambda p: os.path.getmtime(p), reverse=True)
            for stale in entries[keep:]:
                try: os.remove(stale)
                except Exception: pass
    except Exception: pass
def write_json_atomic(path: str, data: typing.Any, indent: typing.Optional[int]=None):
    tmp_path = f"{path}.{os.getpid()}.{threading.get_ident()}.tmp"
    try:
        with open(tmp_path, "w", encoding="utf-8") as f: json.dump(data, f, indent=indent)
        for attempt in range(3):
            try: os.replace(tmp_path, path); return
            except PermissionError:
                if attempt == 2: raise
                time.sleep(0.03)
    finally:
        if os.path.exists(tmp_path):
            try: os.remove(tmp_path)
            except Exception: pass
def parse_version_key(ver: str) -> tuple:
    if not ver: return ((0, 0, 0), 1, "")
    m = re.match(r"^\s*v?(\d+(?:\.\d+)*)([a-zA-Z0-9._\-]*)\s*$", str(ver))
    if not m: return ((0, 0, 0), 1, str(ver))
    nums = tuple(int(x) for x in m.group(1).split("."))
    while len(nums) < 3: nums = nums + (0,)
    suffix = m.group(2) or ""
    return (nums, 0 if suffix else 1, suffix)
def is_newer_version(candidate: str, current: str) -> bool: return parse_version_key(candidate) > parse_version_key(current)
def trace(e: Exception):
    if not debug_mode: return str(e)
    _, tb_v, tb_b = sys.exc_info()
    tb_lines = traceback.extract_tb(tb_b)
    lines = []
    lines.append(PyKits.Terminal.foreground("Traceback (most recent call last):", color="Magenta", bright=True))
    for fn, ln, f, tx in tb_lines:
        fn_c = PyKits.Terminal.foreground(fn, color="Magenta", bright=True)
        ln_c = PyKits.Terminal.foreground(str(ln), color="Magenta", bright=True)
        f_c = PyKits.Terminal.foreground(f, color="Magenta", bright=True)
        lines.append(f"  File {fn_c}, line {ln_c}, in {f_c}")
        if tx: lines.append(f"    {tx}")
    exc_t = PyKits.Terminal.foreground(PyKits.Terminal.bold(f"{type(tb_v).__name__}:"), color="Magenta", bright=True)
    exc_m = PyKits.Terminal.foreground(str(tb_v), color="Magenta", bright=False)
    lines.append(f"{exc_t} {exc_m}")
    return "\n".join(lines)
def _rebind_terminal_streams():
    global terminal
    if not app_logger: return
    terminal = PyKits.Terminal(logger=app_logger, log_level=logging.INFO)
    err_terminal = PyKits.Terminal(logger=app_logger, log_level=logging.ERROR)
    terminal.inject_stdout()
    err_terminal.inject_stderr()
def read_early_debug_mode() -> bool:
    if "--debug" in sys.argv or os.environ.get("NWC_DEBUG") == "1": return True
    config_path = os.path.join(app_data_path, "ConnectConfig.json")
    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f: data = json.load(f)
            return bool(data.get("debug_mode", False))
        except Exception: pass
    return False
def setup_logging():
    global terminal, app_logger
    handler_name = "Main"
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
    app_logger = logger
    _rebind_terminal_streams()
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
@functools.lru_cache(maxsize=8)
def getImageObj(img: str=None) -> Image.Image:
    try: return Image.open(img)
    except Exception: return Image.new("RGB", (64, 64), color="blue")
def notification(title: str, message: str):
    def a():
        if pystray_icon: pystray_icon.notify(message, title)
    addToThread(a, ())
def set_icon(img: str):
    global current_icon_name
    with icon_lock:
        current_icon_name = img
        if pystray_icon and tray_ready.is_set():
            image = getImageObj(os.path.join(program_files, "Resources", img))
            pystray_icon.icon = image
            update_tray()
def setup():
    pystray_icon.visible = True
    with icon_lock:
        tray_ready.set()
        if current_icon_name != "app_icon.ico":
            image = getImageObj(os.path.join(program_files, "Resources", current_icon_name))
            pystray_icon.icon = image
            update_tray()
def needing_state(): system_class.needing_state()
def unneeding_state(): system_class.unneeding_state()
def quit_app():
    global stop_app
    global pause_until_monotonic
    global paused_target
    pause_until_monotonic = 0.0
    paused_target = None
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
    window.geometry(f"+{x}+{y}")
def apply_window_icon(win: tk.Misc, default: bool=False):
    for base in (program_files, os.path.dirname(os.path.abspath(__file__))):
        p = os.path.join(base, "Resources", "app_icon.png")
        if os.path.exists(p):
            try:
                img = tk.PhotoImage(file=p)
                win._icon_img = img
                win.iconphoto(default, img)
                return
            except Exception: pass
def get_server_location(server: dict) -> tuple[str, str]:
    if not isinstance(server, dict): return "", ""
    locs = server.get("locations")
    loc = locs[0] if isinstance(locs, list) and locs and isinstance(locs[0], dict) else {}
    country_obj = loc.get("country") if isinstance(loc.get("country"), dict) else {}
    city_obj = country_obj.get("city") if isinstance(country_obj.get("city"), dict) else {}
    country_name = country_obj.get("name") or ""
    city_name = city_obj.get("name") or country_name
    return country_name, city_name
def get_wireguard_public_key(server: dict) -> typing.Optional[str]:
    if not isinstance(server, dict): return None
    for m in server.get("technologies", []):
        if not isinstance(m, dict) or m.get("id") != 35: continue
        pivot = m.get("pivot")
        if isinstance(pivot, dict) and pivot.get("status", "online") != "online": return None
        meta = m.get("metadata")
        if not isinstance(meta, list): return None
        for item in meta:
            if isinstance(item, dict) and item.get("name") == "public_key" and item.get("value"): return item.get("value")
        for item in meta:
            if isinstance(item, dict) and item.get("value"): return item.get("value")
    return None
def server_has_group_id(server: dict, group_id: typing.Optional[int]) -> bool:
    if group_id is None: return True
    if not isinstance(server, dict): return False
    return any(isinstance(g, dict) and g.get("id") == group_id for g in server.get("groups", []))
def get_group_query_filter(group_key: str=None) -> str:
    key = group_key if group_key is not None else config_data.get("server_group", "all")
    gid = nord_server_groups.get(key, (None, None))[1]
    return f"&filters[servers_groups]={gid}" if gid is not None else ""
def encrypt_data(key_name: str, data: str) -> str: return system_class.encrypt_data(key_name, data)
def decode_data(key_name: str, encrypted: str) -> str: return system_class.decrypt_data(key_name, encrypted)
def _build_lazy_server_nodes(raw_servers: list) -> list[dict]:
    active_group_id = nord_server_groups.get(config_data.get("server_group", "all"), (None, None))[1]
    country_buckets = {}
    specialty_buckets = {g_key: {} for g_key, (_, g_id) in nord_server_groups.items() if g_id is not None}
    for s in raw_servers:
        if not isinstance(s, dict) or s.get("status") != "online" or s.get("load", 0) == 0: continue
        if not get_wireguard_public_key(s): continue
        country_name, _ = get_server_location(s)
        if not country_name: continue
        for g_key, (_, g_id) in nord_server_groups.items():
            if g_id is not None and server_has_group_id(s, g_id): specialty_buckets[g_key].setdefault(country_name, []).append(s)
        if active_group_id is None or server_has_group_id(s, active_group_id): country_buckets.setdefault(country_name, []).append(s)
    def _server_leaf(s: dict) -> dict:
        name = s["name"]
        return {"label": name, "tunnel_id": name.replace(" ", "").replace("#", ""), "exact_mode": "server", "children": None}
    def _build_city_children(city_raw: str, city_srvs: list) -> list[dict]:
        if not check_asn_server_list(): return [{"label": "Auto", "tunnel_id": city_raw, "exact_mode": "city", "children": None}] + [_server_leaf(s) for s in city_srvs]
        asn_groups = {}
        was_loaded = cached_asn_data["loaded"]
        before_len = len(server_asn_map)
        for s in city_srvs:
            asn_name = get_asn_name_from_ip(s.get("station")) or "Unknown ASN"
            asn_groups.setdefault(asn_name, []).append(_server_leaf(s))
        if cached_asn_data["loaded"] and not was_loaded:
            if len(server_asn_map) > before_len:
                try: write_json_atomic(os.path.join(app_data_path, "NordServerASNMap.json"), server_asn_map)
                except Exception: pass
            unload_asn_index()
        nodes = [{"label": "Auto", "tunnel_id": city_raw, "exact_mode": "city", "children": None}]
        for asn_name in sorted(asn_groups.keys()):
            nodes.append({"label": asn_name, "tunnel_id": f"{city_raw}|||{asn_name}", "exact_mode": "asn", "children": [{"label": "Auto", "tunnel_id": f"{city_raw}|||{asn_name}", "exact_mode": "asn", "children": None}] + asn_groups[asn_name]})
        return nodes
    def _build_country_children(country_name: str) -> list[dict]:
        cities = {}
        for s in country_buckets.get(country_name, []):
            _, city_name = get_server_location(s)
            v_lbl = city_name + (" - Virtual" if get_if_virtual_location(s) else "")
            if v_lbl not in cities: cities[v_lbl] = (city_name, [])
            cities[v_lbl][1].append(s)
        nodes = [{"label": "Auto", "tunnel_id": country_name, "exact_mode": "country", "children": None}]
        for v_lbl in sorted(cities.keys()):
            c_raw, c_srvs = cities[v_lbl]
            nodes.append({"label": v_lbl, "tunnel_id": c_raw, "exact_mode": "city", "children": None, "children_fn": lambda cr=c_raw, cs=c_srvs: _build_city_children(cr, cs)})
        return nodes
    def _build_specialty_country_children(g_key: str, country_name: str) -> list[dict]: return [{"label": "Auto", "tunnel_id": f"{g_key}|||{country_name}", "exact_mode": "group", "children": None}] + [_server_leaf(s) for s in specialty_buckets[g_key].get(country_name, [])]
    def _build_specialty_group_children(g_key: str) -> list[dict]:
        nodes = [{"label": "Auto", "tunnel_id": g_key, "exact_mode": "group", "children": None}]
        for c_name in sorted(specialty_buckets[g_key].keys()): nodes.append({"label": c_name, "tunnel_id": f"{g_key}|||{c_name}", "exact_mode": "group", "children": None, "children_fn": lambda gk=g_key, cn=c_name: _build_specialty_country_children(gk, cn)})
        return nodes
    def _build_specialty_root_children() -> list[dict]:
        nodes = [{"label": "All Servers", "tunnel_id": "auto", "exact_mode": None, "children": None}]
        for g_key in sorted((k for k, (_, gid) in nord_server_groups.items() if gid is not None and specialty_buckets.get(k)), key=lambda k: nord_server_groups[k][0]): nodes.append({"label": nord_server_groups[g_key][0], "tunnel_id": g_key, "exact_mode": "group", "children": None, "children_fn": lambda gk=g_key: _build_specialty_group_children(gk)})
        return nodes
    root_nodes = [{"label": "Auto", "tunnel_id": "auto", "exact_mode": None, "children": None}]
    if any(specialty_buckets.values()): root_nodes.append({"label": "Specialty Groups", "tunnel_id": "auto", "exact_mode": None, "children": None, "children_fn": _build_specialty_root_children})
    for c_name in sorted(country_buckets.keys()): root_nodes.append({"label": c_name, "tunnel_id": c_name, "exact_mode": "country", "children": None, "children_fn": lambda cn=c_name: _build_country_children(cn)})
    return root_nodes
def select_server_list(message: str="Please select a server!", servers: list=None) -> typing.Union[None, dict]:
    if servers is None:
        raw_servers = load_cached_servers()
        if not raw_servers: return
        if check_asn_server_list(): load_server_asn_map()
        root_nodes = _build_lazy_server_nodes(raw_servers)
    else: server_dictionary = servers
    win = tk.Toplevel()
    win.title("Select NordVPN Server")
    win.geometry("320x120")
    win.resizable(False, False)
    win.attributes("-topmost", True)
    center_tkinter_window(win)
    apply_window_icon(win, False)
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
            next_ch = node.get("children")
            if next_ch is None and callable(node.get("children_fn")):
                next_ch = node["children_fn"]()
                node["children"] = next_ch
            if next_ch: add_row(next_ch)
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
            for lbl, tid, em in data: nodes.append({"label": lbl, "tunnel_id": tid, "exact_mode": em, "children": None})
        elif type(data) is dict:
            for key in sorted(data.keys(), key=lambda x: ("\0" if x.lower() == "auto" else ("\x01" if x == "Specialty Groups" else x))):
                val = data[key]
                if key.lower() == "auto":
                    for label, tunnel, exact in val: nodes.append({"label": label, "tunnel_id": tunnel if tunnel else label, "exact_mode": exact, "children": None})
                    continue
                if type(val) is list: info = val[0]
                elif type(val) is dict: info = val.get("Auto", [("Auto", key, "country")])[0]
                chi = format_items(val, root=False)
                nodes.append({"label": key, "tunnel_id": info[1], "exact_mode": info[2], "children": chi if chi else None})
        return nodes
    add_row(root_nodes if servers is None else format_items(server_dictionary))
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
    if not selected_server: return
    addToThread(connect_server, (selected_server["tunnel_id"], selected_server["exact_mode"]), connection=True)
    update_tray()
def connect_history_server_button():
    generated_servers = {}
    if not config_data.get("server_history"): config_data["server_history"] = []
    for s in config_data.get("server_history", []):
        if not isinstance(s, dict): continue
        country = s.get("country")
        city = s.get("city")
        server_name = s.get("name")
        date = s.get("time")
        if not country or not city or not server_name or not isinstance(date, (int, float)): continue
        try: dt = datetime.datetime.fromtimestamp(date).astimezone()
        except Exception: continue
        day = convert_to_local(dt).strftime("%Y-%m-%d")
        hm = convert_to_local(dt).strftime("%H:%M")
        if day not in generated_servers: generated_servers[day] = {}
        if hm not in generated_servers[day]: generated_servers[day][hm] = []
        generated_servers[day][hm].insert(0, (f"{server_name} ({city}, {country})", server_name.replace(" ", "").replace("#", ""), "server"))
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
            if latest_version.ok and isinstance(latest_version.json, dict):
                res = latest_version.json
                for v in res.get("versions", []):
                    if isinstance(v, dict) and v.get("version") and v.get("beta", False) == config_data.get("beta_updates", False): return v
    except Exception as e: errorMessage(f"Error while trying to check for updates: {trace(e)}")
def get_latest_wireguard_version() -> dict:
    try:
        if tunnel_check(timeout=1):
            latest_version = requests.get("https://build.wireguard.com/distros.json")
            if latest_version.ok and isinstance(latest_version.json, dict):
                res = latest_version.json
                return system_class.format_for_latest_wireguard(res)
    except Exception as e: errorMessage(f"Error while trying to check for updates: {trace(e)}")
def get_current_wireguard_version() -> typing.Union[str, None]: return system_class.get_current_wireguard_version()
def auto_check_for_updates():
    if session_data.get("connection_text", "").startswith("Connecting"):
        while session_data.get("connection_text", "").startswith("Connecting"): time.sleep(0.2)
    wait_until_connected()
    latest_version = get_latest_version()
    if latest_version and is_newer_version(latest_version["version"], version): 
        notification("Update Available!", f"NordWireConnect v{latest_version['version']} is now available for download!")
    def wireguard_update():
        latest_wireguard_version = get_latest_wireguard_version()
        current_wireguard_version = get_current_wireguard_version()
        if not latest_wireguard_version: return
        if not current_wireguard_version: return
        if is_newer_version(latest_wireguard_version["version"], current_wireguard_version): 
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
    global stop_app
    try:
        mainMessage(f"Update v{latest_ver} is now available!")
        confirmed = messagebox.askyesno("Update Available!", f"NordWireConnect v{latest_ver} is now available to be downloaded from GitHub!\nDo you wish to proceed?", icon="info")
        if not confirmed: return
        installer_path = os.path.join(app_data_path, "NordWireConnectInstaller.exe")
        download_progress = requests.download(f"https://github.com/EfazDev/nordwireconnect/releases/download/v{latest_ver}/NordWireConnectInstaller.exe", installer_path)
        if download_progress.ok:
            if config_data.get("auto_connect", True): 
                session_data["connected"] = False
                disconnect(only_disconnect=True)
            else:
                disconnect()
                sess = os.path.join(app_data_path, "NordSessionData.json")
                if os.path.exists(sess): os.remove(sess)
            res = system_class.run_sudo_cmd([installer_path])
            if res > 32:
                send_command("ui-closing", timeout=5)
                if pystray_icon: pystray_icon.stop()
                stop_app = True
            else: warnMessage(f"Installer launch was canceled or failed (ShellExecuteW={res}).")
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
        if is_newer_version(latest_version["version"], version) or debug_mode != latest_version.get("debug", False): 
            latest_ver = latest_version["version"]
            test_available = requests.head(f"https://github.com/EfazDev/nordwireconnect/releases/download/v{latest_ver}/NordWireConnectInstaller.exe")
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
        if is_newer_version(latest_version["version"], current_version): 
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
            keep_keys = {"access_token", "encrypted_private_key", "username", "openvpn_username", "openvpn_password", "nordvpn_email", "data_encrypted"}
            for k in list(config_data.keys()):
                if k not in keep_keys: config_data[k] = copy.deepcopy(base_config.get(k))
            write_json_atomic(config_path, config_data, indent=4)
            load_configuration()
            update_tray()
    except Exception as e: errorMessage(f"Unable to clear configuration data: {trace(e)}")
def get_current_time() -> datetime.datetime: return datetime.datetime.now(tz=datetime.timezone.utc)
def convert_to_local(datetime: datetime.datetime) -> datetime.datetime: return datetime.astimezone()
def ipv4_to_int(ip: str) -> int: return int.from_bytes(socket.inet_aton(ip), "big")
def _lookup_asn_index_ip(ip: str) -> str:
    try:
        if not cached_asn_data["loaded"] or not ip: return ""
        try: ip_int = ipv4_to_int(ip)
        except Exception: ip_int = int(ipaddress.ip_address(ip))
        i = bisect.bisect_right(cached_asn_data["starts"], ip_int) - 1
        if i >= 0 and ip_int <= cached_asn_data["ends"][i]: return cached_asn_data["unique_names"][cached_asn_data["name_ids"][i]]
    except Exception: pass
    return ""
def load_server_asn_map():
    global server_asn_map_loaded
    map_path = os.path.join(app_data_path, "NordServerASNMap.json")
    if os.path.exists(map_path):
        try:
            with open(map_path, "r", encoding="utf-8") as f: data = json.load(f)
            if isinstance(data, dict): server_asn_map.clear(); server_asn_map.update(data)
        except Exception: pass
    server_asn_map_loaded = True
def unload_asn_index():
    with asn_lock:
        cached_asn_data["starts"] = array.array("I")
        cached_asn_data["ends"] = array.array("I")
        cached_asn_data["name_ids"] = array.array("I")
        cached_asn_data["unique_names"] = []
        cached_asn_data["loaded"] = False
def premap_server_asns(servers: list=None, force: bool=False, unload_after: bool=True):
    global server_asn_map_loaded
    if not server_asn_map_loaded: load_server_asn_map()
    srvs = servers if servers is not None else load_cached_servers()
    station_ips = {s.get("station") for s in srvs if isinstance(s, dict) and isinstance(s.get("station"), str) and "." in s.get("station")}
    if not station_ips: return
    missing = any(ip not in server_asn_map for ip in station_ips)
    if force or missing:
        ensure_asn_loaded(force=force)
        if cached_asn_data["loaded"]:
            updated = {ip: _lookup_asn_index_ip(ip) for ip in sorted(station_ips)}
            server_asn_map.clear()
            server_asn_map.update(updated)
            server_asn_map_loaded = True
            try: write_json_atomic(os.path.join(app_data_path, "NordServerASNMap.json"), server_asn_map)
            except Exception: pass
    if unload_after: unload_asn_index()
def get_asn_name_from_ip(ip: str) -> typing.Union[str, None]:
    if not ip: return None
    if not server_asn_map_loaded: load_server_asn_map()
    if ip in server_asn_map:
        val = server_asn_map[ip]
        return val if val else None
    if not cached_asn_data["loaded"]: ensure_asn_loaded(force=False)
    if not cached_asn_data["loaded"]: return None
    val = _lookup_asn_index_ip(ip)
    server_asn_map[ip] = val
    return val if val else None
def get_vpn_renewal_data() -> dict:
    try:
        if tunnel_check(timeout=1):
            if not config_data.get("access_token", ""): return {"active": False, "expires_at": None, "raw_expiry": "0000-00-00 00:00:00"}
            decrypted_token = decode_data("access_token", config_data.get("access_token", ""))
            service_data = requests.get(f"https://api.nordvpn.com/v1/users/services", auth=["token", decrypted_token])
            if service_data.ok and isinstance(service_data.json, list):
                for s in service_data.json:
                    if isinstance(s, dict) and isinstance(s.get("service"), dict) and s["service"].get("identifier") == "vpn":
                        raw_exp = s.get("expires_at") or "0000-00-00 00:00:00"
                        expiry = convert_nordtime_to_datetime(raw_exp)
                        if expiry:
                            return {
                                "active": True,
                                "expires_at": expiry,
                                "raw_expiry": raw_exp
                            }
    except Exception as e: errorMessage(f"Unable to get VPN plan renewal data: {trace(e)}")
    return {"active": False, "expires_at": None, "raw_expiry": "0000-00-00 00:00:00"}
def get_if_within_reminder_period(expiry) -> typing.Union[int, None]:
    exp_dt = convert_nordtime_to_datetime(expiry)
    if not exp_dt: return -1
    diff = exp_dt - get_current_time()
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
        total = usage["download"] + usage["upload"]
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
    try: write_json_atomic(os.path.join(app_data_path, "NordSessionData.json"), session_data, indent=4)
    except Exception: pass
@functools.lru_cache(maxsize=8)
def convert_nordtime_to_datetime(time_str: str) -> typing.Optional[datetime.datetime]:
    if not time_str or not isinstance(time_str, str) or time_str.startswith("0000-"): return None
    try: return datetime.datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S").replace(tzinfo=datetime.timezone.utc)
    except Exception: return None
def _flatten_asn_intervals(raw_ranges: list[tuple[int, int, str]]) -> tuple[list[int], list[int], list[str]]:
    if not raw_ranges: return [], [], []
    raw_ranges.sort(key=lambda x: (x[0], -(x[1] - x[0])))
    starts, ends, names = [], [], []
    stack = []
    pos = 0
    for s, e, name in raw_ranges:
        while stack and stack[-1][0] < s:
            top_end, top_name = stack.pop()
            if pos <= top_end:
                starts.append(pos)
                ends.append(top_end)
                names.append(top_name)
                pos = top_end + 1
        if stack and pos < s:
            starts.append(pos)
            ends.append(s - 1)
            names.append(stack[-1][1])
        pos = max(pos, s)
        stack.append((e, name))
    while stack:
        top_end, top_name = stack.pop()
        if pos <= top_end:
            starts.append(pos)
            ends.append(top_end)
            names.append(top_name)
            pos = top_end + 1
    return starts, ends, names
def _compile_asn_to_index(starts: list[int], ends: list[int], names: list[str], idx_path: str) -> bool:
    count = len(starts)
    if count < 1000: return False
    unique_names = []
    name_to_id = {}
    id_list = []
    for n in names:
        nid = name_to_id.get(n)
        if nid is None:
            nid = len(unique_names)
            name_to_id[n] = nid
            unique_names.append(n)
        id_list.append(nid)
    starts_arr = array.array("I", starts)
    ends_arr = array.array("I", ends)
    ids_arr = array.array("I", id_list)
    names_blob = "\n".join(unique_names).encode("utf-8")
    header = struct.pack("<4sII", b"NWC1", count, len(names_blob))
    w_starts, w_ends, w_ids = starts_arr, ends_arr, ids_arr
    if sys.byteorder != "little":
        w_starts = array.array("I", starts_arr); w_starts.byteswap()
        w_ends = array.array("I", ends_arr); w_ends.byteswap()
        w_ids = array.array("I", ids_arr); w_ids.byteswap()
    tmp_path = f"{idx_path}.{os.getpid()}.{threading.get_ident()}.tmp"
    try:
        with open(tmp_path, "wb") as f:
            f.write(header)
            f.write(names_blob)
            f.write(w_starts.tobytes())
            f.write(w_ends.tobytes())
            f.write(w_ids.tobytes())
        os.replace(tmp_path, idx_path)
    finally:
        if os.path.exists(tmp_path):
            try: os.remove(tmp_path)
            except Exception: pass
    cached_asn_data["starts"] = starts_arr
    cached_asn_data["ends"] = ends_arr
    cached_asn_data["name_ids"] = ids_arr
    cached_asn_data["unique_names"] = unique_names
    cached_asn_data["loaded"] = True
    return True
def _load_asn_index(idx_path: str) -> bool:
    try:
        with open(idx_path, "rb") as f: data = f.read()
        if len(data) < 12: return False
        magic, count, names_len = struct.unpack("<4sII", data[:12])
        if magic != b"NWC1" or count < 1000 or len(data) != 12 + names_len + (count * 12): return False
        pos = 12
        unique_names = data[pos:pos + names_len].decode("utf-8").split("\n")
        pos += names_len
        arr_bytes = count * 4
        starts_arr = array.array("I"); starts_arr.frombytes(data[pos:pos + arr_bytes]); pos += arr_bytes
        ends_arr = array.array("I"); ends_arr.frombytes(data[pos:pos + arr_bytes]); pos += arr_bytes
        ids_arr = array.array("I"); ids_arr.frombytes(data[pos:pos + arr_bytes])
        if sys.byteorder != "little":
            starts_arr.byteswap(); ends_arr.byteswap(); ids_arr.byteswap()
        cached_asn_data["starts"] = starts_arr
        cached_asn_data["ends"] = ends_arr
        cached_asn_data["name_ids"] = ids_arr
        cached_asn_data["unique_names"] = unique_names
        cached_asn_data["loaded"] = True
        return True
    except Exception: return False
def ensure_asn_loaded(force: bool=False):
    with asn_lock:
        if cached_asn_data["loaded"] and not force: return
        idx_path = os.path.join(app_data_path, "asn-ipv4.idx")
        for stale_name in ("ip2asn-v4.tsv.gz", "NordASNCache.json", "asn-ipv4.db"):
            stale_path = os.path.join(app_data_path, stale_name)
            if os.path.exists(stale_path):
                try: os.remove(stale_path)
                except Exception: pass
        is_fresh = False
        if not force and os.path.exists(idx_path):
            try: is_fresh = (time.time() - os.path.getmtime(idx_path)) <= (7 * 86400)
            except Exception: is_fresh = False
            if is_fresh:
                if _load_asn_index(idx_path): return
                try: os.remove(idx_path)
                except Exception: pass
        if tunnel_check(timeout=2.5):
            asn_sources = (
                "https://github.com/sapics/ip-location-db/releases/download/latest/dbip-asn-ipv4.csv",
                "https://iptoasn.com/data/ip2asn-v4.tsv.gz"
            )
            dl_headers = {"Accept-Encoding": "identity", "User-Agent": f"NordWireConnect/{version}"}
            for src in asn_sources:
                tmp_path = f"{idx_path}.{os.getpid()}.dl"
                try:
                    dl = requests.download(src, tmp_path, headers=dl_headers, timeout=25)
                    if dl and dl.ok and os.path.exists(tmp_path) and os.path.getsize(tmp_path) > 100000:
                        all_asns = system_class.unarchive_asn_database(tmp_path)
                        raw_ranges = []
                        for d in all_asns:
                            try:
                                s_int, e_int = ipv4_to_int(d["s_ip"]), ipv4_to_int(d["e_ip"])
                                asn_name = d.get("asn_name")
                                if asn_name and s_int <= e_int: raw_ranges.append((s_int, e_int, asn_name))
                            except Exception: continue
                        if len(raw_ranges) >= 1000:
                            starts, ends, names = _flatten_asn_intervals(raw_ranges)
                            if _compile_asn_to_index(starts, ends, names, idx_path): return
                except Exception as e: warnMessage(f"Unable to load ASN database from {src}: {e}")
                finally:
                    if os.path.exists(tmp_path):
                        try: os.remove(tmp_path)
                        except Exception: pass
        if os.path.exists(idx_path) and _load_asn_index(idx_path): return
        warnMessage("Unable to load ASN database; servers may show Unknown ASN.")
def run_cache(force: bool=False):
    if not cache_lock.acquire(blocking=False): return
    try:
        server_cache_path = os.path.join(app_data_path, "NordServerCache.json")
        country_cache_path = os.path.join(app_data_path, "NordCountryCache.json")
        city_cache_path = os.path.join(app_data_path, "NordCityCache.json")
        id_cache_path = os.path.join(app_data_path, "NordServerIDCache.json")
        if not force and all(os.path.exists(p) for p in (server_cache_path, country_cache_path, city_cache_path, id_cache_path)):
            try:
                if (time.time() - os.path.getmtime(server_cache_path)) < 900:
                    if check_asn_server_list(): premap_server_asns(unload_after=True)
                    return
            except Exception: pass
        if not force:
            conn_deadline = time.monotonic() + 25.0
            while session_data.get("connection_text", "").startswith("Connecting"):
                if stop_app or time.monotonic() > conn_deadline: return
                time.sleep(0.3)
        mainMessage("Updating server cache..")
        wait_deadline = time.monotonic() + 15.0
        while not tunnel_check(timeout=1.5):
            if stop_app or time.monotonic() > wait_deadline: return
            time.sleep(0.5)
        all_servers = requests.get(f"https://api.nordvpn.com/v1/servers/recommendations?{server_filter}{wireguard_tech_filter}{get_coordinates_query()}&limit=10000", timeout=20)
        raw_all_servers = all_servers.json if (all_servers.ok and isinstance(all_servers.json, list)) else []
        raw_servers = (raw_all_servers[:500] if config_data.get("optimize_server_list", False) else raw_all_servers)
        if not raw_all_servers or not raw_servers:
            mainMessage("Unable to get current servers from NordVPN API. Depending on stored cache.")
            return
        all_cities = {}
        all_countries_resp = requests.get("https://api.nordvpn.com/v1/countries", timeout=15)
        all_countries = all_countries_resp.json if (all_countries_resp.ok and isinstance(all_countries_resp.json, list)) else []
        all_servers_by_id = {}
        if not all_countries:
            mainMessage("Unable to get current list of Countries from NordVPN API. Depending on stored cache.")
            if os.path.exists(country_cache_path):
                with open(country_cache_path, "r", encoding="utf-8") as f: all_countries = json.load(f)
            else: return
        for s in raw_all_servers:
            for l in s.get("locations", []):
                country = l.get("country") if isinstance(l, dict) else None
                c = country.get("city") if isinstance(country, dict) else None
                if isinstance(c, dict) and "id" in c and "name" in c: all_cities[c["id"]] = c["name"]
            if "id" in s and "name" in s: all_servers_by_id[s["name"].replace(" ", "").replace("#", "")] = s["id"]
        if not all_cities:
            for country in all_countries:
                if not isinstance(country, dict): continue
                for c in country.get("cities") or []:
                    if isinstance(c, dict) and "id" in c and "name" in c: all_cities[c["id"]] = c["name"]
        all_cities_list = [{"id": i, "name": n} for i, n in all_cities.items()]
        write_json_atomic(server_cache_path, raw_servers)
        server_cache_mem["servers"] = raw_servers
        try: server_cache_mem["mtime"] = os.path.getmtime(server_cache_path)
        except Exception: server_cache_mem["mtime"] = 0.0
        write_json_atomic(country_cache_path, all_countries)
        write_json_atomic(city_cache_path, all_cities_list)
        write_json_atomic(id_cache_path, all_servers_by_id)
        if check_asn_server_list(): premap_server_asns(raw_servers, force=force, unload_after=True)
        mainMessage("Server cache updated successfully.")
    finally: cache_lock.release()
def send_command(cmd: str, args: dict=None, timeout: float=20.0) -> dict: 
    payload = {"cmd": cmd}
    if args: payload.update(args)
    res = nord_socket.request("NordWireConnectCommands", payload, timeout=timeout)
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
        if abs(n) < 1024: return f"{n:.2f} {unit}"
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
    inputted = simpledialog.askstring(title="DNS Input", prompt="Enter IPv4 DNS servers to use (comma separated):", initialvalue=config_data.get("dns", "103.86.96.100, 103.86.99.100"))
    if inputted is None: return
    ips = inputted.replace(" ", "").split(",")
    added = [ip for ip in ips if ip and requests.get_if_ip(ip)]
    if added:
        config_data["dns"] = ", ".join(added)
        save_configuration()
        successMessageBox(f"Successfully changed DNS to: {', '.join(added)}", "NordWireConnect")
        if session_data["connected"] == True: differ_from_status_action2()
        update_tray()
    else: errorMessage("No valid DNS servers were provided.")
def fetch_and_cache_credentials(decrypted_token: str) -> bool:
    global private_key
    creds = requests.get("https://api.nordvpn.com/v1/users/services/credentials", auth=["token", decrypted_token], timeout=12)
    if creds.ok and isinstance(creds.json, dict):
        pk = creds.json.get("nordlynx_private_key", "")
        if pk:
            private_key = pk
            config_data["encrypted_private_key"] = encrypt_data("private_key", pk)
            save_configuration()
            return True
    return False
def fetch_cybersec_threat_protection_ips() -> typing.List[str]:
    global cached_cybersec_ips
    now = time.time()
    if cached_cybersec_ips["ips"] and (now - cached_cybersec_ips["updated_at"]) < 1800: return list(cached_cybersec_ips["ips"])
    try:
        response = requests.get("https://downloads.nordcdn.com/configs/dns/cybersec.json", timeout=10)
        if response.ok and isinstance(response.json, dict) and isinstance(response.json.get("servers"), list):
            ips = [ip for ip in response.json.get("servers") if requests.get_if_ip(ip)]
            if ips:
                cached_cybersec_ips = {"ips": ips, "updated_at": now}
                return list(ips)
    except Exception as e: errorMessage(f"Unable to fetch Threat Protection DNS IPs: {trace(e)}")
    return ["103.86.96.108", "103.86.99.108"]
def resolve_tunnel_dns() -> str:
    if check_threat_protection_lite():
        ips = list(fetch_cybersec_threat_protection_ips())
        if len(ips) > 1: random.shuffle(ips)
        return ", ".join(ips)
    raw_dns = config_data.get("dns", "103.86.96.100, 103.86.99.100") or "103.86.96.100, 103.86.99.100"
    tokens = [x.strip() for x in raw_dns.split(",") if x.strip()]
    if len(tokens) > 1 and all(t in ("103.86.96.100", "103.86.99.100") for t in tokens):
        random.shuffle(tokens)
        return ", ".join(tokens)
    return raw_dns
def resolve_nordlynx_mtu() -> int:
    global active_tunnel_mtu
    cur_mtu = config_data.get("mtu", None)
    if isinstance(cur_mtu, int) and 1280 <= cur_mtu <= 1500: mtu_val = cur_mtu
    else:
        gw_mtu = 1500
        try: gw_mtu = system_class.get_physical_gateway_mtu()
        except Exception: pass
        mtu_val = max(1280, min(1420, gw_mtu - 80))
    active_tunnel_mtu = mtu_val
    return mtu_val
def get_active_conf_dns() -> str:
    conf_path = os.path.join(app_data_path, "NordWireConnect.conf")
    if os.path.exists(conf_path):
        try:
            with open(conf_path, "r", encoding="utf-8") as f:
                for ln in f:
                    if ln.strip().lower().startswith("dns"):
                        parts = ln.split("=", 1)
                        if len(parts) == 2 and parts[1].strip(): return parts[1].strip()
        except Exception: pass
    return resolve_tunnel_dns()
def get_active_conf_mtu() -> int:
    conf_path = os.path.join(app_data_path, "NordWireConnect.conf")
    if os.path.exists(conf_path):
        try:
            with open(conf_path, "r", encoding="utf-8") as f:
                for ln in f:
                    if ln.strip().lower().startswith("mtu"):
                        parts = ln.split("=", 1)
                        if len(parts) == 2 and parts[1].strip().isdigit():
                            val = int(parts[1].strip())
                            if 1280 <= val <= 1500: return val
        except Exception: pass
    return resolve_nordlynx_mtu()
def _df_ping_check(host: str, payload: int, timeout_ms: int=600) -> tuple[bool, bool]:
    try:
        res = System.shell_run(["ping", host, "-n", "1", "-w", str(timeout_ms), "-f", "-l", str(payload)], timeout=2)
        out = ((res.stdout or "") + "\n" + (res.stderr or "")).lower()
        frag = ("fragment" in out) or ("too big" in out) or (res.returncode != 0 and "df" in out)
        ok = (res.returncode == 0) and ("ttl=" in out) and (not frag)
        return ok, frag
    except Exception: return False, False
def probe_tunnel_path_mtu(initial_mtu: int) -> int:
    if initial_mtu <= 1280: return 1280
    target_host = "1.1.1.1"
    ok, frag = _df_ping_check("1.1.1.1", initial_mtu - 28)
    if ok: return initial_mtu
    if not frag:
        ok2, frag2 = _df_ping_check("8.8.8.8", initial_mtu - 28)
        if ok2: return initial_mtu
        frag = frag or frag2
        if frag2: target_host = "8.8.8.8"
    if not frag:
        ok_min, _ = _df_ping_check("1.1.1.1", 1280 - 28)
        if ok_min: target_host = "1.1.1.1"
        else:
            ok_min2, _ = _df_ping_check("8.8.8.8", 1280 - 28)
            if ok_min2: target_host = "8.8.8.8"
            else: return initial_mtu
        frag = True
    for cand in (1412, 1400, 1392, 1380, 1360, 1348, 1320, 1280):
        if cand >= initial_mtu: continue
        ok_c, _ = _df_ping_check(target_host, cand - 28)
        if ok_c: return cand
    return 1280 if frag else initial_mtu
def clamp_tunnel_path_mtu(initial_mtu: int) -> int:
    global active_tunnel_mtu
    if config_data.get("mtu") is not None: return initial_mtu
    best_mtu = probe_tunnel_path_mtu(initial_mtu)
    if best_mtu < initial_mtu:
        if send_command("set-tunnel-mtu", {"mtu": best_mtu}, timeout=8).get("suc") == True:
            active_tunnel_mtu = best_mtu
            mainMessage(f"Clamped WireGuard tunnel MTU to {best_mtu} via Path MTU probe.")
            return best_mtu
    active_tunnel_mtu = initial_mtu
    return initial_mtu
def change_access_token():
    global private_key
    inputted = simpledialog.askstring(title="Access Token Input", prompt="Enter your generated NordVPN Access Token from the NordVPN website (https://my.nordaccount.com/dashboard/nordvpn/):", show="*")
    if inputted is None: return
    inputted = inputted.strip()
    if len(inputted) > 50:
        test_request = requests.get("https://api.nordvpn.com/v1/users/services/credentials", auth=["token", inputted], timeout=12)
        if test_request.ok and isinstance(test_request.json, dict):
            test_request2 = requests.get("https://api.nordvpn.com/v1/users/current", auth=["token", inputted], timeout=12)
            if test_request2.ok and isinstance(test_request2.json, dict):
                test_request = test_request.json
                test_request2 = test_request2.json
                config_data["openvpn_username"] = encrypt_data("openvpn_username", test_request.get("username", ""))
                config_data["openvpn_password"] = encrypt_data("openvpn_password", test_request.get("password", ""))
                config_data["nordvpn_email"] = test_request2.get("email", "")
                config_data["username"] = test_request2.get("username", "")
                config_data["access_token"] = encrypt_data("access_token", inputted)
                private_key = test_request.get("nordlynx_private_key", "")
                if private_key: config_data["encrypted_private_key"] = encrypt_data("private_key", private_key)
                successMessageBox(f"Successfully logged into {test_request2.get('email')}!", "NordWireConnect")
                save_configuration()
                update_tray()
            else: errorMessage("The provided NordVPN Access Token is invalid.")
        else: errorMessage("The provided NordVPN Access Token is invalid.")
    else: errorMessage("The provided NordVPN Access Token is invalid.")
def change_connection_channel():
    cur_ch = config_data.get("connection_channel")
    inputted = simpledialog.askstring(title="Connection Channel Input", prompt="Enter the Connection Channel you would like to assign this computer (1-10) (enter \"none\" for no channels):", initialvalue="none" if cur_ch is None else str(cur_ch))
    if inputted is None: return
    try:
        if inputted.strip().lower() == "none":
            config_data["connection_channel"] = None
            successMessageBox(f"Successfully cleared Connection Channel!", "NordWireConnect")
            save_configuration()
            update_tray()
            return
        val = int(inputted.strip())
        if 1 <= val <= 10:
            config_data["connection_channel"] = val
            successMessageBox(f"Successfully changed Connection Channel to {val}!", "NordWireConnect")
            save_configuration()
            update_tray()
        else: errorMessage("The given channel number is not in range from 1 to 10.")
    except Exception: errorMessage("The given channel number is invalid.")
def change_keepalive():
    inputted = simpledialog.askinteger(title="Persistent Keepalive Input", prompt="Enter the amount of keepalive (seconds) you would like to set (1-999):", initialvalue=config_data.get("persistent_keepalive", 25), minvalue=1, maxvalue=999)
    if inputted is None: return
    try:
        if 1 <= inputted <= 999:
            config_data["persistent_keepalive"] = inputted
            successMessageBox(f"Successfully changed Persistent Keepalive to {inputted}s!", "NordWireConnect")
            save_configuration()
            if session_data["connected"] == True: differ_from_status_action2()
            update_tray()
        else: errorMessage("The given keepalive value is not in range from 1 to 999.")
    except Exception: errorMessage("The given keepalive value is invalid.")
def change_mtu():
    cur_mtu = config_data.get("mtu", None)
    inputted = simpledialog.askstring(title="WireGuard MTU Input", prompt="Enter WireGuard interface MTU (1280-1500, or \"auto\" for dynamic NordLynx MTU):", initialvalue="auto" if not cur_mtu else str(cur_mtu))
    if inputted is None: return
    txt = inputted.strip().lower()
    if txt == "auto":
        config_data["mtu"] = None
        save_configuration()
        successMessageBox(f"Successfully set WireGuard MTU to Auto ({resolve_nordlynx_mtu()})!", "NordWireConnect")
        if session_data["connected"] == True: differ_from_status_action2()
        update_tray()
        return
    try:
        mtu_val = int(txt)
        if 1280 <= mtu_val <= 1500:
            config_data["mtu"] = mtu_val
            save_configuration()
            successMessageBox(f"Successfully changed WireGuard MTU to {mtu_val}!", "NordWireConnect")
            if session_data["connected"] == True: differ_from_status_action2()
            update_tray()
        else: errorMessage("MTU must be between 1280 and 1500 (or \"auto\").")
    except Exception: errorMessage("Invalid MTU value. Enter 1280-1500 or \"auto\".")
def change_allowlisted_subnets():
    cur = ", ".join(config_data.get("allowlisted_subnets") or [])
    inputted = simpledialog.askstring(title="Allowlisted Subnets Input", prompt="Enter comma-separated CIDRs/IPs to bypass VPN (e.g. 192.168.50.0/24, 1.2.3.4):", initialvalue=cur)
    if inputted is None: return
    valid = []
    invalid = []
    for raw in inputted.split(","):
        tok = raw.strip()
        if not tok: continue
        try:
            net = ipaddress.ip_network(tok, strict=False)
            if net.prefixlen < 8: invalid.append(f"{tok} (prefix < /8)")
            elif str(net) not in valid: valid.append(str(net))
        except Exception: invalid.append(tok)
    if invalid:
        errorMessage(f"Ignored invalid or overly wide subnet(s): {', '.join(invalid)}")
        if not valid and inputted.strip(): return
    config_data["allowlisted_subnets"] = valid
    save_configuration()
    successMessageBox(f"Allowlisted subnets set to: {', '.join(valid) if valid else 'None'}", "NordWireConnect")
    if session_data["connected"] == True: differ_from_status_action2()
    update_tray()
def change_load_swap_threshold():
    inputted = simpledialog.askinteger(title="Load Swapping Threshold Input", prompt="Enter the load percentage threshold for swapping servers (0-100):", initialvalue=config_data.get("load_swapping_threshold", 50), minvalue=0, maxvalue=100)
    if inputted is None: return
    try:
        if 0 <= inputted <= 100:
            config_data["load_swapping_threshold"] = inputted
            successMessageBox(f"Successfully changed Load Swapping Threshold to {inputted}%!", "NordWireConnect")
            save_configuration()
            update_tray()
        else: errorMessage("The given load percentage is not in range from 0 to 100.")
    except Exception: errorMessage("The given load percentage is invalid.")
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
    if usage.get("month") != cur_date.month or usage.get("year") != cur_date.year: usage = {"month": cur_date.month, "year": cur_date.year, "upload": 0, "download": 0}
    usage["upload"] += upload
    usage["download"] += download
    config_data["monthly_data_usage"] = usage
    total_usage["upload"] += upload
    total_usage["download"] += download
    config_data["total_data_usage"] = total_usage
    if data_reminded == False: refresh_data_usage()
    add_data_count += 1
    if add_data_count % 10 == 0: save_configuration(quiet=True)
def default_location(server_name: str, exact_mode: str=None): 
    if exact_mode: config_data["default_location"] = f"{exact_mode}_{server_name}"
    else: config_data["default_location"] = server_name
    save_configuration()
def mark_not_connected():
    session_data["session_time"] = 0
    session_data["connection_text"] = "Not Connected"
    session_data["connected"] = False
    session_data["server"] = None
    session_data["current_load"] = None
    session_data["server_health"] = "Good"
    session_data["download_data_usage"] = "0.00 B"
    session_data["upload_data_usage"] = "0.00 B"
    session_data_modified()
    update_tray()
def load_configuration():
    global config_data
    config_path = os.path.join(app_data_path, "ConnectConfig.json")
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f: temp_conf = json.load(f)
        for k, v in temp_conf.items():
            allowed_type = config_data_type_allowed.get(k)
            if allowed_type is not None and isinstance(v, allowed_type): config_data[k] = v
        if "load_swap_threshold" in temp_conf and isinstance(temp_conf["load_swap_threshold"], int): config_data["load_swapping_threshold"] = temp_conf["load_swap_threshold"]
        if config_data.get("data_encrypted") == False:
            for k in ["access_token", "openvpn_username", "openvpn_password"]:
                if type(config_data.get(k)) is str and config_data.get(k): config_data[k] = encrypt_data(k, config_data[k])
            config_data["data_encrypted"] = True
            save_configuration()
def save_configuration(quiet: bool=False):
    if not quiet: mainMessage("Saving configuration..")
    write_json_atomic(os.path.join(app_data_path, "ConnectConfig.json"), config_data, indent=4)

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
    save_configuration()
    addToThread(run_cache, (True,))
    update_tray()
def check_server_list(): return config_data.get("optimize_server_list", False) == True
def change_asn_server_list():
    config_data["asn_server_list"] = not config_data.get("asn_server_list", False)
    save_configuration()
    if config_data["asn_server_list"]: addToThread(premap_server_asns, (None, False, True))
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
def change_debug_mode():
    global debug_mode
    enabled = not debug_mode
    config_data["debug_mode"] = enabled
    debug_mode = enabled
    save_configuration()
    if enabled:
        system_class.enable_debug_console(f"NordWireConnect v{version} - Debugging", on_streams_changed=_rebind_terminal_streams)
        warnMessage("Debug Mode enabled (disable from tray menu to close console).")
    else:
        warnMessage("Debug Mode disabled.")
        system_class.disable_debug_console(on_streams_changed=_rebind_terminal_streams)
    update_tray()
def check_debug_mode(): return bool(debug_mode)
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
def change_routing_leak_protection():
    config_data["routing_leak_protection"] = not config_data.get("routing_leak_protection", True)
    save_configuration()
    if session_data["connected"] == True: differ_from_status_action2()
    update_tray()
def check_routing_leak_protection(): return config_data.get("routing_leak_protection", True) == True
def change_instant_reject_ipv6():
    config_data["instant_reject_ipv6"] = not config_data.get("instant_reject_ipv6", True)
    save_configuration()
    if session_data["connected"] == True: addToThread(send_command, ("set-ipv6-instant-reject", {"enabled": check_instant_reject_ipv6(), "split_lan": check_split_lan_routing(), "reject_ipv6_cidrs": compute_instant_reject_ipv6_cidrs()}, 10))
    update_tray()
def check_instant_reject_ipv6(): return config_data.get("instant_reject_ipv6", True) == True
def change_server_health_monitor():
    config_data["server_health_monitor"] = not config_data.get("server_health_monitor", True)
    save_configuration()
    update_tray()
def check_server_health_monitor(): return config_data.get("server_health_monitor", True) == True
def change_server_group(group_key: str):
    if group_key not in nord_server_groups: return
    config_data["server_group"] = group_key
    save_configuration()
    if session_data["connected"] == True: differ_from_status_action2()
    update_tray()
def check_server_group(group_key: str): return config_data.get("server_group", "all") == group_key

# Differenting Status
def differ_from_status_text() -> str:
    if is_vpn_paused(): return "Resume VPN Now"
    elif session_data["connected"] == True: return "Disconnect from Server"
    elif session_data["connection_text"].startswith("Connecting"): return "Cancel Connection"
    else: return "Connect to Server"
def differ_from_status_action():
    if is_vpn_paused():
        resume_paused_vpn()
        return
    if session_data["connection_text"].startswith("Connecting"):
        session_data["session_time"] = 0
        session_data["connection_text"] = "Not Connected"
        session_data["connected"] = False
        session_data["server"] = None
        reconnect_queued.clear()
        set_icon("app_icon.ico")
    elif session_data["connected"] == True: 
        s = disconnect()
        if s.get("suc") == True: successMessageBox("Successfully disconnected from NordVPN server!", "NordWireConnect")
    else: 
        default_loc = config_data.get("default_location", "auto") or "auto"
        if default_loc == "auto" or "_" not in default_loc:
            config_data["server"] = "auto"
            addToThread(connect, (), connection=True)
        else:
            mode, server_loc = default_loc.split("_", 1)
            config_data["server"] = server_loc
            addToThread(connect, (mode,), connection=True)
def differ_from_status_action2():
    if reconnect_queued.is_set() or is_vpn_paused() or session_data["connection_text"].startswith("Connecting") or session_data["connection_text"].startswith("Not"): return
    reconnect_queued.set()
    def pre_connect():
        try: connect(exact_mode=session_data.get("exact_connect"))
        finally: reconnect_queued.clear()
    addToThread(pre_connect, (), connection=True)
def differ_to_status(): return session_data["connection_text"]
def differ_to_city() -> str:
    if session_data["connected"] == True and session_data.get("server"): return f"City: {get_server_location(session_data['server'])[1] or 'N/A'}"
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
def differ_to_health() -> str:
    if session_data["connected"] == True: return f"Health: {session_data.get('server_health', 'Good')}"
    else: return "Health: N/A"
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
def differ_to_keepalive() -> str: return f"Persistent Keepalive: {config_data.get('persistent_keepalive', 25)}s"
def differ_to_mtu() -> str:
    if isinstance(config_data.get("mtu"), int): return f"WireGuard MTU: {config_data['mtu']}"
    return f"WireGuard MTU: Auto ({active_tunnel_mtu or 1420})"
def differ_to_allowlist() -> str:
    subs = config_data.get("allowlisted_subnets") or []
    return f"Allowlisted Subnets: {len(subs)} rule(s)" if subs else "Allowlisted Subnets: None"
def differ_to_server_group() -> str: return f"Server Group: {nord_server_groups.get(config_data.get('server_group', 'all'), ('All Servers', None))[0]}"
def differ_to_def_loc() -> str: 
    default_loc = config_data.get("default_location")
    if default_loc and default_loc.lower() != "auto" and "_" in default_loc:
        _, server_loc = default_loc.split("_", 1)
        return f"Default Location: {server_loc}"
    else: return "Default Location: Auto"
def differ_to_load_swap_usage() -> str: return f"Load Swapping Threshold: {config_data.get('load_swapping_threshold', 50)}%"
def differ_to_renewal_active() -> str:
    if config_data.get("plan_renewal_data", {}).get("active"): return "Plan Active: True"
    else: return "Plan Active: False"
def differ_to_renewal_date() -> str:
    if config_data.get("plan_renewal_data", {}).get("active") and config_data["plan_renewal_data"].get("expiry"):
        expiry_date = convert_nordtime_to_datetime(config_data["plan_renewal_data"]["expiry"])
        if not expiry_date: return "Plan Renewal Date: N/A"
        return f"Plan Renewal Date: {expiry_date.astimezone().strftime('%a %B %d, %Y at %H:%M:%S')}"
    else: return "Plan Renewal Date: N/A"
def differ_to_renewal_remaining() -> str:
    if config_data.get("plan_renewal_data", {}).get("active") and config_data["plan_renewal_data"].get("expiry"):
        current_date = get_current_time()
        expiry_date = convert_nordtime_to_datetime(config_data["plan_renewal_data"]["expiry"])
        if not expiry_date: return "Plan Renewal Remaining: N/A"
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
@functools.lru_cache(maxsize=8)
def compute_allowed_ips(split_lan: bool, leak_protection: bool=True, custom_subnets: tuple=()) -> str:
    if not split_lan and not custom_subnets:
        if leak_protection: return "0.0.0.0/0, ::/0"
        return "0.0.0.0/1, 128.0.0.0/1"
    protected_v4 = [ipaddress.ip_network("10.5.0.0/16"), ipaddress.ip_network("103.86.96.0/22"), ipaddress.ip_network("103.86.99.0/24")]
    v4_exclusions = []
    v6_exclusions = []
    if split_lan:
        v4_exclusions.extend([
            "10.0.0.0/14", "10.4.0.0/16", "10.6.0.0/15", "10.8.0.0/13",
            "10.16.0.0/12", "10.32.0.0/11", "10.64.0.0/10", "10.128.0.0/9",
            "172.16.0.0/12", "192.168.0.0/16", "169.254.0.0/16",
            "224.0.0.0/24", "239.255.255.250/32", "224.0.0.0/4", "255.255.255.255/32"
        ])
    for sub in custom_subnets:
        try: net = ipaddress.ip_network(str(sub).strip(), strict=False)
        except Exception: continue
        if net.prefixlen < 8: continue
        if net.version == 4:
            cur = [net]
            for p in protected_v4:
                nxt = []
                for c in cur:
                    if c.overlaps(p):
                        inter = max(c.network_address, p.network_address)
                        mask = max(c.prefixlen, p.prefixlen)
                        nxt.extend(list(c.address_exclude(ipaddress.ip_network(f"{inter}/{mask}"))))
                    else: nxt.append(c)
                cur = nxt
            v4_exclusions.extend(str(c) for c in cur)
        else: v6_exclusions.append(str(net))
    if v4_exclusions:
        v4_allowed = calculate_allowed_ips(["0.0.0.0/0"], v4_exclusions)
        v4_nets = [ipaddress.ip_network(n) for n in v4_allowed] + protected_v4
        v4_allowed = [str(n) for n in ipaddress.collapse_addresses(v4_nets)]
    else: v4_allowed = ["0.0.0.0/1", "128.0.0.0/1"]
    if leak_protection:
        v6_ex = ["::/16", "fc00::/7", "fe80::/10", "ff00::/8"] + v6_exclusions
        v6_allowed = calculate_allowed_ips(["::/0"], v6_ex)
        allowed = v4_allowed + v6_allowed
    else: allowed = v4_allowed
    return ", ".join(allowed)
@functools.lru_cache(maxsize=8)
def _build_reject_ipv6_cidrs(split_lan: bool, custom_subnets: tuple=()) -> str:
    v6_ex = ["::/16", "fe80::/10", "ff00::/8"] + (["fc00::/7"] if split_lan else [])
    for sub in custom_subnets:
        try:
            net = ipaddress.ip_network(str(sub).strip(), strict=False)
            if net.version == 6 and net.prefixlen >= 8: v6_ex.append(str(net))
        except Exception: pass
    return ",".join(calculate_allowed_ips(["::/0"], v6_ex))
def compute_instant_reject_ipv6_cidrs() -> str:
    subs = tuple(str(x).strip() for x in (config_data.get("allowlisted_subnets") or []) if str(x).strip())
    return _build_reject_ipv6_cidrs(bool(config_data.get("split_lan_routing", False)), subs)
def get_allowed_ips() -> str:
    subs = tuple(str(x).strip() for x in (config_data.get("allowlisted_subnets") or []) if str(x).strip())
    return compute_allowed_ips(bool(config_data.get("split_lan_routing", False)), bool(config_data.get("routing_leak_protection", True)), subs)
def get_physical_ipv4_set() -> frozenset:
    try:
        ips = socket.gethostbyname_ex(socket.gethostname())[2]
        return frozenset(ip for ip in ips if ip and not ip.startswith("10.5.") and not ip.startswith("127.") and not ip.startswith("169.254."))
    except Exception: return frozenset()
def haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    r_lat1, r_lon1, r_lat2, r_lon2 = math.radians(lat1), math.radians(lon1), math.radians(lat2), math.radians(lon2)
    d_lat = r_lat2 - r_lat1
    d_lon = r_lon2 - r_lon1
    a = (math.sin(d_lat / 2.0) ** 2) + math.cos(r_lat1) * math.cos(r_lat2) * (math.sin(d_lon / 2.0) ** 2)
    return 12742.0 * math.asin(min(1.0, math.sqrt(a)))
def refresh_user_insights(force: bool=False) -> dict:
    global cached_user_insights
    insights_path = os.path.join(app_data_path, "NordInsightsCache.json")
    if not cached_user_insights["loaded"] and os.path.exists(insights_path):
        try:
            with open(insights_path, "r", encoding="utf-8") as f: data = json.load(f)
            if isinstance(data, dict) and isinstance(data.get("latitude"), (int, float)) and isinstance(data.get("longitude"), (int, float)):
                cached_user_insights.update({"loaded": True, "latitude": float(data["latitude"]), "longitude": float(data["longitude"]), "country_code": str(data.get("country_code") or ""), "updated_at": float(data.get("updated_at") or 0.0)})
        except Exception: pass
    now = time.time()
    if session_data.get("connected") or send_command("connection-status", timeout=3).get("active") == True: return cached_user_insights
    if not force and cached_user_insights["loaded"] and (now - cached_user_insights["updated_at"]) < 1800: return cached_user_insights
    try:
        resp = requests.get("https://api.nordvpn.com/v1/helpers/ips/insights", timeout=6)
        if resp.ok and isinstance(resp.json, dict) and resp.json.get("protected") != True:
            lat = resp.json.get("latitude")
            lon = resp.json.get("longitude")
            if isinstance(lat, (int, float)) and isinstance(lon, (int, float)):
                cached_user_insights = {"loaded": True, "latitude": float(lat), "longitude": float(lon), "country_code": str(resp.json.get("country_code") or ""), "updated_at": now}
                write_json_atomic(insights_path, cached_user_insights)
    except Exception: pass
    return cached_user_insights
def get_coordinates_query() -> str:
    lat = cached_user_insights.get("latitude")
    lon = cached_user_insights.get("longitude")
    if isinstance(lat, (int, float)) and isinstance(lon, (int, float)): return f"&coordinates[longitude]={lon:.6f}&coordinates[latitude]={lat:.6f}"
    return ""
def _extract_server_geo(server: dict) -> tuple:
    locs = server.get("locations") if isinstance(server, dict) else None
    loc = locs[0] if isinstance(locs, list) and locs and isinstance(locs[0], dict) else {}
    country = loc.get("country") if isinstance(loc.get("country"), dict) else {}
    city = country.get("city") if isinstance(country.get("city"), dict) else {}
    lat = city.get("latitude")
    lon = city.get("longitude")
    cc = str(country.get("code") or "")
    hub = city.get("hub_score")
    return lat, lon, cc, hub
def rank_servers_by_penalty(servers: list) -> list:
    if not servers or len(servers) <= 1: return list(servers) if servers else []
    u_lat = cached_user_insights.get("latitude")
    u_lon = cached_user_insights.get("longitude")
    u_cc = str(cached_user_insights.get("country_code") or "").upper()
    has_coords = isinstance(u_lat, (int, float)) and isinstance(u_lon, (int, float))
    parsed = []
    distances = []
    for s in servers:
        if not isinstance(s, dict): continue
        s_lat, s_lon, s_cc, s_hub = _extract_server_geo(s)
        d = haversine_km(u_lat, u_lon, float(s_lat), float(s_lon)) if (has_coords and isinstance(s_lat, (int, float)) and isinstance(s_lon, (int, float))) else None
        if d is not None: distances.append(d)
        parsed.append((s, d, s_cc, s_hub))
    d_min = min(distances) if distances else 0.0
    d_max = max(distances) if distances else 0.0
    d_span = d_max - d_min
    scored = []
    for s, d, s_cc, s_hub in parsed:
        dist_p = 0.7 * (((d - d_min) / d_span) ** 0.5) if (d is not None and d_span > 1e-9) else 0.0
        f_load = max(0.0, float(s.get("load", 50) or 0)) / 10.0
        load_p = (f_load ** f_load) if f_load > 0 else 0.0
        country_p = 0.0 if (u_cc and s_cc and u_cc == s_cc.upper()) else 1.0
        hub_p = float(s_hub) if isinstance(s_hub, (int, float)) else 0.0
        total_p = dist_p + random.uniform(0.0, 0.05) - (country_p * hub_p) + load_p
        scored.append((total_p, s))
    scored.sort(key=lambda x: x[0])
    return [x[1] for x in scored]
def prune_excluded_hosts():
    now = time.monotonic()
    for h in [k for k, exp in excluded_maintenance_hosts.items() if exp <= now]: del excluded_maintenance_hosts[h]
def exclude_maintenance_host(hostname: str, ttl: float=900.0):
    if hostname: excluded_maintenance_hosts[hostname] = time.monotonic() + ttl
def verify_account_ens_state() -> tuple[str, str]:
    global private_key
    raw_token = config_data.get("access_token", "")
    if not raw_token: return "ok", ""
    try: decrypted_token = decode_data("access_token", raw_token)
    except Exception: return "ok", ""
    try:
        creds = requests.get("https://api.nordvpn.com/v1/users/services/credentials", auth=["token", decrypted_token], timeout=7)
        if creds.status_code in (401, 403): return "unauthenticated", "Your NordVPN Access Token has expired or been revoked. Please generate a new token in your NordAccount dashboard."
        if creds.ok and isinstance(creds.json, dict):
            new_pk = creds.json.get("nordlynx_private_key", "")
            if new_pk and new_pk != private_key:
                private_key = new_pk
                config_data["encrypted_private_key"] = encrypt_data("private_key", new_pk)
                save_configuration()
                return "key_rotated", new_pk
        services = requests.get("https://api.nordvpn.com/v1/users/services", auth=["token", decrypted_token], timeout=7)
        if services.status_code in (401, 403): return "unauthenticated", "Your NordVPN Access Token has expired or been revoked. Please generate a new token in your NordAccount dashboard."
        if services.ok and isinstance(services.json, list):
            has_vpn = False
            latest_raw = "0000-00-00 00:00:00"
            latest_dt = None
            now_utc = get_current_time()
            for item in services.json:
                if isinstance(item, dict) and isinstance(item.get("service"), dict) and item["service"].get("identifier") == "vpn":
                    has_vpn = True
                    raw_exp = item.get("expires_at") or "0000-00-00 00:00:00"
                    dt = convert_nordtime_to_datetime(raw_exp)
                    if dt and (latest_dt is None or dt > latest_dt): latest_dt, latest_raw = dt, raw_exp
            if has_vpn and latest_dt is not None:
                active = latest_dt > now_utc
                config_data["plan_renewal_data"] = {"expiry": latest_raw, "active": active}
                save_configuration()
                if not active: return "expired", "Your NordVPN subscription has expired. Please renew your plan to connect."
    except Exception: pass
    return "ok", ""
def prompt_session_limit_guide():
    if tk_root:
        def ask_open():
            if messagebox.askyesno("Session Limit Reached", f"NordVPN rejected the handshake across multiple online servers.\n\nWait a moment before retrying (retrying immediately can extend NordVPN's cooldown), or close active sessions on other devices (max 10 simultaneous connections).\n\nOpen NordVPN's Session Limit help guide?", icon="warning"): webbrowser.open("https://support.nordvpn.com/hc/en-us/articles/47181405478417-I-get-the-Session-Limit-Reached-error-on-NordVPN")
        ui_queue.put(ask_open)
def evaluate_server_health(srv: dict, handshake_age: typing.Optional[float]=None) -> str:
    if not isinstance(srv, dict) or srv.get("status") != "online" or get_wireguard_public_key(srv) is None: return "Maintenance"
    if handshake_age is not None and handshake_age > 180.0: return "Stalled Handshake"
    load = int(srv.get("load", 0) or 0)
    if load >= 80: return "Overloaded"
    if load >= 50: return "High Load"
    if load <= 25: return "Excellent"
    return "Good"
def is_vpn_paused() -> bool: return pause_until_monotonic > time.monotonic()
def pause_vpn(seconds: int, label: str=""):
    global pause_until_monotonic
    global paused_target
    if not session_data.get("connected"): return
    paused_target = {"server": session_data.get("org_server_id") or config_data.get("server", "auto"), "exact_mode": session_data.get("exact_connect")}
    pause_until_monotonic = time.monotonic() + float(seconds)
    disconnect(only_disconnect=True)
    session_data["session_time"] = 0
    session_data["connected"] = False
    session_data["server"] = None
    session_data["current_load"] = None
    session_data["server_health"] = "Good"
    session_data["download_data_usage"] = "0.00 B"
    session_data["upload_data_usage"] = "0.00 B"
    session_data["connection_text"] = f"Paused ({format_seconds(seconds)} left)"
    session_data_modified()
    set_icon("warning.ico")
    update_tray()
    successMessageBox(f"Paused VPN for {label or format_seconds(seconds)}. Connection will resume automatically.", "VPN Paused")
def resume_paused_vpn():
    global pause_until_monotonic
    global paused_target
    pause_until_monotonic = 0.0
    target = paused_target or {"server": config_data.get("server", "auto"), "exact_mode": None}
    paused_target = None
    config_data["server"] = target.get("server") or "auto"
    session_data["connection_text"] = "Connecting..."
    update_tray()
    addToThread(connect, (target.get("exact_mode"),), connection=True)
def cancel_pause_vpn():
    global pause_until_monotonic
    global paused_target
    pause_until_monotonic = 0.0
    paused_target = None
    mark_not_connected()
    set_icon("app_icon.ico")
    successMessageBox("VPN pause canceled.", "NordWireConnect")
@functools.lru_cache(maxsize=256)
def convert_to_channel(hostname: str) -> int: return ((zlib.crc32(hostname.encode("utf-8")) & 0xffffffff) % 10) + 1
def connect_server(server_name: str, exact_mode: str=None):
    config_data["server"] = server_name
    try: connect(exact_mode=exact_mode)
    except Exception: pass
def connect_session():
    global session_data
    global connect_monotonic_time
    global last_tunnel_rx
    global last_tunnel_tx
    global last_physical_ips
    global active_tunnel_mtu
    if session_data.get("connected") == True and config_data.get("auto_connect", True) == True and session_data.get("server"):
        stats = send_command("tunnel-stats", timeout=5)
        if stats.get("active") == True and tunnel_check(timeout=2):
            mainMessage(f"Adopted active WireGuard session on {session_data['server']['name']}!")
            connect_monotonic_time = time.monotonic() - float(session_data.get("session_time", 0) or 0)
            last_tunnel_rx = int(stats.get("rx_bytes", 0))
            last_tunnel_tx = int(stats.get("tx_bytes", 0))
            last_physical_ips = get_physical_ipv4_set()
            session_data["server_health"] = evaluate_server_health(session_data["server"], stats.get("handshake_age"))
            session_data["download_data_usage"] = format_wireguard_bytes(last_tunnel_rx)
            session_data["upload_data_usage"] = format_wireguard_bytes(last_tunnel_tx)
            send_command("set-nrpt-dns-lock", {"enabled": check_routing_leak_protection(), "dns_servers": get_active_conf_dns()}, timeout=8)
            send_command("set-ipv6-instant-reject", {"enabled": check_instant_reject_ipv6(), "split_lan": check_split_lan_routing(), "reject_ipv6_cidrs": compute_instant_reject_ipv6_cidrs()}, timeout=8)
            if config_data.get("mtu") is None: clamp_tunnel_path_mtu(get_active_conf_mtu())
            elif isinstance(config_data.get("mtu"), int): active_tunnel_mtu = config_data["mtu"]
            set_icon("connected.ico")
            update_tray()
            return
        mainMessage("Restoring previous session...")
        org = session_data.get("org_server_id", "auto")
        org_exact_connect = session_data.get("exact_connect")
        prev_server_short = session_data["server"]["name"].replace(" ", "").replace("#", "")
        connect_server(prev_server_short, exact_mode="server")
        config_data["server"] = org
        session_data["exact_connect"] = org_exact_connect
        if not session_data.get("connected") and org:
            mainMessage("Previous node unavailable; falling back to saved location preference..")
            config_data["server"] = org
            connect(exact_mode=org_exact_connect)
    elif session_data.get("connection_text", "").startswith("Connecting"):
        mainMessage("Reconnecting..")
        session_data["connected"] = False
        session_data["connection_text"] = "Not Connected"
        config_data["server"] = None
        session_data["server"] = None
        session_data["current_load"] = None
        session_data["exact_connect"] = None
        differ_from_status_action()
    elif send_command("connection-status", timeout=5).get("active") == True: disconnect()
def fetch_server_recommendations(url: str, timeout: float=10.0):
    resp = requests.get(url, timeout=timeout)
    if resp.ok and isinstance(resp.json, list): return resp
    if send_command("connection-status", timeout=5).get("active") == True:
        warnMessage("API unreachable over current tunnel; disconnecting stale tunnel and retrying..")
        disconnect(only_disconnect=True)
        deadline = time.monotonic() + 6.0
        while time.monotonic() < deadline and not tunnel_check(timeout=1.5): time.sleep(0.3)
        resp = requests.get(url, timeout=timeout)
    return resp
def load_cached_servers() -> list:
    cache_path = os.path.join(app_data_path, "NordServerCache.json")
    if os.path.exists(cache_path):
        try:
            mt = os.path.getmtime(cache_path)
            if server_cache_mem["servers"] and server_cache_mem["mtime"] == mt: return server_cache_mem["servers"]
            with open(cache_path, "r", encoding="utf-8") as f: data = json.load(f)
            if isinstance(data, list):
                server_cache_mem["mtime"] = mt
                server_cache_mem["servers"] = data
                return data
        except Exception: pass
    return []
def verify_new_tunnel(timeout: float=7.0) -> tuple[bool, dict]:
    deadline = time.monotonic() + timeout
    tunnel_check(timeout=1.5)
    while time.monotonic() < deadline:
        if session_data["connection_text"] == "Not Connected" and session_data["connected"] == False: return False, {}
        stats = send_command("tunnel-stats", timeout=5)
        if stats.get("has_handshake") == True and tunnel_check(timeout=2.0): return True, stats
        time.sleep(0.25)
        tunnel_check(timeout=1.5)
    stats = send_command("tunnel-stats", timeout=5)
    if stats.get("has_handshake") == True and tunnel_check(timeout=2.5): return True, stats
    return False, stats
def connect(exact_mode: str=None):
    global session_data
    global private_key
    global connect_monotonic_time
    global last_tunnel_rx
    global last_tunnel_tx
    global last_swapped_server_id
    global last_physical_ips
    global pause_until_monotonic
    global paused_target
    needing_state()
    try:
        pause_until_monotonic = 0.0
        paused_target = None
        session_data["session_time"] = 0
        session_data["connected"] = False
        session_data["server_health"] = "Good"
        session_data["connection_text"] = "Connecting..."
        session_data_modified()
        set_icon("awaiting.ico")

        if not private_key and config_data.get("encrypted_private_key"):
            try: private_key = decode_data("private_key", config_data["encrypted_private_key"])
            except Exception: private_key = None
        if not private_key and config_data.get("access_token"):
            try: fetch_and_cache_credentials(decode_data("access_token", config_data["access_token"]))
            except Exception: pass
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

        if not tunnel_check(timeout=2.5):
            disconnect(only_disconnect=True)
            time.sleep(0.5)
        if not tunnel_check(timeout=2.5):
            if config_data.get("auto_connect", True) == True:
                errorMessage("There was an issue trying to connect to your internet. Press Cancel Connection if you want to disconnect.")
                set_icon("warning.ico")
                while not tunnel_check(timeout=1.5):
                    if session_data["connection_text"] == "Not Connected" and session_data["connected"] == False: break
                    time.sleep(1)
                if session_data["connection_text"] == "Not Connected" and session_data["connected"] == False:
                    mainMessage("Canceled connection.")
                    return
                set_icon("awaiting.ico")
            else: 
                errorMessage("There was an issue trying to connect to your internet.")
                mark_not_connected()
                set_icon("error.ico")
                return

        cached_plan = config_data.get("plan_renewal_data") or {}
        if cached_plan.get("expiry"):
            cached_exp_dt = convert_nordtime_to_datetime(cached_plan["expiry"])
            if cached_exp_dt and cached_exp_dt <= get_current_time():
                ens_state, ens_msg = verify_account_ens_state()
                if ens_state == "unauthenticated":
                    errorMessage(ens_msg, "Access Token Revoked")
                    mark_not_connected()
                    set_icon("error.ico")
                    return
                elif ens_state == "expired":
                    errorMessage(ens_msg, "NordVPN Plan Expired")
                    mark_not_connected()
                    set_icon("error.ico")
                    return

        refresh_user_insights(force=False)
        prune_excluded_hosts()
        active_group_key = config_data.get("server_group", "all")
        active_group_id = nord_server_groups.get(active_group_key, (None, None))[1]
        group_filter = get_group_query_filter(active_group_key)
        coord_query = get_coordinates_query()
        should_filter = False
        from_cache = False
        asn_city_name = None
        sel_group_id = None
        group_country_norm = None
        all_server_recommendations = []
        if exact_mode == "group":
            raw_group_target = str(config_data.get("server") or "all")
            if "|||" in raw_group_target:
                group_key, group_country = raw_group_target.split("|||", 1)
                group_country_norm = group_country.replace(" ", "")
            else:
                group_key, group_country = raw_group_target, ""
            sel_group_id = nord_server_groups.get(group_key, (None, None))[1]
            c_filter = ""
            if group_country:
                country_list_path = os.path.join(app_data_path, "NordCountryCache.json")
                if not os.path.exists(country_list_path): run_cache(force=True)
                if os.path.exists(country_list_path):
                    with open(country_list_path, "r", encoding="utf-8") as f:
                        for c in json.load(f):
                            if c.get("name") == group_country: c_filter = f"&filters[country_id]={c.get('id')}"; break
            g_q = f"&filters[servers_groups]={sel_group_id}" if sel_group_id is not None else ""
            resp = fetch_server_recommendations(f"https://api.nordvpn.com/v1/servers/recommendations?{server_filter}{c_filter}{wireguard_tech_filter}{g_q}{coord_query}&limit=100")
            if resp.ok and isinstance(resp.json, list): all_server_recommendations = resp.json
            else:
                from_cache = True
                all_server_recommendations = rank_servers_by_penalty([s for s in load_cached_servers() if server_has_group_id(s, sel_group_id) and (not group_country or get_server_location(s)[0] == group_country)])[:100]
            if not all_server_recommendations:
                errorMessage("Unable to get NordVPN specialty server recommendations.")
                mark_not_connected()
                set_icon("error.ico")
                return
            should_filter = True
        elif config_data["server"] == "auto" or config_data["server"] == None or exact_mode == None:
            resp = fetch_server_recommendations(f"https://api.nordvpn.com/v1/servers/recommendations?{server_filter}{wireguard_tech_filter}{group_filter}{coord_query}&limit=100")
            if resp.ok and isinstance(resp.json, list): all_server_recommendations = resp.json
            else:
                from_cache = True
                all_server_recommendations = rank_servers_by_penalty([s for s in load_cached_servers() if server_has_group_id(s, active_group_id)])[:100]
            if not all_server_recommendations:
                errorMessage("Unable to get NordVPN server recommendations.")
                mark_not_connected()
                set_icon("error.ico")
                return
            should_filter = True
        elif exact_mode == "country":
            country_list_path = os.path.join(app_data_path, "NordCountryCache.json")
            if not os.path.exists(country_list_path): run_cache(force=True)
            if not os.path.exists(country_list_path):
                errorMessage("Unable to find NordVPN country list.")
                mark_not_connected()
                set_icon("error.ico")
                return
            with open(country_list_path, "r", encoding="utf-8") as f: country_list = json.load(f)
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
            resp = fetch_server_recommendations(f"https://api.nordvpn.com/v1/servers/recommendations?{server_filter}&filters[country_id]={country_id}{wireguard_tech_filter}{group_filter}{coord_query}&limit=100")
            if resp.ok and isinstance(resp.json, list): all_server_recommendations = resp.json
            else:
                from_cache = True
                all_server_recommendations = rank_servers_by_penalty([s for s in load_cached_servers() if get_server_location(s)[0] == config_data["server"] and server_has_group_id(s, active_group_id)])[:100]
            if not all_server_recommendations:
                errorMessage("Unable to get NordVPN server recommendations.")
                mark_not_connected()
                set_icon("error.ico")
                return
            should_filter = True
        elif exact_mode == "city":
            city_list_path = os.path.join(app_data_path, "NordCityCache.json")
            if not os.path.exists(city_list_path): run_cache(force=True)
            if not os.path.exists(city_list_path):
                errorMessage("Unable to find NordVPN city list.")
                mark_not_connected()
                set_icon("error.ico")
                return
            with open(city_list_path, "r", encoding="utf-8") as f: city_list = json.load(f)
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
            resp = fetch_server_recommendations(f"https://api.nordvpn.com/v1/servers/recommendations?{server_filter}&filters[country_city_id]={city_id}{wireguard_tech_filter}{group_filter}{coord_query}&limit=100")
            if resp.ok and isinstance(resp.json, list): all_server_recommendations = resp.json
            else:
                from_cache = True
                all_server_recommendations = rank_servers_by_penalty([s for s in load_cached_servers() if get_server_location(s)[1] == config_data["server"] and server_has_group_id(s, active_group_id)])[:100]
            if not all_server_recommendations:
                errorMessage("Unable to get NordVPN server recommendations.")
                mark_not_connected()
                set_icon("error.ico")
                return
            should_filter = True
        elif exact_mode == "asn":
            load_server_asn_map()
            try: city_name, target_asn = config_data["server"].rsplit("|||", 1)
            except Exception:
                errorMessage("Invalid ASN format given.")
                mark_not_connected()
                set_icon("error.ico")
                return
            asn_city_name = city_name
            city_list_path = os.path.join(app_data_path, "NordCityCache.json")
            if not os.path.exists(city_list_path): run_cache(force=True)
            if not os.path.exists(city_list_path):
                errorMessage("Unable to find NordVPN city list.")
                mark_not_connected()
                set_icon("error.ico")
                return
            with open(city_list_path, "r", encoding="utf-8") as f: city_list = json.load(f)
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
            req = fetch_server_recommendations(f"https://api.nordvpn.com/v1/servers/recommendations?{server_filter}&filters[country_city_id]={city_id}{wireguard_tech_filter}{group_filter}{coord_query}&limit=5000", timeout=15)
            city_servers = req.json if (req.ok and isinstance(req.json, list)) else [s for s in load_cached_servers() if get_server_location(s)[1] == city_name]
            all_server_recommendations = []
            for server in city_servers:
                if len(all_server_recommendations) >= 100: break
                if active_group_id is not None and not server_has_group_id(server, active_group_id): continue
                server_ip = server.get("station") 
                if server_ip:
                    asn_name = get_asn_name_from_ip(server_ip)
                    if asn_name == target_asn: all_server_recommendations.append(server)
            if cached_asn_data["loaded"]: unload_asn_index()
            if not all_server_recommendations:
                errorMessage(f"No available servers found for {target_asn} in {city_name}.")
                mark_not_connected()
                set_icon("error.ico")
                return
            all_server_recommendations = rank_servers_by_penalty(all_server_recommendations)
            should_filter = True
        elif exact_mode == "server":
            server_list_path = os.path.join(app_data_path, "NordServerIDCache.json")
            if not os.path.exists(server_list_path): run_cache(force=True)
            if not os.path.exists(server_list_path):
                errorMessage("Unable to find NordVPN server list.")
                mark_not_connected()
                set_icon("error.ico")
                return
            with open(server_list_path, "r", encoding="utf-8") as f: server_id_list = json.load(f)
            server_id = server_id_list.get(config_data["server"])
            if server_id:
                resp = fetch_server_recommendations(f"https://api.nordvpn.com/v1/servers?{server_filter}&filters[servers.id]={server_id}&limit=1")
                if resp.ok and isinstance(resp.json, list) and resp.json: all_server_recommendations = resp.json
                else: all_server_recommendations = [s for s in load_cached_servers() if s.get("id") == server_id]
                if not all_server_recommendations:
                    errorMessage("Unable to get NordVPN server recommendations.")
                    mark_not_connected()
                    set_icon("error.ico")
                    return
            else:
                errorMessage("Unable to find NordVPN server.")
                mark_not_connected()
                set_icon("error.ico")
                return
        else:
            resp = fetch_server_recommendations(f"https://api.nordvpn.com/v1/servers/recommendations?{server_filter}{wireguard_tech_filter}{group_filter}{coord_query}&limit=100")
            if resp.ok and isinstance(resp.json, list): all_server_recommendations = resp.json
            else:
                from_cache = True
                all_server_recommendations = rank_servers_by_penalty([s for s in load_cached_servers() if server_has_group_id(s, active_group_id)])[:100]
            if not all_server_recommendations:
                errorMessage("Unable to get NordVPN server recommendations.")
                mark_not_connected()
                set_icon("error.ico")
                return
            should_filter = True

        if should_filter and excluded_maintenance_hosts:
            non_excluded = [s for s in all_server_recommendations if s.get("hostname") not in excluded_maintenance_hosts]
            if non_excluded: all_server_recommendations = non_excluded
        if should_filter and config_data.get("connection_channel"):
            channeled_list = [s for s in all_server_recommendations if convert_to_channel(s["hostname"]) == config_data.get("connection_channel")]
            if channeled_list: all_server_recommendations = channeled_list
        if should_filter and len(all_server_recommendations) > 1:
            if check_load_swapping():
                threshold = config_data.get("load_swapping_threshold", 50)
                low_load = [s for s in all_server_recommendations if s.get("load", 100) < threshold]
                if low_load: all_server_recommendations = low_load
            if from_cache: all_server_recommendations = rank_servers_by_penalty(all_server_recommendations)
            if last_swapped_server_id is not None:
                others = [s for s in all_server_recommendations if s.get("id") != last_swapped_server_id]
                if others: all_server_recommendations = others + [s for s in all_server_recommendations if s.get("id") == last_swapped_server_id]
                last_swapped_server_id = None
        mainMessage(f"Loaded {len(all_server_recommendations)} NordVPN server recommendation(s).")

        # Start Finding Servers
        target_server_cfg = config_data.get("server")
        target_norm = target_server_cfg.replace(" ", "").replace("#", "") if isinstance(target_server_cfg, str) else ""
        asn_city_norm = asn_city_name.replace(" ", "") if isinstance(asn_city_name, str) else None
        mainMessage(f"Connecting to NordVPN Servers in {target_server_cfg}..")
        connected_server = None
        connected_stats = {}
        attempts = 0
        max_attempts = 8
        zero_hs_verified_servers = 0
        checked_account_ens = False
        allowed_ips_str = get_allowed_ips()
        nrpt_enabled = check_routing_leak_protection()
        reject_v6_enabled = check_instant_reject_ipv6()
        split_lan_enabled = check_split_lan_routing()
        reject_v6_cidrs = compute_instant_reject_ipv6_cidrs()
        dns = resolve_tunnel_dns()
        mtu_val = resolve_nordlynx_mtu()
        mtu_line = f"\nMTU = {mtu_val}"
        keepalive_val = config_data.get("persistent_keepalive", 25)
        for s in all_server_recommendations:
            if attempts >= max_attempts: break
            if not s.get("name"): continue
            shortened_name = s["name"].replace(" ", "").replace("#", "")
            raw_country, raw_city = get_server_location(s)
            country_name = raw_country.replace(" ", "")
            city_name = raw_city.replace(" ", "")
            if exact_mode == "group":
                if not server_has_group_id(s, sel_group_id) or (group_country_norm and country_name != group_country_norm): continue
            elif target_server_cfg not in ("auto", None):
                matched = shortened_name == target_norm or city_name == target_norm or country_name == target_norm or (asn_city_norm is not None and city_name == asn_city_norm)
                if not matched: continue
            if session_data["connection_text"] == "Not Connected" and session_data["connected"] == False: break
            if s.get("status") != "online" or s.get("load", 0) == 0:
                mainMessage(f"Server offline: {s['name']}")
                continue
            wg_pub_key = get_wireguard_public_key(s)
            if not wg_pub_key:
                mainMessage(f"WireGuard pivot offline on: {s['name']}")
                continue
            mainMessage(f"Attempting connection for: {s['name']}")
            configuration = f"# {shortened_name}.{city_name}.conf\n[Interface]\nPrivateKey = {private_key}\nAddress = 10.5.0.2/16\nDNS = {dns}{mtu_line}\n\n[Peer]\nPublicKey = {wg_pub_key}\nEndpoint = {s.get('station')}:51820\nAllowedIPs = {allowed_ips_str}\nPersistentKeepalive = {keepalive_val}\n"
            config_path = os.path.join(app_data_path, "NordWireConnect.conf")
            with open(config_path, "w", encoding="utf-8") as f: f.write(configuration)
            if session_data["connection_text"] == "Not Connected" and session_data["connected"] == False: break
            vpn_res = send_command("install-wire-tunnel", {"config_path": config_path, "nrpt_dns_lock": nrpt_enabled, "dns_servers": dns, "instant_reject_ipv6": reject_v6_enabled, "split_lan": split_lan_enabled, "reject_ipv6_cidrs": reject_v6_cidrs}, timeout=35)
            attempts += 1
            if vpn_res and vpn_res.get("suc") == True:
                mainMessage(f"Added WireGuard tunnel for server: {s['name']}; verifying handshake..")
                ok, stats = verify_new_tunnel(timeout=7.0)
                if ok:
                    connected_server = s
                    connected_stats = stats
                    break
                else:
                    warnMessage(f"Tunnel handshake/connectivity check failed on {s['name']}; running ENS diagnostic..")
                    disconnect(only_disconnect=True)
                    net_deadline = time.monotonic() + 3.0
                    while time.monotonic() < net_deadline and not tunnel_check(timeout=1.5): time.sleep(0.25)
                    if stats.get("has_handshake") == False:
                        if not tunnel_check(timeout=1.5):
                            warnMessage("Internet connectivity unavailable while diagnosing handshake failure.")
                            break
                        srv_chk = requests.get(f"https://api.nordvpn.com/v1/servers?{server_filter}&filters[servers.id]={s['id']}&limit=1", timeout=6)
                        if srv_chk.ok and isinstance(srv_chk.json, list):
                            matched_srv = next((item for item in srv_chk.json if item.get("id") == s.get("id")), None)
                            if matched_srv is None or matched_srv.get("status") != "online" or get_wireguard_public_key(matched_srv) is None:
                                exclude_maintenance_host(s.get("hostname", ""))
                                warnMessage(f"ENS: Server {s['name']} entered maintenance; going to next server..")
                                attempts = max(0, attempts - 1)
                                continue
                        if not checked_account_ens:
                            checked_account_ens = True
                            ens_state, ens_msg = verify_account_ens_state()
                            if ens_state == "unauthenticated":
                                errorMessage(ens_msg, "Access Token Revoked")
                                mark_not_connected()
                                set_icon("error.ico")
                                return
                            elif ens_state == "expired":
                                errorMessage(ens_msg, "NordVPN Plan Expired")
                                mark_not_connected()
                                set_icon("error.ico")
                                return
                            elif ens_state == "key_rotated":
                                warnMessage("ENS: NordLynx private key rotated on account; retrying server..")
                                configuration = f"# {shortened_name}.{city_name}.conf\n[Interface]\nPrivateKey = {private_key}\nAddress = 10.5.0.2/16\nDNS = {dns}{mtu_line}\n\n[Peer]\nPublicKey = {wg_pub_key}\nEndpoint = {s.get('station')}:51820\nAllowedIPs = {allowed_ips_str}\nPersistentKeepalive = {keepalive_val}\n"
                                with open(config_path, "w", encoding="utf-8") as f: f.write(configuration)
                                vpn_res2 = send_command("install-wire-tunnel", {"config_path": config_path, "nrpt_dns_lock": nrpt_enabled, "dns_servers": dns, "instant_reject_ipv6": reject_v6_enabled, "split_lan": split_lan_enabled, "reject_ipv6_cidrs": reject_v6_cidrs}, timeout=35)
                                if vpn_res2 and vpn_res2.get("suc") == True:
                                    ok2, stats2 = verify_new_tunnel(timeout=7.0)
                                    if ok2:
                                        connected_server = s
                                        connected_stats = stats2
                                        break
                                    disconnect(only_disconnect=True)
                        zero_hs_verified_servers += 1
                        if zero_hs_verified_servers >= 2:
                            warnMessage("ENS: Multiple online NordLynx servers returned no handshake; stopping retries to avoid extending session cooldown.")
                            errorMessage("Too many active sessions or connection attempts (NordVPN allows up to 10 simultaneous devices). Wait a moment before retrying.", "Session Limit Reached")
                            prompt_session_limit_guide()
                            mark_not_connected()
                            set_icon("error.ico")
                            return
                    continue
            else:
                warnMessage(f"Failed to install WireGuard tunnel for server: {s['name']}")
                continue

        if session_data["connection_text"] == "Not Connected" and session_data["connected"] == False:
            if connected_server: disconnect(only_disconnect=True)
            return
        elif connected_server:
            if config_data.get("mtu") is None: clamp_tunnel_path_mtu(mtu_val)
            connect_monotonic_time = time.monotonic()
            last_tunnel_rx = int(connected_stats.get("rx_bytes", 0))
            last_tunnel_tx = int(connected_stats.get("tx_bytes", 0))
            last_physical_ips = get_physical_ipv4_set()
            conn_country, conn_city = get_server_location(connected_server)
            successMessage(f"Successfully connected to server: {connected_server['name']}!")
            successMessage(f"Load: {connected_server['load']}%")
            successMessage(f"City: {conn_city or 'N/A'}")
            successMessage(f"IP Address: {connected_server['station']}")
            successMessage(f"Hostname: {connected_server['hostname']}")
            session_data["connected"] = True
            session_data["server"] = connected_server
            session_data["current_load"] = connected_server["load"]
            session_data["server_health"] = evaluate_server_health(connected_server, connected_stats.get("handshake_age"))
            session_data["connection_text"] = f"Connected to {connected_server['name']}"
            session_data["session_time"] = 0
            session_data["download_data_usage"] = format_wireguard_bytes(last_tunnel_rx)
            session_data["upload_data_usage"] = format_wireguard_bytes(last_tunnel_tx)
            session_data["exact_connect"] = exact_mode
            session_data["org_server_id"] = config_data["server"]
            session_data_modified()
            add_server_to_history({
                "city": conn_city or "N/A",
                "country": conn_country or "N/A",
                "time": get_current_time().timestamp(),
                "name": connected_server["name"],
            })
            successMessageBox(f"Successfully connected to server: {connected_server['name']}!", "Yay!")
            set_icon("connected.ico")
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
    finally: unneeding_state()
def disconnect(only_disconnect: bool=False) -> dict:
    global last_tunnel_rx
    global last_tunnel_tx
    global pause_until_monotonic
    global paused_target
    mainMessage("Disconnecting Wireguard..")
    awared = None
    if session_data.get("connected") == True and session_data.get("server"):
        server_name = session_data["server"]["name"].replace(" ", "").replace("#", "")
        awared = server_name
    last_tunnel_rx = 0
    last_tunnel_tx = 0
    if only_disconnect == False:
        pause_until_monotonic = 0.0
        paused_target = None
        reconnect_queued.clear()
        session_data["session_time"] = 0
        session_data["connection_text"] = "Not Connected"
        session_data["connected"] = False
        session_data["server"] = None
        session_data["current_load"] = None
        session_data["server_health"] = "Good"
        session_data["download_data_usage"] = "0.00 B"
        session_data["upload_data_usage"] = "0.00 B"
        session_data_modified()
        set_icon("app_icon.ico")
        update_tray()
    if awared: return send_command("combined-disconnect", {"known_server": awared}, timeout=30)
    else: return send_command("combined-disconnect", timeout=30)
def brute_end_wireguard():
    disconnect()
    mark_not_connected()
def update_ui_timer():
    global last_tray_signature
    try:
        if session_data.get("connected") == True and session_data.get("server"):
            if connect_monotonic_time > 0: session_data["session_time"] = max(0, int(time.monotonic() - connect_monotonic_time))
            srv_name = session_data["server"].get("name", "Unknown")
            city_name = get_server_location(session_data["server"])[1] or "N/A"
            if pystray_icon: pystray_icon.title = f"NordWireConnect v{version}\nConnected to {srv_name}!\nCity: {city_name}\nSession Time: {format_seconds(session_data['session_time'])}"
            sig = (True, srv_name, session_data.get("current_load"), session_data.get("server_health"), session_data.get("download_data_usage"), session_data.get("upload_data_usage"), session_data["session_time"] // 5)
            if sig != last_tray_signature:
                last_tray_signature = sig
                update_tray()
        elif pause_until_monotonic > 0:
            rem = int(round(pause_until_monotonic - time.monotonic()))
            if rem <= 0: resume_paused_vpn()
            else:
                rem_txt = format_seconds(rem)
                session_data["connection_text"] = f"Paused ({rem_txt} left)"
                if pystray_icon: pystray_icon.title = f"NordWireConnect v{version}\nVPN Paused ({rem_txt} left)"
                sig = ("paused", rem // 5)
                if sig != last_tray_signature:
                    last_tray_signature = sig
                    update_tray()
        else:
            if pystray_icon: pystray_icon.title = f"NordWireConnect v{version}"
            sig = (False, session_data.get("connection_text"))
            if sig != last_tray_signature:
                last_tray_signature = sig
                update_tray()
    except Exception: pass
    if not stop_app: tk_root.after(1000, update_ui_timer)
def connection_monitor_loop():
    global last_tunnel_rx
    global last_tunnel_tx
    global last_swapped_server_id
    global last_physical_ips
    tick = 0
    while not stop_app:
        time.sleep(5)
        if stop_app: break
        tick += 1
        try:
            if tick % 60 == 0:
                if config_data.get("plan_renewal_reminder", True) and config_data.get("plan_renewal_data", {}).get("active") and not plan_reminded: refresh_plan_info()
            elif not plan_reminded and config_data.get("plan_renewal_reminder", True) and config_data.get("plan_renewal_data", {}).get("active") and config_data.get("plan_renewal_data", {}).get("expiry"):
                if get_if_within_reminder_period(config_data["plan_renewal_data"]["expiry"]) != -1: refresh_plan_info()
            if tick % 360 == 0 and not session_data.get("connected") and not is_vpn_paused(): refresh_user_insights(force=False)
            if is_vpn_paused() or not session_data.get("connected") or reconnect_queued.is_set() or session_data.get("connection_text", "").startswith("Connecting"): continue
            stats = send_command("tunnel-stats", timeout=8)
            if not stats or stats.get("suc") != True: continue
            if stats.get("service_running"):
                rx_now = int(stats.get("rx_bytes", 0))
                tx_now = int(stats.get("tx_bytes", 0))
                delta_rx = (rx_now - last_tunnel_rx) if rx_now >= last_tunnel_rx else rx_now
                delta_tx = (tx_now - last_tunnel_tx) if tx_now >= last_tunnel_tx else tx_now
                last_tunnel_rx = rx_now
                last_tunnel_tx = tx_now
                if delta_rx > 0 or delta_tx > 0: add_data_usage_to_history(delta_rx, delta_tx)
                session_data["download_data_usage"] = format_wireguard_bytes(rx_now)
                session_data["upload_data_usage"] = format_wireguard_bytes(tx_now)
            in_connect_grace = (time.monotonic() - connect_monotonic_time) < 15.0
            since_resume = stats.get("seconds_since_resume")
            in_resume_grace = (since_resume is not None and since_resume < 12.0)
            cur_ips = get_physical_ipv4_set()
            if cur_ips and cur_ips != last_physical_ips:
                prev_ips = last_physical_ips
                last_physical_ips = cur_ips
                if prev_ips and not in_connect_grace:
                    mainMessage(f"Physical network change detected ({sorted(cur_ips)}); checking VPN tunnel..")
                    prev_mtu = active_tunnel_mtu
                    target_mtu = resolve_nordlynx_mtu()
                    conf_mtu = get_active_conf_mtu()
                    mtu_shrunk = config_data.get("mtu") is None and prev_mtu is not None and target_mtu < prev_mtu
                    if mtu_shrunk or not tunnel_check(timeout=1.8):
                        warnMessage("Rebinding VPN tunnel after network change..")
                        differ_from_status_action2()
                        continue
                    elif config_data.get("mtu") is None:
                        probe_ceil = min(target_mtu, conf_mtu)
                        if prev_mtu is not None and probe_ceil > prev_mtu: send_command("set-tunnel-mtu", {"mtu": probe_ceil}, timeout=8)
                        clamp_tunnel_path_mtu(probe_ceil)
            if since_resume is not None and 4.0 <= since_resume <= 14.0 and not tunnel_check(timeout=2.0):
                warnMessage("Tunnel unresponsive after system wake; reconnecting..")
                differ_from_status_action2()
                continue
            hs_age = stats.get("handshake_age")
            if hs_age is not None and hs_age > 175.0: tunnel_check(timeout=1.5)
            if session_data.get("server"): session_data["server_health"] = evaluate_server_health(session_data["server"], hs_age)
            if config_data.get("loss_connect_protection", True) and not in_connect_grace and not in_resume_grace:
                if stats.get("service_running") == False:
                    warnMessage("WireGuard tunnel service stopped unexpectedly; reconnecting..")
                    differ_from_status_action2()
                    continue
                if stats.get("active") == False:
                    tunnel_check(timeout=2.0)
                    time.sleep(2.0)
                    recheck = send_command("tunnel-stats", timeout=8)
                    if recheck.get("suc") == True and recheck.get("active") == False and not tunnel_check(timeout=3.0):
                        warnMessage("WireGuard handshake expired and tunnel unresponsive; reconnecting..")
                        differ_from_status_action2()
                        continue
                if tick % 6 == 0:
                    if not tunnel_check(timeout=2.5):
                        time.sleep(2.5)
                        if not tunnel_check(timeout=3.0) and not session_data["connection_text"].startswith("Connecting"):
                            warnMessage("End-to-end connectivity lost; reconnecting..")
                            differ_from_status_action2()
                            continue
            if tick % 12 == 0 and session_data.get("server"):
                server_id = session_data["server"]["id"]
                servers_resp = requests.get(f"https://api.nordvpn.com/v1/servers?{server_filter}&filters[servers.id]={server_id}&limit=1", timeout=8)
                if servers_resp.ok and isinstance(servers_resp.json, list):
                    matched_srv = next((srv for srv in servers_resp.json if srv.get("id") == server_id), None)
                    if matched_srv is None or matched_srv.get("status") != "online" or get_wireguard_public_key(matched_srv) is None:
                        session_data["server_health"] = "Maintenance"
                        if check_server_health_monitor():
                            cur_srv = session_data["server"]
                            exclude_maintenance_host(cur_srv.get("hostname", ""))
                            last_swapped_server_id = server_id
                            if session_data.get("exact_connect") == "server":
                                c_country, c_city = get_server_location(cur_srv)
                                if c_city and c_city != c_country:
                                    config_data["server"] = c_city
                                    session_data["exact_connect"] = "city"
                                elif c_country:
                                    config_data["server"] = c_country
                                    session_data["exact_connect"] = "country"
                                else:
                                    config_data["server"] = "auto"
                                    session_data["exact_connect"] = None
                            warnMessage(f"Server {cur_srv.get('name')} entered maintenance; migrating to a healthy server..")
                            notification("Server Maintenance", f"{cur_srv.get('name')} entered maintenance. Reconnecting to a healthy server..")
                            differ_from_status_action2()
                            continue
                    else:
                        session_data["current_load"] = matched_srv.get("load", session_data["current_load"])
                        session_data["server_health"] = evaluate_server_health(matched_srv, hs_age)
                threshold = config_data.get("load_swapping_threshold", 50)
                if check_load_swapping() and session_data.get("exact_connect") != "server" and isinstance(session_data.get("current_load"), (int, float)) and session_data["current_load"] >= threshold:
                    mainMessage(f"Server load ({session_data['current_load']}%) >= {threshold}%; swapping to lower-load server..")
                    last_swapped_server_id = server_id
                    differ_from_status_action2()
        except Exception: pass
def poll_ui_queue():
    while not ui_queue.empty():
        try:
            func = ui_queue.get_nowait()
            if func: func()
        except Exception as e: errorMessage(f"Error handling UI event: {trace(e)}")
    if not stop_app: tk_root.after(50, poll_ui_queue)
def tunnel_check(timeout: float=3.0) -> bool:
    for host, port in probe_endpoints:
        if requests.get_if_connected(server=host, port=port, timeout=timeout): return True
    return False
def wait_until_connected(max_wait: float=30.0) -> bool:
    deadline = time.monotonic() + max_wait
    while not tunnel_check(timeout=2.0):
        if stop_app or time.monotonic() > deadline: return False
        time.sleep(0.5)
    return True

# Main Runtime
def app():
    global one_instance_lock
    global config_data
    global pystray_icon
    global private_key
    global session_data
    global debug_mode
    global tk_root
    os.makedirs(app_data_path, exist_ok=True)
    debug_mode = read_early_debug_mode()
    if debug_mode: system_class.enable_debug_console(f"NordWireConnect v{version} - Debugging")
    setup_logging()
    systemMessage(f"{'-'*5:^5} NordWireConnect v{version} {'-'*5:^5}")

    # Setup Tkinter
    mainMessage("Starting Tkinter Library..")
    tk_root = tk.Tk()
    tk_root.withdraw()

    # Debug Message
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
            with open(os.path.join(app_data_path, "NordSessionData.json"), "r", encoding="utf-8") as f:
                loaded_sess = json.load(f)
                if isinstance(loaded_sess, dict): session_data.update(loaded_sess)
        legacy_status = os.path.join(app_data_path, "ConnectionStatus")
        if os.path.exists(legacy_status):
            try: os.remove(legacy_status)
            except Exception: pass
    except Exception as e: warnMessage(f"Configuration/session recovery warning: {trace(e)}")

    # Handle Starting with Service
    mainMessage("Verifying Service Status..")
    if not system_class.service_exists():
        mainMessage("NordWireConnectService is not running. Attempting to start it now..")
        system_class.start_service()
    was_old = config_data.get("last_version") != version
    if was_old == True:
        setup_shortcuts()
        config_data["last_version"] = version
        save_configuration()
    send_command("ui-opening", timeout=5)

    # Break Hanging Connections
    mainMessage("Checking for hanging tunnels..")
    try:
        if not tunnel_check(timeout=2.0) and send_command("connection-status", timeout=5).get("active") == True:
            mainMessage("Clearing unresponsive WireGuard tunnel from previous boot..")
            disconnect(only_disconnect=True)
            recover_deadline = time.monotonic() + 5.0
            while time.monotonic() < recover_deadline and not tunnel_check(timeout=1.0): time.sleep(0.25)
    except Exception as e: warnMessage(f"Warning while checking hanging connections: {trace(e)}")

    # Fetch NordVPN Credentials
    if config_data.get("encrypted_private_key"):
        try: private_key = decode_data("private_key", config_data["encrypted_private_key"])
        except Exception: private_key = None
    def background_startup_sync():
        refresh_user_insights(force=False)
        if config_data.get("access_token"):
            try:
                if wait_until_connected(max_wait=20.0):
                    decrypted_token = decode_data("access_token", config_data["access_token"])
                    ok = fetch_and_cache_credentials(decrypted_token)
                    if not ok and not private_key: errorMessage("Unable to get NordVPN credentials. Is your Access Token valid?")
                    refresh_plan_info()
            except Exception as e: warnMessage(f"Background credential/plan sync failed: {trace(e)}")
        run_cache(force=False)

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
                pystray.MenuItem(unicon(differ_to_health), lambda icon, item: None, enabled=False, visible=lambda icon: session_data["connected"] == True),
                pystray.MenuItem(unicon(differ_to_session_time), lambda icon, item: None, enabled=False, visible=lambda icon: session_data["connected"] == True),
                pystray.MenuItem(unicon(differ_to_data_usage), lambda icon, item: None, enabled=False, visible=lambda icon: session_data["connected"] == True),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem(unicon(differ_from_status_text), create_mini_func(differ_from_status_action, connected=True)),
                pystray.MenuItem("Reconnect Servers", create_mini_func(differ_from_status_action2, connected=True), visible=lambda icon: session_data["connected"] == True),
                pystray.MenuItem("Pause VPN", pystray.Menu(
                    *[pystray.MenuItem(lbl, create_mini_func(functools.partial(pause_vpn, secs, lbl))) for secs, lbl in pause_intervals]
                ), visible=lambda icon: session_data["connected"] == True),
                pystray.MenuItem("Cancel Pause", create_mini_func(cancel_pause_vpn), visible=lambda icon: is_vpn_paused()),
                pystray.MenuItem("Servers", lambda icon, item: ui_queue.put(connect_selected_server_button)),
                pystray.MenuItem("History", lambda icon, item: ui_queue.put(connect_history_server_button)),
                pystray.MenuItem("Configuration", pystray.Menu(
                    pystray.MenuItem("Connection", pystray.Menu(
                        pystray.MenuItem(unicon(differ_to_dns), lambda icon, item: None, enabled=False),
                        pystray.MenuItem(unicon(differ_to_nord_acc), lambda icon, item: None, enabled=False),
                        pystray.MenuItem(unicon(differ_to_channel), lambda icon, item: None, enabled=False),
                        pystray.MenuItem(unicon(differ_to_keepalive), lambda icon, item: None, enabled=False),
                        pystray.MenuItem(unicon(differ_to_mtu), lambda icon, item: None, enabled=False),
                        pystray.MenuItem(unicon(differ_to_server_group), lambda icon, item: None, enabled=False),
                        pystray.MenuItem(unicon(differ_to_allowlist), lambda icon, item: None, enabled=False),
                        pystray.Menu.SEPARATOR,
                        pystray.MenuItem("Change DNS Servers", lambda icon, item: ui_queue.put(change_dns), visible=lambda icon: check_threat_protection_lite() == False),
                        pystray.MenuItem("Change Access Token", lambda icon, item: ui_queue.put(change_access_token)),
                        pystray.MenuItem("Change Connection Channel", lambda icon, item: ui_queue.put(change_connection_channel)),
                        pystray.MenuItem("Change Persistent Keepalive", lambda icon, item: ui_queue.put(change_keepalive)),
                        pystray.MenuItem("Change WireGuard MTU", lambda icon, item: ui_queue.put(change_mtu)),
                        pystray.MenuItem("Change Allowlisted Subnets", lambda icon, item: ui_queue.put(change_allowlisted_subnets)),
                        pystray.MenuItem("Server Group", pystray.Menu(
                            *[pystray.MenuItem(g_lbl, unicon(change_server_group, [g_key]), checked=unicon(check_server_group, [g_key]), radio=True) for g_key, (g_lbl, _) in nord_server_groups.items()]
                        )),
                        pystray.MenuItem(unicon(differ_to_def_loc), lambda icon, item: ui_queue.put(save_auto_selected_server_button)),
                        pystray.Menu.SEPARATOR,
                        pystray.MenuItem("Load Swapping", unicon(change_load_swapping), checked=unicon(check_load_swapping), radio=True),
                        pystray.MenuItem(unicon(differ_to_load_swap_usage), lambda icon, item: ui_queue.put(change_load_swap_threshold), visible=unicon(check_load_swapping)),
                        pystray.MenuItem("Auto Reconnect", unicon(change_auto_connect), checked=unicon(check_auto_connect), radio=True),
                        pystray.MenuItem("Split Private LAN Routing", unicon(change_split_lan_routing), checked=unicon(check_split_lan_routing), radio=True),
                        pystray.MenuItem("Routing Leak Protection", unicon(change_routing_leak_protection), checked=unicon(check_routing_leak_protection), radio=True),
                        pystray.MenuItem("Instant-Reject IPv6", unicon(change_instant_reject_ipv6), checked=unicon(check_instant_reject_ipv6), radio=True),
                        pystray.MenuItem("Loss Connect Protection", unicon(change_loss_connect_protection), checked=unicon(check_loss_connect_protection), radio=True),
                        pystray.MenuItem("Server Health Monitor", unicon(change_server_health_monitor), checked=unicon(check_server_health_monitor), radio=True),
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
                        pystray.MenuItem("Debug Mode", unicon(change_debug_mode), checked=unicon(check_debug_mode), radio=True),
                        pystray.MenuItem("Track Servers Connected", unicon(change_track_servers), checked=unicon(check_track_servers), radio=True),
                    )),
                    pystray.MenuItem("Tools", pystray.Menu(
                        pystray.MenuItem("Set Shortcuts", create_mini_func(unicon(setup_shortcuts, [True]))),
                        pystray.MenuItem("Clear Server Connected History", create_mini_func(unicon(clear_server_connected_history))),
                        pystray.MenuItem("Brute End Wireguard", create_mini_func(brute_end_wireguard)),
                        pystray.MenuItem("Reset DNS Cache", create_mini_func(reset_dns_cache)),
                        pystray.MenuItem("Reset Server Cache", lambda icon, item: addToThread(run_cache, (True,))),
                        pystray.MenuItem("Reset Networking", lambda icon, item: ui_queue.put(reset_networking)),
                        pystray.MenuItem("Clear Wireguard .conf Files", lambda icon, item: ui_queue.put(clear_wireguard_conf_files)),
                        pystray.MenuItem("Clear Configuration", lambda icon, item: ui_queue.put(clear_configuration)),
                        pystray.MenuItem("Clear Configuration (except Account)", lambda icon, item: ui_queue.put(clear_configuration_except_account)),
                    )),
                )),
                pystray.MenuItem("About NordWireConnect", unicon(about)),
                pystray.MenuItem("Check for Updates", create_mini_func(check_for_updates)),
                pystray.MenuItem("Check for Wireguard Updates", create_mini_func(check_for_wireguard_updates)),
                pystray.MenuItem("Quit", unicon(quit_app))
            ),
            darwin_nsapplication=shared_app if main_os == "Darwin" else None
        )
    except Exception as e: errorMessage(f"Unable to run app. Exception: {trace(e)}"); sys.exit(1); return

    # Run Main Loop
    try:
        mainMessage("Finishing app load!")
        apply_window_icon(tk_root, True)
        pystray_icon.run_detached(unicon(setup))
        threading.Thread(target=background_startup_sync, daemon=True).start()
        addToThread(connect_session, (), connection=True)
        threading.Thread(target=worker, daemon=True).start()
        threading.Thread(target=connection_monitor_loop, daemon=True).start()
        tk_root.after(0, update_ui_timer)
        if config_data.get("check_for_updates", True) == True: threading.Thread(target=auto_check_for_updates, daemon=True).start()
        tk_root.after(100, check_stop_flag)
        tk_root.after(50, poll_ui_queue)
        tk.mainloop()
    except Exception as e: errorMessage(f"Unable to run app. Exception: {trace(e)}"); sys.exit(1); return
if __name__ == "__main__": app()