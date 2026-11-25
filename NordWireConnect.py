"""
NordWireConnect

Made by @EfazDev
https://www.efaz.dev
"""

# Modules
import os
import sys

# Main Runtime
def get_exe_path() -> str:
    if os.name == "nt":
        try:
            import ctypes
            buf = ctypes.create_unicode_buffer(1024)
            length = ctypes.windll.kernel32.GetModuleFileNameW(None, buf, 1024)
            if length > 0: return buf.value
        except Exception: pass
    return sys.argv[0] or sys.executable
def sanitize_std_handles():
    if os.name != "nt": return
    import ctypes
    from ctypes import wintypes
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.GetStdHandle.argtypes = [wintypes.DWORD]
    kernel32.GetStdHandle.restype = wintypes.HANDLE
    kernel32.SetStdHandle.argtypes = [wintypes.DWORD, wintypes.HANDLE]
    kernel32.SetStdHandle.restype = wintypes.BOOL
    kernel32.GetCurrentProcess.argtypes = []
    kernel32.GetCurrentProcess.restype = wintypes.HANDLE
    kernel32.DuplicateHandle.argtypes = [
        wintypes.HANDLE, wintypes.HANDLE, wintypes.HANDLE,
        ctypes.POINTER(wintypes.HANDLE), wintypes.DWORD, wintypes.BOOL, wintypes.DWORD
    ]
    kernel32.DuplicateHandle.restype = wintypes.BOOL
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel32.CloseHandle.restype = wintypes.BOOL
    DUPLICATE_SAME_ACCESS = 0x00000002
    cur_proc = kernel32.GetCurrentProcess()
    for std_id in (0xFFFFFFF6, 0xFFFFFFF5, 0xFFFFFFF4):
        h = kernel32.GetStdHandle(std_id)
        if not h: continue
        dup = wintypes.HANDLE()
        if kernel32.DuplicateHandle(cur_proc, h, cur_proc, ctypes.byref(dup), 0, False, DUPLICATE_SAME_ACCESS): kernel32.CloseHandle(dup)
        else: kernel32.SetStdHandle(std_id, None)
def main():
    sanitize_std_handles()
    exe_name = os.path.basename(get_exe_path()).lower()
    if "service" in exe_name or "--service" in sys.argv:
        import WindowsService
        WindowsService.main()
    else:
        import Main
        Main.app()
if __name__ == "__main__": main()