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
def main():
    exe_name = os.path.basename(get_exe_path()).lower()
    if "service" in exe_name or "--service" in sys.argv:
        import WindowsService
        WindowsService.main()
    else:
        import Main
        Main.app()
if __name__ == "__main__": main()