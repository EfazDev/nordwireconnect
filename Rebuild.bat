@echo off

rem Build & Sign NordWireConnectService
pyinstaller --hidden-import pywin32 --hidden-import win32timezone --noconsole --hidden-import win32cred --hidden-import pywintypes -i "resources/app_icon.ico" -n NordWireConnectService --windowed --noconsole --clean --onefile Service.py
signtool sign /a /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "dist\NordWireConnectService.exe"

rem Build & Sign NordWireConnect
pyinstaller --hidden-import PyKits -i "resources/app_icon.ico" --add-data "resources;resources" --add-data "dist/NordWireConnectService.exe;." --argv-emulation --windowed --noconsole --clean -n NordWireConnect --onefile Main.py
signtool sign /a /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "dist\NordWireConnect.exe"