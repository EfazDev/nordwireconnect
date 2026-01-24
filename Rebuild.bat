@echo off

rem Build & Sign NordWireConnectService
pyinstaller --hidden-import pywin32 --hidden-import logging.config --hidden-import win32timezone --windowed --noconsole --hidden-import win32cred --noconfirm --hidden-import pywintypes -i "resources/app_icon.ico" -n NordWireConnectService --clean --distpath installer --strip Service.py
signtool sign /a /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "installer\NordWireConnectService\NordWireConnectService.exe"

rem Build & Sign NordWireConnect
pyinstaller --hidden-import logging.config --hidden-import PyKits -i "resources/app_icon.ico" --argv-emulation --windowed --noconfirm --noconsole --clean -n NordWireConnect --distpath installer --strip Main.py
signtool sign /a /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "installer\NordWireConnect\NordWireConnect.exe"

rem Create NordWireConnect Installation ZIP
powershell Compress-Archive -Path installer\* -Update -DestinationPath dist\NordWireConnect.zip

rem Build & Sign NordWireConnectInstaller
pyinstaller --hidden-import PyKits -i "resources/app_icon.ico" --add-data "resources;resources" --add-data "dist\NordWireConnect.zip;." --argv-emulation --clean -n NordWireConnectInstaller --uac-admin --strip --onefile Installer.py
signtool sign /a /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "dist\NordWireConnectInstaller.exe"