@echo off

rem Build & Sign NordWireConnectService
pyinstaller ^
    -i "resources/app_icon.ico" ^
    --hidden-import PyKits ^
    --hidden-import win32timezone ^
    --hidden-import win32cred ^
    --hidden-import pywintypes ^
    --windowed ^
    --noconsole ^
    --noconfirm ^
    --clean ^
    --distpath installer ^
    -n NordWireConnectService ^
    --strip Service.py
signtool sign /a /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "installer\NordWireConnectService\NordWireConnectService.exe"

rem Build & Sign NordWireConnect
pyinstaller ^
    -i "resources/app_icon.ico" ^
    --hidden-import PyKits ^
    --argv-emulation ^
    --windowed ^
    --noconfirm ^
    --noconsole ^
    --clean ^
    --distpath installer ^
    -n NordWireConnect ^
    --strip Main.py
signtool sign /a /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "installer\NordWireConnect\NordWireConnect.exe"

rem Create NordWireConnect Installation ZIP
powershell Compress-Archive -Path installer\* -Update -DestinationPath dist\NordWireConnect.zip

rem Build & Sign NordWireConnectInstaller
pyinstaller ^
    -i "resources/app_icon.ico" ^
    --hidden-import PyKits ^
    --add-data "resources;resources" ^
    --add-data "dist\NordWireConnect.zip;." ^
    --argv-emulation ^
    --clean ^
    --uac-admin ^
    --strip ^
    -n NordWireConnectInstaller ^
    --onefile Installer.py
signtool sign /a /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "dist\NordWireConnectInstaller.exe"