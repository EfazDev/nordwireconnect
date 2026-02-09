@echo off

rem Build & Sign NordWireConnect
pyinstaller ^
    --clean ^
    --distpath Installer ^
    --upx-dir "UPX" ^
    --noconfirm ^
    NordWireConnect.spec
signtool sign /a /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "Installer\NordWireConnect\NordWireConnect.exe"
signtool sign /a /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "Installer\NordWireConnect\NordWireConnectService.exe"

rem Create NordWireConnect Installation ZIP
powershell Compress-Archive -Path Installer\NordWireConnect\* -Update -DestinationPath dist\NordWireConnect.zip
rmdir /s /q Installer build

rem Build & Sign NordWireConnectInstaller
pyinstaller ^
    -i "Resources/app_icon.ico" ^
    --hidden-import PyKits ^
    --collect-submodules NordWireConnectInstaller ^
    --add-data "Resources;Resources" ^
    --add-data "dist\NordWireConnect.zip;." ^
    --argv-emulation ^
    --clean ^
    --uac-admin ^
    --upx-dir "UPX" ^
    -n NordWireConnectInstaller ^
    --onefile Installer.py
signtool sign /a /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "dist\NordWireConnectInstaller.exe"