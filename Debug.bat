@echo off

rem Build & Sign NordWireConnect
pyinstaller ^
    --clean ^
    --distpath Installer ^
    --upx-dir "UPX" ^
    --noconfirm ^
    NordWireConnectDebug.spec
signtool sign /a /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "Installer\NordWireConnect\NordWireConnect.exe"
signtool sign /a /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "Installer\NordWireConnect\NordWireConnectService.exe"
mkdir dist

rem Create NordWireConnect Installation ZIP
powershell Compress-Archive -Path Installer\NordWireConnect\* -Update -DestinationPath dist\NordWireConnect.zip
rmdir /s /q Installer build

rem Build & Sign NordWireConnectDebugInstaller
pyinstaller ^
    -i "resources/app_icon.ico" ^
    --hidden-import PyKits ^
    --collect-submodules NordWireConnectDebugInstaller ^
    --add-data "resources;resources" ^
    --add-data "dist\NordWireConnect.zip;." ^
    --argv-emulation ^
    --clean ^
    --uac-admin ^
    --upx-dir "UPX" ^
    -n NordWireConnectDebugInstaller ^
    --onefile Installer.py
signtool sign /a /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "dist\NordWireConnectDebugInstaller.exe"