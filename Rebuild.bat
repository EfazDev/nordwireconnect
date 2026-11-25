@echo off
setlocal
set "PATH=%~dp0UPX;%PATH%"

rem Build & Sign NordWireConnect (Nuitka Standalone)
echo false> debug_flag
if exist Installer rmdir /s /q Installer
python -m nuitka ^
    --standalone ^
    --lto=no ^
    --onefile-no-compression ^
    --assume-yes-for-downloads ^
    --remove-output ^
    --enable-plugin=tk-inter ^
    --enable-plugin=upx ^
    --upx-binary="%~dp0UPX\upx.exe" ^
    --windows-console-mode=disable ^
    --windows-icon-from-ico="Resources/app_icon.ico" ^
    --include-data-files="debug_flag=debug_flag" ^
    --include-module=PyKits ^
    --include-module=System ^
    --include-module=Main ^
    --include-module=WindowsService ^
    --include-module=win32timezone ^
    --include-module=win32cred ^
    --include-module=servicemanager ^
    --include-module=win32serviceutil ^
    --include-package=truststore ^
    --company-name="EfazDev" ^
    --product-name="NordWireConnect" ^
    --file-version="1.3.5.0" ^
    --product-version="1.3.5.0" ^
    --output-dir="Installer" ^
    --output-filename="NordWireConnect.exe" ^
    NordWireConnect.py
if errorlevel 1 exit /b %errorlevel%

move "Installer\NordWireConnect.dist" "Installer\NordWireConnect"
copy /Y "Installer\NordWireConnect\NordWireConnect.exe" "Installer\NordWireConnect\NordWireConnectService.exe" >nul
if not defined SKIP_SIGN signtool sign /a /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "Installer\NordWireConnect\NordWireConnect.exe"
if not defined SKIP_SIGN signtool sign /a /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "Installer\NordWireConnect\NordWireConnectService.exe"
if not exist dist mkdir dist

rem Create NordWireConnect Installation ZIP
powershell Compress-Archive -Path "Installer\NordWireConnect\*" -Force -DestinationPath "dist\NordWireConnect.zip"
rmdir /s /q Installer

rem Build & Sign NordWireConnectInstaller (Nuitka Onefile)
python -m nuitka ^
    --onefile ^
    --lto=no ^
    --assume-yes-for-downloads ^
    --remove-output ^
    --windows-console-mode=force ^
    --windows-uac-admin ^
    --windows-icon-from-ico="Resources/app_icon.ico" ^
    --include-data-dir="Resources=Resources" ^
    --include-data-files="dist\NordWireConnect.zip=NordWireConnect.zip" ^
    --include-module=PyKits ^
    --include-module=win32api ^
    --include-module=win32con ^
    --company-name="EfazDev" ^
    --product-name="NordWireConnect Installer" ^
    --file-version="1.3.5.0" ^
    --product-version="1.3.5.0" ^
    --output-dir="dist" ^
    --output-filename="NordWireConnectInstaller.exe" ^
    Installer.py
if errorlevel 1 exit /b %errorlevel%
if not defined SKIP_SIGN signtool sign /a /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "dist\NordWireConnectInstaller.exe"
