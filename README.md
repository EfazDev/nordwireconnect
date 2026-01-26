<h1 align="center"><img align="center" src="resources/app_icon.png" width="40" height="40"> NordWireConnect</h1>
<h2 align="center">Connect to NordVPN using Wireguard!</h2>

## What is NordWireConnect?
NordWireConnect is a Python application where you can connect to NordVPN using the official Wireguard app without using the NordVPN app. This may be useful due to startup issues in the NordVPN app and you can't get the most suggested server in the original Wireguard app.

## Access Token Setup
NordVPN Access Token is required in order to communicate to NordVPN servers. Find it by logging into your NordVPN account and create an access token for the duration you want. [NordVPN Dashboard](https://my.nordaccount.com/dashboard/nordvpn/)

## Installation
NordWireConnect and Wireguard will require Administrator permissions before installing. After opening the EXE and allowing the UAC permission, it should start installing and launch automatically. After that, it will load in the bottom right of your task bar (right click to open the menu). After that, set up your Access Token in the Configuration menu and connect to a server!

## Features & Tweaks
- Custom DNS (default is set to NordVPN)
- Private LAN Splitting (Connect to private IP addresses while using VPN for the internet) 
- Auto Reconnect (similar to Wireguard where if you restart while connected, you reconnect again after reboot)
- Load Swapping (If load is more than or equal to 50%, reconnect servers)
- Optimize Server List (Load only 500 servers on the Server dropdown menu instead of all 8000+ servers)
- Loss Connect Protection (Auto Reconnect if connection was lost by Wireguard)
- Set Shortcuts (Creates Desktop Shortcuts)
- Brute End Wireguard (As a last resort, if Wireguard breaks, you can force end the process.)

## Rebuilding
The following command (package dependencies) and Python 3.12+ are needed to rebuild the NordWireConnect app: `pip install pywin32 Pillow pystray pyinstaller`. After that all the packages are installed, run the Rebuild.bat file to build and create an exe file in a dist folder (will automatically create if not found).

## Uninstall
If you want to uninstall the program, you may go to the Installed Apps menu in Windows Settings and uninstall the app. This will also require administrative permissions.