# ComWatcher - Distribution Package

English | [日本語](DISTRIBUTION.md)

## Contents
- `ComWatcher.exe` - Main program
- `*.dll` - Required WPF runtime libraries

## Installation
1. Extract the ZIP file
2. Run `ComWatcher.exe`
3. The app will start minimized in the system tray
4. USB-Serial ports will be automatically detected and notified

## System Requirements
- Windows 10/11 (64-bit)
- No .NET Runtime required (self-contained)

## Notes
- Windows Defender may show a warning on first run
- Keep the EXE and DLLs in the same folder
- To uninstall: Simply delete the folder

## Tray Icon Controls
- **Left click** - Show window
- **Right click** - Show menu (Exit)

## Features
- Automatic detection of USB-Serial ports (Arduino, etc.)
- Balloon notification when ports are added
- Device name and VID/PID display
- Most recently connected port appears at the top
