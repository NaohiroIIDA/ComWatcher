# ComWatcher

COM Watcher is a lightweight Windows tray application (WPF, .NET 8) that monitors serial (COM) ports and notifies you when devices are added or removed. It focuses on USB serial devices and shows friendly names and VID/PID when available.

## Features
- Tray-resident COM port monitor (window hidden at startup)
- 1-second polling with diff detection (added/removed)
- USB serial filtering, friendly name display, VID/PID extraction
- Toast balloon on addition (removal toast optional)
- Robust acquisition using WMI with SerialPort fallback

## Build
- Prerequisites: .NET 8 SDK, Windows 10/11
- Build & publish (single-file, self-contained, English resources only):

```powershell
# from the ComWatcher project directory
 dotnet publish ComWatcher.csproj -c Release -r win-x64 -p:PublishSingleFile=true -p:SelfContained=true -p:SatelliteResourceLanguages=en
```

Artifacts will be in:
```
./bin/Release/net8.0-windows/win-x64/publish/
```

## Usage
- Run `ComWatcher.exe` (starts in the system tray)
- Left click tray icon: show window
- Right click tray icon: menu (Show, Exit)
- The list shows USB serial COM ports; the most recently inserted appears at the top

## Notes
- Single-file publish may still require satellite assemblies for UI localization; this release uses English-only (`SatelliteResourceLanguages=en`) so the EXE should work standalone.
- If you prefer multi-language UI resources, publish the entire folder and keep language subfolders (e.g., `ja`).

## Distribution
For end-user distribution instructions, see [docs/DISTRIBUTION.md](docs/DISTRIBUTION.md).

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
