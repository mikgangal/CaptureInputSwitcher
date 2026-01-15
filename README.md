# CaptureInputSwitcher

A command-line tool to switch video inputs (HDMI/DVI) on CY3014/StarTech capture devices, with OBS integration for live switching.

## Features

- Switch capture card inputs via command line
- No UAC prompts (uses scheduled tasks)
- Live OBS source restart via WebSocket
- Perfect for foot pedal / remote triggering

## Supported Devices

- StarTech.com USB3HDCAP (CY3014)
- Other Yuan/Multimedia capture devices with the `AnalogCrossbarVideoInputProperty` registry setting

## Quick Start

### One-Time Setup (requires admin)

1. **Run `setup_tasks.bat` as Administrator** - creates scheduled tasks that run elevated without UAC prompts

2. **Enable OBS WebSocket:**
   - In OBS: Tools → WebSocket Server Settings
   - Check "Enable WebSocket server"
   - Note the password (update batch files if different from default)

3. **Edit batch files** if your OBS password differs:
   - Open `SwitchToHDMI.bat` and `SwitchToDVI.bat`
   - Update the `--password` parameter

### Usage

Simply run the batch files - no admin prompts, works from foot pedals or hotkeys:

```cmd
SwitchToHDMI.bat    :: Switch to HDMI input
SwitchToDVI.bat     :: Switch to DVI input
```

These batch files:
1. Change the registry setting (via scheduled task - no UAC)
2. Restart the OBS capture source (via WebSocket - makes OBS see the change live)

## Manual Commands

```cmd
:: List all video capture devices
CaptureInputSwitcher.exe list

:: Show available inputs
CaptureInputSwitcher.exe inputs CY3014

:: Get current input
CaptureInputSwitcher.exe get CY3014

:: Switch inputs (requires admin OR use scheduled tasks)
CaptureInputSwitcher.exe switch 0 CY3014   :: HDMI
CaptureInputSwitcher.exe switch 1 CY3014   :: DVI

:: Restart OBS source (after registry change)
OBSControl.exe --password "yourpassword" --scene "Scene" --source "Video Capture Device 2"
```

## Files

| File | Description |
|------|-------------|
| `CaptureInputSwitcher.exe` | Registry-based input switcher |
| `OBSControl.exe` | OBS WebSocket source restarter |
| `SwitchToHDMI.bat` | One-click HDMI switch + OBS restart |
| `SwitchToDVI.bat` | One-click DVI switch + OBS restart |
| `setup_tasks.bat` | Creates scheduled tasks (run as admin once) |

## How It Works

1. **Registry Change:** The tool modifies `AnalogCrossbarVideoInputProperty` in the driver's registry key. This requires admin rights, so we use scheduled tasks to avoid UAC prompts.

2. **OBS Restart:** The driver only reads the registry at initialization. OBSControl.exe connects to OBS via WebSocket and disables/enables the capture source, forcing it to reinitialize with the new input.

3. **No Network Required:** OBS WebSocket runs on localhost only - no firewall or network admin permissions needed.

## Building

Uses the built-in Windows .NET Framework compiler (no SDK required):

```cmd
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe /out:CaptureInputSwitcher.exe /target:exe /platform:x86 CaptureInputSwitcher.cs
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe /out:OBSControl.exe /target:exe OBSControl.cs
```

## Requirements

- Windows 7 or later
- .NET Framework 4.0+ (included in Windows)
- OBS Studio with WebSocket plugin (included in OBS 28+)

## Customization

**Different scene/source names:**
```cmd
OBSControl.exe --password "pass" --scene "MyScene" --source "My Capture"
```

**Different WebSocket port:**
```cmd
OBSControl.exe --password "pass" --port 4455
```

## License

MIT
