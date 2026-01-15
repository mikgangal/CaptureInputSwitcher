# CaptureInputSwitcher

A command-line tool to switch video inputs (HDMI/DVI) on CY3014/StarTech capture devices.

## Use Case

Programmatically switch between video inputs on capture cards that store their input selection in the Windows registry. Useful for:
- OBS Studio automation
- Scripted input switching
- Hotkey-triggered source changes

## Supported Devices

- StarTech.com USB3HDCAP (CY3014)
- Other Yuan/Multimedia capture devices with the `AnalogCrossbarVideoInputProperty` registry setting

## Building

Uses the built-in Windows .NET Framework compiler (no SDK required):

```cmd
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe /out:CaptureInputSwitcher.exe /target:exe /platform:x86 CaptureInputSwitcher.cs
```

## Usage

```cmd
:: List all video capture devices (switchable ones are marked)
CaptureInputSwitcher.exe list

:: Show available inputs for a device
CaptureInputSwitcher.exe inputs CY3014

:: Switch to HDMI (input 0) - requires admin rights
CaptureInputSwitcher.exe switch 0 CY3014

:: Switch to DVI (input 1) - requires admin rights
CaptureInputSwitcher.exe switch 1 CY3014

:: Get current input
CaptureInputSwitcher.exe get CY3014
```

The device name is a partial, case-insensitive match.

## Example Output

```
> CaptureInputSwitcher.exe list
Video Capture Devices:
----------------------
  Integrated Camera
  CY3014 USB, Analog 01 Capture [Switchable]
  OBS Virtual Camera

> CaptureInputSwitcher.exe inputs CY3014
Looking for device matching: CY3014
Found: CY3014 USB, Analog 01 Capture

Available inputs:
  0: HDMI <-- CURRENT
  1: DVI
```

## Requirements

- Windows 7 or later
- .NET Framework 4.0+ (included in Windows)
- **Administrator rights required** for switching inputs (writes to HKLM registry)

## How It Works

The tool modifies the `AnalogCrossbarVideoInputProperty` registry value in the device's driver settings:
```
HKLM\SYSTEM\CurrentControlSet\Control\Class\{...}\xxxx\AnalogCrossbarVideoInputProperty
```

After changing the registry value, you may need to restart the capture source in OBS for changes to take effect.

## Batch Files for Easy Switching

Create these batch files and run them as administrator:

**switch-hdmi.bat**
```batch
@echo off
CaptureInputSwitcher.exe switch 0 CY3014
```

**switch-dvi.bat**
```batch
@echo off
CaptureInputSwitcher.exe switch 1 CY3014
```

## License

MIT
