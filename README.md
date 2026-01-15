# CaptureInputSwitcher

A command-line tool to switch video inputs (HDMI/DVI/etc.) on DirectShow capture devices like the USB3HDCAP.

## Use Case

Programmatically switch between video inputs on capture cards that expose their input selection through DirectShow's `IAMCrossbar` interface. Useful for:
- OBS Studio automation
- Scripted input switching
- Hotkey-triggered source changes

## Building

Uses the built-in Windows .NET Framework compiler (no SDK required):

```cmd
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe /out:CaptureInputSwitcher.exe /target:exe /platform:x86 CaptureInputSwitcher.cs
```

## Usage

```cmd
:: List all video capture devices
CaptureInputSwitcher.exe list

:: Show available inputs for a device
CaptureInputSwitcher.exe inputs USB3HDCAP

:: Switch to a specific input pin
CaptureInputSwitcher.exe switch 0 USB3HDCAP   # Switch to pin 0 (e.g., HDMI)
CaptureInputSwitcher.exe switch 1 USB3HDCAP   # Switch to pin 1 (e.g., DVI)

:: Get current input
CaptureInputSwitcher.exe get USB3HDCAP
```

The device name is a partial, case-insensitive match.

## Example Output

```
> CaptureInputSwitcher.exe inputs USB3HDCAP
Looking for device matching: USB3HDCAP

Found 2 input(s) and 1 output(s):

Video Inputs:
  Pin 0: HDMI <-- CURRENT
  Pin 1: DVI
```

## Requirements

- Windows 7 or later
- .NET Framework 4.0+ (included in Windows)
- No admin rights required to run

## License

MIT
