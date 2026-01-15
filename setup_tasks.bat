@echo off
echo Creating scheduled tasks for capture input switching...
echo These tasks run elevated without UAC prompts.
echo.

schtasks /create /tn "CaptureSwitch_HDMI" /tr "\"%~dp0CaptureInputSwitcher.exe\" switch 0 CY3014" /sc once /st 00:00 /rl highest /f
schtasks /create /tn "CaptureSwitch_DVI" /tr "\"%~dp0CaptureInputSwitcher.exe\" switch 1 CY3014" /sc once /st 00:00 /rl highest /f

echo.
echo Done! You can now use SwitchToHDMI.bat and SwitchToDVI.bat
pause
