@echo off
REM Switch to HDMI (input 0) and restart OBS source
schtasks /run /tn "CaptureSwitch_HDMI" >nul 2>&1
timeout /t 1 /nobreak >nul
"%~dp0OBSControl.exe" --password "8Alz3MgImnSeBOqa" --scene "Scene" --source "Video Capture Device 2"
