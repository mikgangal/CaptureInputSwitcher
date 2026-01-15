# Zero-Admin Input Switching Implementation Notes

## Goal
Switch CY3014/USB3HDCAP capture device inputs (HDMI/DVI) without any admin privileges.

## Current Problem
The current implementation modifies the registry at:
```
HKLM\SYSTEM\CurrentControlSet\Control\Class\{4d36e96c-e325-11ce-bfc1-08002be10318}\0021\AnalogCrossbarVideoInputProperty
```
This requires admin privileges, which the user doesn't have on the target machine.

## Key Discovery
OBS's "Configure Video" dialog CAN switch inputs without admin. It does this by using the **IKsPropertySet** COM interface to access the device's custom property interface directly at runtime, bypassing the registry.

## Device Details
- Device Name: `CY3014 USB, Analog 01 Capture`
- VID: 1164, PID: F533 (StarTech/Yuan)
- Custom Property GUID: `{D1E5209F-68FD-4529-BEE0-5E7A1F47921F}`
- Input Values: 0 = HDMI, 1 = DVI

## Solution: IKsPropertySet Approach

Replace registry-based switching with direct device property access:

```
DirectShow Graph -> Capture Filter -> IKsPropertySet -> Set Property
                                                         |
                                        GUID: {D1E5209F-68FD-4529-BEE0-5E7A1F47921F}
                                        Property ID: TBD (likely 0)
                                        Value: 0 (HDMI) or 1 (DVI)
```

### C# Implementation

```csharp
// 1. Define IKsPropertySet COM interface
[ComImport, Guid("31EFAC30-515C-11d0-A9AA-00AA0061BE93")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IKsPropertySet
{
    int Set(
        [In] ref Guid guidPropSet,
        [In] int dwPropID,
        [In] IntPtr pInstanceData,
        [In] int cbInstanceData,
        [In] IntPtr pPropData,
        [In] int cbPropData);

    int Get(
        [In] ref Guid guidPropSet,
        [In] int dwPropID,
        [In] IntPtr pInstanceData,
        [In] int cbInstanceData,
        [Out] IntPtr pPropData,
        [In] int cbPropData,
        [Out] out int pcbReturned);

    int QuerySupported(
        [In] ref Guid guidPropSet,
        [In] int dwPropID,
        [Out] out int pTypeSupport);
}

// 2. Get capture filter and query for IKsPropertySet
IBaseFilter captureFilter = GetCaptureFilter("CY3014");
IKsPropertySet props = captureFilter as IKsPropertySet;

// 3. Set the input property
Guid propSetGuid = new Guid("D1E5209F-68FD-4529-BEE0-5E7A1F47921F");
int propId = 0; // May need to discover this
int inputValue = 0; // 0=HDMI, 1=DVI
props.Set(ref propSetGuid, propId, ...);
```

## Implementation Steps

1. Add IKsPropertySet interface definition to CaptureInputSwitcher.cs
2. Add IBaseFilter interface (needed to cast device to IKsPropertySet)
3. Discover the correct property ID - Try 0, 1, etc.
4. Test without elevation - Should work without admin
5. Update batch files - Remove scheduled task dependency

## Fallback: UI Automation

If IKsPropertySet doesn't work, automate the "Configure Video" dialog:
1. Use Windows UI Automation API
2. Open device property dialog programmatically
3. Send click to HDMI/DVI button
4. Close dialog

## Why FFmpeg Won't Work
- FFmpeg only supports standard IAMCrossbar interface
- CY3014 uses custom property interface, not IAMCrossbar
- No way to access custom properties through FFmpeg command line

## OBS Configuration (for testing)
- Scene: "Scene"
- Source: "Video Capture Device 2"
- WebSocket password: "8Alz3MgImnSeBOqa"
- Port: 4455 (default)

## References
- IKsPropertySet: https://learn.microsoft.com/en-us/windows/win32/directshow/ikspropertyset
- DirectShow Property Sets: https://learn.microsoft.com/en-us/windows/win32/directshow/property-sets
