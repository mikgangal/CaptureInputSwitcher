using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Win32;

namespace CaptureInputSwitcher
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length == 0)
            {
                return PrintUsage();
            }

            string command = args[0].ToLower();
            string deviceName = args.Length > 1 ? args[1] : null;

            try
            {
                switch (command)
                {
                    case "list":
                        return ListDevices();
                    case "inputs":
                        return ListInputs(deviceName ?? "CY3014");
                    case "switch":
                        if (args.Length < 2)
                        {
                            Console.WriteLine("Error: switch requires an input number (0 or 1)");
                            return 1;
                        }
                        return SwitchInput(
                            args.Length >= 3 ? args[2] : "CY3014",
                            int.Parse(args[1]));
                    case "get":
                        return GetCurrentInput(deviceName ?? "CY3014");
                    default:
                        return PrintUsage();
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: " + ex.Message);
                return 1;
            }
        }

        static int PrintUsage()
        {
            Console.WriteLine("Capture Device Input Switcher");
            Console.WriteLine("=============================");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  CaptureInputSwitcher list                    - List all video capture devices");
            Console.WriteLine("  CaptureInputSwitcher inputs [device]         - List available inputs for device");
            Console.WriteLine("  CaptureInputSwitcher switch <input> [device] - Switch to input (0=HDMI, 1=DVI)");
            Console.WriteLine("  CaptureInputSwitcher get [device]            - Get current input");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  CaptureInputSwitcher list");
            Console.WriteLine("  CaptureInputSwitcher inputs CY3014");
            Console.WriteLine("  CaptureInputSwitcher switch 0 CY3014         - Switch to HDMI (input 0)");
            Console.WriteLine("  CaptureInputSwitcher switch 1 CY3014         - Switch to DVI (input 1)");
            Console.WriteLine();
            Console.WriteLine("Note: Device name is a partial match (case-insensitive)");
            Console.WriteLine("      This tool works with CY3014/StarTech capture devices");
            return 0;
        }

        [DllImport("ole32.dll")]
        static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

        static int ListDevices()
        {
            Console.WriteLine("Video Capture Devices:");
            Console.WriteLine("----------------------");

            var devices = EnumerateVideoDevices();
            if (devices.Count == 0)
            {
                Console.WriteLine("  No video capture devices found.");
                return 1;
            }

            foreach (var device in devices)
            {
                string driverKey = FindDriverKey(device.Item2);
                string marker = !string.IsNullOrEmpty(driverKey) ? " [Switchable]" : "";
                Console.WriteLine("  " + device.Item1 + marker);
            }
            return 0;
        }

        static int ListInputs(string deviceName)
        {
            Console.WriteLine("Looking for device matching: " + deviceName);

            var devices = EnumerateVideoDevices();
            foreach (var device in devices)
            {
                if (device.Item1.IndexOf(deviceName, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine("Found: " + device.Item1);

                    string driverKey = FindDriverKey(device.Item2);
                    if (string.IsNullOrEmpty(driverKey))
                    {
                        Console.WriteLine("Device does not have switchable inputs via registry.");
                        return 1;
                    }

                    int currentInput = GetRegistryInput(driverKey);
                    Console.WriteLine();
                    Console.WriteLine("Available inputs:");
                    Console.WriteLine("  0: HDMI" + (currentInput == 0 ? " <-- CURRENT" : ""));
                    Console.WriteLine("  1: DVI" + (currentInput == 1 ? " <-- CURRENT" : ""));
                    return 0;
                }
            }

            Console.WriteLine("Device not found.");
            return 1;
        }

        static int SwitchInput(string deviceName, int input)
        {
            if (input < 0 || input > 1)
            {
                Console.WriteLine("Invalid input. Use 0 for HDMI or 1 for DVI.");
                return 1;
            }

            Console.WriteLine("Looking for device matching: " + deviceName);

            var devices = EnumerateVideoDevices();
            foreach (var device in devices)
            {
                if (device.Item1.IndexOf(deviceName, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine("Found: " + device.Item1);

                    string driverKey = FindDriverKey(device.Item2);
                    if (string.IsNullOrEmpty(driverKey))
                    {
                        Console.WriteLine("Device does not support input switching via registry.");
                        return 1;
                    }

                    bool success = SetRegistryInput(driverKey, input);
                    if (success)
                    {
                        string inputName = input == 0 ? "HDMI" : "DVI";
                        Console.WriteLine("Successfully switched to input " + input + " (" + inputName + ")");
                        Console.WriteLine();
                        Console.WriteLine("NOTE: You may need to restart the capture source in OBS for changes to take effect.");
                        return 0;
                    }
                    else
                    {
                        Console.WriteLine("Failed to switch input. Try running as administrator.");
                        return 1;
                    }
                }
            }

            Console.WriteLine("Device not found.");
            return 1;
        }

        static int GetCurrentInput(string deviceName)
        {
            var devices = EnumerateVideoDevices();
            foreach (var device in devices)
            {
                if (device.Item1.IndexOf(deviceName, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    string driverKey = FindDriverKey(device.Item2);
                    if (string.IsNullOrEmpty(driverKey))
                    {
                        Console.WriteLine("Device does not support input switching via registry.");
                        return -1;
                    }

                    int currentInput = GetRegistryInput(driverKey);
                    string inputName = currentInput == 0 ? "HDMI" : "DVI";
                    Console.WriteLine("Current input: " + currentInput + " (" + inputName + ")");
                    return currentInput;
                }
            }

            Console.WriteLine("Device not found.");
            return -1;
        }

        static List<Tuple<string, string>> EnumerateVideoDevices()
        {
            var devices = new List<Tuple<string, string>>();

            Type devEnumType = Type.GetTypeFromCLSID(new Guid("62BE5D10-60EB-11d0-BD3B-00A0C911CE86"));
            object devEnumObj = Activator.CreateInstance(devEnumType);
            ICreateDevEnum devEnum = (ICreateDevEnum)devEnumObj;

            IEnumMoniker enumMoniker;
            Guid videoInputCategory = new Guid("860BB310-5D01-11D0-BD3B-00A0C911CE86");
            int hr = devEnum.CreateClassEnumerator(ref videoInputCategory, out enumMoniker, 0);

            if (hr != 0 || enumMoniker == null)
            {
                Marshal.ReleaseComObject(devEnumObj);
                return devices;
            }

            IMoniker[] moniker = new IMoniker[1];
            while (enumMoniker.Next(1, moniker, IntPtr.Zero) == 0)
            {
                IBindCtx bindCtx;
                CreateBindCtx(0, out bindCtx);

                object bagObj;
                Guid propBagGuid = typeof(IPropertyBag).GUID;
                moniker[0].BindToStorage(bindCtx, null, ref propBagGuid, out bagObj);

                string name = "";
                string devicePath = "";
                if (bagObj != null)
                {
                    IPropertyBag bag = (IPropertyBag)bagObj;
                    object nameObj, pathObj;
                    bag.Read("FriendlyName", out nameObj, null);
                    bag.Read("DevicePath", out pathObj, null);
                    name = nameObj != null ? nameObj.ToString() : "Unknown Device";
                    devicePath = pathObj != null ? pathObj.ToString() : "";
                    Marshal.ReleaseComObject(bag);
                }

                devices.Add(new Tuple<string, string>(name, devicePath));

                Marshal.ReleaseComObject(bindCtx);
                Marshal.ReleaseComObject(moniker[0]);
            }

            Marshal.ReleaseComObject(enumMoniker);
            Marshal.ReleaseComObject(devEnumObj);

            return devices;
        }

        static string FindDriverKey(string devicePath)
        {
            // Extract VID and PID from device path
            // Format: \\?\usb#vid_1164&pid_f533#...
            if (string.IsNullOrEmpty(devicePath))
                return null;

            string vidPid = "";
            int vidIdx = devicePath.IndexOf("vid_", StringComparison.OrdinalIgnoreCase);
            if (vidIdx >= 0)
            {
                int hashIdx = devicePath.IndexOf('#', vidIdx);
                if (hashIdx > vidIdx)
                {
                    vidPid = devicePath.Substring(vidIdx, hashIdx - vidIdx).ToUpper();
                }
            }

            if (string.IsNullOrEmpty(vidPid))
                return null;

            // Search for the device in USB enum
            try
            {
                using (RegistryKey usbKey = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Enum\USB"))
                {
                    if (usbKey == null) return null;

                    foreach (string subKeyName in usbKey.GetSubKeyNames())
                    {
                        if (subKeyName.IndexOf(vidPid, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            using (RegistryKey vidPidKey = usbKey.OpenSubKey(subKeyName))
                            {
                                if (vidPidKey == null) continue;

                                foreach (string instanceName in vidPidKey.GetSubKeyNames())
                                {
                                    using (RegistryKey instanceKey = vidPidKey.OpenSubKey(instanceName))
                                    {
                                        if (instanceKey == null) continue;

                                        object driverRef = instanceKey.GetValue("Driver");
                                        if (driverRef != null)
                                        {
                                            string driverPath = @"SYSTEM\CurrentControlSet\Control\Class\" + driverRef.ToString();

                                            // Verify this driver has the AnalogCrossbarVideoInputProperty
                                            using (RegistryKey driverKey = Registry.LocalMachine.OpenSubKey(driverPath))
                                            {
                                                if (driverKey != null && driverKey.GetValue("AnalogCrossbarVideoInputProperty") != null)
                                                {
                                                    return driverPath;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                return null;
            }

            return null;
        }

        static int GetRegistryInput(string driverKeyPath)
        {
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(driverKeyPath))
                {
                    if (key != null)
                    {
                        object value = key.GetValue("AnalogCrossbarVideoInputProperty");
                        if (value != null)
                        {
                            return Convert.ToInt32(value);
                        }
                    }
                }
            }
            catch { }
            return -1;
        }

        static bool SetRegistryInput(string driverKeyPath, int input)
        {
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(driverKeyPath, true))
                {
                    if (key != null)
                    {
                        key.SetValue("AnalogCrossbarVideoInputProperty", input, RegistryValueKind.DWord);
                        return true;
                    }
                }
            }
            catch (UnauthorizedAccessException)
            {
                Console.WriteLine("Access denied. Run as administrator to switch inputs.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            return false;
        }
    }

    // COM Interop definitions

    [ComImport, Guid("29840822-5B84-11D0-BD3B-00A0C911CE86"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ICreateDevEnum
    {
        [PreserveSig]
        int CreateClassEnumerator([In] ref Guid type, out IEnumMoniker enumMoniker, [In] int flags);
    }

    [ComImport, Guid("55272A00-42CB-11CE-8135-00AA004BB851"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IPropertyBag
    {
        [PreserveSig]
        int Read([In, MarshalAs(UnmanagedType.LPWStr)] string propName, [Out, MarshalAs(UnmanagedType.Struct)] out object val, object errorLog);
        [PreserveSig]
        int Write([In, MarshalAs(UnmanagedType.LPWStr)] string propName, [In] ref object val);
    }
}
