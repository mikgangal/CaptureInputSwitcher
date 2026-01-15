using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

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
                        return ListInputs(deviceName ?? "USB3HDCAP");
                    case "switch":
                        if (args.Length < 2)
                        {
                            Console.WriteLine("Error: switch requires a pin number");
                            return 1;
                        }
                        return SwitchInput(
                            args.Length >= 3 ? args[2] : "USB3HDCAP",
                            int.Parse(args[1]));
                    case "get":
                        return GetCurrentInput(deviceName ?? "USB3HDCAP");
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
            Console.WriteLine("  CaptureInputSwitcher switch <pin> [device]   - Switch to input pin number");
            Console.WriteLine("  CaptureInputSwitcher get [device]            - Get current input");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  CaptureInputSwitcher list");
            Console.WriteLine("  CaptureInputSwitcher inputs USB3HDCAP");
            Console.WriteLine("  CaptureInputSwitcher switch 0 USB3HDCAP      - Switch to first input (usually HDMI)");
            Console.WriteLine("  CaptureInputSwitcher switch 1 USB3HDCAP      - Switch to second input (usually DVI)");
            Console.WriteLine();
            Console.WriteLine("Note: Device name is a partial match (case-insensitive)");
            return 0;
        }

        static int ListDevices()
        {
            Console.WriteLine("Video Capture Devices:");
            Console.WriteLine("----------------------");

            List<string> devices = EnumerateVideoDevices();
            if (devices.Count == 0)
            {
                Console.WriteLine("  No video capture devices found.");
                return 1;
            }

            foreach (string device in devices)
            {
                Console.WriteLine("  " + device);
            }
            return 0;
        }

        static int ListInputs(string deviceName)
        {
            Console.WriteLine("Looking for device matching: " + deviceName);

            using (FilterGraph graph = new FilterGraph(deviceName))
            {
                IAMCrossbar crossbar = graph.FindCrossbar();

                if (crossbar == null)
                {
                    Console.WriteLine("No crossbar found. Device may not support input switching.");
                    return 1;
                }

                int outputPins, inputPins;
                crossbar.get_PinCounts(out outputPins, out inputPins);
                Console.WriteLine();
                Console.WriteLine("Found " + inputPins + " input(s) and " + outputPins + " output(s):");
                Console.WriteLine();

                // Get current routing
                int currentInput = -1;
                for (int outPin = 0; outPin < outputPins; outPin++)
                {
                    int outIndex;
                    PhysicalConnectorType outType;
                    crossbar.get_CrossbarPinInfo(false, outPin, out outIndex, out outType);
                    if (outType == PhysicalConnectorType.Video_VideoDecoder)
                    {
                        crossbar.get_IsRoutedTo(outPin, out currentInput);
                        break;
                    }
                }

                Console.WriteLine("Video Inputs:");
                for (int i = 0; i < inputPins; i++)
                {
                    int relatedPin;
                    PhysicalConnectorType pinType;
                    crossbar.get_CrossbarPinInfo(true, i, out relatedPin, out pinType);
                    string current = (i == currentInput) ? " <-- CURRENT" : "";
                    Console.WriteLine("  Pin " + i + ": " + GetPinTypeName(pinType) + current);
                }

                return 0;
            }
        }

        static int SwitchInput(string deviceName, int inputPin)
        {
            Console.WriteLine("Switching device '" + deviceName + "' to input pin " + inputPin + "...");

            using (FilterGraph graph = new FilterGraph(deviceName))
            {
                IAMCrossbar crossbar = graph.FindCrossbar();

                if (crossbar == null)
                {
                    Console.WriteLine("No crossbar found. Device may not support input switching.");
                    return 1;
                }

                int outputPins, inputPins;
                crossbar.get_PinCounts(out outputPins, out inputPins);

                if (inputPin < 0 || inputPin >= inputPins)
                {
                    Console.WriteLine("Invalid input pin. Device has " + inputPins + " input(s) (0-" + (inputPins - 1) + ")");
                    return 1;
                }

                // Find the video decoder output pin
                int videoOutPin = -1;
                for (int outPin = 0; outPin < outputPins; outPin++)
                {
                    int related;
                    PhysicalConnectorType outType;
                    crossbar.get_CrossbarPinInfo(false, outPin, out related, out outType);
                    if (outType == PhysicalConnectorType.Video_VideoDecoder)
                    {
                        videoOutPin = outPin;
                        break;
                    }
                }

                if (videoOutPin == -1)
                {
                    // Just try output pin 0
                    videoOutPin = 0;
                }

                // Check if routing is possible
                int hr = crossbar.CanRoute(videoOutPin, inputPin);
                if (hr != 0)
                {
                    Console.WriteLine("Cannot route input " + inputPin + " to video output. Error: 0x" + hr.ToString("X8"));
                    return 1;
                }

                // Perform the switch
                hr = crossbar.Route(videoOutPin, inputPin);
                if (hr != 0)
                {
                    Console.WriteLine("Failed to switch input. Error: 0x" + hr.ToString("X8"));
                    return 1;
                }

                int relatedPin;
                PhysicalConnectorType pinType;
                crossbar.get_CrossbarPinInfo(true, inputPin, out relatedPin, out pinType);
                Console.WriteLine("Successfully switched to input " + inputPin + " (" + GetPinTypeName(pinType) + ")");
                return 0;
            }
        }

        static int GetCurrentInput(string deviceName)
        {
            using (FilterGraph graph = new FilterGraph(deviceName))
            {
                IAMCrossbar crossbar = graph.FindCrossbar();

                if (crossbar == null)
                {
                    Console.WriteLine("No crossbar found.");
                    return 1;
                }

                int outputPins, inputPins;
                crossbar.get_PinCounts(out outputPins, out inputPins);

                for (int outPin = 0; outPin < outputPins; outPin++)
                {
                    int related;
                    PhysicalConnectorType outType;
                    crossbar.get_CrossbarPinInfo(false, outPin, out related, out outType);
                    if (outType == PhysicalConnectorType.Video_VideoDecoder)
                    {
                        int inputPinResult;
                        crossbar.get_IsRoutedTo(outPin, out inputPinResult);
                        int relatedPin;
                        PhysicalConnectorType pinType;
                        crossbar.get_CrossbarPinInfo(true, inputPinResult, out relatedPin, out pinType);
                        Console.WriteLine("Current input: Pin " + inputPinResult + " (" + GetPinTypeName(pinType) + ")");
                        return inputPinResult;
                    }
                }

                Console.WriteLine("Could not determine current input.");
                return -1;
            }
        }

        static List<string> EnumerateVideoDevices()
        {
            List<string> devices = new List<string>();

            ICreateDevEnum devEnum = (ICreateDevEnum)new CreateDevEnum();
            IEnumMoniker enumMoniker;
            Guid videoInputCategory = FilterCategory.VideoInputDevice;
            int hr = devEnum.CreateClassEnumerator(ref videoInputCategory, out enumMoniker, 0);

            if (hr != 0 || enumMoniker == null)
                return devices;

            IMoniker[] moniker = new IMoniker[1];
            while (enumMoniker.Next(1, moniker, IntPtr.Zero) == 0)
            {
                object bagObj;
                Guid propBagGuid = typeof(IPropertyBag).GUID;
                moniker[0].BindToStorage(null, null, ref propBagGuid, out bagObj);
                IPropertyBag bag = (IPropertyBag)bagObj;

                object nameObj;
                bag.Read("FriendlyName", out nameObj, null);
                devices.Add(nameObj != null ? nameObj.ToString() : "Unknown Device");

                Marshal.ReleaseComObject(bag);
                Marshal.ReleaseComObject(moniker[0]);
            }

            Marshal.ReleaseComObject(enumMoniker);
            Marshal.ReleaseComObject(devEnum);

            return devices;
        }

        static string GetPinTypeName(PhysicalConnectorType type)
        {
            switch (type)
            {
                case PhysicalConnectorType.Video_Tuner: return "Tuner";
                case PhysicalConnectorType.Video_Composite: return "Composite";
                case PhysicalConnectorType.Video_SVideo: return "S-Video";
                case PhysicalConnectorType.Video_RGB: return "RGB";
                case PhysicalConnectorType.Video_YRYBY: return "Component (YPbPr)";
                case PhysicalConnectorType.Video_SerialDigital: return "SDI";
                case PhysicalConnectorType.Video_ParallelDigital: return "Parallel Digital";
                case PhysicalConnectorType.Video_SCSI: return "SCSI";
                case PhysicalConnectorType.Video_AUX: return "AUX";
                case PhysicalConnectorType.Video_1394: return "FireWire";
                case PhysicalConnectorType.Video_USB: return "USB";
                case PhysicalConnectorType.Video_VideoDecoder: return "Video Decoder";
                case PhysicalConnectorType.Video_VideoEncoder: return "Video Encoder";
                case PhysicalConnectorType.Video_SCART: return "SCART";
                case PhysicalConnectorType.Video_Black: return "Black";
                case (PhysicalConnectorType)0x1000: return "HDMI";
                case (PhysicalConnectorType)0x1001: return "DVI";
                case (PhysicalConnectorType)0x1002: return "DisplayPort";
                default: return "Unknown (" + ((int)type).ToString("X") + ")";
            }
        }
    }

    // Filter graph helper class
    class FilterGraph : IDisposable
    {
        private IGraphBuilder _graph;
        private IBaseFilter _captureFilter;
        private ICaptureGraphBuilder2 _captureBuilder;
        private bool _disposed;

        public FilterGraph(string deviceNamePart)
        {
            _graph = (IGraphBuilder)new FilterGraphManager();
            _captureBuilder = (ICaptureGraphBuilder2)new CaptureGraphBuilder2();
            _captureBuilder.SetFiltergraph(_graph);

            _captureFilter = FindAndCreateCaptureFilter(deviceNamePart);
            if (_captureFilter == null)
                throw new Exception("Device containing '" + deviceNamePart + "' not found.");

            _graph.AddFilter(_captureFilter, "Capture");
        }

        public IAMCrossbar FindCrossbar()
        {
            if (_captureBuilder == null || _captureFilter == null)
                return null;

            object crossbarObj;
            Guid crossbarGuid = typeof(IAMCrossbar).GUID;
            Guid pinCat = PinCategory.Capture;
            Guid mediaType = MediaType.Video;

            // Try to find crossbar upstream of the capture filter
            int hr = _captureBuilder.FindInterface(
                ref pinCat,
                ref mediaType,
                _captureFilter,
                ref crossbarGuid,
                out crossbarObj);

            if (hr == 0 && crossbarObj != null)
                return (IAMCrossbar)crossbarObj;

            // Try without pin category
            Guid emptyGuid = Guid.Empty;
            hr = _captureBuilder.FindInterface(
                ref emptyGuid,
                ref mediaType,
                _captureFilter,
                ref crossbarGuid,
                out crossbarObj);

            if (hr == 0 && crossbarObj != null)
                return (IAMCrossbar)crossbarObj;

            // Try to find it by enumerating filters
            return FindCrossbarByEnumeration();
        }

        private IAMCrossbar FindCrossbarByEnumeration()
        {
            // Some devices need the graph to be built first
            if (_captureBuilder != null && _captureFilter != null)
            {
                Guid pinCat = PinCategory.Capture;
                Guid mediaType = MediaType.Video;
                _captureBuilder.RenderStream(ref pinCat, ref mediaType, _captureFilter, null, null);
            }

            if (_graph == null) return null;

            IEnumFilters enumFilters;
            _graph.EnumFilters(out enumFilters);
            if (enumFilters == null) return null;

            IBaseFilter[] filters = new IBaseFilter[1];
            while (enumFilters.Next(1, filters, IntPtr.Zero) == 0)
            {
                IAMCrossbar crossbar = filters[0] as IAMCrossbar;
                if (crossbar != null)
                {
                    Marshal.ReleaseComObject(enumFilters);
                    return crossbar;
                }
                Marshal.ReleaseComObject(filters[0]);
            }
            Marshal.ReleaseComObject(enumFilters);
            return null;
        }

        private static IBaseFilter FindAndCreateCaptureFilter(string deviceNamePart)
        {
            ICreateDevEnum devEnum = (ICreateDevEnum)new CreateDevEnum();
            IEnumMoniker enumMoniker;
            Guid videoInputCategory = FilterCategory.VideoInputDevice;
            devEnum.CreateClassEnumerator(ref videoInputCategory, out enumMoniker, 0);

            if (enumMoniker == null)
                return null;

            IMoniker[] moniker = new IMoniker[1];
            IBaseFilter result = null;

            while (enumMoniker.Next(1, moniker, IntPtr.Zero) == 0)
            {
                object bagObj;
                Guid propBagGuid = typeof(IPropertyBag).GUID;
                moniker[0].BindToStorage(null, null, ref propBagGuid, out bagObj);
                IPropertyBag bag = (IPropertyBag)bagObj;

                object nameObj;
                bag.Read("FriendlyName", out nameObj, null);
                string name = nameObj != null ? nameObj.ToString() : "";

                Marshal.ReleaseComObject(bag);

                if (name.IndexOf(deviceNamePart, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    object filterObj;
                    Guid baseFilterGuid = typeof(IBaseFilter).GUID;
                    moniker[0].BindToObject(null, null, ref baseFilterGuid, out filterObj);
                    result = (IBaseFilter)filterObj;
                    Marshal.ReleaseComObject(moniker[0]);
                    break;
                }

                Marshal.ReleaseComObject(moniker[0]);
            }

            Marshal.ReleaseComObject(enumMoniker);
            Marshal.ReleaseComObject(devEnum);

            return result;
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            if (_captureFilter != null)
            {
                Marshal.ReleaseComObject(_captureFilter);
                _captureFilter = null;
            }
            if (_captureBuilder != null)
            {
                Marshal.ReleaseComObject(_captureBuilder);
                _captureBuilder = null;
            }
            if (_graph != null)
            {
                Marshal.ReleaseComObject(_graph);
                _graph = null;
            }
        }
    }

    // COM Interop definitions

    [ComImport, Guid("29840822-5B84-11D0-BD3B-00A0C911CE86")]
    class CreateDevEnum { }

    [ComImport, Guid("E436EBB3-524F-11CE-9F53-0020AF0BA770")]
    class FilterGraphManager { }

    [ComImport, Guid("BF87B6E1-8C27-11D0-B3F0-00AA003761C5")]
    class CaptureGraphBuilder2 { }

    static class FilterCategory
    {
        public static readonly Guid VideoInputDevice = new Guid("860BB310-5D01-11D0-BD3B-00A0C911CE86");
    }

    static class PinCategory
    {
        public static readonly Guid Capture = new Guid("FB6C4281-0353-11D1-905F-0000C0CC16BA");
    }

    static class MediaType
    {
        public static readonly Guid Video = new Guid("73646976-0000-0010-8000-00AA00389B71");
    }

    enum PhysicalConnectorType
    {
        Video_Tuner = 1,
        Video_Composite = 2,
        Video_SVideo = 3,
        Video_RGB = 4,
        Video_YRYBY = 5,
        Video_SerialDigital = 6,
        Video_ParallelDigital = 7,
        Video_SCSI = 8,
        Video_AUX = 9,
        Video_1394 = 10,
        Video_USB = 11,
        Video_VideoDecoder = 12,
        Video_VideoEncoder = 13,
        Video_SCART = 14,
        Video_Black = 15,
    }

    [ComImport, Guid("29840822-5B84-11D0-BD3B-00A0C911CE86"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ICreateDevEnum
    {
        [PreserveSig]
        int CreateClassEnumerator([In] ref Guid type, out IEnumMoniker enumMoniker, [In] int flags);
    }

    [ComImport, Guid("00000102-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IEnumMoniker
    {
        [PreserveSig]
        int Next([In] int count, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] IMoniker[] monikers, IntPtr fetched);
        [PreserveSig]
        int Skip([In] int count);
        void Reset();
        void Clone(out IEnumMoniker enumMoniker);
    }

    [ComImport, Guid("0000000F-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMoniker
    {
        void QueryInterface_OnlyForPadding();
        void AddRef_OnlyForPadding();
        void Release_OnlyForPadding();
        void GetClassID(out Guid classID);
        void IsDirty();
        void Load(object stream);
        void Save(object stream, bool clearDirty);
        void GetSizeMax(out long size);
        void BindToObject(object bindCtx, IMoniker moniker, [In] ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object obj);
        void BindToStorage(object bindCtx, IMoniker moniker, [In] ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object obj);
    }

    [ComImport, Guid("55272A00-42CB-11CE-8135-00AA004BB851"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IPropertyBag
    {
        [PreserveSig]
        int Read([In, MarshalAs(UnmanagedType.LPWStr)] string propName, [Out, MarshalAs(UnmanagedType.Struct)] out object val, object errorLog);
        [PreserveSig]
        int Write([In, MarshalAs(UnmanagedType.LPWStr)] string propName, [In] ref object val);
    }

    [ComImport, Guid("56A86895-0AD4-11CE-B03A-0020AF0BA770"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IBaseFilter
    {
        void GetClassID(out Guid classID);
        void Stop();
        void Pause();
        void Run(long start);
        void GetState(int timeout, out int state);
        void SetSyncSource(IntPtr clock);
        void GetSyncSource(out IntPtr clock);
        void EnumPins(out IntPtr enumPins);
        void FindPin(string id, out IntPtr pin);
        void QueryFilterInfo(out IntPtr info);
        void JoinFilterGraph(IntPtr graph, string name);
        void QueryVendorInfo(out IntPtr vendorInfo);
    }

    [ComImport, Guid("56A868A9-0AD4-11CE-B03A-0020AF0BA770"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IGraphBuilder
    {
        void AddFilter([In] IBaseFilter filter, [In, MarshalAs(UnmanagedType.LPWStr)] string name);
        void RemoveFilter([In] IBaseFilter filter);
        void EnumFilters(out IEnumFilters enumFilters);
        void FindFilterByName([In, MarshalAs(UnmanagedType.LPWStr)] string name, out IBaseFilter filter);
        void ConnectDirect(IntPtr pinOut, IntPtr pinIn, IntPtr mediaType);
        void Reconnect(IntPtr pin);
        void Disconnect(IntPtr pin);
        void SetDefaultSyncSource();
        void Connect(IntPtr pinOut, IntPtr pinIn);
        void Render(IntPtr pinOut);
        void RenderFile([In, MarshalAs(UnmanagedType.LPWStr)] string file, [In, MarshalAs(UnmanagedType.LPWStr)] string playList);
        void AddSourceFilter([In, MarshalAs(UnmanagedType.LPWStr)] string fileName, [In, MarshalAs(UnmanagedType.LPWStr)] string filterName, out IBaseFilter filter);
        void SetLogFile(IntPtr file);
        void Abort();
        void ShouldOperationContinue();
    }

    [ComImport, Guid("56A86893-0AD4-11CE-B03A-0020AF0BA770"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IEnumFilters
    {
        [PreserveSig]
        int Next([In] int count, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] IBaseFilter[] filters, IntPtr fetched);
        [PreserveSig]
        int Skip([In] int count);
        void Reset();
        void Clone(out IEnumFilters enumFilters);
    }

    [ComImport, Guid("93E5A4E0-2D50-11D2-ABFA-00A0C9C6E38D"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ICaptureGraphBuilder2
    {
        void SetFiltergraph([In] IGraphBuilder graph);
        void GetFiltergraph(out IGraphBuilder graph);
        void SetOutputFileName(ref Guid type, [In, MarshalAs(UnmanagedType.LPWStr)] string fileName, out IBaseFilter filter, out IntPtr sink);

        [PreserveSig]
        int FindInterface([In] ref Guid category, [In] ref Guid type, [In] IBaseFilter filter, [In] ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object obj);

        [PreserveSig]
        int RenderStream([In] ref Guid category, [In] ref Guid type, [In, MarshalAs(UnmanagedType.IUnknown)] object source, [In] IBaseFilter compressor, [In] IBaseFilter renderer);

        void ControlStream(ref Guid category, ref Guid type, IBaseFilter filter, long start, long stop, short sendExtra, out short dropped);
        void AllocCapFile([In, MarshalAs(UnmanagedType.LPWStr)] string fileName, long size);
        void CopyCaptureFile([In, MarshalAs(UnmanagedType.LPWStr)] string oldFile, [In, MarshalAs(UnmanagedType.LPWStr)] string newFile, int allowEscAbort, IntPtr callback);
        void FindPin(object source, int direction, ref Guid category, ref Guid type, bool unconnected, int index, out IntPtr pin);
    }

    [ComImport, Guid("C6E13380-30AC-11D0-A18C-00A0C9118956"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IAMCrossbar
    {
        [PreserveSig]
        int get_PinCounts(out int outputPins, out int inputPins);

        [PreserveSig]
        int CanRoute([In] int outputPin, [In] int inputPin);

        [PreserveSig]
        int Route([In] int outputPin, [In] int inputPin);

        [PreserveSig]
        int get_IsRoutedTo([In] int outputPin, out int inputPin);

        [PreserveSig]
        int get_CrossbarPinInfo([In] bool isInput, [In] int pinIndex, out int relatedPinIndex, out PhysicalConnectorType type);
    }
}
