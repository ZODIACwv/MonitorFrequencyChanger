using System.Runtime.InteropServices;

class Program
{
    #region P/Invoke

    [DllImport("user32.dll")]
    static extern int GetDisplayConfigBufferSizes(uint flags, out uint numPathArrayElements, out uint numModeInfoArrayElements);

    [DllImport("user32.dll")]
    static extern int QueryDisplayConfig(uint flags, ref uint numPathArrayElements,
        [Out] DISPLAYCONFIG_PATH_INFO[] pathArray, ref uint numModeInfoArrayElements,
        [Out] DISPLAYCONFIG_MODE_INFO[] modeInfoArray, IntPtr currentTopologyId);

    [DllImport("user32.dll")]
    static extern int SetDisplayConfig(uint numPathArrayElements,
        [In] DISPLAYCONFIG_PATH_INFO[] pathArray, uint numModeInfoArrayElements,
        [In] DISPLAYCONFIG_MODE_INFO[] modeInfoArray, uint flags);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    static extern int DisplayConfigGetDeviceInfo(ref DISPLAYCONFIG_TARGET_DEVICE_NAME requestPacket);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    static extern int DisplayConfigGetDeviceInfo(ref DISPLAYCONFIG_SOURCE_DEVICE_NAME requestPacket);

    const uint QDC_ONLY_ACTIVE_PATHS = 0x00000002;
    const uint SDC_APPLY = 0x00000080;
    const uint SDC_USE_SUPPLIED_DISPLAY_CONFIG = 0x00000020;
    const uint SDC_ALLOW_CHANGES = 0x00000400;
    const uint SDC_SAVE_TO_DATABASE = 0x00000200;

    const uint DISPLAYCONFIG_DEVICE_INFO_GET_TARGET_NAME = 2;
    const uint DISPLAYCONFIG_DEVICE_INFO_GET_SOURCE_NAME = 1;

    #endregion

    #region Structs

    [StructLayout(LayoutKind.Sequential)]
    struct LUID
    {
        public uint LowPart;
        public int HighPart;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct DISPLAYCONFIG_RATIONAL
    {
        public uint Numerator;
        public uint Denominator;
        public double ToDouble() => Denominator == 0 ? 0 : (double)Numerator / Denominator;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct DISPLAYCONFIG_PATH_SOURCE_INFO
    {
        public LUID adapterId;
        public uint id;
        public uint modeInfoIdx;
        public uint statusFlags;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct DISPLAYCONFIG_PATH_TARGET_INFO
    {
        public LUID adapterId;
        public uint id;
        public uint modeInfoIdx;
        public uint outputTechnology;
        public uint rotation;
        public uint scaling;
        public DISPLAYCONFIG_RATIONAL refreshRate;
        public uint scanLineOrdering;
        public int targetAvailable;
        public uint statusFlags;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct DISPLAYCONFIG_PATH_INFO
    {
        public DISPLAYCONFIG_PATH_SOURCE_INFO sourceInfo;
        public DISPLAYCONFIG_PATH_TARGET_INFO targetInfo;
        public uint flags;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct DISPLAYCONFIG_2DREGION
    {
        public uint cx;
        public uint cy;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct DISPLAYCONFIG_VIDEO_SIGNAL_INFO
    {
        public ulong pixelRate;
        public DISPLAYCONFIG_RATIONAL hSyncFreq;
        public DISPLAYCONFIG_RATIONAL vSyncFreq;
        public DISPLAYCONFIG_2DREGION activeSize;
        public DISPLAYCONFIG_2DREGION totalSize;
        public uint videoStandard;
        public uint scanLineOrdering;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct DISPLAYCONFIG_TARGET_MODE
    {
        public DISPLAYCONFIG_VIDEO_SIGNAL_INFO targetVideoSignalInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct POINTL
    {
        public int x;
        public int y;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct DISPLAYCONFIG_SOURCE_MODE
    {
        public uint width;
        public uint height;
        public uint pixelFormat;
        public POINTL position;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct DISPLAYCONFIG_DESKTOP_IMAGE_INFO
    {
        public POINTL PathSourceSize;
        public RECT DesktopImageRegion;
        public RECT DesktopImageClip;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct RECT
    {
        public int left, top, right, bottom;
    }

    [StructLayout(LayoutKind.Explicit)]
    struct DISPLAYCONFIG_MODE_INFO_UNION
    {
        [FieldOffset(0)]
        public DISPLAYCONFIG_TARGET_MODE targetMode;
        [FieldOffset(0)]
        public DISPLAYCONFIG_SOURCE_MODE sourceMode;
        [FieldOffset(0)]
        public DISPLAYCONFIG_DESKTOP_IMAGE_INFO desktopImageInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct DISPLAYCONFIG_MODE_INFO
    {
        public uint infoType;
        public uint id;
        public LUID adapterId;
        public DISPLAYCONFIG_MODE_INFO_UNION info;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct DISPLAYCONFIG_DEVICE_INFO_HEADER
    {
        public uint type;
        public uint size;
        public LUID adapterId;
        public uint id;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct DISPLAYCONFIG_TARGET_DEVICE_NAME
    {
        public DISPLAYCONFIG_DEVICE_INFO_HEADER header;
        public uint flags;
        public uint outputTechnology;
        public ushort edidManufactureId;
        public ushort edidProductCodeId;
        public uint connectorInstance;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
        public string monitorFriendlyDeviceName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string monitorDevicePath;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct DISPLAYCONFIG_SOURCE_DEVICE_NAME
    {
        public DISPLAYCONFIG_DEVICE_INFO_HEADER header;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string viewGdiDeviceName;
    }

    #endregion

    static int Main(string[] args)
    {
        if (args.Length == 0)
            return ListMonitors();

        var cmd = args[0].ToLower();

        return cmd switch
        {
            "list" => ListMonitors(),
            "set" => SetRefreshRate(args),
            _ => ShowHelp()
        };
    }

    static int ShowHelp()
    {
        Console.WriteLine("MonitorFrequencyChanger - set exact monitor refresh rate via Numerator/Denominator");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  MonitorFrequencyChanger                         - list all monitors");
        Console.WriteLine("  MonitorFrequencyChanger list                    - list all monitors");
        Console.WriteLine("  MonitorFrequencyChanger set <index> <hz>        - set refresh rate by index");
        Console.WriteLine("  MonitorFrequencyChanger set <name> <hz>         - set refresh rate by name");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine("  MonitorFrequencyChanger set 0 59.89             - set monitor 0 to 59.89 Hz");
        Console.WriteLine("  MonitorFrequencyChanger set \"ASUS VW193D\" 59.89 - set by name");
        Console.WriteLine("  MonitorFrequencyChanger set 1 144              - set monitor 1 to 144 Hz");
        Console.WriteLine();
        Console.WriteLine("Note: Uses Windows CCD API with DISPLAYCONFIG_RATIONAL for precise fractional rates.");
        Console.WriteLine("      This bypasses Windows rounding that occurs with standard display settings.");
        return 0;
    }

    static (DISPLAYCONFIG_PATH_INFO[] paths, DISPLAYCONFIG_MODE_INFO[] modes, uint pathCount, uint modeCount)? GetDisplayConfig()
    {
        int err = GetDisplayConfigBufferSizes(QDC_ONLY_ACTIVE_PATHS, out uint pathCount, out uint modeCount);
        if (err != 0)
        {
            Console.WriteLine($"Error GetDisplayConfigBufferSizes: {err}");
            return null;
        }

        var paths = new DISPLAYCONFIG_PATH_INFO[pathCount];
        var modes = new DISPLAYCONFIG_MODE_INFO[modeCount];

        err = QueryDisplayConfig(QDC_ONLY_ACTIVE_PATHS, ref pathCount, paths, ref modeCount, modes, IntPtr.Zero);
        if (err != 0)
        {
            Console.WriteLine($"Error QueryDisplayConfig: {err}");
            return null;
        }

        return (paths, modes, pathCount, modeCount);
    }

    static string GetMonitorName(ref DISPLAYCONFIG_PATH_INFO path)
    {
        var targetName = new DISPLAYCONFIG_TARGET_DEVICE_NAME();
        targetName.header.type = DISPLAYCONFIG_DEVICE_INFO_GET_TARGET_NAME;
        targetName.header.size = (uint)Marshal.SizeOf<DISPLAYCONFIG_TARGET_DEVICE_NAME>();
        targetName.header.adapterId = path.targetInfo.adapterId;
        targetName.header.id = path.targetInfo.id;
        DisplayConfigGetDeviceInfo(ref targetName);
        return targetName.monitorFriendlyDeviceName ?? "";
    }

    static string GetSourceName(ref DISPLAYCONFIG_PATH_INFO path)
    {
        var sourceName = new DISPLAYCONFIG_SOURCE_DEVICE_NAME();
        sourceName.header.type = DISPLAYCONFIG_DEVICE_INFO_GET_SOURCE_NAME;
        sourceName.header.size = (uint)Marshal.SizeOf<DISPLAYCONFIG_SOURCE_DEVICE_NAME>();
        sourceName.header.adapterId = path.sourceInfo.adapterId;
        sourceName.header.id = path.sourceInfo.id;
        DisplayConfigGetDeviceInfo(ref sourceName);
        return sourceName.viewGdiDeviceName ?? "";
    }

    static int ListMonitors()
    {
        var config = GetDisplayConfig();
        if (config == null) return 1;

        var (paths, modes, pathCount, modeCount) = config.Value;

        Console.WriteLine("Active monitors:");
        Console.WriteLine();

        for (int i = 0; i < pathCount; i++)
        {
            ref var path = ref paths[i];

            var monitorName = GetMonitorName(ref path);
            var sourceName = GetSourceName(ref path);

            uint width = 0, height = 0;
            double vSync = 0;

            var sourceModeIdx = path.sourceInfo.modeInfoIdx & 0xFFFF;
            if (sourceModeIdx < modeCount && modes[sourceModeIdx].infoType == 1)
            {
                width = modes[sourceModeIdx].info.sourceMode.width;
                height = modes[sourceModeIdx].info.sourceMode.height;
            }

            var targetModeIdx = path.targetInfo.modeInfoIdx & 0xFFFF;
            if (targetModeIdx < modeCount && modes[targetModeIdx].infoType == 2)
            {
                vSync = modes[targetModeIdx].info.targetMode.targetVideoSignalInfo.vSyncFreq.ToDouble();
                if (width == 0)
                {
                    width = modes[targetModeIdx].info.targetMode.targetVideoSignalInfo.activeSize.cx;
                    height = modes[targetModeIdx].info.targetMode.targetVideoSignalInfo.activeSize.cy;
                }
            }

            var refreshRate = path.targetInfo.refreshRate;

            Console.WriteLine($"[{i}] {sourceName}");
            Console.WriteLine($"    Name: {(string.IsNullOrEmpty(monitorName) ? "(unknown)" : monitorName)}");
            Console.WriteLine($"    Resolution: {width}x{height}");
            Console.WriteLine($"    Refresh: {refreshRate.ToDouble():F2} Hz ({refreshRate.Numerator}/{refreshRate.Denominator})");
            Console.WriteLine();
        }

        return 0;
    }

    static int SetRefreshRate(string[] args)
    {
        if (args.Length < 3)
        {
            Console.WriteLine("Usage: MonitorFrequencyChanger set <index|name> <hz>");
            Console.WriteLine("Example: MonitorFrequencyChanger set 0 59.89");
            return 1;
        }

        var target = args[1];
        if (!double.TryParse(args[2], System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out double hz))
        {
            // Try with current culture
            if (!double.TryParse(args[2], out hz))
            {
                Console.WriteLine($"Invalid refresh rate: {args[2]}");
                return 1;
            }
        }

        if (hz <= 0 || hz > 500)
        {
            Console.WriteLine($"Invalid refresh rate: {hz} Hz (must be 0-500)");
            return 1;
        }

        var config = GetDisplayConfig();
        if (config == null) return 1;

        var (paths, modes, pathCount, modeCount) = config.Value;

        // Find target monitor
        int targetIndex = -1;

        // Try parse as index first
        if (int.TryParse(target, out int idx) && idx >= 0 && idx < pathCount)
        {
            targetIndex = idx;
        }
        else
        {
            // Search by name
            for (int i = 0; i < pathCount; i++)
            {
                var name = GetMonitorName(ref paths[i]);
                if (!string.IsNullOrEmpty(name) && name.Contains(target, StringComparison.OrdinalIgnoreCase))
                {
                    targetIndex = i;
                    break;
                }
            }
        }

        if (targetIndex == -1)
        {
            Console.WriteLine($"Monitor not found: {target}");
            Console.WriteLine("Use 'MonitorFrequencyChanger list' to see available monitors.");
            return 1;
        }

        // Calculate numerator/denominator for the desired refresh rate
        // Use precision of 2 decimal places (multiply by 100)
        uint numerator = (uint)Math.Round(hz * 100);
        uint denominator = 100;

        // Simplify fraction if possible
        uint gcd = GCD(numerator, denominator);
        numerator /= gcd;
        denominator /= gcd;

        var monitorName = GetMonitorName(ref paths[targetIndex]);
        Console.WriteLine($"Target: [{targetIndex}] {(string.IsNullOrEmpty(monitorName) ? "(unknown)" : monitorName)}");
        Console.WriteLine($"Setting refresh rate: {hz:F2} Hz ({numerator}/{denominator})");

        // Set refresh rate in path
        paths[targetIndex].targetInfo.refreshRate.Numerator = numerator;
        paths[targetIndex].targetInfo.refreshRate.Denominator = denominator;

        // Update vSyncFreq in mode info
        var targetModeIdx = paths[targetIndex].targetInfo.modeInfoIdx & 0xFFFF;
        if (targetModeIdx < modeCount && modes[targetModeIdx].infoType == 2)
        {
            modes[targetModeIdx].info.targetMode.targetVideoSignalInfo.vSyncFreq.Numerator = numerator;
            modes[targetModeIdx].info.targetMode.targetVideoSignalInfo.vSyncFreq.Denominator = denominator;
        }

        int result = SetDisplayConfig(pathCount, paths, modeCount, modes,
            SDC_APPLY | SDC_USE_SUPPLIED_DISPLAY_CONFIG | SDC_ALLOW_CHANGES | SDC_SAVE_TO_DATABASE);

        if (result != 0)
        {
            Console.WriteLine($"Error SetDisplayConfig: {result} (0x{result:X8})");
            return 1;
        }

        Console.WriteLine("Success!");
        return 0;
    }

    static uint GCD(uint a, uint b)
    {
        while (b != 0)
        {
            uint t = b;
            b = a % b;
            a = t;
        }
        return a;
    }
}
