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
        public uint modeInfoIdx;  // union with cloneGroupId
        public uint statusFlags;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct DISPLAYCONFIG_PATH_TARGET_INFO
    {
        public LUID adapterId;
        public uint id;
        public uint modeInfoIdx;  // union with desktopModeInfoIdx
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
        public uint videoStandard;  // union
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

    // MODE_INFO union - use explicit layout
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
        var cmd = args.Length > 0 ? args[0].ToLower() : "list";

        return cmd switch
        {
            "list" => ListMonitors(),
            "set" => SetMonitorRefreshRate(),
            "debug" => DebugStructSizes(),
            _ => ShowHelp()
        };
    }

    static int DebugStructSizes()
    {
        Console.WriteLine($"DISPLAYCONFIG_PATH_INFO: {Marshal.SizeOf<DISPLAYCONFIG_PATH_INFO>()} bytes");
        Console.WriteLine($"DISPLAYCONFIG_MODE_INFO: {Marshal.SizeOf<DISPLAYCONFIG_MODE_INFO>()} bytes");
        Console.WriteLine($"DISPLAYCONFIG_TARGET_DEVICE_NAME: {Marshal.SizeOf<DISPLAYCONFIG_TARGET_DEVICE_NAME>()} bytes");
        Console.WriteLine($"DISPLAYCONFIG_SOURCE_DEVICE_NAME: {Marshal.SizeOf<DISPLAYCONFIG_SOURCE_DEVICE_NAME>()} bytes");
        Console.WriteLine($"DISPLAYCONFIG_VIDEO_SIGNAL_INFO: {Marshal.SizeOf<DISPLAYCONFIG_VIDEO_SIGNAL_INFO>()} bytes");
        Console.WriteLine($"DISPLAYCONFIG_TARGET_MODE: {Marshal.SizeOf<DISPLAYCONFIG_TARGET_MODE>()} bytes");
        Console.WriteLine($"DISPLAYCONFIG_SOURCE_MODE: {Marshal.SizeOf<DISPLAYCONFIG_SOURCE_MODE>()} bytes");
        Console.WriteLine($"DISPLAYCONFIG_DESKTOP_IMAGE_INFO: {Marshal.SizeOf<DISPLAYCONFIG_DESKTOP_IMAGE_INFO>()} bytes");
        Console.WriteLine($"DISPLAYCONFIG_MODE_INFO_UNION: {Marshal.SizeOf<DISPLAYCONFIG_MODE_INFO_UNION>()} bytes");
        return 0;
    }

    static int ShowHelp()
    {
        Console.WriteLine("MonitorFrequencyChanger - установка частоты 59.89Hz для ASUS VW193D");
        Console.WriteLine();
        Console.WriteLine("Команды:");
        Console.WriteLine("  list  - показать мониторы и их частоты (по умолчанию)");
        Console.WriteLine("  set   - установить 1440x900 @ 59.89Hz на ASUS VW193D");
        return 0;
    }

    static int ListMonitors()
    {
        int err = GetDisplayConfigBufferSizes(QDC_ONLY_ACTIVE_PATHS, out uint pathCount, out uint modeCount);
        if (err != 0)
        {
            Console.WriteLine($"Ошибка GetDisplayConfigBufferSizes: {err}");
            return 1;
        }

        Console.WriteLine($"pathCount={pathCount}, modeCount={modeCount}");

        var paths = new DISPLAYCONFIG_PATH_INFO[pathCount];
        var modes = new DISPLAYCONFIG_MODE_INFO[modeCount];

        err = QueryDisplayConfig(QDC_ONLY_ACTIVE_PATHS, ref pathCount, paths, ref modeCount, modes, IntPtr.Zero);
        if (err != 0)
        {
            Console.WriteLine($"Ошибка QueryDisplayConfig: {err}");
            return 1;
        }

        Console.WriteLine($"После query: pathCount={pathCount}, modeCount={modeCount}");
        Console.WriteLine();
        Console.WriteLine("Активные мониторы:");
        Console.WriteLine();

        for (int i = 0; i < pathCount; i++)
        {
            ref var path = ref paths[i];

            Console.WriteLine($"[{i}] path.sourceInfo.adapterId = {path.sourceInfo.adapterId.LowPart}:{path.sourceInfo.adapterId.HighPart}");
            Console.WriteLine($"    path.sourceInfo.id = {path.sourceInfo.id}");
            Console.WriteLine($"    path.sourceInfo.modeInfoIdx = {path.sourceInfo.modeInfoIdx}");
            Console.WriteLine($"    path.targetInfo.adapterId = {path.targetInfo.adapterId.LowPart}:{path.targetInfo.adapterId.HighPart}");
            Console.WriteLine($"    path.targetInfo.id = {path.targetInfo.id}");
            Console.WriteLine($"    path.targetInfo.modeInfoIdx = {path.targetInfo.modeInfoIdx}");
            Console.WriteLine($"    path.targetInfo.refreshRate = {path.targetInfo.refreshRate.Numerator}/{path.targetInfo.refreshRate.Denominator} = {path.targetInfo.refreshRate.ToDouble():F2} Hz");

            // Получаем имя монитора
            var targetName = new DISPLAYCONFIG_TARGET_DEVICE_NAME();
            targetName.header.type = DISPLAYCONFIG_DEVICE_INFO_GET_TARGET_NAME;
            targetName.header.size = (uint)Marshal.SizeOf<DISPLAYCONFIG_TARGET_DEVICE_NAME>();
            targetName.header.adapterId = path.targetInfo.adapterId;
            targetName.header.id = path.targetInfo.id;
            int r = DisplayConfigGetDeviceInfo(ref targetName);
            Console.WriteLine($"    DisplayConfigGetDeviceInfo(target) = {r}");
            Console.WriteLine($"    Монитор: '{targetName.monitorFriendlyDeviceName}'");

            // Получаем имя источника
            var sourceName = new DISPLAYCONFIG_SOURCE_DEVICE_NAME();
            sourceName.header.type = DISPLAYCONFIG_DEVICE_INFO_GET_SOURCE_NAME;
            sourceName.header.size = (uint)Marshal.SizeOf<DISPLAYCONFIG_SOURCE_DEVICE_NAME>();
            sourceName.header.adapterId = path.sourceInfo.adapterId;
            sourceName.header.id = path.sourceInfo.id;
            r = DisplayConfigGetDeviceInfo(ref sourceName);
            Console.WriteLine($"    DisplayConfigGetDeviceInfo(source) = {r}");
            Console.WriteLine($"    GDI Name: '{sourceName.viewGdiDeviceName}'");

            // Получаем режим для разрешения
            uint width = 0, height = 0;
            double vSync = 0;

            var sourceModeIdx = path.sourceInfo.modeInfoIdx & 0xFFFF; // lower 16 bits
            if (sourceModeIdx < modeCount)
            {
                ref var modeInfo = ref modes[sourceModeIdx];
                Console.WriteLine($"    sourceMode[{sourceModeIdx}].infoType = {modeInfo.infoType}");
                if (modeInfo.infoType == 1) // SOURCE
                {
                    width = modeInfo.info.sourceMode.width;
                    height = modeInfo.info.sourceMode.height;
                }
            }

            var targetModeIdx = path.targetInfo.modeInfoIdx & 0xFFFF;
            if (targetModeIdx < modeCount)
            {
                ref var modeInfo = ref modes[targetModeIdx];
                Console.WriteLine($"    targetMode[{targetModeIdx}].infoType = {modeInfo.infoType}");
                if (modeInfo.infoType == 2) // TARGET
                {
                    vSync = modeInfo.info.targetMode.targetVideoSignalInfo.vSyncFreq.ToDouble();
                    if (width == 0)
                    {
                        width = modeInfo.info.targetMode.targetVideoSignalInfo.activeSize.cx;
                        height = modeInfo.info.targetMode.targetVideoSignalInfo.activeSize.cy;
                    }
                }
            }

            Console.WriteLine($"    Разрешение: {width}x{height}");
            Console.WriteLine($"    vSync: {vSync:F2} Hz");
            Console.WriteLine();
        }

        return 0;
    }

    static int SetMonitorRefreshRate()
    {
        const string TARGET_MONITOR = "ASUS VW193D";
        const uint TARGET_WIDTH = 1440;
        const uint TARGET_HEIGHT = 900;
        const uint TARGET_REFRESH_NUM = 5989;
        const uint TARGET_REFRESH_DEN = 100;

        if (GetDisplayConfigBufferSizes(QDC_ONLY_ACTIVE_PATHS, out uint pathCount, out uint modeCount) != 0)
        {
            Console.WriteLine("Ошибка: не удалось получить размеры буферов");
            return 1;
        }

        var paths = new DISPLAYCONFIG_PATH_INFO[pathCount];
        var modes = new DISPLAYCONFIG_MODE_INFO[modeCount];

        if (QueryDisplayConfig(QDC_ONLY_ACTIVE_PATHS, ref pathCount, paths, ref modeCount, modes, IntPtr.Zero) != 0)
        {
            Console.WriteLine("Ошибка: не удалось получить конфигурацию дисплеев");
            return 1;
        }

        int targetIndex = -1;

        for (int i = 0; i < pathCount; i++)
        {
            ref var path = ref paths[i];

            var targetName = new DISPLAYCONFIG_TARGET_DEVICE_NAME();
            targetName.header.type = DISPLAYCONFIG_DEVICE_INFO_GET_TARGET_NAME;
            targetName.header.size = (uint)Marshal.SizeOf<DISPLAYCONFIG_TARGET_DEVICE_NAME>();
            targetName.header.adapterId = path.targetInfo.adapterId;
            targetName.header.id = path.targetInfo.id;
            DisplayConfigGetDeviceInfo(ref targetName);

            Console.WriteLine($"[{i}] {targetName.monitorFriendlyDeviceName}");

            if (!string.IsNullOrEmpty(targetName.monitorFriendlyDeviceName) &&
                targetName.monitorFriendlyDeviceName.Contains(TARGET_MONITOR, StringComparison.OrdinalIgnoreCase))
            {
                targetIndex = i;
                Console.WriteLine($"    ^ НАЙДЕН целевой монитор!");
                break;
            }
        }

        if (targetIndex == -1)
        {
            Console.WriteLine($"Монитор '{TARGET_MONITOR}' не найден по имени.");
            Console.WriteLine("Использую монитор с индексом 1 (второй монитор)...");

            if (pathCount > 1)
                targetIndex = 1;
            else
            {
                Console.WriteLine("Ошибка: второй монитор не обнаружен");
                return 1;
            }
        }

        Console.WriteLine($"Целевой монитор: индекс {targetIndex}");

        // Устанавливаем частоту в path
        paths[targetIndex].targetInfo.refreshRate.Numerator = TARGET_REFRESH_NUM;
        paths[targetIndex].targetInfo.refreshRate.Denominator = TARGET_REFRESH_DEN;

        // Обновляем разрешение и частоту в mode info
        var sourceModeIdx = paths[targetIndex].sourceInfo.modeInfoIdx & 0xFFFF;
        if (sourceModeIdx < modeCount && modes[sourceModeIdx].infoType == 1)
        {
            modes[sourceModeIdx].info.sourceMode.width = TARGET_WIDTH;
            modes[sourceModeIdx].info.sourceMode.height = TARGET_HEIGHT;
        }

        var targetModeIdx = paths[targetIndex].targetInfo.modeInfoIdx & 0xFFFF;
        if (targetModeIdx < modeCount && modes[targetModeIdx].infoType == 2)
        {
            modes[targetModeIdx].info.targetMode.targetVideoSignalInfo.vSyncFreq.Numerator = TARGET_REFRESH_NUM;
            modes[targetModeIdx].info.targetMode.targetVideoSignalInfo.vSyncFreq.Denominator = TARGET_REFRESH_DEN;
            modes[targetModeIdx].info.targetMode.targetVideoSignalInfo.activeSize.cx = TARGET_WIDTH;
            modes[targetModeIdx].info.targetMode.targetVideoSignalInfo.activeSize.cy = TARGET_HEIGHT;
        }

        Console.WriteLine($"Устанавливаю {TARGET_WIDTH}x{TARGET_HEIGHT} @ {(double)TARGET_REFRESH_NUM / TARGET_REFRESH_DEN:F2} Hz...");

        int result = SetDisplayConfig(pathCount, paths, modeCount, modes,
            SDC_APPLY | SDC_USE_SUPPLIED_DISPLAY_CONFIG | SDC_ALLOW_CHANGES | SDC_SAVE_TO_DATABASE);

        if (result != 0)
        {
            Console.WriteLine($"Ошибка SetDisplayConfig: {result} (0x{result:X8})");
            return 1;
        }

        Console.WriteLine("Успешно!");
        return 0;
    }
}
