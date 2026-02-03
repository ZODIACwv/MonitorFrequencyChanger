# MonitorFrequencyChanger

CLI tool to set **exact** monitor refresh rates on Windows using the CCD API (Connected Display Configuration).

## Problem

Windows rounds refresh rates when setting them through standard display settings. For example, if you need exactly 59.89 Hz (common for VGA monitors), Windows will round it to 60 Hz or 59.68 Hz, which can cause display issues.

## Solution

This tool uses Windows `SetDisplayConfig` API with `DISPLAYCONFIG_RATIONAL` structure, which allows specifying refresh rate as a fraction (Numerator/Denominator). This bypasses Windows rounding.

## Requirements

- Windows 7 or later
- .NET 9.0

## Build

```bash
dotnet build
```

## Usage

```bash
# List all monitors
MonitorFrequencyChanger

# or
MonitorFrequencyChanger list

# Set refresh rate by monitor index
MonitorFrequencyChanger set 0 59.89

# Set refresh rate by monitor name (partial match)
MonitorFrequencyChanger set "ASUS VW193D" 59.89

# Set to 144 Hz
MonitorFrequencyChanger set 1 144
```

## Example Output

```
Active monitors:

[0] \\.\DISPLAY1
    Name: (unknown)
    Resolution: 1920x1080
    Refresh: 60.05 Hz (142000000/2364600)

[1] \\.\DISPLAY2
    Name: ASUS VW193D
    Resolution: 1440x900
    Refresh: 59.89 Hz (5989/100)
```

## How It Works

1. Queries current display configuration via `QueryDisplayConfig`
2. Finds target monitor by index or name
3. Sets refresh rate using Numerator/Denominator (e.g., 59.89 Hz = 5989/100)
4. Applies configuration via `SetDisplayConfig`

## License

MIT

## TESTED
Tested on Windows 11, ASUS VW193D, frequency successfully setted up to 59.89 Hz via this utility

## Contributors
General developer: Claude Code
Code was fully GENERATED, I only did the review, keep in mind, possible it could require micro structs, enums or functional improvements
