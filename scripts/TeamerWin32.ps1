<#
.SYNOPSIS
    Teamer Win32 API Definitions - Shared module for window manipulation.

.DESCRIPTION
    Contains all Win32 API definitions used by Teamer for:
    - Window enumeration and properties
    - Window positioning and resizing
    - DWM (Desktop Window Manager) APIs
    - Monitor detection

.NOTES
    This module is automatically loaded by Manage-Teamer.ps1 and Manage-TeamerEnvironment.ps1
    Only loads types once to avoid redefinition errors.
#>

#region Win32 Type Definitions

# Check if types are already loaded to avoid redefinition errors
if (-not ([System.Management.Automation.PSTypeName]'TeamerWin32').Type) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections.Generic;

public class TeamerWin32 {
    // Window Rectangle Functions
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool IsZoomed(IntPtr hWnd);

    // Process and Thread Functions
    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    // Window Text Functions
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    // Foreground Window
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    // Window Enumeration
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    // Window Class
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    // Window Style
    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    // Monitor Functions
    [DllImport("user32.dll")]
    public static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFOEX lpmi);

    // Window Messaging
    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    // Window Positioning
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    // DWM API for getting extended frame bounds (visible window area without invisible borders)
    [DllImport("dwmapi.dll")]
    public static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out RECT pvAttribute, int cbAttribute);

    // Delegate for EnumWindows callback
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    // Constants
    public const int GWL_STYLE = -16;
    public const int GWL_EXSTYLE = -20;
    public const uint WS_VISIBLE = 0x10000000;
    public const uint WS_CAPTION = 0x00C00000;
    public const uint WS_EX_TOOLWINDOW = 0x00000080;
    public const uint WS_EX_APPWINDOW = 0x00040000;
    public const uint MONITOR_DEFAULTTONEAREST = 2;
    public const uint WM_CLOSE = 0x0010;
    public const uint SWP_NOACTIVATE = 0x0010;
    public const uint SWP_NOZORDER = 0x0004;

    // DWMWA_EXTENDED_FRAME_BOUNDS = 9 - gets the visible window bounds excluding invisible borders
    public const int DWMWA_EXTENDED_FRAME_BOUNDS = 9;

    // Structures
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct MONITORINFOEX {
        public int cbSize;
        public RECT rcMonitor;
        public RECT rcWork;
        public uint dwFlags;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string szDevice;
    }

    // Helper Methods
    public static string GetWindowTitle(IntPtr hWnd) {
        int length = GetWindowTextLength(hWnd);
        if (length == 0) return "";
        StringBuilder sb = new StringBuilder(length + 1);
        GetWindowText(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    public static string GetWindowClassName(IntPtr hWnd) {
        StringBuilder sb = new StringBuilder(256);
        GetClassName(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    public static List<IntPtr> GetAllWindows() {
        List<IntPtr> windows = new List<IntPtr>();
        EnumWindows((hWnd, lParam) => {
            windows.Add(hWnd);
            return true;
        }, IntPtr.Zero);
        return windows;
    }

    public static bool IsAltTabWindow(IntPtr hWnd) {
        if (!IsWindowVisible(hWnd)) return false;

        int style = GetWindowLong(hWnd, GWL_STYLE);
        int exStyle = GetWindowLong(hWnd, GWL_EXSTYLE);

        if ((exStyle & WS_EX_TOOLWINDOW) != 0 && (exStyle & WS_EX_APPWINDOW) == 0)
            return false;

        if ((style & WS_CAPTION) != WS_CAPTION)
            return false;

        if (GetWindowTextLength(hWnd) == 0)
            return false;

        return true;
    }

    public static string GetMonitorDeviceName(IntPtr hWnd) {
        IntPtr hMonitor = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST);
        MONITORINFOEX mi = new MONITORINFOEX();
        mi.cbSize = Marshal.SizeOf(typeof(MONITORINFOEX));
        if (GetMonitorInfo(hMonitor, ref mi)) {
            return mi.szDevice;
        }
        return "Unknown";
    }

    public static bool CloseWindow(IntPtr hWnd) {
        return PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
    }

    /// <summary>
    /// Gets the invisible border offset for a window using DWM API.
    /// Windows 10/11 have invisible borders that affect positioning.
    /// Returns the difference between window rect and extended frame bounds.
    /// </summary>
    public static bool GetFrameOffset(IntPtr hwnd, out int left, out int top, out int right, out int bottom) {
        left = top = right = bottom = 0;

        RECT windowRect, extendedRect;
        if (!GetWindowRect(hwnd, out windowRect)) return false;

        int result = DwmGetWindowAttribute(hwnd, DWMWA_EXTENDED_FRAME_BOUNDS, out extendedRect, Marshal.SizeOf(typeof(RECT)));
        if (result != 0) return false;

        // Invisible border = window rect - extended frame bounds
        left = extendedRect.Left - windowRect.Left;
        top = extendedRect.Top - windowRect.Top;
        right = windowRect.Right - extendedRect.Right;
        bottom = windowRect.Bottom - extendedRect.Bottom;

        return true;
    }
}
"@
}

#endregion

#region PowerShell Helper Functions

function Get-TeamerWindowFrameOffset {
    <#
    .SYNOPSIS
        Gets the invisible border offset for a window using DWM API.
        Windows 10/11 have invisible borders (~7px) that affect positioning.
    .PARAMETER Handle
        Window handle (IntPtr)
    .OUTPUTS
        Hashtable with Left, Top, Right, Bottom offsets
    .NOTES
        Falls back to hardcoded values if DWM query fails
    #>
    param(
        [Parameter(Mandatory)]
        [IntPtr]$Handle
    )

    $left = 0
    $top = 0
    $right = 0
    $bottom = 0

    # Try to get actual frame offset from DWM
    $success = [TeamerWin32]::GetFrameOffset($Handle, [ref]$left, [ref]$top, [ref]$right, [ref]$bottom)

    if ($success -and ($left -gt 0 -or $right -gt 0 -or $bottom -gt 0)) {
        return @{
            Left = $left
            Top = $top
            Right = $right
            Bottom = $bottom
            Source = 'DWM'
        }
    }

    # Fallback to Windows 10/11 standard invisible border sizes
    # These are consistent across most applications:
    # Left: ~7px, Top: 0px, Right: ~7px, Bottom: ~7px
    return @{
        Left = 7
        Top = 0
        Right = 7
        Bottom = 7
        Source = 'Fallback'
    }
}

function Move-TeamerWindowPosition {
    <#
    .SYNOPSIS
        Moves and resizes a window with automatic invisible border compensation
    .PARAMETER Handle
        Window handle (IntPtr)
    .PARAMETER X
        Target X position (visible window position)
    .PARAMETER Y
        Target Y position (visible window position)
    .PARAMETER Width
        Target width (visible window width)
    .PARAMETER Height
        Target height (visible window height)
    .OUTPUTS
        Boolean indicating success
    #>
    param(
        [Parameter(Mandatory)]
        [IntPtr]$Handle,

        [Parameter(Mandatory)]
        [int]$X,

        [Parameter(Mandatory)]
        [int]$Y,

        [Parameter(Mandatory)]
        [int]$Width,

        [Parameter(Mandatory)]
        [int]$Height
    )

    if ($Handle -eq [IntPtr]::Zero) {
        return $false
    }

    # Get frame offset to compensate for invisible borders
    $frameOffset = Get-TeamerWindowFrameOffset -Handle $Handle

    # Adjust position: move frame left/up so visible part is at target position
    $adjustedX = $X - $frameOffset.Left
    $adjustedY = $Y - $frameOffset.Top

    # Adjust size: include invisible borders in total window size
    $adjustedWidth = $Width + $frameOffset.Left + $frameOffset.Right
    $adjustedHeight = $Height + $frameOffset.Top + $frameOffset.Bottom

    # Use SetWindowPos for reliable positioning
    $result = [TeamerWin32]::SetWindowPos(
        $Handle,
        [IntPtr]::Zero,
        $adjustedX,
        $adjustedY,
        $adjustedWidth,
        $adjustedHeight,
        [TeamerWin32]::SWP_NOZORDER -bor [TeamerWin32]::SWP_NOACTIVATE
    )

    return $result
}

#endregion

Write-Verbose "TeamerWin32 module loaded"
