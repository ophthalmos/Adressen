using System.Runtime.InteropServices;
using System.Text;

namespace Adressen.cls;

internal static class NativeMethods
{
    internal const int EC_LEFTMARGIN = 1;
    internal const int EC_RIGHTMARGIN = 2;
    internal const int EM_SETMARGINS = 0xD3;
    internal const int VK_CONTROL = 0x11;
    internal const int EM_SETCUEBANNER = 0x1501;
    private const uint GW_HWNDNEXT = 2;
    internal const int WM_SETTINGCHANGE = 0x001A;


    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern int SendMessage(nint hWnd, int msg, int wParam, [MarshalAs(UnmanagedType.LPWStr)] string lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern nint SendMessage(nint hWnd, uint Msg, nint wParam, nint lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    internal static extern short GetKeyState(int nVirtKey);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, PreserveSig = false)]
    internal static extern string SHGetKnownFolderPath([MarshalAs(UnmanagedType.LPStruct)] Guid rfid, uint dwFlags, nint hToken = default);


    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool SetForegroundWindow(nint hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool ShowScrollBar(nint hWnd, int wBar, [MarshalAs(UnmanagedType.Bool)] bool bShow);

    [DllImport("user32.dll")]
    private static extern nint GetTopWindow(nint hWnd);

    [DllImport("user32.dll")]
    private static extern nint GetWindow(nint hWnd, uint uCmd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(nint hWnd, StringBuilder lpString, int nMaxCount);

    public static nint GetLastVisibleHandleByTitleEnd(string endString)
    {
        var currentWindow = GetTopWindow(nint.Zero);
        while (currentWindow != nint.Zero)
        {
            var sb = new StringBuilder(256);
            _ = GetWindowText(currentWindow, sb, sb.Capacity);
            if (sb.ToString().EndsWith(endString, StringComparison.OrdinalIgnoreCase)) { return currentWindow; }
            currentWindow = GetWindow(currentWindow, GW_HWNDNEXT); // Zum nächsten Fenster in der Z-Reihenfolge wechseln
        }
        return nint.Zero;
    }
}
