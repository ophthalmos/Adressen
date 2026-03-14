using System.Runtime.InteropServices;

namespace Adressen.cls;

internal static partial class NativeMethods
{
    private const uint GW_HWNDNEXT = 2;
    internal const int EC_LEFTMARGIN = 1;
    internal const int EC_RIGHTMARGIN = 2;
    internal const int EM_SETRECT = 0x00B3;
    internal const int EM_SETMARGINS = 0xD3;
    internal const int VK_CONTROL = 0x11;
    internal const int EM_SETCUEBANNER = 0x1501;
    internal const int WM_SETTINGCHANGE = 0x001A;
    internal const int WM_LBUTTONDBLCLK = 0x0203;
    internal const int EM_GETRECT = 0x00B2;
    internal const int WM_PAINT = 0x000F;
    internal const int WM_PRINTCLIENT = 0x0318;
    internal const int PRF_CLIENT = 0x04;
    internal const int PRF_ERASEBKGND = 0x08;


    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [LibraryImport("user32.dll", EntryPoint = "SendMessageW")]
    public static partial IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, ref RECT lParam);

    // Hier den EntryPoint auf "SendMessageW" setzen
    [LibraryImport("user32.dll", EntryPoint = "SendMessageW", StringMarshalling = StringMarshalling.Utf16)]
    internal static partial int SendMessage(nint hWnd, int msg, int wParam, string lParam);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool ValidateRect(IntPtr hWnd, IntPtr lpRect);

    // Auch hier für die numerische Variante
    [LibraryImport("user32.dll", EntryPoint = "SendMessageW")]
    internal static partial nint SendMessage(nint hWnd, uint Msg, nint wParam, nint lParam);

    [LibraryImport("user32.dll")]
    internal static partial short GetKeyState(int nVirtKey);

    [LibraryImport("shell32.dll")]
    internal static partial int SHGetKnownFolderPath(in Guid rfid, uint dwFlags, nint hToken, out nint ppszPath);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool SetForegroundWindow(nint hWnd);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool ShowScrollBar(nint hWnd, int wBar, [MarshalAs(UnmanagedType.Bool)] bool bShow);

    [LibraryImport("user32.dll")]
    private static partial nint GetTopWindow(nint hWnd);

    [LibraryImport("user32.dll")]
    private static partial nint GetWindow(nint hWnd, uint uCmd);

    // GetWindowText braucht ebenfalls das "W" am Ende des EntryPoints
    [LibraryImport("user32.dll", EntryPoint = "GetWindowTextW", StringMarshalling = StringMarshalling.Utf16)]
    private static partial int GetWindowText(nint hWnd, [Out] char[] lpString, int nMaxCount); public static nint GetLastVisibleHandleByTitleEnd(string endString)
    {
        var currentWindow = GetTopWindow(nint.Zero);

        while (currentWindow != nint.Zero)
        {
            var buffer = new char[256];
            var length = GetWindowText(currentWindow, buffer, buffer.Length);

            if (length > 0)
            {
                var windowText = new string(buffer, 0, length);

                if (windowText.EndsWith(endString, StringComparison.OrdinalIgnoreCase))
                {
                    return currentWindow;
                }
            }

            currentWindow = GetWindow(currentWindow, GW_HWNDNEXT); // Zum nächsten Fenster in der Z-Reihenfolge wechseln
        }

        return nint.Zero;
    }
}

//using System.Runtime.InteropServices;
//using System.Text;

//namespace Adressen.cls;

//internal static class NativeMethods
//{
//    internal const int EC_LEFTMARGIN = 1;
//    internal const int EC_RIGHTMARGIN = 2;
//    internal const int EM_SETMARGINS = 0xD3;
//    internal const int VK_CONTROL = 0x11;
//    internal const int EM_SETCUEBANNER = 0x1501;
//    private const uint GW_HWNDNEXT = 2;
//    internal const int WM_SETTINGCHANGE = 0x001A;


//    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
//    internal static extern int SendMessage(nint hWnd, int msg, int wParam, [MarshalAs(UnmanagedType.LPWStr)] string lParam);

//    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
//    internal static extern nint SendMessage(nint hWnd, uint Msg, nint wParam, nint lParam);

//    [DllImport("user32.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
//    internal static extern short GetKeyState(int nVirtKey);

//    //[DllImport("shell32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, PreserveSig = false)]
//    //internal static extern string SHGetKnownFolderPath([MarshalAs(UnmanagedType.LPStruct)] Guid rfid, uint dwFlags, nint hToken = default);
//    [DllImport("shell32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, PreserveSig = false)]
//    internal static extern void SHGetKnownFolderPath([MarshalAs(UnmanagedType.LPStruct)] Guid rfid, uint dwFlags, nint hToken, out string ppszPath); // 'out string' sorgt hier automatisch für das Freigeben des Speichers


//    [DllImport("user32.dll")]
//    [return: MarshalAs(UnmanagedType.Bool)]
//    internal static extern bool SetForegroundWindow(nint hWnd);

//    [DllImport("user32.dll")]
//    [return: MarshalAs(UnmanagedType.Bool)]
//    internal static extern bool ShowScrollBar(nint hWnd, int wBar, [MarshalAs(UnmanagedType.Bool)] bool bShow);

//    [DllImport("user32.dll")]
//    private static extern nint GetTopWindow(nint hWnd);

//    [DllImport("user32.dll")]
//    private static extern nint GetWindow(nint hWnd, uint uCmd);

//    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
//    private static extern int GetWindowText(nint hWnd, StringBuilder lpString, int nMaxCount);

//    public static nint GetLastVisibleHandleByTitleEnd(string endString)
//    {
//        var currentWindow = GetTopWindow(nint.Zero);
//        while (currentWindow != nint.Zero)
//        {
//            var sb = new StringBuilder(256);
//            _ = GetWindowText(currentWindow, sb, sb.Capacity);
//            if (sb.ToString().EndsWith(endString, StringComparison.OrdinalIgnoreCase)) { return currentWindow; }
//            currentWindow = GetWindow(currentWindow, GW_HWNDNEXT); // Zum nächsten Fenster in der Z-Reihenfolge wechseln
//        }
//        return nint.Zero;
//    }
//}
