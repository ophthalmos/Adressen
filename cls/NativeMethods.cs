using System.Runtime.InteropServices;

namespace Adressen.cls;

internal static partial class NativeMethods
{
    private const uint GW_HWNDNEXT = 2;
    internal const int EC_LEFTMARGIN = 1;
    internal const int EC_RIGHTMARGIN = 2;
    internal const int EM_SETRECT = 0x00B3;
    internal const int EM_SETMARGINS = 0xD3;
    internal const int WM_SETFONT = 0x0030;
    internal const int WM_UNDO = 0x304;
    internal const int EM_CANUNDO = 0x00C6;
    internal const int VK_CONTROL = 0x11;
    internal const int EM_SETCUEBANNER = 0x1501;
    internal const int WM_SETTINGCHANGE = 0x001A;
    internal const int WM_LBUTTONDBLCLK = 0x0203;
    internal const int EM_GETRECT = 0x00B2;
    internal const uint WM_SETREDRAW = 0x000B; // Typ auf uint geändert, passend zur Signatur
    internal const int WM_PAINT = 0x000F;
    internal const int WM_PASTE = 0x0302;
    internal const int WM_PRINTCLIENT = 0x0318;
    internal const int PRF_CLIENT = 0x04;
    internal const int PRF_ERASEBKGND = 0x08;
    internal const int VK_UP = 0x26;
    internal const int VK_DOWN = 0x28;
    internal static readonly nint HWND_TOPMOST = -1;
    internal const uint SWP_NOSIZE = 0x0001;
    internal const uint SWP_NOMOVE = 0x0002;
    internal const uint SWP_NOACTIVATE = 0x0010;
    internal const int WM_TRAY_RESTORE = 0x8001; // Eigener Weckruf für AHK
    internal const int WM_TRAY_MINIMIZE = 0x8002; // Befehl von AHK: "Minimiere mich (intelligent)"

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool SetCursorPos(int x, int y);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool SetWindowPos(nint hWnd, nint hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    [LibraryImport("user32.dll")]
    internal static partial nint GetActiveWindow();

    [LibraryImport("user32.dll")]
    internal static partial short GetAsyncKeyState(int vKey);

    [LibraryImport("user32.dll", EntryPoint = "SendMessageW")]
    internal static partial nint SendMessage(nint hWnd, int msg, nint wParam, ref RECT lParam);

    [LibraryImport("user32.dll", EntryPoint = "SendMessageW", StringMarshalling = StringMarshalling.Utf16)]
    internal static partial int SendMessage(nint hWnd, int msg, int wParam, string lParam);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool ValidateRect(nint hWnd, nint lpRect);

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

    [LibraryImport("user32.dll", EntryPoint = "GetWindowTextW", StringMarshalling = StringMarshalling.Utf16)]
    private static partial int GetWindowText(nint hWnd, [Out] char[] lpString, int nMaxCount);

    public static nint GetLastVisibleHandleByTitleEnd(string endString)
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

            currentWindow = GetWindow(currentWindow, GW_HWNDNEXT);
        }
        return nint.Zero;
    }
}