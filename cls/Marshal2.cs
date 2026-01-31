using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security;

namespace Adressen.cls;

internal static class Marshal2
{
    internal const string OLEAUT32 = "oleaut32.dll";
    internal const string OLE32 = "ole32.dll";
    private const int S_OK = 0;

    [SecurityCritical]
    public static object? GetActiveObject(string progID)
    {
        Guid clsid;
        try { CLSIDFromProgIDEx(progID, out clsid); }
        catch (Exception) { CLSIDFromProgID(progID, out clsid); }

        // Aufruf mit manuellem HRESULT-Check statt Exception-Handling
        var hr = GetActiveObject(ref clsid, IntPtr.Zero, out var obj);
        if (hr == S_OK) { return obj; }
        return null; // Kein Fehler werfen, einfach null zurückgeben, wenn Word nicht läuft (Objekt nicht gefunden)
    }

    // PreserveSig = true ist Standard, hier explizit gesetzt zur Verdeutlichung.
    // Rückgabetyp ist int (HRESULT), nicht void.
    [DllImport(OLEAUT32, PreserveSig = true)]
    [ResourceExposure(ResourceScope.None)]
    [SuppressUnmanagedCodeSecurity]
    [SecurityCritical]
    private static extern int GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.Interface)] out object? ppunk);

    // Bei diesen Methoden ist Exception-Werfen okay/gewünscht, daher lassen wir PreserveSig = false
    [DllImport(OLE32, PreserveSig = false)]
    [ResourceExposure(ResourceScope.None)]
    [SuppressUnmanagedCodeSecurity]
    [SecurityCritical]
    private static extern void CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

    [DllImport(OLE32, PreserveSig = false)]
    [ResourceExposure(ResourceScope.None)]
    [SuppressUnmanagedCodeSecurity]
    [SecurityCritical]
    private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);
}