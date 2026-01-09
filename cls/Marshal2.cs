using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security;

namespace Adressen.cls;

internal static class Marshal2
{
    internal const string OLEAUT32 = "oleaut32.dll";
    internal const string OLE32 = "ole32.dll";

    [SecurityCritical]
    public static object GetActiveObject(string progID)
    {
        Guid clsid;
        try { CLSIDFromProgIDEx(progID, out clsid); }
        catch (Exception) { CLSIDFromProgID(progID, out clsid); }
        GetActiveObject(ref clsid, nint.Zero, out var obj);
        return obj;
    }

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

    [DllImport(OLEAUT32, PreserveSig = false)]
    [ResourceExposure(ResourceScope.None)]
    [SuppressUnmanagedCodeSecurity]
    [SecurityCritical]
    private static extern void GetActiveObject(ref Guid rclsid, nint reserved, [MarshalAs(UnmanagedType.Interface)] out object ppunk);
}
