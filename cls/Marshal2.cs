using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security;

namespace Adressen.cls;

internal static partial class Marshal2
{
    internal const string OLEAUT32 = "oleaut32.dll";
    internal const string OLE32 = "ole32.dll";
    private const int S_OK = 0;

    [SecurityCritical]
    public static object? GetActiveObject(string progID)
    {
        var hr = CLSIDFromProgIDEx(progID, out var clsid);
        if (hr != S_OK)
        {
            hr = CLSIDFromProgID(progID, out clsid);
            if (hr != S_OK) { Marshal.ThrowExceptionForHR(hr); }
        }
        var hr2 = GetActiveObject(in clsid, nint.Zero, out var pUnk);
        if (hr2 == S_OK && pUnk != nint.Zero)
        {
            try { return Marshal.GetObjectForIUnknown(pUnk); }
            finally { Marshal.Release(pUnk); }
        }
        return null;
    }

    [LibraryImport(OLEAUT32)]
    [ResourceExposure(ResourceScope.None)]
    [SuppressUnmanagedCodeSecurity]
    [SecurityCritical]
    private static partial int GetActiveObject(in Guid rclsid, nint reserved, out nint ppunk);

    [LibraryImport(OLE32, StringMarshalling = StringMarshalling.Utf16)]
    [ResourceExposure(ResourceScope.None)]
    [SuppressUnmanagedCodeSecurity]
    [SecurityCritical]
    private static partial int CLSIDFromProgIDEx(string progId, out Guid clsid);

    [LibraryImport(OLE32, StringMarshalling = StringMarshalling.Utf16)]
    [ResourceExposure(ResourceScope.None)]
    [SuppressUnmanagedCodeSecurity]
    [SecurityCritical]
    private static partial int CLSIDFromProgID(string progId, out Guid clsid);
}
