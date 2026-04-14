using System.Runtime.InteropServices;

namespace Keboo.VocalSlide.Infrastructure;

internal static class ComInteropHelper
{
    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgIDEx(string progId, out Guid clsid);

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string progId, out Guid clsid);

    [DllImport("oleaut32.dll")]
    private static extern int GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object? instance);

    public static object GetRunningObject(string progId)
    {
        int result = CLSIDFromProgIDEx(progId, out Guid clsid);
        if (result < 0)
        {
            result = CLSIDFromProgID(progId, out clsid);
        }

        Marshal.ThrowExceptionForHR(result);

        result = GetActiveObject(ref clsid, IntPtr.Zero, out object? instance);
        Marshal.ThrowExceptionForHR(result);

        return instance ?? throw new InvalidOperationException($"Could not resolve a running COM object for '{progId}'.");
    }

    public static void ReleaseIfNeeded(object? comObject)
    {
        if (comObject is not null && Marshal.IsComObject(comObject))
        {
            Marshal.ReleaseComObject(comObject);
        }
    }
}
