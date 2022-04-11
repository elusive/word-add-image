namespace word_add_diagram
{
    using System;
    using System.Runtime.InteropServices;

    public class NativeMethods
    {
        [DllImport("oleaut32.dll", PreserveSig=false)]
        public static extern void GetActiveObject(
            ref Guid   rclsid,
            IntPtr pvReserved,
            [MarshalAs(UnmanagedType.IUnknown)] out Object ppunk
        );

        [DllImport("ole32.dll")]
        public static extern int CLSIDFromProgID(
            [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID,
            out Guid   pclsid
        );

        public static object GetActiveObject(string progId) {
           Guid clsid;
           CLSIDFromProgID(progId, out clsid);

            object obj;
            GetActiveObject(ref clsid, IntPtr.Zero, out obj);

            return obj;
        }
    }
}
