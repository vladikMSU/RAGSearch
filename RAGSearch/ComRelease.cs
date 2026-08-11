using System;
using System.Runtime.InteropServices;

namespace RAGSearch
{
    internal static class ComRelease
    {
        public static void Final(object value)
        {
            if (value == null || !Marshal.IsComObject(value))
            {
                return;
            }

            try
            {
                Marshal.ReleaseComObject(value);
            }
            catch (InvalidComObjectException)
            {
                // The RCW was already released by another cleanup path.
            }
        }
    }
}
