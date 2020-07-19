using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace Carbuncle
{
    public static class TypeInformation
    {
        public static string GetTypeName(object comObject)
        {
            var dispatch = comObject as IDispatch;

            if (dispatch == null)
            {
                return null;
            }

            var pTypeInfo = dispatch.GetTypeInfo(0, 1033);

            string pBstrName;
            string pBstrDocString;
            int pdwHelpContext;
            string pBstrHelpFile;
            pTypeInfo.GetDocumentation(
                -1,
                out pBstrName,
                out pBstrDocString,
                out pdwHelpContext,
                out pBstrHelpFile);

            string str = pBstrName;
            if (str[0] == 95)
            {
                // remove leading '_'
                str = str.Substring(1);
            }

            return str;
        }

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("00020400-0000-0000-C000-000000000046")]
        private interface IDispatch
        {
            int GetTypeInfoCount();

            [return: MarshalAs(UnmanagedType.Interface)]
            ITypeInfo GetTypeInfo(
                [In, MarshalAs(UnmanagedType.U4)] int iTInfo,
                [In, MarshalAs(UnmanagedType.U4)] int lcid);

            void GetIDsOfNames(
                [In] ref Guid riid,
                [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames,
                [In, MarshalAs(UnmanagedType.U4)] int cNames,
                [In, MarshalAs(UnmanagedType.U4)] int lcid,
                [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        }
    }
}
