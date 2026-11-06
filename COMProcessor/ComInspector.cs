using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;

namespace COMProcessor
{
    public class ComInspector
    {
        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void LoadRegTypeLib(
            ref Guid guid,
            short major,
            short minor,
            int lcid,
            out ITypeLib typeLib);

        public static ITypeLib LoadTypeLibrary(Guid clsid)
        {
            string clsidPath = @"CLSID\" + clsid.ToString("B");

            using (RegistryKey key =
                Registry.ClassesRoot.OpenSubKey(clsidPath))
            {
                if (key == null)
                    throw new ArgumentException("CLSID not registered.");

                string libid = key.GetValue("TypeLib") as string;
                string version = key.GetValue("Version") as string;

                if (libid == null || version == null)
                    throw new InvalidOperationException(
                        "CLSID has no registered Type Library.");

                string[] parts = version.Split('.');

                Guid typeLibId = new Guid(libid);
                short major = short.Parse(parts[0]);
                short minor = short.Parse(parts[1]);

                ITypeLib typeLib;

                LoadRegTypeLib(
                    ref typeLibId,
                    major,
                    minor,
                    0,
                    out typeLib);

                return typeLib;
            }
        }

        public static void EnumerateTypes(Guid clsid)
        {
            ITypeLib typeLib = LoadTypeLibrary(clsid);

            try
            {
                int count = typeLib.GetTypeInfoCount();

                Console.WriteLine("Types found: " + count);

                for (int i = 0; i < count; i++)
                {
                    System.Runtime.InteropServices.ComTypes.TYPEKIND kind;
                    typeLib.GetTypeInfoType(i, out kind);

                    ITypeInfo info;
                    typeLib.GetTypeInfo(i, out info);

                    try
                    {
                        string name;
                        string documentation;
                        int helpContext;
                        string helpFile;

                        info.GetDocumentation(
                            -1,
                            out name,
                            out documentation,
                            out helpContext,
                            out helpFile);

                        IntPtr attrPtr;
                        info.GetTypeAttr(out attrPtr);

                        try
                        {
                            var attr =
                                (System.Runtime.InteropServices.ComTypes.TYPEATTR)Marshal.PtrToStructure(
                                    attrPtr, typeof(System.Runtime.InteropServices.ComTypes.TYPEATTR));

                            Console.WriteLine(
                                "{0,-18} {1,-35} {2}",
                                kind,
                                name,
                                attr.guid);
                        }
                        finally
                        {
                            info.ReleaseTypeAttr(attrPtr);
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(info);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(typeLib);
            }
        }
    }
}
