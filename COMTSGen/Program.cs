using COMProcessor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;

namespace COMTSGen
{
    internal class Program
    {
        static int Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("Usage: COMProcessor.CLI <CLSID> [output.d.ts]");
                return 1;
            }

            try
            {
                Guid clsid = new Guid(args[0]);

                string output = args.Length >= 2
                    ? args[1]
                    : "com.d.ts";

                ITypeLib typeLib = ComInspector.LoadTypeLibrary(clsid);

                try
                {
                    TypeScriptGenerator.Generate(typeLib, output);
                }
                finally
                {
                    Marshal.ReleaseComObject(typeLib);
                }

                Console.WriteLine("Generated: " + output);
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex);
                return 1;
            }
        }
    }
}
