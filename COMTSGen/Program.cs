using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using COMProcessor;

namespace COMTSGen
{
    internal class Program
    {
        static int Main(string[] args)
        {
            if (args.Length != 1)
            {
                Console.WriteLine("Usage: ComTypeGen.CLI <CLSID>");
                return 1;
            }

            try
            {

                Guid clsid = new Guid(args[0]);

                

                ComInspector.EnumerateTypes(clsid);

                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
                return 1;
            }
        }
    }
}
