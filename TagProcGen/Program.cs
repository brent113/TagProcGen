using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TagProcGen
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static int Main(string[] args)
        {
            if (args?.Length > 0)
            {
                if (args.Length != 1 || !File.Exists(args[0]))
                {
                    WriteUsage();
                    return 1;
                }

                GenTags.Generate(args[0]);
            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new FormMain());
            }

            return 0;
        }

        static void WriteUsage()
        {
            Console.WriteLine("TagProcGen");
            Console.WriteLine("RTAC and OSI SCADA Configuration Builder");
            Console.WriteLine("");
            Console.WriteLine("Usage:");
            Console.WriteLine("  TagProcGen.exe path_to_configuration.xls[x]");
        }
    }
}
