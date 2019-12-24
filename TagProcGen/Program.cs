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
            var cn = new ConsoleNotifier();

            if (args?.Length > 0)
            {
                if (args.Length != 1 || !File.Exists(args[0]))
                {
                    WriteUsage();
                    return 1;
                }

                GenTags.Generate(args[0], cn);
            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new FormMain());
            }

            return cn.ErrorHasOccured ? 1 : 0;
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

    /// <summary>Console Notifier</summary>
    public class ConsoleNotifier : INotifier
    {
        /// <summary>Tracks if an error has been logged</summary>
        public bool ErrorHasOccured { get; set; } = false;

        /// <summary>Write to log</summary>
        /// <param name="Log">Log text</param>
        /// <param name="Title">Log Title</param>
        /// <param name="Severity">Log Severity</param>
        public void Log(string Log, string Title, LogSeverity Severity)
        {
            string severityText;
            switch (Severity)
            {
                case LogSeverity.Info:
                    severityText = "Info"; break;
                case LogSeverity.Warning:
                    severityText = "Warning"; break;
                case LogSeverity.Error:
                default:
                    severityText = "Error";
                    ErrorHasOccured = true;
                    break;
            }
            var logLines = Log.Split('\n').ToList();
            // Print Title / Severity only on first line
            Console.WriteLine("{0, -10} {1, -20} {2}", severityText, Title, logLines[0]);
            logLines.RemoveAt(0);
            foreach (var s in logLines)
                Console.WriteLine("{0, -10} {1, -20} {2}", "", "", s);
        }
    }
}
