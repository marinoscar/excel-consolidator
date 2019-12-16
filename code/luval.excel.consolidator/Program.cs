using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace luval.excel.consolidator
{
    class Program
    {
        private static FileInfo _log;
        private static StreamWriter _stream;
        static void Main(string[] args)
        {
            var options = new Arguments(args);
            var consolidator = new Consolidator();
            consolidator.Status += Consolidator_Status;
            var dirInfo = new DirectoryInfo(options.InputFolder);
            var files = dirInfo.GetFiles("*.xlsx", SearchOption.AllDirectories)
                .Where(i => !i.Name.Equals(options.OutputFile)).ToArray();
            _log = new FileInfo(options.LogFile);
            using (_stream = new StreamWriter(_log.FullName))
            {
                try
                {
                    Console.WriteLine();
                    Console.WriteLine("Runnning version: {0}", Assembly.GetExecutingAssembly().GetName().Version);
                    Console.WriteLine();
                    Console.WriteLine("Starting the process with {0} files", files.Length);
                    Console.WriteLine();
                    consolidator.Execute(options, new FileInfo(options.OutputFile), files);
                    Console.WriteLine();
                }
                catch(Exception ex)
                {
                    LogMessage("** ERROR ** " + ex.ToString());
                }
            }
            Console.WriteLine();
            Console.WriteLine("Press any key to end...");
            Console.ReadKey();
        }
        private static void Consolidator_Status(object sender, ConsolidatorEventArgs e)
        {
            LogMessage(e.Message);
        }

        private static void LogMessage(string message)
        {
            var formatted = string.Format("[{0}] - {1}", DateTime.Now, message);
            Console.WriteLine(formatted);
            _stream.WriteLine(formatted);
        }
    }
}
