using System;
using System.IO;
using System.Diagnostics;

namespace CrystalReportsNinja
{
    public class LogWriter
    {
        private static string _progDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\";
        private static string _logFilename;
        private static Boolean _LogToConsole;

        public LogWriter(string filename, Boolean logToConsole)
        {
            _logFilename = filename;

            _LogToConsole = logToConsole;
        }

        public void Write(string text)
        {
            Trace.WriteLine(string.Format("{0}\t{1}\t{2}", DateTime.Now.ToString("dd-MM-yyyy"), DateTime.Now.ToString("HH:mm:ss"), text));

            if (_logFilename.Length > 0)
            {
                StreamWriter writer;

                if (!File.Exists(_progDir + _logFilename))
                    writer = File.CreateText(_progDir + _logFilename);
                else
                    writer = File.AppendText(_progDir + _logFilename);

                string date, time;
                date = DateTime.Now.ToString("dd-MM-yyyy");
                time = DateTime.Now.ToString("HH:mm:ss");
                                
                writer.WriteLine(string.Format("{0}\t{1}\t{2}", date, time, text));
                writer.Close();
                writer.Dispose();
            }

            if (_LogToConsole)
            {
                Console.WriteLine(string.Format("{0}\t{1}\t{2}", DateTime.Now.ToString("dd-MM-yyyy"), DateTime.Now.ToString("HH:mm:ss"), text));
            }            
        }
    }
}
