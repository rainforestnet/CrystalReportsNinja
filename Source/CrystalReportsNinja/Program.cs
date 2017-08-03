using System;

namespace CrystalReportsNinja
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // read program arguments
                ArgumentContainer argContainer = new ArgumentContainer();
                argContainer.ReadArguments(args);

                if (argContainer.GetHelp)
                    Helper.DisplayMessage(2);
                else
                {
                    string _logFilename = string.Empty;

                    if (argContainer.EnableLog)
                        _logFilename = "ninja-" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".log";

                    ReportProcessor reportNinja = new ReportProcessor(_logFilename)
                    {
                        ReportArguments = argContainer,
                    };

                    reportNinja.Run();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Exception: {0}. Inner Exception: {1}",ex.Message, ex.InnerException));
            }
        }
    }
}
