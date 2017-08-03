
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CrystalReportsNinja
{
    public class ArgumentContainer
    {
        /// <summary>
        /// -U Report database login username (mandatory unless integrated security login)
        /// </summary>
        public string UserName { get; set; }
        
        /// <summary>
        /// -P Report database login password (mandatory unless integrated security login)
        /// </summary>
        public string Password { get; set; }                        
        
        /// <summary>
        /// -F Crystal Report path and filename (mandatory)
        /// </summary>
        public string ReportPath { get; set; }                      
        
        /// <summary>
        /// -O Output file path and filename 
        /// </summary>
        public string OutputPath { get; set; }                      
        
        /// <summary>
        /// -S Server name of specified crystal Report (mandatory)
        /// </summary>
        public string ServerName { get; set; }                      
        
        /// <summary>
        /// -D Database name of specified crystal Report
        /// </summary>
        public string DatabaseName { get; set; }                      
        
        /// <summary>
        /// -E Crystal Report exported format (pdf,xls,htm). "print" to print to printer
        /// </summary>
        public string OutputFormat { get; set; }                      
        
        /// <summary>
        /// -a Report Parameter set (small letter p). eg:{customer : "Microsoft Corporation"} or {customer : "Microsoft Corporation" | "Google Inc"}
        /// </summary>
        public List<string> ParameterCollection { get; set; }

        //below are action related properties

        /// <summary>
        /// To print or to export as file. true - print, false
        /// </summary>
        public bool PrintOutput { get; set; }

        /// <summary>
        /// -C number of copy to be printed out
        /// </summary>
        public int PrintCopy { get; set; }
        
        /// <summary>
        /// -I Printer name
        /// </summary>
        public string PrinterName { get; set; }

        /// To display Help message
        /// </summary>
        public bool GetHelp { get; set; }
        
        /// <summary>
        /// To produce log file
        /// </summary>
        public bool EnableLog { get; set; }
        
        /// <summary>
        /// Email address to email to
        /// </summary>
        public string MailTo { get; set; }
        
        /// <summary>
        /// To refresh report or not
        /// </summary>
        public bool Refresh { get; set; }

        public ArgumentContainer()
        {
            // Assigning default values
            GetHelp = false;
            EnableLog = false;
            PrintOutput = false;
            PrintCopy = 1;
            PrinterName = "";
            MailTo = "";
            Refresh = true;

            // Collection of string to store parameters
            ParameterCollection = new List<string>();
        }

        public void ReadArguments(string[] parameters)
        {
            if (parameters.Length == 0)
                throw new Exception("No parameter is specified!");

            #region Assigning crexport parameters to variables
            for (int i = 0; i < parameters.Count(); i++)
            {
                if (i + 1 < parameters.Count())
                {
                    if (parameters[i + 1].Length > 0)
                    {
                        if (parameters[i].ToUpper() == "-U")
                            UserName = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-P")
                            Password = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-F")
                            ReportPath = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-O")
                            OutputPath = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-S")
                            ServerName = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-D")
                            DatabaseName = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-E")
                            OutputFormat = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-N")
                            PrinterName = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-C")
                        {
                            try
                            {
                                PrintCopy = Convert.ToInt32(parameters[i + 1]);
                            }
                            catch (Exception ex)
                            { throw ex; }
                        }
                        else if (parameters[i].ToUpper() == "-A")
                            ParameterCollection.Add(parameters[i + 1]);
                        else if (parameters[i].ToUpper() == "-TO")
                            MailTo = parameters[i + 1];
                    }
                }

                if (parameters[i] == "-?" || parameters[i] == "/?")
                    GetHelp = true;

                if (parameters[i].ToUpper() == "-L")
                    EnableLog = true;

                if (parameters[i].ToUpper() == "-NR")
                    Refresh = false;

            }
            #endregion
        }
    }
}
