
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
        /// -I Use Integrated Security for Database Credentials
        /// </summary>
        public Boolean IntegratedSecurity { get; set; }

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

        /// <summary>
        /// To print or to export as file. true - print, false
        /// </summary>
        public bool PrintOutput { get; set; }

        /// <summary>
        /// -C number of copy to be printed out
        /// </summary>
        public int PrintCopy { get; set; }
        
        /// <summary>
        /// -N Printer name
        /// </summary>
        public string PrinterName { get; set; }

        /// To display Help message
        /// </summary>
        public bool GetHelp { get; set; }

        /// <summary>
        /// -M To Email out the report or not
        /// </summary>
        public bool EmailOutput { get; set; }

        /// <summary>
        /// -MF From Email Address.  Defaults to noreply@noreply.com.au
        /// </summary>
        public string MailFrom { get; set; }

        /// <summary>
        /// -MN From Email Address.  Defaults to Crystal Reports
        /// </summary>
        public string MailFromName { get; set; }
        /// <summary>
        /// -MT Email address to email to.  (Mandatory for emailing report)
        /// </summary>
        public string MailTo { get; set; }

        /// <summary>
        /// -MC Email address to CC to.  (Optional for emailing report)
        /// </summary>
        public string MailCC { get; set; }

        /// <summary>
        /// -MB Email address to Bcc to.  (Optional for emailing report)
        /// </summary>
        public string MailBcc { get; set; }

        /// <summary>
        /// -MSA SMTP Address.  (Mandatory for emailing report)
        /// </summary>
        public String SmtpServer { get; set; }

        /// <summary>
        /// -MSP SMTP Port.  Defaults to Port 25
        /// </summary>
        public int SmtpPort { get; set; }

        /// <summary>
        /// -MSE SMTP Enable SSL. Defaults to False
        /// </summary>
        public Boolean SmtpSSL { get; set; }

        /// <summary>
        /// -MSC SMTP Use Current User Credentials. Defaults to False
        /// </summary>
        public Boolean SmtpAuth { get; set; }

        /// <summary>
        /// -MSU SMTP Username.
        /// </summary>
        public string SmtpUN { get; set; }

        /// <summary>
        /// -MSP SMTP Password.
        /// </summary>
        public string SmtpPW { get; set; }

        /// <summary>
        /// -MS Email Subject.  Defaults to Crystal Reports
        /// </summary>
        public String EmailSubject { get; set; }

        /// <summary>
        /// -MI Email Body.
        /// </summary>
        public String EmailBody { get; set; }

        /// <summary>
        /// -MK Email Keep File.  Defaults to false
        /// </summary>
        public Boolean EmailKeepFile { get; set; }

        /// <summary>
        /// -L To produce log file
        /// </summary>
        public bool EnableLog { get; set; }

        /// <summary>
        /// -LC Write log output to Console
        /// </summary>
        public bool EnableLogToConsole { get; set; }

        /// <summary>
        /// Log File Name
        /// </summary>
        public String LogFileName { get; set; }

        /// <summary>
        /// -NR To refresh report or not
        /// </summary>
        public bool Refresh { get; set; }

        /// <summary>
        /// -SF Report Selection Formula
        /// </summary>
        public string SelectionFormula { get; set; }

        /// <summary>
        /// The Locale ID to set in the ReportClientDocument
        /// </summary>
        public Int32? Culture { get; set; }

        public ArgumentContainer()
        {
            // Assigning default values
            GetHelp = false;
            EnableLog = false;
            PrintOutput = false;
            PrintCopy = 1;
            PrinterName = "";
            Refresh = true;
            SelectionFormula = null;
            Culture = null;
            IntegratedSecurity = false;
            UserName = null;
            Password = null;
            ReportPath = null;
            OutputPath = null;
            ServerName = null;
            OutputFormat = null;

            //Email Config
            MailTo = null;
            MailBcc = "NA";
            MailCC = "NA";
            EmailOutput = false;
            MailFrom = "noreply@noreply.com";
            MailFromName = "Crystal Reports";
            SmtpServer = null;
            SmtpPort = 25;
            SmtpSSL = false;
            SmtpAuth = false;
            SmtpUN = null;
            SmtpPW = null;
            EmailSubject = "Crystal Reports";
            EmailBody = "NA";
            EmailKeepFile = false;

            //Logging Options
            EnableLogToConsole = false;
            LogFileName = String.Empty;

            // Collection of string to store parameters
            ParameterCollection = new List<string>();
        }

        public void ReadArguments(string[] parameters)
        {
            if (parameters.Length == 0)
                //throw new Exception("No parameter is specified!");
                GetHelp = true;

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
                        else if (parameters[i].Equals("-I"))
                            IntegratedSecurity = true;
                        else if (parameters[i].ToUpper() == "-F")
                            ReportPath = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-O")
                            OutputPath = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-S")
                            ServerName = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-D")
                            DatabaseName = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-E")
                        { OutputFormat = parameters[i + 1]; if (OutputFormat.ToUpper() == "PRINT") { PrintOutput = true; } }
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
                        else if (parameters[i].ToUpper() == "-SF")
                            SelectionFormula = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-CU")
                            Culture = int.Parse(parameters[i + 1]);

                        //Email Config
                        else if (parameters[i].ToUpper() == "-M")
                            EmailOutput = true;
                        else if (parameters[i].ToUpper() == "-MF")
                            MailFrom = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-MN")
                            MailFromName = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-MS")
                            EmailSubject = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-MI")
                            EmailBody = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-MT")
                            MailTo = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-MC")
                            MailCC = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-MB")
                            MailBcc = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-MK")
                            EmailKeepFile = true;
                        else if (parameters[i].ToUpper() == "-MSA")
                            SmtpServer = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-MSP")
                            SmtpPort = Convert.ToInt32(parameters[i + 1]);
                        else if (parameters[i].ToUpper() == "-MSE")
                            SmtpSSL = true;
                        else if (parameters[i].ToUpper() == "-MSC")
                            SmtpAuth = true;
                        else if (parameters[i].ToUpper() == "-MUN")
                            SmtpUN = parameters[i + 1];
                        else if (parameters[i].ToUpper() == "-MPW")
                            SmtpPW = parameters[i + 1];
                    }
                }

                if (parameters[i] == "-?" || parameters[i] == "/?")
                    GetHelp = true;

                if (parameters[i].ToUpper() == "-L")
                    EnableLog = true;

                if (parameters[i].ToUpper() == "-NR")
                    Refresh = false;

                if (parameters[i].ToUpper() == "-LC")
                    EnableLogToConsole = true;
            }
            #endregion
        }
    }
}
