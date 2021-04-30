using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Globalization;
using System.IO;
using System.Net.Mail;

namespace CrystalReportsNinja
{
    public class ReportProcessor
    {
        private string _sourceFilename;
        private string _outputFilename;
        private string _outputFormat;
        private bool _printToPrinter;

        private ReportDocument _reportDoc;
        private LogWriter _logger;

        public ArgumentContainer ReportArguments { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public ReportProcessor()
        {
            ReportArguments = new ArgumentContainer();

            _reportDoc = new ReportDocument();
        }

        /// <summary>
        /// Load external Crystal Report file into Report Document
        /// </summary>
        private void LoadReport()
        {
            _sourceFilename = ReportArguments.ReportPath.Trim();
            if (_sourceFilename == null || _sourceFilename == string.Empty)
            {
                throw new Exception("Invalid Crystal Reports file");
            }

            if (_sourceFilename.LastIndexOf(".rpt") == -1)
                throw new Exception("Invalid Crystal Reports file");

            if (ReportArguments.Culture.HasValue)
            {
                CultureInfo _CultureInfo = CultureInfo.GetCultureInfo(ReportArguments.Culture.Value);

                _logger.Write(string.Format("Using locale {0} {1}", _CultureInfo.LCID, _CultureInfo.EnglishName));

                _reportDoc.ReportClientDocument.LocaleID = (CrystalDecisions.ReportAppServer.DataDefModel.CeLocale)_CultureInfo.LCID;
                _reportDoc.ReportClientDocument.PreferredViewingLocaleID = (CrystalDecisions.ReportAppServer.DataDefModel.CeLocale)_CultureInfo.LCID;
                _reportDoc.ReportClientDocument.ProductLocaleID = (CrystalDecisions.ReportAppServer.DataDefModel.CeLocale)_CultureInfo.LCID;
            }

            _reportDoc.Load(_sourceFilename, OpenReportMethod.OpenReportByDefault);
            var filenameOnly = System.IO.Path.GetFileNameWithoutExtension(_sourceFilename);
            _logger.Write(string.Format("Report {0}.rpt loaded successfully", filenameOnly));
        }

        /// <summary>
        /// Match User input parameter values with Report parameters
        /// </summary>
        private void ProcessParameters()
        {
            var paramCount = _reportDoc.ParameterFields.Count;
            _logger.Write(string.Format("Number of Parameters detected in the report = {0}", _reportDoc.ParameterFields.Count));
            if (_reportDoc.DataDefinition.ParameterFields.Count > 0)
            {
                ParameterCore paraCore = new ParameterCore(ReportArguments.LogFileName, ReportArguments);
                paraCore.ProcessRawParameters();
                _logger.Write(string.Format(""));
                foreach (ParameterFieldDefinition _ParameterFieldDefinition in _reportDoc.DataDefinition.ParameterFields)
                {
                    if (!_ParameterFieldDefinition.IsLinked())
                        {
                            _logger.Write(string.Format("Applied Parameter '{0}' as MultiValue '{1}'", _ParameterFieldDefinition.Name, _ParameterFieldDefinition.EnableAllowMultipleValue));
                            ParameterValues values = paraCore.GetParameterValues(_ParameterFieldDefinition);
                            _ParameterFieldDefinition.ApplyCurrentValues(values);
                        }
                    else
                        _logger.Write(string.Format("Skipped '{1}' as MultiValue '{2}' Parameter in SubReport = '{0}' as its Linked to Main Report", _ParameterFieldDefinition.ReportName, _ParameterFieldDefinition.Name, _ParameterFieldDefinition.EnableAllowMultipleValue));
                        _logger.Write(string.Format(""));
                }
            }
        }

        /// <summary>
        /// Add selection formula to ReportDocument
        /// </summary>
        private void ProcessSelectionFormula()
        {
            if (!String.IsNullOrWhiteSpace(ReportArguments.SelectionFormula))
            {
                _reportDoc.RecordSelectionFormula = ReportArguments.SelectionFormula;
            }
        }

        /// <summary>
        /// Validate configurations related to program output.
        /// </summary>
        /// <remarks>
        /// Program output can be TWO forms
        /// 1. Export as a file
        /// 2. Print to printer
        /// </remarks>
        private void ValidateOutputConfigurations()
        {
            _outputFilename = ReportArguments.OutputPath;
            _outputFormat = ReportArguments.OutputFormat;
            _printToPrinter = ReportArguments.PrintOutput;

            bool specifiedFileName = _outputFilename != null ? true : false;
            bool specifiedFormat = _outputFormat != null ? true : false;

            if (!_printToPrinter)
            {
                string fileExt = "";

                //default set to text file
                if (!specifiedFileName && !specifiedFormat)
                    _outputFormat = "txt";
                	fileExt = "txt";
                // Use output format to set output file name extension
                if (specifiedFormat)
                {
                    if (_outputFormat.ToUpper() == "XLSDATA")
                        fileExt = "xls";
                    else if (_outputFormat.ToUpper() == "TAB")
                        fileExt = "txt";
                    else if (_outputFormat.ToUpper() == "ERTF")
                        fileExt = "rtf";
                    else if (_outputFormat.ToUpper() == "HTM")
                        fileExt = "html";
                    else
                        fileExt = _outputFormat;
                }

                // Use output file name extension to set output format
                if (specifiedFileName && !specifiedFormat)
                {
                    int lastIndexDot = _outputFilename.LastIndexOf(".");
                    if (_outputFilename.Length == lastIndexDot + 4)
                    {
                    fileExt = _outputFilename.Substring(lastIndexDot + 1, 3);

                    //ensure filename extension has 3 char after the dot (.)
                        if (fileExt.ToUpper() == "RTF" || fileExt.ToUpper() == "TXT" || fileExt.ToUpper() == "CSV" || fileExt.ToUpper() == "PDF" || fileExt.ToUpper() == "RPT" || fileExt.ToUpper() == "DOC" || fileExt.ToUpper() == "XLS" || fileExt.ToUpper() == "XML" || fileExt.ToUpper() == "HTM")
                            _outputFormat = _outputFilename.Substring(lastIndexDot + 1, 3);
                    }
                    else if (_outputFilename.Length == lastIndexDot + 5)
                    {
                        fileExt = _outputFilename.Substring(lastIndexDot + 1, 4); // Fix for file ext has 4 char
                        if (fileExt.ToUpper() == "XLSX" || fileExt.ToUpper() == "HTML")
                            _outputFormat = _outputFilename.Substring(lastIndexDot + 1, 4);
                    }
                }
                // Use output file name and extension to set output format and Name if Matching criteria
                if (specifiedFileName && specifiedFormat)
                {
                    int lastIndexDot = _outputFilename.LastIndexOf(".");
                    if ((_outputFilename.Length == lastIndexDot + 4) && (fileExt.ToUpper() != _outputFilename.Substring(lastIndexDot + 1, 3).ToUpper())) //file ext has 3 char
                    {
                        _outputFilename = string.Format("{0}.{1}", _outputFilename, fileExt);
                    }
                    else if ((_outputFilename.Length == lastIndexDot + 5) && (fileExt.ToUpper() != _outputFilename.Substring(lastIndexDot + 1, 4).ToUpper())) //file ext has 4 char
                    {
                        _outputFilename = string.Format("{0}.{1}", _outputFilename, fileExt);
                    }
                }

                if (!specifiedFileName)
                    _outputFilename = String.Format("{0}-{1}.{2}", _sourceFilename.Substring(0, _sourceFilename.LastIndexOf(".rpt")), DateTime.Now.ToString("yyyyMMddHHmmss"), fileExt);

                _logger.Write(string.Format("Output Filename : {0}", _outputFilename));
                _logger.Write(string.Format("Output Format : {0}", _outputFormat));
            }
        }

        /// <summary>
        /// Perform Login to database tables
        /// </summary>
        private void PerformDBLogin()
        {
            if (ReportArguments.Refresh)
            {
                TableLogOnInfo logonInfo;
                foreach (Table table in _reportDoc.Database.Tables)
                {
                    logonInfo = table.LogOnInfo;
                    if (!String.IsNullOrWhiteSpace(ReportArguments.ServerName))
                    { logonInfo.ConnectionInfo.ServerName = ReportArguments.ServerName; }

                    if (!String.IsNullOrWhiteSpace(ReportArguments.DatabaseName))
                    { logonInfo.ConnectionInfo.DatabaseName = ReportArguments.DatabaseName; }

                    if (ReportArguments.IntegratedSecurity)
                    { logonInfo.ConnectionInfo.IntegratedSecurity = true; }
                    else
                    {
                        if (!String.IsNullOrWhiteSpace(ReportArguments.UserName))
                        { logonInfo.ConnectionInfo.UserID = ReportArguments.UserName; }

                        if (!String.IsNullOrWhiteSpace(ReportArguments.Password))
                        { logonInfo.ConnectionInfo.Password = ReportArguments.Password; }
                    }
                    table.ApplyLogOnInfo(logonInfo);
                }
                _logger.Write(string.Format("Logged into {1} Database on {0} successfully with User id: {2}", ReportArguments.ServerName, ReportArguments.DatabaseName, ReportArguments.UserName));
            }
        }

        /// <summary>
        /// Set export file type or printer to Report Document object.
        /// </summary>
        private void ApplyReportOutput()
        {
            if (_printToPrinter)
            {
                var printerName = ReportArguments.PrinterName != null ? ReportArguments.PrinterName.Trim() : "";
                _logger.Write(string.Format("The default printer set in the Report is '{0}'", _reportDoc.PrintOptions.PrinterName));
                if (printerName.Length > 0 && printerName.ToUpper() != "DFLT")
                //Print to the Specified Printer
                {
                    _reportDoc.PrintOptions.NoPrinter = false; //Changes the report option "No Printer: Optimized for Screen"
                    _reportDoc.PrintOptions.PrinterName = printerName;
                    _logger.Write(string.Format("The Specified PrinterName '{0}' is set by Parameter will be used", printerName));
                }
                else if (_reportDoc.PrintOptions.PrinterName.Length > 0 && printerName.ToUpper() == "DFLT")
                //Print to the reports default Printer
                {
                    _reportDoc.PrintOptions.NoPrinter = false; //Changes the report option "No Printer: Optimized for Screen"
                    _logger.Write(string.Format("The Specified PrinterName '{0}' is set in the report and DFLT flag will be used", _reportDoc.PrintOptions.PrinterName));
                }
                else
                //Print to the Windows default Printer
                {
                    System.Drawing.Printing.PrinterSettings prinSet = new System.Drawing.Printing.PrinterSettings();
                    _logger.Write(string.Format("Printer is not specified - The Windows Default Printer '{0}' will be used", prinSet.PrinterName));
                        if (prinSet.PrinterName.Trim().Length > 0)
                        {
                            _reportDoc.PrintOptions.NoPrinter = false; //Changes the report option "No Printer: Optimized for Screen"
                            _reportDoc.PrintOptions.PrinterName = prinSet.PrinterName;
                        }
                    else
                        throw new Exception("No printer name is specified");
                }
            }
            else
            {
                if (_outputFormat.ToUpper() == "RTF")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.RichText;
                else if (_outputFormat.ToUpper() == "TXT")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.Text;
                else if (_outputFormat.ToUpper() == "TAB")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.TabSeperatedText;
                else if (_outputFormat.ToUpper() == "CSV")
                {
                    CharacterSeparatedValuesFormatOptions csvExpOpts = new CharacterSeparatedValuesFormatOptions();
                    csvExpOpts.ExportMode = CsvExportMode.Standard;
                    csvExpOpts.GroupSectionsOption = CsvExportSectionsOption.Export;
                    csvExpOpts.ReportSectionsOption = CsvExportSectionsOption.Export;
                    csvExpOpts.GroupSectionsOption = CsvExportSectionsOption.ExportIsolated;
                    csvExpOpts.ReportSectionsOption = CsvExportSectionsOption.ExportIsolated;
                    csvExpOpts.SeparatorText = ",";
                    csvExpOpts.Delimiter = "\"";
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.CharacterSeparatedValues;
                    _reportDoc.ExportOptions.FormatOptions = csvExpOpts;
                }
                else if (_outputFormat.ToUpper() == "PDF")
                {
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;

                    var grpCnt = _reportDoc.DataDefinition.Groups.Count;

                    if (grpCnt > 0)
                        _reportDoc.ExportOptions.ExportFormatOptions = new PdfFormatOptions { CreateBookmarksFromGroupTree = true };
                }
                else if (_outputFormat.ToUpper() == "RPT")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.CrystalReport;
                else if (_outputFormat.ToUpper() == "DOC")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.WordForWindows;
                else if (_outputFormat.ToUpper() == "XLS")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.Excel;
                else if (_outputFormat.ToUpper() == "XLSDATA")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.ExcelRecord;
                else if (_outputFormat.ToUpper() == "XLSX")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.ExcelWorkbook;
                else if (_outputFormat.ToUpper() == "ERTF")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.EditableRTF;
                else if (_outputFormat.ToUpper() == "XML")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.Xml;
                else if (_outputFormat.ToUpper() == "HTM" || _outputFormat.ToUpper() == "HTML")
                {
                    HTMLFormatOptions htmlFormatOptions = new HTMLFormatOptions();

                    if (_outputFilename.LastIndexOf("\\") > 0) //if absolute output path is specified
                        htmlFormatOptions.HTMLBaseFolderName = _outputFilename.Substring(0, _outputFilename.LastIndexOf("\\"));

                    htmlFormatOptions.HTMLFileName = _outputFilename;
                    htmlFormatOptions.HTMLEnableSeparatedPages = false;
                    htmlFormatOptions.HTMLHasPageNavigator = true;
                    htmlFormatOptions.FirstPageNumber = 1;
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.HTML40;
                    _reportDoc.ExportOptions.FormatOptions = htmlFormatOptions;
                }
            }
        }

        /// <summary>
        /// Refresh Crystal Report if no input of parameters
        /// </summary>
        private void PerformRefresh()
        {
            bool toRefresh = ReportArguments.Refresh;
            bool noParameter = (_reportDoc.ParameterFields.Count == 0) ? true : false;

            if (toRefresh && noParameter)
                _reportDoc.Refresh();
        }

        /// <summary>
        /// Print or Export Crystal Report
        /// </summary>
        private void PerformOutput()
        {
            if (_printToPrinter)
            {
                var copy = ReportArguments.PrintCopy;
                _reportDoc.PrintToPrinter(copy, true, 0, 0);
                _logger.Write(string.Format("Report printed to : {0} - {1} Copies", _reportDoc.PrintOptions.PrinterName,copy));
            }
            else
            {
                _reportDoc.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;

                DiskFileDestinationOptions diskOptions = new DiskFileDestinationOptions();
                _reportDoc.ExportOptions.DestinationOptions = diskOptions;
                diskOptions.DiskFileName = _outputFilename;

                _reportDoc.Export();
                _logger.Write(string.Format("Report exported to : {0}", _outputFilename));

                if (ReportArguments.EmailOutput)
                {
                    using (MailMessage _MailMessage = new MailMessage())
                    {
                        _MailMessage.Attachments.Add(new Attachment(_outputFilename));
                        _MailMessage.From = new MailAddress(ReportArguments.MailFrom, ReportArguments.MailFromName);
                        _MailMessage.Subject = ReportArguments.EmailSubject;
                        if (ReportArguments.EmailBody != "NA")
                        {
                            _MailMessage.Body = ReportArguments.EmailBody;
                        }
                        _MailMessage.To.Add(ReportArguments.MailTo);
                        if (ReportArguments.MailCC != "NA")
                        {
                            _MailMessage.CC.Add(ReportArguments.MailCC);
                        }
                        if (ReportArguments.MailBcc != "NA")
                        {
                            _MailMessage.Bcc.Add(ReportArguments.MailBcc);
                        }
                        SmtpClient smtpClient = new SmtpClient();
                        smtpClient.Host = ReportArguments.SmtpServer;
                        smtpClient.Port = ReportArguments.SmtpPort;
                        smtpClient.EnableSsl = ReportArguments.SmtpSSL;

                        if (ReportArguments.SmtpUN != null && ReportArguments.SmtpPW != null)
                        {
                                //Uses Specified credentials to send email
                                smtpClient.UseDefaultCredentials = true;
                                System.Net.NetworkCredential credentials = new System.Net.NetworkCredential(ReportArguments.SmtpUN, ReportArguments.SmtpPW);
                                smtpClient.Credentials = credentials;
                            }
                        else
                            {
                                //If Set - uses the currently logged in user credentials to send email otherwise sent using Anonymous
                                smtpClient.UseDefaultCredentials = ReportArguments.SmtpAuth;
                            }
                        smtpClient.Send(_MailMessage);
                        _logger.Write(string.Format("Report {0} Emailed to : {1} CC'd to: {2} BCC'd to: {3}", _outputFilename, ReportArguments.MailTo, ReportArguments.MailCC, ReportArguments.MailBcc));
                        _logger.Write(string.Format("SMTP Details: Server:{0}, Port:{1}, SSL:{2} Auth:{3}, UN:{4}", smtpClient.Host, smtpClient.Port, smtpClient.EnableSsl, smtpClient.UseDefaultCredentials, ReportArguments.SmtpUN));
                    }

                    if (!ReportArguments.EmailKeepFile)
                    { File.Delete(_outputFilename); }
                }
            }
            Console.WriteLine("Completed");
        }

        /// <summary>
        /// Run the Crystal Reports Exporting or Printing process.
        /// </summary>
        public void Run()
        {
            try
            {
                _logger = new LogWriter(ReportArguments.LogFileName, ReportArguments.EnableLogToConsole);
                _logger.Write(string.Format(""));
                _logger.Write(string.Format("================== LoadReport, Output File and DB logion ==========="));
                LoadReport();
                ValidateOutputConfigurations();

                PerformDBLogin();
                ApplyReportOutput();
                _logger.Write(string.Format(""));
                _logger.Write(string.Format("================== ProcessParameters ==============================="));
                ProcessParameters();
                ProcessSelectionFormula();

                PerformRefresh();
                _logger.Write(string.Format(""));
                _logger.Write(string.Format("================== STEP = PerformOutput ============================"));
                PerformOutput();
            }
            catch (Exception ex)
            {
                _logger.Write(string.Format(""));
                _logger.Write(string.Format("===================Logs and Errors ================================="));
                _logger.Write(string.Format("Message: {0}", ex.Message));
                _logger.Write(string.Format("HResult: {0}", ex.HResult));
                _logger.Write(string.Format("Data: {0}", ex.Data));
				_logger.Write(string.Format("Inner Exception: {0}", ex.InnerException));
                _logger.Write(string.Format("StackTrace: {0}", ex.StackTrace));
                _logger.Write(string.Format("===================================================================="));
                throw;
            }
            finally
            {
                _reportDoc.Close();
            }
        }
    }
}
