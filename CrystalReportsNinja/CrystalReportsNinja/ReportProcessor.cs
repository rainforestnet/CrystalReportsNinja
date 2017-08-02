using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;

namespace CrystalReportsNinja
{
    internal class UserParameter
    {
        public string ParameterName { get; set; }
        public string ParameterValue { get; set; }
    }

    public class ReportProcessor
    {
        private string _sourceFilename;
        private string _outputFilename;
        private string _outputFormat;
        private bool _printToPrinter;

        private ReportDocument _reportDoc;
        private List<UserParameter> _userParams;
        private LogWriter _logger;

        public ArgumentContainer ReportArguments { get; set; }

        public ReportProcessor(string logfilename)
        {
            _reportDoc = new ReportDocument();
            _userParams = new List<UserParameter>();

            _logger = new LogWriter(logfilename);
            _logger.Write("instantiated");
        }

        private void ProcessRawParameters()
        {
            foreach (string input in ReportArguments.ParameterCollection)
            {
                _userParams.Add(new UserParameter
                {
                    ParameterName = input.Substring(0, input.IndexOf(":")).Trim(),
                    ParameterValue = (input.Substring(input.IndexOf(":") + 1, input.Length - (input.IndexOf(":") + 1))).Trim(),
                });
            }
        }

        /// <summary>
        /// A report can contains more than One parameters, hence
        /// we loop through all parameters that user has input and match
        /// it with parameter definition of Crystal Reports file.
        /// </summary>
        private void ApplyingParameters()
        {
            if (_reportDoc.ParameterFields.Count > 0)
            {
                ParameterFieldDefinitions paramDefs = _reportDoc.DataDefinition.ParameterFields;
                ParameterValues paramValues = new ParameterValues();
                ParameterValue paramValue;

                for (int i = 0; i < paramDefs.Count; i++)
                {
                    for (int j = 0; j < _userParams.Count; j++)
                    {
                        if (paramDefs[i].Name == _userParams[j].ParameterName)
                        {
                            if (paramDefs[i].EnableAllowMultipleValue && _userParams[j].ParameterValue.IndexOf("|") != -1) 
                            {
                                // For multiple value parameter
                                List<string> values = new List<string>();
                                values = Helper.SplitIntoSingleValue(_userParams[j].ParameterValue); //split multiple value into single value regardless discrete or range

                                for (int k = 0; k < values.Count; k++)
                                {
                                    paramValue = GetParamValue(paramDefs[i].DiscreteOrRangeKind, values[k], paramDefs[i].Name);
                                    paramValues.Add(paramValue);
                                }
                            }
                            else 
                            {
                                // For simple single value parameter
                                paramValue = GetParamValue(paramDefs[i].DiscreteOrRangeKind, _userParams[j].ParameterValue, paramDefs[i].Name);
                                paramValues.Add(paramValue);
                            }
                            paramDefs[i].ApplyCurrentValues(paramValues);
                        }
                    }
                }
            }
        }

        // Load external Crystal Report file into Report Document
        private void LoadReport()
        {
            _sourceFilename = ReportArguments.ReportPath.Trim();
            if (_sourceFilename == null || _sourceFilename == string.Empty)
            {
                throw new Exception("Invalid Crystal Reports file");
            }

            if (_sourceFilename.LastIndexOf(".rpt") == -1)
                throw new Exception("Invalid Crystal Reports file");

            _reportDoc.Load(_sourceFilename, OpenReportMethod.OpenReportByDefault);
            _logger.Write(string.Format("Report loaded successfully"));
        }

        private void SetReportOutput()
        {
            _outputFilename = ReportArguments.OutputPath;
            _outputFormat = ReportArguments.OutputFormat;
            _printToPrinter = ReportArguments.PrintOutput;

            if (!_printToPrinter)
            {
                string fileExt = "";

                //default set to text file
                if (_outputFormat == null && _outputFilename == null)
                    _outputFormat = "txt";

                //if output path isn't specified but there is a output format
                //use output format to set output path
                if (_outputFilename == null && _outputFormat != null)
                {
                    if (_outputFormat.ToUpper() == "XLSDATA")
                        fileExt = "xls";
                    else if (_outputFormat.ToUpper() == "TAB")
                        fileExt = "txt";
                    else if (_outputFormat.ToUpper() == "ERTF")
                        fileExt = "rtf";
                    else
                        fileExt = _outputFormat;
                }
                _outputFilename = String.Format("{0}-{1}.{2}", _sourceFilename.Substring(0, _sourceFilename.LastIndexOf(".rpt")), DateTime.Now.ToString("yyyyMMddHHmmss"), fileExt);
                _logger.Write(string.Format("Output Filename : {0}", _outputFilename));

                if (_outputFormat == null && _outputFilename != null)
                {
                    int lastIndexDot = _outputFilename.LastIndexOf(".");
                    fileExt = _outputFilename.Substring(lastIndexDot + 1, 3);

                    //ensure filename extension has 3 char after the dot (.)
                    if ((_outputFilename.Length == lastIndexDot + 4) && (fileExt.ToUpper() == "RTF" || fileExt.ToUpper() == "TXT" || fileExt.ToUpper() == "CSV" || fileExt.ToUpper() == "PDF" || fileExt.ToUpper() == "RPT" || fileExt.ToUpper() == "DOC" || fileExt.ToUpper() == "XLS" || fileExt.ToUpper() == "XML" || fileExt.ToUpper() == "HTM"))
                        _outputFormat = _outputFilename.Substring(lastIndexDot + 1, 3);
                }
                _logger.Write(string.Format("Output format : {0}", _outputFormat));
            }
        }

        private void ApplyDBLogin()
        {
            bool toRefresh = ReportArguments.Refresh;

            var server = ReportArguments.ServerName;
            var database = ReportArguments.DatabaseName;
            var username = ReportArguments.UserName;
            var password = ReportArguments.Password;

            if (toRefresh)
            {
                TableLogOnInfo logonInfo = new TableLogOnInfo();
                foreach (Table table in _reportDoc.Database.Tables)
                {
                    if (server != null)
                    {
                        logonInfo.ConnectionInfo.ServerName = server;
                        _logger.Write(string.Format("Logon to Server = {0}", server));
                    }

                    if (database != null)
                    {
                        logonInfo.ConnectionInfo.DatabaseName = database;
                        _logger.Write(string.Format("Logon to Database = {0}", database));
                    }

                    if (username == null && password == null)
                    {
                        logonInfo.ConnectionInfo.IntegratedSecurity = true;
                        _logger.Write(string.Format("Integrated Security = true"));
                    }
                    else
                    {
                        if (username != null && username.Length > 0)
                        {
                            logonInfo.ConnectionInfo.UserID = username;
                            _logger.Write(string.Format("Logon with user id = {0}", username));
                        }

                        if (password == null) //to support blank password
                        {
                            logonInfo.ConnectionInfo.Password = "";
                            _logger.Write(string.Format("Logon with blank password"));
                        }
                        else
                        {
                            logonInfo.ConnectionInfo.Password = password;
                            _logger.Write(string.Format("Logon with password = {0}", password));
                        }
                    }
                    table.ApplyLogOnInfo(logonInfo);
                }
            }
        }

        /// <summary>
        /// Extract value from a raw parameter string.
        /// "paraInputText" must be just a single value and not a multiple value parameters
        /// </summary>
        /// <param name="paraType"></param>
        /// <param name="paraInputText"></param>
        /// <param name="paraName"></param>
        /// <remarks>
        /// Complex parameter input can be as such, 
        /// pipe "|" is used to split multiple values, comma "," is used to split Start and End value of a range.
        /// 
        /// -a "date:(01-01-2001,28-02-2001)|(02-01-2002,31-10-2002)|(02-08-2002,31-12-2002)" 
        /// -a "Client:(Ace Soft Inc,Best Computer Inc)|(Xtreme Bike Inc,Zebra Design Inc)"
        /// </remarks>
        /// <returns></returns>
        private ParameterValue GetParamValue(DiscreteOrRangeKind paraType, string paraInputText, string paraName)
        {
            ParameterValues paraValues = new ParameterValues();
            bool isDiscreateType = paraType == DiscreteOrRangeKind.DiscreteValue ? true : false;
            bool isDiscreateAndRangeType = paraType == DiscreteOrRangeKind.DiscreteAndRangeValue ? true : false;
            bool isRangeType = paraType == DiscreteOrRangeKind.RangeValue ? true : false;
            bool paraTextIsRange = paraInputText.IndexOf("(") != -1 ? true : false;

            if (isDiscreateType || (isDiscreateAndRangeType && !paraTextIsRange))
            {
                var paraValue = new ParameterDiscreteValue()
                {
                    Value = paraInputText
                };
                _logger.Write(string.Format("Discrete Parameter : {0} = {1}", paraName, ((ParameterDiscreteValue)paraValue).Value));
                return paraValue;
                //paraValues.Add(paraValue);
            }
            else if (isRangeType || (isDiscreateAndRangeType && paraTextIsRange))
            {
                // sample of range parameter (01-01-2001,28-02-2001)
                var paraValue = new ParameterRangeValue()
                {
                    StartValue = Helper.GetStartValue(paraInputText),
                    EndValue = Helper.GetEndValue(paraInputText)
                };
                _logger.Write(string.Format("Range Parameter : {0} = {1} to {2} ", paraName, ((ParameterRangeValue)paraValue).StartValue, ((ParameterRangeValue)paraValue).EndValue));
                return paraValue;
                //paraValues.Add(paraValue);
            }
            return null;
            //return paraValues;
        }

        private void ApplyReportOutput()
        {
            if (_printToPrinter)
            {
                var printerName = ReportArguments.PrinterName != null ? ReportArguments.PrinterName.Trim() : "";

                if (printerName.Length > 0)
                {
                    _reportDoc.PrintOptions.PrinterName = printerName;
                }
                else
                {
                    System.Drawing.Printing.PrinterSettings prinSet = new System.Drawing.Printing.PrinterSettings();

                    if (prinSet.PrinterName.Trim().Length > 0)
                        _reportDoc.PrintOptions.PrinterName = prinSet.PrinterName;
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
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.CharacterSeparatedValues;
                else if (_outputFormat.ToUpper() == "PDF")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                else if (_outputFormat.ToUpper() == "RPT")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.CrystalReport;
                else if (_outputFormat.ToUpper() == "DOC")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.WordForWindows;
                else if (_outputFormat.ToUpper() == "XLS")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.Excel;
                else if (_outputFormat.ToUpper() == "XLSDATA")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.ExcelRecord;
                else if (_outputFormat.ToUpper() == "ERTF")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.EditableRTF;
                else if (_outputFormat.ToUpper() == "XML")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.Xml;
                else if (_outputFormat.ToUpper() == "HTM")
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

        private void DoRefresh()
        {
            bool toRefresh = ReportArguments.Refresh;
            bool noParameter = (_reportDoc.ParameterFields.Count == 0) ? true : false;

            if (toRefresh && noParameter)
                _reportDoc.Refresh();
        }

        private void DoOutput()
        {
            if (_printToPrinter)
            {
                var copy = ReportArguments.PrintCopy;
                _reportDoc.PrintToPrinter(copy, true, 0, 0);
                _logger.Write(string.Format("Report printed to : {0}", _reportDoc.PrintOptions.PrinterName));
            }
            else
            {
                _reportDoc.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;

                DiskFileDestinationOptions diskOptions = new DiskFileDestinationOptions();
                _reportDoc.ExportOptions.DestinationOptions = diskOptions;
                diskOptions.DiskFileName = _outputFilename;

                _reportDoc.Export();
                _logger.Write(string.Format("Report exported to : {0}", _outputFilename));
            }
        }

        //private void WonderWoman()
        //{
        //    foreach (string input in ReportArguments.ParameterCollection)
        //    {
        //        reportParameterName.Add(input.Substring(0, input.IndexOf(":")).Trim());
        //        reportParameterValue.Add((input.Substring(input.IndexOf(":") + 1, input.Length - (input.IndexOf(":") + 1))).Trim());
        //    }
        //}

        //private void Spiderman()
        //{
        //    if (_reportDoc.ParameterFields.Count > 0)
        //    {
        //        ParameterFieldDefinitions paramDefs = _reportDoc.DataDefinition.ParameterFields;
        //        ParameterValues paramValues = new ParameterValues();
        //        List<string> singleParamValue = new List<string>();

        //        for (int i = 0; i < paramDefs.Count; i++)
        //        {
        //            for (int j = 0; j < reportParameterName.Count; j++)
        //            {
        //                if (paramDefs[i].Name == reportParameterName[j])
        //                {
        //                    _logger.Write(string.Format("{0} : {1}", reportParameterName[j], reportParameterValue[j]));
        //                    if (paramDefs[i].EnableAllowMultipleValue && reportParameterValue[j].IndexOf("|") != -1)
        //                    {
        //                        singleParamValue = Helper.SplitIntoSingleValue(reportParameterValue[j]); //split multiple value into single value regardless discrete or range

        //                        for (int k = 0; k < singleParamValue.Count; k++)
        //                            AddParameter(ref paramValues, paramDefs[i].DiscreteOrRangeKind, singleParamValue[k], paramDefs[i].Name);

        //                        singleParamValue.Clear();
        //                    }
        //                    else
        //                        AddParameter(ref paramValues, paramDefs[i].DiscreteOrRangeKind, reportParameterValue[j], paramDefs[i].Name);

        //                    paramDefs[i].ApplyCurrentValues(paramValues);
        //                    paramValues.Clear();

        //                    break; //jump into another user input parameter
        //                }
        //            }

        //            //for (int j = 0; j < _userParams.Count; j++)
        //            //{
        //            //    if (paramDefs[i].Name == _userParams[j].ParameterName)
        //            //    {
        //            //        if (paramDefs[i].EnableAllowMultipleValue && _userParams[j].ParameterValue.IndexOf("|") != -1)
        //            //        {
        //            //            singleParamValue = Helper.SplitIntoSingleValue(_userParams[j].ParameterValue); //split multiple value into single value regardless discrete or range

        //            //            for (int k = 0; k < singleParamValue.Count; k++)
        //            //                AddParameter(ref paramValues, paramDefs[i].DiscreteOrRangeKind, singleParamValue[k], paramDefs[i].Name);

        //            //            singleParamValue.Clear();
        //            //        }
        //            //        else
        //            //            AddParameter(ref paramValues, paramDefs[i].DiscreteOrRangeKind, _userParams[j].ParameterValue, paramDefs[i].Name);

        //            //        paramDefs[i].ApplyCurrentValues(paramValues);
        //            //        paramValues.Clear();

        //            //        break; //jump into another user input parameter
        //            //    }
        //            //}
        //        }
        //    }
        //}

        //private void AddParameter(ref ParameterValues pValues, DiscreteOrRangeKind DoR, string inputString, string pName)
        //{
        //    ParameterValue paraValue;
        //    if (DoR == DiscreteOrRangeKind.DiscreteValue || (DoR == DiscreteOrRangeKind.DiscreteAndRangeValue && inputString.IndexOf("(") == -1))
        //    {
        //        paraValue = new ParameterDiscreteValue();
        //        ((ParameterDiscreteValue)paraValue).Value = inputString;
        //        Console.WriteLine("Discrete Parameter : {0} = {1}", pName, ((ParameterDiscreteValue)paraValue).Value);

        //        //if (enableLog)
        //        //    WriteLog("Discrete Parameter : " + pName + " = " + ((ParameterDiscreteValue)paraValue).Value);

        //        pValues.Add(paraValue);
        //        paraValue = null;
        //    }
        //    else if (DoR == DiscreteOrRangeKind.RangeValue || (DoR == DiscreteOrRangeKind.DiscreteAndRangeValue && inputString.IndexOf("(") != -1))
        //    {
        //        paraValue = new ParameterRangeValue();
        //        ((ParameterRangeValue)paraValue).StartValue = Helper.GetStartValue(inputString);
        //        ((ParameterRangeValue)paraValue).EndValue = Helper.GetEndValue(inputString);
        //        Console.WriteLine("Range Parameter : {0} = {1} to {2} ", pName, ((ParameterRangeValue)paraValue).StartValue, ((ParameterRangeValue)paraValue).EndValue);

        //        //if (enableLog)
        //        //    WriteLog("Range Parameter : " + pName + " = " + ((ParameterRangeValue)paraValue).StartValue + " to " + ((ParameterRangeValue)paraValue).EndValue);

        //        pValues.Add(paraValue);
        //        paraValue = null;
        //    }
        //}

        // Run Report
        public void Run()
        {
            try
            {
                LoadReport();
                SetReportOutput();
                ProcessRawParameters();

                ApplyDBLogin();
                ApplyReportOutput();
                ApplyingParameters();

                DoRefresh();
                DoOutput();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                _logger.Write(ex.Message);
            }
            finally
            {
                _reportDoc.Close();
            }
        }
    }
}
