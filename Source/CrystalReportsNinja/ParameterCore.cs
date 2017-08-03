using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Collections.Generic;

namespace CrystalReportsNinja
{
    internal class UserParameter
    {
        public string ParameterName { get; set; }
        public string ParameterValue { get; set; }
    }

    public class ParameterCore
    {
        private List<UserParameter> _userParams;
        private LogWriter _logger;
        private List<string> _parameterCollection;

        public ParameterCore(string logfilename, List<string> paramCollection)
        {
            _userParams = new List<UserParameter>();

            _parameterCollection = paramCollection;
            _logger = new LogWriter(logfilename);
        }

        public void ProcessRawParameters()
        {
            foreach (string input in _parameterCollection)
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
        public ParameterValues GetParameterValues(ParameterFieldDefinition ParameterDef)
        {
            ParameterValues paramValues = new ParameterValues();
            //ParameterValue paramValue;

            for (int j = 0; j < _userParams.Count; j++)
            {
                if (ParameterDef.Name == _userParams[j].ParameterName)
                {
                    if (ParameterDef.EnableAllowMultipleValue && _userParams[j].ParameterValue.IndexOf("|") != -1)
                    {
                        // multiple value parameter
                        List<string> values = new List<string>();
                        values = Helper.SplitIntoSingleValue(_userParams[j].ParameterValue); //split multiple value into single value regardless discrete or range

                        for (int k = 0; k < values.Count; k++)
                        {
                            ParameterValue paramValue = GetSingleParamValue(ParameterDef.DiscreteOrRangeKind, values[k], ParameterDef.Name);
                            paramValues.Add(paramValue);
                        }
                    }
                    else
                    {
                        // simple single value parameter
                        ParameterValue paramValue = GetSingleParamValue(ParameterDef.DiscreteOrRangeKind, _userParams[j].ParameterValue, ParameterDef.Name);
                        paramValues.Add(paramValue);
                    }

                }
            }

            return paramValues;
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
        private ParameterValue GetSingleParamValue(DiscreteOrRangeKind paraType, string paraInputText, string paraName)
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
            }
            return null;
        }
    }
}
