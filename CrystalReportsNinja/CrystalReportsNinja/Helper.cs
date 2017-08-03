using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.Text;

namespace CrystalReportsNinja
{
    public static class Helper
    {
        public static void DisplayMessage(byte mode)
        {
            Version v = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;

            if (mode == 0)
            {

                Console.WriteLine("\nCrystal Reports Exporter Command Line Utility. Version {0}", v.ToString());
                Console.WriteLine("Copyright(c) 2011 Rainforest Software Solution http://www.rainforestnet.com");
            }
            else if (mode == 1)
            {
                Console.WriteLine("Type \"CrystalReportsNinja -?\" for help");
            }
            else if (mode == 2)
            {
                Console.WriteLine("\nCrystal Reports Ninja. Version {0}", v.ToString());
                Console.WriteLine("Copyright(c) 2017 Rainforest Software Solution http://www.rainforestnet.com");
                Console.WriteLine("CrystalReportsNinja Arguments Listing");
                Console.WriteLine("---------------------------------------------------");
                Console.WriteLine("-U database login username");
                Console.WriteLine("-P database login password");
                Console.WriteLine("-F Crystal reports path and filename (Mandatory)");
                Console.WriteLine("-S Database Server Name (instance name)");
                Console.WriteLine("-D Database Name");
                Console.WriteLine("-O Crystal reports Output path and filename");
                Console.WriteLine("-E Export file type.(pdf,doc,xls,rtf,htm,rpt,txt,csv...) If print to printer simply specify \"print\"");
                Console.WriteLine("-a Parameter value");
                Console.WriteLine("-N Printer Name (Network printer : \\\\PrintServer\\Printername or Local printer : printername)");
                Console.WriteLine("-C Number of copy to be printed");
                Console.WriteLine("-l To write a log file. filename ninja-yyyyMMddHHmmss.log");
                Console.WriteLine("---------------------------------------------------");
                Console.WriteLine("\nExample: C:\\> CrystalReportsNinja -U user1 -P mypass -S Server01 -D \"ExtremeDB\" -F c:\\test.rpt -O d:\\test.pdf -a \"Supplier Name:Active Outdoors\" -a \"Date Range:(12-01-2001,12-04-2002)\"");
                Console.WriteLine("Learn more in http://www.rainforestnet.com/crystal-reports-exporter/");
            }
        }

        public static string GetStartValue(string parameterString)
        {
            int delimiter = parameterString.IndexOf(",");
            int leftbracket = parameterString.IndexOf("(");

            if (delimiter == -1 || leftbracket == -1)
                throw new Exception("Invalid Range Parameter value. eg. -a \"parameter name:(1000,2000)\"");

            return parameterString.Substring(leftbracket + 1, delimiter - 1).Trim();
        }

        public static string GetEndValue(string parameterString)
        {
            int delimiter = parameterString.IndexOf(",");
            int rightbracket = parameterString.IndexOf(")");

            if (delimiter == -1 || rightbracket == -1)
                throw new Exception("Invalid Range Parameter value. eg. -a \"parameter name:(1000,2000)\"");

            return parameterString.Substring(delimiter + 1, rightbracket - delimiter - 1).Trim();
        }

        public static List<string> SplitIntoSingleValue(string multipleValueString)
        {
            //if "|" found,means multiple values parameter found
            int pipeStartIndex = 0;
            List<string> singleValue = new List<string>();
            bool loop = true; //loop false when it reaches the last parameter to read

            //pipeIndex is start search position of parameter string
            while (loop)
            {
                if (pipeStartIndex == multipleValueString.LastIndexOf("|") + 1)
                    loop = false;

                if (loop) //if this is not the last parameter
                    singleValue.Add(multipleValueString.Substring(pipeStartIndex, multipleValueString.IndexOf("|", pipeStartIndex + 1) - pipeStartIndex).Trim());
                else
                    singleValue.Add(multipleValueString.Substring(pipeStartIndex, multipleValueString.Length - pipeStartIndex).Trim());

                pipeStartIndex = multipleValueString.IndexOf("|", pipeStartIndex) + 1; //index to the next search of pipe
            }
            return singleValue;
        }

        public static void AddParameter(ref ParameterValues pValues, DiscreteOrRangeKind DoR, string inputString, string pName)
        {
            ParameterValue paraValue;
            if (DoR == DiscreteOrRangeKind.DiscreteValue || (DoR == DiscreteOrRangeKind.DiscreteAndRangeValue && inputString.IndexOf("(") == -1))
            {
                paraValue = new ParameterDiscreteValue();
                ((ParameterDiscreteValue)paraValue).Value = inputString;
                Console.WriteLine("Discrete Parameter : {0} = {1}", pName, ((ParameterDiscreteValue)paraValue).Value);

                pValues.Add(paraValue);
                paraValue = null;
            }
            else if (DoR == DiscreteOrRangeKind.RangeValue || (DoR == DiscreteOrRangeKind.DiscreteAndRangeValue && inputString.IndexOf("(") != -1))
            {
                paraValue = new ParameterRangeValue();
                ((ParameterRangeValue)paraValue).StartValue = GetStartValue(inputString);
                ((ParameterRangeValue)paraValue).EndValue = GetEndValue(inputString);
                Console.WriteLine("Range Parameter : {0} = {1} to {2} ", pName, ((ParameterRangeValue)paraValue).StartValue, ((ParameterRangeValue)paraValue).EndValue);

                pValues.Add(paraValue);
                paraValue = null;
            }
        }

        public static string ConvertStringArrayToString(string[] array)
        {
            //
            // Concatenate all the elements into a StringBuilder.
            //
            StringBuilder builder = new StringBuilder();
            foreach (string value in array)
            {
                builder.Append(value);
                builder.Append(' ');
            }
            return builder.ToString();
        }
    }
}
