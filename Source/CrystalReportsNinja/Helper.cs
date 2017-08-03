using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.Text;

namespace CrystalReportsNinja
{
    public static class Helper
    {
        /// <summary>
        /// Display help message
        /// </summary>
        /// <param name="mode"></param>
        public static void ShowHelpMessage()
        {
            Version v = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;

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

        /// <summary>
        /// Get Start Value of Range parameter
        /// </summary>
        /// <param name="parameterString"></param>
        /// <returns></returns>
        public static string GetStartValue(string parameterString)
        {
            int delimiter = parameterString.IndexOf(",");
            int leftbracket = parameterString.IndexOf("(");

            if (delimiter == -1 || leftbracket == -1)
                throw new Exception("Invalid Range Parameter value. eg. -a \"parameter name:(1000,2000)\"");

            return parameterString.Substring(leftbracket + 1, delimiter - 1).Trim();
        }

        /// <summary>
        /// Get End Value of a range parameter
        /// </summary>
        /// <param name="parameterString"></param>
        /// <returns></returns>
        public static string GetEndValue(string parameterString)
        {
            int delimiter = parameterString.IndexOf(",");
            int rightbracket = parameterString.IndexOf(")");

            if (delimiter == -1 || rightbracket == -1)
                throw new Exception("Invalid Range Parameter value. eg. -a \"parameter name:(1000,2000)\"");

            return parameterString.Substring(delimiter + 1, rightbracket - delimiter - 1).Trim();
        }

        /// <summary>
        /// Split multiple value parameter example "100|22|222" to 3 string in List
        /// </summary>
        /// <param name="multipleValueString"></param>
        /// <returns></returns>
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
    }
}
