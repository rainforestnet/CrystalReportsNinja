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
            Console.WriteLine("\n     -------------- DB and Report Config --------------");
            Console.WriteLine("     -U database login username.                     (Optional, If not set IntegratedSecurity is used");
            Console.WriteLine("     -P database login password.                     (Optional, If not set IntegratedSecurity is used");
            Console.WriteLine("     -F Crystal reports path and filename.           (Mandatory)");
            Console.WriteLine("     -S Database Server Name.                        (instance name)");
            Console.WriteLine("     -D Database Name.");
            Console.WriteLine("\n     ------------------ Output Config -----------------");
            Console.WriteLine("     -O Output path and filename.");
            Console.WriteLine("     -E Export file type.                            (pdf,doc,xls,xlsx,rtf,htm,rpt,txt,csv...) If print to printer simply specify \"print\"");
            Console.WriteLine("     -a Parameter value.");
            Console.WriteLine("     -N Printer Name.                                (Network printer : \\\\PrintServer\\Printername or Local printer : printername)");
            Console.WriteLine("     -C Number of copy to be printed.");
            Console.WriteLine("\n     ----------------- Logging Config -----------------");
            Console.WriteLine("     -L To write a log file. filename ninja-yyyyMMddHHmmss.log");
            Console.WriteLine("     -LC To write log output to console");
            Console.WriteLine("\n     ------------------ Email Config ------------------");
            Console.WriteLine("     -M  Email Report Output.                        (Enable Email Support)");
            Console.WriteLine("     -MF Email Address to SEND From.                 (Optional, Default: noreply@noreply.com)");
            Console.WriteLine("     -MN Email Address to SEND From Display Name.    (Optional, Default: Crystal Reports)"); 
            Console.WriteLine("     -MT Email Address to SEND To.                   (Mandatory)");
            Console.WriteLine("     -MC Email Address to be CC'd.                   (Optional)");
            Console.WriteLine("     -MB Email Address to be BCC'd.                  (Optional)");
            Console.WriteLine("     -MS Email Subject Line of the Email.            (Optional, Default: Crystal Reports)");
            Console.WriteLine("     -MI Email Message Body.                         (Optional, Default: Blank)");
            Console.WriteLine("     -MK Email Keep Output File after sending.       (Optional, Default: False)");
            Console.WriteLine("     -MSA SMTP server address.                       (Mandatory, if SSL enabled FQDN is required)");
            Console.WriteLine("     -MSP SMTP server port.                          (Optional, Default: 25)");
            Console.WriteLine("     -MSE SMTP server Enable SSL.                    (Optional, Default: False");
            Console.WriteLine("     -MSC SMTP Auth - Use Current User Credentials,  (Optional, Default: False ");
            Console.WriteLine("     -MUN SMTP server Username.                      (Optional) \"domain\\username\"");
            Console.WriteLine("     -MPW SMTP server Password.                      (Optional) \"password\"");
            Console.WriteLine("     ----- SMTP Auth Note: If Username and Password is provided, -MSC Param is ignored and specifed Credentials are used to send email -----");
            Console.WriteLine("     ----- SMTP Auth Note: If -MSC + -MUN + -MPW Params are not set, Email is sent using Anonymous User -----");
            Console.WriteLine("     ----- SMTP Auth Note: If -MSC is set and -MUN + -MPW are not, Email is sent using Current User Credentials -----");
            Console.WriteLine("\n     ----------------- Example Scripts ----------------");
            Console.WriteLine("     Example: C:\\> CrystalReportsNinja -U user1 -P mypass -S Server01 -D \"ExtremeDB\" -F c:\\test.rpt -O d:\\test.pdf -a \"Supplier Name:Active Outdoors\" -a \"Date Range:(12-01-2001,12-04-2002)\"");
            Console.WriteLine("     Example Email : Add this to the above line, -M -MSA SMTP-FQDN -MSP 25 -MUN \"smtpUN\" -MPW \"smtpPW\" -MN \"Report User\" -MF noreply@noreply.com.au -MS \"Report Subject\" -MI \"Report Body\" -MN \"Report User\" -MF noreply@noreply.com.au -MT myemail@domain.com");
            Console.WriteLine("\n     Learn more in http://www.rainforestnet.com/crystal-reports-exporter/");
            
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
