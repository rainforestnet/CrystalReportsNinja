# Crystal Reports Ninja

Forked from Tengyongs Crystal Reports Ninja.  https://github.com/rainforestnet/CrystalReportsNinja

Crystal Reports Ninja is an open source Windows application.  The application loads and executes Crystal Report files exporting the results to a directoy or sending via email.

The Crystal Reports Ninja application can be invoked using Windows PowerShell or Windows Command Prompt.

## Supported File Formats for Export

<table class="table table-bordered table-condensed table-hover">
<thead>
<tr>
<th>File type</th>
<th>Descriptions</th>
</tr>
</thead>
<tbody>
<tr>
<td>txt</td>
<td>Plain text file</td>
</tr>
<tr>
<td>pdf</td>
<td>Adobe Portable Document format (pdf)</td>
</tr>
<tr>
<td>doc</td>
<td>Microsoft Word </td>
</tr>
<tr>
<td>xlsx</td>
<td>Microsoft Excel </td>
</tr>
<tr>
<td>xls</td>
<td>Microsoft Excel </td>
</tr>
<tr>
<td>xlsdata</td>
<td>Microsoft Excel with Data Only</td>
</tr>
<tr>
<td>rtf</td>
<td>Rich Text format</td>
</tr>
<tr>
<td>ertf</td>
<td>Editable Rich Text format</td>
</tr>
<tr>
<td>tab</td>
<td>Tab delimited text file </td>
</tr>
<tr>
<td>csv</td>
<td>Comma delimited text file</td>
</tr>
<tr>
<td>xml</td>
<td>xml file </td>
</tr>
<tr>
<td>htm</td>
<td>HTML file </td>
</tr>
<tr>
<td>rpt</td>
<td>Crystal Reports file </td>
</tr>
<tr>
<td>print</td>
<td>Print Crystal Report to a printer </td>
</tr>
</tbody>
</table>


## PreRequisites
* .NET Framework 4.5
* Crystal Reports Runtime 13.0.28 (or later).
	- If using the 64-bit Crystal Reports Ninja you must have the 64-bit Crystal Reports Runtime installed.  
	- If using the 32-bit Crystal Reports Ninja you must have the 32-bit Crystal Reports Runtime installed.  
	- Crystal Reports Runtime installation files can be downloaded from SAP using one of the following link.  
		http://www.crystalreports.com/crystal-reports-visual-studio/ 		
		https://wiki.scn.sap.com/wiki/display/BOBJ/Crystal+Reports%2C+Developer+for+Visual+Studio+Downloads
		https://origin.softwaredownloads.sap.com/public/site/index.html
	- It is suggested that you uninstall any previous versions of Crystal Reports Runtime before installing the latest version.

## Installation
* Copy all the files from either the 64-bit or 32-bit Deployment directories to a new local directory.

## How to use with Windows PowerShell (NEW FEATURE IN 1.4.0.0)

* Start by creating a new PowerShell script.  Using the PowerShell ISE to create the script is prefered, but any text editor will do.  There are a few examples of PowerShell scripts located in the Scripts directory.

* Below I have included the basic steps that need to be coded into the PowerShell script to be able to run Crystal Reports Ninja.
	- 1 Reference the Crystal Reports Ninja executable.
	- 2 Create a new instance of the ReportProcessor object.
	- 3 Set appropriate arguments for desired results.  The only argument that is required to be set is the ReportPath argument.
	- 4 Set any parameters required by the Crystal Report.
	- 5 Use the Report Processors Run method to have Crystal Reports Ninja begin its execution.

* Example 1.  Export results to Excel file using Report with Integrated Security.
```
[System.Reflection.Assembly]::LoadFile("C:\Program Files\CrystalReportsNinja\CrystalReportsNinja.exe")
$ReportProcessor = New-Object CrystalReportsNinja.ReportProcessor
$ReportProcessor.ReportArguments.EnableLogToConsole = $true
$ReportProcessor.ReportArguments.OutputFormat = "xlsx"
$ReportProcessor.ReportArguments.OutputPath = "c:\tmp\Test1"
$ReportProcessor.ReportArguments.ParameterCollection.Add("@DaysToConsider:90")
$ReportProcessor.ReportArguments.ParameterCollection.Add("@MaximumVariance:1")
$ReportProcessor.ReportArguments.ReportPath = "C:\CrystalReports\CrystalReport1.rpt"
$ReportProcessor.Run()
```

## How to use with Windows Command Prompt

Locate the folder of CrystalReportsNinja.exe and run CrystalReportsNinja -? 
The only mandatory argument is "-F", in which is for user to specify a Crystal Reports file.

### List of arguments

* ----------- Report and Database Config -----------;
* -F Crystal Reports Path and Filename.           (Mandatory);
* -S Database Server Name.                        (Optional if using Database Server set within the Crystal Report);
* -D Database Name.                               (Optional if using Database Server set within the Crystal Report);
* -U Database User Account.                       (Optional, If not set IntegratedSecurity is used);
* -P Database User Account Password.              (Optional, If not set IntegratedSecurity is used);
* ------------------ Output Config -----------------;
* -E Export Type.                                 (pdf,doc,xls,xlsx,rtf,htm,rpt,txt,csv,print);
* -O Output Path and Filename.                    (Not needed when printed);
* -N Printer Name.                                (Required when Printing);
* -C Number of copies to be printed.              (Optional, Default: 1);
* ------------- Parameters and Formulas -------------;
* -A Parameter value.;                            (Optional);
* -SF Crystal Reports Selection Forumula          (Optional);
*   ----------------- Logging Config -----------------;
* -L Write a log file.                            (Optional);
* -LC Write log output to console;                (Optional);
* ------------------ Email Config ------------------;
* -M  Email Report Output.                        (Enable Email Support);
* -MF Email Address to be SENT FROM.              (Optional, Default: noreply@noreply.com);
* -MN Email SENT FROM Display Name                (Optional, Default: Crystal Reports)
* -MT Email Address to SEND to.                   (Mandatory);
* -MC Email Address to be CC'd.                   (Optional);
* -MB Email Address to be BCC'd.                  (Optional);
* -MS Email Subject Line of the Email.            (Optional, Default: Crystal Reports);
* -MI Email Message Body.                         (Optional)
* -MSA SMTP server address.                        (Mandatory, if SSL enabled FQDN is required);
* -MSP SMTP server port.                           (Optional, Default: 25);
* -MSE SMTP server Enable SSL.                     (Optional, Default: False);
* -MSC SMTP Auth - Use Current User Credentials,   (Optional, Default: False );
* -MUN SMTP server Username.                      (Optional) \"domain\\username\";
* -MPW SMTP server Password.                      (Optional) \"password\";
* -MK Keep the generated file after sending email (Optional, Boolean True/False)

### -F, Crystal Reports source file
This is the only mandatory (must specify) argument, it allows your to specify the Crystal Reports filename to be exported.

### -S, Server name or Data Source name
Most of the time, you need not specify the server name or data source name of your Crystal Reports file as every Crystal Reports file saves data source information during design time.

Therefore, you only specify a server name when you want Crystal Reports to retrieve data from another data source that is different from what it is saved in Crystal Reports file during design time.

If your crystal Reports data source is one of the ODBC DSN (regardless of User DSN or System DSN), the DSN will be your Server Name (-S).

### -D, Database name
Similarly to Server Name. Database Name is completely optional. You only specify a Database Name when you intend to retrieve data from another database that is different from default database name that is saved in Crystal Reports file.

### -U and - P, Username and Password to login to Data source
If you are using database like Microsoft SQL server, MySQL or Oracle, you are most likely required to supply a username and password to login to database.

If you are using trusted connection to your SQL server, you wouldn't need to supply username and password when you run CrystalReportsNinja.

If you are using a password protected Microsoft Access database, you need only to specify password and not username.

### -E, Crystal Reports Export File Type
Use -E to specify desired file format. There are 13 file formats that you can export a Crystal Reports file.  If -E argument is not supplied, CrystalReportsNinja will look into the output file extension.  If output file is report1.pdf, it will then set to be Adobe PDF format. If file extension cannot be mapped into supported file format, CrystalReportsNinja will export your Crystal Reports into plain text file.

### -O, Crystal Reports output file
If -O is not specified, Crystal Reports Exporter will just export the Crystal Reports into the same directory when Crystal Reports source file is resided.

The filename of the exported file is the same as Crystal Report source file and ending with timestamp (yyyyMMddHHmmss).

### -N, Printer Name
Network printer : \\\\PrintServer\\Printername or Local printer : printername.

### -C, Copies to Print
Number of copies to print to printer.

### -A Passing in Parameters of Crystal Reports
Parameters within the Crystal Report that you would like to pass a value to.  Case must match.

### -L Write Log File
filename ninja-yyyyMMddHHmmss.log

### -LC Write Log Output to Console


### -SF, Crystal Reports Selection Formula
The Crystal Reports Record Selection Formula to use.  The result of the formula must be boolean (at least as far as I can tell).  You can find the field name to use by opening the Crystal Report, right clicking the field and selecting 'Find in Formulas'.  The field name will be displayed next to the "Matches for".

### -M Email Report


### -MF From Email Address
The email address that the email will appear to be from.  Defaults to CrystalReportsNinja@noreply.com.

### -MN From Email Address Display Nmae
The display name that the email will appear to be from.

### -MT To Email Address
The email address that the email will be sent to.  Mandatory for Email Report.

### -MS Subject
The text that will appear in the subject line of the email.  Defaults to Crystal Reports Ninja.

### -MI Email Message Body
The text that will appear in the body of the email.

### -MSA SMTP Server
SMTP server address.  Mandatory for Email Report.

### -MK Keep the generated file after sending by email
This is a boolean that will determine if the generated report should be kept after sending by email. Setting ```-MK True``` will keep the generated file after sending by email. 



###Example 1
Let's take an example of a Crystal Reports file test.rpt located in C: root directory. The Crystal Reports file has two parameters, namely Supplier and Date Range.
```
c:\>CrystalReportsNinja -U user1 -P mypass -S server1 -D Demo 
-F c:\test.rpt -O d:\test.pdf -E pdf -a "Supplier:Active Outdoors" 
-a "Date Range:(12-01-2001,12-04-2002)"
```

Rule to pass parameters
Use double quote (") to enclose the parameter string (parameter name and parameter value)
A colon (:) must be placed between parameter name and parameter value.
Use comma (,) to separate start value and end value of a Range parameter.
Use pipe (|) to pass multiple values into the same parameter.
Parameter name is case sensitive

###Example 2
The following crystal Reports "testTraining.rpt" has 5 discrete parameters and 2 range parameters.
Discrete parameters are:
* company
* customer
* branch

Range parameters are:
* Daterange
* Product

Example of Single value parameters :

```
c:\>CrystalReportsNinja -U user1 -P secret -F C:\Creport\Statement.rpt 
-O D:\Output\Statement.pdf -E pdf -S"AnySystem DB" 
-a "company:The One Computer,Inc" -a "customer:"MyBooks Store" 
-a "branch:Malaysia" -a "daterange:(12-01-2001,12-04-2002)"
-a "product:(AFD001,AFD005)"
```

Example of Multiple values parameters :
```
c:\>CrystalReportsNinja -F testreport.rpt -O testOutput.xls 
-E xlsdata -S order.mdb -a "date:(01-01-2001,28-02-2001)|(02-01-2002,31-10-2002)|(02-08-2002,31-12-2002)"
-a "Client:(Ace Soft Inc,Best Computer Inc)|(Xtreme Bike Inc,Zebra Design Inc)"
```

Example to print Crystal Reports output to printer
```
c:\>CrystalReportsNinja -F report101.rpt -E print -N "HP LaserJet 1200" -C 3
```

Email Report Example
```
c:\>CrystalReportsNinja -F Z:\CrystalReportsNinja\CrystalReports\ConsignaStoreInventoryValue.rpt -E pdf -O Z:\CrystalReportsNinja\Output\Test.pdf -a "@CustomerId:12345" -a "@Warehouse:987" -M -MF "Report1@company.com" -MT "good_user@company.com" -MS "Testing Ninja" -MSA "mail.company.com"
```

## Troubleshooting Ninja
Check the ensure that the version of Ninja (64bit\32bit) you are using matches the version of the ODBC Driver (64bit\32bit) you are using.  FYI, as of Oct 2019 the Crystal Reports Developer application is still 32bit.

Make sure that the option to save data in the Crystal Report is not enabled.  Users have reported issues when this option is enabled.

I had an issue with SQLs written by Crystal Reports working in Developer but not in Ninja.  The problem had to do with fully qualified table names.  To solve the problem I switched out the SQL Crystal had written and replaced it with a stored procedure which fixed the problem for me.  There are other solutions to this problem.  This webpage helped when I was searching for solutions to the problem.  https://wisdomofsolomon.wordpress.com/2011/06/18/crystal-reports-tables-not-found-during-verify-database/

## License
[MIT License](https://en.wikipedia.org/wiki/MIT_License)
