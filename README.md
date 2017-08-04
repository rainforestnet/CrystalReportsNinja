# Crystal Reports Ninja
This is a complete rewritten based on Crystal-Reports-Exporter in order to enhance code readability and maintainability.

Crystal Reports Ninja is an open source Windows Console program runs on .NET Framework 4.5.
It loads external Crystal Report file (.rpt) and export into various file formats such as xls, pdf as well as print to printer.

Since it is a console (command-line) app, it can be invoked by Task Scheduler, batch file (.bat), command line file (.cmd), as well as Web API or Web Applications.

## Directory structure
* - Source (Source Code)
* - Deployment
	- CRforVS_redist_install_32bit_13_0_20.zip (runtime needed for 32-bit Windows machine)
	- CRforVS_redist_install_64bit_13_0_20.zip (runtime needed for 64-bit Windows machine)
	- CrystalReportsNinja.exe  (Executable compiled by Visual Studio 2017 for .NET Framework 4.5)

## Installation
Copy CrystalReportsNinja.exe and install the right CRforVS runtime of your platform and it should work.

## How to use
Locate the folder of CrystalReportsNinja.exe and run CrystalReportsNinja -? 
The only mandatory argument is "-F", in which is for user to specify a Crystal Reports file.

### List of arguments

* -F Crystal Reports filename to be loaded (i.e. "C:\Report Source\Report1.rpt") 
* -O Crystal Reports Output filename (i.e. "C:\Reports Output\Report1.pdf" ) [Optional]
* -E Intended file format to be exported.(i.e. pdf, doc, xls .. and etc). If you wish to print Crystal Reports to a printer, simply "-E print" instead of specifying file format.

* -N Printer Name. If printer name is not specified, it looks for default printer in the computer. If network printer, -N \\computer01\printer1
* -C Number of copy to be printed (any integer value i.e. 1,2,3..)
* -S Server Name for server where data resides. Only one server per Crystal Reports is allowed.
* -D Database Name. 
* -U Data source / server login username. Do not specify username and password for Integrated Security connection.
* -P Data source / server login password. Do not specify username and password for Integrated Security connection.
* -a Pass Crystal Reports file parameter set on run-time
* -l Create a log file into CrystalReportsNinja.exe directory

### -F, Crystal Reports source file
This is the only mandatory (must specify) argument, it allows your to specify the Crystal Reports filename to be exported.

### -O, Crystal Reports output file
If -O is not specified, Crystal Reports Exporter will just export the Crystal Reports into the same directory when Crystal Reports source file is resided.

The filename of the exported file is the same as Crystal Report source file and ending with timestamp (yyyyMMddHHmmss).

### -E, Crystal Reports Export File Type
There are 13 file formats that you can export a Crystal Reports file.

Use -E to specify desired file format. If -E argument is not supplied, CrystalReportsNinja will look into the output file extension, if output file is report1.pdf, it will then set to be Adobe PDF format. If file extension cannot be mapped into supported file format, CrystalReportsNinja will export your Crystal Reports into plain text file.

File type	Descriptions
-------------------------
txt		Plain text file
pdf		Adobe Portable Document format (pdf)
doc		Microsoft Word
xls		Microsoft Excel
xlsdata		Microsoft Excel with Data Only
rtf		Rich Text format
ertf	Editable Rich Text format
tab		Tab delimited text file
csv		Comma delimited text file
csv		Comma delimited text file
xml		xml file
htm		HTML file
rpt		Crystal Reports file
print	Print Crystal Report to a printer

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

### -a Passing in Parameters of Crystal Reports
We can have as many Parameters as we want in one Crystal Reports file. 
When we refresh a report, it prompts user to input the value of parameter in runtime in order to produce the result that user wants.

You can pass parameter value to CrystalReportsNinja.exe when it is executed.



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

