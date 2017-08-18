Domain Reporter
 
Rethinkit Inc.
Chris Simotas 
christopher.simotas@tufts.edu
July 27, 2017

Version 1.1
Windows 10

Download whois from https://docs.microsoft.com/en-us/sysinternals/downloads/whois

Program searches and stores various properties of given domains from a whois program and reports back
an email, if it notices any changes or any upcoming expiration dates.

The program begins by importing a list of domains from the starter .csv file. The domains in the
.csv also have a description and any properties that the user doesn't care about if changes occur.
(For the Ignore Changes column, property names should be separated by comma). The name of the domain list
can be altered in an .xml file of the default values.

Next, the program then imports the default values from an xml file. The domain properties, a default value,
are the properties that user wants to search for. They can be adjusted in the DomainReporter xml file.
In some cases, the primary properties will be listed under other names. Thus, secondary or backup
property names can also be used. They can be edited in the xml file as well. Multiple terms can be
used under one primary property. 

Then, the program performs a whois on the domains and stores all
information in a cache .csv. In addition to the properties, The program also calculates how many days 
remain before the domain expires. After storing all the data, Domain Reporter compares the new information
against the old Domain information, previously stored. The program also assumes the occasional error from the whois program. 
Occassionally, the whois program will fail to return properties back. Thus, knowing this, Domain Reporter
will not overwrite NULL values over existing values from the last saved cache .csv file for the
specified potential error properties. The potential error properties are a default value and can be edited
in the xml default values file. In addition, Domain Reporter is built to withstand connection errors from the 
whois program. Moreover, the changes found and the upcoming expiring domains are then emailed to the email
address specified in the xml file. The changes html are saved in the Logs folder.

If the parameter -onlyReport is used, the program will only open a full
.csv file of all the domain information. The file is saved in the Reports folder
The program is built to be run on a schedule.

Notes:
  PowerShell.exe -Command "& 'C:\... \DomainReporter.ps1' -OnlyReport"
  or
  PowerShell.exe -Command "& 'C:\... \DomainReporter.ps1'"

 - False values have been entered in for a few of the default values. Please give true values
   to allow program to work.
 - To enable scripts, Run powershell 'as admin' then type Set-ExecutionPolicy Unrestricted
 - The program whois64.exe must be located in the same folder as the program
   source: https://docs.microsoft.com/en-us/sysinternals/downloads/whois
 - If a property is added, please be sure to add it to the switch statement in the importData function 
   if it can not be found using the default case (or using the findDesiredWord function)
 - If a property is added to the default values, be sure to enter NULL in its corresponding spot
   in the backupProperties array
 - The domain list must be a .csv file with three column headers: Domain, Description, Ignore Changes.
 - If you are planning on using a gmail or other main email providers, additional parameters will have to be added
   to the Send-MailMessage cmdlet in the sendReport function
 - Only compatibale with Windows 10 for now
 - Automatically accepts EULA for whois program (on a new machine)