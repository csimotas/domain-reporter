# # # # # # # # # # # # # # # # # # # # # # # # # # #
#
# Domain Reporter
# 
# Rethinkit Inc.
# Chris Simotas 
# christopher.simotas@tufts.edu
# October 17, 2017
#
# Version 1.2
# Windows 10
#
### README ###
# Program searches and stores various properties of given domains from a whois program and reports back
# an email, if it notices any changes or any upcoming expiration dates.
#
# The program begins by importing a list of domains from the starter .csv file. The domains in the
# .csv also have a description and any properties that the user doesn't care about if changes occur.
# (For the Ignore Changes column, property names should be separated by comma). The name of the domain list
# can be altered in an .xml file of the default values.
#
# Next, the program then imports the default values from an xml file. The domain properties, a default value,
# are the properties that user wants to search for. They can be adjusted in the DomainReporter xml file.
# In some cases, the primary properties will be listed under other names. Thus, secondary or backup
# property names can also be used. They can be edited in the xml file as well. Multiple terms can be
# used under one primary property. 
#
# Then, the program performs a whois on the domains and stores all
# information in a cache .csv. In addition to the properties, The program also calculates how many days 
# remain before the domain expires. After storing all the data, Domain Reporter compares the new information
# against the old Domain information, previously stored. The program also assumes the occasional error from the whois program. 
# Occassionally, the whois program will fail to return properties back. Thus, knowing this, Domain Reporter
# will not overwrite NULL values over existing values from the last saved cache .csv file for the
# specified potential error properties. The potential error properties are a default value and can be edited
# in the xml default values file. In addition, Domain Reporter is built to withstand connection errors from the 
# whois program. Moreover, the changes found and the upcoming expiring domains are then emailed to the email
# address specified in the xml file. The changes html are saved in the Logs folder.
#
# If the parameter -onlyReport is used, the program will only open a full
# .csv file of all the domain information. The file is saved in the Reports folder
# The program is built to be run on a schedule.
#
# Notes:
#   PowerShell.exe -Command "& 'C:\... \DomainReporter.ps1' -OnlyReport"
#   or
#   PowerShell.exe -Command "& 'C:\... \DomainReporter.ps1'"
#
# - False values have been entered in for a few of the default values. Please give true values
#   to allow program to work.
# - To enable scripts, Run powershell 'as admin' then type Set-ExecutionPolicy Unrestricted
# - The program whois64.exe must be located in the same folder as the program
#   source: https://docs.microsoft.com/en-us/sysinternals/downloads/whois
# - If a property is added, please be sure to add it to the switch statement in the importData function 
#   if it can not be found using the default case (or using the findDesiredWord function)
# - If a property is added to the default values, be sure to enter NULL in its corresponding spot
#   in the backupProperties array
# - The domain list must be a .csv file with three column headers: Domain, Description, Ignore Changes.
# - If you are planning on using a gmail or other main email providers, additional parameters will have to be added
#   to the Send-MailMessage cmdlet in the sendReport function
# - Only compatibale with Windows 10 for now
# - Automatically accepts EULA for whois program (on a new machine)
#
#
# Functions are written below followed by the main script at the end
# # # # # # # # # # # # # # # # # # # # # # # # # # #


Param([switch]$onlyReport)

###################################################################################
## Function: pathNoExt
## Purpose: Full path without the extension
Function pathNoExt ($file) 
{
    $file.Substring(0, $file.LastIndexOf('.'))
    return
}

###################################################################################
## Function: Pause
## Purpose: Pauses program until key is pressed
Function pause ($M="Press any key to continue . . . ")
{
    If($psISE)
    {
        $S=New-Object -ComObject "WScript.Shell"
        $B=$S.Popup("Click OK to continue.",0,"Script Paused",0)
        Return
    }
	Else
	{
		Write-Host -NoNewline $M
		$I=16,17,18,20,91,92,93,144,145,166,167,168,169,170,171,172,173,174,175,176,177,178,179,180,181,182,183
		While ($K.VirtualKeyCode -Eq $Null -Or $I -Contains $K.VirtualKeyCode)
		{
			$K=$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
		}
		Write-Host "(OK)"
	}
}

###################################################################################
## Function: globalsSave
## Purpose: Saves an xml to be used for global variables
Function globalsSave
{
    $GlobalsXML = (pathNoExt($script:MyInvocation.MyCommand.Path)) + ".xml"
    Export-Clixml -InputObject $Globals -Path $GlobalsXML
    return
}

###################################################################################
## Function: globalsLoad
## Purpose: Loads an xml of global variables and if sml doesn't exists, it pulls from a default list
Function globalsLoad
{
    $GlobalsXML = (pathNoExt($script:MyInvocation.MyCommand.Path)) + ".xml"
    if (Test-Path($GlobalsXML))
    {
        $script:Globals = Import-Clixml -Path $GlobalsXML
    }
    else
   	{
    	Write-Host "Couldn't find settings file ($GlobalsXML).  Creating a new file with default values."
    	globalsInit
    }
    return
}

## *** DEFAULT VALUES *** ##

###################################################################################
## Function: globalsIntit 
## Purpose: Sets all the global variables for the xml file. The variables include
##          mailServer - the server for the email
##          mailFrom - the outgoing email address
##          mailTo - the email address the email is being sent to
##          reportFilename - the filename of the report
##          domainList - the csv list of all the domains to check
##          daysLimit - the number of days left till expiration when the user should be notified
##          prop - the array of all the properties being checked in installed program
##                 check .\whois664.exe -v domain to see of possible properties
##          bprop - the secondary/backup property terms to check if the primary 
##                  property term doesn't exist
##          eprop - properties with potential errors where whois returns occassional blanks for
Function globalsInit
{
    $script:Globals = @{}
    #### Values used to create a missing settings file
    #### Note: - do not edit these values here.  Edit 'CheckDNS.xml' settings file instead.
    ####         the program will purposely fail
    ####       - each value in the backupProperties array can be an array

    $Globals.Add("mailServer", "smtp2.server.com")
    $Globals.Add("mailFrom", "domainreporter@server.com")
    $Globals.Add("mailTo", "chris@server.com")
    $Globals.Add("reportFilename", "Domain_Cache.csv")
    $Globals.Add("domainList", "DomainsToCheckPublic.csv") 
    $Globals.Add("daysLimit", 30)

    # when you add a value to properties, be sure to add a value or NULL in the correct place in backupProperties as well
    #
    # [System.Collections.ArrayList] allows items to be removed from the array
    $prop = [System.Collections.ArrayList]("Domain", "Description", "Ignore Changes", "Registrar", "Mail Exchange", "Name Server", "WWW Record", "Autodiscover Record", "Registrant Name", "Registrant Email", "Updated Date", "Creation Date", "Expiration Date", "Days Till Expiration")
    $Globals.Add("properties", $prop)

    # when you add a value to properties, be sure to add a value or NULL in the correct place in backupProperties as well
    $bprop = ($NULL, $NULL, $NULL, "Sponsoring Registrar", $NULL, $NULL, $NULL, $NULL, "Registrant Organization", $NULL, $NULL, $NULL, ("Registry Expiry Date", "Registrar Registration Expiration Date"), $NULL)
    $Globals.Add("backupProperties", $bprop)
    
    # properties with potential errors (whois returns occassional blanks for them)
    $eprop = ("Registrant Name", "Registrant Email")
    $Globals.Add("propertiesWithPotError", $eprop) 
    ######
    $GlobalsXML = (pathNoExt($script:MyInvocation.MyCommand.Path)) + ".xml"
    Export-Clixml -InputObject $Globals -Path $GlobalsXML
    return
}


###################################################################################
## Function: setupObject
## Input: properties - the array of property names
## Output: domainObject - the object of ther domain
## Purpose: Creates an "empty" object for the domain with property members
Function setupObject ($properties)
{
    $domainObject = new-object PSObject

    foreach($p in $properties)
    {
        $domainObject | add-member -type NoteProperty -name $p -Value $NULL
    }
    return $domainObject
}

###################################################################################
## Function: findLineContaining
## Input: word - desired word user is searching for
##        arrstring - the array of string consisting of several lines
##        supress - boolean that indicates to supress the "No match" warning
##        searchWholeLine - boolean that indicates whether to search the entire line or not
##        returnWholeLine - boolean that indicates whether to return the whole line or not
## Output: the portion of the line containing the word
## Purpose: Searches in the script for the desired line based on beginning word
Function findLineContaining 
{
    param 
        (
        [string]   $word = '' , 
        [string[]] $arrstring = @("black","white","yellow","blue") ,
        [boolean]  $supress = $false , 
        [boolean]  $searchWholeLine = $false, #true = only lines that begin with word,  false = anywhere in line
        [boolean]  $returnWholeLine = $false  #true = whole line is returned, false = only portion at end of line
        )
 
    # $stringSplit = ($arrstring.Trim() -split '[\r\n]') |? {$_}
    # trims all indentation splits string by line and 
    # search each line
    $done = $false
    foreach ($line in $arrstring)
    {
        $trimline = $line.Trim()
        # checks to see if desired word matches the word in the line
        if ($searchWholeLine)
        {
            if($trimline.contains($word))
            {
                $done = $true;
                break
            }  
        }
        else
        {
            if($trimline.StartsWith($word))
            {
                $done = $true;
                break
            }     
        }
        
    }
    if($done)
    { # found a match find the position of the 1st instance of the word (works either way)
        if ($returnWholeLine)
        {
            $data=$line
        }
        else
        {
            $pos = $line.indexof($word,[System.StringComparison]::CurrentCultureIgnoreCase)
            # return portion of line after word
            $data = $line.Substring($pos+$word.length)
        }
    }
    else
    { # no match
        $data = $NULL
        if (!($supress))
        { # show an error
            Write-Host "No Word Match Found" -ForegroundColor Red
        }

    }
    # returns the info
    return $data
}

###################################################################################
## Function: importData
## Input: folder - current folder of the program
##        entry - the entry domain being worked on
##        properties - the array of properties
##        backupProp - the array of backup properties names if the main property names can not be found
## Output: output - an object of domainObject (the object of the domain containing all its information)
##                  and failedDom (a string of the Domain if it failed) 
## Purpose: Imports data/information from whois and powershell about a domain and stores it into an object.
Function importData ($folder, $entry, $properties, $backupProp)
{
    # runs the whois64 program
    $importData = & ($folder + "\whois64.exe") -v $entry.domain -nobanner
    
    # if the connection with the whois server fails, the properties of the domain will be left NULL and later
    # filled in with values from the old data set. The output will be less than 6 lines with no domain information
    # if the connection failed.
     
    if ($importData.count -gt 6)
    {
        # checks for other whois servers to obtain more information
        $serverTerms = @("Registrar WHOIS Server", "WHOIS Server")
        foreach($term in $serverTerms)
        {
            $whoIsServ = FindLineContaining -word ($term + ": ") -arrstring $importData -supress $true
            if(-NOT [string]::IsNULLOrEmpty($whoIsServ))
            {
                $importData += & ($folder + "\whois64.exe") -v $entry.domain $whoIsServ -nobanner
                break
            }
        }

        $object = setupObject -properties $properties
     
        # loops through all the properties and uses FindLineContaining to find the data being searched
        # skips over the properties of Domain, Mail Exchange, WWW Record, Autodiscover Record, and Days Till Expiration since those
        # do not require the whois64 to find values

        foreach($p in $properties)
        {
            # switch for all the properties 
            # ** If property is added, please add it to below if it wouldn't be found under default **
            switch ($p)
            {
                "Mail Exchange" 
                {
                    # uses built-in Powershell cmdlet to find Mail Exchange, takes highest preference
                    $d = Resolve-DnsName -Name ($entry.domain) -Type MX
                    
                    # incase an error occurs with Resolve-DnsName. It will skip over following lines
                    $m = $d
                    if($d -ne $NULL)
                    {
                        $r = $d | Sort-Object -Property "Preference", "NameExchange"
                        $m =$r[0].NameExchange
                    }
    
                    $object.$p= $m
                }

                "Description"
                {
                    $object.$p = $entry.Description
                }
    
                "Ignore Changes"
                {
                    $object.$p = $entry."Ignore Changes"
                }
    
                "Domain"
                {
                    $object.$p = $entry.Domain
                } 
    
                "Days Till Expiration"
                {
                    # creates an array of the possible terms for the property Expiration Date based off of the backup terms
                    $b = [array]::IndexOf($properties, "Expiration Date")
                    $possibleOpt = @()
                    $possibleOpt += "Expiration Date"
                    $possibleOpt += $backupProp[$b]
                    # removes all NULL values
                    $possibleOpt = ($possibleOpt.Where({ $_ -ne $NULL }))
                
                    # runs through each option until it gets a hit
                    foreach($opt in $possibleOpt)
                    {
                        $expDate = FindLineContaining -word ($opt + ": ") -arrstring $importData -supress $true
                         
                        if($expDate-ne $NULL)
                        {
                            break
                        }
                    }

                    # formats the date string if necessary
                    $ind = ($expDate).IndexOf('T')
                    if($ind -ne -1) 
                    {
                        $expDate = ($expDate).Substring(0, $ind)
                    } 

                    # calculates the number of days left till expiration using New-TimeSpan and Get-Date cmdlet
                    $object.$p = (New-TimeSpan -start (Get-Date) -end $expDate).Days + 1
                }


                {($_ -eq "WWW Record") -or ($_ -eq "Autodiscover Record")}
                {
                    # since not every domain will have a WWW record or Autodiscover record, the error is checked to see if the value should be
                    # set to NULL to allow us to use the error variable to detect if successful

                    $record = ""
                    switch ($p)
                    {
                        "WWW Record" {$record = "www."}
                        "Autodiscover Record" {$record = "autodiscover."}
                    }

                    $error.clear()
                    $w = Resolve-DnsName -Name ($record + $entry.Domain) -Type CNAME -ErrorAction SilentlyContinue
                    $result = $w.NameHost

                    if($result -eq $NULL)
                    {
                        $w = Resolve-DnsName -Name ($record + $entry.Domain) -Type A -ErrorAction SilentlyContinue
                        $result = [string]($w.IPAddress | Sort-Object)
                        if($error.count -gt 0) {$result = $NULL}
                    }
                    $object.$p = $result
                }
                    
                Default
                {
                    # creates an array of the possible terms (different terms) for the property
                    # the array begins with the original property name and then is followed by
                    # all the other options
                    $b = [array]::IndexOf($properties, $p)
                    $possibleOpt = @()
                    $possibleOpt += $p
                    $possibleOpt += $backupProp[$b]
                    # removes all NULL values
                    $possibleOpt = ($possibleOpt.Where({ $_ -ne $NULL }))
                
                    # runs through each option until it gets a hit
                    foreach($opt in $possibleOpt)
                    {
                        if($possibleOpt.Indexof($opt) -gt 0)
                        {
                            Write-Host ("Checking for backup term: " + $opt) -ForegroundColor Yellow
                        }
                    
                        $supress = ($possibleOpt.Indexof($opt) -ne ($possibleOpt.count - 1))
                        $object.$p = FindLineContaining -word ($opt + ": ") -arrstring $importData -supress $supress
                         
                        if($object.$p-ne $NULL)
                        {
                            break
                        }
                    }
                
                    # Converts all the dates to the same format
                    if($p -like '*Date*')
                    {
                        # formatting a specific date format to avoid error
                        $ind = ($object.$p).IndexOf('T')
                        if($ind -ne -1) 
                        {
                            $object.$p = ($object.$p).Substring(0, $ind)
                        } 

                        $date = [dateTime](($object.$p))
                        $object.$p = $date.ToString("dd-MMM-yyyy")
                    } 
                }
            }   

            # writes to screen the property and info
            Write-Host ($p + ": ")$object.$p
        }
        
        $failedDom = $NULL
    }
    else
    {
        # if the whois program fails
        Write-Host "`nWhois failed to return information. Skipping to next domain`n" -ForegroundColor Red
        $object = setupObject -properties $properties
        $object.Domain = $entry.Domain

        # adds failed domain to variable
        $failedDom = $entry.Domain
    }

    $output = "" | Select-Object -Property domainObject, failedDom
    $output.domainObject = $object
    $output.failedDom = $failedDom
    return $output
}


###################################################################################
## Function: modifyInfo
## Input: newInfo - the newly gathered array of domain objects
##        oldInfo - the imported older array of domain objects
##        properties - the array of properties
##        potentialProp2Mod - an array of properties that may cause error from whois
##        failedDomains - the domains the failed during the whois
## Purpose: The whois program sometimes fails to gather all the data during every single run.
##          At times it will fail to connect to the server and leave all the properties for the
##          domain blank.
##          Other times it will output blank values for some of the properties.
##          This function copies overwrites the empty new data with values from the old dataset, 
##          or, prevents blanks from overwriting over existing values imported
##          from the last saved domain properties information file, assuming no one will
##          change the property to blank.
Function modifyInfo ($newInfo, $oldInfo, $properties, $potentialProp2Mod, $failedDomains)
{
    # for each failed domain, the domain's info from oldInfo is just copied over to newInfo
    foreach($fail in $failedDomains)
    {
        $n = [array]::IndexOf($oldInfo.Domain, $fail)
        $i = [array]::IndexOf($newInfo.Domain, $fail)

        if($n -ne -1)
        {
            $newInfo[$i] = $oldInfo[$n]
        }
    }
    
    # checks to make sure that the properties that will potentially be modified
    # do actually exist in the properties array
    $modifyProps = @()
    foreach($pot in $potentialProp2Mod)
    {
       if($properties -contains $pot)
       {
            [array]$modifyProps += $pot
       }
    }

    # runs through each domain for the select properties and looks for 
    # blanks in the newInfo that will overwrite values in the oldInfo
    # it then replaces the blank from the new with the value from the old
    foreach($domain in $newInfo)
    {
        foreach($p in $modifyProps)
        {
            $n = [array]::IndexOf($oldInfo.Domain, $domain.Domain)
                
            if($domain.$p -eq $NULL -and ($oldInfo[$n].$p -ne $NULL))
            {
                $domain.$p = $oldInfo[$n].$p
            }
        }
    }
    return
}
             

###################################################################################
## Function: checkForChangesAndExport
## Input: folder - current directory of folder
##        entries - the list of all the domains being checked
##        newInfo - the new array of objects of all the domains
##        filename - the name of the .csv file that contains all the domain properties information
##        properties - the array of properties
##        modifyProps - the array of potential properties that will have to be modified
##        failedDomains - the domains the failed during the whois
## Output: output - an object containing the variables changes (an object containing all the change info)
##         and newInfo (the array of domain objects of all the information
## Purpose: Check to see if there have been any changes to the imported data since the last test and exports
##          domain information
Function checkForChangesAndExport ($folder, $entries, $newInfo, $filename, $properties, $modifyProps, $failedDomains)
{
    $changes = @()
    # tests to see if a current csv file exists
    if(Test-Path ($folder + "\" + $filename))
    {
        # imports array from current csv file
        $oldInfo = import-csv ($folder + "\" + $filename)
        $oldInfo | foreach{ foreach($p in $properties){ if($_.$p -eq ""){ $_.$p = $NULL}}}

        # For Testing
        # $oldInfo[2]."Creation Date" = "ZZZZZZZZZZZZZZZ"
        # $oldInfo[0]."Registrant Name" = "ZZZZZZZZZZZZZZZ"
        # $newInfo[5]."Registrant Email" = $NULL
        # $newInfo[0]."Registrant Name" = $NULL

        # modifies info to prevent NULL from overwriting over existing values for the specified properties or for a whole domain
        modifyInfo -newInfo $newInfo -oldInfo $oldInfo -properties $properties -potentialProp2Mod $modifyProps -failedDomains $failedDomains

        # skips over Domain (since that won't change), Description (since changes are not important) and Days Till Expiration (since that is always changing)
        # also skips over below code if there are no differences from Compare-Object
        $properties.Remove("Domain")
        $properties.Remove("Days Till Expiration")
        $properties.Remove("Description")

        $members = "Domain", "Property", "Action", "Data"
        Write-Host "`n***Changes to Domain Properties***"

        # loops through each property of every object to see if there have been any changes
        foreach($prop in $properties)
        {
            $comp = Compare-Object -ReferenceObject $oldInfo -DifferenceObject $newInfo -Property "Domain", $prop | Sort-Object "Domain"

            if($comp.count -ne 0)
            {
                foreach($line in $comp)
                {
                    # Checks to see if there are any properties to not include. If there are, it splits
                    # them up into an array and then when there is a match between $p and the list of 
                    # properties not to include, it skips it                   
                    $ind = [array]::IndexOf($entries.Domain, $line.Domain)
                    
                    if($entries[$ind]."Ignore Changes" -eq $NULL -or -Not (((($entries[$ind]."Ignore Changes").Split(",")).Trim()) -contains $prop))    
                    {
                        $object = setupObject ($members) 
                        if($line.SideIndicator -eq "<=")    
                        {
                            $object.Action = "removed"
                        }
                        elseif($line.SideIndicator -eq "=>")
                        {
                            $object.Action = "added"
                        }

                        # adds results to an array containing all the changes of the domains
                        $object.Domain = $line.Domain
                        $object.Property = $prop
                        $object.Data = ($line.$prop)*($line.$prop -ne $NULL) + ('[BLANK]')*($line.$prop -eq $NULL)
                        $changes = [Array]$changes + $object

                        # prints to screen the data added/removed from the property of the domain
                        Write-Host $object.Domain": " $object.Property "--" $object.Action "--" $object.Data -ForegroundColor Cyan

                    }
                }
            }
        }

        if($changes.count -eq 0)
        {
            Write-Host "None" -ForegroundColor Cyan
        }
    }
    
    # rewrites old .csv file
    # warning: If you open this .csv in excel and then save. The dates will be automatically formatted. Thus, during the next run of the script,
    #          the program will think there have been changes to the date.

    $newInfo | Export-csv ($folder + "\" + $filename) -notypeinformation

    $output = "" | Select-Object -Property changes, newInfo
    $output.changes = $changes
    $output.newInfo = $newInfo
    return $output
}

###################################################################################
## Function: prepHTML
## Input: body - the body text of the html
## Purpose: Prepares an HTML file with a header, body, and foot
function prepHTML ($body)
{
#####
    $head=@"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
<style type="text/css">
body {
	font-family: Verdana, Geneva, sans-serif;
	font-size: 10pt;
}
#added {
	background-color: #0C0;
}
#removed {
	background-color: #F00;
}
</style>
</head>
<body>
"@
#####   
    $foot=@"
</body>
</html>
"@
#####
    $return = $head
    $return += $body -join "`r`n"  ##body is an array of strings, join them with a newline after each
    $return += $foot
	$return
}

###################################################################################
## Function: bodyHTML
## Input: changes - an array of objects containing each of the changes occurring
##        data - the array of domain objects containg all the domain information
##        daysLimit - the number of days until expiration when a notice should be sent
## Output: body - the array of lines which will be the text of the 
## Purpose: Prepares the body text for an HTML
function bodyHTML ($changes, $data, $daysLimit)
{
    $body = ("Administrative Notice: Domain Reporter Summary<br>")
    $body += ("--------------------------------------------------------------------------<br>")
    $scriptFullname = $PSCommandPath ; if (!($scriptFullname)) {$scriptFullname =$MyInvocation.InvocationName }
    $scriptName     = Split-Path -Path $scriptFullname -Leaf
    $body += ("<font size=1><I>$scriptName        Computer:$env:computername User:$env:username PSver:"+($PSVersionTable.PSVersion.Major) + "</I></font>")

    $body += ("<br><br><B>Upcoming Expiration Dates (<=$daysLimit days)</B><br>")
    $i = 0

    $sortedData = $data | Sort-Object "Days Till Expiration"
    # adds the domains with expiring dates to the body if they exist
    foreach($d in $sortedData)
    {
        if($d."Days Till Expiration" -le $daysLimit)
        {
            $domain1 = $d.Domain
            $des = $d.Description
            $descrip = "($des)"*($d.Description -ne $NULL)
            $regis = $d.Registrar
            $daysLeft = $d."Days Till Expiration"
            $expDate = $d."Expiration Date"

            $body += ("<U>$domain1</U> $descrip expires in <B><font color= 'red'>$daysLeft days</font></B> on $expDate [$regis]<br>")
            $i++
        }
    }

    if($i -eq 0)
    {
        $body += "None<br>"
    }


    # adds a table of the changes to domain properties to the body if they exist
    $body += ("<br><br><B>Changes to domain properties since last run</B><br>")

    if($changes.count -ne 0)
    {
        $body +=  @('<table width="90%" border="1">')
        $body += ("<tr><td><B>Domain</B></td><td><B>Property</B></td><td><B>Action</B></td><td><B>Data</B></td></tr>")
        foreach($line in $changes)
        {
            $domain2 = $line.Domain
            $property = $line.Property
            $action = $line.Action
            $info = $line.Data
            if($action -eq "added")
            {
                $body += ("<tr><td>$domain2</td><td>$property</td><td id=added>$action</td><td>$info</td></tr>")
            }
            else
            {
                $body += ("<tr><td>$domain2</td><td>$property</td><td id=removed>$action</td><td>$info</td></tr>")
            }
        }
        $body += ("</table>")
    }
    else
    {
        $body += "None<br>"
    }
    
    # writes a list of all the domains checked
    $body += ("<br><br><U><font size='1'>Domains:</font></U><br>")
    foreach($ent in $data)
    {
        $domain3 = $ent.Domain
        $body += ("<font size='1'>$domain3</font><br>")
    }
    return $body
}

###################################################################################
## Function: sendReport
## Input: changes - array of the changes to the domain properties
##        server - the email server
##        from - the mail address to be sent from
##        to - the mail address to be sent to
##        data - the array of domain objects containg all the domain information
##        folder - the current folder where the program is located
##        daysLimit - the number of days until expiration when a notice should be sent
## Purpose: Sends a report via email of all the changes that have been made and the expiring domains
Function sendReport($changes, $server, $from, $to, $data, $folder, $daysLimit) 
{
    $listOfDaysLeft = $data."Days Till Expiration"
    $listOfDaysLeft = ($listOfDaysLeft.Where({ $_ -ne $NULL }))

    if($changes.count -gt 0 -or ($listOfDaysLeft -lt $daysLimit).count -gt 0)
    {         
        $subject = "Administrative Alert: Domain Reporter Summary"
        # prepares the html for the email
        $msg = bodyHTML -changes $changes -data $data -daysLimit $daysLimit
        $html = prepHTML $msg

        # finds out if report folder has been created already
        if(-Not (Test-Path ($folder + "\Logs")))
        {
            New-Item ($folder + "\Logs") -type directory | Out-NULL
        }

        $dtfmt= "yyyy-MM-dd_HHmmss"
        $logFile = "DomainLog_" + (Get-Date -format "$dtfmt") + ".html"
        # stores the html reports in its own folder
        Add-Content ($folder + "\Logs\" + $logFile) $html

        # sends an email if any changes occurred or any expiration dates are approaching            
        #Send-MailMessage -SmtpServer $server -From $from -To $to -Subject $Subject -BodyAsHtml -Body $html
        Write-Host "`nEmail Sent" -ForegroundColor Green
    }
    else
    {
        Write-Host "`nNo changes to the domain properties or upcoming expiring dates. Email not sent" -ForegroundColor Yellow
    }
    return
}

#####################################################################
## MAIN SCRIPT
## 
## To enable scripts, Run powershell 'as admin' then type
## Set-ExecutionPolicy Unrestricted
##
## whois.exe must be located in the same folder as script
## source: https://technet.microsoft.com/en-us/sysinternals/bb89743
#####################################################################

## --- Main Script Header --- ##
$scriptFullname = $PSCommandPath ; if (!($scriptFullname)) {$scriptFullname =$MyInvocation.InvocationName }
$scriptXML      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".xml"  ### replace .ps1 with .xml
$scriptDir      = Split-Path -Path $scriptFullname -Parent
$scriptName     = Split-Path -Path $scriptFullname -Leaf
$scriptBase     = $scriptName.Substring(0, $scriptName.LastIndexOf('.'))
Write-Host "-----------------------------------------------------------------------------"
Write-Host ("$scriptName        Computer:$env:computername User:$env:username PSver:"+($PSVersionTable.PSVersion.Major))
Write-Host "-----------------------------------------------------------------------------"
$folder = $scriptDir
# on first run of whois you must accepteula 

$checkWhois =  & ($folder + "\whois64.exe") -v -nobanner
$acceptEula = findLineContaining -word "accepteula" -arrstring $checkWhois -supress $true -searchWholeLine $true -returnWholeLine $true
if ($acceptEula)
{
    $checkWhois =  & ($folder + "\whois64.exe") -accepteula
}

# loads xml file with variables
globalsLoad

## --- Sends email or opens domain information report, based on switch parameter --- ##

if($onlyReport.IsPresent)
{
    # when in only report mode, the program will simply copy over the last stored domain
    # information csv file into a new folder and be named DomainReport
    $filename = $Globals["reportFilename"]

    if(-Not (Test-Path ($folder + "\" + $filename)))
    {
        Write-Host "Domain information file does not exist.`nRun Domain Reporter without onlyReport parameter to create file`n" -ForegroundColor Red
        pause
        exit
    }

    Write-Host "Opening full domain information report..." -ForegroundColor Green
    
    if(-Not (Test-Path ($folder + "\Reports")))
    {
        New-Item ($folder + "\Reports") -type directory | Out-NULL
    }

    $dtfmt= "yyyy-MM-dd_HHmmss"
    $reportFile = "DomainReport_" + (Get-Date -format "$dtfmt") + ".csv"

    Copy-Item ($folder + "\" + $filename) ($folder + "\Reports\" + $reportFile)
    
    Start-Sleep -s 3 
    Invoke-Item ($folder + "\Reports\" + $reportFile)    
}
else
{
    ## --- Fill $entries with contents of file --- ##

    $domainList = $Globals["domainList"]

    # removes any empty rows in the domains list csv file
    #(Get-Content -path ($folder+ "\" + $domainList)) -notmatch '(^[\s,-]*$)|rows\s*affected' | Set-Content -Path ($folder+ "\" + $domainList)

    # imports the domains with their description and the properties to ignore changes
    $entries = import-csv ($folder+ "\" + $domainList)

    # changes all "" to NULL
    $entries| foreach{ if($_."Description" -eq ""){ $_."Description" = $NULL} if($_."Ignore Changes" -eq ""){ $_."Ignore Changes" = $NULL}}
    $entriescount = $entries.count

    $processed = 0
    $i = 0

    # sets-up an empty array
    $domainInfo = @()
    $failedDomains = @()

    # imports values from globals
    $filename = $Globals["reportFilename"] 
    $properties = $Globals["properties"]
    $mailServer = $Globals["mailServer"]
    $mailFrom = $Globals["mailFrom"]
    $mailTo = $Globals["mailTo"]
    $daysLimit = $Globals["daysLimit"]
    $backup = $Globals["backupProperties"]
    $modifyProps = $Globals["propertiesWithPotError"]


    ## --- Gather Data --- ##
    foreach ($x in $entries)
    {
        $i++
        Write-Host "-----" $i of $entriescount $x.Domain

        $processed++
                
        $output1 = importData -folder $folder -entry $x -properties $properties -backupProp $backup
        $domainInfo = [Array]$domainInfo + $output1.domainObject
        $failedDomains += $output1.failedDom                                                                                      
    }

    # removes NULL
    $failedDomains = ($failedDomains.Where({ $_ -ne $NULL }))

    ## --- Check Data against old copy --- ##

    $output2 = checkForChangesAndExport -folder $folder -entries $entries -newInfo $domainInfo -filename $filename -properties $properties -modifyProps $modifyProps -failedDomains $failedDomains
    $domainInfo = $output2.newInfo
    $changes = $output2.changes
    
    # goes through sending email process
    sendReport -changes $changes -server $mailServer -from $mailFrom -to $mailTo -data $domainInfo -folder $folder -daysLimit $daysLimit

    Start-Sleep -s 10  
    Write-Host "------------------------------------------------------------------------------------"
    Write-Host "------------------------------------------------------------------------------------"
    $entries.Domain
}