Set-Location C:\temp
#connecting to outlook and retriveing mailaddress 
add-type -AssemblyName "Microsoft.Office.Interop.Outlook"
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNameSpace("MAPI")
$mailaddress = $namespace.Accounts |select -ExpandProperty SmtpAddress
#retriveing computer name 
$hostname = $env:COMPUTERNAME

#retriveing last logon name 

$log = get-eventlog system -Source Microsoft-Windows-Winlogon| where InstanceID -EQ 7001 |select -First 1
$username = (New-Object System.Security.Principal.SecurityIdentifier $log.ReplacementStrings[1]).Translate([System.Security.Principal.NTAccount]).Value

#pushing to excel 
$report = import-csv 'C:\temp\cross report.csv'

$report.email = $mailaddress
$report.username = $username
$report.hostname = $hostname

$report |Export-Csv 'c:\temp\cross report.csv' -Append
