# Daily HTML Email Report from OpenManage Essentials
# This report can be ran either as a standalone or within OME
# It will gather basic inventory, provide a disk space report
# processes report, service report, system and application
# event log.
#
# This script will generate a valid XHTML 1.0 Transitional HTML file
# that is emailed out to users specified below.
#
# Check validation at http://validator.w3.org. Results from validator:
# This document was successfully checked as XHTML 1.0 Transitional!
#
# Example usage as standalone, from PowerShell command window type:
# .\dailyreport-OME.ps1 .\servers.txt
# where servers.txt is a file containing a list of Server names or IP's to run this against.
#
# Example usage using OME is to create a generic command line task and run this script against
# the servers.txt task that creates the txt file that is used for a scheduled tasks.


Function Get-DailyHTMLReport
{
$Path = "c:\tmp"
$OutputFile= "DailyReport_$(get-date -format ddMMyyyy).html" # Name of the file that gets emailed out
$list = $args[0] #This accepts the argument you add to your scheduled task in OME or list.txt
$computerList = Get-Content $list # grab the namesof the servers/computers from file

#Set warning and critical thresholds below in percent for disk report
$FreePercentWarningThreshold=30
$FreePercentCriticalThreshold=10
#Number of proccess to fetch that are using the most amount of memory
$ProccessNumToFetch = 10
#Number of events to gather from system and application logs that are
either warning or critical
$EventNum = 3
#Email settings for report
$users = "sysadmin@mycompany.com" #List of users to email your report to (separated by comma)
$fromemail = "OMEDailyReports@mycompany.com"
$server = "smtp.mycompany.com" #SMTP server to use for sending out email
Write-Host "Starting to Generate HTML Daily Email Report...."
#Create a new report file to be emailed out
New-Item -ItemType File -Name $OutputFile -Path $Path -Force | Out-Null
#Write the HTML header information to file
writeHtmlHeader "$Path\$OutputFile"

#Process each server and run through script
foreach ($computer in $ComputerList)
{
$ErrorActionPreference = "silentlycontinue"
#Test to make sure computer is up and that you are using the proper
credentials
if ((Test-Connection -ComputerName $computer -Quiet -Count 1) -and
(Test-Path \\$Computer\admin`$ ) )
{
#Convert IP into an FQDN name for our report
write-host "$computer - UP" -ForegroundColor Green
$IP =
[System.Net.Dns]::GetHostEntry($computer).AddressList |
%{$_.IPAddressToString}
$IP | %{$HostName =
[System.Net.Dns]::GetHostEntry($_).HostName}
If ($IP)
{
Write-Host "IP is $IP"
Write-Host "Hostname is $hostname"
$wmi = (gwmi -computer $hostname win32_service)
If ($wmi)
{
#Create a header for each server we are
processing
$compHeader = @"
<table>
 <tr>
 <td colspan="6"><h2>Report for:
$hostname</h2></td>
 </tr>
</table><p></p>
"@
Add-Content "$Path\$OutputFile" $compHeader
# Run the inventory report
InventoryReport
# Run the disk report usage
DiskReport
# Run the top processes report
ProcessReport
# Run the services report
ServiceReport
# Run the System Event Log Report
SystemLogReport
# Run the Application Event Log Report
AppLogReport
}
else
{
Write-Host "Unable to access to WMI, please check the
user has admin access to server"
}
}
else

{
write-host "No Hostname found make sure you have DNS
PTR records for your servers"
}
}
else
{
Write-Host "$computer Wrong Credentials, Not
Responding or Not a Windows Server" -ForegroundColor Red
}
}
# Close out all open HTML tags
Add-Content "$Path\$OutputFile" "</div></div></body></html>"
# Finish up Report
Write-Host "Daily HTML Report File Path is: $Path\$OutputFile" -
ForegroundColor Green
# Email out report and add HTML file as an attachment
$HTMLmessage = get-content $Path\$OutputFile
send-mailmessage -from $fromemail -to $users -subject "Daily Systems
Report" -BodyAsHTML -body "$HTMLmessage" -attachment $Path\$OutputFile -
priority Normal -smtpServer $server
}


# Write HTML Header information to our Report
# Use CSS to make report more readable
Function writeHtmlHeader
{
$date = (get-date -Format F)
$header = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Daily Reports</title>
<style type="text/css">
<!--
body {
font: 100%/1.4 Verdana, Arial, Helvetica, sans-serif;
background: #FFFFFF;
margin: 0;
padding: 0;
color: #000;
}
.container {
width: 100%;
margin: 0 auto;
}
h1 {
font-size: 18px;
}
h2 {
color: #FFF;
padding: 0px;
margin: 0px;
font-size: 14px;
background-color: #006400;
}

h3 {
color: #FFF;
padding: 0px;
margin: 0px;
font-size: 14px;
background-color: #191970;
}
h4 {
color: #348017;
padding: 0px;
margin: 0px;
font-size: 10px;
font-style: italic;
}
.header {
text-align: center;
}
.container table {
width: 100%;
font-family: Verdana, Geneva, sans-serif;
font-size: 12px;
font-style: normal;
font-weight: bold;
font-variant: normal;
text-align: center;
border: 2px solid black;
padding: 0px;
margin: 0px;
}
td {
font-weight: normal;
border: 1px solid grey;
}
th {
font-weight: bold;
border: 1px solid grey;
text-align: center;
}
-->
</style></head>
<body>
<div class="container">
 <div class="header">
 <h1>OpenManage Daily Reports</h1>
<h1>$date</h1>
 </div>
 <div class="content">
"@
Add-Content "$Path\$OutputFile" $header
}
# Go through all disks on the server and color code the ones that are
# in the thresholds specified above for warning and critical
Function WriteDiskInfo
(
[string]$fileName,
[string]$devId,
[string]$volName,

[double]$freeSpace,
[double]$totalSpace
)
{
$greenColor = "#638B38"
$yellowColor = "#F5BD22"
$redColor = "#C1281C"
$totalSpace=[math]::Round(($totalSpace/1GB),2)
$freeSpace=[Math]::Round(($freeSpace/1GB),2)
$usedSpace = $totalSpace - $freeSpace
$usedSpace=[Math]::Round($usedSpace,2)
$freePercent = ($freeSpace/$totalSpace)*100
$freePercent = [Math]::Round($freePercent,0)
if ($freePercent -gt $FreePercentWarningThreshold)
{
$color = $greenColor
$dataRow = @"
<tr>
<td>$devid</td>
<td>$volName</td>
<td>$totalSpace</td>
<td>$usedSpace</td>
<td>$freeSpace</td>
<td style="background-color: #638B38;">$freePercent</td>
</tr>
"@
Add-Content $fileName $dataRow
}
elseif ($freePercent -le $FreePercentCriticalThreshold)
{
$color = $redColor
$dataRow = @"
<tr>
<td>$devid</td>
<td>$volName</td>
<td>$totalSpace</td>
<td>$usedSpace</td>
<td>$freeSpace</td>
<td style="background-color: #C1281C;">$freePercent</td>
</tr>
"@
Add-Content $fileName $dataRow
}
else
{
$color = $yellowColor
$dataRow = @"
<tr>
<td>$devid</td>
<td>$volName</td>
<td>$totalSpace</td>
<td>$usedSpace</td>
<td>$freeSpace</td>
<td style="background-color: #F5BD22;">$freePercent</td>
</tr>
"@

Add-Content $fileName $dataRow
}
}
# Provide the time the server has been up
Function Get-HostUptime
{
Write-Host $Computer
$Uptime = Get-WmiObject -Class Win32_OperatingSystem -ComputerName
$Computer
$LastBootUpTime = $Uptime.ConvertToDateTime($Uptime.LastBootUpTime)
$Time = (Get-Date) - $LastBootUpTime
Return '{0:00} Days, {1:00} Hours, {2:00} Minutes, {3:00} Seconds' -f
$Time.Days, $Time.Hours, $Time.Minutes, $Time.Seconds
}
#Gather basic inventory from server
Function InventoryReport
{
$OS = (Get-WmiObject Win32_OperatingSystem -computername
$computer).caption
$SystemInfo = Get-WmiObject -Class Win32_OperatingSystem -computername
$computer | Select-Object Name, TotalVisibleMemorySize, FreePhysicalMemory
$ModelInfo = Get-WmiObject -Class Win32_ComputerSystem -ComputerName
$computer | Select-Object Manufacturer, Model,DNSHostName,Domain
$TotalRAM = $SystemInfo.TotalVisibleMemorySize/1MB
$FreeRAM = $SystemInfo.FreePhysicalMemory/1MB
$UsedRAM = $TotalRAM - $FreeRAM
$RAMPercentFree = ($FreeRAM / $TotalRAM) * 100
$TotalRAM = [Math]::Round($TotalRAM, 2)
$FreeRAM = [Math]::Round($FreeRAM, 2)
$UsedRAM = [Math]::Round($UsedRAM, 2)
$RAMPercentFree = [Math]::Round($RAMPercentFree, 2)
$Make = $ModelInfo.model
$Made = $ModelInfo.manufacturer
$Name = $ModelInfo.DNSHostName
$Domain = $ModelInfo.Domain
$SystemUptime = Get-HostUptime -ComputerName $computer
$InventoryReportHeader = @"
<table>
 <tr>
 <td colspan="6"><h3>Inventory Report</h3>
<h4>This provides a basic overview of the system and some key
statistics.</h4>
</td>
 </tr>
<tr>
<th>System Uptime</th>
<td>$SystemUptime</td>
</tr>
<tr>
<th>Computer Name</th>
<td>$Name</td>
</tr>
<tr>
<th>Computer Domain</th>
<td>$Domain</td>
</tr>
<tr>

<th>Manufacturer</th>
<td>$Made</td>
</tr>
<tr>
<th>Model</th>
<td>$Make</td>
</tr>
<tr>
<th>Operating System</th>
<td>$OS</td>
</tr>
<tr>
<th>Total RAM (GB)</th>
<td>$TotalRAM</td>
</tr>
<tr>
<th>Free RAM (GB)</th>
<td>$FreeRAM</td>
</tr>
<tr>
<th>Percent free RAM</th>
<td>$RAMPercentFree</td>
</tr>
</table><p></p>
"@
Add-Content "$Path\$OutputFile" $InventoryReportHeader
}
# Assemble the disk report
Function DiskReport
{
$DiskReportHeader = @"
<table>
 <tr>
 <td colspan="6"><h3>Disk Report</h3>
<h4>Drive(s) listed below have less than
$thresholdspace % free space. Disk within the thresholds specified are
colored to identify them easily</h4>
</td>
 </tr>
<tr>
 <th>Drive</th>
 <th>Drive Label</th>
 <th>Total Capacity (GB)</th>
 <th>Used Capacity (GB) </th>
 <th>Free Space (GB) </th>
 <th>Freespace %</th>
 </tr>
"@
Add-Content "$Path\$OutputFile" $DiskReportHeader
$disks = Get-WmiObject win32_logicaldisk -
ComputerName $computer | Where-Object {$_.drivetype -eq 3}
foreach ($item in $disks)
{
WriteDiskInfo "$Path\$OutputFile"
$item.DeviceID $item.VolumeName $item.FreeSpace $item.Size

Add-Content "$Path\$OutputFile" "</table><p></p>"
}
# Assemble to top processes that are consuming the highest amount of memory
Function ProcessReport
{
$ProcessReportHeader = @"
<table>
 <tr>
 <td colspan="6"><h3>Processes Report</h3>
<h4>The following $ProccessNumToFetch processes
are those consuming the highest amount of Working Set (WS) Memory (bytes) on
$computer</h4></td>
 </tr>
</table><p></p>
"@
Add-Content "$Path\$OutputFile" $ProcessReportHeader
$TopProcesses = (Get-process -ComputerName $computer
| select-object ws,name | sort-object –property WS -Descending | selectobject
–First $ProccessNumToFetch | convertto-html -Fragment) | Add-Content
"$Path\$OutputFile"
Add-Content "$Path\$OutputFile" "<p></p>"
}
# Assemble Service Report
Function ServiceReport
{
$ServiceReportHeader = @"
<table>
 <tr>
 <td colspan="6"><h3>Services Report</h3>
<h4>The following services are those which are
set to Automatic startup type, yet are currently not running on
$computer</h4></td>
 </tr>
</table><p></p>
"@
$Services = (Get-WmiObject -Class Win32_Service -
ComputerName $computer | Select-Object DisplayName,Name,StartMode,State |
Where-Object {$_.StartMode -eq "Auto" -and $_.State -eq "Stopped"})
If ($Services -ne $null)
{
Add-Content "$Path\$OutputFile" $ServiceReportHeader
$StopServices = (Get-WmiObject -Class Win32_Service -
ComputerName $computer | Select-Object DisplayName,Name,StartMode,State |
Where-Object {$_.StartMode -eq "Auto" -and $_.State -eq "Stopped"} |
ConvertTo-Html -Fragment) -replace "<table>",'<table>' | Add-Content
"$Path\$OutputFile"
Add-Content "$Path\$OutputFile" "<p></p>"
}
}
#Assemble System Event Log Report
Function SystemLogReport
{
$SysReportHeader = @"
<table>
 <tr>
 <td colspan="6"><h3>System Event Report</h3>

 <h4>The following is a list of the last
$EventNum <b>System log</b> events that had an Event Type of either Warning
or Error on $computer</h4></td>
 </tr>
</table><p></p>
"@
Add-Content "$Path\$OutputFile" $SysReportHeader
$SystemEvents = (Get-EventLog -ComputerName $computer
-LogName System -EntryType Error,Warning -Newest $EventNum | Select-Object
Message,Source,EntryType,TimeGenerated | ConvertTo-Html -Fragment) -replace
"<table>",'<table style="text-align: left" >' | Add-Content
"$Path\$OutputFile"
Add-Content "$Path\$OutputFile" "<p></p>"
}
#Assemble Application Event Log Report
Function AppLogReport
{
$AppReportHeader = @"
<table>
 <tr>
 <td colspan="6"><h3>Application Event Report</h3>
<h4>The following is a list of the last
$EventNum <b>Application log</b> events that had an Event Type of either
Warning or Error on $computer</h4></td>
 </tr>
</table><p></p>
"@
Add-Content "$Path\$OutputFile" $AppReportHeader
$ApplicationEvents = (Get-EventLog -ComputerName
$computer -LogName Application -EntryType Error,Warning -Newest $EventNum |
Select-Object Message,Source,EntryType,TimeGenerated | ConvertTo-Html -
Fragment) -replace "<table>",'<table style="text-align: left">' | Add-Content
"$Path\$OutputFile"
Add-Content "$Path\$OutputFile" "<p></p>"
}
# Run Main Report
Get-DailyHTMLReport $args[0]





