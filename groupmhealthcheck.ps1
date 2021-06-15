$CurrentUser = whoami
$logonuser= $CurrentUser.Split("\\")[1]
$date = Get-Date -UFormat "%d%m%Y%H%M"
$strdate = Get-Date
$filedate = Get-Date -UFormat "%d%m%Y"
$reportdate = Get-Date -UFormat "%d/%m/%Y"
$ondreportdate = Get-Date -UFormat "%d/%m/%Y %H:%M EST"
New-Item Logs -Type directory -ErrorAction SilentlyContinue | Out-Null
$AuditLog = ".\Logs\Tasklog.txt"
$Errorlog = ".\Logs\Errors_"+$filedate+".txt"
"===========================================" >> $AuditLog
$strdate >> $AuditLog
"Script is initiated by " + $logonuser >> $AuditLog

$computers = Import-Csv .\ServersPROD.txt
$urls = Import-Csv .\URLs.txt
# $cred = get-Credential
#$cred = [System.Net.CredentialCache]::DefaultNetworkCredentials
"Drive,CapacityGB,FreeSpaceGB,PercentageFree,ServerName" > DiskSpaceReport.csv
"WebSite,Status,ServerName" > WebSiteReport.csv
"ApplicationPool,Status,ServerName" > AppPoolReport.csv
"WinServiceInAutoStartupType,Status,ServerName" > WinServiceReport.csv
"WebURL,ResponseCode,Status" > WebURLReport.csv
$disk=@()
$IISSite=@()
$IISAppPool=@()
$WinService=@()
$UnavailServer=@()
$DiskReport=@()
$IISSiteReport=@()
$IISAppPoolReport=@()
$WinServiceReport=@()
$WebURLReport=@()

Foreach ($computer in $computers) {
try
{
	$disk = Get-WMIObject win32_logicaldisk -computer $computer.Server -ErrorAction Stop |  select DeviceID,Size,Freespace
	$IISSite = get-wmiobject  -class Site -Authentication PacketPrivacy -Impersonation Impersonate -namespace "root/webadministration" -computer $computer.Server -ErrorAction Stop | select name, @{Expression={if($_.GetState().ReturnValue -eq 1){"Started"}else{"Not Started"}};Label="State"}
	$IISAppPool = get-wmiobject  -class applicationpool -Authentication PacketPrivacy -Impersonation Impersonate -namespace "root/webadministration"  -computer $computer.Server -ErrorAction Stop | select name, @{Expression={if($_.GetState().ReturnValue -eq 1){"Started"}else{"Not Started"}};Label="State"} 
	$WinService = Get-WmiObject Win32_Service -ComputerName $computer.Server -ErrorAction Stop | where {$_.StartMode -eq "Auto" -and $_.State -eq "Stopped"} | select displayname,name,state
}
catch
{
	"===========================================" >> $Errorlog
	$strdate >> $Errorlog
	"Either RPC Server is unavailable or Access is denied for Server " + $computer.Server >> $Errorlog
	"===========================================" >> $Errorlog
	"				" >> $Errorlog
	$UnavailServer+=,$computer.Server
	continue;
}

if ($disk.Count)
{
$i=0
Do {
if ($disk[$i].Size -gt 0)
{
$deviceID = $disk[$i].deviceid
$capacity = [math]::Round($disk[$i].Size/1GB)
$freespace = [math]::Round($disk[$i].freespace/1GB)
$percentfree = [math]::Round(($disk[$i].freespace/$disk[$i].Size) * 100,2)
$deviceID + "," + $capacity + "," + $freespace + "," + $percentfree + "," + $computer.Server >> DiskSpaceReport.csv
}
$i++
} While ($i -lt $disk.Count)
}
elseif ($disk.Size -gt 0)
{
$deviceID = $disk.deviceid
$capacity = [math]::Round($disk.Size/1GB)
$freespace = [math]::Round($disk.freespace/1GB)
$percentfree = [math]::Round(($disk.freespace/$disk.Size) * 100,2)
$deviceID + "," + $capacity + "," + $freespace + "," + $percentfree + "," + $computer.Server >> DiskSpaceReport.csv
}
$disk = @()

if ($IISSite.Count)
{
$j=0
Do {
$WebSite = $IISSite[$j].name
$Status = $IISSite[$j].State
$WebSite + "," + $Status  + "," + $computer.Server >> WebSiteReport.csv
$j++
} While ($j -lt $IISSite.Count)
}
else
{
$WebSite = $IISSite.name
$Status = $IISSite.State
$WebSite + "," + $Status  + "," + $computer.Server >> WebSiteReport.csv
}
$IISSite = @()

if ($IISAppPool.Count)
{
$k=0
Do {
$AppPool = $IISAppPool[$k].name
$Status = $IISAppPool[$k].State
$AppPool + "," + $Status + "," + $computer.Server >> AppPoolReport.csv
$k++
} While ($k -lt $IISAppPool.Count)
}
else
{
$AppPool = $IISAppPool.name
$Status = $IISAppPool.State
$AppPool + "," + $Status + "," + $computer.Server >> AppPoolReport.csv
}
$IISAppPool = @()

if ($WinService.Count)
{
$l=0
Do {
$service = $WinService[$l].DisplayName
$Status = $WinService[$l].State
$service + "," + $Status + "," + $computer.Server >> WinServiceReport.csv
$l++
} While ($l -lt $WinService.Count)
}
else
{
$service = $WinService.DisplayName
$Status = $WinService.State
$service + "," + $Status + "," + $computer.Server >> WinServiceReport.csv
}
$WinService = @()

}

$DiskReport = Import-Csv .\DiskSpaceReport.csv | Sort {[decimal]$_.percentagefree} | where {[decimal]$_.percentagefree -lt 15 -and ($_.Drive -ne "F:" -or ($_.ServerName -ne "PSCSPSU00119.AD.INSIDEMEDIA.NET" -and $_.ServerName -ne "PSCSPSU00120.AD.INSIDEMEDIA.NET"))}
$IISSiteReport = Import-Csv .\WebSiteReport.csv | where {$_.Status -eq "Not Started" -and $_.WebSite -ne "Default Web Site"}
$IISAppPoolReport = Import-Csv .\AppPoolReport.csv | where {$_.Status -eq "Not Started" -and $_.ApplicationPool -ne "SharePoint Web Services Root" -and $_.ApplicationPool -ne "eawiki.insidemedia.net" -and ($_.ApplicationPool -ne "ControlPoint" -or $_.ServerName -NotLike "PSCSPSP0010*") -and ($_.ApplicationPool -ne "SharePoint - contenttypes-uatext.insidemedia.net80" -or $_.ServerName -ne "NYCSPSU01117.AD.INSIDEMEDIA.NET") -and ($_.ApplicationPool -ne "SharePoint - inside-uat.wppgts.com80" -or $_.ServerName -ne "NYCSPSU01117.AD.INSIDEMEDIA.NET") -and ($_.ApplicationPool -ne "SharePoint - spine-UAT.mediacom.com443" -or $_.ServerName -ne "NYCSPSU01117.AD.INSIDEMEDIA.NET") -and ($_.ApplicationPool -ne "SharePoint - my-uat.groupm.com80" -or ($_.ServerName -ne "NYCSPSU01116.AD.INSIDEMEDIA.NET" -and $_.ServerName -ne "NYCSPSU01117.AD.INSIDEMEDIA.NET"))}
$WinServiceReport = Import-Csv .\WinServiceReport.csv | where {$_.WinServiceInAutoStartupType -ne "Google Update Service (gupdate)" -and $_.WinServiceInAutoStartupType -ne "Remote Registry" -and $_.WinServiceInAutoStartupType -ne "Software Protection" -and $_.WinServiceInAutoStartupType -ne "Microsoft .NET Framework NGEN v4.0.30319_X86" -and $_.WinServiceInAutoStartupType -ne "Microsoft .NET Framework NGEN v4.0.30319_X64" -and $_.WinServiceInAutoStartupType -ne "Windows Firewall" -and $_.WinServiceInAutoStartupType -ne "Sophos Web Intelligence Update" -and $_.WinServiceInAutoStartupType -ne "Shell Hardware Detection" -and $_.WinServiceInAutoStartupType -ne "BES Client" -and $_.WinServiceInAutoStartupType -ne "Windows Modules Installer" -and ($_.WinServiceInAutoStartupType -NotLike "Nintex*" -or ($_.ServerName -NotLike "PSCSPSU*" -and $_.ServerName -NotLike "PSCSPST*"))}

Foreach ($url in $urls) {
try
{
$webstats = Invoke-WebRequest $url.WebURL -UseDefaultCredentials | select StatusCode
if ($webstats.StatusCode -eq "200")
{
$url.WebURL + "," + $webstats.StatusCode + ",OK" >> WebURLReport.csv
}
else
{
$url.WebURL + "," + $webstats.StatusCode + ",UnexpectedResponse" >> WebURLReport.csv
}
}
catch
{
$url.WebURL + ",No Response,ConnectionError" >> WebURLReport.csv
}
}
$WebURLReport = Import-Csv .\WebURLReport.csv | where {$_.Status -ne "OK"}

#$a = "<p style='color:red;'><i>Critical, Uncreachable or Not Started details visible in mail. Detailed Info in attachments.</i><span style='color:black;'><bold> , </bold></span><span style='color:yellow;'><i>Warning</i></span><span style='color:black;'><bold> , </bold></span><span style='color:green;'><i>Information</i></span></p>"
$a = "<p style='color:black;'><b>GroupM SharePoint PROD Daily HealthCheck Summary Report - $ondreportdate</b></p>"
$a = $a + "<p style='color:red;'><i>Items to be looked in priority(if any) are listed in Red. Detailed Info in attachments.</i></p>"
#$a = $a + "<p style='color:green;'><i>Note: Nintex Workflow Service in Test and UAT SharePoint Servers are excluded from summary.</i></p>"
$a = $a + "<style>"
$a = $a + "BODY{background-color:white;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:thistle}"
$a = $a + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:white}"
$a = $a + "</style>"
$b = @()
$c = "<p style='color:black;'><i>Please click <a href='https://inside.groupm.com/sites/HealthCheck/SharePoint/SPDailyHealthCheck/SPHealthCheckSummary/LatestSummary.htm'>here</a> after 15 minutes of receiving this mail to access above report through browser</i></p>"
$c = $c + "<p style='color:black;'><i>Please click <a href='https://inside.groupm.com/sites/HealthCheck/SharePoint/SPDailyHealthCheck/'>here</a> to view the current and past reports</i></p>"
$c = $c + "<style>"
$c = $c + "BODY{background-color:white;}"
$c = $c + "</style>"
$d = "<p style='color:black;'></p>"
$d = $d + "<style>"
$d = $d + "BODY{background-color:white;}"
$d = $d + "</style>"

##Formatting HTML Table for Unreachable servers
if ($UnavailServer)
{
$strTableStartHTML0 = "<TABLE CELLSPACING=1 CELLPADDING=1 BORDER=1>`n<TR><TH>UnreachableServers</TH></TR>"
$strTableBodyHTML0 = for ($x=0;$x -lt $UnavailServer.Count;$x++) {
	$unavail = $UnavailServer[$x]
	$strRowStyleHTML0 = " STYLE='color:red'"
	"<TR$strRowStyleHTML0><TD><b>$($unavail)</b></TD></TR>`n"
	}
$strTableEndHTML0 = "</TABLE>"
$strTable0 = $strTableStartHTML0 + $strTableBodyHTML0 + $strTableEndHTML0
$UnavailServer = @()
}
else
{
$strTableStartHTML0 = "<TABLE CELLSPACING=1 CELLPADDING=1 BORDER=1>`n<TR><TH>Connectivity Report</TH></TR>"
$strRowStyleHTML0 = " STYLE='color:green'"
$strTableBodyHTML0 = "<TR$strRowStyleHTML0><TD><b>All Servers are reachable</b></TD></TR>`n"
$strTableEndHTML0 = "</TABLE>"
$strTable0 = $strTableStartHTML0 + $strTableBodyHTML0 + $strTableEndHTML0
}

##Formatting HTML Table for Disk Space Report
if ($DiskReport)
{
$strTableStartHTML1 = "<TABLE CELLSPACING=1 CELLPADDING=1 BORDER=1>`n<TR><TH>Drive</TH><TH>Capacity(GB)</TH><TH>FreeSpace(GB)</TH><TH>PercentageFree</TH><TH>ServerName</TH></TR>"
$strTableBodyHTML1 = foreach ($diskr in $DiskReport) {
	$perc = [decimal]$diskr.PercentageFree
	$strRowStyleHTML1 = if ($perc -lt 10) {" STYLE='color:red'"} elseif ($perc -lt 15) {" STYLE='color:yellow'"} else {" STYLE='color:green'"}
	"<TR$strRowStyleHTML1><TD><b>$($diskr.Drive)</b></TD><TD><b>$($diskr.CapacityGB)</b></TD><TD><b>$($diskr.FreeSpaceGB)</b></TD><TD><b>$($diskr.PercentageFree)%</b></TD><TD><b>$($diskr.ServerName)</b></TD></TR>`n"
	}
$strTableEndHTML1 = "</TABLE>"
$strTable1 = $strTableStartHTML1 + $strTableBodyHTML1 + $strTableEndHTML1
$DiskReport = @()
}
else
{
$strTableStartHTML1 = "<TABLE CELLSPACING=1 CELLPADDING=1 BORDER=1>`n<TR><TH>Storage Report</TH></TR>"
$strRowStyleHTML1 = " STYLE='color:green'"
$strTableBodyHTML1 = "<TR$strRowStyleHTML1><TD><b>Disk Free Space of all Servers are in Safe Limit</b></TD></TR>`n"
$strTableEndHTML1 = "</TABLE>"
$strTable1 = $strTableStartHTML1 + $strTableBodyHTML1 + $strTableEndHTML1
}

##Formatting HTML Table for WebSite Status Report
if ($IISSiteReport)
{
$strTableStartHTML2 = "<TABLE CELLSPACING=1 CELLPADDING=1 BORDER=1>`n<TR><TH>WebSite</TH><TH>Status</TH><TH>ServerName</TH></TR>"
$strTableBodyHTML2 = foreach ($site in $IISSiteReport) {
	$strRowStyleHTML2 = " STYLE='color:red'"
	"<TR$strRowStyleHTML2><TD><b>$($site.WebSite)</b></TD><TD><b>$($site.Status)</b></TD><TD><b>$($site.ServerName)</b></TD></TR>`n"
	}
$strTableEndHTML2 = "</TABLE>"
$strTable2 = $strTableStartHTML2 + $strTableBodyHTML2 + $strTableEndHTML2
$IISSiteReport = @()
}
else
{
$strTableStartHTML2 = "<TABLE CELLSPACING=1 CELLPADDING=1 BORDER=1>`n<TR><TH>WebSite Report</TH></TR>"
$strRowStyleHTML2 = " STYLE='color:green'"
$strTableBodyHTML2 = "<TR$strRowStyleHTML2><TD><b>All WebSites are Running</b></TD></TR>`n"
$strTableEndHTML2 = "</TABLE>"
$strTable2 = $strTableStartHTML2 + $strTableBodyHTML2 + $strTableEndHTML2
}

##Formatting HTML Table for Application Pool Status Report
if ($IISAppPoolReport)
{
$strTableStartHTML3 = "<TABLE CELLSPACING=1 CELLPADDING=1 BORDER=1>`n<TR><TH>ApplicationPool</TH><TH>Status</TH><TH>ServerName</TH></TR>"
$strTableBodyHTML3 = foreach ($pool in $IISAppPoolReport) {
	$strRowStyleHTML3 = " STYLE='color:red'"
	"<TR$strRowStyleHTML3><TD><b>$($pool.ApplicationPool)</b></TD><TD><b>$($pool.Status)</b></TD><TD><b>$($pool.ServerName)</b></TD></TR>`n"
	}
$strTableEndHTML3 = "</TABLE>"
$strTable3 = $strTableStartHTML3 + $strTableBodyHTML3 + $strTableEndHTML3
$IISAppPoolReport = @()
}
else
{
$strTableStartHTML3 = "<TABLE CELLSPACING=1 CELLPADDING=1 BORDER=1>`n<TR><TH>ApplicationPool Report</TH></TR>"
$strRowStyleHTML3 = " STYLE='color:green'"
$strTableBodyHTML3 = "<TR$strRowStyleHTML3><TD><b>All Application Pools are Running</b></TD></TR>`n"
$strTableEndHTML3 = "</TABLE>"
$strTable3 = $strTableStartHTML3 + $strTableBodyHTML3 + $strTableEndHTML3
}

##Formatting HTML Table for Windows Service Status Report
if ($WinServiceReport)
{
$strTableStartHTML4 = "<TABLE CELLSPACING=1 CELLPADDING=1 BORDER=1>`n<TR><TH>WinServiceInAutoStartupType</TH><TH>Status</TH><TH>ServerName</TH></TR>"
$strTableBodyHTML4 = foreach ($service in $WinServiceReport) {
	$strRowStyleHTML4 = " STYLE='color:red'"
	"<TR$strRowStyleHTML4><TD><b>$($service.WinServiceInAutoStartupType)</b></TD><TD><b>$($service.Status)</b></TD><TD><b>$($service.ServerName)</b></TD></TR>`n"
	}
$strTableEndHTML4 = "</TABLE>"
$strTable4 = $strTableStartHTML4 + $strTableBodyHTML4 + $strTableEndHTML4
$WinServiceReport = @()
}
else
{
$strTableStartHTML4 = "<TABLE CELLSPACING=1 CELLPADDING=1 BORDER=1>`n<TR><TH>Windows Service Report</TH></TR>"
$strRowStyleHTML4 = " STYLE='color:green'"
$strTableBodyHTML4 = "<TR$strRowStyleHTML4><TD><b>All Services(Filtered) in Automatic Startup Type are Running</b></TD></TR>`n"
$strTableEndHTML4 = "</TABLE>"
$strTable4 = $strTableStartHTML4 + $strTableBodyHTML4 + $strTableEndHTML4
}

##Formatting HTML Table for Web URL Status Report
if ($WebURLReport)
{
$strTableStartHTML5 = "<TABLE CELLSPACING=1 CELLPADDING=1 BORDER=1>`n<TR><TH>WebURL</TH><TH>ResponseCode</TH><TH>Status</TH></TR>"
$strTableBodyHTML5 = foreach ($web in $WebURLReport) {
	$strRowStyleHTML5 = " STYLE='color:red'"
	"<TR$strRowStyleHTML5><TD><b>$($web.WebURL)</b></TD><TD><b>$($web.ResponseCode)</b></TD><TD><b>$($web.Status)</b></TD></TR>`n"
	}
$strTableEndHTML5 = "</TABLE>"
$strTable5 = $strTableStartHTML5 + $strTableBodyHTML5 + $strTableEndHTML5
$WebURLReport = @()
}
else
{
$strTableStartHTML5 = "<TABLE CELLSPACING=1 CELLPADDING=1 BORDER=1>`n<TR><TH>URL Check Report</TH></TR>"
$strRowStyleHTML5 = " STYLE='color:green'"
$strTableBodyHTML5 = "<TR$strRowStyleHTML5><TD><b>All Monitored URLs Returned a Success Status</b></TD></TR>`n"
$strTableEndHTML5 = "</TABLE>"
$strTable5 = $strTableStartHTML5 + $strTableBodyHTML5 + $strTableEndHTML5
}

#Search Health Check

$farms = @()
$farmfiles = ls SearchHealthCheck\SearchReports\*SearchFarmInitStatus.txt | select name
foreach ($farmfile in $farmfiles) {
$item = $farmfile.Name -split "-" 
$farms += $item[0]
}
if ($farms)
{
"Farm,ApplicationName,RunningStatus,HealthStatus" > SearchFarmsReport.csv
"Farm,ApplicationName,Component,Server,State,Remarks" > SearchComponentsReport.csv
foreach ($farm in $farms) {
if ($farm -eq "SP2013Prod")
{
$srchupldtime = Get-Content .\SearchHealthCheck\FileUploadTime.txt
$b = "<p style='color:black;'><b>Please click <a href='https://inside.groupm.com/sites/HealthCheck/SharePoint/_layouts/15/osssearchresults.aspx?k=IBM_$srchupldtime.docx'>here</a> to verify if the document uploaded on Prod within an hour is displaying in Search results</b></p>"
#$b = $b + "<p style='color:black;'><b>Please click <a href='https://uat13inside.groupm.com/sites/TEST-Connect/_layouts/15/osssearchresults.aspx?k=IBM_$srchupldtime.docx'>here</a> to verify if the document uploaded on UAT within an hour is displaying in Search results</b></p>"
#$b = $b + "<p style='color:black;'><b>Please copy the link https://test-inside.groupm.com/sites/test1/_layouts/15/osssearchresults.aspx?k=IBM_$srchupldtime.docx and access through GroupM VPN to verify if the document uploaded on Test within an hour is displaying in Search results</b></p>"
$b = $b + "<style>"
$b = $b + "BODY{background-color:white;}"
$b = $b + "</style>"
del .\SearchHealthCheck\FileUploadTime.txt
}
$SearchFarmInitReport = @()
$SearchFarmReport = @()
$SearchCompReport = @()
$SearchFarmInitReport = Import-Csv .\SearchHealthCheck\SearchReports\$farm-SearchFarmInitStatus.txt
try
{
$SearchFarmReport = Import-Csv .\SearchHealthCheck\SearchReports\$farm-SearchFarmStatus.txt -ErrorAction SilentlyContinue
$SearchCompReport = Import-Csv .\SearchHealthCheck\SearchReports\$farm-SearchComponentStatus.txt -ErrorAction SilentlyContinue
}
catch
{}
if ($SearchFarmInitReport.Status -eq "Online")
{
	$farm + "," + $SearchFarmInitReport.ServiceApplication + "," + $SearchFarmInitReport.Status + "," + $SearchFarmReport.Status >> SearchFarmsReport.csv
}
else
{
	$farm + "," + $SearchFarmInitReport.ServiceApplication + "," + $SearchFarmInitReport.Status + ",DOWN" >> SearchFarmsReport.csv
}
if ($SearchCompReport)
{
	foreach ($scr in $SearchCompReport) {
		$farm + "," + $SearchFarmInitReport.ServiceApplication + "," + $scr.Component + "," + $scr.Server + "," + $scr.State + "," + $scr.Remarks >> SearchComponentsReport.csv
	}
}
}
$SearchFarmsReport = Import-Csv .\SearchFarmsReport.csv
if (Test-Path .\SearchComponentsReport.csv)
{
	$SearchComponentsReport = Import-Csv .\SearchComponentsReport.csv | where {$_.State -ne "Active"}
}

##Formatting HTML Table for Search Health Report
if ($SearchFarmsReport)
{
$strTableStartHTML6 = "<TABLE CELLSPACING=1 CELLPADDING=1 BORDER=1>`n<TR><TH>Farm</TH><TH>SearchAppName</TH><TH>Running Status</TH><TH>Health Status</TH></TR>"
$strTableBodyHTML6 = foreach ($sfr in $SearchFarmsReport) {
	$strRowStyleHTML6 = if ($sfr.HealthStatus -eq "OK") {" STYLE='color:green'"} elseif ($sfr.HealthStatus -eq "Degraded") {" STYLE='color:yellow'"} else {" STYLE='color:red'"}
	"<TR$strRowStyleHTML6><TD><b>$($sfr.Farm)</b></TD><TD><b>$($sfr.ApplicationName)</b></TD><TD><b>$($sfr.RunningStatus)</b></TD><TD><b>$($sfr.HealthStatus)</b></TD></TR>`n"
	}
$strTableEndHTML6 = "</TABLE>"
$strTable6 = $strTableStartHTML6 + $strTableBodyHTML6 + $strTableEndHTML6
$SearchFarmsReport = @()
}

if($SearchComponentsReport)
{
$strTableStartHTML7 = "<TABLE CELLSPACING=1 CELLPADDING=1 BORDER=1>`n<TR><TH>Farm</TH><TH>SearchAppName</TH><TH>Component</TH><TH>Server</TH><TH>State</TH><TH>Remarks</TH></TR>"
$strTableBodyHTML7 = foreach ($scsr in $SearchComponentsReport) {
	$strRowStyleHTML7 = " STYLE='color:red'"
	"<TR$strRowStyleHTML7><TD><b>$($scsr.Farm)</b></TD><TD><b>$($scsr.ApplicationName)</b></TD><TD><b>$($scsr.Component)</b></TD><TD><b>$($scsr.Server)</b></TD><TD><b>$($scsr.State)</b></TD><TD><b>$($scsr.Remarks)</b></TD></TR>`n"
	}
$strTableEndHTML7 = "</TABLE>"
$strTable7 = $strTableStartHTML7 + $strTableBodyHTML7 + $strTableEndHTML7
$SearchComponentsReport = @()
}
else
{
$strTableStartHTML7 = "<TABLE CELLSPACING=1 CELLPADDING=1 BORDER=1>`n<TR><TH>Search Component Level Report</TH></TR>"
$strRowStyleHTML7 = " STYLE='color:green'"
$strTableBodyHTML7 = "<TR$strRowStyleHTML7><TD><b>All Components are found to be Active</b></TD></TR>`n"
$strTableEndHTML7 = "</TABLE>"
$strTable7 = $strTableStartHTML7 + $strTableBodyHTML7 + $strTableEndHTML7
}
}

##Concatenating all HTML tables
$strTable = $strTable0 + $d + $strTable1 + $d + $strTable2 + $d + $strTable3 + $d + $strTable4 + $d + $strTable5 + $d + $strTable6 + $d + $strTable7

##Generating HTML Page
$html = ConvertTo-HTML -head $a$strTable$b$c -PostContent "Scripted by Vijayakumar Sankaramoorthy"
$BodyToString = $html | out-string
$html | Out-File .\Summary.htm
#Invoke-Expression .\Summary.htm


$Frm = "GroupM Automation<Automation@Groupm.com>"
#$Tomail = "pramodkn@in.ibm.com"
$Tomail = "GlobalSharepoint@wwpdl.vnet.ibm.com"
$Sub = "GroupM SharePoint PROD Daily HealthCheck Report - " + $reportdate
$smtpserver = "RELAY.NLBINT.INSIDEMEDIA.NET"
if (Test-Path .\SearchFarmsReport.csv)
{
$attachment = "DiskSpaceReport.csv","WebSiteReport.csv","AppPoolReport.csv","WinServiceReport.csv","WebURLReport.csv","Summary.htm","SearchFarmsReport.csv","SearchComponentsReport.csv"
}
else
{
$attachment = "DiskSpaceReport.csv","WebSiteReport.csv","AppPoolReport.csv","WinServiceReport.csv","WebURLReport.csv","Summary.htm"
}
$CCs = "SPDailyHealthCheck@sp.groupm.com" ,"SP_Issu.4pgdvy4ca2jnc446@u.box.com"
try
{
	Send-MailMessage -EA Stop -From $Frm -To $Tomail -Subject $Sub -SmtpServer $smtpserver -Body $BodyToString -BodyAsHtml -Attachments $attachment -CC $CCs
}
catch
{
	$strdate = Get-Date
	"===========================================" >> $Errorlog
	$strdate >> $Errorlog
	$_.Exception.Message >> $Errorlog
	"===========================================" >> $Errorlog
	"				" >> $Errorlog
}
#Summary
$htmlS = ConvertTo-HTML -head $a$strTable$b -PostContent "Scripted by Vijayakumar Sankaramoorthy"
$htmlS | Out-File .\LatestSummary.htm
#Invoke-Expression .\LatestSummary.htm
$TomailS = "SPDailyHealthCheck@sp.groupm.com"
$SubS = "SPHealthCheckSummary"
$BodyS = "Summary Report - " + $reportdate
$attachmentS = "LatestSummary.htm"
try
{
	Send-MailMessage -EA Stop -From $Frm -To $TomailS -Subject $SubS -SmtpServer $smtpserver -Body $BodyS -Attachments $attachmentS
}
catch
{
	$strdate = Get-Date
	"===========================================" >> $Errorlog
	$strdate >> $Errorlog
	$_.Exception.Message >> $Errorlog
	"===========================================" >> $Errorlog
	"				" >> $Errorlog
}

del LatestSummary.htm

New-Item Reports\$filedate -Type directory -ErrorAction SilentlyContinue | Out-Null
move .\DiskSpaceReport.csv .\Reports\$filedate\DiskSpaceReport_$date.csv
move .\WebSiteReport.csv .\Reports\$filedate\WebSiteReport_$date.csv
move .\AppPoolReport.csv .\Reports\$filedate\AppPoolReport_$date.csv
move .\WinServiceReport.csv .\Reports\$filedate\WinServiceReport_$date.csv
move .\WebURLReport.csv .\Reports\$filedate\WebURLReport_$date.csv
move .\Summary.htm .\Reports\$filedate\Summary_$date.htm
$srchcsvfiles = ls SearchHealthCheck\SearchReports | select name
if ($srchcsvfiles)
{
foreach ($srchcsvfile in $srchcsvfiles) {
$item = $srchcsvfile.Name -split "\."
$filename = $item[0]
$srcfile = ".\SearchHealthCheck\SearchReports\" + $filename + ".txt"
$dstfile = ".\Reports\" + $filedate + "\" + $filename + "_" + $date + ".txt"
move $srcfile $dstfile
}
}
if (Test-Path .\SearchFarmsReport.csv)
{
	move .\SearchFarmsReport.csv .\Reports\$filedate\SearchFarmsReport_$date.csv
}
if (Test-Path .\SearchComponentsReport.csv)
{
	move .\SearchComponentsReport.csv .\Reports\$filedate\SearchComponentsReport_$date.csv
}


$strdate = Get-Date
$strdate >> $AuditLog
"Script Execution is Completed" >> $AuditLog
"===========================================" >> $AuditLog
"				" >> $AuditLog
