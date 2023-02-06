#############################################################################
# Author:Madhiri Jayavishnu
#
#
# Description: Script to collect Server Information in HTML amd EXCEL File
#############################################################################
#############################################################################
$outputpath="E:\powershell"
Ree-Item -Path "$outputpath\Servers Information.html" -ErrorAction SilentlyContinue #the path which is given in output path with file name to delete 
Remove-Item -Path "$outputpath\Servers Information.xlsx"  -Force -ErrorAction SilentlyContinue
Remove-Variamovble -Name check,test,auto -ErrorAction SilentlyContinue
##checking the module
 
try {
 
if (Get-Module PSExcel -ListAvailable) #ImportExcel
{
Import-Module PSExcel -ErrorAction SilentlyContinue
}
else
{
Install-Module -Name PSExcel -Scope CurrentUser
Import-Module PSExcel -ErrorAction SilentlyContinue
}
}
catch
{
Write-Host "Check the Module and version"
}


 
$servername = @("localhost","13.233.247.69") #list of input servers
 


$finalout=foreach ($computername in $servername)
{
$test=Test-Connection -ComputerName $computername -count 1 -ErrorAction SilentlyContinue

if( $test)
{
if($computername -match "localhost" ){
$session=New-PSSession -ComputerName $computername  -EnableNetworkAccess
write-host "match if"
}
else{
Write-Host "match else"

$cred = get-credential -UserName administrator -Message "Please Enter Credential for $computername "
$session=New-PSSession -ComputerName $computername -Credential $cred
}
Invoke-Command -Session $session -ScriptBlock{ 

$AVGProc = Get-WmiObject -Class win32_processor  | Measure-Object -property LoadPercentage -Average | Select Average
$drivepercent = gwmi -Class win32_operatingsystem |
Select-Object @{Name = "MemoryUsage"; Expression = {“{0:N2}” -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory)*100)/ $_.TotalVisibleMemorySize) }}
$os=(get-Wmiobject win32_operatingsystem ).version
$cpu_details=Get-WmiObject –class Win32_processor  | Select Name,DeviceID,NumberOfCores,NumberOfLogicalProcessors,NumberOfProcessors
$osname=(get-Wmiobject win32_operatingsystem ).caption
$freespace = Get-WmiObject win32_logicaldisk | select @{n="ComputerName";e={$computername}},
@{n="DriveName";e={$_.DeviceID}},@{n="Total_Space";e={"$([math]::Round($_.Size/1gb,2)) GB"}},
@{n="Freespace";e={"$([math]::Round($_.Freespace/1gb,2)) GB"}},
@{n = 'UsedSpace';e={ "{0:N2}" -f $([math]::Round(($_.Size - $_.FreeSpace)*100) / $_.size)}}
$auto=Get-Service | where{$_.StartType -eq "Automatic" -and $_.Status -eq "Stopped" } | Select Name, Starttype, status
foreach($ii in $freespace){
[PSCustomObject]@{
"Server Name" = $using:computername
"pingingstatus"= "UP"
"Operating System"=$osname
"OS version"=$os
"DriveName" = $ii.DriveName
"TotalSpace(GB)" = $ii.Total_Space
"Freespace(GB)" = $ii.Freespace
"DriveUsedPercent"= $ii.UsedSpace
"MemLoad" = "$($drivepercent.MemoryUsage)%"
"CPULoad" = "$($AVGProc.Average)%"
}
}

foreach($vv in $auto){

[PSCustomObject]@{
"Service Name" = $vv.Name
"Start type"= $vv.Starttype
"Status"=$vv.status

}

}
}
}
else
{
 [PSCustomObject] @{
"Server Name" = $computername
"pingingstatus"= "Down"
"Operating System"="None"
"OS version"="None"
"DriveName" = "None"
"TotalSpace(GB)" = "None"
"Freespace(GB)" = "None"
"DriveUsedPercent"= "None"
"MemLoad" = "None"
"CPULoad" = "None"
"Service Name" = "None"
"Start type"= "None"
"Status"= "None"

}
}

}
 
$check= $finalout|select 'Server Name','pingingstatus' -Unique
$kreach=0
$ireach=0
foreach($j in $check.pingingstatus){
if($j -match "UP"){
$ireach++
}
if($j -match "Down"){
$kreach++
}
}





###############################HTML OUTPUT########################################
$z="<html>
<head>
<style>
table{
font-size:12px
}
#Header{font-family:'Trebuchet MS', Arial, Helvetica, sans-serif;width:100%;border-collapse:collapse;}
#Header td, #Header th {font-size:14px;border:1px solid #DDDDDD;padding:5px 7px 5px 7px;}
#Header th {font-size:14px;text-align:left;padding-top:10px;padding-bottom:10px;background-color:#D6D6D6;color:blue;}
#Header tr.alt td {color:#000;background-color:#FAF2D3;}
</style>
</head>
<div class='col-md-12'>
<h4 style = 'color:orange;'>Server Details</h4>
<Table border=1 cellpadding=0 cellspacing=0 id=Header>
<thead style='color:blue;'>
<tr>
<th>Server Name</th>
<th>pingingstatus</th>
<th>Operating System</th>
<th>OS version</th>
<th>DriveName</th>
<th>TotalSpace(GB)</th>
<th>Freespace(GB)</th>
<th>DriveUsedPercent</th>
<th>MemLoad</th>
<th>CPULoad</th>
</tr>
</thead>
<tbody>"
$z += $finalout|Where-Object{$_.'Server Name' -and $_.'Pingingstatus' -and $_.'os version' }| foreach{
"<tr>
<td>$($_.'server name')</td>
<td>$($_.'pingingstatus')</td>
<td>$($_.'operating system' ) </td>
<td>$($_.'os version')</td>
<td>$($_.'Drivename')</td>
<td>$($_.'TotalSpace(GB)' ) </td>
<td>$($_.'Freespace(GB)')</td>"
if($_.DriveUsedPercent -gt 90){
"<td style=background-color:red;>$($_.'DriveUsedPercent')</td>"}
elseif($_.DriveUsedPercent -gt 85 -and $_.DriveUsedPercent -lt 90){
"<td style=background-color:yellow;>$($_.'DriveUsedPercent')</td>"}
else{
"<td style=background-color:green;>$($_.'DriveUsedPercent')</td>"}
if($_.MemLoad -gt 90){
"<td style=background-color:red;>$($_.'MemLoad' ) </td>"}
elseif($_.MemLoad -gt 85 -and $_.MemLoad -lt 90){
"<td style=background-color:yellow;>$($_.'MemLoad' ) </td>"}
else{
"<td style=background-color:green;>$($_.'MemLoad' ) </td>"}
if($_.CPULoad -gt 90){
"<td style=background-color:red;>$($_.'CPULoad')</td>"}
elseif($_.CPULoad -gt 85 -and $_.CPULoad -lt 90) {
"<td style=background-color:yellow;>$($_.'CPULoad')</td>"}
else{
"<td style=background-color:green>$($_.'CPULoad')</td>"}
}
$z += "<tbody>
</table>
</div>
</div>
</body>
</html>
"
$y="<html>
<head>
<style>
table{
font-size:12px
}
#Header{font-family:'Trebuchet MS', Arial, Helvetica, sans-serif;width:100%;border-collapse:collapse;}
#Header td, #Header th {font-size:14px;border:1px solid #DDDDDD;padding:5px 7px 5px 7px;}
#Header th {font-size:14px;text-align:left;padding-top:10px;padding-bottom:10px;background-color:#D6D6D6;color:blue;}
#Header tr.alt td {color:#000;background-color:#FAF2D3;}
</style>
</head>
<body>
<center>
<h4 style = 'color:orange;'>Automatic Services Status</h4>
<Table border=1 cellpadding=0 cellspacing=0 id=Header>
<thead style='color:blue;'>
</center>
<tr>
<th>Server Name</th>
<th>Service Name</th>
<th>Start Type</th>
<th>Status</th>
</tr>
</thead>
<tbody>"
$y += $finalout | %{
if($_.status -match "stopped"){
"<tr>
<td>$($_.'PSComputerName')</td>
<td>$($_.'Service Name')</td>
<td>$($_.'Start type')</td>
<td style=background-color:red;>$($_.'Status' )</td>
</tr>"
}
elseif($_.'pingingstatus' -match "DOWN")
{
"<tr style='color:red'>
<td>$($_.'Server Name')</td>
<td>$($_.'Service Name')</td>
<td>$($_.'Start type')</td>
<td>$($_.'status' )</td>
</tr>"
}
}
$y += "<tbody>
</table>
</div>
</div>
</body>
</html>
"
$c="<html>
<head>
<style>
table{
font-size:12px
}
#Header{font-family:'Trebuchet MS', Arial, Helvetica, sans-serif;width:100%;border-collapse:collapse;}
#Header td, #Header th {font-size:14px;border:1px solid #DDDDDD;padding:5px 7px 5px 7px;}
#Header th {font-size:14px;text-align:left;padding-top:10px;padding-bottom:10px;background-color:#D6D6D6;color:blue;}
#Header tr.alt td {color:#000;background-color:#FAF2D3;}
</style>
<center>
<h1 style='color:darkblue'>Server Health Check Report</h1>
<center>
<h3 style='text-align: left;'><b>$(get-date -Format 'dddd, MMMM dd, yyyy hh:mm:ss tt')</b></h3>
<div class ='col-md-12'>
<h3></h3>
</head>
<div class ='col-md-12'>
<h3></h3>
<div class='col-md-12'>
<center>
<h4 style = 'color:orange;'>Server Reachability</h4>
<Table border=1 cellpadding=0 cellspacing=0 id=Header>
<thead style='color:blue;'>
</center>
<tr>
<th>Total Servers</th>
<th>Reachable Servers</th>
<th> Not Reachable Servers </th>
</tr>
</thead>
<tbody>"
$c +=
"<tr>
<td>$($ireach + $kreach)</td>
<td>$($ireach)</td>
<td>$($kreach) </td>
</tr>"
$c += "<tbody>
</table>
</div>
</div>
</body>
</html>
"
$c,$z,$y| Out-File "$outputpath\Servers Information.html"
$total=$ireach+$kreach
$server=@()
$server += [pscustomobject]@{
TotalServers= "$total"
ReachbleServers = "$ireach"
NotreachbleServers = "$kreach"
}



##########################EXCEL Output#####################################
$finalout|?{$_.'Server Name' -and $_.'Pingingstatus' -and $_.'os version' }`
|Select-Object -Property 'Server Name','pingingstatus','OS version','DriveName','TotalSpace(GB)','Freespace(GB)','DriveUsedPercent','MemLoad','CPULoad'`
| Export-XLSX -Path "$outputpath\Servers Information.xlsx" -WorksheetName ServerInfo 
$finalout|?{$_.'Service Name' -and $_.'Start type' -and $_.'Status'}|Select-Object -Property 'PsComputerName','Service Name','Start type','Status' |Export-XLSX -Path "$outputpath\Servers Information.xlsx" -WorksheetName ServicesInfo 
$server | Export-XLSX -Path "$outputpath\Servers Information.xlsx" -WorksheetName ServerReachability 



############################################# mail sending ############################################


##sending mail if cpu or memory values greater than 90 percent
foreach($serverinfo in $finalout){

if($serverinfo.CPULoad -gt 90 -or $serverinfo.MemLoad -gt 90){

write-host "Threshold Value Exceeded.Details sent in mail" -BackgroundColor Red
$email = "madhiri.jayavishnu2003@gmail.com"
$password = "zzfmrrhvpsjjuoqw"
$smtpServer = "smtp.gmail.com"
$smtpPort = "587"
$to = "madhiri.jayavishnu2003@gmail.com"
$subject = "HTML attachment email"
$attachment1 = "$outputpath\Servers Information.html"
$attachment2="$outputpath\Servers Information.xlsx"

$body = Get-Content $attachment1 | Out-String

$message = New-Object System.Net.Mail.MailMessage $email, $to, $subject, $body
$message.IsBodyHTML = $true
$attachment = New-Object System.Net.Mail.Attachment($attachment2)
$message.Attachments.Add($attachment)

$smtp = New-Object Net.Mail.SmtpClient($smtpServer, $smtpPort)
$smtp.UseDefaultCredentials = $false
$smtp.Credentials = New-Object System.Net.NetworkCredential($email, $password)
$smtp.EnableSsl = $true
$smtp.Send($message)

}

}
