<#
Automation_Tracert_Report_OptionA_Fast_v6.ps1
Parallel traceroute monitoring with SSID naming by IP range.
#>
$ErrorActionPreference='Continue'
$Desktop=[Environment]::GetFolderPath('Desktop')
try {
 $IPv4=(Get-NetIPAddress -AddressFamily IPv4 | Where-Object {$_.IPAddress -notlike '169.254*' -and $_.IPAddress -ne '127.0.0.1'} | Select-Object -First 1 -ExpandProperty IPAddress)
} catch {$IPv4='Unavailable'}
$WiFiSSID='Unknown_Network'
if($IPv4 -like '192.168.10.*'){ $WiFiSSID='Corporate_Office' }
elseif($IPv4 -like '192.168.11.*'){ $WiFiSSID='Corporate_Executive' }
elseif($IPv4 -like '192.168.12.*'){ $WiFiSSID='Corporate_Guest' }
elseif($IPv4 -like '192.168.13.*'){ $WiFiSSID='Corporate_WIFI' }
$DateStamp=Get-Date -Format 'MMMM_dd_yyyy'
$ExcelFile=Join-Path $Desktop "Automation_Tracert_${WiFiSSID}_Report_$DateStamp.xlsx"
if(Test-Path $ExcelFile){Remove-Item $ExcelFile -Force -ErrorAction SilentlyContinue}
if(-not (Get-Module -ListAvailable -Name ImportExcel)){
 if(-not(Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)){Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null}
 Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber
}
Import-Module ImportExcel -Force
try{$OS=(Get-CimInstance Win32_OperatingSystem).Caption}catch{$OS='Unavailable'}
$Targets=@(
@{Sheet='Drive';Host='drive.google.com'},
@{Sheet='Meet';Host='meet.google.com'},
@{Sheet='Outlook';Host='outlook.office365.com'},
@{Sheet='Microsoft';Host='www.microsoft.com'},
@{Sheet='SharePoint';Host='microsoft.sharepoint.com'},
@{Sheet='OneDrive';Host='onedrive.live.com'})
foreach($t in $Targets){Start-Job -Name $t.Sheet -ArgumentList $t.Host -ScriptBlock {param($h) tracert -d -h 20 -w 1000 $h 2>&1} | Out-Null}
Get-Job | Wait-Job | Out-Null
foreach($t in $Targets){
 $rows=@()
 $rows += [pscustomobject]@{Property='Computer Name';Value=$env:COMPUTERNAME}
 $rows += [pscustomobject]@{Property='Logged In User';Value="$env:USERDOMAIN\$env:USERNAME"}
 $rows += [pscustomobject]@{Property='IPv4 Address';Value=$IPv4}
 $rows += [pscustomobject]@{Property='WiFi SSID';Value=$WiFiSSID}
 $rows += [pscustomobject]@{Property='Operating System';Value=$OS}
 $rows += [pscustomobject]@{Property='Date';Value=(Get-Date -Format 'yyyy-MM-dd')}
 $rows += [pscustomobject]@{Property='Time';Value=(Get-Date -Format 'HH:mm:ss')}
 foreach($line in (Receive-Job -Name $t.Sheet)){ $rows += [pscustomobject]@{Property='Traceroute Output';Value=$line.ToString()} }
 $rows | Export-Excel -Path $ExcelFile -WorksheetName $t.Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
}
Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
Write-Host "Report saved to: $ExcelFile"
Read-Host 'Press Enter to close'
