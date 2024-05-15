########################################################################################
#                          Windows Update issues resolver.                             #
#                         Written by Chaminda Prasad - V1.3                            #
########################################################################################

#Script Logging
$transcriptpath = "$env:SystemDrive\PowerShell_Transcript"
$transcriptpathstatus = Test-Path -Path $transcriptpath

if ( -not $transcriptpathstatus -eq 'True') {
  New-Item -ItemType 'Directory' -Path $transcriptpath | Out-Null;
  Start-Transcript -Path "$transcriptpath\Update_Resolver_Transcript.log" -Append -ErrorAction 'SilentlyContinue' | Out-Null;
}
else {
  Start-Transcript -Path "$transcriptpath\Update_Resolver_Transcript.log" -Append -ErrorAction 'SilentlyContinue' | Out-Null;
}

#Fetch paths information for temporary data and demons 
$Path = "$env:windir\SoftwareDistribution\", "$env:windir\system32\catroot2";
$Service = "wuauserv", "BITS" , "cryptsvc" 
$msii = "msiserver"
$Message = 'Windows Update componets have been reset and update cache has been removed successfully'
$winupreg = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\PolicyManager\current\device\Update'
#$regbackuppath = 'C:\Registry Backup'
$reg_OK = 'Windows Updates settings Registry modifications were made'
$regstatusmg = 'Windows update registry key does not exist. Hence, No registry modifications were made.'

#Fetch services which need for windows updates to stop and delete temp data
$Svcstatus = Get-Service -Name $Service; 
if ($Svcstatus.Status -eq 'Running') { 
  Foreach-Object { $Svcstatus.Name | Stop-Service -Verbose -ErrorAction 'SilentlyContinue' } 
  Get-ChildItem -Path $Path | Remove-Item -Recurse -Verbose -ErrorAction 'SilentlyContinue'
}
else {
  Get-ChildItem -Path $Path | Remove-Item -Recurse -Verbose -ErrorAction 'SilentlyContinue'
}

<#
#Export registry key as a backup
$folderstatus = Test-path -path $regbackuppath

if ( -not $folderstatus -eq 'true') 
{
  New-Item -ItemType 'Directory' -Path $regbackuppath -ErrorAction 'Stop';
}
#>

#Modify update components in registry
$regstatus = Test-Path -Path "Registry::$winupreg"

if ($regstatus -eq 'True') {
  Remove-Item -Path "Registry::$winupreg" -Recurse;
  $reg_OK_MG = $reg_OK
}
else {
  $regmg = $regstatusmg;
}

Start-Service -Name $Service -Verbose -ErrorAction 'SilentlyContinue'

#Set windows updates core services to run on Startup
Get-Service -Name $Service | Set-Service -StartupType 'Automatic' -Verbose -ErrorAction 'Stop';

$installer = Get-Service -name $msii
if ( $installer.Status -eq 'Stopped') {
  $installer | Start-Service -Verbose -ErrorAction 'Stop';
}


Write-Host $reg_OK_MG -ForegroundColor 'Magenta';
Write-Host $regmg -ForegroundColor 'Magenta';
Write-Host $Message -ForegroundColor 'Green';

Stop-Transcript -Verbose

########################################################################################
#                             ##  END of the script ##                                 # 
########################################################################################