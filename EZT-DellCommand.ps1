<#  
    .Name
    EZT-DellCommand
    
    .Version 
    0.23

    .SYNOPSIS
    Automates checking for or installing Dell driver, firmware and other updates available using DellCommand

    .DESCRIPTION
       
    .Configurable Variables:

    .EXAMPLE
    \EZT-DellCommand.ps1

    .OUTPUTS
    System.Management.Automation.PSObject

    .Credits
    Write-Color           - https://github.com/EvotecIT/PSWriteColor
    HTML Dashboard        - https://github.com/EvotecIT/PSWriteHTML

    .NOTES
    Author: EZTechhelp
    DellCommand CLI Commands: https://www.dell.com/support/manuals/en-us/command-update/dellcommandupdate_rg/dell-command--update-cli-commands?guid=guid-92619086-5f7c-4a05-bce2-0d560c15e8ed
#> 

#############################################################################
#region Configurable Script Parameters
#############################################################################

#---------------------------------------------- 
#region Dell Command Update Variables
#----------------------------------------------
$runDellCommand = $true #enables download and execution of dell command. Leave enabled otherwise this script doesnt do very much
$installDellCommand_Drivers_Types = 'none' #enter update types to install if detected, comma seperated. Adding 'All' takes priority and will install all updates. Available Types: bios, firmware, driver, application, others, All
$InstallDellCommand_Drivers_reboot = $false #enables automatic reboot of system after updates are installed
$DownloadURL_Dell = 'https://dl.dell.com/FOLDER06986472M/2/Dell-Command-Update-Application-for-Windows-10_DF2DT_WIN_4.1.0_A00.EXE' #download url for the dell command installer file
$DownloadLocation_Dell = 'C:\DellCommand' #working directory where the dell command installer file will be downloaded and executed
#---------------------------------------------- 
#endregion Dell Command Update Variables
#----------------------------------------------

#---------------------------------------------- 
#region Report Variables
#----------------------------------------------
$dashboardreport = $true #enables generating an HTML Report containing all Dell Updates found
$opendashboardreport = $true #enables immediately opening the report in the default browser after its created. Disable if using through RMM or other remote tools
$save_report_location = 'c:\Dell Reports' #directory where the HTML dashboard report should be saved. If left blank, the report will be deleted after its generated or emailed
$email_report = $false #enables sending the HTML report via email. Configure email variables if enabled
$email_logs = $false # enables attaching the script log file to the email sent with the report. 
#---------------------------------------------- 
#endregion Report Variables
#----------------------------------------------

#---------------------------------------------- 
#region Email Variables
#----------------------------------------------
$SmtpPassword = 'password'
$SmtpUser = 'user@email.com'
$SmtpPort = '587'
$SmtpServer = 'smtp.office365.com'
$MailFrom = 'from@email.com'
$MailTo = 'to@email.com'
$Subject = 'Dell Command Report'
#---------------------------------------------- 
#endregion Email Variables
#----------------------------------------------

#---------------------------------------------- 
#region Log Variables
#----------------------------------------------
$enablelogs = $true # enables creating a log file of executed actions, run history and errors
$logfile_directory = 'C:\Logs\' # directory where log file should be created if enabled
#---------------------------------------------- 
#endregion Log Variables
#----------------------------------------------

#---------------------------------------------- 
#region Global Variables - DO NOT CHANGE UNLESS YOU KNOW WHAT YOU'R DOING
#----------------------------------------------
$stopwatch = [system.diagnostics.stopwatch]::StartNew() #starts stopwatch timer 
$Required_modules = 'PSWriteHTML','PSWriteColor','PowerShellGet' #these modules are automatically installed and imported if not already
$update_modules = $false # enables checking for and updating all required modules for this script. Potentially adds a few seconds to total runtime but ensures all modules are the latest
$force_modules = $false # enables installing and importing of a module even if it is already. Should not be used unless troubleshooting module issues 
$logdateformat = 'MM/dd/yyyy h:mm:ss tt' # sets the date/time appearance format for log file and console messages
#"DellBIOSProvider"
#---------------------------------------------- 
#endregion Global Variables - DO NOT CHANGE UNLESS YOU KNOW WHAT YOU'R DOING
#----------------------------------------------
#############################################################################
#endregion Configurable Script Parameters
#############################################################################

#############################################################################
#region global functions - Must be run first and/or are script agnostic
#############################################################################
#---------------------------------------------- 
#region Get-ThisScriptInfo Function
#----------------------------------------------
function Get-ThisScriptInfo
{
  $Invocation = (Get-Variable MyInvocation -Scope 1).Value
  $ScriptPath = $PSCommandPath
  if(!$ScriptPath)
  {   
    $ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand
    $thisScript = @{File = Get-ChildItem $ScriptPath; Contents = $Invocation.MyCommand}
  }
  else
  {$thisScript = @{File = Get-ChildItem $ScriptPath; Contents = $Invocation.MyCommand.ScriptContents}}
  If ($thisScript.Contents -Match '^\s*\<#([\s\S]*?)#\>') 
  {$thisScript.Help = $Matches[1].Trim()}
  [RegEx]::Matches($thisScript.Help, "(^|[`r`n])\s*\.(.+)\s*[`r`n]|$") | ForEach-Object {
    If ($Caption) 
    {$thisScript.$Caption = $thisScript.Help.SubString($Start, $_.Index - $Start)}
    $Caption = $_.Groups[2].ToString().Trim()
    $Start = $_.Index + $_.Length
  }
  $thisScript.Title = $thisScript.Synopsis.Trim().Split("`r`n")[0].Trim()
  
  $thisScript.Version = $thisScript.Version.Trim()
  
  $thisScript.Name = $thisScript.Name.Trim()
  
  $thisScript.credits = $thisScript.credits -split("`n") | ForEach-Object {$_.trim()}
  $thisScript.SYNOPSIS = $thisScript.SYNOPSIS -split("`n") | ForEach-Object {$_.trim()}
  $thisScript.Description = $thisScript.Description -split("`n") | ForEach-Object {$_.trim()}
  $thisScript.Notes = $thisScript.Notes -split("`n") | ForEach-Object {$_.trim()}
  $thisScript.Path = $thisScript.File.FullName; $thisScript.Folder = $thisScript.File.DirectoryName; $thisScript.BaseName = $thisScript.File.BaseName
  $thisScript.Arguments = (($Invocation.Line + ' ') -Replace ('^.*\\' + $thisScript.File.Name.Replace('.', '\.') + "['"" ]"), '').Trim()
  return $thisScript
}
$thisScript = Get-ThisScriptInfo
$Script_Temp_Folder = "$env:TEMP\$($thisScript.Name)"
if(!(Test-Path $Script_Temp_Folder))
{
  try
  {$null = New-Item $Script_Temp_Folder -ItemType Directory -Force}
  catch
  {Write-EZLogs "[ERROR] Exception creating script temp directory $Script_Temp_Folder - $_" -ShowTime -color Red}
}
Write-Host "#### Executing $($thisScript.Name) - v$($thisScript.Version) ####" -ForegroundColor Black -BackGroundColor yellow
Write-Host " | $($thisScript.SYNOPSIS)"
#---------------------------------------------- 
#endregion Get-ThisScriptInfo Function
#----------------------------------------------

#---------------------------------------------- 
#region Begin Logging
#----------------------------------------------
if ($enablelogs)
{  
  $logfile = [System.IO.Path]::Combine($logfile_directory, "$($thisScript.Name)-$($thisScript.Version).log")
  if (!(Test-Path -LiteralPath $logfile 2> $null))
  {$null = New-Item -Path $logfile_directory -ItemType directory -Force}
  $OriginalPref = $ProgressPreference
  $ProgressPreference = 'SilentlyContinue'
  $Computer_Info = Get-WmiObject Win32_ComputerSystem | Select-Object *
  $OS_Info = Get-CimInstance Win32_OperatingSystem | Select-Object *
  $CPU_Name = (Get-WmiObject Win32_Processor -Property 'Name').name
  $ProgressPreference = $OriginalPref
  $logheader = @"
`n###################### Logging Enabled ######################
Script Name          : $($thisScript.Name)
Synopsis             : $($thisScript.SYNOPSIS)
Log File             : $logfile
Version              : $($thisScript.Version)
Current Username     : $env:username
Powershell           : $($PSVersionTable.psversion)($($PSVersionTable.psedition))
Computer Name        : $env:computername
Operating System     : $($OS_Info.Caption)($($OS_Info.Version))
CPU                  : $($CPU_Name)
RAM                  : $([Math]::Round([int64]($computer_info.TotalPhysicalMemory)/1MB,2)) GB (Available: $([Math]::Round([int64]($OS_Info.FreePhysicalMemory)/1MB,2)) GB)
Manufacturer         : $($computer_info.Manufacturer)
Model                : $($computer_info.Model)
Serial Number        : $((Get-WmiObject Win32_BIOS | Select-Object SerialNumber).SerialNumber)
Domain               : $($computer_info.Domain)
Install Date         : $($OS_Info.InstallDate)
Last Boot Up Time    : $($OS_Info.LastBootUpTime)
Local Date/Time      : $($OS_Info.LocalDateTime)
Windows Directory    : $($OS_Info.WindowsDirectory)
###################### Logging Started - [$(Get-Date)] ##########################
"@
  Write-Output $logheader | Out-File -FilePath $logfile -Encoding unicode -Append
  Write-Host " | Logging is enabled. Log file: $logfile"
}
#---------------------------------------------- 
#endregion Begin Logging
#----------------------------------------------

#---------------------------------------------- 
#region Load-Modules Function
#----------------------------------------------
function Load-Modules ($modules,$force,$update) 
{
  #Make sure we can download and install modules through NuGet
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  if (Get-PackageProvider | Where-Object {$_.Name -eq 'Nuget'}) 
  {Write-Output ' | Required PackageProvider Nuget is installed.' -OutVariable message;if($enablelogs){$message | Out-File -FilePath $logfile -Encoding unicode -Append}}
  else
  {
    try
    {
      Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
      Register-PackageSource -Name nuget.org -Location https://www.nuget.org/api/v2 -ProviderName NuGet
    }
    catch
    {Write-Error "[Load-Module ERROR] $_`n" -ErrorVariable messageerror;if($enablelogs){$messageerror | Out-File -FilePath $logfile -Encoding unicode -Append}}
  }
  #Install latest version of PowerShellGet
  if (Get-Module 'PowershellGet' | Where-Object {$_.Version -lt 2.2.5})
  {
    Write-Output ' | PowershellGet version too low, updating to 2.2.5' -OutVariable message;if($enablelogs){$message | Out-File -FilePath $logfile -Encoding unicode -Append}
    Install-Module -Name 'PowershellGet' -MinimumVersion 2.2.5 -Force 
  }
  foreach ($m in $modules)
  {  
    if (Get-Module | Where-Object {$_.Name -eq $m}) 
    {
      Write-Output " | Required Module $m is imported." -OutVariable message;if($enablelogs){$message | Out-File -FilePath $logfile -Encoding unicode -Append}
      if ($force)
      {
        Write-Output " | Force parameter applied - Installing $m" -OutVariable message;if($enablelogs){$message | Out-File -FilePath $logfile -Encoding unicode -Append}
        Install-Module -Name $m -Scope AllUsers -Force -Verbose 
      }
    }
    else 
    {
      #If module is not imported, but available on disk set module autoloading when needed/called 
      if (Get-Module -ListAvailable | Where-Object {$_.Name -eq $m}) 
      {
        $PSModuleAutoLoadingPreference = 'ModuleQualified'
        if($update)
        {
          Write-Output " | Updating module: $m" -OutVariable message;if($enablelogs){$message | Out-File -FilePath $logfile -Encoding unicode -Append}
          Update-Module -Name $m -Force -ErrorAction Continue
        }
        if($force)
        {
          if($enablelogs){Write-Output " | Force parameter applied - Importing $m" | Out-File -FilePath $logfile -Encoding unicode -Append}
          Import-Module $m -Verbose -force -Scope Global
        }
      }
      else 
      {
        #If module is not imported, not available on disk, but is in online gallery then install and import
        if (Find-Module -Name $m | Where-Object {$_.Name -eq $m}) 
        {
          try
          {
            Install-Module -Name $m -Force -Verbose -Scope AllUsers -AllowClobber
            Import-Module $m -Verbose -force -Scope Global
          }
          catch
          {Write-Error "[Load-Module ERROR] $_" -ErrorVariable messageerror;if($enablelogs){$messageerror | Out-File -FilePath $logfile -Encoding unicode -Append}}      
        }
        else 
        {
          #If module is not imported, not available and not in online gallery then abort
          Write-Error "[Load-Module ERROR] Required module $m not imported, not available and not in online gallery, exiting." -ErrorVariable messageerror;if($enablelogs){$messageerror | Out-File -FilePath $logfile -Encoding unicode -Append}
          EXIT 1
        }
      }
    }
  } 
}
Load-Modules -modules $Required_modules -force:$force_modules -update:$update_modules
#---------------------------------------------- 
#endregion Load-Modules Function
#----------------------------------------------

#---------------------------------------------- 
#region Write-EZLogs Function
#----------------------------------------------
Function Write-EZLogs 
{
  [CmdletBinding(DefaultParameterSetName = 'text')]
  param (
    [string]$text,
    [switch]$VerboseDebug,
    [switch]$enablelogs,
    [string]$logfile = $logfile,
    [switch]$Warning,
    [string]$DateTimeFormat = $logdateformat,
    [ValidateSet('Black','Blue','Cyan','Gray','Green','Magenta','Red','White','Yellow','DarkBlue','DarkCyan','DarkGreen','DarkMagenta','DarkRed','DarkYellow')]
    [string]$color = 'white',
    [switch]$showtime,
    [switch]$logtime,
    [switch]$NoNewLine,
    [ValidateSet('Black','Blue','Cyan','Gray','Green','Magenta','Red','White','Yellow','DarkBlue','DarkCyan','DarkGreen','DarkMagenta','DarkRed','DarkYellow')]
    [string]$BackgroundColor,
    [int]$linesbefore,
    [int]$linesafter
  )
  if($showtime -and !$logtime){$logtime = $true}else{$logtime = $false}
  if($BackgroundColor){$BackgroundColor_param = $BackgroundColor}else{$BackgroundColor_param = $null}
  if($linesBefore){$text = "`n$text"}
  if($linesAfter){$text = "$text`n"}
  if($enablelogs)
  {
    if($VerboseDebug -and $warning)
    {
      $tmp = [System.IO.Path]::GetTempFileName();
      Write-Warning ($wrn = "[$(Get-Date -Format $DateTimeFormat)] $text");Write-Output "[$(Get-Date -Format $DateTimeFormat)] [WARNING] $wrn" | Out-File -FilePath $logfile -Encoding unicode -Append -Verbose:$VerboseDebug 4>$tmp
      $result = "[DEBUG] $(Get-Content $tmp)" | Out-File $logfile -Encoding unicode -Append;Remove-Item $tmp   
    }
    elseif($Warning)
    {Write-Color -showtime -NoNewLine -DateTimeFormat:$DateTimeFormat;Write-Warning ($wrn = "$text");Write-Output "[$(Get-Date -Format $DateTimeFormat)] [WARNING] $wrn" | Out-File -FilePath $logfile -Encoding unicode -Append}
    else
    {Write-Color $text -color:$color -showtime:$showtime -LogFile:$logfile -LogTime:$logtime -NoNewLine:$NoNewLine -DateTimeFormat:$DateTimeFormat -BackGroundColor $BackgroundColor_param}
  }
  else
  {
    if($warning)
    {Write-Color -showtime -NoNewLine -DateTimeFormat:$DateTimeFormat;Write-Warning ($wrn = "$text")}
    else
    {Write-Color $text -color:$color -showtime:$showtime -NoNewLine:$NoNewLine -DateTimeFormat:$DateTimeFormat -BackGroundColor $BackgroundColor_param}     
  }
}
#---------------------------------------------- 
#endregion Write-EZLogs Function
#----------------------------------------------

#---------------------------------------------- 
#region Use Run-As Function
#----------------------------------------------
function Use-RunAs 
{    
  # Check if script is running as Adminstrator and if not use RunAs 
  # Use Check Switch to check if admin 
  # http://gallery.technet.microsoft.com/scriptcenter/63fd1c0d-da57-4fb4-9645-ea52fc4f1dfb
    
  param([Switch]$Check) 
  $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator') 
  if ($Check) { return $IsAdmin }     
  if ($MyInvocation.ScriptName -ne '') 
  {  
    if (-not $IsAdmin)  
    {  
      try 
      {  
        $arg = "-file `"$($MyInvocation.ScriptName)`"" 
        Start-Process "$psHome\powershell.exe" -Verb Runas -ArgumentList $arg -ErrorAction 'stop'  
      } 
      catch 
      { 
        Write-Warning 'Error - Failed to restart script with runas'  
        break               
      } 
      exit # Quit this session of powershell 
    }  
  }  
  else  
  {  
    Write-EZLogs 'Script must be saved as a .ps1 file first' -showtime -LogFile $logfile -LinesAfter 1 -Warning  
    break  
  }  
}
#---------------------------------------------- 
#endregion Use Run-As Function
#----------------------------------------------
#############################################################################
#endregion global functions
#############################################################################

#############################################################################
#region Core functions - The primary functions specific to this script
#############################################################################

#---------------------------------------------- 
#region Send Email Function
#---------------------------------------------- 
Function Send-Email
{
  param 
  (   
    [switch]$enablelogs,
    [String]$dashboardreport_file = $dashboardreport_file,
    [String]$logfile = $logfile,
    [String]$SmtpUser,
    [String]$SmtpPort,
    [String]$SmtpServer,
    [String]$MailFrom,
    [String]$MailTo,
    [String]$Subject,
    $SmtpPassword
  )
  $SmtpPassword =  ConvertTo-SecureString -AsPlainText -Force $SmtpPassword
  [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
  # Create the email.
  $Subject = "$Subject - $env:computername"
  Write-EZLogs "Creating email with SUBJECT ($subject) FROM ($mailfrom) TO ($mailto)" -ShowTime -enablelogs:$enablelogs
  $email = New-Object System.Net.Mail.MailMessage($MailFrom , $MailTo)
  $email.Subject = $Subject
  $email.IsBodyHtml = $true
  $email.Body = @"
See attached report
"@

  if ($dashboardreport_file)
  {
    Write-EZLogs -text "Attaching HTML Dashboard Report File ($dashboardreport_file)" -ShowTime -enablelogs:$enablelogs
    $email.attachments.add($dashboardreport_file)
  }
  if($enablelogs)
  {
    $emaillog =  [System.IO.Path]::Combine($env:temp, "$($thisScript.Name)-$($thisScript.Version)-EM.log")
    Write-EZLogs -text "Attaching Log File ($emaillog)" -ShowTime -enablelogs:$enablelogs
    $null = Copy-Item $logfile -Destination $emaillog -Force
    Write-Output "[$(Get-Date -Format $logdateformat)] Sending Email...."  | Out-File -FilePath $emaillog -Encoding unicode -Append
    Write-Output "###################### Logging Finished - [$(Get-Date -Format $logdateformat)] ######################`n" | Out-File -FilePath $emaillog -Encoding unicode -Append
    Start-Sleep 1    
    $email.attachments.add($emaillog)  
  }
  
  # Send the email.
  $SMTPClient=New-Object System.Net.Mail.SmtpClient( $SmtpServer , $SmtpPort )
  $SMTPClient.EnableSsl=$true
  $SMTPClient.Credentials=New-Object System.Net.NetworkCredential( $SmtpUser , $SmtpPassword );
  Write-EZLogs "Sending email via $SmtpServer\:$Smtpport" -showtime
  try
  {
    $SMTPClient.Send( $email )
    $emailstatus = '[SUCCESS] Email successfuly sent!' 
    $emailcolor = 'green'
  }
  catch
  {
    $emailstatus = "[ERROR] Sending email failed! $_"
    $emailcolor = 'red'
  }
  $email.Dispose();
  Write-EZLogs $emailstatus -showtime -color:$emailcolor -enablelogs:$enablelogs
}
#---------------------------------------------- 
#endregion Send Email Function
#---------------------------------------------- 

#---------------------------------------------- 
#region Dell Command Function
#----------------------------------------------
Function Get-DellCommand
{
  [CmdletBinding()]
  param(
    [switch] $ApplyAllUpdates,
    [switch] $Reboot,
    [string] $ApplyUpdateTypes
  )
  Write-EZLogs '#### Running Dell Command Update ####' -color yellow -linesbefore 1 -enablelogs:$enablelogs
  $Computer_Info = Get-WmiObject Win32_ComputerSystem
  if ($($computer_info.Manufacturer) -notmatch 'Dell')
  {
    $null = Write-EZLogs "The current hardware manufacturer ($($computer_info.Manufacturer)) is not detected as a Dell. Skipping Dell Command..." -ShowTime -color Red -enablelogs:$enablelogs -Warning
    return $null
  }
  #Check if already installed
  $Check_Install = Test-Path -literalpath "$env:ProgramW6432\Dell\CommandUpdate\dcu-cli.exe"
  if($Check_Install)
  {Write-EZLogs 'Dell Command Update is installed...skipping download' -ShowTime -enablelogs:$enablelogs}
  else
  {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
    Write-EZLogs -text 'Downloading Dell Command Update' -showtime -enablelogs:$enablelogs
    # Download Dell Command Update
    $start_time2 = Get-Date
    $processtimeout = 300 #Amount of time in seconds to wait until process is timedout and canceled
    $App_name = 'DellCommand'
    $dell_log = "$DownloadLocation_dell\DellCommand.log"
    $execute_file_name = 'DellCommandUpdate.exe'
    $arguments = "/scan -report=$DownloadLocation_dell -outputlog=$dell_log"
    $download_output_file = [System.IO.Path]::Combine($DownloadLocation_dell, $execute_file_name)  
    #Download BrowsingHistoryView
    try 
    {
      $TestDownloadLocation = Test-Path $DownloadLocation_dell
      if (!$TestDownloadLocation) 
      {
        $null = New-Item $DownloadLocation_dell -ItemType Directory -force
        Write-EZLogs "Creating destination directory and downloading file $downloadurl_dell" -ShowTime -enablelogs:$enablelogs 
      }
      else
      {Write-EZLogs -text "Destination directory exists...downloading file $downloadurl_dell" -showtime -enablelogs:$enablelogs}    
      Invoke-WebRequest -Uri $DownloadURL_dell -OutFile $download_output_file -UseBasicParsing 
      Write-EZLogs -text "Download Time taken for file $DownloadURL_dell : $((Get-Date).Subtract($start_time2).Seconds) second(s)" -ShowTime -enablelogs:$enablelogs 
      Write-EZLogs -text "$App_name downloaded to $download_output_file" -ShowTime -enablelogs:$enablelogs 
      Write-EZLogs "Installing $execute_file_name...." -ShowTime -enablelogs:$enablelogs 
      $null = Start-Process -FilePath $download_output_file -ArgumentList '/s' -Verbose -Wait
    }
    catch 
    {  
      Write-EZLogs -text "[ERROR] The download and extraction of $execute_file_name failed: $($_.Exception.Message)" -ShowTime -red -enablelogs:$enablelogs
      exit 1
    }  
  }
  $DCU_Args = $null
  $block = $Null
  $DCU_Args = "/scan -report=$DownloadLocation_dell -outputlog=$DownloadLocation_dell\DellCommand.log"
  $block = 
  {
    Param
    (
      [string]$DownloadLocation_dell,
      [string]$dell_log,
      [switch]$enablelogs,
      [string]$DCU_Args
      
    )
    #Run Dell Command Update process
    $proc = Start-Process "$env:ProgramW6432\Dell\CommandUpdate\dcu-cli.exe" -ArgumentList $using:DCU_Args -Wait -WindowStyle Hidden
    #$proc = Start-Process "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe" -Wait -Argumentlist "/scan -report=$DownloadLocation_dell -outputlog=$DownloadLocation_dell\DellCommand.log" -WindowStyle Hidden #-PassThru
    $proc | Wait-Process -Timeout 400 -ErrorAction Continue -ErrorVariable timeouted
    if ($timeouted)
    {
      # terminate the process
      $proc | Stop-Process
      Write-EZLogs -text 'Process failed to finish before the timeout period and was canceled. Removing downloaded file and exiting' -color red -enablelogs:$using:enablelogs
      Write-EZLogs -text "Timedout: $timeouted" -color red -enablelogs:$using:enablelogs
      $Null = Remove-Item $DownloadLocation_dell -Recurse
      exit
    }
    elseif ($proc.ExitCode -ne 0)
    {Write-EZLogs -text "Process Exit Code: $($proc.ExitCode)" -showtime -Warning -LogFile $using:logfile}     
  }   
   
  #Remove all jobs and set max threads
  Get-Job | Remove-Job -Force
  $MaxThreads = 3
  
  #Start the jobs. Max 4 jobs running simultaneously.
  While ($(Get-Job -state running).count -ge $MaxThreads)
  {Start-Sleep -Milliseconds 3}
  Write-EZLogs -text ">>>> Running Dell Command Update`n" -showtime -color cyan -enablelogs:$enablelogs
  $dellupdates_code = $Null
  $Null = Start-Job -Scriptblock $Block -ArgumentList $DownloadLocation_dell,$dell_log,$enablelogs,$DCU_Args
  Write-EZLogs '-----------DellCommand Log Entries-----------' -enablelogs:$enablelogs           
  #Wait for all jobs to finish.
  While ($(Get-Job -State Running).count -gt 0)
  {
    #Check last line of the log, if it matches our exit trigger text, sleep until it changes indicating new log entries are being added
    if (!(Test-Path "$DownloadLocation_dell\DellCommand.log"))
    {Start-Sleep -Milliseconds 3}
    else
    {
      $last_line = Get-Content -Path "$DownloadLocation_dell\DellCommand.log" -force -Tail 1 2> $Null
      if($last_line -match 'Program exited with return code:')
      {Start-Sleep -Milliseconds 3}
      #Watch the log file and output all new lines. If the new line matches our exit trigger text, break out of wait
      $count = 0
      Get-Content -Path "$DownloadLocation_dell\DellCommand.log" -force -Tail 1 -wait  | ForEach-Object {
        $count++
        Write-EZLogs "$_" -enablelogs:$enablelogs
        if($_ -match 'The program exited with return code: 500 '){ $dellexit_code = 500 ;break}
        if($_ -match 'Number of applicable updates for the current system configuration:'){ $dellupdates_code = $_.Substring(($_.IndexOf('configuration: ')+15))}
        if($_ -match 'The program exited with return code: 0 '){ $dellexit_code = 0 ;break}  
        if($_ -match 'Program exited with return code:' -and $_ -notmatch 'Exiting with exit code: InvalidParameters'){break}
        if($(Get-Job -State Running).count -eq 0){$delljob_code = 0;break }
      }
    }      
  }
  
  #Get information from each job.
  foreach($job in Get-Job)
  {$info=Receive-Job -Id ($job.Id)}
  
  #Remove all jobs created.
  Get-Job | Remove-Job -Force 
  Write-EZLogs '---------------END Log Entries---------------' -enablelogs:$enablelogs
  Write-EZLogs -text ">>>> Dell Command Finished. Final loop count: $count" -showtime -enablelogs:$enablelogs -color Cyan
  $dellupdates_code = $dellupdates_code.trim()
  if($dellexit_code -eq 500)
  {
    Write-EZLogs '[INFO] Dell Command found no updates that are available for this system' -showtime -enablelogs:$enablelogs -color Cyan
    return $false
  }
  elseif($dellupdates_code)
  {Write-EZLogs "[INFO] Dell Command found $dellupdates_code updates that are available for this system" -showtime -enablelogs:$enablelogs -color Cyan}
  try
  {[xml]$XMLReport = Get-Content "$DownloadLocation_dell\DCUApplicableUpdates.xml" -ErrorAction stop}
  catch
  {
    Write-EZLogs "[ERROR] Unable to process DCUApplicableUpdates.xml - $_" -ShowTime -color red -enablelogs:$enablelogs
    return $false
  }
  
  $AvailableUpdates = $XMLReport.updates.update
  $BIOSUpdates = ($XMLReport.updates.update | Where-Object { $_.type -eq 'BIOS' }).name.Count
  $ApplicationUpdates = ($XMLReport.updates.update | Where-Object { $_.type -eq 'Application' }).name.Count
  $DriverUpdates = ($XMLReport.updates.update | Where-Object { $_.type -eq 'Driver' }).name.Count
  $FirmwareUpdates = ($XMLReport.updates.update | Where-Object { $_.type -eq 'Firmware' }).name.Count
  $OtherUpdates = ($XMLReport.updates.update | Where-Object { $_.type -eq 'Other' }).name.Count
  $PatchUpdates = ($XMLReport.updates.update | Where-Object { $_.type -eq 'Patch' }).name.Count
  $UtilityUpdates = ($XMLReport.updates.update | Where-Object { $_.type -eq 'Utility' }).name.Count
  $UrgentUpdates = ($XMLReport.updates.update | Where-Object { $_.Urgency -eq 'Urgent' }).name.Count
  
  if($BIOSUpdates){Write-EZLogs "BIOS Updates: $BIOSUpdates" -showtime -enablelogs:$enablelogs}
  if($ApplicationUpdates){Write-EZLogs "Application Updates: $ApplicationUpdates" -showtime -enablelogs:$enablelogs}
  if($DriverUpdates){Write-EZLogs "Driver Updates: $DriverUpdates" -showtime -enablelogs:$enablelogs}
  if($FirmwareUpdates){Write-EZLogs "Firmware Updates: $FirmwareUpdates" -showtime -enablelogs:$enablelogs}
  if($OtherUpdates){Write-EZLogs "Other Updates: $OtherUpdates" -showtime -enablelogs:$enablelogs}
  if($PatchUpdates){Write-EZLogs "Patch Updates: $PatchUpdates" -showtime -enablelogs:$enablelogs}
  if($UtilityUpdates){Write-EZLogs "Utility Updates: $UtilityUpdates" -showtime -enablelogs:$enablelogs}
  if($UrgentUpdates){Write-EZLogs "Urgent Updates: $UrgentUpdates" -showtime -enablelogs:$enablelogs}
  
  foreach ($update in $AvailableUpdates)
  {
    $filename = Split-Path $($update.file) -Leaf
    if($update.urgency -match 'Recommended')
    {$urgency_msg = 'Dell recommends applying this update during your next scheduled update cycle. The update contains feature enhancements or changes that will help keep your system software current and compatible with other system modules (firmware, BIOS, drivers and software).'}
    elseif($update.urgency -match 'Urgent')
    {$urgency_msg = 'Dell highly recommends applying this update as soon as possible. The update contains changes to improve the reliability and availability of your Dell system.'}
    elseif($update.urgency -match 'Optional')
    {$urgency_msg = 'Dell recommends the customer review specifics about the update to determine if it applies to your system. The update contains changes that impact only certain configurations, or provides new features that may/may not apply to your environment.'}
    else
    {$urgency_msg = $Null}
 
    if($urgency_msg)
    {   
      $urgency_link = @"
<style>
.popup {
  position: relative;
  display: inline-block;
  cursor: pointer;
  -webkit-user-select: none;
  -moz-user-select: none;
  -ms-user-select: none;
  user-select: none;
}
/* The actual popup */
.popup .popuptext {
  font-style:bold;
  font-weight:400;
  text-transform:none;
  letter-spacing:normal;
  word-break:normal;
  word-spacing:normal;
  white-space:normal;
  line-break:auto;
  word-wrap:break-word;
  text-decoration:none;
  line-height:1.5;
  text-align: left;
  text-align:start;
  visibility: hidden;
  width: 200px;
  background-color:#f4f4f4;
  z-index:1060;
  max-width:280px;
  text-align:left;
  text-align:start;
  background-color:#f4f4f4;
  background-clip:padding-box;
  border:1px solid #ddd;
  border-radius:.125rem;
 -webkit-box-shadow:2px 2px 8px rgba(0,0,0,.1);
 box-shadow:2px 2px 8px rgba(0,0,0,.1);
  color: black;
  padding: 10px 12px;
  position: absolute;
  top: 125%;
  left: 0%;
  margin-left: -80px;
}
/* Popup arrow */
.popup .popuptext::before {
  content: "";
  position: absolute;
 border-color:transparent;
  bottom: 95%;
  left: 50%;
  margin-left: -5px;
  border-width: 5px;
  border-style: solid;
  border-color: #555 transparent transparent transparent;
}
/* Toggle this class - hide and show the popup */
.popup .show {
  visibility: visible;
  -webkit-animation: fadeIn 1s;
  animation: fadeIn 1s;
}
.popup #b1:hover + .popuptext {
  visibility: visible;
  opacity: 1;
}
/* Add animation (fade in the popup) */
@-webkit-keyframes fadeIn {
  from {
    opacity: 0;
  }
  to {
    opacity: 1;
  }
}
@keyframes fadeIn {
  from {
    opacity: 0;
  }
  to {
    opacity: 1;
  }
}
</style>
  <span class=popup>
    <a href=# id=b1>$($update.urgency)</a>
    <span class=popuptext id=myPopup>$urgency_msg</span>
  </span>
"@.replace("`n",'')
    }
    else
    {$urgency_link = $update.urgency}   
    $dellcommandOutput  = New-Object -Type PSObject
    $dellcommandOutput | Add-Member -MemberType NoteProperty -Name 'DisplayName' -Value $update.name
    $dellcommandOutput | Add-Member -MemberType NoteProperty -Name 'Name' -Value "<a href='https://www.dell.com/support/home/en-us/drivers/driversdetails?driverid=$($update.release)' target='_blank'>$($update.name)</a>"
    $dellcommandOutput | Add-Member -MemberType NoteProperty -Name 'Release' -Value $update.release
    $dellcommandOutput | Add-Member -MemberType NoteProperty -Name 'Version' -Value $update.version
    $dellcommandOutput | Add-Member -MemberType NoteProperty -Name 'Release Date' -Value $update.date
    $dellcommandOutput | Add-Member -MemberType NoteProperty -Name 'Urgency' -Value $urgency_link
    $dellcommandOutput | Add-Member -MemberType NoteProperty -Name 'Type' -Value $update.type
    $dellcommandOutput | Add-Member -MemberType NoteProperty -Name 'Category' -Value $update.category
    $dellcommandOutput | Add-Member -MemberType NoteProperty -Name 'File Name' -Value $update.file
    $dellcommandOutput | Add-Member -MemberType NoteProperty -Name 'File Size' -Value ('{0:N2}MB' -f ($update.bytes/1MB)) 
    $dellcommandOutput | Add-Member -MemberType NoteProperty -Name 'Download' -Value "<a href='https://dl.dell.com/$($update.file)' target='_blank'>$filename</a>"
    $dellcommandOutput
  }
  $ApplyUpdate_Array = $ApplyUpdateTypes.split(',')
  $updates_types_toinstall = Compare-Object $AvailableUpdates.type -DifferenceObject $ApplyUpdate_Array -IncludeEqual -ExcludeDifferent -passthru
  if($AvailableUpdates.type)
  {Write-EZLogs "Availabe Update Types Found: $($AvailableUpdates.type | Select-Object -Unique)" -ShowTime -enablelogs:$enablelogs}
  if($ApplyUpdate_Array)
  {Write-EZLogs "Selected Update Types to install: $($ApplyUpdate_Array)" -ShowTime -enablelogs:$enablelogs}
  if($ApplyAllUpdates)
  {
    $InstallUpdates = $true
    Write-EZLogs 'Selected Update Types to install: All' -ShowTime -enablelogs:$enablelogs
    Write-EZLogs "Update Types to be installed: $($AvailableUpdates.type | Select-Object -Unique)" -ShowTime -enablelogs:$enablelogs
  }
  elseif($updates_types_toinstall)
  {
    $InstallUpdates = $true
    Write-EZLogs "Update Types to be installed: $updates_types_toinstall" -ShowTime -enablelogs:$enablelogs
  }
  else
  {$InstallUpdates = $false}
  if($AvailableUpdates -and $InstallUpdates)
  {
    Write-EZLogs -text "`n#### Intalling Available Updates ####" -color yellow -enablelogs:$enablelogs -LogTime:$false
    if ($reboot)
    {
      $rebootarg = 'enable'
      $reboot_msg1 = '>>>> A reboot will occur automatically after installation of updates'
    }
    else
    {
      $rebootarg='disable'
      $reboot_msg1 = '>>>> No reboot will occur automatically'
    }
    if($ApplyAllUpdates)
    {$updates_types_toinstall_arg = $null}
    else
    {$updates_types_toinstall_arg = "-updateType=$ApplyUpdateTypes"}
    try
    {
      $Global:DCU_Args = "/applyUpdates -autoSuspendBitLocker=enable $updates_types_toinstall_arg -reboot=$rebootarg -outputlog=$DownloadLocation_dell\DellCommand.log"
      Get-Job | Remove-Job -Force
      $MaxThreads = 3
      
      While ($(Get-Job -state running).count -ge $MaxThreads)
      {Start-Sleep -Milliseconds 3}
      Write-EZLogs -text ">>>> Running Dell Command Update`n" -showtime -color cyan -enablelogs:$enablelogs
      $Null = Start-Job -Scriptblock $Block -ArgumentList $DownloadLocation_dell,$dell_log,$enablelogs,$DCU_Args
      Write-EZLogs '-----------DellCommand Log Entries-----------' -enablelogs:$enablelogs           
      #Wait for all jobs to finish.
      While ($(Get-Job -State Running).count -gt 0)
      {
        #Check last line of the log, if it matches our exit trigger text, sleep until it changes indicating new log entries are being added
        if (!(Test-Path "$DownloadLocation_dell\DellCommand.log"))
        {Start-Sleep -Milliseconds 3}
        else
        {
          $last_line = Get-Content -Path "$DownloadLocation_dell\DellCommand.log" -force -Tail 1 2> $Null
          if($last_line -match 'Program exited with return code:')
          {Start-Sleep -Milliseconds 2}
          #Watch the log file and output all new lines. If the new line matches our exit trigger text, break out of wait
          $count2 = 0
          Get-Content -Path "$DownloadLocation_dell\DellCommand.log" -force -Tail 1 -wait  | ForEach-Object {
            $count2++
            Write-EZLogs "$_" -enablelogs:$enablelogs
            if($_ -match 'The program exited with return code: 500 '){ $dellexit_code = 500 ;break}
            if($_ -match 'Finished installing the updates'){ $dellupdates_code = 11}
            if($_ -match 'Pending self-update installation for these updates, will get installed after a system reboot'){ $dellreboot_code = 1;$reboot_msg = '>>>> Self-update installations will complete after a system reboot'}
            if($_ -match 'The system has been updated and requires a reboot to complete the process.'){ $dellreboot_code = 2;$reboot_msg = '>>>> A reboot is required to complete the update process'}
            if($_ -match 'Warning:'){$dell_warnings += "$_"}
            if($_ -match 'The program exited with return code: 0 '){ $dellexit_code = 0 ;break}  
            if($_ -match 'Program exited with return code:' -and $_ -notmatch 'Exiting with exit code: InvalidParameters'){break} 
            if($(Get-Job -State Running).count -eq 0){$delljob_code = 0;break }
          }
        }      
      }
  
      #Get information from each job.
      foreach($job in Get-Job)
      {$info=Receive-Job -Id ($job.Id)}
  
      #Remove all jobs created.
      Get-Job | Remove-Job -Force 
      Write-EZLogs '---------------END Log Entries---------------' -enablelogs:$enablelogs
      Write-EZLogs -text ">>>> Dell Command Finished. Final loop count: $count2" -showtime -enablelogs:$enablelogs -color Cyan      
      if($dellexit_code -eq 11)
      {Write-EZLogs '[INFO] Dell Command finished installing updates' -showtime -enablelogs:$enablelogs -color Cyan}      
      #start-process "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe" -ArgumentList "/applyUpdates -autoSuspendBitLocker=enable $updates_types_toinstall_arg -reboot=$rebootarg -outputlog=$DownloadLocation_dell\DellCommand.log" -Wait -WindowStyle Hidden
      Write-EZLogs -text $reboot_msg -ShowTime -Color Cyan -enablelogs:$enablelogs
      Write-EZLogs -text $reboot_msg1 -ShowTime -Color Cyan -enablelogs:$enablelogs 
    }
    catch
    {Write-EZLogs "[ERROR] An exception occurred while trying to apply updates: $($_.exception.message)`n$($_.InvocationInfo.positionmessage)`n$($_.ScriptStackTrace)" -showtime -color red -enablelogs:$enablelogs}
  }
  else
  {Write-EZLogs -text '>>>> No available updates match the types specified to install or none were configured, we are done here' -ShowTime -Color Cyan -enablelogs:$enablelogs}
}
#---------------------------------------------- 
#endregion Dell Command Function
#----------------------------------------------

#############################################################################
#endregion Core functions
#############################################################################

#############################################################################
#region Execution and Output - Code that executes required actions and/or performs output 
#############################################################################
 
#---------------------------------------------- 
#region Run Dell Command Update
#----------------------------------------------
if ($runDellCommand)
{
  #Ensure the script runs with elevated priviliges
  if($InstallDellCommand_Drivers_reboot -eq 1){$reboot_arg = $true}else{$reboot_arg = $false}
  Use-RunAs
  if($installDellCommand_Drivers_Types -and $installDellCommand_Drivers_Types -notmatch 'All')
  {$dellcommand_results = Get-DellCommand -ApplyUpdateTypes $installDellCommand_Drivers_Types -reboot:$reboot_arg}  
  elseif ($installDellCommand_Drivers_Types -match 'All')
  {$dellcommand_results = Get-DellCommand -ApplyAllUpdates -reboot:$reboot_arg}
  else
  {$dellcommand_results = Get-DellCommand}
}
#---------------------------------------------- 
#endregion Run Dell Command Update
#----------------------------------------------

#---------------------------------------------- 
#region Build HTML Dashboard Report
#----------------------------------------------
if($dashboardreport -eq 1){$create_report = $true}else{$create_report = $false}
if($opendashboardreport -eq 1){$open_report = $true}else{$open_report = $false}
if ($create_report -and $dellcommand_results)
{
  Write-EZLogs -Text "`n#### Generating HTML Report ####" -Color yellow -enablelogs:$enablelogs
  $dellcommand_results_dash_all = $dellcommand_results | Select-Object * | Sort-Object 'Date' -Descending
  $dellcommand_results_dash_drivers = $dellcommand_results | Select-Object * | Where-Object {$_.type -eq 'Driver'} | Sort-Object 'Date' -Descending
  $dellcommand_results_dash_application = $dellcommand_results | Select-Object * | Where-Object {$_.type -eq 'Application'} | Sort-Object 'Date' -Descending
  $dellcommand_results_dash_firmware = $dellcommand_results | Select-Object * | Where-Object {$_.type -eq 'Firmware'} | Sort-Object 'Date' -Descending
  $dellcommand_results_dash_bios = $dellcommand_results | Select-Object * | Where-Object {$_.type -eq 'BIOS'} | Sort-Object 'Date' -Descending
  
  $dashboardreport_name = "$env:computername-DellCommandUpdates-$(Get-Date -Format yyyy-mm-dd-hhmm)"
  if($save_report_location)
  {
    if(!(Test-Path $save_report_location))
    {
      Write-EZLogs "Save directory doesnt exist...attempting to create at $save_report_location" -ShowTime -enablelogs:$enablelogs
      try
      {$null = New-Item $save_report_location -ItemType Directory -force}
      catch
      {
        Write-EZLogs "[ERROR] Exception creating directory $save_report_location - $_" -ShowTime -color Red -enablelogs:$enablelogs
        Write-EZLogs "Setting save report location to: $DownloadLocation_Dell" -ShowTime -color Red -enablelogs:$enablelogs
        $save_report_location = $DownloadLocation_Dell
      }
    }
    $dashboardreport_file = [System.IO.Path]::Combine($save_report_location, "$dashboardreport_name.html")
  }
  else
  {$dashboardreport_file = [System.IO.Path]::Combine($DownloadLocation_Dell, "$dashboardreport_name.html")}
  New-HTML -Name $dashboardreport_name -FilePath $dashboardreport_file -ShowHTML:$open_report  {
    New-HTMLMain -BackgroundColor dimgray
    New-HTMLTabStyle -SlimTabs -LinearGradient  -BackgroundColor lightblue -FontWeightActive bold -BorderStyle outset -BorderBottomStyleActive groove -BorderBottomColorActive lightgrey
    New-HTMLTab -Name "$env:computername - Dell Command Report" -IconSolid desktop -TextSize 14 -IconColor white -IconSize 15 -HtmlData {
      New-HTMLTab -Name 'All Available Updates' -IconSolid hdd -HtmlData  {  
        New-HTMLSection -BackgroundColor LightGrey -HeaderBackGroundColor navy -Content {
          New-HTMLPanel -AlignContentText center -BackgroundColor white -Content  {
            New-HTMLTableOption -DataStore HTML 
            New-HTMLTable -DataTable $dellcommand_results_dash_all -DefaultSortIndex 0 -ExcludeProperty 'DisplayName' -filtering -AutoSize -DefaultSortOrder Descending  -FreezeColumnsLeft 1 -PagingOptions 100 -ScreenSizePercent 65 -DateTimeSortingFormat 'M-dd-yyy HH:mm:ss tt' -InvokeHTMLTags  {
              New-TableCondition -Name 'Urgency' -ComparisonType string -Operator contains -Value 'Urgent' -FontWeight bold -FontSize 12 -BackgroundColor red -color white
            }
          }
        }
      }
      if($dellcommand_results_dash_drivers)
      {
        New-HTMLTab -Name 'Driver Updates' -IconSolid hdd -HtmlData  {  
          New-HTMLSection -BackgroundColor LightGrey -HeaderBackGroundColor navy -Content {
            New-HTMLPanel -AlignContentText center -BackgroundColor white -Content  {
              New-HTMLTable -DataTable $dellcommand_results_dash_drivers -DefaultSortIndex 0 -ExcludeProperty 'DisplayName' -filtering -AutoSize -DefaultSortOrder Descending  -FreezeColumnsLeft 1 -PagingOptions 100 -ScreenSizePercent 65 -DateTimeSortingFormat 'M-dd-yyy HH:mm:ss tt' -InvokeHTMLTags {
                New-TableCondition -Name 'Urgency' -ComparisonType string -Operator contains -Value 'Urgent' -FontWeight bold -FontSize 12 -BackgroundColor red -color white
              }
            }
          }
        }
      }
      if($dellcommand_results_dash_application)
      {
        New-HTMLTab -Name 'Application Updates' -IconSolid hdd -HtmlData  {  
          New-HTMLSection -BackgroundColor LightGrey -HeaderBackGroundColor navy -Content {
            New-HTMLPanel -AlignContentText center -BackgroundColor white -Content  {
              New-HTMLTable -DataTable $dellcommand_results_dash_application -DefaultSortIndex 0 -ExcludeProperty 'DisplayName' -filtering -AutoSize -DefaultSortOrder Descending  -FreezeColumnsLeft 1 -PagingOptions 100 -ScreenSizePercent 65 -DateTimeSortingFormat 'M-dd-yyy HH:mm:ss tt'   -InvokeHTMLTags {
                New-TableCondition -Name 'Urgency' -ComparisonType string -Operator contains -Value 'Urgent' -FontWeight bold -FontSize 12 -BackgroundColor red -color white
              }
            }
          }
        }
      }
      if($dellcommand_results_dash_firmware)
      {
        New-HTMLTab -Name 'Firmware Updates' -IconSolid hdd -HtmlData  {  
          New-HTMLSection -BackgroundColor LightGrey -HeaderBackGroundColor navy -Content {
            New-HTMLPanel -AlignContentText center -BackgroundColor white -Content  {
              New-HTMLTable -DataTable $dellcommand_results_dash_firmware -DefaultSortIndex 0 -ExcludeProperty 'DisplayName' -filtering -AutoSize -DefaultSortOrder Descending  -FreezeColumnsLeft 1 -PagingOptions 100 -ScreenSizePercent 65 -DateTimeSortingFormat 'M-dd-yyy HH:mm:ss tt' -InvokeHTMLTags  {
                New-TableCondition -Name 'Urgency' -ComparisonType string -Operator contains -Value 'Urgent' -FontWeight bold -FontSize 12 -BackgroundColor red -color white
              }
            }
          }
        }
      }
      if($dellcommand_results_dash_bios)
      {
        New-HTMLTab -Name 'BIOS Updates' -IconSolid hdd -HtmlData  {  
          New-HTMLSection -BackgroundColor LightGrey -HeaderBackGroundColor navy -Content {
            New-HTMLPanel -AlignContentText center -BackgroundColor white -Content  {
              New-HTMLTable -DataTable $dellcommand_results_dash_bios -DefaultSortIndex 0 -ExcludeProperty 'DisplayName' -filtering -AutoSize -DefaultSortOrder Descending  -FreezeColumnsLeft 1 -PagingOptions 100 -ScreenSizePercent 65 -DateTimeSortingFormat 'M-dd-yyy HH:mm:ss tt' -InvokeHTMLTags {
                New-TableCondition -Name 'Urgency' -ComparisonType string -Operator contains -Value 'Urgent' -FontWeight bold -FontSize 12 -BackgroundColor red -color white
              }
            }
          }
        } 
      }
    }
  }
  if(Test-Path $dashboardreport_file)
  {Write-EZLogs "[SUCCESS] HTML Report file successfully generated and saved to: $dashboardreport_file" -showtime -color green -enablelogs:$enablelogs}
  else
  {Write-EZLogs -Text 'No report file was generated or there was no valid data returned from Dell Command to create a report from' -Color red -enablelogs:$enablelogs -Warning}  
}

#---------------------------------------------- 
#endregion Build HTML Dashboard Report
#----------------------------------------------

#---------------------------------------------- 
#region Send Email and Cleanup
#----------------------------------------------
if($email_report -eq 1){$send_email = $true}else{$send_email = $false}
if($email_logs -eq 1){$send_email_logs = $true}else{$send_email_logs = $false}
if($send_email -and $dashboardreport_file)
{
  Write-EZLogs -Text "`n#### Generating and Sending Email ####" -Color yellow -enablelogs:$enablelogs
  try 
  {Send-Email -dashboardreport_file $dashboardreport_file -SmtpUser $SmtpUser -SmtpPort $SmtpPort -SmtpServer $SmtpServer -MailFrom $MailFrom -MailTo $MailTo -Subject $Subject -SmtpPassword $SmtpPassword -enablelogs:$send_email_logs -logfile:$logfile}
  catch
  {Write-EZLogs "[ERROR] $_" -Color Red -enablelogs:$enablelogs}
}
if(!$save_report_location -and $dashboardreport_file)
{
  try 
  {
    Write-EZLogs -Text "Removing Dashboard Report File $dashboardreport_file" -ShowTime -enablelogs:$enablelogs
    Remove-Item $dashboardreport_file -Recurse
  }
  catch
  {Write-EZLogs "[ERROR] $_" -Color Red -enablelogs:$enablelogs}  
}
Write-EZLogs "`n#### Cleaning Up Temp Files and Folders ####" -color yellow -enablelogs:$enablelogs
if($dellcommand_results)
{
  try
  {
    Write-EZLogs "Removing file: $DownloadLocation_dell\DCUApplicableUpdates.xml" -showtime -enablelogs:$enablelogs
    $null = Remove-Item "$DownloadLocation_dell\DCUApplicableUpdates.xml" -Force
    Write-EZLogs "Removing file: $DownloadLocation_dell\DellCommand.log" -showtime -enablelogs:$enablelogs
    $null = Remove-Item "$DownloadLocation_dell\DellCommand.log" -Force
  }
  catch
  {Write-EZLogs "[ERROR] Unable to remove file(s) -- $_" -ShowTime -color red -enablelogs:$enablelogs}
}
try
{
  Write-EZLogs "Removing $Script_Temp_Folder" -ShowTime -enablelogs:$enablelogs
  Remove-Item $Script_Temp_Folder -Recurse -Force
  Write-EZLogs '[SUCCESS] Execution and cleanup complete' -showtime -color green -enablelogs:$enablelogs
}
catch
{Write-EZLogs "[ERROR] - $_" -ShowTime -color Red -enablelogs:$enablelogs}
#---------------------------------------------- 
#endregion Send Email and Cleanup
#----------------------------------------------

#---------------------------------------------- 
#region Finish Logging
#----------------------------------------------
if ($enablelogs)
{
  if($error)
  {
    Write-Output "`n`n[-----ALL ERRORS------]" | Out-File -FilePath $logfile -Encoding unicode -Append
    $e_index = 0
    foreach ($e in $error)
    {
      $e_index++
      Write-Output "[ERROR $e_index Message] =========================================================================`n$($e.exception.message)`n$($e.InvocationInfo.positionmessage)`n$($e.ScriptStackTrace)`n`n" | Out-File -FilePath $logfile -Encoding unicode -Append
    }
    Write-Output '-----------------' | Out-File -FilePath $logfile -Encoding unicode -Append
    $error.Clear()
  }
  Write-EZLogs "`n======== Total Script Execution Time ========" -enablelogs:$enablelogs -LogTime:$false
  Write-EZLogs "Minutes      : $($stopwatch.elapsed.Minutes)`nSeconds      : $($stopwatch.elapsed.Seconds)`nMilliseconds : $($stopwatch.elapsed.Milliseconds)" -enablelogs:$enablelogs -LogTime:$false
  $($stopwatch.stop())
  $($stopwatch.reset()) 
  Write-Output "###################### Logging Finished - [$(Get-Date -Format $logdateformat)] ######################`n" | Out-File -FilePath $logfile -Encoding unicode -Append
}  
#---------------------------------------------- 
#endregion Finish Logging
#----------------------------------------------
#############################################################################
#endregion Execution and Output Functions
#############################################################################