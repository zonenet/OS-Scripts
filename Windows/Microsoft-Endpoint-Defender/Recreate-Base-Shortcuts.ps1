#Requires -RunAsAdministrator
# Recreate Base Shortcuts - https://github.com/TheAlienDrew/OS-Scripts/blob/master/Windows/Microsoft-Endpoint-Defender/Recreate-Base-Shortcuts.ps1
# Script only recreates shortcuts to applications it knows are installed, and also works for user profile installed applications.
# If a program you use isn't in any of the lists here, either fork/edit/push, or create an issue at:
# https://github.com/TheAlienDrew/OS-Scripts/issues/new?title=%5BAdd%20App%5D%20Recreate-Base-Shortcuts.ps1&body=%3C%21--%20Please%20enter%20the%20app%20you%20need%20added%20below%2C%20and%20a%20link%20to%20the%20installer%20--%3E%0A%0A

# About the issue: https://www.bleepingcomputer.com/news/microsoft/buggy-microsoft-defender-asr-rule-deletes-windows-app-shortcuts/

# Application objects are setup like so:
<# @{
       Name="[name of shortcut here]";
       TargetPath="[path to exe/url/folder here]";
       Arguments="[any arguments that an app starts with here]";
       SystemLnk="[path to lnk or name of app here]";
       StartIn="[start in path, if needed, here]";
       Description="[comment, that shows up in tooltip, here]";
       IconLocation="[path to ico|exe|ico w/ index]";
       RunAsAdmin="[true or false, if needed]"
   } #>


if (!(Test-Path C:\Skyttel -PathType Container)) {mkdir C:\Skyttel}
Start-Transcript -Path "C:\Skyttel\Recreate-Base-Shortcuts.log"
Write-Host "" # Makes log look better

# Constants

Set-Variable NotInstalled -Option Constant -Value "NOT-INSTALLED"



# Variables

$isWindows11 = ((Get-WMIObject win32_operatingsystem).Caption).StartsWith("Microsoft Windows 11")

$isWin10orNewer = [System.Environment]::OSVersion.Version.Major -ge 10

# Functions

function New-Shortcut {
  param(
    [Parameter(Mandatory = $true)]
    [Alias("name", "n")]
    [string]$sName,

    [Parameter(Mandatory = $true)]
    [Alias("targetpath", "tp")]
    [string]$sTargetPath,

    [Alias("arguments", "a")]
    [string]$sArguments, # Optional (for special shortcuts)

    [Alias("systemlnk", "sl")]
    [string]$sSystemLnk, # Optional (for if name / path is different from normal)

    [Alias("startin", "si")]
    [string]$sStartIn, # Optional (for special shortcuts)

    [Alias("description", "d")]
    [string]$sDescription, # Optional (some shortcuts have comments for tooltips)

    [Alias("iconlocation", "il")]
    [string]$sIconLocation, # Optional (some shortcuts have a custom icon)
    
    [Alias("runasadmin", "r")]
    [bool]$sRunAsAdmin, # Optional (if the shortcut should be ran as admin)

    [Alias("user", "u")]
    [string]$sUser # Optional (username of the user to install shortcut to)
  )

  $result = $true
  $resultMsg = @()
  $warnMsg = @()
  $errorMsg = @()

  Set-Variable ProgramShortcutsPath -Option Constant -Value "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
  Set-Variable UserProgramShortcutsPath -Option Constant -Value "C:\Users\${sUser}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"

  # validate name and target path
  if ($sName -And $sTargetPath -And (Test-Path $sTargetPath -PathType leaf)) {
    # if shortcut path not given, create one at default location with $sName
    if (-Not ($sSystemLnk)) { $sSystemLnk = $sName }
    # if doesn't have $ProgramShortcutsPath or $UserProgramShortcutsPath (and not start with drive letter), it'll assume a path for it
    if (-Not ($sSystemLnk -match '^[a-zA-Z]:\\.*' -Or $sSystemLnk -match ('^' + [Regex]::Escape($ProgramShortcutsPath) + '.*') -Or $sSystemLnk -match ('^' + [Regex]::Escape($UserProgramShortcutsPath) + '.*'))) {
      $sSystemLnk = $(if ($sUser) { $UserProgramShortcutsPath } else { $ProgramShortcutsPath }) + '\' + $sSystemLnk
    }
    # if it ends with '\', then we append the name to the end
    if ($sSystemLnk.EndsWith('\')) { $sSystemLnk = $sSystemLnk + $sName }
    # if doesn't end with .lnk, add it
    if (-Not ($sSystemLnk -match '.*\.lnk$')) { $sSystemLnk = $sSystemLnk + '.lnk' }

    # only create shortcut if it doesn't already exist
    if (Test-Path $sSystemLnk -PathType leaf) {
      $resultMsg += "A shortcut already exists at:`n${sSystemLnk}"
      $result = $false
    }
    else {
      $WScriptObj = New-Object -ComObject WScript.Shell
      $newLNK = $WscriptObj.CreateShortcut($sSystemLnk)

      $newLNK.TargetPath = $sTargetPath
      if ($sArguments) { $newLNK.Arguments = $sArguments }
      if ($sStartIn) { $newLNK.WorkingDirectory = $sStartIn }
      if ($sDescription) { $newLNK.Description = $sDescription }
      if ($sIconLocation) { $newLNK.IconLocation = $sIconLocation }

      $newLNK.Save()
      $result = $?
      [Runtime.InteropServices.Marshal]::ReleaseComObject($Shell) | Out-Null

      if ($result) {
        $resultMsg += "Created shortcut at:`n${sSystemLnk}"

        # set to run as admin if needed
        if ($sRunAsAdmin) {
          $bytes = [System.IO.File]::ReadAllBytes($sSystemLnk)
          $bytes[0x15] = $bytes[0x15] -bor 0x20 #set byte 21 (0x15) bit 6 (0x20) ON
          [System.IO.File]::WriteAllBytes($sSystemLnk, $bytes)
          $result = $?
          if ($result) { $resultMsg += "Shortcut set to Run as Admin, at: ${sSystemLnk}" }
          else { $errorMsg += "Failed to set shortcut to Run as Admin, at: ${sSystemLnk}" }
        }
      }
      else { $errorMsg += "Failed to create shortcut, with target at: ${sTargetPath}" }
    }
  }
  elseif (-Not ($sName -Or $sTargetPath)) {
    if (-Not $sName) {
      $errorMsg += "Error! Name is missing!"
    }
    if (-Not $sTargetPath) {
      $errorMsg += "Error! Target is missing!"
    }

    $result = $false
  }
  else {
    $warnMsg += "Target invalid! Doesn't exist or is spelled wrong:`n${sTargetPath}"

    $result = $false
  }

  if ($result) { Write-Host -ForegroundColor Green $sName }
  else { Write-Host -ForegroundColor Red $sName }

  if ($resultMsg.length -gt 0) {
    for ($msgNum = 0; $msgNum -lt $resultMsg.length; $msgNum++) {
      Write-Host $resultMsg[$msgNum]
    }
  }
  elseif ($errortMsg.length -gt 0) {
    for ($msgNum = 0; $msgNum -lt $errorMsg.length; $msgNum++) {
      Write-Error $errorMsg[$msgNum]
    }
  }
  if ($warnMsg.length -gt 0) {
    for ($msgNum = 0; $msgNum -lt $warnMsg.length; $msgNum++) {
      Write-Warning $warnMsg[$msgNum]
    }
  }
  Write-Host ""

  return $result
}



# MAIN

$ScriptResults = $true

if (-Not $isWin10orNewer) {
  Write-Error "This script is only meant to be ran on Windows 10 and newer!"
  exit 1
}



# System Applications

# App arguments dependant on uninstall strings

## App Name
#$App_Arguments = ...

# App paths dependant on app version

# Powershell (7 or newer)
$PowerShell_TargetPath = "C:\Program Files\PowerShell\"
$PowerShell_Version = if (Test-Path -Path $PowerShell_TargetPath) { Get-ChildItem -Directory -Path $PowerShell_TargetPath | Where-Object { $_.Name -match '^[0-9]+$' } | Sort-Object -Descending }
$PowerShell_Version = if ($PowerShell_Version.length -ge 1) { $PowerShell_Version[0].name } else { $NotInstalled }
$PowerShell_TargetPath += "${PowerShell_Version}\pwsh.exe"
$PowerShell_32bit_TargetPath = "C:\Program Files (x86)\PowerShell\"
$PowerShell_32bit_Version = if (Test-Path -Path $PowerShell_32bit_TargetPath) { Get-ChildItem -Directory -Path $PowerShell_32bit_TargetPath | Where-Object { $_.Name -match '^[0-9]+$' } | Sort-Object -Descending }
$PowerShell_32bit_Version = if ($PowerShell_32bit_Version.length -ge 1) { $PowerShell_32bit_Version[0].name } else { $NotInstalled }
$PowerShell_32bit_TargetPath += "${PowerShell32bit_Version}\pwsh.exe"
# PowerToys
$PowerToys_TargetPath = "C:\Program Files\PowerToys\PowerToys.exe"

# App names dependant on OS or app version

# PowerShell (7 or newer)
$PowerShell_Name = "PowerShell " + $(if ($PowerShell_Version) { $PowerShell_Version } else { $NotInstalled }) + " (x64)"
$PowerShell_32bit_Name = "PowerShell " + $(if ($PowerShell_32bit_Version) { $PowerShell_32bit_Version } else { $NotInstalled }) + " (x86)"
# PowerToys
$PowerToys_isPreview = if (Test-Path -Path $PowerToys_TargetPath -PathType Leaf) { (Get-Item $PowerToys_TargetPath).VersionInfo.FileVersionRaw.Major -eq 0 }
$PowerToys_Name = "PowerToys" + $(if ($PowerToys_isPreview) { " (Preview)" })
# Windows Accessories
$WindowsMediaPlayerOld_Name = "Windows Media Player" + $(if ($isWindows11) { " Legacy" })

$sysAppList = @(
  # Azure
  @{Name = "Azure Data Studio"; TargetPath = "C:\Program Files\Azure Data Studio\azuredatastudio.exe"; SystemLnk = "Azure Data Studio\"; StartIn = "C:\Program Files\Azure Data Studio" },
  # Edge
  @{Name = "Microsoft Edge"; TargetPath = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"; StartIn = "C:\Program Files\Microsoft\Edge\Application"; Description = "Browse the web" }, # it's the only install on 32-bit
  @{Name = "Microsoft Edge"; TargetPath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"; StartIn = "C:\Program Files (x86)\Microsoft\Edge\Application"; Description = "Browse the web" }, # it's the only install on 64-bit
  # Intune Management Extension
  @{Name = "Microsoft Intune Management Extension"; TargetPath = "C:\Program Files\Microsoft Intune Management Extension\AgentExecutor.exe"; SystemLnk = "Microsoft Intune Management Extension\"; Description = "Microsoft Intune Management Extension" }, # it's the only install on 32-bit
  @{Name = "Microsoft Intune Management Extension"; TargetPath = "C:\Program Files (x86)\Microsoft Intune Management Extension\AgentExecutor.exe"; SystemLnk = "Microsoft Intune Management Extension\"; Description = "Microsoft Intune Management Extension" }, # it's the only install on 64-bit
  # Office
  @{Name = "Access"; TargetPath = "C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE"; Description = "Build a professional app quickly to manage data." },
  @{Name = "Excel"; TargetPath = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"; Description = "Easily discover, visualize, and share insights from your data." },
  @{Name = "OneNote"; TargetPath = "C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE"; Description = "Take notes and have them when you need them." },
  @{Name = "Outlook"; TargetPath = "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"; Description = "Manage your email, schedules, contacts, and to-dos." },
  @{Name = "PowerPoint"; TargetPath = "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"; Description = "Design and deliver beautiful presentations with ease and confidence." },
  @{Name = "Project"; TargetPath = "C:\Program Files\Microsoft Office\root\Office16\WINPROJ.EXE"; Description = "Easily collaborate with others to quickly start and deliver winning projects." },
  @{Name = "Publisher"; TargetPath = "C:\Program Files\Microsoft Office\root\Office16\MSPUB.EXE"; Description = "Create professional-grade publications that make an impact." },
  @{Name = "Visio"; TargetPath = "C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE"; Description = "Create professional and versatile diagrams that simplify complex information." },
  @{Name = "Word"; TargetPath = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"; Description = "Create beautiful documents, easily work with others, and enjoy the read." },
  @{Name = "Database Compare"; TargetPath = "C:\Program Files\Microsoft Office\root\Client\AppVLP.exe"; Arguments = "`"C:\Program Files (x86)\Microsoft Office\Office16\DCF\DATABASECOMPARE.EXE`""; SystemLnk = "Microsoft Office Tools\"; Description = "Compare versions of an Access database." },
  @{Name = "Office Language Preferences"; TargetPath = "C:\Program Files\Microsoft Office\root\Office16\SETLANG.EXE"; SystemLnk = "Microsoft Office Tools\"; Description = "Change the language preferences for Office applications." },
  @{Name = "Spreadsheet Compare"; TargetPath = "C:\Program Files\Microsoft Office\root\Client\AppVLP.exe"; Arguments = "`"C:\Program Files (x86)\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE`""; SystemLnk = "Microsoft Office Tools\"; Description = "Compare versions of an Excel workbook." },
  @{Name = "Telemetry Log for Office"; TargetPath = "C:\Program Files\Microsoft Office\root\Office16\msoev.exe"; SystemLnk = "Microsoft Office Tools\"; Description = "View critical errors, compatibility issues and workaround information for your Office solutions by using Office Telemetry Log." },
  @{Name = "Access (32-bit)"; TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE"; Description = "Build a professional app quickly to manage data." },
  @{Name = "Excel (32-bit)"; TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"; Description = "Easily discover, visualize, and share insights from your data." },
  @{Name = "OneNote (32-bit)"; TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\ONENOTE.EXE"; Description = "Take notes and have them when you need them." },
  @{Name = "Outlook (32-bit)"; TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"; Description = "Manage your email, schedules, contacts, and to-dos." },
  @{Name = "PowerPoint (32-bit)"; TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE"; Description = "Design and deliver beautiful presentations with ease and confidence." },
  @{Name = "Project (32-bit)"; TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\WINPROJ.EXE"; Description = "Easily collaborate with others to quickly start and deliver winning projects." },
  @{Name = "Publisher (32-bit)"; TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\MSPUB.EXE"; Description = "Create professional-grade publications that make an impact." },
  @{Name = "Visio (32-bit)"; TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\VISIO.EXE"; Description = "Create professional and versatile diagrams that simplify complex information." },
  @{Name = "Word (32-bit)"; TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"; Description = "Create beautiful documents, easily work with others, and enjoy the read." },
  @{Name = "Database Compare (32-bit)"; TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Client\AppVLP.exe"; Arguments = "`"C:\Program Files (x86)\Microsoft Office\Office16\DCF\DATABASECOMPARE.EXE`""; SystemLnk = "Microsoft Office Tools\"; Description = "Compare versions of an Access database." },
  @{Name = "Office Language Preferences (32-bit)"; TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\SETLANG.EXE"; SystemLnk = "Microsoft Office Tools\"; Description = "Change the language preferences for Office applications." },
  @{Name = "Spreadsheet Compare (32-bit)"; TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Client\AppVLP.exe"; Arguments = "`"C:\Program Files (x86)\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE`""; SystemLnk = "Microsoft Office Tools\"; Description = "Compare versions of an Excel workbook." },
  @{Name = "Telemetry Log for Office (32-bit)"; TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\msoev.exe"; SystemLnk = "Microsoft Office Tools\"; Description = "View critical errors, compatibility issues and workaround information for your Office solutions by using Office Telemetry Log." },
  # OneDrive
  @{Name = "OneDrive"; TargetPath = "C:\Program Files\Microsoft OneDrive\OneDrive.exe"; Description = "Keep your most important files with you wherever you go, on any device." },
  @{Name = "OneDrive (32-bit)"; TargetPath = "C:\Program Files (x86)\Microsoft OneDrive\OneDrive.exe"; Description = "Keep your most important files with you wherever you go, on any device." },
  # PowerShell (7 or newer)
  @{Name = $PowerShell_Name; TargetPath = $PowerShell_TargetPath; Arguments = "-WorkingDirectory ~"; SystemLnk = "PowerShell\"; Description = $PowerShell_Name },
  @{Name = $PowerShell_32bit_Name; TargetPath = $PowerShell_32bit_TargetPath; Arguments = "-WorkingDirectory ~"; SystemLnk = "PowerShell\"; Description = $PowerShell_32bit_Name },
  # PowerToys
  @{Name = $PowerToys_Name; TargetPath = $PowerToys_TargetPath; SystemLnk = $PowerToys_Name + '\'; StartIn = "C:\Program Files\PowerToys\"; Description = "PowerToys - Windows system utilities to maximize productivity" },
  # Visual Studio
  @{Name = "Visual Studio 2022"; TargetPath = "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\devenv.exe"; StartIn = "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\"; Description = "Microsoft Visual Studio 2022" },
  @{Name = "Visual Studio 2022"; TargetPath = "C:\Program Files\Microsoft Visual Studio\2022\Professional\Common7\IDE\devenv.exe"; StartIn = "C:\Program Files\Microsoft Visual Studio\2022\Professional\Common7\IDE\"; Description = "Microsoft Visual Studio 2022" },
  @{Name = "Visual Studio 2022"; TargetPath = "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\Common7\IDE\devenv.exe"; StartIn = "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\Common7\IDE\"; Description = "Microsoft Visual Studio 2022" },
  @{Name = "Blend for Visual Studio 2022"; TargetPath = "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\Blend.exe"; StartIn = "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\"; Description = "Microsoft Blend for Visual Studio 2022" },
  @{Name = "Blend for Visual Studio 2022"; TargetPath = "C:\Program Files\Microsoft Visual Studio\2022\Professional\Common7\IDE\Blend.exe"; StartIn = "C:\Program Files\Microsoft Visual Studio\2022\Professional\Common7\IDE\"; Description = "Microsoft Blend for Visual Studio 2022" },
  @{Name = "Blend for Visual Studio 2022"; TargetPath = "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\Common7\IDE\Blend.exe"; StartIn = "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\Common7\IDE\"; Description = "Microsoft Blend for Visual Studio 2022" },
  @{Name = "Visual Studio 2019"; TargetPath = "C:\Program Files\Microsoft Visual Studio\2019\Community\Common7\IDE\devenv.exe"; StartIn = "C:\Program Files\Microsoft Visual Studio\2019\Community\Common7\IDE\"; Description = "Microsoft Visual Studio 2019" },
  @{Name = "Visual Studio 2019"; TargetPath = "C:\Program Files\Microsoft Visual Studio\2019\Professional\Common7\IDE\devenv.exe"; StartIn = "C:\Program Files\Microsoft Visual Studio\2019\Professional\Common7\IDE\"; Description = "Microsoft Visual Studio 2019" },
  @{Name = "Visual Studio 2019"; TargetPath = "C:\Program Files\Microsoft Visual Studio\2019\Enterprise\Common7\IDE\devenv.exe"; StartIn = "C:\Program Files\Microsoft Visual Studio\2019\Enterprise\Common7\IDE\"; Description = "Microsoft Visual Studio 2019" },
  @{Name = "Blend for Visual Studio 2019"; TargetPath = "C:\Program Files\Microsoft Visual Studio\2019\Community\Common7\IDE\Blend.exe"; StartIn = "C:\Program Files\Microsoft Visual Studio\2019\Community\Common7\IDE\"; Description = "Microsoft Blend for Visual Studio 2019" },
  @{Name = "Blend for Visual Studio 2019"; TargetPath = "C:\Program Files\Microsoft Visual Studio\2019\Professional\Common7\IDE\Blend.exe"; StartIn = "C:\Program Files\Microsoft Visual Studio\2019\Professional\Common7\IDE\"; Description = "Microsoft Blend for Visual Studio 2019" },
  @{Name = "Blend for Visual Studio 2019"; TargetPath = "C:\Program Files\Microsoft Visual Studio\2019\Enterprise\Common7\IDE\Blend.exe"; StartIn = "C:\Program Files\Microsoft Visual Studio\2019\Enterprise\Common7\IDE\"; Description = "Microsoft Blend for Visual Studio 2019" },
  @{Name = "Visual Studio 2017"; TargetPath = "C:\Program Files\Microsoft Visual Studio\2017\Community\Common7\IDE\devenv.exe"; StartIn = "C:\Program Files\Microsoft Visual Studio\2017\Community\Common7\IDE\"; Description = "Microsoft Visual Studio 2017" },
  @{Name = "Visual Studio 2017"; TargetPath = "C:\Program Files\Microsoft Visual Studio\2017\Professional\Common7\IDE\devenv.exe"; StartIn = "C:\Program Files\Microsoft Visual Studio\2017\Professional\Common7\IDE\"; Description = "Microsoft Visual Studio 2017" },
  @{Name = "Visual Studio 2017"; TargetPath = "C:\Program Files\Microsoft Visual Studio\2017\Enterprise\Common7\IDE\devenv.exe"; StartIn = "C:\Program Files\Microsoft Visual Studio\2017\Enterprise\Common7\IDE\"; Description = "Microsoft Visual Studio 2017" },
  @{Name = "Blend for Visual Studio 2017"; TargetPath = "C:\Program Files\Microsoft Visual Studio\2017\Community\Common7\IDE\Blend.exe"; StartIn = "C:\Program Files\Microsoft Visual Studio\2017\Community\Common7\IDE\"; Description = "Microsoft Blend for Visual Studio 2017" },
  @{Name = "Blend for Visual Studio 2017"; TargetPath = "C:\Program Files\Microsoft Visual Studio\2017\Professional\Common7\IDE\Blend.exe"; StartIn = "C:\Program Files\Microsoft Visual Studio\2017\Professional\Common7\IDE\"; Description = "Microsoft Blend for Visual Studio 2017" },
  @{Name = "Blend for Visual Studio 2017"; TargetPath = "C:\Program Files\Microsoft Visual Studio\2017\Enterprise\Common7\IDE\Blend.exe"; StartIn = "C:\Program Files\Microsoft Visual Studio\2017\Enterprise\Common7\IDE\"; Description = "Microsoft Blend for Visual Studio 2017" },
  @{Name = "Visual Studio Code"; TargetPath = "C:\Program Files\Microsoft VS Code\Code.exe"; SystemLnk = "Visual Studio Code\"; StartIn = "C:\Program Files\Microsoft VS Code" },
  @{Name = "Visual Studio Code (32-bit)"; TargetPath = "C:\Program Files (x86)\Microsoft VS Code\Code.exe"; SystemLnk = "Visual Studio Code\"; StartIn = "C:\Program Files\Microsoft VS Code" },
  @{Name = "Visual Studio Installer"; TargetPath = "C:\Program Files\Microsoft Visual Studio\Installer\setup.exe"; StartIn = "C:\Program Files\Microsoft Visual Studio\Installer" }, # it's the only install on 32-bit
  @{Name = "Visual Studio Installer"; TargetPath = "C:\Program Files (x86)\Microsoft Visual Studio\Installer\setup.exe"; StartIn = "C:\Program Files (x86)\Microsoft Visual Studio\Installer" }, # it's the only install on 64-bit
  # SQL Management Studio
  @{Name = "Microsoft SQL Server Management Studio 18"; TargetPath = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE\Ssms.exe"; SystemLnk = "Microsoft SQL Server Tools 18\"; StartIn = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE" },
  @{Name = "Analysis Services Deployment Wizard 18"; TargetPath = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE\Microsoft.AnalysisServices.Deployment.exe"; SystemLnk = "Microsoft SQL Server Tools 18\"; StartIn = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE" },
  @{Name = "Database Engine Tuning Advisor 18"; TargetPath = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\DTASHELL.EXE"; SystemLnk = "Microsoft SQL Server Tools 18\Performance Tools\"; StartIn = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE" },
  @{Name = "SQL Server Profiler 18"; TargetPath = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\PROFILER.EXE"; SystemLnk = "Microsoft SQL Server Tools 18\Performance Tools\"; StartIn = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE" },

  @{Name = "Microsoft SQL Server Management Studio 17"; TargetPath = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 17\Common7\IDE\Ssms.exe"; SystemLnk = "Microsoft SQL Server Tools 17\"; StartIn = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 17\Common7\IDE" },
  @{Name = "Analysis Services Deployment Wizard 17"; TargetPath = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 17\Common7\IDE\Microsoft.AnalysisServices.Deployment.exe"; SystemLnk = "Microsoft SQL Server Tools 17\"; StartIn = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 17\Common7\IDE" },
  @{Name = "Database Engine Tuning Advisor 17"; TargetPath = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 17\Common7\DTASHELL.EXE"; SystemLnk = "Microsoft SQL Server Tools 17\Performance Tools\"; StartIn = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 17\Common7\IDE" },
  @{Name = "SQL Server Profiler 17"; TargetPath = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 17\Common7\PROFILER.EXE"; SystemLnk = "Microsoft SQL Server Tools 17\Performance Tools\"; StartIn = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 17\Common7\IDE" },
  # Custom Navision 2017 + regular client
  @{Name = "NAVCS"; TargetPath = "C:\Program Files (x86)\Skyttel\NAV Servervelger\NavChangeServer.exe"; StartIn = "C:\Program Files (x86)\Skyttel\NAV Servervelger" },
  @{Name = "Dynamics Nav 2017"; TargetPath = "C:\Program Files (x86)\Microsoft Dynamics NAV\100\RoleTailored Client\Microsoft.Dynamics.Nav.Client.exe"; StartIn = "C:\Program Files (x86)\Microsoft Dynamics NAV\100\RoleTailored Client" },
  # Windows Accessories
  @{Name = "Remote Desktop Connection"; TargetPath = "%windir%\system32\mstsc.exe"; SystemLnk = "Accessories\"; StartIn = "%windir%\system32\"; Description = "Use your computer to connect to a computer that is located elsewhere and run programs or access files." },
  @{Name = "Steps Recorder"; TargetPath = "%windir%\system32\psr.exe"; SystemLnk = "Accessories\"; Description = "Capture steps with screenshots to save or share." },
  @{Name = "Windows Fax and Scan"; TargetPath = "%windir%\system32\WFS.exe"; SystemLnk = "Accessories\"; Description = "Send and receive faxes or scan pictures and documents." },
  @{Name = $WindowsMediaPlayerOld_Name; TargetPath = "%ProgramFiles%\Windows Media Player\wmplayer.exe"; Arguments = "/prefetch:1"; SystemLnk = "Accessories\"; StartIn = "%ProgramFiles(x86)%\Windows Media Player" }, # it's the only install on 32-bit
  @{Name = $WindowsMediaPlayerOld_Name; TargetPath = "%ProgramFiles(x86)%\Windows Media Player\wmplayer.exe"; Arguments = "/prefetch:1"; SystemLnk = "Accessories\"; StartIn = "%ProgramFiles(x86)%\Windows Media Player" }, # it's the only install on 64-bit
  @{Name = "WordPad"; TargetPath = "%ProgramFiles%\Windows NT\Accessories\wordpad.exe"; SystemLnk = "Accessories\"; Description = "Creates and edits text documents with complex formatting." },
  @{Name = "Character Map"; TargetPath = "%windir%\system32\charmap.exe"; SystemLnk = "Accessories\System Tools\"; Description = "Selects special characters and copies them to your document." }
  #@{ Name=""; TargetPath=""; Arguments=""; SystemLnk=""; StartIn=""; Description=""; IconLocation=""; RunAsAdmin=($true -Or $false) },
)

for ($i = 0; $i -lt $sysAppList.length; $i++) {
  $app = $sysAppList[$i]
  $aName = $app.Name
  $aTargetPath = $app.TargetPath
  $aArguments = if ($app.Arguments) { $app.Arguments } else { "" }
  $aSystemLnk = if ($app.SystemLnk) { $app.SystemLnk } else { "" }
  $aStartIn = if ($app.StartIn) { $app.StartIn } else { "" }
  $aDescription = if ($app.Description) { $app.Description } else { "" }
  $aIconLocation = if ($app.IconLocation) { $app.IconLocation } else { "" }
  $aRunAsAdmin = if ($app.RunAsAdmin) { $app.RunAsAdmin } else { $false }

  $ScriptResults = New-Shortcut -n $aName -tp $aTargetPath -a $aArguments -sl $aSystemLnk -si $aStartIn -d $aDescription -il $aIconLocation -r $aRunAsAdmin
}







# Third-Party System Applications (not made by Microsoft)

# App arguments dependant on uninstall strings


# App paths dependant on app version

# GIMP
$GIMP_TargetPath = "C:\Program Files\"
$GIMP_FindFolder = Get-ChildItem -Directory -Path $GIMP_TargetPath | Where-Object { $_.Name -match '^GIMP' } | Sort-Object -Descending
$GIMP_FindFolder = if ($GIMP_FindFolder.length -ge 1) { $GIMP_FindFolder[0].name } else { $NotInstalled }
$GIMP_TargetPath += "${GIMP_FindFolder}\bin\"
$GIMP_FindExe = if (Test-Path -Path $GIMP_TargetPath) { Get-ChildItem -File -Path $GIMP_TargetPath | Where-Object { $_.Name -match '^gimp\-[.0-9]+exe$' } | Sort-Object -Descending }
$GIMP_FindExe = if ($GIMP_FindExe.length -ge 1) { $GIMP_FindExe[0].name } else { "${NotInstalled}.exe" }
$GIMP_TargetPath += $GIMP_FindExe
$GIMP_32bit_TargetPath = "C:\Program Files (x86)\"
$GIMP_32bit_FindFolder = Get-ChildItem -Directory -Path $GIMP_32bit_TargetPath | Where-Object { $_.Name -match '^GIMP' } | Sort-Object -Descending
$GIMP_32bit_FindFolder = if ($GIMP_32bit_FindFolder.length -ge 1) { $GIMP_32bit_FindFolder[0].name } else { $NotInstalled }
$GIMP_32bit_TargetPath += "${GIMP_32bit_FindFolder}\bin\"
$GIMP_32bit_FindExe = if (Test-Path -Path $GIMP_32bit_TargetPath) { Get-ChildItem -File -Path $GIMP_32bit_TargetPath | Where-Object { $_.Name -match '^gimp\-[.0-9]+exe$' } | Sort-Object -Descending }
$GIMP_32bit_FindExe = if ($GIMP_32bit_FindExe.length -ge 1) { $GIMP_32bit_FindExe[0].name } else { "${NotInstalled}.exe" }
$GIMP_32bit_TargetPath += $GIMP_32bit_FindExe
# Google
$GoogleDrive_TargetPath = "C:\Program Files\Google\Drive File Stream\"
$GoogleDrive_Version = if (Test-Path -Path $GoogleDrive_TargetPath) { Get-ChildItem -Directory -Path $GoogleDrive_TargetPath | Where-Object { $_.Name -match '^[.0-9]+$' } | Sort-Object -Descending }
$GoogleDrive_Version = if ($GoogleDrive_Version.length -ge 1) { $GoogleDrive_Version[0].name } else { $NotInstalled }
$GoogleDrive_TargetPath += "${GoogleDrive_Version}\GoogleDriveFS.exe"
$GoogleOneVPN_TargetPath = "C:\Program Files\Google\VPN by Google One\"
$GoogleOneVPN_Version = if (Test-Path -Path $GoogleOneVPN_TargetPath) { Get-ChildItem -Directory -Path $GoogleOneVPN_TargetPath | Where-Object { $_.Name -match '^[.0-9]+$' } | Sort-Object -Descending }
$GoogleOneVPN_Version = if ($GoogleOneVPN_Version.length -ge 1) { $GoogleOneVPN_Version[0].name } else { $NotInstalled }
$GoogleOneVPN_TargetPath += "${GoogleOneVPN_Version}\googleone.exe"
# KeePass
$KeePass_StartIn = "C:\Program Files\"
$KeePass_FindFolder = Get-ChildItem -Directory -Path $KeePass_StartIn | Where-Object { $_.Name -match '^KeePass Password Safe' } | Sort-Object -Descending
$KeePass_FindFolder = if ($KeePass_FindFolder.length -ge 1) { $KeePass_FindFolder[0].name } else { $NotInstalled }
$KeePass_TargetPath = "${KeePass_FindFolder}\KeePass.exe"
$KeePass_32bit_StartIn = "C:\Program Files (x86)\"
$KeePass_32bit_FindFolder = Get-ChildItem -Directory -Path $KeePass_32bit_StartIn | Where-Object { $_.Name -match '^KeePass Password Safe' } | Sort-Object -Descending
$KeePass_32bit_FindFolder = if ($KeePass_32bit_FindFolder.length -ge 1) { $KeePass_32bit_FindFolder[0].name } else { $NotInstalled }
$KeePass_32bit_TargetPath = "${KeePass_32bit_FindFolder}\KeePass.exe"

# App names dependant on OS or app version

# GIMP
$GIMP_ProductVersion = if (Test-Path -Path $GIMP_TargetPath -PathType Leaf) { (Get-Item $GIMP_TargetPath).VersionInfo.ProductVersion }
$GIMP_Version = if ($GIMP_ProductVersion) { $GIMP_ProductVersion } else { $NotInstalled }
$GIMP_Name = "GIMP ${GIMP_Version}"
$GIMP_32bit_ProductVersion = if (Test-Path -Path $GIMP_32bit_TargetPath -PathType Leaf) { (Get-Item $GIMP_32bit_TargetPath).VersionInfo.ProductVersion }
$GIMP_32bit_Version = if ($GIMP_32bit_ProductVersion) { $GIMP_32bit_ProductVersion } else { $NotInstalled }
$GIMP_32bit_Name = "GIMP ${GIMP_32bit_Version}"
# KeePass
$KeePass_FileVersionRaw = if (Test-Path -Path $KeePass_TargetPath -PathType Leaf) { (Get-Item $KeePass_TargetPath).VersionInfo.FileVersionRaw }
$KeePass_Version = if ($KeePass_FileVersionRaw) { $KeePass_FileVersionRaw.Major } else { $NotInstalled }
$KeePass_Name = "KeePass ${KeePass_Version}"
$KeePass_32bit_FileVersionRaw = if (Test-Path -Path $KeePass_32bit_TargetPath -PathType Leaf) { (Get-Item $KeePass_32bit_TargetPath).VersionInfo.FileVersionRaw }
$KeePass_32bit_Version = if ($KeePass_32bit_FileVersionRaw) { $KeePass_32bit_FileVersionRaw.Major } else { $NotInstalled }
$KeePass_32bit_Name = "KeePass ${KeePass_32bit_Version}"


$sys3rdPartyAppList = @(
  # 7-Zip
  @{Name = "7-Zip File Manager"; TargetPath = "C:\Program Files\7-Zip\7zFM.exe"; SystemLnk = "7-Zip\" },
  @{Name = "7-Zip Help"; TargetPath = "C:\Program Files\7-Zip\7-zip.chm"; SystemLnk = "7-Zip\" },
  @{Name = "7-Zip File Manager (32-bit)"; TargetPath = "C:\Program Files (x86)\7-Zip\7zFM.exe"; SystemLnk = "7-Zip\" },
  @{Name = "7-Zip Help"; TargetPath = "C:\Program Files (x86)\7-Zip\7-zip.chm"; SystemLnk = "7-Zip\" },
  # Adobe
  @{Name = "Adobe Acrobat"; TargetPath = "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe" },
  @{Name = "Adobe Acrobat Distiller"; TargetPath = "C:\Program Files\Adobe\Acrobat DC\Acrobat\acrodist.exe" },
  @{Name = "Adobe Creative Cloud"; TargetPath = "C:\Program Files\Adobe\Adobe Creative Cloud\ACC\Creative Cloud.exe" },
  @{Name = "Adobe UXP Developer Tool"; TargetPath = "C:\Program Files\Adobe\Adobe UXP Developer Tool\Adobe UXP Developer Tool.exe"; StartIn = "C:\Program Files\Adobe\Adobe UXP Developer Tool" },
  @{Name = "Adobe Acrobat (32-bit)"; TargetPath = "C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe" },
  @{Name = "Adobe Acrobat Distiller (32-bit)"; TargetPath = "C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\acrodist.exe" },
  @{Name = "Adobe Acrobat Reader"; TargetPath = "C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe" }, # old version; it's the only install on 32-bit
  @{Name = "Adobe Acrobat Distiller"; TargetPath = "C:\Program Files\Adobe\Acrobat Reader DC\Reader\acrodist.exe" }, # old version; it's the only install on 32-bit
  @{Name = "Adobe Acrobat Reader"; TargetPath = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe" }, # old version; it's the only install on 64-bit
  @{Name = "Adobe Acrobat Distiller"; TargetPath = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\acrodist.exe" }, # old version; it's the only install on 64-bit
  # Cisco
  @{Name = "Cisco AnyConnect Secure Mobility Client"; TargetPath = "C:\Program Files\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe"; SystemLnk = "C:\Program Files\Cisco\Cisco AnyConnect Secure Mobility Client\"; Description = "Cisco AnyConnect Secure Mobility Client" }, # it's the only install on 32-bit
  @{Name = "Cisco AnyConnect Secure Mobility Client"; TargetPath = "C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe"; SystemLnk = "C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\"; Description = "Cisco AnyConnect Secure Mobility Client" }, # it's the only install on 64-bit
  # Citrix Workspace
  @{Name = "Citrix Workspace"; TargetPath = "C:\Program Files\Citrix\ICA Client\SelfServicePlugin\SelfService.exe"; Arguments = "-showAppPicker"; StartIn = "C:\Program Files\Citrix\ICA Client\SelfServicePlugin\"; Description = "Select applications you want to use on your computer" }, # it's the only install on 32-bit
  @{Name = "Citrix Workspace"; TargetPath = "C:\Program Files (x86)\Citrix\ICA Client\SelfServicePlugin\SelfService.exe"; Arguments = "-showAppPicker"; StartIn = "C:\Program Files (x86)\Citrix\ICA Client\SelfServicePlugin\"; Description = "Select applications you want to use on your computer" }, # it's the only install on 64-bit
  # Docker
  @{Name = "Docker Desktop"; TargetPath = "C:\Program Files\Docker\Docker\Docker Desktop.exe"; SystemLnk = "C:\ProgramData\Microsoft\Windows\Start Menu\"; Description = "Docker Desktop" },
  # draw.io
  @{Name = "draw.io"; TargetPath = "C:\Program Files\draw.io\draw.io.exe"; StartIn = "C:\Program Files\draw.io"; Description = "draw.io desktop" },
  @{Name = "draw.io (32-bit)"; TargetPath = "C:\Program Files (x86)\draw.io\draw.io.exe"; StartIn = "C:\Program Files (x86)\draw.io"; Description = "draw.io desktop" },
  # GIMP
  @{Name = $GIMP_Name; TargetPath = $GIMP_TargetPath; StartIn = "%USERPROFILE%"; Description = $GIMP_Name },
  @{Name = $GIMP_32bit_Name; TargetPath = $GIMP_32bit_TargetPath; StartIn = "%USERPROFILE%"; Description = $GIMP_32bit_Name },
  # Google
  @{Name = "Google Chrome"; TargetPath = "C:\Program Files\Google\Chrome\Application\chrome.exe"; StartIn = "C:\Program Files\Google\Chrome\Application"; Description = "Access the Internet" },
  @{Name = "Google Chrome (32-bit)"; TargetPath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"; StartIn = "C:\Program Files (x86)\Google\Chrome\Application"; Description = "Access the Internet" },
  @{Name = "Google Drive"; TargetPath = $GoogleDrive_TargetPath; Description = "Google Drive" },
  @{Name = "VPN by Google One"; TargetPath = $GoogleOneVPN_TargetPath; Description = "VPN by Google One" },
  # KeePass
  @{Name = $KeePass_Name; TargetPath = $KeePass_TargetPath; StartIn = $KeePass_StartIn }, # new version 2+
  @{Name = $KeePass_32bit_Name; TargetPath = $KeePass_32bit_TargetPath; StartIn = $KeePass_32bit_StartIn }, # new version 2+
  @{Name = "KeePass"; TargetPath = "C:\Program Files\KeePass Password Safe\KeePass.exe"; StartIn = "C:\Program Files\KeePass Password Safe" }, # old version 1.x; it's the only install on 32-bit
  @{Name = "KeePass"; TargetPath = "C:\Program Files (x86)\KeePass Password Safe\KeePass.exe"; StartIn = "C:\Program Files (x86)\KeePass Password Safe" }, # old version 1.x; it's the only install on 64-bit
  # Mozilla
  @{Name = "Firefox"; TargetPath = "C:\Program Files\Mozilla Firefox\firefox.exe"; StartIn = "C:\Program Files\Mozilla Firefox" },
  @{Name = "Firefox Private Browsing"; TargetPath = "C:\Program Files\Mozilla Firefox\private_browsing.exe"; StartIn = "C:\Program Files\Mozilla Firefox"; Description = "Firefox Private Browsing" },
  @{Name = "Firefox (32-bit)"; TargetPath = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"; StartIn = "C:\Program Files (x86)\Mozilla Firefox" },
  @{Name = "Firefox Private Browsing (32-bit)"; TargetPath = "C:\Program Files (x86)\Mozilla Firefox\private_browsing.exe"; StartIn = "C:\Program Files (x86)\Mozilla Firefox"; Description = "Firefox Private Browsing" },
  @{Name = "Thunderbird"; TargetPath = "C:\Program Files\Mozilla Thunderbird\thunderbird.exe"; StartIn = "C:\Program Files\Mozilla Thunderbird" },
  @{Name = "Thunderbird (32-bit)"; TargetPath = "C:\Program Files (x86)\Mozilla Thunderbird\thunderbird.exe"; StartIn = "C:\Program Files (x86)\Mozilla Thunderbird" },
  # Notepad++
  @{Name = "Notepad++"; TargetPath = "C:\Program Files\Notepad++\notepad++.exe"; StartIn = "C:\Program Files\Notepad++" },
  @{Name = "Notepad++ (32-bit)"; TargetPath = "C:\Program Files (x86)\Notepad++\notepad++.exe"; StartIn = "C:\Program Files (x86)\Notepad++" },
  # OpenVPN
  @{Name = "OpenVPN"; TargetPath = "C:\Program Files\OpenVPN\bin\openvpn-gui.exe"; SystemLnk = "OpenVPN\OpenVPN GUI"; StartIn = "C:\Program Files\OpenVPN\bin\" },
  @{Name = "OpenVPN Manual Page"; TargetPath = "C:\Program Files\OpenVPN\doc\openvpn.8.html"; SystemLnk = "OpenVPN\Documentation\"; StartIn = "C:\Program Files\OpenVPN\doc\" },
  @{Name = "OpenVPN Windows Notes"; TargetPath = "C:\Program Files\OpenVPN\doc\INSTALL-win32.txt"; SystemLnk = "OpenVPN\Documentation\"; StartIn = "C:\Program Files\OpenVPN\doc\" },
  @{Name = "OpenVPN Configuration File Directory"; TargetPath = "C:\Program Files\OpenVPN\config"; SystemLnk = "OpenVPN\Shortcuts\"; StartIn = "C:\Program Files\OpenVPN\config\" },
  @{Name = "OpenVPN Log File Directory"; TargetPath = "C:\Program Files\OpenVPN\log"; SystemLnk = "OpenVPN\Shortcuts\"; StartIn = "C:\Program Files\OpenVPN\log\" },
  @{Name = "OpenVPN Sample Configuration Files"; TargetPath = "C:\Program Files\OpenVPN\sample-config"; SystemLnk = "OpenVPN\Shortcuts\"; StartIn = "C:\Program Files\OpenVPN\sample-config\" },
  @{Name = "Add a new TAP-Windows6 virtual network adapter"; TargetPath = "C:\Program Files\OpenVPN\bin\tapctl.exe"; Arguments = "create --hwid root\tap0901"; SystemLnk = "OpenVPN\Utilities\"; StartIn = "C:\Program Files\OpenVPN\bin\" },
  @{Name = "Add a new Wintun virtual network adapter"; TargetPath = "C:\Program Files\OpenVPN\bin\tapctl.exe"; Arguments = "create --hwid wintun"; SystemLnk = "OpenVPN\Utilities\"; StartIn = "C:\Program Files\OpenVPN\bin\" },
  @{Name = "OpenVPN (32-bit)"; TargetPath = "C:\Program Files (x86)\OpenVPN\bin\openvpn-gui.exe"; SystemLnk = "OpenVPN\OpenVPN GUI"; StartIn = "C:\Program Files (x86)\OpenVPN\bin\" },
  @{Name = "OpenVPN Manual Page"; TargetPath = "C:\Program Files (x86)\OpenVPN\doc\openvpn.8.html"; SystemLnk = "OpenVPN\Documentation\"; StartIn = "C:\Program Files (x86)\OpenVPN\doc\" },
  @{Name = "OpenVPN Windows Notes"; TargetPath = "C:\Program Files (x86)\OpenVPN\doc\INSTALL-win32.txt"; SystemLnk = "OpenVPN\Documentation\"; StartIn = "C:\Program Files (x86)\OpenVPN\doc\" },
  @{Name = "OpenVPN Configuration File Directory"; TargetPath = "C:\Program Files (x86)\OpenVPN\config"; SystemLnk = "OpenVPN\Shortcuts\"; StartIn = "C:\Program Files (x86)\OpenVPN\config\" },
  @{Name = "OpenVPN Log File Directory"; TargetPath = "C:\Program Files (x86)\OpenVPN\log"; SystemLnk = "OpenVPN\Shortcuts\"; StartIn = "C:\Program Files (x86)\OpenVPN\log\" },
  @{Name = "OpenVPN Sample Configuration Files"; TargetPath = "C:\Program Files (x86)\OpenVPN\sample-config"; SystemLnk = "OpenVPN\Shortcuts\"; StartIn = "C:\Program Files (x86)\OpenVPN\sample-config\" },
  @{Name = "Add a new TAP-Windows6 virtual network adapter (32-bit)"; TargetPath = "C:\Program Files (x86)\OpenVPN\bin\tapctl.exe"; Arguments = "create --hwid root\tap0901"; SystemLnk = "OpenVPN\Utilities\"; StartIn = "C:\Program Files (x86)\OpenVPN\bin\" },
  @{Name = "Add a new Wintun virtual network adapter (32-bit)"; TargetPath = "C:\Program Files (x86)\OpenVPN\bin\tapctl.exe"; Arguments = "create --hwid wintun"; SystemLnk = "OpenVPN\Utilities\"; StartIn = "C:\Program Files (x86)\OpenVPN\bin\" },
  # Oracle
  @{Name = "License (English)"; TargetPath = "C:\Program Files\Oracle\VirtualBox\License_en_US.rtf"; SystemLnk = "Oracle VM VirtualBox\"; StartIn = "C:\Program Files\Oracle\VirtualBox\"; Description = "License" },
  @{Name = "Oracle VM VirtualBox"; TargetPath = "C:\Program Files\Oracle\VirtualBox\VirtualBox.exe"; SystemLnk = "Oracle VM VirtualBox\"; StartIn = "C:\Program Files\Oracle\VirtualBox\"; Description = "Oracle VM VirtualBox" },
  @{Name = "User manual (CHM, English)"; TargetPath = "C:\Program Files\Oracle\VirtualBox\VirtualBox.chm"; SystemLnk = "Oracle VM VirtualBox\"; Description = "User manual" },
  @{Name = "User manual (PDF, English)"; TargetPath = "C:\Program Files\Oracle\VirtualBox\doc\UserManual.pdf"; SystemLnk = "Oracle VM VirtualBox\"; Description = "User manual" },
  @{Name = "License (English)"; TargetPath = "C:\Program Files (x86)\Oracle\VirtualBox\License_en_US.rtf"; SystemLnk = "Oracle VM VirtualBox\"; StartIn = "C:\Program Files (x86)\Oracle\VirtualBox\"; Description = "License" },
  @{Name = "Oracle VM VirtualBox (32-bit)"; TargetPath = "C:\Program Files (x86)\Oracle\VirtualBox\VirtualBox.exe"; SystemLnk = "Oracle VM VirtualBox\"; StartIn = "C:\Program Files (x86)\Oracle\VirtualBox\"; Description = "Oracle VM VirtualBox" },
  @{Name = "User manual (CHM, English)"; TargetPath = "C:\Program Files (x86)\Oracle\VirtualBox\VirtualBox.chm"; SystemLnk = "Oracle VM VirtualBox\"; Description = "User manual" },
  @{Name = "User manual (PDF, English)"; TargetPath = "C:\Program Files (x86)\Oracle\VirtualBox\doc\UserManual.pdf"; SystemLnk = "Oracle VM VirtualBox\"; Description = "User manual" },
  # paint.net
  @{Name = "paint.net"; TargetPath = "C:\Program Files\paint.net\paintdotnet.exe"; StartIn = "C:\Program Files\paint.net"; Description = "Create, edit, scan, and print images and photographs." },
  @{Name = "paint.net (32-bit)"; TargetPath = "C:\Program Files (x86)\paint.net\paintdotnet.exe"; StartIn = "C:\Program Files (x86)\paint.net"; Description = "Create, edit, scan, and print images and photographs." },
  # VideoLAN
  @{Name = "Documentation"; TargetPath = "C:\Program Files\VideoLAN\VLC\Documentation.url"; SystemLnk = "VideoLAN\"; StartIn = "C:\Program Files\VideoLAN\VLC" },
  @{Name = "Release Notes"; TargetPath = "C:\Program Files\VideoLAN\VLC\NEWS.txt"; SystemLnk = "VideoLAN\"; StartIn = "C:\Program Files\VideoLAN\VLC" },
  @{Name = "VideoLAN Website"; TargetPath = "C:\Program Files\VideoLAN\VLC\VideoLAN Website.url"; SystemLnk = "VideoLAN\"; StartIn = "C:\Program Files\VideoLAN\VLC" },
  @{Name = "VLC media player - reset preferences and cache files"; TargetPath = "C:\Program Files\VideoLAN\VLC\vlc.exe"; Arguments = "--reset-config --reset-plugins-cache vlc://quit"; SystemLnk = "VideoLAN\"; StartIn = "C:\Program Files\VideoLAN\VLC" },
  @{Name = "VLC media player skinned"; TargetPath = "C:\Program Files\VideoLAN\VLC\vlc.exe"; Arguments = "-Iskins"; SystemLnk = "VideoLAN\"; StartIn = "C:\Program Files\VideoLAN\VLC" },
  @{Name = "VLC media player"; TargetPath = "C:\Program Files\VideoLAN\VLC\vlc.exe"; SystemLnk = "VideoLAN\"; StartIn = "C:\Program Files\VideoLAN\VLC" },
  @{Name = "Documentation"; TargetPath = "C:\Program Files x86\VideoLAN\VLC\Documentation.url"; SystemLnk = "VideoLAN\"; StartIn = "C:\Program Files x86\VideoLAN\VLC" },
  @{Name = "Release Notes"; TargetPath = "C:\Program Files x86\VideoLAN\VLC\NEWS.txt"; SystemLnk = "VideoLAN\"; StartIn = "C:\Program Files x86\VideoLAN\VLC" },
  @{Name = "VideoLAN Website"; TargetPath = "C:\Program Files x86\VideoLAN\VLC\VideoLAN Website.url"; SystemLnk = "VideoLAN\"; StartIn = "C:\Program Files x86\VideoLAN\VLC" },
  @{Name = "VLC media player - reset preferences and cache files (32-bit)"; TargetPath = "C:\Program Files x86\VideoLAN\VLC\vlc.exe"; Arguments = "--reset-config --reset-plugins-cache vlc://quit"; SystemLnk = "VideoLAN\"; StartIn = "C:\Program Files x86\VideoLAN\VLC" },
  @{Name = "VLC media player skinned (32-bit)"; TargetPath = "C:\Program Files x86\VideoLAN\VLC\vlc.exe"; Arguments = "-Iskins"; SystemLnk = "VideoLAN\"; StartIn = "C:\Program Files x86\VideoLAN\VLC" },
  @{Name = "VLC media player (32-bit)"; TargetPath = "C:\Program Files x86\VideoLAN\VLC\vlc.exe"; SystemLnk = "VideoLAN\"; StartIn = "C:\Program Files x86\VideoLAN\VLC" },
  # WinSCP
  @{Name = "WinSCP"; TargetPath = "C:\Program Files\WinSCP\WinSCP.exe"; StartIn = "C:\Program Files\WinSCP"; Description = "WinSCP: SFTP, FTP, WebDAV and SCP client" }, # it's the only install on 32-bit
  @{Name = "WinSCP"; TargetPath = "C:\Program Files (x86)\WinSCP\WinSCP.exe"; StartIn = "C:\Program Files (x86)\WinSCP"; Description = "WinSCP: SFTP, FTP, WebDAV and SCP client" } # it's the only install on 64-bit
  #@{ Name=""; TargetPath=""; Arguments=""; SystemLnk=""; StartIn=""; Description=""; IconLocation=""; RunAsAdmin=($true -Or $false) },
)

for ($i = 0; $i -lt $sys3rdPartyAppList.length; $i++) {
  $app = $sys3rdPartyAppList[$i]
  $aName = $app.Name
  $aTargetPath = $app.TargetPath
  $aArguments = if ($app.Arguments) { $app.Arguments } else { "" }
  $aSystemLnk = if ($app.SystemLnk) { $app.SystemLnk } else { "" }
  $aStartIn = if ($app.StartIn) { $app.StartIn } else { "" }
  $aDescription = if ($app.Description) { $app.Description } else { "" }
  $aIconLocation = if ($app.IconLocation) { $app.IconLocation } else { "" }
  $aRunAsAdmin = if ($app.RunAsAdmin) { $app.RunAsAdmin } else { $false }

  $ScriptResults = New-Shortcut -n $aName -tp $aTargetPath -a $aArguments -sl $aSystemLnk -si $aStartIn -d $aDescription -il $aIconLocation -r $aRunAsAdmin
}



# User Applications (per user installed apps)

# get all users 
$Users = (Get-ChildItem -Directory -Path "C:\Users\" | ForEach-Object { if (($_.name -ne "Default") -And ($_.name -ne "Public")) { $_.name } })
if ($Users -And ($Users[0].length -eq 1)) { $Users = @("$Users") } # if only one user, array needs to be recreated

# System app arguments dependant on uninstall strings

## App Name
#$App_Arguments = ...

# System app paths dependant on app version

# System app names dependant on OS or app version

# App names dependant on OS or app version

# Microsoft Teams
$MicrosoftTeams_Name = "Microsoft Teams" + $(if ($isWindows11) { " (work or school)" })

for ($i = 0; $i -lt $Users.length; $i++) {
  # get user
  $aUser = $Users[$i]

  # User app paths dependant on app version


  # Discord
  $Discord_StartIn = "C:\Users\${aUser}\AppData\Local\Discord\"
  $Discord_TargetPath = $Discord_StartIn + "Update.exe"
  $Discord_FindFolder = if (Test-Path -Path $Discord_StartIn) { Get-ChildItem -Directory -Path $Discord_StartIn | Where-Object { $_.Name -match '^app\-[.0-9]+$' } | Sort-Object -Descending }
  $Discord_FindFolder = if ($Discord_FindFolder.length -ge 1) { $Discord_FindFolder[0].name } else { $NotInstalled }
  $Discord_StartIn += $Discord_FindFolder
  # GitHub
  $GitHubDesktop_StartIn = "C:\Users\${aUser}\AppData\Local\GitHubDesktop\"
  $GitHubDesktop_TargetPath = $GitHubDesktop_StartIn + "GitHubDesktop.exe"
  $GitHubDesktop_FindFolder = if (Test-Path -Path $GitHubDesktop_StartIn) { Get-ChildItem -Directory -Path $GitHubDesktop_StartIn | Where-Object { $_.Name -match '^app\-[.0-9]+$' } | Sort-Object -Descending }
  $GitHubDesktop_FindFolder = if ($GitHubDesktop_FindFolder.length -ge 1) { $GitHubDesktop_FindFolder[0].name } else { $NotInstalled }
  $GitHubDesktop_StartIn += $GitHubDesktop_FindFolder
  
  # User app names dependant on OS or app version

  $userAppList = @( # all instances of "${aUser}" get's replaced with the username
    # Discord
    @{Name = "Discord"; TargetPath = $Discord_TargetPath; Arguments = "--processStart Discord.exe"; SystemLnk = "Discord Inc\"; StartIn = $Discord_StartIn; Description = "Discord - https://discord.com" },
    # GitHub
    @{Name = "GitHub Desktop"; TargetPath = $GitHubDesktop_TargetPath; SystemLnk = "GitHub, Inc\"; StartIn = $GitHubDesktop_StartIn; Description = "Simple collaboration from your desktop" },
    # Google
    @{Name = "Google Chrome"; TargetPath = "C:\Users\${aUser}\AppData\Local\Google\Chrome\Application\chrome.exe"; StartIn = "C:\Users\${aUser}\AppData\Local\Google\Chrome\Application"; Description = "Access the Internet" },
    # Microsoft
    @{Name = "Azure Data Studio"; TargetPath = "C:\Users\${aUser}\AppData\Local\Programs\Azure Data Studio\azuredatastudio.exe"; SystemLnk = "Azure Data Studio\"; StartIn = "C:\Users\${aUser}\AppData\Local\Programs\Azure Data Studio" },
    @{Name = "Visual Studio Code"; TargetPath = "C:\Users\${aUser}\AppData\Local\Programs\Microsoft VS Code\Code.exe"; SystemLnk = "Visual Studio Code\"; StartIn = "C:\Users\${aUser}\AppData\Local\Programs\Microsoft VS Code" },
    @{Name = "OneDrive"; TargetPath = "C:\Users\${aUser}\AppData\Local\Microsoft\OneDrive\OneDrive.exe"; Description = "Keep your most important files with you wherever you go, on any device." },
    @{Name = $MicrosoftTeams_Name; TargetPath = "C:\Users\${aUser}\AppData\Local\Microsoft\Teams\Update.exe"; Arguments = "--processStart `"Teams.exe`""; StartIn = "C:\Users\${aUser}\AppData\Local\Microsoft\Teams" },
    # Mozilla
    @{Name = "Firefox"; TargetPath = "C:\Users\${aUser}\AppData\Local\Mozilla Firefox\firefox.exe"; StartIn = "C:\Users\${aUser}\AppData\Local\Mozilla Firefox" },
    # NVIDIA Corporation
    @{Name = "NVIDIA GeForce NOW"; TargetPath = "C:\Users\${aUser}\AppData\Local\NVIDIA Corporation\GeForceNOW\CEF\GeForceNOW.exe"; StartIn = "C:\Users\${aUser}\AppData\Local\NVIDIA Corporation\GeForceNOW\CEF" }

    #@{ Name=""; TargetPath=""; Arguments=""; SystemLnk=""; StartIn=""; Description=""; IconLocation=""; RunAsAdmin=($true -Or $false) },
  )

  for ($j = 0; $j -lt $userAppList.length; $j++) {
    $app = $userAppList[$j]
    $aName = $app.Name
    $aTargetPath = $app.TargetPath
    $aArguments = if ($app.Arguments) { $app.Arguments } else { "" }
    $aSystemLnk = if ($app.SystemLnk) { $app.SystemLnk } else { "" }
    $aStartIn = if ($app.StartIn) { $app.StartIn } else { "" }
    $aDescription = if ($app.Description) { $app.Description } else { "" }
    $aIconLocation = if ($app.IconLocation) { $app.IconLocation } else { "" }
    $aRunAsAdmin = if ($app.RunAsAdmin) { $app.RunAsAdmin } else { $false }

    $ScriptResults = New-Shortcut -n $aName -tp $aTargetPath -a $aArguments -sl $aSystemLnk -si $aStartIn -d $aDescription -il $aIconLocation -r $aRunAsAdmin -u $aUser
  }
}

Stop-Transcript

if ($ScriptResults) { Write-Host "Script completed successfully." }
else { Write-Warning "Script completed with warnings and/or errors." }
