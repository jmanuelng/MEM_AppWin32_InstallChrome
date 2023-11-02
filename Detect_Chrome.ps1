<#
.DESCRIPTION
    Detects Google Chrome browser's installation  status on a Windows device and retrieve its version if found. 
    By inspecting familiar file paths and registry entries where Google Chrome is generally installed, 
    the script identifies location path and version of the browser.

    The script is divided into two major sections: 
    - A Functions region which encompasses the primary logic for identifying the installation.
    - A Main region that uses the function, assesses the results, and provides 
      a summarized output of findings.

    If Google Chrome is detected, script will return an "OK" status along with the 
    version number. If it's not found, it will return "FAIL" status.

.HOW IT WORKS
    1. `Get-ChromeExeDetails` function searches for `chrome.exe` executable, it also searches specific registry paths.
       Function returns a custom object with two properties:
       - InstallLocation: Specifies where `chrome.exe` is located.
       - DisplayVersion: Indicates the version of Google Chrome.
       
    3. Main region calls `Get-ChromeExeDetails`. If found outcome will be:
       - "OK" if Google Chrome is found and the version is successfully extracted, and will "exit 0"
       - "FAIL" if Google Chrome isn't found, with "exit 1"
       - "NOTE" as a default status for other scenarios.

.USAGE
    Distribute as an Intune Win32 package.

.NOTES
    Keep in mind that changes in Google Chrome's installation procedures or the Windows OS 
    might necessitate modifications in the future. It's always a good practice to test 
    scripts in a controlled setting before applying them in a production environment.

.DISCLAIMER
    This script is delivered as-is without any guarantees or warranties. Always ensure 
    you have backups and take necessary precautions when executing scripts, particularly 
    in production environments.

.LAST MODIFIED
    August 6th, 2023

#>

#region Functions

function Get-ChromeExeDetails {
    <#
    .DESCRIPTION
        Searches for Google Chrome's installation location and display version
        by checking known file paths and registry paths.

    .EXAMPLE
        $chromeDetails = Get-ChromeExeDetails
        Write-Host "Google Chrome is installed at $($chromeDetails.InstallLocation) and the version is $($chromeDetails.DisplayVersion)"
    #>
    # Define the known file paths and registry paths for Google Chrome
    $chromePaths = [System.IO.Path]::Combine($env:ProgramW6432, "Google\Chrome\Application\chrome.exe"),
                   [System.IO.Path]::Combine(${env:ProgramFiles(x86)}, "Google\Chrome\Application\chrome.exe")
    $registryPaths = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome",
                     "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome",
                     "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"

    # Check the known file paths
    $installedPath = $chromePaths | Where-Object { Test-Path $_ }
    if ($installedPath) {
        $chromeDetails = New-Object PSObject -Property @{
            InstallLocation = $installedPath
            DisplayVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($installedPath).FileVersion
        }
        return $chromeDetails
    }

    # Check the known registry paths
    $registryInstalled = $registryPaths | Where-Object { Get-ItemProperty -Path $_ -ErrorAction SilentlyContinue }
    if ($registryInstalled) {
        $chromeRegistryPath = $registryInstalled | ForEach-Object { Get-ItemProperty -Path $_ }
        foreach ($registryPath in $chromeRegistryPath) {
            if ($registryPath.InstallLocation) {
                $chromeDetails = New-Object PSObject -Property @{
                    InstallLocation = [System.IO.Path]::GetFullPath((Join-Path -Path $registryPath.InstallLocation -ChildPath "..\Application\chrome.exe"))
                    DisplayVersion = $registryPath.DisplayVersion
                }
                if (Test-Path $chromeDetails.InstallLocation) {
                    return $chromeDetails
                }
            } elseif ($registryPath.'(Default)') {
                $chromeDetails = New-Object PSObject -Property @{
                    InstallLocation = $registryPath.'(Default)'
                    DisplayVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($registryPath.'(Default)').FileVersion
                }
                if (Test-Path $chromeDetails.InstallLocation) {
                    return $chromeDetails
                }
            }
        }
    }

    # If Google Chrome is not found, return $null
    return $null
}

function Test-WingetAndDependencies {
    <#
    .SYNOPSIS
    Tests for presence of Winget and required dependencies on the system.

    .DESCRIPTION
    Checks if the Windows Package Manager (Winget) is installed and verifies necessary dependencies, including the Desktop App Installer and Microsoft.UI.Xaml. Returns a status code indicating the result of the check.
    Ensures that the necessary components for Winget to function correctly are present before attempting any package installations.

    .EXAMPLE
    $checkResult = Test-WingetAndDependencies
    if ($checkResult -eq 0) {
        Write-Host "Winget and all dependencies are present."
    } elseif ($checkResult -eq 1) {
        Write-Host "Winget is not installed."
    } elseif ($checkResult -eq 2) {
        Write-Host "Winget is installed, but one or more dependencies are missing."
    }

    This example calls the Test-WingetAndDependencies function and acts based on the returned status code.

    .OUTPUTS
    Int
    Returns an integer value indicating the status of the check:
    0 - Winget and all dependencies are detected successfully.
    1 - Winget is not detected.
    2 - Winget is detected but one or more dependencies are not found.

    .NOTES
    Date: November 2, 2023

    Function does not attempt to install Winget or its dependencies. It only checks for their presence and reports the findings.

    .LINK
    Documentation for Winget: https://docs.microsoft.com/en-us/windows/package-manager/winget/
    #>

    # Initialize the summary variable
    $usrFeedback = ""
    $returnCode = 0

    # Check if Winget is installed
    $wingetPath = (Get-Command -Name winget -ErrorAction SilentlyContinue).Source
    if (-not $wingetPath) {
        $usrFeedback += "Winget NOT detected. "
        $returnCode = 1
    } else {
        $usrFeedback += "Winget detected at $wingetPath. "
    }

    # Check for Desktop App Installer
    $desktopAppInstaller = Get-AppxPackage -Name Microsoft.DesktopAppInstaller
    if (-not $desktopAppInstaller) {
        $usrFeedback += "Desktop App Installer NOT detected. "
        $returnCode = 2
    }
    else {
        $usrFeedback += "Desktop App Installer confirmed. "
    }

    # Check for Microsoft.UI.Xaml
    $uiXaml = Get-AppxPackage -Name Microsoft.UI.Xaml.2* # Assuming version 2.x is required
    if (-not $uiXaml) {
        $usrFeedback += "Microsoft.UI.Xaml NOT detected. "
        $returnCode = 2
    }
    else {
        $usrFeedback += "Microsoft.UI.Xaml confirmed. "
    }

    Write-Host $usrFeedback

    # Return the appropriate code
    return $returnCode
}

#endregion Functions


#region Main

#region Variables
$appChrome = $null              #Stores Google Chrome's details
$verChrome = $null              #for Google Chrome's version
$detectSummary = ""             #Summary of script execution
$result = 0                    #Script execution result
#endregion Variables

# Clear errors
$Error.Clear()

# Check if Google Chrome is installed
$appChrome = Get-ChromeExeDetails

# Some spaces to make it easier to read in log file
Write-Host `n`n

if ($null -ne $appChrome)  {
    # Get the current version of Google Chrome
    $verChrome = $appChrome.DisplayVersion
    Write-Host "Found Installed Chrome version $verChrome"
    $detectSummary += "Chrome Installed version = $verChrome. " 
}
else {
    Write-Host "Google Chrome not installed on device."
    $detectSummary += "Chrome not found on device. "

    # If Chrome not installed, check Winget and dependencies
    $wingetCheckResult = Test-WingetAndDependencies
    switch ($wingetCheckResult) {
        0 { $detectSummary += "Winget and all dependencies detected successfully. "}
        1 { $detectSummary += "Winget NOT detected. "}
        2 { $detectSummary += "Winget detected but one or more dependencies NOT found. " }
    }

    $result = 1
}

# Some spaces to make it easier to read in log file
Write-Host `n`n

#Return result
if ($result -eq 0) {
    Write-Host "OK $([datetime]::Now) : $detectSummary"
    Exit 0
}
elseif ($result -eq 1) {
    Write-Host "FAIL $([datetime]::Now) : $detectSummary"
    Exit 1
}
else {
    Write-Host "NOTE $([datetime]::Now) : $detectSummary"
    Exit 0
}