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
    November 9th, 2023

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
    Tests for the presence of Winget and required dependencies on the system.

    .DESCRIPTION
    Checks if the Windows Package Manager (Winget) is installed and verifies necessary dependencies, 
    including the Desktop App Installer, Microsoft.UI.Xaml, and the Visual C++ Redistributable. 
    Returns a string with unique identifiers indicating the result of the check and outputs feedback to the console.
    This allows for precise identification of which components are missing.

    .EXAMPLE
    $checkResult = Test-WingetAndDependencies
    if ($checkResult -eq "0") {
        Write-Host "Winget and all dependencies are present."
    } else {
        Write-Host "Missing components: $checkResult"
    }
    This example calls the Test-WingetAndDependencies function and acts based on the returned status string.

    .OUTPUTS
    String
    Returns a string value with concatenated identifiers indicating the status of the check:
    "0" - Winget and all dependencies are detected successfully.
    "W" - Winget is not detected.
    "D" - Desktop App Installer is not detected.
    "U" - Microsoft.UI.Xaml is not detected.
    "V" - Visual C++ Redistributable is not detected.
    Concatenated string for multiple missing components, e.g., "DU" for missing Desktop App Installer and Microsoft.UI.Xaml.

    .NOTES
    Date: November 9, 2023
    The function does not attempt to install Winget or its dependencies. It only checks for their presence, reports the findings, and outputs feedback to the console.

    .LINK
    Documentation for Winget: https://docs.microsoft.com/en-us/windows/package-manager/winget/
    #>

    # Initialize an array to hold missing component identifiers
    $missingComponents = @()

    # Check if Winget is installed
    $wingetPath = (Get-Command -Name winget -ErrorAction SilentlyContinue).Source
    if (-not $wingetPath) {
        $missingComponents += "W" # Add 'W' to the array if Winget is missing
        Write-Host "Winget is NOT installed."
    } else {
        Write-Host "Winget is installed."
    }

    # Check for Desktop App Installer
    $desktopAppInstaller = Get-AppxPackage -Name Microsoft.DesktopAppInstaller -ErrorAction SilentlyContinue
    if (-not $desktopAppInstaller) {
        $missingComponents += "D" # Add 'D' to the array if Desktop App Installer is missing
        Write-Host "Desktop App Installer is NOT installed."
    } else {
        Write-Host "Desktop App Installer is installed."
    }

    # Check for Microsoft.UI.Xaml
    $uiXaml = Get-AppxPackage -Name Microsoft.UI.Xaml.2* -ErrorAction SilentlyContinue # Assuming version 2.x is required
    if (-not $uiXaml) {
        $missingComponents += "U" # Add 'U' to the array if Microsoft.UI.Xaml is missing
        Write-Host "Microsoft.UI.Xaml is NOT installed."
    } else {
        Write-Host "Microsoft.UI.Xaml is installed."
    }

    # Check for Visual C++ Redistributable
    $vcDisplayName = "Microsoft Visual C++ 2015-2022 Redistributable (x64)"
    $vcInstalled = Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall, 
                                  HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall |
                   Get-ItemProperty |
                   Where-Object { $_.DisplayName -like "*$vcDisplayName*" } -ErrorAction SilentlyContinue
    if (-not $vcInstalled) {
        $missingComponents += "V" # Add 'V' to the array if Visual C++ Redistributable is missing
        Write-Host "Visual C++ Redistributable is NOT installed."
    } else {
        Write-Host "Visual C++ Redistributable is installed."
    }

    # Return a concatenated string of missing component identifiers
    # If no components are missing, return '0'
    if ($missingComponents.Length -eq 0) {
        return "0"
    } else {
        return [String]::Join('', $missingComponents)
    }
}

function Test-InternetConnectivity {
    <#
    .SYNOPSIS
    Confirms internet connectivity to download content from github.com and nuget.org.

    .DESCRIPTION
    Tests the TCP connection to github.com and nuget.org on port 443 (HTTPS) to confirm internet connectivity.
    Returns a string of characters that clearly identifies if there is a connectivity issue, and if so, to which URL or site.
    Additionally, outputs simplified but clear feedback to the console.

    .EXAMPLE
    Test-InternetConnectivity
    This example calls the Test-InternetConnectivity function and outputs the result to the console.

    .OUTPUTS
    String
    Returns a string of characters indicating the connectivity status:
    '0' - No connectivity issues.
    'G' - Connectivity issue with github.com.
    'N' - Connectivity issue with nuget.org.
    'GN' - Connectivity issues with both sites.

    .NOTES
    Date: November 2, 2023
    #>

    # Initialize a variable to hold the connectivity status
    $connectivityStatus = ''

    # Test connectivity to github.com
    $githubTest = Test-NetConnection -ComputerName 'github.com' -Port 443 -ErrorAction SilentlyContinue
    if (-not $githubTest.TcpTestSucceeded) {
        $connectivityStatus += 'G'
        Write-Host "Connectivity issue with github.com."
    } else {
        Write-Host "Successfully connected to github.com."
    }

    # Test connectivity to nuget.org
    $nugetTest = Test-NetConnection -ComputerName 'nuget.org' -Port 443 -ErrorAction SilentlyContinue
    if (-not $nugetTest.TcpTestSucceeded) {
        $connectivityStatus += 'N'
        Write-Host "Connectivity issue with nuget.org."
    } else {
        Write-Host "Successfully connected to nuget.org."
    }

    # Determine the return value based on the tests
    if ($connectivityStatus -eq '') {
        Write-Host "Internet connectivity to both github.com and nuget.org is confirmed."
        return '0' # No issues
    } else {
        Write-Host "Connectivity test completed with issues: $connectivityStatus"
        return $connectivityStatus # Return the specific issue(s)
    }
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

if ($null -ne $appChrome) {
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
    # Adjust the switch to handle string identifiers
    switch -Regex ($wingetCheckResult) {
        '0' { 
            $detectSummary = "Winget and all dependencies detected successfully. " # Set summary exclusively for this case
            break # Exit the switch to avoid processing other cases
        }
        'W' { $detectSummary += "Winget NOT detected. " }
        'D' { $detectSummary += "Desktop App Installer NOT detected. " }
        'U' { $detectSummary += "Microsoft.UI.Xaml NOT detected. " }
        'V' { $detectSummary += "Visual C++ Redistributable NOT detected. " }
        Default { $detectSummary += "Unknown dependency check result: $wingetCheckResult " }
    }

    # Check internet connectivity to github.com and nuget.org
    $internetConnectivityResult = Test-InternetConnectivity
    # Adjust the switch to handle string identifiers for connectivity results
    switch -Regex ($internetConnectivityResult) {
        '0' { 
            $detectSummary += "Connectivity to github.com and nuget.org confirmed. "
        }
        'G' { $detectSummary += "Connectivity issue with github.com. " }
        'N' { $detectSummary += "Connectivity issue with nuget.org. " }
        'GN' { $detectSummary += "Connectivity issues with both github.com and nuget.org. " }
        Default { $detectSummary += "Unknown connectivity check result: $internetConnectivityResult " }
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