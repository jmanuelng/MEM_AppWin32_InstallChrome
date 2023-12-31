<#
.DESCRIPTION
    Automate the process of installing software using the Windows Package Manager (Winget).
    It primarily focuses on Google Chrome, but the structure can be adjusted to support other software applications by changing the WingetAppID variable.

.EXAMPLE
    $WingetAppID = "Your.ApplicationID"
    .\ScriptName.ps1
    Installs the application associated with "Your.ApplicationID" using Winget, 
    assuming Winget is available on the device.

.NOTES
    Inspired by John Bryntze; Twitter: @JohnBryntze

.DISCLAIMER
    This script is delivered as-is without any guarantees or warranties. Always ensure 
    you have backups and take necessary precautions when executing scripts, particularly 
    in production environments.

.LAST MODIFIED
    November 9th, 2023

#>

#region Functions

function Invoke-Ensure64bitEnvironment {
    <#
    .SYNOPSIS
        Check if the script is running in a 32-bit or 64-bit environment, and relaunch using 64-bit PowerShell if necessary.

    .NOTES
        This script checks the processor architecture to determine the environment.
        If it's running in a 32-bit environment on a 64-bit system (WOW64), 
        it will relaunch using the 64-bit version of PowerShell.
        Place the function at the beginning of the script to ensure a switch to 64-bit when necessary.
    #>
    if ($ENV:PROCESSOR_ARCHITECTURE -eq "x86" -and $ENV:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
        Write-Output "Detected 32-bit PowerShell on 64-bit system. Relaunching script in 64-bit environment..."
        Start-Process -FilePath "$ENV:WINDIR\SysNative\WindowsPowershell\v1.0\PowerShell.exe" -ArgumentList "-WindowStyle Hidden -NonInteractive -File `"$($PSCommandPath)`" " -Wait -NoNewWindow
        exit # Terminate the 32-bit process
    } elseif ($ENV:PROCESSOR_ARCHITECTURE -eq "x86") {
        Write-Output "Detected 32-bit PowerShell on a 32-bit system. Stopping script execution."
        exit # Terminate the script if it's a pure 32-bit system
    }
}

function Find-WingetPath {
    <#
    .SYNOPSIS
        Locates the winget.exe executable within a system.

    .DESCRIPTION
        Finds the path of the `winget.exe` executable on a Windows system. 
        Aimed at finding Winget when main script is executed as SYSTEM, but will also work under USER
        
        Windows Package Manager (`winget`) is a command-line tool that facilitates the 
        installation, upgrade, configuration, and removal of software packages. Identifying the 
        exact path of `winget.exe` allows for execution (installations) under SYSTEM context.

        METHOD
        1. Defining Potential Paths:
        - Specifies potential locations of `winget.exe`, considering:
            - Standard Program Files directory (64-bit systems).
            - 32-bit Program Files directory (32-bit applications on 64-bit systems).
            - Local application data directory.
            - Current user's local application data directory.
        - Paths may utilize wildcards (*) for flexible directory naming, e.g., version-specific folder names.

        2. Iterating Through Paths:
        - Iterates over each potential location.
        - Resolves paths containing wildcards to their actual path using `Resolve-Path`.
        - For each valid location, uses `Get-ChildItem` to search for `winget.exe`.

        3. Returning Results:
        - If `winget.exe` is located, returns the full path to the executable.
        - If not found in any location, outputs an error message and returns `$null`.

    .EXAMPLE
        $wingetLocation = Find-WingetPath
        if ($wingetLocation) {
            Write-Output "Winget found at: $wingetLocation"
        } else {
            Write-Error "Winget was not found on this system."
        }

    .NOTES
        While this function is designed for robustness, it relies on current naming conventions and
        structures used by the Windows Package Manager's installation. Future software updates may
        necessitate adjustments to this function.

    .DISCLAIMER
        This function and script is provided as-is with no warranties or guarantees of any kind. 
        Always test scripts and tools in a controlled environment before deploying them in a production setting.
        
        This function's design and robustness were enhanced with the assistance of ChatGPT, it's important to recognize that 
        its guidance, like all automated tools, should be reviewed and tested within the specific context it's being 
        applied. 

    #>
    # Define possible locations for winget.exe
    $possibleLocations = @(
        "${env:ProgramFiles}\WindowsApps\Microsoft.DesktopAppInstaller*_x64__8wekyb3d8bbwe\winget.exe", 
        "${env:ProgramFiles(x86)}\WindowsApps\Microsoft.DesktopAppInstaller*_8wekyb3d8bbwe\winget.exe",
        "${env:LOCALAPPDATA}\Microsoft\WindowsApps\winget.exe",
        "${env:USERPROFILE}\AppData\Local\Microsoft\WindowsApps\winget.exe"
    )

    # Iterate through the potential locations and return the path if found
    foreach ($location in $possibleLocations) {
        try {
            # Resolve path if it contains a wildcard
            if ($location -like '*`**') {
                $resolvedPaths = Resolve-Path $location -ErrorAction SilentlyContinue
                # If the path is resolved, update the location for Get-ChildItem
                if ($resolvedPaths) {
                    $location = $resolvedPaths.Path
                }
                else {
                    # If path couldn't be resolved, skip to the next iteration
                    Write-Warning "Couldn't resolve path for: $location"
                    continue
                }
            }
            
            # Try to find winget.exe using Get-ChildItem
            $items = Get-ChildItem -Path $location -ErrorAction Stop
            if ($items) {
                Write-Host "Found Winget at: $items"
                return $items[0].FullName
                break
            }
        }
        catch {
            Write-Warning "Couldn't search for winget.exe at: $location"
        }
    }

    Write-Error "Winget wasn't located in any of the specified locations."
    return $null
}

function Install-VisualCIfMissing {
    <#
    .SYNOPSIS
        Checks for the presence of Microsoft Visual C++ Redistributable on the system and installs it if missing.

    .DESCRIPTION
        This function is designed to ensure that the Microsoft Visual C++ 2015-2022 Redistributable (x64) is installed on the system.
        It checks the system's uninstall registry keys for an existing installation of the specified version of Visual C++ Redistributable.
        If not found, proceeds to download the installer from the official Microsoft link and installs it silently without user interaction.
        Function returns a boolean value indicating the success or failure of the installation or the presence of the redistributable.


    .PARAMETER vcRedistUrl
        The URL from which the Visual C++ Redistributable installer will be downloaded.
        Default is set to the latest supported Visual C++ Redistributable direct download link from Microsoft.

    .PARAMETER vcRedistFilePath
        The local file path where the Visual C++ Redistributable installer will be downloaded to.
        Default is set to the Windows TEMP directory with the filename 'vc_redist.x64.exe'.

    .PARAMETER vcDisplayName
        The display name of the Visual C++ Redistributable to check for in the system's uninstall registry keys.
        This is used to determine if the redistributable is already installed.

    .EXAMPLE
        $vcInstalled = Install-VisualCIfMissing
        This example calls the function and stores the result in the variable $vcInstalled.
        After execution, $vcInstalled will be true if the redistributable is installed, otherwise false.

    .NOTES
        This function requires administrative privileges to install the Visual C++ Redistributable.
        Ensure that the script is run in a context that has the necessary permissions.

        The function uses the Start-Process cmdlet to execute the installer, which requires the '-Wait' parameter to ensure
        that the installation process completes before the script proceeds.

        Error handling is implemented to catch any exceptions during the download and installation process.
        If an error occurs, the function will return false and output the error message.

        It is recommended to test this function in a controlled environment before deploying it in a production setting.

    .LINK
        For more information on Microsoft Visual C++ Redistributable, visit:
        https://docs.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist

    #>

    # Define the Visual C++ Redistributable download URL and file path
    $vcRedistUrl = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
    $vcRedistFilePath = "$env:TEMP\vc_redist.x64.exe"

    # Define the display name for the Visual C++ Redistributable to check if it's installed
    $vcDisplayName = "Microsoft Visual C++ 2015-2022 Redistributable (x64)"

    # Check if Visual C++ Redistributable is already installed
    $vcInstalled = Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall |
                   Get-ItemProperty |
                   Where-Object { $_.DisplayName -like "*$vcDisplayName*" }

    if ($vcInstalled) {
        # Visual C++ is already installed, no action needed
        Write-Host "Microsoft Visual C++ Redistributable is already installed."
        return $true
    } else {
        # Visual C++ is not installed, proceed with download and installation
        Write-Host "Microsoft Visual C++ Redistributable not found. Attempting to install..."

        # Attempt to download the Visual C++ Redistributable installer
        try {
            Invoke-WebRequest -Uri $vcRedistUrl -OutFile $vcRedistFilePath -ErrorAction Stop
            Write-Host "Download of Visual C++ Redistributable succeeded."
        } catch {
            # Log detailed error message and halt execution if download fails
            Write-Error "Failed to download Visual C++ Redistributable: $($_.Exception.Message)"
            return $false
        }

        # Attempt to install the Visual C++ Redistributable
        try {
            # Start the installer and wait for it to complete, capturing the process object
            $process = Start-Process -FilePath $vcRedistFilePath -ArgumentList '/install', '/quiet', '/norestart' -Wait -PassThru -ErrorAction Stop
            # Check the exit code of the installer process to determine success
            if ($process.ExitCode -eq 0) {
                Write-Host "Successfully installed Microsoft Visual C++ Redistributable."
                return $true
            } else {
                # Log detailed error message if installation fails
                Write-Error "Visual C++ Redistributable installation failed with exit code $($process.ExitCode)."
                return $false
            }
        } catch {
            # Log detailed error message and halt execution if installation process fails
            Write-Error "Failed to install Visual C++ Redistributable: $($_.Exception.Message)"
            return $false
        }
    }
                
}


function Get-LoggedOnUser {
    <#
    .SYNOPSIS
    Retrieves the user identifier of the currently logged-on user or the most recently logged-on user based on active explorer.exe processes.

    .DESCRIPTION
    The function performs a two-step verification to determine the active user on a Windows system. 
    Initially, it attempts to identify the currently logged-on user via the Win32_ComputerSystem class. 
    Should this approach fail, it proceeds to evaluate all running explorer.exe processes to ascertain 
    which user session was initiated most recently.

    .OUTPUTS
    System.String
    Outputs a string in the format "DOMAIN\Username" representing the active or most recent user. 
    If no user can be identified, it outputs $null.

    .EXAMPLE
    $UserId = Get-LoggedOnUser
    if ($UserId) {
        Write-Host "The active or most recently connected user is: $UserId"
    } else {
        Write-Host "Unable to identify the active or most recently connected user."
    }

    In this example, the function retrieves the user identifier of the active or most recent user
    and prints it to the console. If no user can be determined, it conveys an appropriate message.

    .NOTES
    Execution context: This function is intended to be run with administrative privileges to ensure accurate retrieval of user information.

    Assumptions: This function assumes that the presence of an explorer.exe process correlates with an interactive user session
     and utilizes this assumption to determine the user identity.

    Error Handling: If the function encounters any issues while attempting to identify the user via Win32_ComputerSystem,
     it outputs a warning and falls back to the process-based identification method.

    #>

    # Initialization and Win32_ComputerSystem user retrieval
    $loggedOnUser = $null
    try {
        $loggedOnUser = Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty UserName
        if ($loggedOnUser) {
            return $loggedOnUser
        }
    } catch {
        Write-Warning "Query to Win32_ComputerSystem failed to retrieve the logged-on user."
    }

    # Fallback method using explorer.exe processes
    $explorerProcesses = Get-WmiObject Win32_Process -Filter "name = 'explorer.exe'"
    $userSessions = @()
    foreach ($process in $explorerProcesses) {
        $ownerInfo = $process.GetOwner()
        $startTime = $process.ConvertToDateTime($process.CreationDate)
        $userSessions += New-Object PSObject -Property @{
            User      = "$($ownerInfo.Domain)\$($ownerInfo.User)"
            StartTime = $startTime
        }
    }

    # Identification of the most recent user session
    $mostRecentUserSession = $userSessions | Sort-Object StartTime -Descending | Select-Object -First 1
    if ($mostRecentUserSession) {
        return $mostRecentUserSession.User
    } else {
        return $null
    }
}


function Install-WingetAsSystem {
    <#
    .SYNOPSIS
        Installs the Windows Package Manager (winget) as a system app by creating a scheduled task.

    .DESCRIPTION
        This function creates a scheduled task that runs a PowerShell script to install the latest version of winget and its dependencies.
        It is designed to install winget in the system context, making it available to all users on the device.
        The installation script is adapted from the winget-pkgs repository on GitHub, ensuring the latest version and dependencies are installed.

        The function is based on the 'InstallWingetAsSystem' function from the 'Winget-InstallPackage.ps1' script
        by djust270, which can be found at:
        https://github.com/djust270/Intune-Scripts/blob/master/Winget-InstallPackage.ps1

        The installation script within the function is adapted from:
        https://github.com/microsoft/winget-pkgs/blob/master/Tools/SandboxTest.ps1

    .EXAMPLE
        Install-WingetAsSystem

        Installs winget as a system app by creating and running a scheduled task.

    .NOTES
        Administrative privileges are required to create scheduled tasks and install winget as a system app.

    .LINK
        Original script source for Install-WingetAsSystem: https://github.com/djust270/Intune-Scripts/blob/master/Winget-InstallPackage.ps1
        Original script source for winget installation: https://github.com/microsoft/winget-pkgs/blob/master/Tools/SandboxTest.ps1

    #>
    # PowerShell script block that will be executed by the scheduled task
    $scriptBlock = @'
        # Function to install the latest version of WinGet and its dependencies
        function Install-WinGet {
            $tempFolderName = 'WinGetInstall'
            $tempFolder = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath $tempFolderName
            New-Item $tempFolder -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
            
            $apiLatestUrl = if ($Prerelease) { 'https://api.github.com/repos/microsoft/winget-cli/releases?per_page=1' }
            else { 'https://api.github.com/repos/microsoft/winget-cli/releases/latest' }
            
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $WebClient = New-Object System.Net.WebClient
            
            function Get-LatestUrl
            {
                ((Invoke-WebRequest $apiLatestUrl -UseBasicParsing | ConvertFrom-Json).assets | Where-Object { $_.name -match '^Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle$' }).browser_download_url
            }
            
            function Get-LatestHash
            {
                $shaUrl = ((Invoke-WebRequest $apiLatestUrl -UseBasicParsing | ConvertFrom-Json).assets | Where-Object { $_.name -match '^Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.txt$' }).browser_download_url
                
                $shaFile = Join-Path -Path $tempFolder -ChildPath 'Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.txt'
                $WebClient.DownloadFile($shaUrl, $shaFile)
                
                Get-Content $shaFile
            }
            
            $desktopAppInstaller = @{
                fileName = 'Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle'
                url	     = $(Get-LatestUrl)
                hash	 = $(Get-LatestHash)
            }
            
            $vcLibsUwp = @{
                fileName = 'Microsoft.VCLibs.x64.14.00.Desktop.appx'
                url	     = 'https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx'
                hash	 = '9BFDE6CFCC530EF073AB4BC9C4817575F63BE1251DD75AAA58CB89299697A569'
            }
            $uiLibsUwp = @{
                fileName = 'Microsoft.UI.Xaml.2.7.zip'
                url	     = 'https://www.nuget.org/api/v2/package/Microsoft.UI.Xaml/2.7.0'
                hash	 = '422FD24B231E87A842C4DAEABC6A335112E0D35B86FAC91F5CE7CF327E36A591'
            }

            $dependencies = @($desktopAppInstaller, $vcLibsUwp, $uiLibsUwp)
            
            Write-Host '--> Checking dependencies'
            
            foreach ($dependency in $dependencies)
            {
                $dependency.file = Join-Path -Path $tempFolder -ChildPath $dependency.fileName
                #$dependency.pathInSandbox = (Join-Path -Path $tempFolderName -ChildPath $dependency.fileName)
                
                # Only download if the file does not exist, or its hash does not match.
                if (-Not ((Test-Path -Path $dependency.file -PathType Leaf) -And $dependency.hash -eq $(Get-FileHash $dependency.file).Hash))
                {
                    Write-Host "`t- Downloading: `n`t$($dependency.url)"
                    
                    try
                    {
                        $WebClient.DownloadFile($dependency.url, $dependency.file)
                    }
                    catch
                    {
                        #Pass the exception as an inner exception
                        throw [System.Net.WebException]::new("Error downloading $($dependency.url).", $_.Exception)
                    }
                    if (-not ($dependency.hash -eq $(Get-FileHash $dependency.file).Hash))
                    {
                        throw [System.Activities.VersionMismatchException]::new('Dependency hash does not match the downloaded file')
                    }
                }
            }
            
            # Extract Microsoft.UI.Xaml from zip (if freshly downloaded).
            # This is a workaround until https://github.com/microsoft/winget-cli/issues/1861 is resolved.
            
            if (-Not (Test-Path (Join-Path -Path $tempFolder -ChildPath \Microsoft.UI.Xaml.2.7\tools\AppX\x64\Release\Microsoft.UI.Xaml.2.7.appx)))
            {
                Expand-Archive -Path $uiLibsUwp.file -DestinationPath ($tempFolder + '\Microsoft.UI.Xaml.2.7') -Force
            }
            $uiLibsUwp.file = (Join-Path -Path $tempFolder -ChildPath \Microsoft.UI.Xaml.2.7\tools\AppX\x64\Release\Microsoft.UI.Xaml.2.7.appx)
            Add-AppxPackage -Path $($desktopAppInstaller.file) -DependencyPath $($vcLibsUwp.file), $($uiLibsUwp.file)
            # Clean up files
            Remove-Item $tempFolder -recurse -force
    }
    # Call the Install-WinGet function to perform the installation
    Install-WinGet
'@

    # Name for Temp Script.
    $tmpScript = "WingetScript.ps1"
    
    # Ensure the automation directory exists
    if (!(Test-Path "$env:systemdrive\automation")) {
        New-Item "$env:systemdrive\automation" -ItemType Directory | Out-Null
    }

    # Write the script block to a file in the automation directory
    $scriptBlock | Out-File "$env:systemdrive\automation\$tmpScript"

    # Create the scheduled task action to run the PowerShell script
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-executionpolicy bypass -WindowStyle minimized -file %SYSTEMDRIVE%\automation\$tmpScript"

    # Create the scheduled task trigger to run at log on
    $trigger = New-ScheduledTaskTrigger -AtLogOn

    # Get the current user's username to set as the principal of the task
    $UserId = Get-LoggedOnUser
    $principal = New-ScheduledTaskPrincipal -UserId $UserId

    # Create the scheduled task
    $task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal

    # Register and start the scheduled task
    Register-ScheduledTask RunScript -InputObject $task
    Start-ScheduledTask -TaskName RunScript

    # Wait for the task to complete
    Start-Sleep -Seconds 120

    # Unregister and remove the scheduled task and script file
    Unregister-ScheduledTask -TaskName RunScript -Confirm:$false
    Remove-Item "$env:systemdrive\automation\$tmpScript"
}

#endregion Functions

#region Main

#region Initialization
$wingetPath = ""                # Path to Winget executable
$detectSummary = ""             # Script execution summary
$result = 0                     # Exit result (default to 0)
$WingetAppID = "Google.Chrome"  # Winget Application ID
$processResult = $null          # Winget process result
$exitCode = $null               # Software installation exit code
$installInfo                    # Information about the Winget installation process
#endregion Initialization

# Make the log easier to read
Write-Host `n`n

# Invoke the function to ensure we're running in a 64-bit environment if available
Invoke-Ensure64bitEnvironment
Write-Host "Script running in 64-bit environment."

# Find if Visual C++ redistributable is installed using Install-VisualCIfMissing function and capture the result
$vcInstalled = Install-VisualCIfMissing

if ($vcInstalled) {
    $detectSummary += "Visual C++ Redistributable installed. "
} else {
    $detectSummary += "Failed to verify or install Visual C++ Redistributable. "
    $result = 5
}

# Check if Winget is available and, if not, find it
$wingetPath = (Get-Command -Name winget -ErrorAction SilentlyContinue).Source

if (-not $wingetPath) {
    Write-Host "Winget not detected, attempting to locate in system..."
    $wingetPath = Find-WingetPath
}

# If not present, try to install it
if (-not $wingetPath) {
    Write-Host "Trying to install latest Winget using Install-WingetAsSystem...."
    Install-WingetAsSystem
    $wingetPath = Find-WingetPath
}

# If still not present, notify, or maybe it did find it, yei!!
if (-not $wingetPath) {
    Write-Host "Winget (Windows Package Manager) is absent on this device." 
    $detectSummary += "Winget NOT detected. "
    $result = 6
} else {
    $detectSummary += "Winget located at $wingetPath. "
    $result = 0
}

# Use Winget to install the desired software
if ($result -eq 0) {
    try {
        $tempFile = New-TemporaryFile
        Write-Host "Initiating App $WingetAppID Installation"
        $processResult = Start-Process -FilePath "$wingetPath" -ArgumentList "install -e --id ""$WingetAppID"" --scope=machine --silent --accept-package-agreements --accept-source-agreements --force" -NoNewWindow -Wait -RedirectStandardOutput $tempFile.FullName -PassThru

        $exitCode = $processResult.ExitCode
        $installInfo = Get-Content $tempFile.FullName
        Remove-Item $tempFile.FullName

        Write-Host "Winget install exit code: $exitCode"
        #Write-Host "Winget installation output: $installInfo"          #Remove comment to troubleshoot.
        
        if ($exitCode -eq 0) {
            Write-Host "Winget successfully installed application."
            $detectSummary += "Installed $WingetAppID via Winget. "
            $result = 0
        } else {
            $detectSummary += "Error during installation, exit code: $exitCode. "
            $result = 1
        }
    }
    catch {
        Write-Host "Encountered an error during installation: $_"
        $detectSummary += "Installation failed with exit code $($processResult.ExitCode). "
        $result = 1
    }
}

# Simplify reading in the AgentExecutor Log
Write-Host `n`n

# Output the final results
if ($result -eq 0) {
    Write-Host "OK $([datetime]::Now) : $detectSummary"
    Exit 0
} elseif ($result -eq 1) {
    Write-Host "FAIL $([datetime]::Now) : $detectSummary"
    Exit 1
} else {
    Write-Host "NOTE $([datetime]::Now) : $detectSummary"
    Exit 0
}

#endregion Main
