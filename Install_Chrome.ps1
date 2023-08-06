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

#>

#region Functions

function Find-WingetPath {
    # Define possible locations for winget.exe
    $possibleLocations = @(
        "${env:ProgramFiles}\WindowsApps\Microsoft.DesktopAppInstaller_*\app\winget.exe",
        "${env:ProgramFiles(x86)}\WindowsApps\Microsoft.DesktopAppInstaller_*\app\winget.exe",
        "${env:LOCALAPPDATA}\Microsoft\WindowsApps\winget.exe",
        "${env:USERPROFILE}\AppData\Local\Microsoft\WindowsApps\winget.exe"
    )

    # Iterate through the potential locations and return the path if found
    foreach ($location in $possibleLocations) {
        try {
            $items = Get-ChildItem -Path $location -ErrorAction Stop
            if ($items) {
                return $items[0].FullName
            }
        }
        catch {
            Write-Warning "Couldn't search for winget.exe at: $location"
        }
    }

    Write-Error "Winget wasn't located in any of the specified locations."
    return $null
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

# Check if Winget is installed and, if not, find it
$wingetPath = (Get-Command -Name winget -ErrorAction SilentlyContinue).Source
if (-not $wingetPath) {
    Write-Host "Winget not detected in user path, attempting to locate in system..."
    $wingetPath = Find-WingetPath
}

if (-not $wingetPath) {
    Write-Host "Winget (Windows Package Manager) is absent on this device." 
    $detectSummary += "Winget NOT detected. "
    $result = 5
} else {
    $detectSummary += "Winget located at $wingetPath. "
}

# Use Winget to install the desired software
if ($result -eq 0) {
    try {
        $tempFile = New-TemporaryFile
        $processResult = Start-Process -FilePath "$wingetPath" -ArgumentList "install -e --id ""$WingetAppID"" --scope=machine --silent --accept-package-agreements --accept-source-agreements --force" -NoNewWindow -Wait -RedirectStandardOutput $tempFile.FullName -PassThru

        $exitCode = $processResult.ExitCode
        $installInfo = Get-Content $tempFile.FullName
        Remove-Item $tempFile.FullName

        Write-Host "Winget install exit code: $exitCode"
        Write-Host "Winget installation output: $installInfo"
        
        if ($exitCode -eq 0) {
            Write-Host "Winget successfully installed application."
            $detectSummary += "Application installed via Winget. "
            $result = 0
        } else {
            $detectSummary += "Error during installation: $installInfo, exit code: $exitCode. "
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
