# MEM_AppWin32_InstallChrome

## Overview

The `MEM_AppWin32_InstallChrome` repository offers scripts designed to ensure a consistently up-to-date installation of Google Chrome. These can be deployed as a Win32 package.

While some apps, such as Google Chrome, are not yet available via the "New" Microsoft Store method, this repository provides a solution. By harnessing the capabilities of `winget`, Microsoft's package manager for Windows, these scripts guarantee that upon deployment, devices always receive the latest version of Google Chrome. This approach not only streamlines the deployment process but also ensures devices are protected from potential vulnerabilities present in outdated software versions.

By employing `winget`, devices are set to fetch the latest Google Chrome version directly, eliminating the need for manual update packaging.

A special thanks to John Bryntze! His videos, referenced [here](https://www.youtube.com/watch?v=0Ov4AcRM4jI) and [here](https://www.youtube.com/watch?v=MnFL2FQLjp4), provide a comprehensive understanding of creating a Win32 package to distribute software installations using PowerShell and `winget`. They explain packaging, setting detection rules, and addressing challenges.

## Scripts in this Repository

- **[Detect_Chrome.ps1](https://github.com/jmanuelng/MEM_AppWin32_InstallChrome/blob/main/Detect_Chrome.ps1)**: This script checks if Google Chrome is already installed on the device.
  
- **[Install_Chrome.ps1](https://github.com/jmanuelng/MEM_AppWin32_InstallChrome/blob/main/Install_Chrome.ps1)**: Using `winget`, this script facilitates the installation of the most recent Google Chrome version.

## Usage

1. **Preparation**: Before deploying the scripts, ensure you have the IntuneWinAppUtil tool. This utility is essential for packaging the scripts into a format suitable for Intune.

2. **Packaging for Intune**:
   - Navigate to the directory containing the scripts and the [IntuneWinAppUtil](https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool) tool.
   - Run the IntuneWinAppUtil tool.
   - When prompted, provide the source folder, the setup file (script), and the output folder.
   - The tool will generate an `.intunewin` file, which is suitable for uploading to Intune.

3. **Uploading to Intune**:
   - Go to the Microsoft Endpoint Manager admin center.
   - Navigate to Apps > All apps > Add.
   - Select `Windows app (Win32)` from the list.
   - Upload the `.intunewin` file generated in the previous step.
   - Configure the app information, settings, and assignments as needed.
   - Save and assign the app to desired group(s).

4. **Setting Detection Rules**:
   - For detection, utilize the `Detect_Chrome.ps1` script. This script verifies if Google Chrome is already installed on the device, ensuring the installation process is initiated only when necessary.

5. **Deployment**:
   - Once the app is assigned, devices in the target group(s) will receive the latest version of Google Chrome upon their next check-in with Intune.

## Credits

A special acknowledgment to [John Bryntze](https://twitter.com/JohnBryntze), creator of the referenced YouTube videos, for his tutorials on packaging applications for Intune using `winget`:
- [Packaging Zoom with Winget and Intune - Part 1](https://www.youtube.com/watch?v=0Ov4AcRM4jI)
- [Packaging Zoom with Winget and Intune - Part 2](https://www.youtube.com/watch?v=MnFL2FQLjp4)

