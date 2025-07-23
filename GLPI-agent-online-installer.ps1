<#
.SYNOPSIS
    Script for the unattended installation of the GLPI agent (x64 only).
.DESCRIPTION
    Downloads, verifies, and installs the GLPI agent from GitHub or a local location. Allows configuration of installation options, uninstallation of OCS agents, and verification of system architecture.
.PARAMETER setupOptions
    Installation options for MsiExec (e.g., '/quiet RUNNOW=1 SERVER="http://..."'). Required.
.PARAMETER expectedSha256
    Expected SHA256 hash of the installer. Leave blank to skip verification.
.PARAMETER setupVersion
    Version of the agent to install. Use 'Latest' for the latest available version.
.PARAMETER setupLocation
    URL or local path of the installer. By default, uses GitHub for the specified version.
.PARAMETER setupNightlyLocation
    URL for nightly versions of the GLPI agent.
.PARAMETER setup
    Temporary path of the downloaded MSI file.
.PARAMETER allowVerbose
    Enables verbose messages if set to 'Yes'. Default: 'Yes'.
.PARAMETER runUninstallFusionInventoryAgent
    Uninstalls the FusionInventory agent if set to 'Yes'. Default: 'No'. (Function not implemented).
.PARAMETER uninstallOcsAgent
    Uninstalls the OCS Inventory agent if set to 'Yes'. Default: 'No'.
.PARAMETER reconfigure
    Reconfigures the existing installation if set to 'Yes'. Default: 'No'.
.PARAMETER repair
    Repairs the existing installation if set to 'Yes'. Default: 'No'.
#>
# GLPI Agent Unattended Deployment PowerShell Script (x64 only)
# USER SETTINGS
param (
	[string]$setupOptions = '/quiet RUNNOW=1 SERVER=http://soporte.ctl.tech/marketplace/glpiinventory/ ADD_FIREWALL_EXCEPTION=1 ADDLOCAL=feat_AGENT,feat_DEPLOY EXECMODE=1',
    [string]$expectedSha256 = "",
    [string]$setupVersion = "Latest",
    [string]$setupLocation = "https://github.com/glpi-project/glpi-agent/releases/download/$setupVersion",
    [string]$setupNightlyLocation = "https://nightly.glpi-project.org/glpi-agent",
    [string]$setup = "$env:temp\GLPI-Agent-$setupVersion-x64.msi",
    [string]$allowVerbose = "Yes",
    [string]$runUninstallFusionInventoryAgent = "No",
    [string]$uninstallOcsAgent = "No",
    [string]$reconfigure = "No",
    [string]$repair = "No"
)
function Test-Http {
    param ($strng)
    return [System.Uri]::IsWellFormedUriString($strng, [System.UriKind]::Absolute)
}
function Test-Nightly {
    param ($strng)
    return $strng -match "-(git[0-9a-f]{8})$"
}
function Test-InstallationNeeded {
    param ($setupVersion)
    $regPaths = @("HKLM:\SOFTWARE\GLPI-Agent\Installer", "HKLM:\SOFTWARE\Wow6432Node\GLPI-Agent\Installer")
    foreach ($path in $regPaths) {
        $currentVersion = (Get-ItemProperty -Path $path -Name "Version" -ErrorAction SilentlyContinue).Version
        if ($currentVersion) {
            if ($currentVersion -ne $setupVersion) {
                if ($allowVerbose -ne "No") { Write-Verbose "Installation needed: $currentVersion -> $setupVersion" -Verbose }
                return $true
            }
            return $false
        }
    }
    if ($allowVerbose -ne "No") { Write-Verbose "Installation needed: $setupVersion" -Verbose }
    return $true
}
function Save-WebBinary {
    param ($setupLocation, $setup)
    try {
        $url = "$setupLocation/$setup"
        $tempPath = Join-Path $env:TEMP $setup
        if ($allowVerbose -ne "No") { Write-Verbose "Downloading: $url" -Verbose }
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($url, $tempPath)
        if ($expectedSha256) {
            $actualHash = Get-Sha256Hash -filePath $tempPath
            if ($actualHash -ne $expectedSha256) {
                if ($allowVerbose -ne "No") { Write-Verbose "SHA256 hash verification failed!" -Verbose }
                Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
                return $null
            }
            if ($allowVerbose -ne "No") { Write-Verbose "SHA256 hash verification passed." -Verbose }
        }
        return $tempPath
    } catch {
        if ($allowVerbose -ne "No") { Write-Verbose "Error downloading '$url': $_" -Verbose }
        return $null
    } finally {
        if ($webClient) { $webClient.Dispose() }
    }
}
function Get-GLPIAgentWin64Info {
    $webClient = New-Object System.Net.WebClient
    $releasesUrl = "https://api.github.com/repos/glpi-project/glpi-agent/releases/latest"
    try {
        $webClient.Headers.Add("User-Agent", "PowerShell")
        $releaseJson = $webClient.DownloadString($releasesUrl)
        $release = ConvertFrom-Json $releaseJson
        $version = $release.tag_name
        $x64Asset = $release.assets | Where-Object { $_.name -like "GLPI-Agent-$version-x64.msi" }
        if ($x64Asset -and $x64Asset.digest -and ($x64Asset.digest -match 'sha256:([0-9a-fA-F]+)')) {
            $result = @(
                $x64Asset.browser_download_url,
                $matches[1]
            )
            return $result
        } else {
            if ($allowVerbose -ne "No") { Write-Verbose "No files or digest found for version $version of Windows x64" -Verbose }
            return $null
        }
    } catch {
        if ($allowVerbose -ne "No") { Write-Verbose "Error retrieving information: $_" -Verbose }
        return $null
    } finally {
        $webClient.Dispose()
    }
}
function Invoke-OCSAgentCleanup {
    try {
        $uninstallPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\OCS Inventory Agent",
            "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\OCS Inventory Agent",
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\OCS Inventory NG Agent",
            "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\OCS Inventory NG Agent"
        )
        foreach ($path in $uninstallPaths) {
            $uninstallString = (Get-ItemProperty -Path $path -ErrorAction SilentlyContinue).UninstallString
            if ($uninstallString -and (Test-Path $path)) {
                Stop-Service -Name "OCS INVENTORY SERVICE" -Force -ErrorAction SilentlyContinue
                Start-Process -FilePath "cmd.exe" -ArgumentList "/C $uninstallString /S /NOSPLASH" -Wait -NoNewWindow
                if (Test-Path "$env:ProgramFiles\OCS Inventory Agent") {
                    Remove-Item -Path "$env:ProgramFiles\OCS Inventory Agent" -Recurse -Force -ErrorAction SilentlyContinue
                }
                if (Test-Path "$env:ProgramFiles(x86)\OCS Inventory Agent") {
                    Remove-Item -Path "$env:ProgramFiles(x86)\OCS Inventory Agent" -Recurse -Force -ErrorAction SilentlyContinue
                }
                if (Test-Path "$env:SystemDrive\ocs-ng") {
                    Remove-Item -Path "$env:SystemDrive\ocs-ng" -Recurse -Force -ErrorAction SilentlyContinue
                }
                Start-Process -FilePath "sc.exe" -ArgumentList "delete 'OCS INVENTORY'" -Wait -NoNewWindow
            }
        }
    } catch {
        if ($allowVerbose -ne "No") { Write-Verbose "Error removing OCS Agents: $_" -Verbose }
    }
}
function Test-OptionPresent {
    param ($opt)
    $pattern = "\b$opt=.+\b"
    return $setupOptions -match $pattern
}
function Test-SelectedReconfigure {
    if ($reconfigure -ne "No") {
        if ($allowVerbose -ne "No") { Write-Verbose "Installation reconfigure: $setupVersion" -Verbose }
        return $true
    }
    return $false
}
function Test-SelectedRepair {
    if ($repair -ne "No") {
        if ($allowVerbose -ne "No") { Write-Verbose "Installation repairing: $setupVersion" -Verbose }
        return $true
    }
    return $false
}
function Get-Sha256Hash {
    param ($filePath)
    try {
        $sha256 = (-join ([System.Security.Cryptography.SHA256]::Create().ComputeHash([System.IO.File]::OpenRead($filePath)) | ForEach-Object { $_.ToString("x2") }))
        return $sha256
    } catch {
        if ($allowVerbose -ne "No") { Write-Verbose "Error calculating SHA256 hash: $_" -Verbose }
        return $null
    }
}
function Test-MsiServerAvailable {
    $maxLoops = 6
    $loopCount = 0
    while ($loopCount -lt $maxLoops) {
        $wmiService = Get-CimInstance -ClassName Win32_Service -Filter "Name='MsiServer'"
		if ($loopCount -gt 0) { Start-Sleep -Seconds 10 }
        if ($wmiService.State -eq "Stopped") { return $true }
        try {
            $result = $wmiService.StopService()
            if ($result.ReturnValue -eq 0) { return $true }
        } catch {
            if ($allowVerbose -ne "No") { Write-Verbose "Could not determine MsiServer status!" -Verbose }
        }
        $loopCount++
    }
    return $false
}
function Invoke-MsiExec {
    param ($options, $setup)
    $maxLoops = 3
    $loopCount = 0
    $result = 0
    if (-not (Test-Path $setup)) {
        if ($allowVerbose -ne "No") { Write-Verbose "Installer file not found: $setup" -Verbose }
        return 1
    }
    while ($loopCount -lt $maxLoops) {
        if ($loopCount -gt 0) {
            if ($allowVerbose -ne "No") { Write-Verbose "Next attempt in 30 seconds..." -Verbose }
            Start-Sleep -Seconds 30
        }
        if (Test-MsiServerAvailable) {
            if ($allowVerbose -ne "No") { Write-Verbose "Running: MsiExec.exe $options" -Verbose }
            $process = Start-Process -FilePath "MsiExec.exe" -ArgumentList $options -Wait -PassThru -NoNewWindow
            $result = $process.ExitCode
            if ($result -ne 1618) { break }
        } else {
            $result = 1618
        }
        $loopCount++
    }
    if ($result -eq 0) {
        if ($allowVerbose -ne "No") { Write-Verbose "Deployment done!" -Verbose }
    } elseif ($result -eq 1618) {
        if ($allowVerbose -ne "No") { Write-Verbose "Deployment failed: MSI Installer is busy!" -Verbose }
    } else {
        if ($allowVerbose -ne "No") { Write-Verbose "Deployment failed! (Err=$result)" -Verbose }
    }
    return $result
}
function Invoke-DeleteOrSchedule {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    try {
        Start-Sleep 5
        Remove-Item -Path "$Path" -Force -ErrorAction Stop
        if ($allowVerbose -ne "No") { Write-Verbose "Deleted: $Path" -Verbose }
    } catch {
        if (-not (Test-Path $Path)) {
            if ($allowVerbose -ne "No") { Write-Verbose "File does not exist: $Path" -Verbose }
        } else {
            if ($allowVerbose -ne "No") { Write-Verbose "Failed to delete the file: $Path" -Verbose }
        }
    }
}
################
##### MAIN #####
################
if ($uninstallOcsAgent -eq "Yes") { Invoke-OCSAgentCleanup }
if ($env:PROCESSOR_ARCHITECTURE -ne "AMD64") {
    if ($allowVerbose -ne "No") {
        Write-Verbose "This script only supports x64 architecture. Current architecture: $env:PROCESSOR_ARCHITECTURE" -Verbose
        Write-Verbose "Deployment aborted!" -Verbose
    }
    exit 1
} else {
    if ($allowVerbose -ne "No") { Write-Verbose "System architecture detected: $env:PROCESSOR_ARCHITECTURE" -Verbose }
}
if ($setupVersion -eq "Latest") {
    $info = Get-GLPIAgentWin64Info
    if ($info) {
        $downloadUrl = $info[0]
        $setup = ($downloadUrl -split '/')[-1]
        $setupVersion = ($setup -replace "^GLPI-Agent-", "") -replace "-x64\.msi$", ""
        $setupLocation = $downloadUrl -replace "/$setup$", ""
        $expectedSha256 = $info[1]
        if ($allowVerbose -ne "No") {
            Write-Verbose "Latest version: $setupVersion" -Verbose
            Write-Verbose "Download: $setupLocation" -Verbose
            Write-Verbose "SHA256: $expectedSha256" -Verbose
        }
    } else {
        if ($allowVerbose -ne "No") { Write-Verbose "Failed to fetch latest version info. Deployment aborted!" -Verbose }
        exit 5
    }
}
$setup = "GLPI-Agent-$setupVersion-x64.msi"
$bInstall = $false
$installOrRepair = "/i"
if (Test-InstallationNeeded -SetupVersion $setupVersion) {
    $bInstall = $true
} elseif (Test-SelectedRepair) {
    $installOrRepair = "/fa"
    $bInstall = $true
} elseif (Test-SelectedReconfigure) {
    if (-not (Test-OptionPresent "REINSTALL")) {
        $setupOptions += " REINSTALL=feat_AGENT"
    }
    $bInstall = $true
}
if ($bInstall) {
    if (Test-Nightly $setupVersion) {
        $setupLocation = $setupNightlyLocation
    }
    if (Test-Http $setupLocation) {
        $installerPath = Save-WebBinary -SetupLocation $setupLocation -Setup $setup
        if ($installerPath) {
            $null = Invoke-MsiExec -options "$installOrRepair `"$installerPath`" $setupOptions" -setup $installerPath
            if ($allowVerbose -ne "No") { Write-Verbose "Deleting `"$installerPath`"" -Verbose }
            Invoke-DeleteOrSchedule -Path $installerPath
        } else {
            if ($allowVerbose -ne "No") { Write-Verbose "Installer download or verification failed. Aborting installation." -Verbose }
            exit 6
        }
    } else {
        if ($setupLocation -and $setupLocation -ne ".") {
            $setup = Join-Path $setupLocation $setup
            if (-not (Test-Path $setup)) {
                if ($allowVerbose -ne "No") { Write-Verbose "Local installer not found: $setup. Aborting installation." -Verbose }
                exit 7
            }
        }
        Invoke-MsiExec -options "$installOrRepair `"$setup`" $setupOptions" -setup $setup
    }
} else {
    if ($allowVerbose -ne "No") { Write-Verbose "No installation needed for '$setup'." -Verbose }
}