# GLPI Agent Unattended Deployment PowerShell Script (x64 only)
# USER SETTINGS
param (
	# Mandatory
    [string]$glpiServer = "http://YOUR_SERVER/",
    # Recomended
	[string]$expectedSha256 = "",  # Leave blank to skip
	# Optional
    [string]$setupVersion = "Latest",  # Enter 'Latest' to install the latest available version (hash is not required).
    [string]$setupLocation = "https://github.com/glpi-project/glpi-agent/releases/download/$setupVersion",
    [string]$setupNightlyLocation = "https://nightly.glpi-project.org/glpi-agent",
    [string]$setup = "GLPI-Agent-$setupVersion-x64.msi",
    [string]$Reconfigure = "Yes",
    [string]$Repair = "No",
    [string]$Verbose = "Yes",
    [string]$RunUninstallFusionInventoryAgent = "No",
    [string]$UninstallOcsAgent = "No"
)
########################################
#                                      #
#   🚫 Do not modify anything below    #
#        this line.                    #
#                                      #
########################################

$setupOptions= "/quiet RUNNOW=1 SERVER=$glpiServer"
function Is-Http {
    param ($strng) return $strng -match "^(http(s?)).*"
}
function Is-Nightly {
    param ($strng) return $strng -match "-(git[0-9a-f]{8})$"
}
function Is-InstallationNeeded {
    param ($setupVersion)
    $regPaths = @("HKLM:\SOFTWARE\GLPI-Agent\Installer", "HKLM:\SOFTWARE\Wow6432Node\GLPI-Agent\Installer")
    foreach ($path in $regPaths) {
        $currentVersion = (Get-ItemProperty -Path $path -Name "Version" -ErrorAction SilentlyContinue).Version
        if ($currentVersion) {
            if ($currentVersion -ne $setupVersion) {
                if ($Verbose -ne "No") { Write-Verbose "Installation needed: $currentVersion -> $setupVersion" -Verbose }
                return $true
            }
            return $false
        }
    }
    if ($Verbose -ne "No") { Write-Verbose "Installation needed: $setupVersion" -Verbose }
    return $true
}
function Save-WebBinary {
    param ($setupLocation, $setup)
    try {
        $url = "$setupLocation/$setup"
        $tempPath = Join-Path $env:TEMP $setup
        if ($Verbose -ne "No") { Write-Verbose "Downloading: $url" -Verbose }
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($url, $tempPath)
        if ($expectedSha256) {
            $actualHash = Get-Sha256Hash -filePath $tempPath
            if ($actualHash -ne $expectedSha256) {
                if ($Verbose -ne "No") { Write-Verbose "SHA256 hash verification failed!" -Verbose }
                Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
                return $null
            }
            if ($Verbose -ne "No") { Write-Verbose "SHA256 hash verification passed." -Verbose }
        }
        return $tempPath
    } catch {
        if ($Verbose -ne "No") { Write-Verbose "Error downloading '$url': $_" -Verbose }
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
        if ($x64Asset -and $x64Asset.digest) {
            $result = @(
                $x64Asset.browser_download_url,
                $x64Asset.digest -replace 'sha256:', ''
            )
            return $result
        } else {
            Write-Verbose "No files or digest found for version $version of Windows x64" -Verbose
            return $null
        }
    } catch {
        Write-Verbose "Error retrieving information: $_" -Verbose
        return $null
    } finally {
        $webClient.Dispose()
    }
}
function Remove-OCSAgents {
    try {
        $uninstallPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\OCS Inventory Agent",
            "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\OCS Inventory Agent",
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\OCS Inventory NG Agent",
            "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\OCS Inventory NG Agent"
        )
        foreach ($path in $uninstallPaths) {
            $uninstallString = (Get-ItemProperty -Path $path -ErrorAction SilentlyContinue).UninstallString
            if ($uninstallString) {
                Stop-Service -Name "OCS INVENTORY SERVICE" -Force -ErrorAction SilentlyContinue
                Start-Process -FilePath "cmd.exe" -ArgumentList "/C $uninstallString /S /NOSPLASH" -Wait -NoNewWindow
                Remove-Item -Path "$env:ProgramFiles\OCS Inventory Agent" -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item -Path "$env:ProgramFiles(x86)\OCS Inventory Agent" -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item -Path "$env:SystemDrive\ocs-ng" -Recurse -Force -ErrorAction SilentlyContinue
                Start-Process -FilePath "sc.exe" -ArgumentList "delete 'OCS INVENTORY'" -Wait -NoNewWindow
            }
        }
    } catch {
        if ($Verbose -ne "No") { Write-Verbose "Error removing OCS Agents: $_" -Verbose }
    }
}
function Has-Option { 
    param ($opt) $pattern = "\b$opt=.+\b"; return $setupOptions -match $pattern
}
function Is-SelectedReconfigure {
    if ($Reconfigure -ne "No") {
        if ($Verbose -ne "No") { Write-Verbose "Installation reconfigure: $setupVersion" -Verbose }
        return $true
    }
    return $false
}
function Is-SelectedRepair {
    if ($Repair -ne "No") {
        if ($Verbose -ne "No") { Write-Verbose "Installation repairing: $setupVersion" -Verbose }
        return $true
    }
    return $false
}
function Get-Sha256Hash {
    param ($filePath)
    try {
        $sha256 = Get-FileHash -Path $filePath -Algorithm SHA256 -ErrorAction Stop
        return $sha256.Hash
    } catch {
        if ($Verbose -ne "No") { Write-Verbose "Error calculating SHA256 hash: $_" -Verbose }
        return $null
    }
}
function Msi-ServerAvailable {
    $maxLoops = 120
    $loopCount = 0
    $wmiService = Get-WmiObject -Class Win32_Service -Filter "Name='MsiServer'"
    while ($loopCount -lt $maxLoops) {
        if ($loopCount -gt 0) { Start-Sleep -Seconds 1 }
        if ($wmiService.State -eq "Stopped") { return $true }
        try {
            $result = $wmiService.StopService()
            if ($result.ReturnValue -eq 0) { return $true }
        } catch {}
        $loopCount++
    }
    return $false
}
function Msi-Exec {
    param ($options)
    $maxLoops = 3
    $loopCount = 0
    $result = 0
    while ($loopCount -lt $maxLoops) {
        if ($loopCount -gt 0) {
            if ($Verbose -ne "No") { Write-Verbose "Next attempt in 30 seconds..." -Verbose }
            Start-Sleep -Seconds 30
        }
        if (Msi-ServerAvailable) {
            if ($Verbose -ne "No") { Write-Verbose "Running: MsiExec.exe $options" -Verbose }
            $process = Start-Process -FilePath "MsiExec.exe" -ArgumentList $options -Wait -PassThru -NoNewWindow
            $result = $process.ExitCode
            if ($result -ne 1618) { break }
        } else {
            $result = 1618
        }
        $loopCount++
    }
    if ($result -eq 0) {
        if ($Verbose -ne "No") { Write-Verbose "Deployment done!" -Verbose }
    } elseif ($result -eq 1618) {
        if ($Verbose -ne "No") { Write-Verbose "Deployment failed: MSI Installer is busy!" -Verbose }
    } else {
        if ($Verbose -ne "No") { Write-Verbose "Deployment failed! (Err=$result)" -Verbose }
    }
    return $result
}
function Try-DeleteOrSchedule {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    try {
        Start-Sleep 5
        Remove-Item -Path "$Path" -Force -ErrorAction Stop
        if ($Verbose -ne "No") {Write-Verbose "Deleted: $Path" -Verbose}
    } catch {
        if (-not (Test-Path $path)) {
            if ($Verbose -ne "No") {Write-Warning "File does not exist: $path"}
        }else{
           if ($Verbose -ne "No") {Write-Warning "Failed to delete the file: $path"}
        }
    }
}
################
##### MAIN #####
################
# Get the latest version if necessary
if ($UninstallOcsAgent -eq "Yes") { Remove-OCSAgents }
if ($RunUninstallFusionInventoryAgent -eq "Yes") { Uninstall-FusionInventoryAgent }
if ($env:PROCESSOR_ARCHITECTURE -ne "AMD64") {
    if ($Verbose -ne "No") {
        if ($Verbose -ne "No") {Write-Verbose "This script only supports x64 architecture. Current architecture: $env:PROCESSOR_ARCHITECTURE" -Verbose}
        if ($Verbose -ne "No") {Write-Verbose "Deployment aborted!" -Verbose}
    }
    exit 1
} else {
    if ($Verbose -ne "No") { Write-Verbose "System architecture detected: $env:PROCESSOR_ARCHITECTURE" -Verbose }
}
if ($setupVersion -eq "Latest") {
    $info = Get-GLPIAgentWin64Info
    if ($info) {
        $downloadUrl = $info[0]
        $setup = ($downloadUrl -split '/')[-1]
        $setupVersion = ($setup -replace "^GLPI-Agent-", "") -replace "-x64\.msi$", ""
        $setupLocation = $downloadUrl -replace "/$setup$", ""
        $expectedSha256 = $info[1]
        if ($Verbose -ne "No") {
            Write-Verbose "Latest version: $setupVersion" -Verbose
            Write-Verbose "Download: $setupLocation" -Verbose
            Write-Verbose "SHA256: $expectedSha256" -Verbose
        }
    } else {
        if ($Verbose -ne "No") { Write-Verbose "Failed to fetch latest version info. Deployment aborted!" -Verbose }
        exit 5
    }
}

$setup = "GLPI-Agent-$setupVersion-x64.msi"
$bInstall = $false
$installOrRepair = "/i"
if (Is-InstallationNeeded -SetupVersion $setupVersion) {
    $bInstall = $true
} elseif (Is-SelectedRepair) {
    $installOrRepair = "/fa"
    $bInstall = $true
} elseif (Is-SelectedReconfigure) {
    if (-not (Has-Option "REINSTALL")) {
        $setupOptions += " REINSTALL=feat_AGENT"
    }
    $bInstall = $true
}
if ($bInstall) {
    if (Is-Nightly $setupVersion) {
        $setupLocation = $setupNightlyLocation
    }
    if (Is-Http $setupLocation) {
        $installerPath = Save-WebBinary -SetupLocation $setupLocation -Setup $setup
        if ($installerPath) {
            $msiResult = Msi-Exec -options "$installOrRepair `"$installerPath`" $setupOptions"
            if ($Verbose -ne "No") { Write-Verbose "Deleting `"$installerPath`"" -Verbose }
            Start-Sleep -Seconds 5
            Try-DeleteOrSchedule -Path $installerPath -Verbose
        } else {
            if ($Verbose -ne "No") { Write-Verbose "Installer download or verification failed. Aborting installation." -Verbose }
            exit 6
        }
    } else {
        if ($setupLocation -and $setupLocation -ne ".") {
            $setup = Join-Path $setupLocation $setup
        }
        Msi-Exec -options "$installOrRepair `"$setup`" $setupOptions"
    }
} else {
    if ($Verbose -ne "No") { Write-Verbose "It isn't needed the installation of '$setup'." -Verbose }
}
