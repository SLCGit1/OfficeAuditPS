# OfficeAuditv3.ps1
# Created by Jesus Ayala From Sarah Lawrence College
# Audits Office type, architecture, name, version, and displays details in the PowerShell console

function Get-InteractiveUser {
    try {
        $user = (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty UserName)
        if ([string]::IsNullOrWhiteSpace($user)) {
            return "N/A"
        } else {
            return ($user -replace '^.*\\')  # Strip domain if present
        }
    } catch {
        return "N/A"
    }
}

# Initialize values
$ComputerName     = $env:COMPUTERNAME
$CurrentUser      = Get-InteractiveUser
$OSVersion        = (Get-CimInstance Win32_OperatingSystem).Caption
$OfficeArch       = "Unknown"
$OfficeType       = "None"
$OfficeVersion    = "Not Found"
$OfficeBuild      = "Unknown"
$OfficeVersionMap = "Unknown"
$OfficeName       = "Unknown"
$Timestamp        = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Detect Click-to-Run Office via registry
$C2RPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
)

foreach ($path in $C2RPaths) {
    if (Test-Path $path) {
        try {
            $props = Get-ItemProperty -Path $path
            if ($props.Platform) {
                $OfficeArch = $props.Platform
                $OfficeType = "Click-to-Run"
            }
            if ($props.VersionToReport) {
                $OfficeVersion = $props.VersionToReport
                if ($OfficeVersion -match "^16\.0\.(\d+)\.") {
                    $OfficeBuild = $Matches[1]
                }
            }
            if ($props.ProductReleaseIds) {
                $OfficeName = $props.ProductReleaseIds -replace ";.*", ""  # Grab first product ID
            }
            break
        } catch {}
    }
}

# If not Click-to-Run, check for MSI Office using WMI
if ($OfficeType -eq "None") {
    try {
        $OfficeProducts = Get-WmiObject -Class Win32_Product | Where-Object {
            $_.Name -match "Microsoft Office" -and $_.Name -notmatch "Click" -and $_.Name -notmatch "365"
        }

        foreach ($product in $OfficeProducts) {
            if ($product -ne $null) {
                $OfficeType    = "MSI"
                $OfficeVersion = $product.Version
                $OfficeName    = $product.Name
                if ($OfficeVersion -match "^16\.0\.(\d+)\.") {
                    $OfficeBuild = $Matches[1]
                }
                break
            }
        }

        $OfficeExePaths = @(
            "$env:ProgramFiles\Microsoft Office",
            "$env:ProgramFiles (x86)\Microsoft Office"
        )

        foreach ($base in $OfficeExePaths) {
            $exe = Get-ChildItem -Path $base -Recurse -Include "winword.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($exe) {
                $OfficeArch = if ($exe.FullName -like "*Program Files (x86)*") { "32-bit" } else { "64-bit" }
                break
            }
        }

    } catch {
        $OfficeType    = "MSI Detection Failed"
        $OfficeArch    = "Unknown"
        $OfficeVersion = "Unknown"
        $OfficeBuild   = "Unknown"
        $OfficeName    = "Unknown"
    }
}

# Map build to marketing version
if ($OfficeVersion -match "^16\.0\.(\d+)\.") {
    $build = [int]$Matches[1]
    $OfficeBuild = $build
    switch ($build) {
        { $_ -ge 17000 } { $OfficeVersionMap = "2024"; break }
        { $_ -ge 14000 } { $OfficeVersionMap = "2021"; break }
        { $_ -ge 10300 } { $OfficeVersionMap = "2019"; break }
        { $_ -ge 4266  } { $OfficeVersionMap = "2016"; break }
        default          { $OfficeVersionMap = "16.x (Unknown)" }
    }
} elseif ($OfficeVersion -match "^15\.0") {
    $OfficeVersionMap = "2013"
} elseif ($OfficeVersion -match "^14\.0") {
    $OfficeVersionMap = "2010"
}

# Display results
$AuditSummary = [PSCustomObject]@{
    ComputerName       = $ComputerName
    Username           = $CurrentUser
    OSVersion          = $OSVersion
    OfficeType         = $OfficeType
    OfficeArch         = $OfficeArch
    OfficeName         = $OfficeName
    OfficeVersion      = $OfficeVersionMap
    OfficeBuildNumber  = $OfficeBuild
    OfficeFullVersion  = $OfficeVersion
    Timestamp          = $Timestamp
}

Write-Host "`nOffice Audit Summary:`n" -ForegroundColor Cyan
$AuditSummary | Format-List
