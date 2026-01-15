<#
.SYNOPSIS
Automates creating a PS App Deploy Toolkit deployment for Google Chrome in MCM.

.DESCRIPTION
Copies the PSADT Chrome template into the MSI directory, drops the MSI into Files,
refreshes the Chrome-specific Invoke-AppDeployToolkit.ps1, and updates version,
MSI name, and Active Setup ProductCode references.

Includes a -Preflight mode to validate paths, read MSI ProductCode, and test write access
WITHOUT copying or editing anything.

.PARAMETER MsiPath
Literal path to the Chrome MSI (UNC supported).

.PARAMETER Preflight
If set, performs validation and stops before making changes.

.NOTES
Requires Windows with access to the template and configuration UNC shares.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$MsiPath,

    [switch]$Preflight
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Fixed sources
$TemplatePath = "\\psfs.blue.psu.edu\s9-dle\DLE\CLM Tasks Dev\CLM DEV Templates\PSADT\PSAppDeployToolkit_Template_v4.1.7"
$ChromeInvokeSource = "\\psfs.blue.psu.edu\S9-DLE\DLE\Staff\DTK\PSADT_Confgs\Chrome\Invoke-AppDeployToolkit.ps1"

function Resolve-MsiProductCode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "MSI not found: $Path"
    }

    $installer = New-Object -ComObject WindowsInstaller.Installer
    try {
        $database = $installer.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $installer, @($Path, 0))
        $view = $database.OpenView("SELECT `Value` FROM `Property` WHERE `Property`='ProductCode'")
        $view.Execute()
        $record = $view.Fetch()

        if (-not $record) {
            throw "Unable to read ProductCode from MSI: $Path"
        }

        $productCode = $record.StringData(1)
        $view.Close()

        return $productCode
    }
    finally {
        # Best-effort COM cleanup
        foreach ($obj in @($record, $view, $database, $installer)) {
            if ($null -ne $obj) {
                try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null } catch {}
            }
        }
    }
}

function Update-InvokeScript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [string]$Version,
        [Parameter(Mandatory)]
        [string]$MsiName,
        [Parameter(Mandatory)]
        [string]$ProductCode
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Invoke script not found: $Path"
    }

    $raw = Get-Content -LiteralPath $Path -Raw
    $updated = $raw
    $changed = $false

    # 1) AppVersion = '...'
    $before = $updated
    $updated = [regex]::Replace($updated, "AppVersion\s*=\s*'[^']*'", "AppVersion = '$Version'")
    if ($updated -ne $before) { $changed = $true } else { Write-Warning "Did not find AppVersion assignment to replace." }

    # 2) Start-ADTMsiProcess -Action 'Install' ... -FilePath '...'
    # More resilient: keeps any other args on that line intact; only replaces the quoted FilePath value.
    $before = $updated
    $updated = [regex]::Replace(
        $updated,
        "(?m)^(?<line>\s*Start-ADTMsiProcess\b.*?-Action\s*'Install'\b.*?-FilePath\s*)'[^']*'(?<rest>.*)$",
        "`${line}'$MsiName'`${rest}"
    )
    if ($updated -ne $before) { $changed = $true } else { Write-Warning "Did not find Start-ADTMsiProcess Install line to update." }

    # 3) Replace GUIDs in these two specific Remove-ADTRegistryKey lines
    $before = $updated
    $updated = [regex]::Replace(
        $updated,
        "(?m)^(?<pfx>\s*Remove-ADTRegistryKey\s+-Key\s+'HKLM:\\SOFTWARE\\Wow6432Node\\Microsoft\\Active Setup\\Installed Components\\)\{[0-9A-Fa-f-]+\}(?<sfx>'\s*)$",
        "`${pfx}$ProductCode`${sfx}"
    )
    if ($updated -ne $before) { $changed = $true } else { Write-Warning "Did not find Wow6432Node Active Setup key line to update." }

    $before = $updated
    $updated = [regex]::Replace(
        $updated,
        "(?m)^(?<pfx>\s*Remove-ADTRegistryKey\s+-Key\s+'HKLM:\\SOFTWARE\\Microsoft\\Active Setup\\Installed Components\\)\{[0-9A-Fa-f-]+\}(?<sfx>'\s*)$",
        "`${pfx}$ProductCode`${sfx}"
    )
    if ($updated -ne $before) { $changed = $true } else { Write-Warning "Did not find non-Wow6432Node Active Setup key line to update." }

    if ($changed) {
        Set-Content -LiteralPath $Path -Value $updated -Encoding UTF8
    }
    else {
        Write-Warning "No changes were made to $Path (patterns may not match file contents)."
    }
}

# -------- Validate input & environment --------

if (-not (Test-Path -LiteralPath $MsiPath)) {
    throw "MSI path not found: $MsiPath"
}

$resolvedMsiPath = (Resolve-Path -LiteralPath $MsiPath).ProviderPath

if (-not (Test-Path -LiteralPath $resolvedMsiPath -PathType Leaf)) {
    throw "MSI path is not a file: $resolvedMsiPath"
}

if ([IO.Path]::GetExtension($resolvedMsiPath).ToLowerInvariant() -ne '.msi') {
    throw "MSI path must point to an .msi file: $resolvedMsiPath"
}

if (-not (Test-Path -LiteralPath $TemplatePath -PathType Container)) {
    throw "Template path not found: $TemplatePath"
}

if (-not (Test-Path -LiteralPath $ChromeInvokeSource -PathType Leaf)) {
    throw "Chrome Invoke-AppDeployToolkit.ps1 not found: $ChromeInvokeSource"
}

$targetFolder = Split-Path -Path $resolvedMsiPath -Parent
$version      = Split-Path -Path $targetFolder -Leaf
$msiName      = Split-Path -Path $resolvedMsiPath -Leaf

Write-Host "=== Inputs / Derived ==="
Write-Host "MSI        : $resolvedMsiPath"
Write-Host "Target     : $targetFolder"
Write-Host "Version    : $version"
Write-Host "MSI Name   : $msiName"
Write-Host "Template   : $TemplatePath"
Write-Host "Invoke Src : $ChromeInvokeSource"

# Read product code early (fast fail)
$productCode = Resolve-MsiProductCode -Path $resolvedMsiPath
Write-Host "ProductCode: $productCode"
$productCode = [string](Resolve-MsiProductCode -Path $resolvedMsiPath)


# Ensure target exists (your flow implies it does, but this keeps things robust)
if (-not (Test-Path -LiteralPath $targetFolder -PathType Container)) {
    New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null
}

# Write test (prevents the “copied nothing because perms” time sink)
$writeTest = Join-Path $targetFolder ".psadt_write_test.$([guid]::NewGuid().ToString('N')).tmp"
try {
    New-Item -ItemType File -Path $writeTest -Force | Out-Null
    Remove-Item -LiteralPath $writeTest -Force
    Write-Host "Write test : OK"
}
catch {
    throw "Write test failed in target folder: $targetFolder. $($_.Exception.Message)"
}

if ($Preflight) {
    Write-Host "Preflight mode: stopping before copy/edit." -ForegroundColor Yellow
    return
}

# -------- Do the work --------

# Copy template contents into target
Write-Host "Copying PSADT template contents into target..."
Get-ChildItem -LiteralPath $TemplatePath -Force | ForEach-Object {
    Copy-Item -LiteralPath $_.FullName -Destination $targetFolder -Recurse -Force
}

# Copy MSI into Files folder
$filesFolder = Join-Path -Path $targetFolder -ChildPath 'Files'
if (-not (Test-Path -LiteralPath $filesFolder -PathType Container)) {
    New-Item -ItemType Directory -Path $filesFolder -Force | Out-Null
}

Write-Host "Copying MSI into Files..."
Copy-Item -LiteralPath $resolvedMsiPath -Destination (Join-Path -Path $filesFolder -ChildPath $msiName) -Force

# Refresh Invoke-AppDeployToolkit.ps1 and update it
$invokeDest = Join-Path -Path $targetFolder -ChildPath 'Invoke-AppDeployToolkit.ps1'

Write-Host "Refreshing Invoke-AppDeployToolkit.ps1 (overwrite)..."
Copy-Item -LiteralPath $ChromeInvokeSource -Destination $invokeDest -Force

Write-Host "Updating Invoke-AppDeployToolkit.ps1..."
Update-InvokeScript -Path $invokeDest -Version $version -MsiName $msiName -ProductCode $productCode

Write-Host "`nDeployment scaffolding completed" -ForegroundColor Green
Write-Host "Target     : $targetFolder"
Write-Host "Version    : $version"
Write-Host "MSI        : $msiName"
Write-Host "ProductCode: $productCode"
