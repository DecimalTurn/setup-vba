param(
    [string]$OfficeApps = "Excel,Word,PowerPoint,Access"
)

$ErrorActionPreference = "Stop"

function Get-NormalizedOfficeApps {
    param(
        [Parameter(Mandatory = $true)]
        [string]$OfficeApps
    )

    $validApps = @("Excel", "Word", "PowerPoint", "Access")
    $resolved = @()

    foreach ($rawApp in ($OfficeApps -split ",")) {
        $candidate = $rawApp.Trim()
        if ([string]::IsNullOrWhiteSpace($candidate)) {
            continue
        }

        $match = $validApps | Where-Object { $_ -ieq $candidate } | Select-Object -First 1
        if (-not $match) {
            throw "Unsupported Office app '$candidate'. Valid values: $($validApps -join ', ')."
        }

        if ($resolved -notcontains $match) {
            $resolved += $match
        }
    }

    if ($resolved.Count -eq 0) {
        throw "No valid Office apps were provided."
    }

    return $resolved
}

. "$PSScriptRoot/Prepare-OfficeApps.ps1"
. "$PSScriptRoot/Enable-VbaSecurity.ps1"

$apps = Get-NormalizedOfficeApps -OfficeApps $OfficeApps
Write-Host "Preparing Office apps: $($apps -join ', ')"

Prepare-OfficeApps -OfficeApps $apps

foreach ($app in $apps) {
    Write-Host "Configuring VBA security for $app"
    Enable-VbaSecurity -App $app
}

"office-apps=$($apps -join ',')" | Out-File -FilePath $env:GITHUB_OUTPUT -Append -Encoding utf8
Write-Host "Setup complete."
