function Get-OfficeVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$App
    )

    $appKeyPath = "Registry::HKEY_CLASSES_ROOT\$App.Application\CurVer"
    if (-not (Test-Path $appKeyPath)) {
        throw "Registry path not found: $appKeyPath"
    }

    $curVer = Get-ItemProperty -Path $appKeyPath
    $version = $curVer.'(default)'.Replace("$App.Application.", "") + ".0"
    return $version
}

function Set-RegistryValueIfPathExists {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Paths,
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [int]$Value
    )

    $updated = $false

    foreach ($path in $Paths) {
        if (Test-Path $path) {
            Set-ItemProperty -Path $path -Name $Name -Value $Value
            Write-Host "Set $Name=$Value at $path"
            $updated = $true
        }
    }

    if (-not $updated) {
        Write-Host "No matching registry path found for '$Name'."
    }
}

function Enable-VbaSecurity {
    param(
        [Parameter(Mandatory = $true)]
        [string]$App
    )

    $version = Get-OfficeVersion -App $App

    $accessVbomPaths = @(
        "HKCU:\Software\Microsoft\Office\$version\$App\Security",
        "HKLM:\Software\Microsoft\Office\$version\$App\Security",
        "HKLM:\Software\WOW6432Node\Microsoft\Office\$version\$App\Security",
        "HKCU:\Software\Microsoft\Office\$version\Common\TrustCenter",
        "HKLM:\Software\Microsoft\Office\$version\Common\TrustCenter"
    )

    $macroSecurityPaths = @(
        "HKCU:\Software\Microsoft\Office\$version\$App\Security",
        "HKLM:\Software\Microsoft\Office\$version\$App\Security"
    )

    Set-RegistryValueIfPathExists -Paths $accessVbomPaths -Name "AccessVBOM" -Value 1
    Set-RegistryValueIfPathExists -Paths $macroSecurityPaths -Name "VBAWarnings" -Value 1
}
