param(
    [Parameter(Mandatory)] [string] $Destination,
    [Parameter(Mandatory)] [string] $Started,
    [Parameter(Mandatory)] [string] $ChocolateyLog,
    [Parameter(Mandatory)] [int] $ChocolateyLogLines
)

$ErrorActionPreference = "Stop"
$startedAt = [DateTime]::Parse($Started).ToUniversalTime().AddSeconds(-2)
$details = [System.Collections.Generic.List[string]]::new()
$details.Add("Collection started for install at $Started (UTC).")

if (Test-Path $ChocolateyLog) {
    $lines = @(Get-Content -Path $ChocolateyLog)
    if ($lines.Count -lt $ChocolateyLogLines) {
        Write-Warning "Chocolatey log was truncated during installation; collecting the available log."
        $ChocolateyLogLines = 0
    }
    $installLog = @($lines | Select-Object -Skip $ChocolateyLogLines)
    $installLog | Set-Content -Path (Join-Path $Destination "chocolatey-install.log")
    $bootstrapperLines = @($installLog | Where-Object {
        $_ -match 'officedeploymenttool_[^ ''"]+\.exe' -and
        ($_ -match 'Completed download of' -or $_ -match 'Running Install-ChocolateyInstallPackage')
    })
    if ($bootstrapperLines.Count) {
        $details.Add("ODT bootstrapper download/version information from Chocolatey:")
        foreach ($line in $bootstrapperLines) { $details.Add($line) }
    }
    $downloadLines = @($installLog | Where-Object {
        $_ -match '/download' -and
        ($_ -match 'Running Start-ChocolateyProcessAsAdmin' -or
         $_ -match 'Elevating permissions and running \[' -or
         $_ -match 'Command \[' -or $_ -match 'ERROR: Running \[')
    })
    if ($downloadLines.Count) {
        $details.Add("Actual ODT download invocation and result from Chocolatey:")
        foreach ($line in $downloadLines) { $details.Add($line) }
        $startLine = $downloadLines | Where-Object { $_ -match "Running Start-ChocolateyProcessAsAdmin" } | Select-Object -Last 1
        if ($startLine -match "-exeToRun '([^']+)'") {
            $setupPath = $Matches[1]
            if (Test-Path -LiteralPath $setupPath) {
                $version = (Get-Item -LiteralPath $setupPath).VersionInfo
                $details.Add("ODT executable: $setupPath")
                $fileVersion = if ($version.FileVersion) { $version.FileVersion } else { "unavailable" }
                $productVersion = if ($version.ProductVersion) { $version.ProductVersion } else { "unavailable" }
                $details.Add("ODT file version: $fileVersion")
                $details.Add("ODT product version: $productVersion")
            } else {
                $details.Add("ODT executable no longer available for version inspection: $setupPath")
            }
        } else {
            $details.Add("ODT executable and version unavailable in Chocolatey log.")
        }
    } else {
        $details.Add("ODT download invocation not found in Chocolatey log.")
    }
} else {
    $details.Add("Chocolatey log not found: $ChocolateyLog")
}

$roots = @(
    @{ Name = "runner-temp"; Path = $env:RUNNER_TEMP },
    @{ Name = "process-temp"; Path = $env:TEMP }
)
foreach ($root in $roots) {
    if (-not $root.Path -or -not (Test-Path -LiteralPath $root.Path)) { continue }
    $files = @(Get-ChildItem -LiteralPath $root.Path -File -Filter "*.log" |
        Where-Object { $_.LastWriteTimeUtc -ge $startedAt })
    if ($files.Count) {
        $target = Join-Path $Destination $root.Name
        New-Item -ItemType Directory -Path $target -Force | Out-Null
        foreach ($file in $files) {
            Copy-Item -LiteralPath $file.FullName -Destination $target
            $details.Add("Collected temp log: $($file.FullName)")
        }
    } else {
        $details.Add("No recent top-level logs found in $($root.Path).")
    }
}

$extractionLog = Join-Path $env:TEMP "chocolatey/Office365ProPlus/OfficeInstall.log"
if ((Test-Path -LiteralPath $extractionLog) -and
    (Get-Item -LiteralPath $extractionLog).LastWriteTimeUtc -ge $startedAt) {
    $target = Join-Path $Destination "extraction"
    New-Item -ItemType Directory -Path $target -Force | Out-Null
    Copy-Item -LiteralPath $extractionLog -Destination $target
    $details.Add("Collected ODT bootstrapper extraction log: $extractionLog")
} else {
    $details.Add("No recent ODT bootstrapper extraction log found: $extractionLog")
}

$details | Set-Content -Path (Join-Path $Destination "odt-details.txt")
Write-Host "Office diagnostics directory: $Destination"
Get-ChildItem -LiteralPath $Destination -File -Recurse | ForEach-Object {
    Write-Host "  $($_.FullName)"
}
Write-Host "ODT diagnostics:"
$details | ForEach-Object { Write-Host $_ }
