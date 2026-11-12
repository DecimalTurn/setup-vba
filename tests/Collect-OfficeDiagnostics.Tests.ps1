$ErrorActionPreference = "Stop"
$root = Join-Path ([System.IO.Path]::GetTempPath()) "setup-vba-diagnostics-test-$([guid]::NewGuid())"
$oldRunnerTemp = $env:RUNNER_TEMP
$oldTemp = $env:TEMP

try {
    $env:RUNNER_TEMP = Join-Path $root "runner"
    $env:TEMP = Join-Path $root "temp"
    $destination = Join-Path $root "diagnostics"
    $extraction = Join-Path $env:TEMP "chocolatey/Office365ProPlus"
    New-Item -ItemType Directory -Force -Path $env:RUNNER_TEMP, $extraction, $destination | Out-Null
    $setup = Join-Path $extraction "Setup.exe"
    Set-Content $setup "dummy"
    Set-Content (Join-Path $extraction "OfficeInstall.log") "extraction"
    Set-Content (Join-Path $env:TEMP "OfficeSetup.log") "download"
    $oldLog = Join-Path $env:TEMP "old.log"
    Set-Content $oldLog "old"
    (Get-Item $oldLog).LastWriteTimeUtc = [DateTime]::UtcNow.AddDays(-1)
    $chocoLog = Join-Path $root "chocolatey.log"
    @(
        "old Chocolatey entry"
        "Completed download of officedeploymenttool_17531-20046.exe"
        "Running Start-ChocolateyProcessAsAdmin -workingDirectory 'C:\work' -statements '/download C:\config.xml extra' -exeToRun '$setup'"
        "Command [`"$setup`" /download C:\config.xml extra] exited with '30183'."
    ) | Set-Content $chocoLog

    & "$PSScriptRoot/../scripts/Collect-OfficeDiagnostics.ps1" `
        -Destination $destination -Started ([DateTime]::UtcNow.ToString("o")) `
        -ChocolateyLog $chocoLog -ChocolateyLogLines 1

    $details = Get-Content (Join-Path $destination "odt-details.txt") -Raw
    $installLog = Get-Content (Join-Path $destination "chocolatey-install.log") -Raw
    if ($details -notmatch '30183' -or $details -notmatch 'workingDirectory' -or
        $details -notmatch '/download C:\\config.xml extra' -or
        $details -notmatch '17531-20046' -or $details -notmatch 'ODT file version:') {
        throw "Missing ODT invocation, exit code, working directory, or version details."
    }
    if ($installLog -match 'old Chocolatey entry') { throw "Collected a previous Chocolatey entry." }
    if ((Get-Content (Join-Path $destination "extraction/OfficeInstall.log") -Raw).Trim() -ne "extraction") {
        throw "Extraction log was not preserved separately."
    }
    if ((Get-Content (Join-Path $destination "process-temp/OfficeSetup.log") -Raw).Trim() -ne "download") {
        throw "Download log was not preserved."
    }
    if (Test-Path (Join-Path $destination "process-temp/old.log")) { throw "Collected a stale temp log." }

    $missingDestination = Join-Path $root "missing"
    New-Item -ItemType Directory $missingDestination | Out-Null
    & "$PSScriptRoot/../scripts/Collect-OfficeDiagnostics.ps1" `
        -Destination $missingDestination -Started ([DateTime]::UtcNow.AddDays(1).ToString("o")) `
        -ChocolateyLog (Join-Path $root "missing.log") -ChocolateyLogLines 0
    $missingDetails = Get-Content (Join-Path $missingDestination "odt-details.txt") -Raw
    if ($missingDetails -notmatch 'Chocolatey log not found' -or
        $missingDetails -notmatch 'No recent top-level logs' -or
        $missingDetails -notmatch 'No recent ODT bootstrapper extraction log found') {
        throw "Missing diagnostic sources were not reported."
    }
    if (Test-Path (Join-Path $missingDestination "extraction/OfficeInstall.log")) {
        throw "Collected a stale extraction log."
    }
    Write-Host "Office diagnostics tests passed."
}
finally {
    $env:RUNNER_TEMP = $oldRunnerTemp
    $env:TEMP = $oldTemp
    if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
}
