function Stop-OfficeProcesses {
    $officeProcesses = @(
        "WINWORD",
        "EXCEL",
        "POWERPNT",
        "MSACCESS",
        "OUTLOOK",
        "ONENOTE",
        "MSPUB",
        "VISIO",
        "OfficeClickToRun",
        "OfficeC2RClient"
    )

    $runningProcesses = Get-Process | Where-Object { $officeProcesses -contains $_.Name }
    if ($runningProcesses.Count -eq 0) {
        Write-Host "No Office processes are running."
        return
    }

    foreach ($process in $runningProcesses) {
        try {
            $process | Stop-Process -Force -ErrorAction Stop
            Write-Host "Closed process: $($process.Name)"
        }
        catch {
            Write-Host "Could not close process $($process.Name): $($_.Exception.Message)"
        }
    }

    Start-Sleep -Seconds 5
}

function Prepare-OfficeApps {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$OfficeApps
    )

    foreach ($app in $OfficeApps) {
        Write-Host "Opening $app application"
        try {
            $comObject = New-Object -ComObject "$app.Application"
            if ($null -ne $comObject) {
                try {
                    $comObject.Quit()
                }
                catch {
                    Write-Host "Could not quit $app COM object cleanly: $($_.Exception.Message)"
                }

                try {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) | Out-Null
                }
                catch {
                    Write-Host "Could not release $app COM object: $($_.Exception.Message)"
                }
            }
            Write-Host "$app prepared successfully"
        }
        catch {
            throw "Failed to open $app application: $($_.Exception.Message)"
        }
    }

    Stop-OfficeProcesses
}
