function Take-Screenshot {
    param (
        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $screen = [System.Windows.Forms.Screen]::PrimaryScreen
    $bounds = $screen.Bounds
    $bitmap = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
    $graphic = [System.Drawing.Graphics]::FromImage($bitmap)

    try {
        $graphic.CopyFromScreen($bounds.X, $bounds.Y, 0, 0, $bounds.Size)
        $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $OutputPath = $OutputPath -replace "{{timestamp}}", $timestamp
        $bitmap.Save($OutputPath)
    }
    finally {
        $graphic.Dispose()
        $bitmap.Dispose()
    }

    Write-Host "Screenshot saved to: $OutputPath"
}