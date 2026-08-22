function Minimize-Window {
    param (
        [Parameter(Mandatory=$true)]
        [string]$WindowTitlePart
    )

    if (-not ("NativeWindow" -as [type])) {
        Add-Type @"
using System;
using System.Runtime.InteropServices;
public class NativeWindow {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
    }

    Get-Process |
        Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -like "*$WindowTitlePart*" } |
        ForEach-Object {
            [NativeWindow]::ShowWindow($_.MainWindowHandle, 6) | Out-Null
            Write-Host "Minimized window: $($_.MainWindowTitle)"
        }
}