#---------------------------
# Simple File Monitor for Y:\Dwnl
#---------------------------

param(
    [string]$DownloadsPath = "Y:\Dwnl"
)

#---------------------------
# Logging function
#---------------------------
function Write-Log {
    param(
        [string]$Message,
        [ConsoleColor]$Color = "White"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] $Message" -ForegroundColor $Color
}

#---------------------------
# Setup FileSystemWatcher
#---------------------------
try {
    if (-not (Test-Path $DownloadsPath)) {
        Write-Log -Message "Path not found: $DownloadsPath" -Color Red
        exit
    }

    Write-Log -Message "Simple File Monitor Starting..." -Color Cyan
    Write-Log -Message "Watching folder: $DownloadsPath" -Color Green

    $watcher = New-Object System.IO.FileSystemWatcher
    $watcher.Path = $DownloadsPath
    $watcher.IncludeSubdirectories = $false
    $watcher.EnableRaisingEvents = $true

    # File Created
    Register-ObjectEvent $watcher Created -Action {
        $path = $Event.SourceEventArgs.FullPath
        Write-Log -Message "File created: $path" -Color Yellow
    }

    # File Changed
    Register-ObjectEvent $watcher Changed -Action {
        $path = $Event.SourceEventArgs.FullPath
        Write-Log -Message "File modified: $path" -Color Blue
    }

    # File Deleted
    Register-ObjectEvent $watcher Deleted -Action {
        $path = $Event.SourceEventArgs.FullPath
        Write-Log -Message "File deleted: $path" -Color Red
    }

    # File Renamed
    Register-ObjectEvent $watcher Renamed -Action {
        $old = $Event.SourceEventArgs.OldFullPath
        $new = $Event.SourceEventArgs.FullPath
        Write-Log -Message "File renamed: '$old' -> '$new'" -Color Magenta
    }

    Write-Log -Message "Monitoring started. Press Ctrl + C to stop." -Color Cyan
    while ($true) { Start-Sleep -Seconds 1 }

} finally {
    if ($watcher) { $watcher.Dispose() }
    Get-EventSubscriber | Unregister-Event
    Write-Log -Message "File monitor stopped." -Color Red
}

#---------------------------
# End of Script
#---------------------------
