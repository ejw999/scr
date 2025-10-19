#---------------------------
# Enhanced File Monitor for Y:\Dwnl
#---------------------------

param(
    [string]$DownloadsPath = "Y:\Dwnl",
    [string]$LogFilePath = "$($DownloadsPath)\FileMonitor.log",
    [int]$ExeLaunchDelay = 5 # Seconds to wait before launching .exe
)

#---------------------------
# Simple file type mappings
#---------------------------
$FileTypeMappings = @{
    # Documents
    '.pdf' = 'Documents\PDFs'; '.doc' = 'Documents\Word'; '.docx' = 'Documents\Word'; '.dot' = 'Documents\Word'
    '.dotx' = 'Documents\Word'; '.docm' = 'Documents\Word'; '.dotm' = 'Documents\Word'
    '.xls' = 'Documents\Excel'; '.xlsx' = 'Documents\Excel'; '.xlsm' = 'Documents\Excel'
    '.xlsb' = 'Documents\Excel'; '.xlt' = 'Documents\Excel'; '.xltx' = 'Documents\Excel'
    '.xltm' = 'Documents\Excel'; '.csv' = 'Documents\Excel'
    '.ppt' = 'Documents\PowerPoint'; '.pptx' = 'Documents\PowerPoint'; '.pptm' = 'Documents\PowerPoint'
    '.pot' = 'Documents\PowerPoint'; '.potx' = 'Documents\PowerPoint'; '.potm' = 'Documents\PowerPoint'
    '.pps' = 'Documents\PowerPoint'; '.ppsx' = 'Documents\PowerPoint'; '.ppsm' = 'Documents\PowerPoint'
    '.txt' = 'Documents\Text'; '.rtf' = 'Documents\Text'; '.odt' = 'Documents\Text'; '.pages' = 'Documents\Text'
    '.epub' = 'Documents\eBooks'; '.mobi' = 'Documents\eBooks'; '.azw' = 'Documents\eBooks'
    '.azw3' = 'Documents\eBooks'; '.fb2' = 'Documents\eBooks'; '.chm' = 'Documents\eBooks'

    # Images
    '.jpg' = 'Images\JPEG'; '.jpeg' = 'Images\JPEG'; '.jpe' = 'Images\JPEG'; '.jfif' = 'Images\JPEG'; '.webp' = 'Images\JPEG'
    '.png' = 'Images\PNG'; '.gif' = 'Images\GIF'
    '.bmp' = 'Images\Other'; '.tiff' = 'Images\Other'; '.tif' = 'Images\Other'
    '.ico' = 'Images\Other'; '.svg' = 'Images\Other'; '.psd' = 'Images\Other'; '.ai' = 'Images\Other'; '.eps' = 'Images\Other'
    '.raw' = 'Images\Other'; '.cr2' = 'Images\Other'; '.nef' = 'Images\Other'; '.dng' = 'Images\Other'
    '.heic' = 'Images\Other'; '.heif' = 'Images\Other'

    # Videos (all → MP4)
    '.mp4' = 'Videos\MP4'; '.m4v' = 'Videos\MP4'; '.avi' = 'Videos\MP4'; '.mkv' = 'Videos\MP4'
    '.mov' = 'Videos\MP4'; '.wmv' = 'Videos\MP4'; '.flv' = 'Videos\MP4'; '.webm' = 'Videos\MP4'
    '.3gp' = 'Videos\MP4'; '.ogv' = 'Videos\MP4'; '.m2ts' = 'Videos\MP4'; '.ts' = 'Videos\MP4'
    '.mts' = 'Videos\MP4'; '.vob' = 'Videos\MP4'; '.asf' = 'Videos\MP4'; '.divx' = 'Videos\MP4'; '.xvid' = 'Videos\MP4'

    # Audio
    '.mp3' = 'Audio\MP3'; '.wav' = 'Audio\WAV'; '.flac' = 'Audio\Other'; '.aac' = 'Audio\Other'; '.ogg' = 'Audio\Other'
    '.wma' = 'Audio\Other'; '.m4a' = 'Audio\Other'; '.opus' = 'Audio\Other'; '.ape' = 'Audio\Other'; '.ac3' = 'Audio\Other'
    '.dts' = 'Audio\Other'; '.amr' = 'Audio\Other'; '.3ga' = 'Audio\Other'

    # Archives (all → ZIP folder)
    '.zip' = 'Archives\ZIP'; '.rar' = 'Archives\ZIP'; '.7z' = 'Archives\ZIP'
    '.tar' = 'Archives\ZIP'; '.gz' = 'Archives\ZIP'; '.bz2' = 'Archives\ZIP'; '.xz' = 'Archives\ZIP'
    '.lz' = 'Archives\ZIP'; '.lzma' = 'Archives\ZIP'; '.cab' = 'Archives\ZIP'
    '.iso' = 'Archives\ZIP'; '.dmg' = 'Archives\ZIP'; '.pkg' = 'Archives\ZIP'
    '.deb' = 'Archives\ZIP'; '.rpm' = 'Archives\ZIP'; '.apk' = 'Archives\ZIP'

    # Executables
    '.exe' = 'Executables\Windows'; '.msi' = 'Executables\Windows'; '.msix' = 'Executables\Windows'
    '.appx' = 'Executables\Windows'; '.bat' = 'Executables\Windows'; '.cmd' = 'Executables\Windows'
    '.ps1' = 'Executables\Windows'; '.vbs' = 'Executables\Windows'; '.js' = 'Executables\Windows'; '.jar' = 'Executables\Windows'
    '.app' = 'Executables\Other'; '.run' = 'Executables\Other'; '.bin' = 'Executables\Other'; '.sh' = 'Executables\Other'

    # Fonts
    '.ttf' = 'Fonts'; '.otf' = 'Fonts'; '.woff' = 'Fonts'; '.woff2' = 'Fonts'

    # Scripts & Configs
    '.json' = 'Scripts\Configs'; '.xml' = 'Scripts\Configs'; '.ini' = 'Scripts\Configs'; '.cfg' = 'Scripts\Configs'
    '.yml' = 'Scripts\Configs'; '.yaml' = 'Scripts\Configs'; '.psm1' = 'Scripts\Configs'

    # Miscellaneous / Other
    '.torrent' = 'Miscellaneous'; '.log' = 'Miscellaneous'; '.md' = 'Miscellaneous'; '.url' = 'Miscellaneous'
    '.desktop' = 'Miscellaneous'; '.lnk' = 'Miscellaneous'; '.tmp' = 'Miscellaneous'; '.part' = 'Miscellaneous'
}

#---------------------------
# Logging function
#---------------------------
function Write-Log {
    param(
        [string]$Message,
        [ConsoleColor]$Color = "White"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] $Message"
    Write-Host $logEntry -ForegroundColor $Color
    Add-Content -Path $LogFilePath -Value $logEntry
}

#---------------------------
# Move PowerShell console to second monitor
#---------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
}
"@

function Move-ConsoleToSecondMonitor {
    $screens = [System.Windows.Forms.Screen]::AllScreens
    if ($screens.Count -lt 2) {
        Write-Log -Message "Only one monitor detected - skipping move." -Color Yellow
        return
    }

    $second = $screens[1]

    Start-Sleep -Milliseconds 500
    $handle = [Win32]::GetForegroundWindow()

    if ($handle -ne [IntPtr]::Zero) {
        $width  = 1200
        $height = 800
        $x = $second.Bounds.X + 100
        $y = $second.Bounds.Y + 100
        [Win32]::MoveWindow($handle, $x, $y, $width, $height, $true) | Out-Null
        Write-Log -Message "Moved console to second monitor." -Color Cyan
    } else {
        Write-Log -Message "Could not find console window handle." -Color Red
    }
}

Move-ConsoleToSecondMonitor

#---------------------------
# Function to delete temporary files
#---------------------------
function Clear-TemporaryFiles {
    $tempPath = Join-Path -Path $DownloadsPath -ChildPath "Temporary"
    if (Test-Path $tempPath) {
        Write-Log -Message "Clearing temporary files from '$tempPath'" -Color DarkGray
        try {
            Get-ChildItem -Path $tempPath -File | ForEach-Object {
                Remove-Item -Path $_.FullName -Force
            }
            Write-Log -Message "Temporary files cleared." -Color Green
        } catch {
            Write-Log -Message "Error clearing temporary files: $_" -Color Red
        }
    } else {
        Write-Log -Message "Temporary folder not found: '$tempPath'" -Color DarkGray
    }
}

#---------------------------
# Setup FileSystemWatcher
#---------------------------
try {
    if (-not (Test-Path $DownloadsPath)) {
        Write-Log -Message "Path not found: $DownloadsPath" -Color Red
        exit
    }

    Write-Log -Message "Enhanced File Monitor Starting..." -Color Cyan
    Write-Log -Message "Watching folder: $DownloadsPath" -Color Green

    $watcher = New-Object System.IO.FileSystemWatcher
    $watcher.Path = $DownloadsPath
    $watcher.IncludeSubdirectories = $false
    $watcher.EnableRaisingEvents = $true

    # File Created
    Register-ObjectEvent $watcher Created -Action {
        $path = $Event.SourceEventArgs.FullPath
        $extension = [System.IO.Path]::GetExtension($path).ToLower()
        $fileName = [System.IO.Path]::GetFileName($path)

        if ($FileTypeMappings.ContainsKey($extension)) {
            $relativeFolder = $FileTypeMappings[$extension]
        } else {
            $relativeFolder = 'Miscellaneous\Unknown'
            Write-Log -Message "Unmapped extension '$extension' - using fallback folder." -Color DarkYellow
        }

        $destinationFolder = Join-Path -Path $DownloadsPath -ChildPath $relativeFolder

        if (-not (Test-Path $destinationFolder)) {
            New-Item -Path $destinationFolder -ItemType Directory | Out-Null
            Write-Log -Message "Created folder: $destinationFolder" -Color Green
        }

        $destinationPath = Join-Path -Path $destinationFolder -ChildPath $fileName

        try {
            Move-Item -Path $path -Destination $destinationPath -Force
            Write-Log -Message "Moved '$fileName' to '$relativeFolder'" -Color Yellow

            # Launch .exe files with a delay
            if ($extension -eq ".exe") {
                Start-Sleep -Seconds $ExeLaunchDelay
                Write-Log -Message "Launching executable: '$destinationPath'" -Color Green
                try {
                    Start-Process -FilePath $destinationPath
                } catch {
                    Write-Log -Message "Failed to launch '$fileName': $_" -Color Red
                }
            }

        } catch {
            Write-Log -Message "Failed to move '$fileName': $_" -Color Red
        }
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

    # Clear temporary files on startup
    Clear-TemporaryFiles

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
