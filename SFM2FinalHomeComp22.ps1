#---------------------------
# Enhanced File Monitor for Y:\Dwnl
#---------------------------

param(
    [string]$DownloadsPath = "Y:\Dwnl",
    [string]$LogFilePath = "$($DownloadsPath)\FileMonitor.log",
    [int]$ExeLaunchDelay = 5, # Seconds to wait before launching .exe
    [int]$PollInterval = 2 # Seconds between each scan for files
)

#---------------------------
# File type mappings
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

    # Videos
    '.mp4' = 'Videos\MP4'; '.m4v' = 'Videos\MP4'; '.avi' = 'Videos\MP4'; '.mkv' = 'Videos\MP4'
    '.mov' = 'Videos\MP4'; '.wmv' = 'Videos\MP4'; '.flv' = 'Videos\MP4'; '.webm' = 'Videos\MP4'
    '.3gp' = 'Videos\MP4'; '.ogv' = 'Videos\MP4'; '.m2ts' = 'Videos\MP4'; '.ts' = 'Videos\MP4'
    '.mts' = 'Videos\MP4'; '.vob' = 'Videos\MP4'; '.asf' = 'Videos\MP4'; '.divx' = 'Videos\MP4'; '.xvid' = 'Videos\MP4'

    # Audio
    '.mp3' = 'Audio\MP3'; '.wav' = 'Audio\WAV'; '.flac' = 'Audio\Other'; '.aac' = 'Audio\Other'; '.ogg' = 'Audio\Other'
    '.wma' = 'Audio\Other'; '.m4a' = 'Audio\Other'; '.opus' = 'Audio\Other'; '.ape' = 'Audio\Other'; '.ac3' = 'Audio\Other'
    '.dts' = 'Audio\Other'; '.amr' = 'Audio\Other'; '.3ga' = 'Audio\Other'

    # Archives
    '.zip' = 'Archives\ZIP'; '.rar' = 'Archives\RAR'; '.7z' = 'Archives\7Z'
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

    # Misc
    '.torrent' = 'Miscellaneous'; '.log' = 'Miscellaneous'; '.md' = 'Miscellaneous'; '.url' = 'Miscellaneous'
    '.desktop' = 'Miscellaneous'; '.lnk' = 'Miscellaneous'; '.tmp' = 'Miscellaneous'; '.part' = 'Miscellaneous'
}

#---------------------------
# Logging function
#---------------------------
function Write-Log {
    param([string]$Message, [ConsoleColor]$Color = "White")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp] $Message"
    Write-Host $entry -ForegroundColor $Color
    Add-Content -Path $LogFilePath -Value $entry
}

#---------------------------
# Move console to second monitor
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
        Write-Log "Only one monitor detected - skipping move." Yellow
        return
    }
    $second = $screens[1]
    Start-Sleep -Milliseconds 500
    $handle = [Win32]::GetForegroundWindow()
    if ($handle -ne [IntPtr]::Zero) {
        [Win32]::MoveWindow($handle, $second.Bounds.X + 100, $second.Bounds.Y + 100, 1200, 800, $true) | Out-Null
        Write-Log "Moved console to second monitor." Cyan
    }
}
Move-ConsoleToSecondMonitor

#---------------------------
# Clear temporary files
#---------------------------
function Clear-TemporaryFiles {
    $temp = Join-Path $DownloadsPath "Temporary"
    if (Test-Path $temp) {
        Write-Log "Clearing temporary files from '$temp'" DarkGray
        try {
            Get-ChildItem $temp -File | Remove-Item -Force
            Write-Log "Temporary files cleared." Green
        } catch {
            Write-Log "Error clearing temporary files: $_" Red
        }
    }
}

#---------------------------
# Wait until file is ready
#---------------------------
function Wait-ForFileReady {
    param([string]$FilePath, [int]$MaxWaitSeconds = 300)
    $waited = 0
    while ($waited -lt $MaxWaitSeconds) {
        try {
            if ($FilePath -match '\.crdownload$' -or $FilePath -match '\.part$' -or $FilePath -match '\.tmp$') { return $false }
            $stream = [System.IO.File]::Open($FilePath, 'Open', 'Read', 'None')
            if ($stream) { $stream.Close(); return $true }
        } catch {}
        Start-Sleep -Milliseconds 500
        $waited += 0.5
    }
    return $false
}

#---------------------------
# Function: Update existing files
#---------------------------
function Update-ExistingFiles {
    Write-Log "Updating existing files in $DownloadsPath..." Cyan
    Get-ChildItem -Path $DownloadsPath -File | ForEach-Object {
        $path  = $_.FullName
        $name  = $_.Name
        $ext   = $_.Extension.ToLower()
        $folder = if ($FileTypeMappings.ContainsKey($ext)) { 
            $FileTypeMappings[$ext] 
        } else { 
            'Miscellaneous\Unknown' 
        }

        $dest = Join-Path $DownloadsPath $folder
        if (-not (Test-Path $dest)) { 
            New-Item -Path $dest -ItemType Directory | Out-Null 
        }

        try {
            Move-Item -Path $path -Destination (Join-Path $dest $name) -Force
            Write-Log "Moved existing file '$name' to '$folder'" Yellow
        } catch {
            Write-Log "Failed to move existing file '$name': $_" Red
        }
    }
    Write-Log "Existing files updated." Green
}

#---------------------------
# Function: Continuously move files from root download folder
#---------------------------
function Move-FilesFromRoot {
    try {
        # Get all files in the root of DownloadsPath (not in subdirectories)
        $files = Get-ChildItem -Path $DownloadsPath -File -ErrorAction SilentlyContinue
        
        foreach ($file in $files) {
            $path = $file.FullName
            $name = $file.Name
            $ext = $file.Extension.ToLower()
            
            # Skip log file, temporary files, and partial downloads
            if ($path -eq $LogFilePath) { continue }
            if ($name -match '\.(crdownload|part|tmp)$') { continue }
            
            # Wait for file to be ready (not locked)
            if (-not (Wait-ForFileReady -FilePath $path)) {
                Write-Log "File '$name' is still being written or locked, skipping..." DarkGray
                continue
            }
            
            # Determine destination folder
            if ($FileTypeMappings.ContainsKey($ext)) {
                $relativeFolder = $FileTypeMappings[$ext]
            } else {
                $relativeFolder = 'Miscellaneous\Unknown'
            }
            
            $destinationFolder = Join-Path -Path $DownloadsPath -ChildPath $relativeFolder
            if (-not (Test-Path $destinationFolder)) {
                New-Item -Path $destinationFolder -ItemType Directory -Force | Out-Null
                Write-Log "Created folder: $relativeFolder" Green
            }
            
            $destinationPath = Join-Path -Path $destinationFolder -ChildPath $name
            
            # Check if destination already exists and handle duplicates
            if (Test-Path $destinationPath) {
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($name)
                $counter = 1
                do {
                    $newName = "${baseName}_${counter}$ext"
                    $destinationPath = Join-Path -Path $destinationFolder -ChildPath $newName
                    $counter++
                } while (Test-Path $destinationPath)
                Write-Log "Duplicate detected, renaming to: $newName" DarkYellow
            }
            
            try {
                Move-Item -Path $path -Destination $destinationPath -Force -ErrorAction Stop
                Write-Log "Moved '$name' to '$relativeFolder'" Yellow
                
                # Launch .exe files with a delay
                if ($ext -eq ".exe") {
                    Start-Sleep -Seconds $ExeLaunchDelay
                    Write-Log "Launching executable: '$destinationPath'" Green
                    try {
                        Start-Process -FilePath $destinationPath
                    } catch {
                        Write-Log "Failed to launch '$name': $_" Red
                    }
                }
            } catch {
                Write-Log "Failed to move '$name': $_" Red
            }
        }
    } catch {
        Write-Log "Error in Move-FilesFromRoot: $_" Red
    }
}

#---------------------------
# FileSystemWatcher Setup
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

    #---------------------------
    # File Created
    #---------------------------
    Register-ObjectEvent $watcher Created -Action {
        $path = $Event.SourceEventArgs.FullPath
        if ($path -eq $LogFilePath) { return } # ✅ Ignore log file

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

    #---------------------------
    # File Changed
    #---------------------------
    Register-ObjectEvent $watcher Changed -Action {
        $path = $Event.SourceEventArgs.FullPath
        if ($path -eq $LogFilePath) { return } # ✅ Ignore log file
        Write-Log -Message "File modified: $path" -Color Blue
    }

    #---------------------------
    # File Deleted
    #---------------------------
    Register-ObjectEvent $watcher Deleted -Action {
        $path = $Event.SourceEventArgs.FullPath
        if ($path -eq $LogFilePath) { return } # ✅ Ignore log file
        Write-Log -Message "File deleted: $path" -Color Red
    }

    #---------------------------
    # File Renamed
    #---------------------------
    Register-ObjectEvent $watcher Renamed -Action {
        $old = $Event.SourceEventArgs.OldFullPath
        $new = $Event.SourceEventArgs.FullPath
        if ($new -eq $LogFilePath -or $old -eq $LogFilePath) { return } # ✅ Ignore log file
        Write-Log -Message "File renamed: '$old' -> '$new'" -Color Magenta
    }

    #---------------------------
    # Update existing files at startup
    #---------------------------
    Update-ExistingFiles

    # Clear temporary files on startup
    Clear-TemporaryFiles

    Write-Log -Message "Continuous file monitoring started. Press Ctrl + C to stop." -Color Cyan
    Write-Log -Message "Scanning every $PollInterval seconds..." -Color Green
    
    # Continuous polling loop to move files
    while ($true) {
        Move-FilesFromRoot
        Start-Sleep -Seconds $PollInterval
    }

} finally {
    if ($watcher) { $watcher.Dispose() }
    Get-EventSubscriber | Unregister-Event
    Write-Log -Message "File monitor stopped." -Color Red
}
#---------------------------
# End of Script
#---------------------------