# Simple File Monitor for Y:\Dwnl
param(
    [string]$DownloadsPath = "Y:\Dwnl"
)

Write-Host "Simple File Monitor Starting..." -ForegroundColor Green
Write-Host "Monitoring: $DownloadsPath" -ForegroundColor Cyan

# Ensure directory exists
if (!(Test-Path $DownloadsPath)) {
    Write-Host "Creating directory: $DownloadsPath" -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $DownloadsPath -Force | Out-Null
}

# Simple file type mappings
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

function Move-FileToFolder {
    param(
        [string]$FilePath,
        [string]$TargetFolder
    )
    
    $fileName = Split-Path $FilePath -Leaf
    $targetPath = Join-Path $TargetFolder $fileName
    $extension = [System.IO.Path]::GetExtension($fileName).ToLower()
    
    try {
        if (!(Test-Path $TargetFolder)) {
            Write-Host "Creating directory: $TargetFolder" -ForegroundColor Yellow
            New-Item -ItemType Directory -Path $TargetFolder -Force | Out-Null
        }
        
        Write-Host "Moving: $fileName" -ForegroundColor Cyan
        Write-Host "   From: $FilePath" -ForegroundColor Gray
        Write-Host "   To:   $targetPath" -ForegroundColor Gray
        
        Move-Item $FilePath $targetPath -Force
        Write-Host "Successfully moved: $fileName -> $TargetFolder" -ForegroundColor Green

        # Open Explorer, show popup, and run executable if it's a Windows exe
        if ($extension -in '.exe', '.msi', '.bat') {
            try {
                Start-Process explorer.exe "/select,`"$targetPath`""
                
                $wshell = New-Object -ComObject WScript.Shell
                $wshell.Popup("Executable moved: $fileName", 3, "File Moved", 64)

                if ($extension -eq '.exe') {
                    Start-Process $targetPath
                    Write-Host "Launched executable: $fileName" -ForegroundColor Green
                }
            } catch {
                Write-Host "Could not open Explorer, show notification, or launch executable: $targetPath" -ForegroundColor Yellow
            }
        }

        return $true
    }
    catch {
        Write-Host "Error moving $fileName : $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Invoke-FileProcessing {
    param([string]$FilePath)
    
    $fileName = Split-Path $FilePath -Leaf
    $extension = [System.IO.Path]::GetExtension($fileName).ToLower()
    
    Write-Host "NEW FILE DETECTED: $fileName" -ForegroundColor Magenta
    
    if ($FileTypeMappings.ContainsKey($extension)) {
        $targetFolder = Join-Path $DownloadsPath $FileTypeMappings[$extension]
        Write-Host "File type: $extension -> $($FileTypeMappings[$extension])" -ForegroundColor Green
        Write-Host "Target: $targetFolder" -ForegroundColor Cyan
        
        Move-FileToFolder $FilePath $targetFolder
    } else {
        $miscFolder = Join-Path $DownloadsPath "Miscellaneous"
        Write-Host "Unknown file type: $extension" -ForegroundColor Yellow
        Write-Host "Moving to: $miscFolder" -ForegroundColor Cyan
        
        Move-FileToFolder $FilePath $miscFolder
    }
    
    Write-Host "----------------------------------------" -ForegroundColor Gray
}

# Create FileSystemWatcher
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $DownloadsPath
$watcher.Filter = "*.*"
$watcher.IncludeSubdirectories = $false
$watcher.EnableRaisingEvents = $true

$action = {
    $path = $Event.SourceEventArgs.FullPath
    $changeType = $Event.SourceEventArgs.ChangeType
    
    if ($changeType -eq 'Created') {
        $fileName = Split-Path $path -Leaf
        
        if ($fileName -like "*.part" -or $fileName -like "*.crdownload" -or $fileName -like "*.tmp") {
            Write-Host "Skipping temporary file: $fileName" -ForegroundColor Yellow
            return
        }
        
        Start-Sleep -Seconds 2
        Invoke-FileProcessing $path
    }
}

Register-ObjectEvent -InputObject $watcher -EventName "Created" -Action $action | Out-Null

Write-Host "File monitor is now active!" -ForegroundColor Green
Write-Host "Drop files into $DownloadsPath to see them organized automatically" -ForegroundColor White
Write-Host "Press Ctrl+C to stop monitoring" -ForegroundColor Yellow
Write-Host ""

try {
    while ($true) { Start-Sleep -Seconds 1 }
}
finally {
    $watcher.Dispose()
    Get-EventSubscriber | Unregister-Event
    Write-Host "File monitor stopped." -ForegroundColor Red
}
