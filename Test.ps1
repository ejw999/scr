#Requires -Version 5.1

<#
.SYNOPSIS
    Organizes files in a specified directory based on file type.
.DESCRIPTION
    This script organizes files in a specified directory (defaulting to the Downloads folder)
    by moving them into subdirectories based on their file extension.
    It uses a predefined file type mapping to determine the destination directory for each file type.
.PARAMETER TargetDirectory
    Specifies the directory to be organized. Defaults to the current user's Downloads folder if not provided.
.EXAMPLE
    Organize-Files -TargetDirectory "C:\Users\YourName\Downloads"
    Organizes the files in the "C:\Users\YourName\Downloads" directory.
.EXAMPLE
    Organize-Files
    Organizes the files in the current user's Downloads folder.
.NOTES
    - Requires PowerShell 5.1 or later due to the use of the -File parameter with Get-ChildItem.
    - The script assumes that the destination directories already exist. You may need to create them manually if they don't.
    - The script handles errors gracefully by skipping files that cannot be moved and logging the errors.
#>
param (
    [string]$TargetDirectory = "$([Environment]::GetFolderPath('UserProfile'))\Downloads"
)

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

    # Videos (all -> MP4)
    '.mp4' = 'Videos\MP4'; '.m4v' = 'Videos\MP4'; '.avi' = 'Videos\MP4'; '.mkv' = 'Videos\MP4'
    '.mov' = 'Videos\MP4'; '.wmv' = 'Videos\MP4'; '.flv' = 'Videos\MP4'; '.webm' = 'Videos\MP4'
    '.3gp' = 'Videos\MP4'; '.ogv' = 'Videos\MP4'; '.m2ts' = 'Videos\MP4'; '.ts' = 'Videos\MP4'
    '.mts' = 'Videos\MP4'; '.vob' = 'Videos\MP4'; '.asf' = 'Videos\MP4'; '.divx' = 'Videos\MP4'; '.xvid' = 'Videos\MP4'

    # Audio
    '.mp3' = 'Audio\MP3'; '.wav' = 'Audio\WAV'; '.flac' = 'Audio\Other'; '.aac' = 'Audio\Other'; '.ogg' = 'Audio\Other'
    '.wma' = 'Audio\Other'; '.m4a' = 'Audio\Other'; '.opus' = 'Audio\Other'; '.ape' = 'Audio\Other'; '.ac3' = 'Audio\Other'
    '.dts' = 'Audio\Other'; '.amr' = 'Audio\Other'; '.3ga' = 'Audio\Other'

    # Archives (all -> ZIP folder)
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

    # Miscellaneous / Other
    '.torrent' = 'Miscellaneous'; '.log' = 'Miscellaneous'; '.md' = 'Miscellaneous'; '.url' = 'Miscellaneous'
    '.desktop' = 'Miscellaneous'; '.lnk' = 'Miscellaneous'; '.tmp' = 'Miscellaneous'; '.part' = 'Miscellaneous'
}

# Get all files in the target directory
try {
    $Files = Get-ChildItem -Path $TargetDirectory -File -ErrorAction Stop
}
catch {
    Write-Error "Error accessing directory: $($_.Exception.Message)"
    return
}

# Iterate through each file
foreach ($File in $Files) {
    # Get the file extension
    $Extension = $File.Extension

    # Check if the extension is in the mapping
    if ($FileTypeMappings.ContainsKey($Extension)) {
        # Get the destination path from the mapping
        $DestinationPath = Join-Path -Path $TargetDirectory -ChildPath $FileTypeMappings[$Extension]

        # Create the directory if it doesn't exist
        if (!(Test-Path -Path $DestinationPath -PathType Container)) {
            try {
                New-Item -ItemType Directory -Path $DestinationPath -ErrorAction Stop
                Write-Host "Created directory: $DestinationPath"
            }
            catch {
                Write-Error "Error creating directory $DestinationPath: $($_.Exception.Message)"
                continue # Skip to the next file
            }
        }

        # Move the file to the destination directory
        $DestinationFile = Join-Path -Path $DestinationPath -ChildPath $File.Name
        try {
            Move-Item -Path $File.FullName -Destination $DestinationFile -ErrorAction Stop
            Write-Host "Moved '$($File.Name)' to '$($FileTypeMappings[$Extension])'"
        }
        catch {
            Write-Error "Error moving file '$($File.Name)' to '$($FileTypeMappings[$Extension])': $($_.Exception.Message)"
        }
    }
    else {
        Write-Host "No mapping found for extension '$Extension' for file '$($File.Name)'"
    }
}

Write-Host "File organization complete."
