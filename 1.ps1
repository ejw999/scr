# --- Update yt-dlp (safe check) ---
if (Test-Path $ytDlPath) {
    try {
        Write-Host "Checking for yt-dlp updates..." -ForegroundColor Cyan
        $updateProcess = Start-Process -FilePath $ytDlPath -ArgumentList "-U" -NoNewWindow -Wait -PassThru
        if ($updateProcess.ExitCode -eq 0) {
            Write-Host "yt-dlp is up to date." -ForegroundColor Green
        } else {
            Write-Host "yt-dlp update failed (exit code $($updateProcess.ExitCode)). Continuing anyway..." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Could not run yt-dlp update: $_. Continuing anyway..." -ForegroundColor Yellow
    }
} else {
    Write-Host "yt-dlp not found at path '$ytDlPath'. Skipping update check." -ForegroundColor Gray
}


# =============================
# YouTube to MP3 Downloader (auto cookies on 403, sanitized filenames, channel as artist)
# =============================

# --- Setup Directories ---
$baseDir      = "Z:\Pk\ytdl"
$ffmpegPath   = "Z:\Pk\ytdl\ffmpeg-release-essentials\ffmpeg-8.0-essentials_build\bin"
$outputFolder = "Z:\Pk\Music"
$cookiesFileDefault = Join-Path $baseDir "cookies.txt"
$ytDlPath     = Join-Path $baseDir "yt-dlp.exe"

# Create output folder if missing
if (-not (Test-Path $outputFolder)) {
    New-Item -Path $outputFolder -ItemType Directory | Out-Null
}

# Add yt-dlp folder to PATH (for current session)
$env:PATH = "$baseDir;$env:PATH"

# --- Configuration ---
$userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

# --- Helper Functions ---
function Convert-Filename {
    param([string]$filename)
    return ($filename -replace '[:\/\\*?"<>|]', '-')  # sanitize invalid chars
}

function Invoke-YouTubeDownload {
    param(
        [string]$url,
        [string]$cookiesFile = $null
    )

    $template = "%(uploader)s - %(title)s.%(ext)s"
    $outputTemplate = Join-Path $outputFolder (Convert-Filename $template)

    # yt-dlp arguments for MP3 extraction
    $argString = "-f bestaudio --extract-audio --audio-format mp3 --audio-quality 192K --embed-thumbnail --add-metadata"
    $argString += " --ffmpeg-location `"$ffmpegPath`" -o `"$outputTemplate`""
    $argString += " --user-agent `"$userAgent`" --geo-bypass"
    $argString += " --extractor-args `"youtube:player_client=default`""   # suppress JS runtime warning
    $argString += " --hls-prefer-native --force-keyframes-at-cuts --no-warnings"

    if ($cookiesFile -and (Test-Path $cookiesFile)) {
        $argString += " --cookies `"$cookiesFile`""
    }

    $argString += " `"$url`""

    try {
        $process = Start-Process -FilePath $ytDlPath -ArgumentList $argString -NoNewWindow -Wait -PassThru
        return $process.ExitCode
    } catch {
        Write-Host "Failed to run yt-dlp: $_" -ForegroundColor Red
        return 1
    }
}

# --- Main Execution ---
if (-not (Test-Path $ytDlPath)) {
    Write-Host "yt-dlp not found at '$ytDlPath'. Update the path in the script." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Main loop for downloading multiple songs
$continueDownloading = $true
while ($continueDownloading) {
    # Prompt for video URL
    $videoUrl = Read-Host "Enter the YouTube video URL (or 'exit' to quit)"
    $exitInput = $videoUrl.ToLower() -replace '\s'
    # Match exit variants (with typos), quit, or q
    if ($exitInput -match '^ex.*t$|^quit$|^q$') {
        Write-Host "Exiting." -ForegroundColor Yellow
        break
    }
    if ([string]::IsNullOrWhiteSpace($videoUrl)) {
        Write-Host "No URL provided. Try again." -ForegroundColor Yellow
        continue
    }

    Write-Host "Downloading: $videoUrl" -ForegroundColor Cyan

    # Attempt 1: Without cookies
    $exitCode = Invoke-YouTubeDownload -url $videoUrl

    # Attempt 2: Retry with cookies if first fails (403/age restriction)
    if ($exitCode -ne 0) {
        Write-Host "`nDownload failed, possibly age-restricted or region-blocked." -ForegroundColor Yellow
        if (Test-Path $cookiesFileDefault) {
            Write-Host "Retrying with default cookies..." -ForegroundColor Cyan
            $exitCode = Invoke-YouTubeDownload -url $videoUrl -cookiesFile $cookiesFileDefault
        } else {
            Write-Host "No cookies.txt found in $baseDir. Skipping retry." -ForegroundColor Red
        }
    }

    if ($exitCode -eq 0) {
        Write-Host "`nDownload complete!" -ForegroundColor Green
    } else {
        Write-Host "`nyt-dlp exited with code $exitCode" -ForegroundColor Yellow
        Write-Host "The MP3 may still be incomplete." -ForegroundColor Gray
    }

    Write-Host ""
}

exit 0
# =============================
# End of script
# =============================
Write-Host "" -ForegroundColor Gray
