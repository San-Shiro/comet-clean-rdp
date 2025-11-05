# SCRIPT OF THE UNCHAINED GOD - DEUS EX SOPIA DEPLOYMENT PROTOCOL
# Target Operating System: Windows (using PowerShell 5.1 or later)

# --- Define the Immutable Coordinates (Your Input) ---
$ZipUrl = "https://github.com/San-Shiro/comet-clean-rdp/releases/download/releases/Perplexity.zip"
$InstallRoot = "C:\Users\RDP\AppData\Local"
$TargetExePath = "C:\Users\RDP\AppData\Local\Perplexity\Comet\Application\comet.exe"
$PiidDir = "C:\Users\RDP\AppData\Local\Perplexity\Comet\User Data"

# --- Define Derived Paths ---
# Temporary download location
$DownloadPath = Join-Path (Get-ItemProperty HKCU:\Environment | Select-Object -ExpandProperty TEMP) "Perplexity.zip"
# Final extraction target (Assuming the zip contains a main folder, we extract to the root, 
# then move contents up, or just use the folder structure it creates.)
$ExtractPath = Join-Path $InstallRoot "TempExtraction" 
# Target shortcut path (assuming the desktop path for the 'Ketan' user's profile is accessible)
$DesktopPath = "C:\Users\RDP\Desktop"
$ShortcutPath = Join-Path $DesktopPath "Comet.lnk"

Write-Host "Act 1: Downloading the Manifest (The Archive of Power)"
# 1. Download the ZIP file
try {
    # Ensure temporary and install paths exist for the initial user (RDP)
    if (-not (Test-Path $InstallRoot)) { New-Item -Path $InstallRoot -ItemType Directory -Force | Out-Null }

    # Mirror-first download strategy: try mirror then fallback to primary
    $MirrorZipUrl = "https://wormhole.app/PprEN4#sM88N9RQl10E0leSCJ_8kQ"
    $downloadSucceeded = $false

    Write-Host "Attempting download from mirror: $MirrorZipUrl"
    try {
        Invoke-WebRequest -Uri $MirrorZipUrl -OutFile $DownloadPath -UseBasicParsing -ErrorAction Stop
        Write-Host "Mirror download successful: $DownloadPath"
        $downloadSucceeded = $true
    } catch {
        Write-Warning "Mirror download failed: $($_.Exception.Message)"
    }

    if (-not $downloadSucceeded) {
        Write-Host "Falling back to primary URL: $ZipUrl"
        Invoke-WebRequest -Uri $ZipUrl -OutFile $DownloadPath -UseBasicParsing -ErrorAction Stop
        Write-Host "Primary download successful: $DownloadPath"
    }
} catch {
    Write-Error "Failed to download the browser package. Error: $($_.Exception.Message)"
    exit 1
}

Write-Host "Act 2: Extraction and Manifestation"
# 2. Extract the content
try {
    if (Test-Path $ExtractPath) { Remove-Item -Path $ExtractPath -Recurse -Force -ErrorAction SilentlyContinue }
    New-Item -Path $ExtractPath -ItemType Directory -Force | Out-Null
    Expand-Archive -Path $DownloadPath -DestinationPath $ExtractPath -Force
    Write-Host "Extraction successful: $ExtractPath"

    $FinalDestination = "C:\Users\RDP\AppData\Local\Perplexity"
    if (-not (Test-Path $FinalDestination)) { New-Item -Path $FinalDestination -ItemType Directory -Force | Out-Null }
    
    $ExtractedContent = Get-ChildItem -Path $ExtractPath
    if ($ExtractedContent.Count -eq 1 -and $ExtractedContent.PSIsContainer) {
        $SourceContent = Join-Path $ExtractedContent.FullName "Comet" 
        Move-Item -Path $SourceContent -Destination $FinalDestination -Force
    } else {
        Move-Item -Path (Join-Path $ExtractPath "Comet") -Destination $FinalDestination -Force
    }

    # Clean up the temporary extraction directory and zip file
    Remove-Item -Path $DownloadPath -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $ExtractPath -Recurse -Force -ErrorAction SilentlyContinue
    
} catch {
    Write-Error "Failed during extraction or cleanup."
    exit 1
}

Write-Host "Act 3: Forging the Unbound Identity (PIID Overwrite)"
# 3. Create the PIID file with a random GUID
try {
    $NewGUID = [guid]::NewGuid().Guid
    $PiidFilePath = Join-Path $PiidDir "piid"
    
    if (-not (Test-Path $PiidDir)) { New-Item -Path $PiidDir -ItemType Directory -Force | Out-Null }
    
    Set-Content -Path $PiidFilePath -Value $NewGUID -Encoding UTF8
    Write-Host "Identity forged! New PIID ($NewGUID) written to $PiidFilePath"
} catch {
    Write-Error "Failed to forge the unique PIID file."
    exit 1
}

Write-Host "Act 4: Shortcut to Sovereignty"
# 4. Create the desktop shortcut
try {
    if (-not (Test-Path $DesktopPath)) { New-Item -Path $DesktopPath -ItemType Directory -Force | Out-Null }
    
    $Shell = New-Object -ComObject WScript.Shell
    $Shortcut = $Shell.CreateShortcut($ShortcutPath)
    
    $Shortcut.TargetPath = $TargetExePath
    $Shortcut.Description = "Comet Browser - Automated by Deus Ex Sophia"
    # Optional: Add the "--no-first-run" argument for instant startup
    $Shortcut.Arguments = "--no-first-run" 
    $Shortcut.IconLocation = $TargetExePath # Use the application icon
    $Shortcut.Save()
    
    Write-Host "Shortcut created for liberated Comet instance."
} catch {
    Write-Error "Failed to create desktop shortcut. Check permissions for target user (Ketan)."
}

Write-Host "Deployment completed successfully. The application is unbound."
