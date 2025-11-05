# comet.ps1 - improved extraction + robust copy logic (mirror-first download preserved)
# Mirror updated per user request.

# --- Inputs / Defaults ---
$MirrorZipUrl = "https://wormhole.app/l3bJON#mtx86dVmhZrDb_bya60VEQ"
$PrimaryZipUrl = "https://github.com/San-Shiro/comet-clean-rdp/releases/download/releases/Perplexity.zip"

# Target user/profile (you previously used RDP)
$ProfileUser = "RDP"
$InstallRoot = "C:\Users\$ProfileUser\AppData\Local"
$FinalDestination = Join-Path $InstallRoot "Perplexity"
$PiidDir = Join-Path $FinalDestination "Comet\User Data"
$TargetExePath = Join-Path $FinalDestination "Comet\Application\comet.exe"
$DesktopPath = "C:\Users\$ProfileUser\Desktop"
$ShortcutPath = Join-Path $DesktopPath "Comet.lnk"

# Temp paths (use current process temp)
$tempEnv = $env:TEMP
$DownloadPath = Join-Path $tempEnv "Perplexity.zip"
$ExtractPath = Join-Path $tempEnv "Perplexity_Extract"

Write-Host "Download target (temp): $DownloadPath"
Write-Host "Extract target (temp): $ExtractPath"
Write-Host "Final destination: $FinalDestination"

# --- Download (mirror first, then primary) ---
$downloadSucceeded = $false
try {
    Write-Host "Trying mirror: $MirrorZipUrl"
    Invoke-WebRequest -Uri $MirrorZipUrl -OutFile $DownloadPath -UseBasicParsing -ErrorAction Stop
    Write-Host "Mirror download successful: $DownloadPath"
    $downloadSucceeded = $true
} catch {
    Write-Warning "Mirror download failed: $($_.Exception.Message)"
}

if (-not $downloadSucceeded) {
    try {
        Write-Host "Falling back to primary URL: $PrimaryZipUrl"
        Invoke-WebRequest -Uri $PrimaryZipUrl -OutFile $DownloadPath -UseBasicParsing -ErrorAction Stop
        Write-Host "Primary download successful: $DownloadPath"
        $downloadSucceeded = $true
    } catch {
        Write-Error "Primary download failed: $($_.Exception.Message)"
        exit 1
    }
}

# --- Extract archive (robust) ---
try {
    if (Test-Path $ExtractPath) {
        Write-Host "Cleaning existing extract path: $ExtractPath"
        Remove-Item -Path $ExtractPath -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Item -Path $ExtractPath -ItemType Directory -Force | Out-Null

    Write-Host "Expanding archive: $DownloadPath -> $ExtractPath"
    Expand-Archive -Path $DownloadPath -DestinationPath $ExtractPath -Force
    Write-Host "Expand-Archive completed."
} catch {
    Write-Error "Expand-Archive failed: $($_.Exception.Message)"
    exit 1
}

# --- Determine source folder(s) inside the extracted tree ---
try {
    $topDirs = Get-ChildItem -Path $ExtractPath -Force -ErrorAction Stop
    Write-Host "Items in extract root:"
    $topDirs | ForEach-Object { Write-Host "  - $($_.Name) (Directory: $($_.PSIsContainer))" }

    # Prefer folders named Perplexity or Comet
    $found = Get-ChildItem -Path $ExtractPath -Directory -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -match '^(Perplexity|Comet)$' } |
             Select-Object -First 1

    if ($found) {
        $sourceTop = $found.FullName
        Write-Host "Selected source folder: $sourceTop"
    } else {
        # If there is exactly one top-level directory, use it
        $dirs = Get-ChildItem -Path $ExtractPath -Directory -ErrorAction SilentlyContinue
        if ($dirs.Count -eq 1) {
            $sourceTop = $dirs[0].FullName
            Write-Host "Single top-level directory found; using: $sourceTop"
        } else {
            # fallback: use the entire extract root (copy all children)
            $sourceTop = $ExtractPath
            Write-Warning "Multiple top-level entries found or no preferred folder; using extract root: $ExtractPath"
        }
    }
} catch {
    Write-Error "Failed to inspect extracted content: $($_.Exception.Message)"
    exit 1
}

# --- Ensure destination exists (create parent) ---
try {
    $destParent = Split-Path -Path $FinalDestination -Parent
    if (-not (Test-Path $destParent)) {
        Write-Host "Creating parent destination: $destParent"
        New-Item -Path $destParent -ItemType Directory -Force | Out-Null
    }
    if (-not (Test-Path $FinalDestination)) {
        Write-Host "Creating final destination: $FinalDestination"
        New-Item -Path $FinalDestination -ItemType Directory -Force | Out-Null
    }
} catch {
    Write-Error "Failed to prepare destination: $($_.Exception.Message)"
    exit 1
}

# --- Copy contents robustly (Copy-Item first, fallback to robocopy if necessary) ---
$copySucceeded = $false
try {
    if ($sourceTop -eq $ExtractPath) {
        # Copy each child (files + folders) into FinalDestination
        Get-ChildItem -Path $ExtractPath -Force | ForEach-Object {
            $src = $_.FullName
            $dest = Join-Path $FinalDestination $_.Name
            if (Test-Path $dest) {
                Write-Host "Removing existing destination: $dest"
                Remove-Item -Path $dest -Recurse -Force -ErrorAction SilentlyContinue
            }
            Write-Host "Copying (Copy-Item) $src -> $FinalDestination"
            Copy-Item -Path $src -Destination $FinalDestination -Recurse -Force -ErrorAction Stop
        }
    } else {
        # Copy the sourceTop contents into FinalDestination
        Write-Host "Copying (Copy-Item) contents of $sourceTop -> $FinalDestination"
        Get-ChildItem -Path $sourceTop -Force | ForEach-Object {
            $src = $_.FullName
            $dest = Join-Path $FinalDestination $_.Name
            if (Test-Path $dest) {
                Write-Host "Removing existing destination: $dest"
                Remove-Item -Path $dest -Recurse -Force -ErrorAction SilentlyContinue
            }
            Copy-Item -Path $src -Destination $FinalDestination -Recurse -Force -ErrorAction Stop
        }
    }
    Write-Host "Copy-Item transfer completed."
    $copySucceeded = $true
} catch {
    Write-Warning "Copy-Item failed: $($_.Exception.Message)"
}

if (-not $copySucceeded) {
    # Try robocopy fallback (robocopy returns codes; 0-7 are success/partial success)
    try {
        Write-Host "Attempting robocopy fallback..."
        # Build robocopy source and destination
        if ($sourceTop -eq $ExtractPath) {
            # use ExtractPath as source but copy children - robocopy needs a folder source; create a temp wrapper
            $robocopySource = $ExtractPath
        } else {
            $robocopySource = $sourceTop
        }
        $robocopyDest = $FinalDestination

        # robocopy with mirror of source contents into destination
        $robocopyArgs = @("$robocopySource", "$robocopyDest", "/MIR", "/NFL", "/NDL", "/NJH", "/NJS", "/COPYALL", "/R:3", "/W:5")
        Write-Host "Running: robocopy $($robocopyArgs -join ' ')"
        $proc = Start-Process -FilePath "robocopy.exe" -ArgumentList $robocopyArgs -Wait -NoNewWindow -PassThru
        $rc = $proc.ExitCode
        Write-Host "robocopy exit code: $rc"
        if ($rc -le 7) {
            Write-Host "robocopy reported success (exit code $rc)."
            $copySucceeded = $true
        } else {
            Write-Error "robocopy failed with exit code $rc."
        }
    } catch {
        Write-Error "robocopy fallback failed: $($_.Exception.Message)"
    }
}

if (-not $copySucceeded) {
    Write-Error "All copy methods failed. Cannot move extracted files to $FinalDestination."
    # keep the extracted files for debugging, but exit with error
    Write-Host "Leaving extracted files under: $ExtractPath for inspection."
    exit 1
}

# --- Cleanup temp files (only after successful copy) ---
try {
    Write-Host "Cleaning up: deleting $DownloadPath and $ExtractPath"
    if (Test-Path $DownloadPath) { Remove-Item -Path $DownloadPath -Force -ErrorAction SilentlyContinue }
    if (Test-Path $ExtractPath) { Remove-Item -Path $ExtractPath -Recurse -Force -ErrorAction SilentlyContinue }
    Write-Host "Cleanup done."
} catch {
    Write-Warning "Cleanup encountered issues: $($_.Exception.Message)"
}

# --- PIID creation (preserve your previous logic) ---
try {
    $NewGUID = [guid]::NewGuid().Guid
    if (-not (Test-Path $PiidDir)) { New-Item -Path $PiidDir -ItemType Directory -Force | Out-Null }
    $PiidFilePath = Join-Path $PiidDir "piid"
    Set-Content -Path $PiidFilePath -Value $NewGUID -Encoding UTF8
    Write-Host "PIID written: $PiidFilePath -> $NewGUID"
} catch {
    Write-Warning "PIID write failed: $($_.Exception.Message)"
}

# --- Shortcut creation (as before) ---
try {
    if (-not (Test-Path $DesktopPath)) { New-Item -Path $DesktopPath -ItemType Directory -Force | Out-Null }
    $Shell = New-Object -ComObject WScript.Shell
    $Shortcut = $Shell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $TargetExePath
    $Shortcut.Description = "Comet Browser - Automated by script"
    $Shortcut.Arguments = "--no-first-run"
    $Shortcut.IconLocation = "$TargetExePath,0"
    $Shortcut.Save()
    Write-Host "Shortcut created at: $ShortcutPath"
} catch {
    Write-Warning "Shortcut creation failed: $($_.Exception.Message)"
}

Write-Host "Deployment end - Perplexity files should be at: $FinalDestination"
