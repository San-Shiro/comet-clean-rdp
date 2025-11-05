$targetUrl = "https://www.perplexity.ai/download-comet"
$outputFile = "$env:TEMP\comet_installer_latest.exe" # Assume the downloaded file is an EXE

$userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"

Write-Host "Initiating download with browser headers..."

try {
    # Added -UserAgent to simulate a legitimate browser request
    Invoke-WebRequest -Uri $targetUrl -OutFile $outputFile -MaximumRedirection 10 -ErrorAction Stop -UserAgent $userAgent
    
    if (-not (Test-Path $outputFile)) {
        throw "Download failed: Installer file was not found at $outputFile."
    }
    
    Write-Host "Installer downloaded successfully to $outputFile"
    
    # Execute the installer silently. 
    Write-Host "Executing silent installation of Comet..."
    
    Start-Process -FilePath $outputFile -ArgumentList "/S" -Wait -NoNewWindow -ErrorAction Stop
    
    Write-Host "Comet Browser installation completed without user intervention."
    
    # Clean up the downloaded installer trace
    Write-Host "Cleaning up the installer file."
    Remove-Item $outputFile -Force
    
} catch {
    # Detailed error output for debugging server responses
    Write-Error "Download or installation failure. Check if the URL is correct or if the server requires more headers."
    Write-Host "Error details: $($_.Exception.Message)"
    exit 1
}