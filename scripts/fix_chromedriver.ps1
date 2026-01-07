# PowerShell script to fix ChromeDriver issues
# Run this if you get "[WinError 193]" or ChromeDriver errors

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "ChromeDriver Fix Script" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Close any running Chrome processes
Write-Host "Closing Chrome processes..." -ForegroundColor Yellow
Get-Process chrome -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

# Find and remove .wdm folder (ChromeDriver cache)
$wdmPath = "$env:USERPROFILE\.wdm"
if (Test-Path $wdmPath) {
    Write-Host "Found ChromeDriver cache at: $wdmPath" -ForegroundColor Yellow
    Write-Host "Removing cache..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force $wdmPath
    Write-Host "✓ ChromeDriver cache cleared!" -ForegroundColor Green
} else {
    Write-Host "No ChromeDriver cache found." -ForegroundColor Gray
}

# Remove local Chrome profile if it exists
$profilePath = ".\chrome_profile"
if (Test-Path $profilePath) {
    Write-Host "Found Chrome profile at: $profilePath" -ForegroundColor Yellow
    Write-Host "Removing profile..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force $profilePath
    Write-Host "✓ Chrome profile cleared!" -ForegroundColor Green
} else {
    Write-Host "No Chrome profile found." -ForegroundColor Gray
}

Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "1. Make sure Chrome browser is installed and up to date" -ForegroundColor White
Write-Host "2. Make sure all Chrome windows are closed" -ForegroundColor White
Write-Host "3. Run: python whatsapp_sender.py" -ForegroundColor White
Write-Host "4. ChromeDriver will be downloaded automatically" -ForegroundColor White
Write-Host ""
Write-Host "If issues persist:" -ForegroundColor Yellow
Write-Host "- Check if antivirus is blocking ChromeDriver" -ForegroundColor White
Write-Host "- Try restarting your computer" -ForegroundColor White
Write-Host "- Update Chrome to the latest version" -ForegroundColor White
Write-Host ""
