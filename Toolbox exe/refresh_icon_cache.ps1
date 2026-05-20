$ErrorActionPreference = "Stop"

Write-Host "Refreshing icon cache..."

try {
    & "$env:WINDIR\System32\ie4uinit.exe" -ClearIconCache | Out-Null
} catch {
    # Some Windows builds only respond to -show; keep going.
}

& "$env:WINDIR\System32\ie4uinit.exe" -show | Out-Null

Write-Host "Restarting Explorer..."
Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
Start-Process explorer.exe -WindowStyle Hidden

Write-Host "Done."
