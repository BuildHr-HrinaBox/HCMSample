# Run after freeing disk space on C: drive
Set-Location $PSScriptRoot
git show "HEAD:app/src/Pages/Statutory.js" | Set-Content -Path "app/src/Pages/Statutory.js" -Encoding utf8
$lines = (Get-Content "app/src/Pages/Statutory.js").Count
Write-Host "Restored Statutory.js ($lines lines)"
Select-String -Path "app/src/Pages/Statutory.js" -Pattern "export default Statutory" | Select-Object -Last 1
