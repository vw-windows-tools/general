# ==========================================
# Nettoyage du cache OneNote
# ==========================================

Write-Host "Fermeture de OneNote..." -ForegroundColor Cyan

# Ferme OneNote s'il est lancé
Get-Process ONENOTE -ErrorAction SilentlyContinue | Stop-Process -Force
Get-Process OneNote -ErrorAction SilentlyContinue | Stop-Process -Force

Start-Sleep -Seconds 3

# Liste des emplacements de cache possibles
$cachePaths = @(
    "$env:LOCALAPPDATA\Microsoft\OneNote",
    "$env:LOCALAPPDATA\Packages\Microsoft.Office.OneNote_8wekyb3d8bbwe\LocalCache",
    "$env:LOCALAPPDATA\Packages\Microsoft.Office.OneNote_8wekyb3d8bbwe\AC"
)

foreach ($path in $cachePaths) {

    if (Test-Path $path) {

        Write-Host "Nettoyage : $path" -ForegroundColor Yellow

        try {
            Remove-Item "$path\*" -Recurse -Force -ErrorAction Stop
            Write-Host "OK : $path vidé." -ForegroundColor Green
        }
        catch {
            Write-Host "Erreur sur $path : $_" -ForegroundColor Red
        }

    }
    else {
        Write-Host "Chemin absent : $path" -ForegroundColor DarkGray
    }
}

Write-Host ""
Write-Host "Nettoyage terminé." -ForegroundColor Green
Write-Host "Relancez OneNote pour reconstruire le cache."