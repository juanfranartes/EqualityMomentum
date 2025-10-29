# Script para limpiar TODOS los archivos sensibles antes de commit
# Ejecutar ANTES de git add / git commit

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "LIMPIEZA DE ARCHIVOS SENSIBLES" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$carpetas = @(
    "01_DATOS_SIN_PROCESAR",
    "02_RESULTADOS",
    "03_LOGS",
    "05_INFORMES"
)

foreach ($carpeta in $carpetas) {
    if (Test-Path $carpeta) {
        Write-Host "Limpiando: $carpeta" -ForegroundColor White
        
        # Borrar archivos Excel
        $excel = Get-ChildItem -Path $carpeta -Filter "*.xlsx" -ErrorAction SilentlyContinue
        if ($excel) {
            Remove-Item -Path "$carpeta\*.xlsx" -Force -ErrorAction SilentlyContinue
            Write-Host "  Borrados $($excel.Count) archivos .xlsx" -ForegroundColor Green
        }
        
        # Borrar archivos Word
        $word = Get-ChildItem -Path $carpeta -Filter "*.docx" -ErrorAction SilentlyContinue
        if ($word) {
            Remove-Item -Path "$carpeta\*.docx" -Force -ErrorAction SilentlyContinue
            Write-Host "  Borrados $($word.Count) archivos .docx" -ForegroundColor Green
        }
        
        # Borrar logs
        $logs = Get-ChildItem -Path $carpeta -Filter "*.log" -ErrorAction SilentlyContinue
        if ($logs) {
            Remove-Item -Path "$carpeta\*.log" -Force -ErrorAction SilentlyContinue
            Write-Host "  Borrados $($logs.Count) archivos .log" -ForegroundColor Green
        }
    }
}

# Borrar archivos temporales
Write-Host ""
Write-Host "Limpiando archivos temporales..." -ForegroundColor White
Remove-Item -Path "temp_*.png" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "temp_*.docx" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "nul" -Force -ErrorAction SilentlyContinue
Write-Host "  Archivos temporales borrados" -ForegroundColor Green

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "LIMPIEZA COMPLETADA" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Ahora puedes hacer:" -ForegroundColor Yellow
Write-Host "  git add ." -ForegroundColor White
Write-Host "  git commit -m tu mensaje" -ForegroundColor White
Write-Host "  git push" -ForegroundColor White
Write-Host ""
