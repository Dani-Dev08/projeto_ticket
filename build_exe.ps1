# Gera o executável TicketsTaxis.exe (um único arquivo, sem janela de terminal)
# Uso:  powershell -ExecutionPolicy Bypass -File build_exe.ps1

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

Write-Host "Limpando builds anteriores..." -ForegroundColor Cyan
foreach ($pasta in @("build", "dist")) {
    if (Test-Path $pasta) { Remove-Item $pasta -Recurse -Force }
}

Write-Host "Gerando o executavel (isso pode levar alguns minutos)..." -ForegroundColor Cyan
pyinstaller --noconfirm --clean TicketsTaxis.spec

$exe = Join-Path $PSScriptRoot "dist\TicketsTaxis.exe"
if (Test-Path $exe) {
    $tamanho = [math]::Round((Get-Item $exe).Length / 1MB, 1)
    Write-Host ""
    Write-Host "OK! Executavel gerado:" -ForegroundColor Green
    Write-Host "  $exe  ($tamanho MB)"
    Write-Host ""
    Write-Host "Copie apenas esse arquivo para o outro computador e execute."
} else {
    Write-Host "Falha na geracao do executavel." -ForegroundColor Red
    exit 1
}
