@echo off
chcp 65001 >nul
echo ==========================================
echo   ATUALIZAR PAINEL DE HOMOLOGACAO
echo   DTP.095850 - FPSE / SIBE / PAT / FI
echo ==========================================
echo.

cd /d "%~dp0"

echo [1/2] Registrando timestamp da atualizacao...
powershell -ExecutionPolicy Bypass -Command ^
    "$ts = @{ data = (Get-Date -Format 'yyyy-MM-dd'); hora = (Get-Date -Format 'HH:mm:ss'); dataHoraCompleta = (Get-Date -Format 'dd/MM/yyyy HH:mm:ss') } | ConvertTo-Json; ^
    $ts | Out-File -FilePath 'ultima-atualizacao.json' -Encoding UTF8; ^
    $ts | Out-File -FilePath 'fpse/ultima-atualizacao.json' -Encoding UTF8; ^
    if (Test-Path 'sibe') { $ts | Out-File -FilePath 'sibe/ultima-atualizacao.json' -Encoding UTF8 }; ^
    if (Test-Path 'pat') { $ts | Out-File -FilePath 'pat/ultima-atualizacao.json' -Encoding UTF8 }; ^
    if (Test-Path 'fluxo-integrado') { $ts | Out-File -FilePath 'fluxo-integrado/ultima-atualizacao.json' -Encoding UTF8 }"

echo.
echo [2/2] Enviando para o GitHub...
git add -A
git status --short | findstr /r "." >nul
if %errorlevel%==0 (
    git commit -m "Atualiza dados da homologacao - %date% %time:~0,5%"
    git push origin main
    echo.
    echo ==========================================
    echo   SUCESSO! Painel atualizado.
    echo   Aguarde 1-2 min para refletir online.
    echo ==========================================
) else (
    echo.
    echo Nenhuma alteracao detectada.
    echo Edite uma das planilhas e tente novamente.
    echo.
    echo Planilhas esperadas:
    echo   fpse\CAF.xlsx
    echo   sibe\SIBE.xlsx
    echo   pat\PAT.xlsx
    echo   fluxo-integrado\FLUXO.xlsx
)

echo.
pause
