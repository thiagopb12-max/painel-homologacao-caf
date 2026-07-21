@echo off
chcp 65001 >nul
echo ==========================================
echo   ATUALIZAR PAINEL DE HOMOLOGACAO CAF
echo ==========================================
echo.

cd /d "%~dp0"

echo [1/2] Registrando timestamp da atualizacao...
powershell -ExecutionPolicy Bypass -Command "$json = @{ data = (Get-Date -Format 'yyyy-MM-dd'); hora = (Get-Date -Format 'HH:mm:ss'); dataHoraCompleta = (Get-Date -Format 'dd/MM/yyyy HH:mm:ss') } | ConvertTo-Json; $json | Out-File -FilePath 'ultima-atualizacao.json' -Encoding UTF8"

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
    echo Edite o arquivo CAF.xlsx e tente novamente.
)

echo.
pause
