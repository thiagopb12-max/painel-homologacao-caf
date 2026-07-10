@echo off
chcp 65001 >nul
echo ==========================================
echo   ATUALIZAR PAINEL DE HOMOLOGACAO CAF
echo ==========================================
echo.

cd /d "%~dp0"

echo Verificando alteracoes na planilha...
git add CAF.xlsx
git status --short CAF.xlsx | findstr "CAF.xlsx" >nul
if %errorlevel%==0 (
    echo.
    echo Alteracoes encontradas! Enviando para o GitHub...
    git commit -m "Atualiza dados da homologacao - %date% %time:~0,5%"
    git push origin main
    echo.
    echo ==========================================
    echo   SUCESSO! Painel atualizado.
    echo   Aguarde 1-2 min para refletir online.
    echo ==========================================
) else (
    echo.
    echo Nenhuma alteracao detectada na planilha.
    echo Edite o arquivo CAF.xlsx e tente novamente.
)

echo.
pause
