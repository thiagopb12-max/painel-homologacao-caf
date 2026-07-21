@echo off
chcp 65001 >nul
echo ==========================================
echo   ATUALIZAR PAINEL DE HOMOLOGACAO CAF
echo ==========================================
echo.

cd /d "%~dp0"

echo [1/3] Gerando snapshot do dia...
powershell -ExecutionPolicy Bypass -File "%~dp0gerar-snapshot.ps1"
if %errorlevel% neq 0 (
    echo.
    echo AVISO: Nao foi possivel gerar snapshot. Verifique se o Excel esta fechado.
    echo Continuando com o push da planilha...
    echo.
)

echo.
echo [2/3] Preparando arquivos para envio...
git add CAF.xlsx historico.json ultima-atualizacao.json
git status --short | findstr /r "." >nul
if %errorlevel%==0 (
    echo.
    echo [3/3] Alteracoes encontradas! Enviando para o GitHub...
    git add -A
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
