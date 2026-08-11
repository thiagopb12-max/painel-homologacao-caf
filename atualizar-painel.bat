@echo off
chcp 65001 >nul
title Atualizar Painel de Homologacao

echo ==========================================
echo   ATUALIZAR PAINEL DE HOMOLOGACAO
echo   DTP.095850 - FPSE / SIBE / PAT / FI
echo ==========================================
echo.

cd /d "%~dp0"

echo [1/3] Gerando timestamp...
for /f "tokens=1-3 delims=/ " %%a in ("%date%") do set DATA=%%c-%%b-%%a
for /f "tokens=1-2 delims=:, " %%a in ("%time%") do set HORA=%%a:%%b
echo {"data":"%DATA%","hora":"%time:~0,8%","dataHoraCompleta":"%date% %time:~0,8%"} > ultima-atualizacao.json
echo {"data":"%DATA%","hora":"%time:~0,8%","dataHoraCompleta":"%date% %time:~0,8%"} > fpse\ultima-atualizacao.json
if exist sibe echo {"data":"%DATA%","hora":"%time:~0,8%","dataHoraCompleta":"%date% %time:~0,8%"} > sibe\ultima-atualizacao.json
if exist pat echo {"data":"%DATA%","hora":"%time:~0,8%","dataHoraCompleta":"%date% %time:~0,8%"} > pat\ultima-atualizacao.json
if exist fluxo-integrado echo {"data":"%DATA%","hora":"%time:~0,8%","dataHoraCompleta":"%date% %time:~0,8%"} > fluxo-integrado\ultima-atualizacao.json
echo Timestamp gerado: %date% %time:~0,8%

echo.
echo [2/3] Preparando arquivos...
git add -A
git status --short

echo.
echo [3/3] Enviando para o GitHub...
git commit -m "Atualiza dados - %date% %time:~0,5%"
git push origin main
echo.
echo ==========================================
echo   SUCESSO! Painel atualizado no GitHub.
echo   O Netlify publicara em ~30 segundos.
echo ==========================================

echo.
pause
