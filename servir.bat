@echo off
echo ==========================================
echo   PAINEL DE HOMOLOGACAO CAF - DATAPREV
echo ==========================================
echo.
echo Iniciando servidor na porta 8080...
echo Acesse: http://localhost:8080
echo.
echo Para compartilhar na rede, use seu IP local:
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| find "IPv4"') do echo   http://%%a:8080
echo.
echo Pressione Ctrl+C para parar o servidor.
echo ==========================================
python -m http.server 8080
