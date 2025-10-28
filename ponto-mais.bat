@echo off
echo Iniciando processo 
echo.

for /f "usebackq tokens=1,2 delims==" %%a in (".env") do (
    if "%%a"=="CAMINHO_ARQUIVO" set CAMINHO_ARQUIVO=%%b
)

cd /d "%CAMINHO_ARQUIVO%"

echo Executando extracao.py...
python extracao/extracao.py
if %errorlevel% neq 0 (
    echo Erro ao executar extracao.py
    pause
    exit /b
)

echo Executando transformacao_turnos.py...
python transformacao/transformacao_turnos.py
if %errorlevel% neq 0 (
    echo Erro ao executar transformacao_turnos.py
    pause
    exit /b
)

echo Executando transformacao.py...
python transformacao/transformacao.py
if %errorlevel% neq 0 (
    echo Erro ao executar transformacao.py
    pause
    exit /b
)

echo.
echo Todos os scripts foram executados com sucesso!
pause
