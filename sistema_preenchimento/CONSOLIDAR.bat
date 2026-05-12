@echo off
title Consolidar Exportacoes - Programacao Frentes
color 0A

echo.
echo  ================================================
echo   PROGRAMACAO FRENTES - Consolidacao de Exports
echo  ================================================
echo.

cd /d "%~dp0"

if not exist "engine\consolidar.exe" (
    color 0C
    echo  ERRO: engine\consolidar.exe nao encontrado.
    echo.
    pause
    exit /b 1
)

if not exist "programacao_frentes.xlsx" (
    color 0C
    echo  ERRO: programacao_frentes.xlsx nao encontrado nesta pasta.
    echo.
    pause
    exit /b 1
)

if not exist "exports\" (
    mkdir exports
    echo  Pasta exports\ criada.
)

set count=0
for %%f in (exports\export_frentes_*.xlsx) do set /a count+=1

if %count%==0 (
    color 0E
    echo.
    echo  AVISO: Nenhum arquivo encontrado em exports\
    echo  Coloque os arquivos export_frentes_*.xlsx na pasta exports\
    echo.
    pause
    exit /b 0
)

echo  %count% arquivo(s) encontrado(s) em exports\
echo.
echo  ------------------------------------------------
echo.

engine\consolidar.exe

echo.
if errorlevel 1 (
    color 0C
    echo  Processo encerrado com erro.
) else (
    color 0A
    echo  Processo concluido com sucesso.
)

echo.
pause
