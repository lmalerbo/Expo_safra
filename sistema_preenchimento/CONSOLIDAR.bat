@echo off
title Consolidar Exportacoes - Programacao Frentes
color 0A

echo.
echo  ================================================
echo   PROGRAMACAO FRENTES - Consolidacao de Exports
echo  ================================================
echo.

cd /d "%~dp0"

if exist "engine\consolidar.exe" (
    set RUNNER=engine\consolidar.exe
) else (
    python --version >nul 2>&1
    if errorlevel 1 (
        color 0C
        echo  ERRO: consolidar.exe nao encontrado e Python nao esta instalado.
        echo.
        pause
        exit /b 1
    )
    set RUNNER=python engine\consolidar.py
    echo  AVISO: consolidar.exe nao encontrado — usando Python diretamente.
    echo.
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

%RUNNER%

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
