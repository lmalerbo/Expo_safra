@echo off
title Atualizar Programacao Frentes
color 0B

echo.
echo  ================================================
echo   PROGRAMACAO FRENTES - Atualizacao da Base
echo  ================================================
echo.

cd /d "%~dp0.."

if exist "engine\atualizar_programacao.exe" (
    set RUNNER=engine\atualizar_programacao.exe
) else (
    python --version >nul 2>&1
    if errorlevel 1 (
        color 0C
        echo  ERRO: atualizar_programacao.exe nao encontrado e Python nao esta instalado.
        echo.
        pause
        exit /b 1
    )
    set RUNNER=python engine\atualizar_programacao.py
)

set xlsm_found=0
for %%f in (base_icol\*.xlsm) do set xlsm_found=1
if %xlsm_found%==0 (
    color 0C
    echo  ERRO: Nenhum arquivo .xlsm encontrado em base_icol\
    echo  Coloque a base ICOL dentro da pasta base_icol\
    echo.
    pause
    exit /b 1
)

if not exist "supabase_config.json" (
    color 0C
    echo  ERRO: supabase_config.json nao encontrado.
    echo  Crie esse arquivo com: { "url": "https://xxxx.supabase.co", "secret_key": "sb_secret_..." }
    echo.
    pause
    exit /b 1
)

echo  Iniciando atualizacao...
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
