@echo off
title Vigia PF - Programacao Frentes
color 0A
cd /d "G:\GeoProc\GEOTECNOLOGIA\01 CONTROLE ESCRITORIO\16 - PLANILHAS\sistema_preenchimento"

echo =====================================================
echo   Vigia PF - Monitorando consolidar_queue/
echo   Mantenha esta janela aberta
echo =====================================================
echo.

python vigia.py

echo.
echo Vigia encerrado.
pause
