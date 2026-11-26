@echo off
title Instalador - Sistema Empresarial
echo.
echo  ==========================================================
echo     INSTALADOR DEL SISTEMA EMPRESARIAL (Excel + VBA)
echo  ==========================================================
echo.
echo  Cierre Excel antes de continuar.
echo  Se creara la carpeta "SISTEMA EMPRESARIAL" en su escritorio.
echo.
pause
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0instalar.ps1"
echo.
pause
