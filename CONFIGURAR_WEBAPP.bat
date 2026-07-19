@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"
title Configurar Web App - Festa Tudo de Helena

echo.
echo ============================================================
echo  CONFIGURAR GOOGLE APPS SCRIPT NO SITE
echo ============================================================
echo.
echo Cole abaixo a URL da implantacao do Google Apps Script.
echo Ela deve comecar com https://script.google.com/macros/s/
echo e terminar com /exec
echo.

set "WEBAPP_URL="
set /p "WEBAPP_URL=URL do Web App: "

if not defined WEBAPP_URL (
  echo.
  echo ERRO: nenhuma URL foi informada.
  pause
  exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Configurar-WebApp.ps1" -WebAppUrl "%WEBAPP_URL%"

if errorlevel 1 (
  echo.
  echo A configuracao nao foi concluida.
  pause
  exit /b 1
)

echo.
echo Agora abra index.html para testar o convite.
echo Abra organizador.html para informar data, horario, local e idades.
echo.
pause
exit /b 0
