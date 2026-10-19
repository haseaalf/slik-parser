@echo off
setlocal enabledelayedexpansion
title SLIK Parser
color 0B

cd /d "%~dp0"

REM ---- Cari perintah python yang BENAR-BENAR jalan (bukan stub Store) ----
set "PY="
call :TEST_PY py       && set "PY=py"
if not defined PY call :TEST_PY python  && set "PY=python"
if not defined PY call :TEST_PY python3 && set "PY=python3"
if not defined PY (
    echo [ERROR] Python belum terpasang / tidak berfungsi.
    echo         Jalankan "setup.bat" terlebih dahulu.
    echo.
    pause
    exit /b 1
)

REM ---- Pastikan ada file PDF yang diseret ----
if "%~1"=="" (
    echo ============================================================
    echo   SLIK OJK Parser
    echo ============================================================
    echo.
    echo   Cara pakai: SERET file PDF ke atas ikon run.bat ini.
    echo.
    echo   Atau ketik manual:
    echo     run.bat "C:\path\ke\file.pdf"
    echo.
    pause
    exit /b 1
)

REM ---- Proses setiap file PDF yang diseret ----
:LOOP
if "%~1"=="" goto DONE
echo.
echo Memproses: %~1
!PY! "%~dp0slik_parser.py" "%~1"
shift
goto LOOP

:DONE
echo.
pause
exit /b 0

REM ---- Subroutine: uji Python asli (menolak stub Microsoft Store) ----
:TEST_PY
%~1 -c "import sys" >nul 2>nul
exit /b %errorlevel%
