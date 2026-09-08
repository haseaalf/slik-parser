@echo off
setlocal enabledelayedexpansion
title SLIK Parser - Build EXE
color 0B

cd /d "%~dp0"

echo ============================================================
echo   SLIK Parser - Build EXE (environment terisolasi)
echo ============================================================
echo.
echo   Script ini TIDAK memakai environment Python global kamu.
echo   Sebuah virtual environment baru (.buildenv) akan dibuat
echo   khusus untuk build, berisi HANYA library yang diperlukan:
echo     - pdfplumber, openpyxl, tkinterdnd2, pyinstaller
echo   Ini mencegah exe membengkak akibat ikut membawa library
echo   lain yang kebetulan terpasang di environment global.
echo.

REM ---- Pastikan file sumber ada di folder ini ----
if not exist "slik_parser.py" (
    echo [ERROR] slik_parser.py tidak ditemukan di folder ini.
    echo         Pastikan file ini ada di folder yang sama dengan build_exe.bat
    pause
    exit /b 1
)
if not exist "gui_slik.py" (
    echo [ERROR] gui_slik.py tidak ditemukan di folder ini.
    echo         Pastikan file ini ada di folder yang sama dengan build_exe.bat
    pause
    exit /b 1
)
if not exist "icon.ico" (
    echo [ERROR] icon.ico tidak ditemukan di folder ini.
    echo         Pastikan file ini ada di folder yang sama dengan build_exe.bat
    pause
    exit /b 1
)

REM ---- Cari perintah python asli (bukan stub Microsoft Store) ----
set "PY="
call :TEST_PY py
if not errorlevel 1 ( set "PY=py"      & goto FOUND_PY )
call :TEST_PY python
if not errorlevel 1 ( set "PY=python"  & goto FOUND_PY )
call :TEST_PY python3
if not errorlevel 1 ( set "PY=python3" & goto FOUND_PY )
echo [ERROR] Python belum terpasang / tidak berfungsi.
echo         Jalankan "setup.bat" terlebih dahulu, atau install dari python.org
pause
exit /b 1

:FOUND_PY
echo [OK] Python ditemukan: !PY!
echo.

REM ---- Buat virtual environment terisolasi khusus untuk build ----
if not exist ".buildenv" (
    echo [INFO] Membuat virtual environment terisolasi di .buildenv ...
    !PY! -m venv .buildenv
    if errorlevel 1 (
        echo [ERROR] Gagal membuat virtual environment.
        pause
        exit /b 1
    )
) else (
    echo [INFO] Virtual environment .buildenv sudah ada, dipakai ulang.
)

set "VENV_PY=.buildenv\Scripts\python.exe"

REM ---- Install HANYA library yang benar-benar dibutuhkan ----
echo.
echo [INFO] Memasang dependensi minimal di environment terisolasi...
"%VENV_PY%" -m pip install --upgrade pip --quiet
"%VENV_PY%" -m pip install --quiet pdfplumber openpyxl tkinterdnd2 pyinstaller
if errorlevel 1 (
    echo [ERROR] Gagal memasang dependensi. Periksa koneksi internet lalu coba lagi.
    pause
    exit /b 1
)
echo [OK] Dependensi terpasang.

REM ---- Bersihkan sisa build sebelumnya (kalau ada) ----
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "SLIK_Parser_GUI.spec" del /q "SLIK_Parser_GUI.spec"

REM ---- Build dengan PyInstaller dari environment terisolasi ----
echo.
echo [INFO] Menjalankan PyInstaller, mohon tunggu...
echo.
"%VENV_PY%" -m PyInstaller --onefile --windowed --name SLIK_Parser_GUI --icon icon.ico --add-data "icon.ico;." --collect-all tkinterdnd2 gui_slik.py

if not exist "dist\SLIK_Parser_GUI.exe" (
    echo.
    echo [ERROR] Build gagal. Lihat pesan error PyInstaller di atas.
    pause
    exit /b 1
)

REM ---- Pindahkan exe hasil build ke folder utama ----
copy /y "dist\SLIK_Parser_GUI.exe" "SLIK_Parser_GUI.exe" >nul

REM ---- Bersihkan folder/file sementara hasil build ----
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "__pycache__" rmdir /s /q "__pycache__"
if exist "SLIK_Parser_GUI.spec" del /q "SLIK_Parser_GUI.spec"

echo.
echo ============================================================
echo   [SELESAI] SLIK_Parser_GUI.exe berhasil dibuat!
echo ============================================================
echo.
echo   File exe : %~dp0SLIK_Parser_GUI.exe
echo.
echo   Catatan  : folder ".buildenv" (virtual environment) sengaja
echo              disimpan supaya build berikutnya lebih cepat -
echo              boleh dihapus manual kapan saja untuk menghemat
echo              ruang disk, tidak akan memengaruhi file .exe yang
echo              sudah jadi.
echo.
pause
exit /b 0

REM ---- Subroutine: uji apakah suatu perintah adalah Python asli ----
:TEST_PY
%~1 -c "import sys" >nul 2>nul
exit /b %errorlevel%
