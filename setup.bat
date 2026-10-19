@echo off
setlocal enabledelayedexpansion
title SLIK Parser - Setup Environment
color 0B

echo ============================================================
echo   SLIK OJK Parser - Setup Environment
echo ============================================================
echo.

REM ---- Pindah ke folder tempat file .bat ini berada ----
cd /d "%~dp0"

REM ============================================================
REM  1. Cek apakah Python sudah terpasang (dan BENAR-BENAR jalan)
REM     Catatan: "python.exe" di WindowsApps sering hanya stub
REM     Microsoft Store yang tidak menjalankan Python sungguhan.
REM     Karena itu setiap kandidat diuji dengan menjalankan kode.
REM ============================================================
set "PY="
call :TEST_PY py
if not errorlevel 1 ( set "PY=py"      & goto FOUND )
call :TEST_PY python
if not errorlevel 1 ( set "PY=python"  & goto FOUND )
call :TEST_PY python3
if not errorlevel 1 ( set "PY=python3" & goto FOUND )
goto NO_PYTHON

:FOUND
echo [OK] Python ditemukan: !PY!
!PY! --version
goto CHECK_VERSION

:NO_PYTHON

REM ============================================================
REM  2. Python belum ada -> coba install otomatis via winget
REM ============================================================
echo [INFO] Python belum terpasang di laptop ini.
echo.
where winget >nul 2>nul
if %errorlevel%==0 (
    echo [INFO] Mencoba memasang Python otomatis lewat winget...
    echo        Silakan setujui bila muncul jendela konfirmasi.
    echo.
    winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements
    echo.
    echo [PENTING] Instalasi selesai. Karena PATH baru diperbarui,
    echo           TUTUP jendela ini lalu jalankan setup.bat SEKALI LAGI.
    echo.
    pause
    exit /b 0
) else (
    echo [ERROR] winget tidak tersedia di laptop ini.
    echo.
    echo   Silakan pasang Python secara manual dari:
    echo   https://www.python.org/downloads/
    echo.
    echo   PENTING: saat instalasi, centang kotak
    echo   "Add Python to PATH", lalu jalankan setup.bat lagi.
    echo.
    start "" "https://www.python.org/downloads/"
    pause
    exit /b 1
)

REM ============================================================
REM  3. Verifikasi versi Python >= 3.8
REM ============================================================
:CHECK_VERSION
echo.
!PY! -c "import sys; assert sys.version_info >= (3,8), sys.version" 1>nul 2>nul
if errorlevel 1 (
    echo [ERROR] Versi Python terlalu lama. Butuh Python 3.8 atau lebih baru.
    echo         Silakan update Python dari https://www.python.org/downloads/
    pause
    exit /b 1
)

REM ============================================================
REM  4. Upgrade pip + install library yang dibutuhkan
REM ============================================================
echo.
echo [INFO] Memperbarui pip...
!PY! -m pip install --upgrade pip --quiet

echo [INFO] Memasang library: pdfplumber, openpyxl ...
!PY! -m pip install --upgrade pdfplumber openpyxl
if errorlevel 1 (
    echo.
    echo [ERROR] Gagal memasang library. Periksa koneksi internet lalu coba lagi.
    pause
    exit /b 1
)

REM ============================================================
REM  5. Selesai
REM ============================================================
echo.
echo ============================================================
echo   [SELESAI] Environment siap digunakan!
echo ============================================================
echo.
echo   Cara pakai:
echo     - Seret (drag) file PDF ke atas "run.bat", atau
echo     - Jalankan:  py slik_parser.py namafile.pdf
echo.
pause
exit /b 0

REM ============================================================
REM  Subroutine: uji apakah suatu perintah adalah Python asli.
REM  Menolak stub Microsoft Store (yang gagal menjalankan kode).
REM  errorlevel 0 = Python jalan, selain itu = gagal/bukan Python.
REM ============================================================
:TEST_PY
%~1 -c "import sys" >nul 2>nul
exit /b %errorlevel%
