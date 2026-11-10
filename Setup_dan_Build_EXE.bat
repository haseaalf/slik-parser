@echo off
setlocal EnableExtensions EnableDelayedExpansion
title SLIK Parser - Setup dan Build EXE
color 0F
cd /d "%~dp0"

echo ============================================================
echo   SLIK Parser - Setup dan Build EXE  (1 klik)
echo ============================================================
echo.
echo   Script ini akan:
echo     1. Menyiapkan Python 3.12 (dipasang otomatis bila belum ada)
echo     2. Memasang library yang dibutuhkan di folder sementara
echo     3. Membuat SLIK_Parser_GUI.exe di folder ini
echo     4. Membuat shortcut "SLIK Parser" di Desktop
echo.
echo   Butuh koneksi internet. Lama proses sekitar 3-10 menit.
echo   Jangan tutup jendela ini sampai muncul tulisan SELESAI.
echo.

REM Folder kerja build sengaja di TEMP (path pendek): venv di dalam folder
REM aplikasi bisa melewati batas 260 karakter path Windows bila folder
REM aplikasi sudah dalam (mis. hasil ekstrak ZIP GitHub / OneDrive), dan
REM itu membuat sebagian library gagal dimuat saat build.
set "WORK=%TEMP%\slik_parser_build"

REM ============================================================
REM  0. Pastikan semua file sumber ada (ZIP sudah diekstrak)
REM ============================================================
set "MISSING="
for %%F in (slik_parser.py gui_slik.py icon.ico logo.png requirements-build.txt) do (
    if not exist "%%F" set "MISSING=!MISSING! %%F"
)
if defined MISSING (
    echo [ERROR] File berikut tidak ditemukan di folder ini:!MISSING!
    echo         Ekstrak dulu SEMUA isi ZIP ke satu folder, lalu jalankan
    echo         script ini dari folder hasil ekstrak tersebut.
    goto FAIL
)

REM ============================================================
REM  1. Cari Python 3.12 asli (bukan stub Microsoft Store)
REM ============================================================
echo [1/4] Mencari Python 3.12 ...
set "PY="
call :TRY_PY py -3.12
if not defined PY call :TRY_PY "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not defined PY call :TRY_PY "%ProgramFiles%\Python312\python.exe"
if defined PY goto HAVE_PY

echo       Python 3.12 belum terpasang. Memasang otomatis lewat winget ...
where winget >nul 2>&1
if errorlevel 1 goto NO_WINGET
echo       (Setujui jika muncul jendela konfirmasi. Tidak perlu hak admin.)
winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements
call :TRY_PY "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not defined PY call :TRY_PY py -3.12
if defined PY goto HAVE_PY
echo [ERROR] Python 3.12 gagal dipasang otomatis.
goto MANUAL_PY

:NO_WINGET
echo [ERROR] winget tidak tersedia di laptop ini.

:MANUAL_PY
echo.
echo   Pasang Python 3.12 secara manual:
echo     1. Buka https://www.python.org/downloads/windows/
echo     2. Unduh "Windows installer (64-bit)" versi 3.12.x
echo     3. Saat instalasi, centang "Add python.exe to PATH"
echo     4. Jalankan lagi script ini
start "" "https://www.python.org/downloads/windows/"
goto FAIL

:HAVE_PY
echo [OK]  Python siap dipakai:
!PY! --version
echo.

REM ============================================================
REM  2. Virtual environment sementara + library terkunci versinya
REM ============================================================
echo [2/4] Memasang library (pdfplumber, openpyxl, tkinterdnd2, pyinstaller) ...
if exist "%WORK%" rmdir /s /q "%WORK%"
!PY! -m venv "%WORK%\venv"
if errorlevel 1 (
    echo [ERROR] Gagal membuat virtual environment Python.
    goto FAIL_CLEAN
)
set "VPY=%WORK%\venv\Scripts\python.exe"
"%VPY%" -m pip install --disable-pip-version-check --quiet --upgrade pip
"%VPY%" -m pip install --disable-pip-version-check -r requirements-build.txt
if errorlevel 1 (
    echo.
    echo [ERROR] Gagal mengunduh/memasang library.
    echo         Periksa koneksi internet. Jika memakai jaringan kantor
    echo         yang memblokir unduhan, coba gunakan jaringan lain.
    goto FAIL_CLEAN
)
echo [OK]  Library terpasang.
echo.

REM ============================================================
REM  3. Build exe dengan PyInstaller
REM ============================================================
echo [3/4] Membuat SLIK_Parser_GUI.exe (1-3 menit) ...
"%VPY%" -m PyInstaller --noconfirm --clean --onefile --windowed --log-level WARN ^
    --name SLIK_Parser_GUI ^
    --icon "%~dp0icon.ico" ^
    --add-data "%~dp0icon.ico;." ^
    --add-data "%~dp0logo.png;." ^
    --collect-all tkinterdnd2 ^
    --distpath "%WORK%\dist" --workpath "%WORK%\work" --specpath "%WORK%" ^
    "%~dp0gui_slik.py"
if not exist "%WORK%\dist\SLIK_Parser_GUI.exe" (
    echo [ERROR] Build gagal. Lihat pesan error di atas.
    goto FAIL_CLEAN
)
copy /y "%WORK%\dist\SLIK_Parser_GUI.exe" "%~dp0SLIK_Parser_GUI.exe" >nul
if errorlevel 1 (
    echo [ERROR] Tidak bisa menyalin SLIK_Parser_GUI.exe ke folder ini.
    echo         Tutup aplikasi SLIK Parser jika sedang terbuka, lalu jalankan lagi.
    goto FAIL_CLEAN
)
echo [OK]  SLIK_Parser_GUI.exe berhasil dibuat.
echo.

REM ============================================================
REM  4. Shortcut di Desktop + bersihkan file sementara
REM ============================================================
echo [4/4] Membuat shortcut di Desktop dan membersihkan file sementara ...
set "EXE_PATH=%~dp0SLIK_Parser_GUI.exe"
set "EXE_DIR=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$d=[Environment]::GetFolderPath('Desktop'); $s=(New-Object -ComObject WScript.Shell).CreateShortcut((Join-Path $d 'SLIK Parser.lnk')); $s.TargetPath=$env:EXE_PATH; $s.WorkingDirectory=$env:EXE_DIR; $s.IconLocation=$env:EXE_PATH + ',0'; $s.Description='SLIK Parser - konversi PDF SLIK ke Excel'; $s.Save()" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Shortcut Desktop tidak bisa dibuat. Jalankan langsung
    echo        SLIK_Parser_GUI.exe dari folder ini.
) else (
    echo [OK]  Shortcut "SLIK Parser" dibuat di Desktop.
)
call :CLEANUP

echo.
echo ============================================================
echo   SELESAI!
echo ============================================================
echo.
echo   Aplikasi : %~dp0SLIK_Parser_GUI.exe
echo   Shortcut : "SLIK Parser" di Desktop
echo.
echo   Mulai sekarang cukup buka aplikasi lewat shortcut di Desktop.
echo   Script ini TIDAK perlu dijalankan lagi.
echo   Catatan: jika folder ini dipindah, shortcut Desktop ikut rusak;
echo   jalankan script ini lagi dari lokasi baru.
echo.
explorer /select,"%~dp0SLIK_Parser_GUI.exe"
pause
exit /b 0

REM ============================================================
REM  Subroutine
REM ============================================================

:TRY_PY
REM Pakai kandidat ini jika benar-benar Python 3.12 lengkap dengan tkinter.
%* -c "import sys, tkinter, venv; sys.exit(0 if sys.version_info[:2] == (3, 12) else 1)" >nul 2>&1
if not errorlevel 1 set "PY=%*"
exit /b 0

:CLEANUP
if exist "%WORK%" rmdir /s /q "%WORK%"
if exist "%~dp0__pycache__" rmdir /s /q "%~dp0__pycache__"
exit /b 0

:FAIL_CLEAN
call :CLEANUP

:FAIL
echo.
echo   Build TIDAK selesai. Perbaiki masalah di atas lalu jalankan lagi.
echo.
pause
exit /b 1
