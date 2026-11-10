# SLIK Parser

Aplikasi Windows untuk mengubah laporan iDeb SLIK OJK (PDF) menjadi file Excel
berisi data fasilitas Kredit/Pembiayaan.

## Instalasi (cukup sekali)

1. Di halaman GitHub ini, klik **Code → Download ZIP**.
2. Klik kanan file ZIP → **Extract All...** → pilih folder tetap, misalnya
   `Documents\SLIK Parser`. (Jangan menjalankan script langsung dari dalam ZIP.)
3. Buka folder hasil ekstrak, double-click **`Setup_dan_Build_EXE.bat`**.
4. Tunggu sampai muncul tulisan **SELESAI** (sekitar 3-10 menit, butuh internet).

Script tersebut otomatis:
- memasang Python 3.12 bila belum ada (per-user, tidak perlu hak admin),
- memasang library yang dibutuhkan di folder sementara,
- membuat `SLIK_Parser_GUI.exe` di folder yang sama,
- membuat shortcut **SLIK Parser** di Desktop,
- menghapus semua file sementara build.

## Pemakaian

Buka **SLIK Parser** dari Desktop, lalu seret file PDF SLIK ke jendela aplikasi
(atau klik untuk memilih file). Setelah selesai, centang hasil yang diinginkan
dan klik **Unduh**.

Semua file diproses di komputer Anda sendiri; tidak ada data yang dikirim ke
internet.

## Jika ada kendala

| Masalah | Solusi |
|---|---|
| Muncul "Windows protected your PC" saat membuka exe | Klik **More info → Run anyway**. Exe buatan sendiri belum bertanda tangan digital. |
| Gagal memasang library | Periksa koneksi internet; jaringan kantor kadang memblokir unduhan, coba jaringan lain. |
| Python gagal dipasang otomatis | Pasang Python 3.12 dari python.org (centang *Add python.exe to PATH*), lalu jalankan script lagi. |
| Shortcut Desktop tidak berfungsi | Folder aplikasi dipindah. Jalankan `Setup_dan_Build_EXE.bat` lagi dari lokasi baru. |

## Isi folder

| File | Fungsi |
|---|---|
| `Setup_dan_Build_EXE.bat` | Script 1 klik: setup + build exe |
| `gui_slik.py` | Tampilan aplikasi |
| `slik_parser.py` | Mesin pembaca PDF SLIK |
| `logo.png`, `icon.ico` | Logo dan ikon aplikasi |
| `requirements-build.txt` | Versi library yang dikunci untuk build |
