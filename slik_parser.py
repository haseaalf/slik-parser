"""
SLIK OJK PDF Parser
====================
Mengekstrak data fasilitas kredit dari laporan SLIK OJK (format PDF teks digital).

Instalasi dependensi:
    pip install pdfplumber openpyxl

Penggunaan:
    python slik_parser.py <path_ke_file_pdf>

Contoh:
    python slik_parser.py slik_nasabah.pdf

Output:
    <nama_file_input>.xlsx (nama sama dengan PDF input, di direktori yang sama)

Requirements:
    Python >= 3.8
    pdfplumber >= 0.9
    openpyxl >= 3.0
"""

import sys
import re
import os

try:
    import pdfplumber
except ImportError:
    print("ERROR: Library 'pdfplumber' tidak ditemukan. Install dengan: pip install pdfplumber")
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERROR: Library 'openpyxl' tidak ditemukan. Install dengan: pip install openpyxl")
    sys.exit(1)


# ─────────────────────────── Helpers ────────────────────────────

BULAN_MAP = {
    "januari": "01", "februari": "02", "maret": "03", "april": "04",
    "mei": "05", "juni": "06", "juli": "07", "agustus": "08",
    "september": "09", "oktober": "10", "november": "11", "desember": "12",
}

# Pola baris kalender (header histori kolektibilitas)
_RE_CALENDAR = re.compile(
    r"^(?:Apr|Mei|Jun|Jul|Agt|Sep|Okt|Nov|Des|Jan|Feb|Mar)\s+\d{2}",
    re.IGNORECASE
)

# Suffix hukum badan usaha yang menandai akhir nama pelapor
_RE_LEGAL_SUFFIX = re.compile(
    r"\b(Tbk\.?|Indonesia|Finance|Multiartha|Teknologi|Tech|Fintek|Mikro|Artos|Sejahtera|Lestari|Nusantara|Bersama|Utama|Mandiri|Makmur)\b"
)

# Pola teks cabang yang selalu muncul setelah nama pelapor di baris header.
# Format: "NNN - <NAMA_PELAPOR> <CABANG> Rp xxx"
# PENTING: "BANK " (huruf kapital semua) = marker cabang.
#          "Bank " (huruf campuran) = bagian nama perusahaan → JANGAN dicocokkan.
# Cabang selalu dimulai dengan salah satu marker berikut:
_RE_CABANG_START = re.compile(
    r"(?:^|\s)"
    r"(?:"
    r"Kantor\s+Pusat"        # "Kantor Pusat"
    r"|KC\s"                  # "KC " (kode cabang)
    r"|KCP\s"                 # "KCP "
    r"|KPO[\s\-]"            # "KPO " atau "KPO-"
    r"|BANK\s+[A-Z]"         # "BANK " all-caps (bukan "Bank" dalam nama pelapor)
    r"|BRI\s+"                # "BRI "
    r"|BNI\s+"                # "BNI "
    r"|BTN\s+"                # "BTN "
    r"|BTPN\s+"               # "BTPN "
    r"|MANDIRI\s+KC"          # "MANDIRI KC"
    r"|MEGA\s+KC"             # "MEGA KC"
    r"|MEGA\s+KPO"            # "MEGA KPO"
    r")"
)


def clean_cid(text):
    """Hapus artefak (cid:N) yang kadang muncul di PDF tertentu."""
    return re.sub(r'\(cid:\d+\)', ' ', text)


def parse_tanggal(text):
    """Ubah 'DD Bulan YYYY' → 'DD/MM/YYYY'. Kembalikan string asli jika gagal."""
    if not text or not text.strip():
        return ""
    t = text.strip().lower()
    for nama, num in BULAN_MAP.items():
        if nama in t:
            t = t.replace(nama, num)
            break
    parts = t.split()
    if len(parts) == 3:
        try:
            return "{:02d}/{}/{}".format(int(parts[0]), parts[1], parts[2])
        except Exception:
            pass
    return text.strip()


def parse_rupiah(text):
    """Ubah 'Rp 1.234.567,00' → 1234567. Kembalikan 0 jika kosong/gagal."""
    if not text or not text.strip():
        return 0
    s = re.sub(r"[Rp\s]", "", text.strip())
    s = s.replace(".", "").split(",")[0]
    try:
        return int(s)
    except Exception:
        return 0


def parse_persen(text):
    """Ubah '21 %' → '21%', '21.5 %' → '21.5%'."""
    if not text or not text.strip():
        return ""
    s = text.strip().replace(" ", "")
    if not s.endswith("%"):
        s += "%"
    return s


# ─────────────────────── Identitas Nasabah ───────────────────────

def extract_identitas(pdf):
    """Ekstrak identitas nasabah dari halaman pertama PDF."""
    text = clean_cid(pdf.pages[0].extract_text() or "")

    # Nama: muncul tepat setelah "Nama Jenis Kelamin <nomor>\n"
    nama = ""
    m = re.search(
        r"Nama\s+Jenis Kelamin\s+\S+\s*\n([A-Z][A-Z\s]+?)\s+(?:PEREMPUAN|LAKI-LAKI|LAKI)",
        text
    )
    if m:
        nama = m.group(1).strip()
    if not nama:
        m = re.search(
            r"Nama Sesuai Identitas.*?\n([A-Z][A-Z\s]+?)\s+NIK\s*/",
            text, re.DOTALL
        )
        if m:
            nama = m.group(1).strip()

    # NIK: 16 digit pertama
    nik = ""
    m = re.search(r"\b(\d{16})\b", text)
    if m:
        nik = m.group(1)

    # Nomor Referensi (Kode Ref. Pengguna)
    nomor_ref = ""
    m = re.search(r"Kode Ref\.\s*Pengguna\s*\n?\s*([\w.\-/]+)", text)
    if m:
        nomor_ref = m.group(1).strip()

    # Nomor Laporan (fallback)
    nomor_laporan = ""
    m = re.search(r"Nomor\s+Laporan\s*\n?\s*([\d/IDEB]+)", text)
    if m:
        nomor_laporan = m.group(1).strip()

    return {
        "nama": nama,
        "nik": nik,
        "nomor_ref": nomor_ref,
        "nomor_laporan": nomor_laporan,
    }


# ─────────────────── Ringkasan Fasilitas ────────────────────────

def extract_total_baki_debet(pdf):
    """
    Ekstrak Total Baki Debet Kredit/Pembiayaan dari section Ringkasan Fasilitas.

    Kolom Ringkasan Fasilitas bervariasi tergantung tipe debitur:
    - 5 kolom: Kredit/Pembiayaan | L/C | Garansi | Fasilitas Lain | Total
    - 6 kolom: Kredit/Pembiayaan | Surat Berharga | L/C | Garansi | Fasilitas Lain | Total

    Kolom pertama (index 1) selalu Kredit/Pembiayaan — ini yang kita ambil.
    Kolom terakhir adalah Total — berguna sebagai fallback.
    """
    for page in pdf.pages:
        text = clean_cid(page.extract_text() or "")
        if "Ringkasan Fasilitas" not in text:
            continue
        if "Baki Debet" not in text:
            continue

        # Coba ambil semua angka setelah "Baki Debet"
        m = re.search(r"Baki Debet\s+((\d[\d.,]*(?:\s+\d[\d.,]*)*))", text)
        if m:
            nums = re.findall(r"[\d.,]+", m.group(1))
            if nums:
                # Kolom pertama = Kredit/Pembiayaan
                kredit_val = parse_rupiah("Rp " + nums[0])
                return kredit_val
    return 0


# ─────────────────── Kol Terendah dari Histori ───────────────────

def extract_kol_terendah(text):
    """
    Cari nilai kolektibilitas terendah (1-5) dari tabel histori
    'Kualitas / Jumlah Hari Tunggakan'.

    Format tabel: setiap bulan ada 2 angka berpasangan → (kol, hari_tunggakan).
    Kita hanya ambil angka di posisi GANJIL (kol), skip posisi GENAP (hari tunggakan).

    Kembalikan None jika histori kosong (fasilitas sudah melewati window 2 tahun).
    """
    # Cari blok histori antara header "Kualitas / Jumlah Hari Tunggakan" dan "No Rekening"
    # Variasi layout header:
    #   (1) "Kualitas / Jumlah Hari\nTunggakan"  → header 2 baris versi A
    #   (2) "Kualitas / Jumlah\nHari Tunggakan"  → header 2 baris versi B
    #   (3) "Kualitas / Jumlah Hari Tunggakan"    → header 1 baris
    # Angka histori bisa langsung setelah header atau di baris berikutnya
    m = re.search(
        r"Kualitas\s*/\s*Jumlah(?:\s+Hari)?\s*\nHari\s+Tunggakan(.*?)No\s+Rekening",
        text,
        re.DOTALL | re.IGNORECASE,
    )
    if not m:
        m = re.search(
            r"Kualitas\s*/\s*Jumlah\s+Hari\s*\nTunggakan(.*?)No\s+Rekening",
            text,
            re.DOTALL | re.IGNORECASE,
        )
    if not m:
        # Pola di mana angka histori langsung di baris yang sama dengan "Tunggakan"
        # mis: "Kualitas / Jumlah 1 0 1 0 ...\nHari Tunggakan\n..."
        m = re.search(
            r"Kualitas\s*/\s*Jumlah(\s+\d.*?)\nHari\s+Tunggakan(.*?)No\s+Rekening",
            text,
            re.DOTALL | re.IGNORECASE,
        )
        if m:
            # Gabungkan grup 1 (angka setelah "Jumlah") dan grup 2 (setelah "Tunggakan")
            combined = m.group(1) + "\n" + m.group(2)
            m = type("M", (), {"group": lambda self, n: combined if n == 1 else ""})()
    if not m:
        # Fallback: apapun setelah kata "Tunggakan" sampai "No Rekening"
        m = re.search(
            r"Tunggakan\s*(.*?)No\s+Rekening",
            text,
            re.DOTALL | re.IGNORECASE,
        )

    if not m:
        return None

    block = m.group(1).strip()

    # Proses SETIAP BARIS angka secara INDEPENDEN.
    # Dalam setiap baris: index 0,2,4,... = kolektibilitas, index 1,3,5,... = hari tunggakan.
    # Jangan gabung semua baris dulu — offset bisa bergeser jika jumlah bulan tiap baris ganjil.
    any_data = False
    kol_values = []

    for line in block.split("\n"):
        line = line.strip()
        if not line:
            continue
        if _RE_CALENDAR.match(line):
            continue
        row_nums = []
        for t in line.split():
            try:
                row_nums.append(int(t))
            except ValueError:
                pass
        if not row_nums:
            continue
        any_data = True
        # Index 0,2,4,... dari baris ini = kolektibilitas
        for i in range(0, len(row_nums), 2):
            v = row_nums[i]
            if 1 <= v <= 5:
                kol_values.append(v)

    if not any_data:
        return None  # histori kosong

    if not kol_values:
        return None

    return max(kol_values)  # Kol terendah = kualitas TERBURUK = angka tertinggi (Kol 5=macet)


# ─────────────────── Pelapor Extraction ─────────────────────────

def _cut_at_last_legal_suffix(s):
    """Potong string tepat di akhir suffix hukum terakhir."""
    best = -1
    for mm in _RE_LEGAL_SUFFIX.finditer(s):
        if mm.end() > best:
            best = mm.end()
    return s[:best].strip() if best > 0 else s.strip()


def _strip_cabang(name):
    """
    Hapus teks cabang dari nama pelapor.
    Cabang selalu dimulai dengan marker tertentu: Kantor Pusat, KC, KCP, KPO,
    BANK XX, BRI, BNI, atau singkatan bank dalam huruf kapital diikuti KC/KCP/KPO.
    Jika tidak ada marker cabang, fallback ke suffix hukum.
    """
    cabang_m = _RE_CABANG_START.search(name)
    if cabang_m:
        cut = cabang_m.start()
        # Jika match dimulai dengan spasi (bukan awal string), potong di spasi itu
        return name[:cut].strip()
    # Fallback: potong di suffix hukum terakhir
    return _cut_at_last_legal_suffix(name)


def extract_pelapor(lines):
    """
    Ekstrak nama pelapor dari baris-baris halaman fasilitas.

    Format baris header:
      "NNN - <NAMA_PELAPOR> <CABANG> Rp xxx DD Bln YYYY"
    Baris berikutnya bisa berisi lanjutan nama (mis. "Tbk") atau baris kalender.

    Strategi utama: potong teks cabang menggunakan _RE_CABANG_START,
    yang mendeteksi marker cabang (Kantor Pusat, KC, BRI KC, BANK XX, dll).
    """
    bank_line = ""
    bank_line_idx = -1
    for idx, line in enumerate(lines):
        if re.match(r"\s*\d+\s+-\s+", line):
            bank_line = line
            bank_line_idx = idx
            break

    if not bank_line:
        return ""

    # Ambil teks antara "NNN - " dan " Rp "
    m = re.match(r"\s*\d+\s+-\s+(.+?)\s+Rp\s+", bank_line)
    if not m:
        return ""
    full = m.group(1).strip()

    # Periksa next line: lanjutan nama atau kalender?
    next_line = ""
    if bank_line_idx + 1 < len(lines):
        nl = lines[bank_line_idx + 1].strip()
        # next_line valid jika: bukan kalender, bukan baris bank lain, tidak kosong
        if nl and not _RE_CALENDAR.match(nl) and not re.match(r"\d+\s+-\s+", nl):
            next_line = nl

    # ── Kasus d/h (dahulu/dba) ──
    if " d/h " in full.lower():
        dh_idx = full.lower().index(" d/h ")
        before = full[:dh_idx]      # nama baru
        after  = full[dh_idx + 5:]  # nama lama + cabang

        # Potong cabang dari "after"
        cabang_m = re.search(
            r"(?:^|\s)(?:KC\b|KCP\b|KPO\b|Kantor\s+Pusat|BANK\s+[A-Z]|BRI\s+|BNI\s+)",
            after, re.IGNORECASE
        )
        old_raw = after[:cabang_m.start()].strip() if cabang_m else after.strip()

        # Jika old_raw kosong atau tidak punya suffix, tambahkan next_line
        if next_line and (not old_raw or not _RE_LEGAL_SUFFIX.search(old_raw)):
            old_raw = (old_raw + " " + next_line).strip()

        nama_baru = _strip_cabang(before)
        return "{} d/h {}".format(nama_baru, old_raw.strip()).strip()

    # ── Kasus normal ──
    # Potong bagian cabang dari full
    pelapor = _strip_cabang(full)

    # Tambahkan next_line jika merupakan lanjutan nama (bukan cabang, bukan kalender)
    if next_line:
        is_cabang = bool(re.search(
            r"(?:Kantor\s+Pusat|KC\b|KCP\b|KPO\b|BANK\s+[A-Z]|BRI\s+|BNI\s+)",
            next_line, re.IGNORECASE
        ))
        if not is_cabang:
            pelapor = (pelapor + " " + next_line).strip()

    # Potong ulang jika next_line membawa serta teks cabang
    pelapor = _strip_cabang(pelapor)

    # Hilangkan duplikasi segmen (mis. "Auto Finance Auto Finance")
    words = pelapor.split()
    n = len(words)
    for half in range(n // 2, 1, -1):
        if words[n - half:] == words[n - 2 * half:n - half]:
            pelapor = " ".join(words[:n - half])
            break

    return pelapor.strip()


# ─────────────────── Parse Satu Halaman Fasilitas ────────────────

def parse_facility_page(raw_text):
    """Parse satu halaman Kredit/Pembiayaan → dict data fasilitas."""
    text = clean_cid(raw_text)
    lines = text.split("\n")

    # ── Pelapor ──
    pelapor = extract_pelapor(lines)

    # ── Baki Debet (dari baris header atas halaman) ──
    # Format: "NNN - Nama Bank   CABANG   Rp xxx,xx   DD Bln YYYY"
    baki_debet = 0
    m = re.search(r"Rp\s+([\d.,]+)\s+\d{2}\s+\w+\s+\d{4}", text)
    if m:
        baki_debet = parse_rupiah("Rp " + m.group(1))

    # ── Kualitas saat ini ──
    kol = ""
    m = re.search(r"Kualitas\s+(\d)\s*-", text)
    if m:
        kol = m.group(1)

    # ── Jenis Penggunaan ──
    # Dua-kolom: "Jenis Penggunaan  <nilai>  Frekuensi Restrukturisasi  <n>"
    jenis = ""
    m = re.search(r"Jenis Penggunaan\s+(.+?)\s+Frekuensi Restrukturisasi", text, re.DOTALL)
    if m:
        jenis = m.group(1).strip()

    # ── Tanggal Mulai (Realisasi) ──
    realisasi = ""
    m = re.search(r"Tanggal Mulai\s+(.+?)\s+Tunggakan\s+Bunga", text, re.DOTALL)
    if m:
        realisasi = parse_tanggal(m.group(1).strip())

    # ── Tanggal Jatuh Tempo ──
    jatuh_tempo = ""
    m = re.search(r"Tanggal Jatuh Tempo\s+(.+?)\s+Frekuensi Tunggakan", text, re.DOTALL)
    if m:
        jatuh_tempo = parse_tanggal(m.group(1).strip())

    # ── Plafon Awal ──
    plafon_awal = 0
    m = re.search(r"Plafon Awal\s+Rp\s+([\d.,]+)", text)
    if m:
        plafon_awal = parse_rupiah("Rp " + m.group(1))

    # ── Plafon (bukan Plafon Awal) ──
    # Ambil kemunculan "Plafon Rp" yang tidak didahului "Awal"
    plafon = 0
    for mo in re.finditer(r"Plafon\s+Rp\s+([\d.,]+)", text):
        prefix = text[max(0, mo.start() - 10):mo.start()]
        if "Awal" not in prefix:
            plafon = parse_rupiah("Rp " + mo.group(1))
            break

    # ── Frekuensi Restrukturisasi → Flag Restruk ──
    freq_restruk = 0
    m = re.search(r"Frekuensi Restrukturisasi\s+(\d+)", text)
    if m:
        try:
            freq_restruk = int(m.group(1))
        except ValueError:
            freq_restruk = 0
    flag_restruk = "YES" if freq_restruk != 0 else "NO"

    # ── Suku Bunga / Imbalan ──
    rate = ""
    m = re.search(r"Suku Bunga/Imbalan\s+([\d.]+)\s*%?", text)
    if m:
        rate = parse_persen(m.group(1))

    # ── Kol Terendah ──
    kol_terendah = extract_kol_terendah(text)
    # None → NULL (disimpan sebagai None, di Excel jadi sel kosong berlabel NULL)

    # ── Kondisi ──
    kondisi = ""
    m = re.search(r"Kab/Kota Lokasi Proyek\s+.+?\s+Kondisi\s+(.+?)(?:\s{2,}|\n|$)", text)
    if m:
        kondisi = m.group(1).strip()
    if not kondisi:
        # Fallback langsung cari "Kondisi <nilai>"
        m = re.search(r"\bKondisi\s+(Fasilitas Aktif|Lunas|Hapus Buku|Hapus Tagih|Write Off)", text)
        if m:
            kondisi = m.group(1).strip()

    return {
        "pelapor": pelapor,
        "jenis": jenis,
        "realisasi": realisasi,
        "jatuh_tempo": jatuh_tempo,
        "plafon_awal": plafon_awal,
        "plafon": plafon,
        "baki_debet": baki_debet,
        "kol": kol,
        "flag_restruk": flag_restruk,
        "rate": rate,
        "kol_terendah": kol_terendah,  # int, atau None jika kosong
        "kondisi": kondisi,
    }


# ────────────────────── Main Parser ─────────────────────────────

def parse_slik(pdf_path):
    print("\n" + "=" * 60)
    print("  SLIK OJK Parser")
    print("=" * 60)
    print("  File: {}".format(pdf_path))

    if not os.path.exists(pdf_path):
        print("\nERROR: File tidak ditemukan: {}".format(pdf_path))
        sys.exit(1)

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        print("  Total halaman PDF: {}".format(total_pages))

        # 1. Identitas Nasabah
        identitas = extract_identitas(pdf)
        has_identitas = bool(identitas["nama"] or identitas["nik"])

        if has_identitas:
            print("\n" + "-" * 60)
            print("  IDENTITAS NASABAH")
            print("-" * 60)
            print("  Nama          : {}".format(identitas["nama"]))
            print("  NIK           : {}".format(identitas["nik"]))
            print("  Nomor Ref     : {}".format(identitas["nomor_ref"]))
            print("  Nomor Laporan : {}".format(identitas["nomor_laporan"]))
        else:
            print("\n  [INFO] Halaman identitas nasabah tidak ditemukan (PDF parsial).")

        # 2. Total Baki Debet ringkasan
        total_baki_ringkasan = extract_total_baki_debet(pdf)
        if total_baki_ringkasan:
            print("\n  Total Baki Debet (Ringkasan): Rp {:,.0f}".format(total_baki_ringkasan))

        # 3. Parse halaman Kredit/Pembiayaan
        # Setiap halaman bisa berisi:
        #   (a) Satu fasilitas murni (baris pertama = "Kredit/Pembiayaan")
        #   (b) Fasilitas embedded di halaman lain (mis. setelah Ringkasan Fasilitas)
        #       → "Kredit/Pembiayaan" muncul di tengah halaman
        fasilitas_list = []
        kredit_page_count = 0

        for page in pdf.pages:
            raw = page.extract_text() or ""
            text = clean_cid(raw)
            if not text.strip():
                continue

            lines = text.strip().split("\n")
            first_line = lines[0].strip()

            if first_line == "Kredit/Pembiayaan":
                # Kasus (a): halaman murni fasilitas
                kredit_page_count += 1
                facility = parse_facility_page(raw)
                fasilitas_list.append(facility)

            elif "Kredit/Pembiayaan" in text:
                # Kasus (b): fasilitas embedded di halaman lain
                # Cari posisi "Kredit/Pembiayaan" di teks dan ambil sub-teks dari situ
                # Bisa ada lebih dari satu fasilitas per halaman (jarang, tapi mungkin)
                split_positions = [m.start() for m in re.finditer(
                    r"(?:^|\n)Kredit/Pembiayaan\n", text
                )]
                for pos in split_positions:
                    # Ambil sub-teks mulai dari "Kredit/Pembiayaan"
                    sub_text = text[pos:].lstrip("\n")
                    kredit_page_count += 1
                    facility = parse_facility_page(sub_text)
                    fasilitas_list.append(facility)

        print("\n  Halaman Kredit/Pembiayaan  : {}".format(kredit_page_count))
        print("  Fasilitas berhasil di-parse: {}".format(len(fasilitas_list)))

        # 4. Verifikasi (hanya jika PDF lengkap)
        print("\n" + "-" * 60)
        print("  VERIFIKASI")
        print("-" * 60)

        if len(fasilitas_list) == kredit_page_count:
            print("  [OK] Jumlah fasilitas ({}) sesuai dengan jumlah halaman ({})".format(
                len(fasilitas_list), kredit_page_count))
        else:
            print("  [WARNING] Jumlah fasilitas ({}) TIDAK SAMA dengan "
                  "jumlah halaman Kredit/Pembiayaan ({})".format(
                      len(fasilitas_list), kredit_page_count))

        if total_baki_ringkasan:
            total_baki_parsed = sum(f["baki_debet"] for f in fasilitas_list)
            print("\n  Total Baki Debet hasil parse : Rp {:,.0f}".format(total_baki_parsed))
            print("  Total Baki Debet Ringkasan   : Rp {:,.0f}".format(total_baki_ringkasan))
            selisih = abs(total_baki_parsed - total_baki_ringkasan)
            if selisih == 0:
                print("  [OK] Total Baki Debet sesuai")
            elif selisih <= 100:
                print("  [OK] Total Baki Debet sesuai (selisih pembulatan: Rp {:,})".format(selisih))
            else:
                print("  [WARNING] Selisih Total Baki Debet = Rp {:,.0f}".format(selisih))
        else:
            print("  [INFO] Ringkasan Fasilitas tidak ditemukan, verifikasi Baki Debet dilewati.")

        # 5. Tulis Excel
        # Nama file output = nama file input, hanya ekstensi yang berubah
        # (mis. "SLIK_Budi_2026.pdf" -> "SLIK_Budi_2026.xlsx")
        pdf_basename = os.path.splitext(os.path.basename(pdf_path))[0]
        out_filename = "{}.xlsx".format(pdf_basename)
        out_dir = os.path.dirname(os.path.abspath(pdf_path))
        out_path = os.path.join(out_dir, out_filename)

        write_excel(out_path, identitas, fasilitas_list)

        print("\n" + "-" * 60)
        print("  OUTPUT")
        print("-" * 60)
        print("  File Excel : {}".format(out_path))
        print("=" * 60 + "\n")

    return out_path


# ──────────────────── Excel Writer ───────────────────────────────

HEADERS = [
    "Pelapor", "Jenis", "Realisasi", "Jatuh Tempo",
    "Plafond Awal", "Plafond", "Baki Debet",
    "Kol", "Flag Restruk", "Rate", "Kol Terendah", "Kondisi"
]

COL_KEYS = [
    "pelapor", "jenis", "realisasi", "jatuh_tempo",
    "plafon_awal", "plafon", "baki_debet",
    "kol", "flag_restruk", "rate", "kol_terendah", "kondisi"
]

# Kolom yang nilainya angka murni (format ribuan)
NUMERIC_COLS = {
    COL_KEYS.index("plafon_awal"),  # 4
    COL_KEYS.index("plafon"),        # 5
    COL_KEYS.index("baki_debet"),    # 6
}

# Kolom Kol Terendah (integer atau NULL)
KOL_TERENDAH_IDX = COL_KEYS.index("kol_terendah")  # 10


def write_excel(out_path, identitas, fasilitas_list):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SLIK"

    num_cols = len(HEADERS)
    last_col = get_column_letter(num_cols)

    font_bold   = Font(name="Arial", bold=True, size=10)
    font_normal = Font(name="Arial", size=10)
    font_header = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    font_total  = Font(name="Arial", bold=True, size=10)

    fill_info   = PatternFill("solid", fgColor="D9E1F2")
    fill_header = PatternFill("solid", fgColor="2F5496")
    fill_total  = PatternFill("solid", fgColor="BDD7EE")
    fill_alt    = PatternFill("solid", fgColor="EEF3FB")

    al_left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    al_center = Alignment(horizontal="center", vertical="center")
    al_right  = Alignment(horizontal="right",  vertical="center")

    thin_side   = Side(style="thin", color="B4C6E7")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # ── Baris 1–3: Identitas ──
    info_vals = [
        "Nama: {}".format(identitas.get("nama", "")),
        "NIK: {}".format(identitas.get("nik", "")),
        "No. Referensi SLIK: {}".format(
            identitas.get("nomor_ref") or identitas.get("nomor_laporan", "")
        ),
    ]
    for row_i, val in enumerate(info_vals, start=1):
        ws.merge_cells("A{}:{}{}".format(row_i, last_col, row_i))
        cell = ws["A{}".format(row_i)]
        cell.value = val
        cell.font = font_bold
        cell.fill = fill_info
        cell.alignment = al_left
        cell.border = thin_border
        ws.row_dimensions[row_i].height = 18

    # ── Baris 4: Header kolom ──
    for col_i, hdr in enumerate(HEADERS, start=1):
        cell = ws.cell(row=4, column=col_i, value=hdr)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = al_center
        cell.border = thin_border
    ws.row_dimensions[4].height = 22

    # ── Baris 5+: Data fasilitas ──
    data_start = 5
    for row_off, fac in enumerate(fasilitas_list):
        row = data_start + row_off
        alt = row_off % 2 == 1
        for col_i, key in enumerate(COL_KEYS, start=1):
            cell = ws.cell(row=row, column=col_i)
            cell.border = thin_border
            cell.font = font_normal
            if alt:
                cell.fill = fill_alt

            raw = fac.get(key)

            if col_i - 1 in NUMERIC_COLS:
                # Nilai numerik: int atau 0
                cell.value = int(raw) if raw else 0
                cell.number_format = "#,##0"
                cell.alignment = al_right

            elif col_i - 1 == KOL_TERENDAH_IDX:
                # Kol Terendah: int jika ada histori, "NULL" jika tidak ada
                if raw is None:
                    cell.value = "NULL"
                else:
                    cell.value = int(raw)
                cell.alignment = al_center

            else:
                cell.value = raw if raw is not None else ""
                cell.alignment = al_left

        ws.row_dimensions[row].height = 16

    # ── Baris Total ──
    total_row = data_start + len(fasilitas_list)
    # Merge A–F untuk label "TOTAL"
    ws.merge_cells("A{}:F{}".format(total_row, total_row))
    tc = ws["A{}".format(total_row)]
    tc.value = "TOTAL"
    tc.font = font_total
    tc.fill = fill_total
    tc.alignment = al_center
    tc.border = thin_border

    # Formula SUM untuk Baki Debet
    baki_col_i  = COL_KEYS.index("baki_debet") + 1
    baki_letter = get_column_letter(baki_col_i)
    bc = ws.cell(row=total_row, column=baki_col_i)
    if fasilitas_list:
        bc.value = "=SUM({}{}:{}{})".format(
            baki_letter, data_start, baki_letter, total_row - 1)
    else:
        bc.value = 0
    bc.font = font_total
    bc.fill = fill_total
    bc.number_format = "#,##0"
    bc.alignment = al_right
    bc.border = thin_border

    # Style kolom lain di baris total
    for col_i in range(2, num_cols + 1):
        if col_i != baki_col_i:
            c = ws.cell(row=total_row, column=col_i)
            c.fill = fill_total
            c.border = thin_border
    ws.row_dimensions[total_row].height = 18

    # ── Lebar kolom ──
    col_widths = [42, 20, 14, 14, 18, 18, 18, 6, 12, 9, 12, 18]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A5"
    wb.save(out_path)


# ──────────────────── Entry Point ────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        print("\nERROR: Harap masukkan path file PDF.")
        print("Contoh: python slik_parser.py slik_nasabah.pdf")
        sys.exit(1)

    parse_slik(sys.argv[1])
