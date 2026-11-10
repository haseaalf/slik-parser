"""
SLIK Parser - GUI Drag & Drop (Multi-file)
============================================
Jendela dengan area drop untuk memproses banyak PDF SLIK sekaligus.
- Area kotak bisa di-DRAG file ke situ, ATAU DIKLIK untuk membuka
  Windows Explorer (pilih file manual).
- Hasil .xlsx TIDAK langsung tersimpan di folder PDF asal. Setelah
  proses selesai, centang file yang diinginkan (default semua
  tercentang), lalu klik "Unduh" per file atau "Unduh yang
  Dicentang..." untuk memilih lokasi penyimpanan sendiri.

Tampilan memakai palet merah-putih-abu yang diambil dari logo
SLIK Parser (selaras dengan identitas visual OJK).

PENTING: file ini harus berada di folder YANG SAMA dengan slik_parser.py,
logo.png, dan icon.ico (script ini memanggil fungsi dari slik_parser.py,
tidak menduplikasi logikanya).

Instalasi dependensi:
    pip install pdfplumber openpyxl tkinterdnd2

Penggunaan:
    python gui_slik.py

Requirements:
    Python >= 3.8
    pdfplumber >= 0.9  (sudah membawa Pillow, dipakai untuk logo)
    openpyxl >= 3.0
    tkinterdnd2 >= 0.3
"""

import os
import sys
import shutil
import tempfile
import atexit
import queue
import threading
import traceback
import webbrowser

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    print("ERROR: Library 'tkinterdnd2' tidak ditemukan. Install dengan: pip install tkinterdnd2")
    sys.exit(1)

import tkinter as tk
from tkinter import ttk, filedialog

try:
    import slik_parser
except ImportError:
    print("ERROR: 'slik_parser.py' tidak ditemukan di folder yang sama dengan gui_slik.py")
    sys.exit(1)


DEVELOPER_EMAIL = "haseaalfian6@gmail.com"
ICON_FILENAME = "icon.ico"
LOGO_FILENAME = "logo.png"

# ─────────── Palet warna (merah-putih-abu, diambil dari logo) ───────────
RED = "#B5121B"          # aksi utama, aksen
RED_DARK = "#980008"     # hover tombol merah, teks di atas latar merah muda
RED_BRIGHT = "#E01820"   # strip header
RED_TINT = "#FDECEC"     # latar hover / area drop aktif
INK = "#2B2B2B"          # teks utama
GREY = "#484848"         # judul
MUTED = "#5F5F5F"        # teks keterangan (kontras >= 6:1 di atas putih)
LINE = "#D6D6D6"         # garis pemisah / border
BG = "#F3F4F6"           # latar jendela
CARD = "#FFFFFF"         # latar kartu
OK_GREEN = "#1E7B34"
WARN_ORANGE = "#9A5B00"
FONT = "Segoe UI"


def resource_path(relative_path):
    """Cari path file resource (mis. icon.ico), baik saat dijalankan
    sebagai script .py biasa MAUPUN setelah dibungkus jadi .exe oleh
    PyInstaller (yang mengekstrak file ke folder sementara _MEIPASS)."""
    base_path = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base_path, relative_path)


def set_window_icon(window):
    """Pasang icon.ico ke title bar & taskbar, mengganti icon 'bulu'
    bawaan Tk. Dibungkus try/except supaya tidak crash kalau file
    icon tidak ditemukan atau berjalan di platform non-Windows."""
    try:
        window.iconbitmap(resource_path(ICON_FILENAME))
    except Exception:
        pass


def enable_dpi_awareness():
    """Minta Windows merender jendela pada resolusi asli layar, supaya teks
    tidak buram di laptop dengan scaling 125%/150%. Harus dipanggil
    sebelum jendela Tk dibuat."""
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

# ─────────── Folder staging sementara (BUKAN folder asal PDF) ───────────
# Semua file .xlsx hasil parsing dipindahkan ke sini dulu, baru user
# memilih lokasi akhir lewat tombol "Unduh". Folder ini otomatis
# dihapus saat aplikasi ditutup.

SESSION_TEMP_DIR = tempfile.mkdtemp(prefix="slik_gui_")
atexit.register(shutil.rmtree, SESSION_TEMP_DIR, True)


# ─────────────────── Redirect print() dari slik_parser ke GUI ───────────

class TextRedirector:
    """Menampung print() ke queue supaya thread background aman mengirim
    teks ke jendela utama (Tkinter tidak boleh diakses langsung dari
    thread lain)."""

    def __init__(self, out_queue):
        self.out_queue = out_queue

    def write(self, s):
        if s:
            self.out_queue.put(("log", s))

    def flush(self):
        pass


# ─────────────────── Proses berjalan di background thread ───────────────

def process_files(file_paths, out_queue):
    """Panggil slik_parser.parse_slik() untuk tiap file, lalu langsung
    pindahkan hasil .xlsx keluar dari folder PDF asal ke folder staging
    sementara -- supaya tidak ada file yang tertinggal di folder sumber."""
    old_stdout = sys.stdout
    sys.stdout = TextRedirector(out_queue)
    total = len(file_paths)
    try:
        for i, path in enumerate(file_paths, start=1):
            display_name = os.path.basename(path)
            out_queue.put(("current", (i, total, display_name)))
            print("\n{}/{}  Memproses: {}".format(i, total, display_name))
            try:
                out_path = slik_parser.parse_slik(path)

                staged_dir = tempfile.mkdtemp(dir=SESSION_TEMP_DIR)
                staged_path = os.path.join(staged_dir, os.path.basename(out_path))
                shutil.move(out_path, staged_path)

                out_queue.put(("result", {
                    "input_name": display_name,
                    "output_name": os.path.basename(staged_path),
                    "staged_path": staged_path,
                    "success": True,
                }))
            except Exception as e:
                print("  [GAGAL] {}: {}".format(display_name, e))
                print(traceback.format_exc())
                out_queue.put(("result", {
                    "input_name": display_name,
                    "output_name": None,
                    "staged_path": None,
                    "success": False,
                }))
    finally:
        sys.stdout = old_stdout
    out_queue.put(("done", None))


# ─────────────────────────── Komponen tampilan ──────────────────────────

BUTTON_STYLES = {
    # bg, fg, hover, bg saat nonaktif, fg saat nonaktif
    "primary": (RED, "#FFFFFF", RED_DARK, "#E7B9BC", "#FFFFFF"),
    "secondary": (RED_TINT, RED_DARK, "#F8D3D5", "#F1F1F1", "#9A9A9A"),
    "quiet": (BG, MUTED, "#E4E5E8", BG, "#9A9A9A"),
}


class FlatButton(tk.Button):
    """Tombol datar berwarna dengan efek hover dan tampilan nonaktif yang
    jelas (tk.Button bawaan Windows tidak mengubah warna latar saat
    dinonaktifkan)."""

    def __init__(self, parent, text, command, kind="primary", font=None, padx=14, pady=6):
        self.colors = BUTTON_STYLES[kind]
        super().__init__(
            parent, text=text, command=command, relief="flat", bd=0,
            highlightthickness=0, padx=padx, pady=pady,
            font=font or (FONT, 10, "bold"),
        )
        self._hover = False
        self.bind("<Enter>", lambda e: self._set_hover(True))
        self.bind("<Leave>", lambda e: self._set_hover(False))
        self._apply()

    def set_enabled(self, enabled):
        self.config(state="normal" if enabled else "disabled")
        self._apply()

    def _set_hover(self, hover):
        self._hover = hover
        self._apply()

    def _apply(self):
        bg, fg, hover, dis_bg, dis_fg = self.colors
        if str(self["state"]) == "disabled":
            self.config(bg=dis_bg, disabledforeground=dis_fg, cursor="arrow")
        else:
            cur = hover if self._hover else bg
            self.config(bg=cur, fg=fg, activebackground=hover,
                        activeforeground=fg, cursor="hand2")


def card(parent, **pack_kw):
    """Kartu putih ber-border tipis."""
    frame = tk.Frame(parent, bg=CARD, highlightthickness=1,
                     highlightbackground=LINE, highlightcolor=LINE)
    frame.pack(**pack_kw)
    return frame


# ─────────────────────────── GUI Utama ───────────────────────────────────

class App:
    def __init__(self, root):
        self.root = root
        root.title("SLIK Parser")
        set_window_icon(root)
        root.configure(bg=BG)

        # Skala piksel mengikuti DPI layar (1.0 = 100%)
        self.scale = max(1.0, root.winfo_fpixels("1i") / 96.0)
        S = self.S
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        width = min(S(780), sw - S(40))
        height = min(S(720), sh - S(90))
        # Posisikan di tengah layar (sedikit ke atas, memberi ruang taskbar)
        x = max(0, (sw - width) // 2)
        y = max(0, (sh - height - S(60)) // 2)
        root.geometry("{}x{}+{}+{}".format(width, height, x, y))
        root.minsize(min(S(620), width), min(S(600), height))

        self.out_queue = queue.Queue()
        self.processing = False
        self.results = []          # hanya hasil yang SUKSES: {..., "var": BooleanVar}
        self.total_processed = 0   # hitungan kumulatif semua file (sukses + gagal)
        self.count_ok = 0
        self.count_fail = 0
        self.batch_total = 0
        self.batch_done = 0
        self.drop_active = False

        self._setup_ttk_styles()
        self._build_header()
        # Footer & log di-pack (side="bottom") SEBELUM daftar hasil, supaya di
        # layar kecil yang menyusut adalah daftar hasil, bukan footer/log.
        self._build_footer()

        body = tk.Frame(root, bg=BG)
        body.pack(fill="both", expand=True, padx=S(16), pady=(S(12), 0))

        self._build_drop_zone(body)
        self._build_status_bar(body)
        self._build_log(body)
        self._build_results(body)

        self.root.after(100, self.poll_queue)

    def S(self, px):
        """Konversi ukuran piksel (desain 100%) ke piksel layar sebenarnya."""
        return int(round(px * self.scale))

    # ── Gaya widget ttk (progress bar & scrollbar) ──
    def _setup_ttk_styles(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "Red.Horizontal.TProgressbar", troughcolor="#ECECEC",
            background=RED, lightcolor=RED, darkcolor=RED, bordercolor=LINE,
            thickness=self.S(8),
        )
        style.configure(
            "Light.Vertical.TScrollbar", background="#D3D3D3", troughcolor=CARD,
            bordercolor=CARD, arrowcolor=MUTED, lightcolor="#D3D3D3",
            darkcolor="#D3D3D3",
        )
        style.map("Light.Vertical.TScrollbar", background=[("active", "#B9B9B9")])

    # ── Header: logo + judul + strip merah ──
    def _build_header(self):
        S = self.S
        header = tk.Frame(self.root, bg=CARD)
        header.pack(fill="x")

        inner = tk.Frame(header, bg=CARD)
        inner.pack(fill="x", padx=S(16), pady=S(10))

        self.logo_img = self._load_logo(S(84))
        if self.logo_img is not None:
            tk.Label(inner, image=self.logo_img, bg=CARD).pack(side="left")
        else:
            tk.Label(inner, text="SLIK Parser", bg=CARD, fg=RED,
                     font=(FONT, 20, "bold")).pack(side="left")

        text_col = tk.Frame(inner, bg=CARD)
        text_col.pack(side="left", fill="x", expand=True, padx=(S(16), 0))
        tk.Label(
            text_col, text="Konversi Laporan iDeb SLIK ke Excel", bg=CARD, fg=GREY,
            font=(FONT, 15, "bold"), anchor="w",
        ).pack(fill="x")
        tk.Label(
            text_col,
            text="Ekstrak data fasilitas Kredit/Pembiayaan dari PDF SLIK OJK "
                 "menjadi file .xlsx siap analisis.",
            bg=CARD, fg=MUTED, font=(FONT, 10), anchor="w", justify="left",
            wraplength=S(520),
        ).pack(fill="x", pady=(S(2), 0))

        tk.Frame(self.root, bg=RED_BRIGHT, height=S(4)).pack(fill="x")

    def _load_logo(self, height):
        """Muat logo.png dan perkecil ke tinggi tertentu. Kembalikan None jika
        logo tidak ditemukan (GUI tetap jalan dengan judul teks)."""
        path = resource_path(LOGO_FILENAME)
        try:
            from PIL import Image, ImageTk
            img = Image.open(path)
            w = round(img.width * height / img.height)
            return ImageTk.PhotoImage(img.resize((w, height), Image.LANCZOS))
        except Exception:
            try:
                img = tk.PhotoImage(file=path)
                factor = max(1, round(img.height() / height))
                return img.subsample(factor, factor)
            except Exception:
                return None

    # ── Area drop (bisa di-drag, bisa juga diklik) ──
    def _build_drop_zone(self, parent):
        S = self.S
        self.drop_canvas = tk.Canvas(
            parent, bg=CARD, height=S(132), highlightthickness=0, cursor="hand2",
        )
        self.drop_canvas.pack(fill="x")
        self.drop_canvas.bind("<Configure>", lambda e: self._draw_drop_zone())
        self.drop_canvas.bind("<Button-1>", lambda e: self.browse_files())
        self.drop_canvas.bind("<Enter>", lambda e: self._set_drop_active(True))
        self.drop_canvas.bind("<Leave>", lambda e: self._set_drop_active(False))

        self.drop_canvas.drop_target_register(DND_FILES)
        self.drop_canvas.dnd_bind("<<Drop>>", self.on_drop)
        self.drop_canvas.dnd_bind("<<DropEnter>>", self._on_drag_enter)
        self.drop_canvas.dnd_bind("<<DropLeave>>", self._on_drag_leave)

    def _draw_drop_zone(self):
        S = self.S
        c = self.drop_canvas
        c.delete("all")
        w, h = c.winfo_width(), c.winfo_height()
        if w < 10 or h < 10:
            return
        active = self.drop_active and not self.processing
        fill = RED_TINT if active else CARD
        outline = RED if active else "#A8A8A8"
        c.configure(bg=fill, cursor="watch" if self.processing else "hand2")
        m = S(6)
        c.create_rectangle(m, m, w - m, h - m, outline=outline, width=2, dash=(6, 4))

        cx = w // 2
        if self.processing:
            c.create_text(cx, h // 2 - S(10), text="Sedang memproses...",
                          fill=GREY, font=(FONT, 14, "bold"))
            c.create_text(cx, h // 2 + S(18), text="Mohon tunggu hingga semua file selesai.",
                          fill=MUTED, font=(FONT, 10))
            return

        c.create_text(cx, h // 2 - S(30), text="Seret & lepas file PDF SLIK di sini",
                      fill=INK, font=(FONT, 14, "bold"))
        c.create_text(cx, h // 2 - S(4),
                      text="bisa lebih dari satu file sekaligus",
                      fill=MUTED, font=(FONT, 10))
        # "Tombol" di dalam area drop (klik di mana pun pada area ini = pilih file)
        bw, bh = S(150), S(32)
        by = h // 2 + S(30)
        c.create_rectangle(cx - bw // 2, by - bh // 2, cx + bw // 2, by + bh // 2,
                           fill=RED_DARK if active else RED, outline="")
        c.create_text(cx, by, text="Pilih File PDF", fill="#FFFFFF",
                      font=(FONT, 10, "bold"))

    def _set_drop_active(self, active):
        self.drop_active = active
        self._draw_drop_zone()

    def _on_drag_enter(self, event):
        self._set_drop_active(True)
        return event.action

    def _on_drag_leave(self, event):
        self._set_drop_active(False)
        return event.action

    # ── Status + progress + tombol unduh yang dicentang ──
    def _build_status_bar(self, parent):
        S = self.S
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", pady=(S(12), 0))

        self.download_all_btn = FlatButton(
            bar, text="Unduh yang Dicentang...", command=self.download_checked,
            kind="primary",
        )
        self.download_all_btn.set_enabled(False)
        self.download_all_btn.pack(side="right")

        left = tk.Frame(bar, bg=BG)
        left.pack(side="left", fill="x", expand=True, padx=(0, S(16)))
        self.status_label = tk.Label(
            left, text="Menunggu file. Seret PDF ke kotak di atas atau klik untuk memilih.",
            bg=BG, fg=INK, font=(FONT, 10), anchor="w",
        )
        self.status_label.pack(fill="x")
        self.progress = ttk.Progressbar(
            left, style="Red.Horizontal.TProgressbar", mode="determinate", maximum=1,
        )
        self.progress.pack(fill="x", pady=(S(4), 0))

    # ── Daftar hasil (scrollable, checkbox + tombol unduh per baris) ──
    def _build_results(self, parent):
        S = self.S
        box = card(parent, fill="both", expand=True, pady=(S(10), 0))

        head = tk.Frame(box, bg=CARD)
        head.pack(fill="x", padx=S(12), pady=(S(8), S(6)))
        tk.Label(head, text="Hasil", bg=CARD, fg=GREY,
                 font=(FONT, 11, "bold")).pack(side="left")
        self.summary_label = tk.Label(head, text="", bg=CARD, fg=MUTED, font=(FONT, 9))
        self.summary_label.pack(side="right")
        tk.Frame(box, bg=LINE, height=1).pack(fill="x")

        list_frame = tk.Frame(box, bg=CARD)
        list_frame.pack(fill="both", expand=True)

        self.results_canvas = tk.Canvas(list_frame, bg=CARD, highlightthickness=0,
                                        height=S(120))
        results_scrollbar = ttk.Scrollbar(
            list_frame, orient="vertical", command=self.results_canvas.yview,
            style="Light.Vertical.TScrollbar",
        )
        self.results_inner = tk.Frame(self.results_canvas, bg=CARD)

        self.results_inner.bind(
            "<Configure>",
            lambda e: self.results_canvas.configure(scrollregion=self.results_canvas.bbox("all"))
        )
        inner_id = self.results_canvas.create_window((0, 0), window=self.results_inner, anchor="nw")
        # Lebar isi daftar selalu mengikuti lebar kotak (tombol Unduh rata kanan)
        self.results_canvas.bind(
            "<Configure>",
            lambda e: self.results_canvas.itemconfigure(inner_id, width=e.width)
        )
        self.results_canvas.configure(yscrollcommand=results_scrollbar.set)
        self.results_canvas.pack(side="left", fill="both", expand=True)
        results_scrollbar.pack(side="right", fill="y")
        self._bind_mousewheel(list_frame, self.results_canvas)

        self.empty_label = tk.Label(
            self.results_inner,
            text="Belum ada file diproses. Hasil konversi akan muncul di sini.",
            bg=CARD, fg=MUTED, font=(FONT, 10),
        )
        self.empty_label.pack(pady=S(18))

    def _bind_mousewheel(self, area, canvas):
        def on_wheel(event):
            if canvas.yview() != (0.0, 1.0):
                canvas.yview_scroll(int(-event.delta / 120), "units")
        area.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", on_wheel))
        area.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

    # ── Log detail (untuk diagnosa) ──
    def _build_log(self, parent):
        S = self.S
        box = card(parent, side="bottom", fill="x", pady=(S(10), 0))

        head = tk.Frame(box, bg=CARD)
        head.pack(fill="x", padx=S(12), pady=(S(6), S(4)))
        tk.Label(head, text="Log proses", bg=CARD, fg=GREY,
                 font=(FONT, 10, "bold")).pack(side="left")
        tk.Label(head, text="Detail verifikasi tiap file", bg=CARD, fg=MUTED,
                 font=(FONT, 9)).pack(side="right")

        log_frame = tk.Frame(box, bg="#FAFAFA")
        log_frame.pack(fill="x", padx=S(1), pady=(0, S(1)))

        self.log_text = tk.Text(
            log_frame, bg="#FAFAFA", fg=INK, font=("Consolas", 9),
            wrap="word", relief="flat", height=6, padx=S(8), pady=S(6),
            selectbackground=RED_TINT, selectforeground=INK,
        )
        log_scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview,
                                      style="Light.Vertical.TScrollbar")
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scrollbar.pack(side="right", fill="y")

        self.log_text.tag_configure("ok", foreground=OK_GREEN)
        self.log_text.tag_configure("warn", foreground=WARN_ORANGE, font=("Consolas", 9, "bold"))
        self.log_text.tag_configure("err", foreground=RED, font=("Consolas", 9, "bold"))
        self.log_text.tag_configure("info", foreground=MUTED)

    # ── Footer: catatan privasi + tombol kontak developer ──
    def _build_footer(self):
        S = self.S
        footer = tk.Frame(self.root, bg=BG)
        footer.pack(side="bottom", fill="x", padx=S(16), pady=(S(8), S(10)))

        FlatButton(
            footer, text="Kontak Developer", command=self.show_contact_popup,
            kind="quiet", font=(FONT, 9), padx=S(10), pady=S(4),
        ).pack(side="right")
        tk.Label(
            footer,
            text="File diproses di komputer ini. Hasil sementara dihapus saat aplikasi ditutup.",
            bg=BG, fg=MUTED, font=(FONT, 9), anchor="w",
        ).pack(side="left", fill="x", expand=True)

    # ── Popup kecil di pojok kanan-bawah layar berisi link email developer ──
    def show_contact_popup(self):
        S = self.S
        popup = tk.Toplevel(self.root)
        popup.title("Kontak Developer")
        set_window_icon(popup)
        popup.resizable(False, False)
        popup.configure(bg=CARD)
        popup.attributes("-topmost", True)

        w, h = S(300), S(140)
        sw = popup.winfo_screenwidth()
        sh = popup.winfo_screenheight()
        x = sw - w - S(20)
        y = sh - h - S(80)  # beri jarak dari taskbar
        popup.geometry("{}x{}+{}+{}".format(w, h, x, y))

        tk.Frame(popup, bg=RED_BRIGHT, height=S(4)).pack(fill="x")
        tk.Label(
            popup, text="Hubungi developer via email:",
            bg=CARD, fg=INK, font=(FONT, 10)
        ).pack(pady=(S(14), S(4)))

        link = tk.Label(
            popup, text=DEVELOPER_EMAIL, bg=CARD, fg=RED,
            font=(FONT, 10, "underline"), cursor="hand2"
        )
        link.pack()
        link.bind("<Button-1>", lambda e: self.open_mail_client())

        FlatButton(
            popup, text="Tutup", command=popup.destroy, kind="secondary",
            font=(FONT, 9, "bold"), padx=S(14), pady=S(3),
        ).pack(pady=(S(12), 0))

    # ── Buka aplikasi email default ke alamat developer ──
    def open_mail_client(self):
        webbrowser.open("mailto:{}".format(DEVELOPER_EMAIL))

    # ── Handler saat file di-drop ──
    def on_drop(self, event):
        self._set_drop_active(False)
        paths = self.root.tk.splitlist(event.data)
        pdf_paths = [p for p in paths if p.lower().endswith(".pdf")]
        skipped = len(paths) - len(pdf_paths)
        if skipped:
            self.log("[INFO] {} file dilewati (bukan .pdf)\n".format(skipped))
        if not pdf_paths:
            self.log("[INFO] Tidak ada file .pdf yang valid di antara file yang di-drop.\n")
            self.status_label.config(text="File yang diseret bukan PDF. Gunakan file PDF SLIK.")
            return event.action
        self.start_processing(pdf_paths)
        return event.action

    # ── Klik area drop -> buka Windows Explorer ──
    def browse_files(self):
        if self.processing:
            return
        paths = filedialog.askopenfilenames(
            title="Pilih file PDF SLIK", filetypes=[("PDF files", "*.pdf")]
        )
        if paths:
            self.start_processing(list(paths))

    # ── Mulai batch processing di background thread ──
    def start_processing(self, pdf_paths):
        if self.processing:
            self.log("[INFO] Masih memproses batch sebelumnya, tunggu selesai dulu.\n")
            return
        self.processing = True
        self.batch_total = len(pdf_paths)
        self.batch_done = 0
        self.progress.config(maximum=self.batch_total, value=0)
        self.status_label.config(text="Memproses {} file...".format(len(pdf_paths)))
        self._draw_drop_zone()
        threading.Thread(
            target=process_files,
            args=(pdf_paths, self.out_queue),
            daemon=True
        ).start()

    # ── Polling queue dari thread background (tiap 100ms) ──
    def poll_queue(self):
        try:
            while True:
                kind, payload = self.out_queue.get_nowait()
                if kind == "log":
                    self.log(payload)
                elif kind == "current":
                    i, total, name = payload
                    self.status_label.config(
                        text="Memproses {} dari {}: {}".format(i, total, name))
                elif kind == "result":
                    self.add_result_row(payload)
                    self.batch_done += 1
                    self.progress.config(value=self.batch_done)
                elif kind == "done":
                    self.processing = False
                    self._draw_drop_zone()
                    self.status_label.config(
                        text="Selesai. Total {} file diproses: {} berhasil, {} gagal.".format(
                            self.total_processed, self.count_ok, self.count_fail)
                    )
        except queue.Empty:
            pass
        self.root.after(100, self.poll_queue)

    # ── Tambah satu baris ke daftar hasil ──
    def add_result_row(self, result):
        S = self.S
        if self.empty_label.winfo_exists():
            self.empty_label.destroy()

        self.total_processed += 1
        row = tk.Frame(self.results_inner, bg=CARD)
        row.pack(fill="x")
        content = tk.Frame(row, bg=CARD)
        content.pack(fill="x", padx=S(8), pady=S(6))

        if result["success"]:
            self.count_ok += 1
            var = tk.BooleanVar(value=True)  # default tercentang
            result["var"] = var

            tk.Checkbutton(
                content, variable=var, bg=CARD, activebackground=CARD,
                selectcolor=CARD, highlightthickness=0, bd=0, cursor="hand2",
            ).pack(side="left", padx=(0, S(4)))

            FlatButton(
                content, text="Unduh", kind="secondary", font=(FONT, 9, "bold"),
                padx=S(12), pady=S(3), command=lambda r=result: self.download_single(r),
            ).pack(side="right", padx=(S(8), 0))
            tk.Label(content, text="✓ Berhasil", bg=CARD, fg=OK_GREEN,
                     font=(FONT, 9, "bold")).pack(side="right")

            texts = tk.Frame(content, bg=CARD)
            texts.pack(side="left", fill="x", expand=True)
            tk.Label(texts, text=result["output_name"], bg=CARD, fg=INK,
                     font=(FONT, 10), anchor="w").pack(fill="x")
            tk.Label(texts, text="dari {}".format(result["input_name"]), bg=CARD,
                     fg=MUTED, font=(FONT, 8), anchor="w").pack(fill="x")

            self.results.append(result)
            self.download_all_btn.set_enabled(True)
        else:
            self.count_fail += 1
            tk.Label(content, text="✕ Gagal", bg=CARD, fg=RED,
                     font=(FONT, 9, "bold")).pack(side="right")
            texts = tk.Frame(content, bg=CARD)
            texts.pack(side="left", fill="x", expand=True, padx=(S(26), 0))
            tk.Label(texts, text=result["input_name"], bg=CARD, fg=INK,
                     font=(FONT, 10), anchor="w").pack(fill="x")
            tk.Label(texts, text="Gagal diproses. Lihat detail di Log proses.",
                     bg=CARD, fg=RED, font=(FONT, 8), anchor="w").pack(fill="x")

        tk.Frame(row, bg="#ECECEC", height=1).pack(fill="x")
        self.summary_label.config(
            text="{} berhasil  ·  {} gagal".format(self.count_ok, self.count_fail))

    # ── Unduh satu file hasil ──
    def download_single(self, result):
        dest = filedialog.asksaveasfilename(
            title="Simpan sebagai",
            initialfile=result["output_name"],
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if dest:
            shutil.copy2(result["staged_path"], dest)
            self.log("[INFO] Disimpan ke: {}\n".format(dest))
            self.status_label.config(text="Tersimpan: {}".format(dest))

    # ── Unduh semua file yang TERCENTANG ke satu folder ──
    def download_checked(self):
        checked = [r for r in self.results if r["var"].get()]
        if not checked:
            self.log("[INFO] Tidak ada file yang dicentang untuk diunduh.\n")
            self.status_label.config(text="Belum ada file yang dicentang untuk diunduh.")
            return
        dest_dir = filedialog.askdirectory(title="Pilih folder tujuan untuk file yang dicentang")
        if not dest_dir:
            return
        for r in checked:
            target = os.path.join(dest_dir, r["output_name"])
            base, ext = os.path.splitext(target)
            counter = 1
            while os.path.exists(target):
                target = "{}({}){}".format(base, counter, ext)
                counter += 1
            shutil.copy2(r["staged_path"], target)
        self.log("[INFO] {} file (dicentang) disimpan ke: {}\n".format(len(checked), dest_dir))
        self.status_label.config(
            text="{} file disimpan ke: {}".format(len(checked), dest_dir))

    def log(self, text):
        # Warna ditentukan per baris: [OK] hijau, [WARNING] oranye, gagal merah
        for line in text.splitlines(keepends=True):
            if "[OK]" in line:
                tag = "ok"
            elif "[WARNING]" in line:
                tag = "warn"
            elif "[GAGAL]" in line or "ERROR" in line or "Traceback" in line:
                tag = "err"
            elif "[INFO]" in line:
                tag = "info"
            else:
                tag = None
            self.log_text.insert("end", line, tag)
        self.log_text.see("end")


if __name__ == "__main__":
    enable_dpi_awareness()
    root = TkinterDnD.Tk()
    App(root)
    root.mainloop()
