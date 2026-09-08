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

PENTING: file ini harus berada di folder YANG SAMA dengan slik_parser.py
(script ini memanggil fungsi dari slik_parser.py, tidak menduplikasi logikanya).

Instalasi dependensi:
    pip install pdfplumber openpyxl tkinterdnd2

Penggunaan:
    python gui_slik.py

Requirements:
    Python >= 3.8
    pdfplumber >= 0.9
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


# ─────────────────────────── GUI Utama ───────────────────────────────────

class App:
    def __init__(self, root):
        self.root = root
        root.title("SLIK Parser - Drag & Drop")
        set_window_icon(root)
        root.geometry("700x680")
        root.configure(bg="#1e293b")
        root.minsize(560, 500)

        self.out_queue = queue.Queue()
        self.processing = False
        self.results = []          # hanya hasil yang SUKSES: {..., "var": BooleanVar}
        self.total_processed = 0   # hitungan kumulatif semua file (sukses + gagal)

        # ── Area drop (bisa di-drag, bisa juga diklik) ──
        self.drop_frame = tk.Frame(
            root, bg="#334155", highlightbackground="#64748b",
            highlightthickness=2, height=120, cursor="hand2"
        )
        self.drop_frame.pack(fill="x", padx=16, pady=(16, 8))
        self.drop_frame.pack_propagate(False)

        self.drop_label = tk.Label(
            self.drop_frame,
            text="\u2b07  Seret file PDF ke sini,  atau  KLIK untuk pilih file  \u2b07",
            bg="#334155", fg="#e2e8f0", font=("Segoe UI", 12, "bold"),
            wraplength=580, justify="center", cursor="hand2"
        )
        self.drop_label.pack(expand=True)

        for widget in (self.drop_frame, self.drop_label):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self.on_drop)
            widget.bind("<Button-1>", lambda e: self.browse_files())

        # ── Status + tombol unduh yang dicentang ──
        top_bar = tk.Frame(root, bg="#1e293b")
        top_bar.pack(fill="x", padx=16)

        self.status_label = tk.Label(
            top_bar, text="Menunggu file...", bg="#1e293b", fg="#94a3b8"
        )
        self.status_label.pack(side="left")

        self.download_all_btn = tk.Button(
            top_bar, text="\u2b07 Unduh yang Dicentang...", command=self.download_checked,
            bg="#16a34a", fg="white", relief="flat", padx=10, pady=4,
            state="disabled", activebackground="#15803d", activeforeground="white"
        )
        self.download_all_btn.pack(side="right")

        # ── Daftar hasil (scrollable, checkbox + tombol unduh per baris) ──
        results_outer = tk.Frame(root, bg="#1e293b")
        results_outer.pack(fill="both", expand=True, padx=16, pady=8)

        tk.Label(
            results_outer, text="Hasil:", bg="#1e293b", fg="#e2e8f0",
            font=("Segoe UI", 10, "bold"), anchor="w"
        ).pack(fill="x")

        canvas_frame = tk.Frame(results_outer, bg="#0f172a", height=180)
        canvas_frame.pack(fill="both", expand=True, pady=(4, 0))
        canvas_frame.pack_propagate(False)

        self.results_canvas = tk.Canvas(canvas_frame, bg="#0f172a", highlightthickness=0)
        results_scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.results_canvas.yview)
        self.results_inner = tk.Frame(self.results_canvas, bg="#0f172a")

        self.results_inner.bind(
            "<Configure>",
            lambda e: self.results_canvas.configure(scrollregion=self.results_canvas.bbox("all"))
        )
        self.results_canvas.create_window((0, 0), window=self.results_inner, anchor="nw")
        self.results_canvas.configure(yscrollcommand=results_scrollbar.set)
        self.results_canvas.pack(side="left", fill="both", expand=True)
        results_scrollbar.pack(side="right", fill="y")

        self.empty_label = tk.Label(
            self.results_inner, text="Belum ada file diproses.",
            bg="#0f172a", fg="#64748b", font=("Segoe UI", 9)
        )
        self.empty_label.pack(pady=12)

        # ── Log detail (kecil, untuk diagnosa) ──
        tk.Label(
            root, text="Log detail:", bg="#1e293b", fg="#94a3b8",
            font=("Segoe UI", 9), anchor="w"
        ).pack(fill="x", padx=16)

        log_frame = tk.Frame(root, bg="#1e293b")
        log_frame.pack(fill="both", expand=False, padx=16, pady=(4, 8))

        self.log_text = tk.Text(
            log_frame, bg="#0f172a", fg="#64748b", font=("Consolas", 8),
            wrap="word", relief="flat", height=7
        )
        log_scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scrollbar.pack(side="right", fill="y")

        # ── Footer: tombol kontak developer (email disembunyikan di tombol utama) ──
        footer = tk.Frame(root, bg="#1e293b")
        footer.pack(fill="x", padx=16, pady=(0, 12))

        tk.Button(
            footer, text="Kontak Developer", command=self.show_contact_popup,
            bg="#334155", fg="#cbd5e1", relief="flat", padx=10, pady=4,
            font=("Segoe UI", 8), activebackground="#475569", activeforeground="white"
        ).pack(side="right")

        self.root.after(100, self.poll_queue)

    # ── Popup kecil di pojok kanan-bawah layar berisi link email developer ──
    def show_contact_popup(self):
        popup = tk.Toplevel(self.root)
        popup.title("Kontak Developer")
        set_window_icon(popup)
        popup.resizable(False, False)
        popup.configure(bg="#1e293b")
        popup.attributes("-topmost", True)

        w, h = 260, 110
        sw = popup.winfo_screenwidth()
        sh = popup.winfo_screenheight()
        x = sw - w - 20
        y = sh - h - 80  # beri jarak dari taskbar
        popup.geometry("{}x{}+{}+{}".format(w, h, x, y))

        tk.Label(
            popup, text="Hubungi developer via email:",
            bg="#1e293b", fg="#e2e8f0", font=("Segoe UI", 9)
        ).pack(pady=(16, 4))

        link = tk.Label(
            popup, text=DEVELOPER_EMAIL, bg="#1e293b", fg="#60a5fa",
            font=("Segoe UI", 10, "underline"), cursor="hand2"
        )
        link.pack()
        link.bind("<Button-1>", lambda e: self.open_mail_client())

        tk.Button(
            popup, text="Tutup", command=popup.destroy,
            bg="#334155", fg="white", relief="flat", padx=10, pady=3,
            font=("Segoe UI", 8)
        ).pack(pady=(12, 0))

    # ── Buka aplikasi email default ke alamat developer ──
    def open_mail_client(self):
        webbrowser.open("mailto:{}".format(DEVELOPER_EMAIL))

    # ── Handler saat file di-drop ──
    def on_drop(self, event):
        paths = self.root.tk.splitlist(event.data)
        pdf_paths = [p for p in paths if p.lower().endswith(".pdf")]
        skipped = len(paths) - len(pdf_paths)
        if skipped:
            self.log("[INFO] {} file dilewati (bukan .pdf)\n".format(skipped))
        if not pdf_paths:
            self.log("[INFO] Tidak ada file .pdf yang valid di antara file yang di-drop.\n")
            return
        self.start_processing(pdf_paths)

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
        self.status_label.config(text="Memproses {} file...".format(len(pdf_paths)))
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
                elif kind == "result":
                    self.add_result_row(payload)
                elif kind == "done":
                    self.processing = False
                    self.status_label.config(
                        text="{} file diproses".format(self.total_processed)
                    )
        except queue.Empty:
            pass
        self.root.after(100, self.poll_queue)

    # ── Tambah satu baris ke daftar hasil ──
    def add_result_row(self, result):
        if self.empty_label.winfo_exists():
            self.empty_label.destroy()

        self.total_processed += 1
        row = tk.Frame(self.results_inner, bg="#0f172a")
        row.pack(fill="x", pady=2)

        if result["success"]:
            var = tk.BooleanVar(value=True)  # default tercentang
            result["var"] = var

            tk.Checkbutton(
                row, variable=var, bg="#0f172a", activebackground="#0f172a",
                selectcolor="#1e293b", highlightthickness=0, bd=0
            ).pack(side="left", padx=(2, 4))

            tk.Label(row, text="\u2705", bg="#0f172a").pack(side="left", padx=(0, 6))
            tk.Label(
                row, text=result["output_name"], bg="#0f172a", fg="#cbd5e1",
                font=("Segoe UI", 9), anchor="w"
            ).pack(side="left", fill="x", expand=True)

            tk.Button(
                row, text="\u2b07 Unduh", bg="#3b82f6", fg="white", relief="flat",
                padx=8, pady=2, font=("Segoe UI", 8),
                command=lambda r=result: self.download_single(r)
            ).pack(side="right", padx=4)

            self.results.append(result)
            self.download_all_btn.config(state="normal")
        else:
            tk.Label(row, text="\u274c", bg="#0f172a").pack(side="left", padx=(4, 6))
            tk.Label(
                row, text="{} (gagal diproses)".format(result["input_name"]),
                bg="#0f172a", fg="#f87171", font=("Segoe UI", 9), anchor="w"
            ).pack(side="left", fill="x", expand=True)

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

    # ── Unduh semua file yang TERCENTANG ke satu folder ──
    def download_checked(self):
        checked = [r for r in self.results if r["var"].get()]
        if not checked:
            self.log("[INFO] Tidak ada file yang dicentang untuk diunduh.\n")
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

    def log(self, text):
        self.log_text.insert("end", text)
        self.log_text.see("end")


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    App(root)
    root.mainloop()
