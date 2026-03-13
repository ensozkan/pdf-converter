"""
PDF Dönüştürücü — Standalone Desktop Uygulaması
Sürükle-bırak veya dosya seç, PDF'e çevir.
"""

import os, io, sys, threading, tempfile, traceback
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

# ── Renkler & tema ──────────────────────────────────────────────
BG        = "#0f0f1a"
CARD      = "#1a1a2e"
SURFACE   = "#16213e"
ACCENT    = "#7c3aed"
ACCENT2   = "#a855f7"
SUCCESS   = "#10b981"
ERROR     = "#ef4444"
WARN      = "#f59e0b"
TEXT      = "#e2e8f0"
MUTED     = "#64748b"
BORDER    = "#2d2d44"

EXT_COLORS = {
    "jpg":"#f59e0b","jpeg":"#f59e0b","png":"#06b6d4","webp":"#06b6d4",
    "gif":"#ec4899","bmp":"#8b5cf6","tiff":"#8b5cf6","tif":"#8b5cf6",
    "txt":"#94a3b8","md":"#14b8a6","docx":"#3b82f6",
    "pptx":"#d97706","ppt":"#d97706",
    "html":"#f97316","htm":"#f97316",
    "csv":"#10b981","xlsx":"#22c55e","xls":"#22c55e","pdf":"#ef4444",
}

SUPPORTED_EXTS = {
    ".jpg",".jpeg",".png",".gif",".bmp",".tiff",".tif",".webp",
    ".txt",".md",".docx",".pptx",".ppt",
    ".html",".htm",".csv",".xlsx",".xls",".pdf",
}

# ── Dönüştürme fonksiyonları ─────────────────────────────────────

def image_to_pdf(data, suffix):
    from PIL import Image
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    img = Image.open(io.BytesIO(data))
    if img.mode in ("RGBA","P","LA"):
        bg = Image.new("RGB", img.size, (255,255,255))
        if img.mode == "P": img = img.convert("RGBA")
        if img.mode in ("RGBA","LA"): bg.paste(img, mask=img.split()[-1])
        else: bg.paste(img)
        img = bg
    elif img.mode != "RGB":
        img = img.convert("RGB")
    buf = io.BytesIO()
    W, H = A4
    iw, ih = img.size
    scale = min(W/iw, H/ih)
    nw, nh = iw*scale, ih*scale
    c = canvas.Canvas(buf, pagesize=A4)
    tmp = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
    img.save(tmp.name)
    tmp.close()
    c.drawImage(tmp.name, (W-nw)/2, (H-nh)/2, nw, nh)
    c.save()
    os.unlink(tmp.name)
    buf.seek(0)
    return buf.read()

def text_to_pdf(data):
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    text = data.decode("utf-8", errors="replace")
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=20*mm, rightMargin=20*mm, topMargin=20*mm, bottomMargin=20*mm)
    style = ParagraphStyle("code", fontName="Courier", fontSize=9,
                           leading=14, textColor=colors.HexColor("#1a1a2e"))
    story = []
    for line in text.split("\n"):
        safe = line.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
        story.append(Paragraph(safe or "&nbsp;", style))
    doc.build(story)
    buf.seek(0)
    return buf.read()

def docx_to_pdf(data):
    from docx import Document
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    tmp.write(data); tmp.close()
    doc = Document(tmp.name)
    os.unlink(tmp.name)
    buf = io.BytesIO()
    h1 = ParagraphStyle("H1", fontSize=18, fontName="Helvetica-Bold", leading=24, spaceAfter=8)
    h2 = ParagraphStyle("H2", fontSize=14, fontName="Helvetica-Bold", leading=20, spaceAfter=6)
    body = ParagraphStyle("Body", fontSize=10, fontName="Helvetica", leading=15, spaceAfter=4)
    pdfdoc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=20*mm, rightMargin=20*mm, topMargin=20*mm, bottomMargin=20*mm)
    story = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text: story.append(Spacer(1, 6)); continue
        safe = text.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
        if para.style.name.startswith("Heading 1"): story.append(Paragraph(safe, h1))
        elif para.style.name.startswith("Heading 2"): story.append(Paragraph(safe, h2))
        else: story.append(Paragraph(safe, body))
    pdfdoc.build(story)
    buf.seek(0)
    return buf.read()

def find_soffice():
    import shutil, platform, glob
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p: return p
    if platform.system() == "Windows":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ] + glob.glob(r"C:\Program Files\LibreOffice*\program\soffice.exe")
        for p in candidates:
            if os.path.exists(p): return p
    for p in ("/Applications/LibreOffice.app/Contents/MacOS/soffice",):
        if os.path.exists(p): return p
    return None

def pptx_to_pdf(data, suffix):
    import subprocess, shutil
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice bulunamadı.\n"
            "https://www.libreoffice.org adresinden indirip kurun."
        )
    tmp_dir = tempfile.mkdtemp()
    try:
        in_path = os.path.join(tmp_dir, f"input{suffix}")
        with open(in_path, "wb") as f: f.write(data)
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp_dir, in_path],
            capture_output=True, text=True, timeout=120
        )
        pdf_path = os.path.join(tmp_dir, "input.pdf")
        if not os.path.exists(pdf_path):
            raise RuntimeError(result.stderr or result.stdout or "LibreOffice dönüşümü başarısız")
        with open(pdf_path, "rb") as f: return f.read()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

def html_to_pdf(data):
    import weasyprint
    html = data.decode("utf-8", errors="replace")
    return weasyprint.HTML(string=html).write_pdf()

def csv_to_pdf(data):
    import pandas as pd
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    df = pd.read_csv(io.BytesIO(data)).fillna("")
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
        leftMargin=10*mm, rightMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm)
    rows = [df.columns.tolist()] + [[str(c) for c in r] for r in df.values.tolist()]
    t = Table(rows, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0), colors.HexColor("#16213e")),
        ("TEXTCOLOR",(0,0),(-1,0), colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1), 8),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.HexColor("#f8f9fa"),colors.white]),
        ("GRID",(0,0),(-1,-1), 0.5, colors.HexColor("#dee2e6")),
        ("PADDING",(0,0),(-1,-1), 4),
    ]))
    doc.build([t])
    buf.seek(0)
    return buf.read()

def xlsx_to_pdf(data):
    import pandas as pd
    df = pd.read_excel(io.BytesIO(data))
    buf = io.BytesIO()
    df.to_csv(buf, index=False)
    return csv_to_pdf(buf.getvalue())

def convert_file(path):
    ext = Path(path).suffix.lower()
    with open(path, "rb") as f:
        data = f.read()
    if ext in (".jpg",".jpeg",".png",".gif",".bmp",".tiff",".tif",".webp"):
        return image_to_pdf(data, ext)
    elif ext in (".txt",".md"):
        return text_to_pdf(data)
    elif ext == ".docx":
        return docx_to_pdf(data)
    elif ext in (".pptx",".ppt"):
        return pptx_to_pdf(data, ext)
    elif ext in (".html",".htm"):
        return html_to_pdf(data)
    elif ext == ".csv":
        return csv_to_pdf(data)
    elif ext in (".xlsx",".xls"):
        return xlsx_to_pdf(data)
    elif ext == ".pdf":
        return data
    else:
        raise ValueError(f'"{ext}" formatı desteklenmiyor')

# ── GUI ──────────────────────────────────────────────────────────

class RoundedButton(tk.Canvas):
    def __init__(self, parent, text, command=None, bg=ACCENT, fg=TEXT,
                 width=160, height=40, radius=10, **kwargs):
        super().__init__(parent, width=width, height=height,
                         bg=CARD, highlightthickness=0, **kwargs)
        self.command = command
        self.bg_color = bg
        self.hover_color = self._lighten(bg)
        self.fg = fg
        self.r = radius
        self.w = width
        self.h = height
        self.text = text
        self._draw(self.bg_color)
        self.bind("<Enter>", lambda e: self._draw(self.hover_color))
        self.bind("<Leave>", lambda e: self._draw(self.bg_color))
        self.bind("<Button-1>", lambda e: command() if command else None)

    def _lighten(self, hex_color):
        r = int(hex_color[1:3], 16)
        g = int(hex_color[3:5], 16)
        b = int(hex_color[5:7], 16)
        r = min(255, r + 30); g = min(255, g + 30); b = min(255, b + 30)
        return f"#{r:02x}{g:02x}{b:02x}"

    def _draw(self, color):
        self.delete("all")
        r, w, h = self.r, self.w, self.h
        self.create_arc(0, 0, 2*r, 2*r, start=90, extent=90, fill=color, outline=color)
        self.create_arc(w-2*r, 0, w, 2*r, start=0, extent=90, fill=color, outline=color)
        self.create_arc(0, h-2*r, 2*r, h, start=180, extent=90, fill=color, outline=color)
        self.create_arc(w-2*r, h-2*r, w, h, start=270, extent=90, fill=color, outline=color)
        self.create_rectangle(r, 0, w-r, h, fill=color, outline=color)
        self.create_rectangle(0, r, w, h-r, fill=color, outline=color)
        self.create_text(w//2, h//2, text=self.text, fill=self.fg,
                         font=("Segoe UI", 10, "bold"))

    def set_state(self, enabled):
        color = self.bg_color if enabled else MUTED
        self._draw(color)
        self.bind("<Button-1>", (lambda e: self.command()) if enabled else lambda e: None)


class FileRow(tk.Frame):
    def __init__(self, parent, filepath, on_remove, **kwargs):
        super().__init__(parent, bg=SURFACE, **kwargs)
        self.filepath = filepath
        self.on_remove = on_remove
        self.status = "wait"  # wait | loading | done | error
        self._build()

    def _build(self):
        ext = Path(self.filepath).suffix.lower().lstrip(".")
        color = EXT_COLORS.get(ext, ACCENT)
        name = Path(self.filepath).name
        size_kb = os.path.getsize(self.filepath) / 1024

        # Badge
        badge = tk.Label(self, text=ext.upper()[:4], bg=SURFACE,
                         fg=color, font=("Segoe UI", 7, "bold"), width=5)
        badge.pack(side="left", padx=(10,6), pady=10)

        # Info
        info = tk.Frame(self, bg=SURFACE)
        info.pack(side="left", fill="x", expand=True, pady=6)

        self.name_lbl = tk.Label(info, text=name, bg=SURFACE, fg=TEXT,
                                 font=("Segoe UI", 9, "bold"), anchor="w")
        self.name_lbl.pack(fill="x")

        self.sub_lbl = tk.Label(info, text=f"{size_kb:.1f} KB  •  Bekliyor",
                                bg=SURFACE, fg=MUTED, font=("Segoe UI", 8), anchor="w")
        self.sub_lbl.pack(fill="x")

        # Remove btn
        self.remove_btn = tk.Label(self, text="✕", bg=SURFACE, fg=MUTED,
                                   font=("Segoe UI", 10), cursor="hand2")
        self.remove_btn.pack(side="right", padx=12)
        self.remove_btn.bind("<Button-1>", lambda e: self.on_remove(self))
        self.remove_btn.bind("<Enter>", lambda e: self.remove_btn.config(fg=ERROR))
        self.remove_btn.bind("<Leave>", lambda e: self.remove_btn.config(fg=MUTED))

        # Separator
        sep = tk.Frame(self, bg=BORDER, height=1)
        sep.pack(side="bottom", fill="x")

    def set_status(self, status, extra=""):
        self.status = status
        if status == "loading":
            self.sub_lbl.config(text="⏳ Dönüştürülüyor...", fg=ACCENT2)
        elif status == "done":
            self.sub_lbl.config(text=f"✅ Kaydedildi → {extra}", fg=SUCCESS)
            self.remove_btn.config(text="")
        elif status == "error":
            self.sub_lbl.config(text=f"❌ {extra}", fg=ERROR)
        self.update_idletasks()


class App(TkinterDnD.Tk if HAS_DND else tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Dönüştürücü")
        self.geometry("620x680")
        self.minsize(500, 500)
        self.configure(bg=BG)
        self._set_icon()
        self.file_rows = []
        self._build_ui()
        if HAS_DND:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_drop)

    def _set_icon(self):
        try:
            # Create a simple icon programmatically
            icon_data = """
R0lGODlhIAAgAMIAAAAAAP+/AP//AP///wAAAAAAAAAAAAAAACH5BAUAAAAALAAAAAAIAAIAAANB
CLrc/jDKSYe9OOvNu/9gKI5kaZ5oqq5s675wLM90bd94ru987//AoHBILBqPyKRyyWw6n9Cot
GqtWq/YrHbL7Xq/4LAAAAA="""
        except:
            pass

    def _build_ui(self):
        # Header
        header = tk.Frame(self, bg=BG)
        header.pack(fill="x", padx=30, pady=(28,0))

        tk.Label(header, text="📄", font=("Segoe UI Emoji", 28),
                 bg=BG, fg=TEXT).pack(side="left")
        title_frame = tk.Frame(header, bg=BG)
        title_frame.pack(side="left", padx=12)
        tk.Label(title_frame, text="PDF Dönüştürücü", font=("Segoe UI", 18, "bold"),
                 bg=BG, fg=TEXT).pack(anchor="w")
        tk.Label(title_frame, text="Dosyaları sürükle bırak veya seç, PDF'e çevir",
                 font=("Segoe UI", 9), bg=BG, fg=MUTED).pack(anchor="w")

        # Drop zone
        self.drop_frame = tk.Frame(self, bg=CARD, relief="flat",
                                   highlightbackground=BORDER, highlightthickness=1)
        self.drop_frame.pack(fill="x", padx=30, pady=18)
        self.drop_frame.bind("<Button-1>", lambda e: self._pick_files())
        self.drop_frame.bind("<Enter>", lambda e: self._hover_drop(True))
        self.drop_frame.bind("<Leave>", lambda e: self._hover_drop(False))

        inner = tk.Frame(self.drop_frame, bg=CARD, pady=24)
        inner.pack()
        inner.bind("<Button-1>", lambda e: self._pick_files())

        tk.Label(inner, text="🗂️", font=("Segoe UI Emoji", 28),
                 bg=CARD, fg=TEXT, cursor="hand2").pack()
        tk.Label(inner, text="Dosyaları buraya sürükle ya da tıkla",
                 font=("Segoe UI", 10, "bold"), bg=CARD, fg=TEXT, cursor="hand2").pack(pady=(6,2))

        tags = tk.Frame(inner, bg=CARD)
        tags.pack(pady=6)
        for label, color in [("JPG/PNG","#06b6d4"),("DOCX","#3b82f6"),("PPTX","#d97706"),
                              ("TXT/MD","#94a3b8"),("HTML","#f97316"),("CSV/XLSX","#22c55e")]:
            tk.Label(tags, text=label, font=("Segoe UI", 8, "bold"),
                     bg=self._hex_alpha(color, 0.15), fg=color,
                     padx=8, pady=3, relief="flat").pack(side="left", padx=3)

        # File list
        list_container = tk.Frame(self, bg=BG)
        list_container.pack(fill="both", expand=True, padx=30)

        self.canvas = tk.Canvas(list_container, bg=BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.list_frame = tk.Frame(self.canvas, bg=BG)
        self.canvas_window = self.canvas.create_window((0,0), window=self.list_frame, anchor="nw")

        self.list_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.empty_label = tk.Label(self.list_frame,
            text="Henüz dosya eklenmedi", font=("Segoe UI", 10),
            bg=BG, fg=MUTED)
        self.empty_label.pack(pady=30)

        # Bottom buttons
        btn_frame = tk.Frame(self, bg=BG)
        btn_frame.pack(fill="x", padx=30, pady=18)

        self.convert_btn = RoundedButton(btn_frame, "✨  Hepsini Dönüştür",
                                         command=self._convert_all,
                                         bg=ACCENT, width=200, height=42)
        self.convert_btn.pack(side="left")

        self.save_btn = RoundedButton(btn_frame, "📁  Kayıt Klasörü",
                                      command=self._pick_output_dir,
                                      bg=SURFACE, width=160, height=42)
        self.save_btn.pack(side="left", padx=10)

        RoundedButton(btn_frame, "Temizle", command=self._clear_all,
                      bg=SURFACE, width=100, height=42).pack(side="right")

        # Status bar
        self.status_var = tk.StringVar(value="Hazır")
        status_bar = tk.Label(self, textvariable=self.status_var,
                              bg=BG, fg=MUTED, font=("Segoe UI", 8), anchor="w")
        status_bar.pack(fill="x", padx=32, pady=(0,10))

        self.output_dir = str(Path.home() / "Downloads")
        self._update_status()

    def _hex_alpha(self, color, alpha):
        r = int(color[1:3], 16); g = int(color[3:5], 16); b = int(color[5:7], 16)
        rb = int(26); gb = int(26); bb = int(46)  # CARD background
        r2 = int(r*alpha + rb*(1-alpha))
        g2 = int(g*alpha + gb*(1-alpha))
        b2 = int(b*alpha + bb*(1-alpha))
        return f"#{r2:02x}{g2:02x}{b2:02x}"

    def _hover_drop(self, on):
        color = self._hex_alpha(ACCENT, 0.08) if on else CARD
        self.drop_frame.configure(bg=color,
            highlightbackground=ACCENT if on else BORDER)

    def _on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _on_drop(self, event):
        paths = self.tk.splitlist(event.data)
        self._add_files(paths)

    def _pick_files(self):
        exts = " ".join(f"*{e}" for e in sorted(SUPPORTED_EXTS))
        paths = filedialog.askopenfilenames(
            title="Dosya Seç",
            filetypes=[("Desteklenen Dosyalar", exts), ("Tüm Dosyalar", "*.*")]
        )
        if paths:
            self._add_files(paths)

    def _pick_output_dir(self):
        d = filedialog.askdirectory(title="Kayıt Klasörü Seç", initialdir=self.output_dir)
        if d:
            self.output_dir = d
            self._update_status()

    def _add_files(self, paths):
        existing = {r.filepath for r in self.file_rows}
        added = 0
        for p in paths:
            p = p.strip("{}")  # Windows DnD quirk
            if p in existing: continue
            if not os.path.isfile(p): continue
            ext = Path(p).suffix.lower()
            if ext not in SUPPORTED_EXTS:
                self.status_var.set(f"⚠  '{Path(p).name}' desteklenmiyor, atlandı")
                continue
            row = FileRow(self.list_frame, p, self._remove_row)
            row.pack(fill="x", pady=2)
            self.file_rows.append(row)
            existing.add(p)
            added += 1
        if added:
            self.empty_label.pack_forget()
        self._update_status()

    def _remove_row(self, row):
        self.file_rows.remove(row)
        row.destroy()
        if not self.file_rows:
            self.empty_label.pack(pady=30)
        self._update_status()

    def _clear_all(self):
        for row in self.file_rows:
            row.destroy()
        self.file_rows.clear()
        self.empty_label.pack(pady=30)
        self._update_status()

    def _update_status(self):
        n = len(self.file_rows)
        if n == 0:
            self.status_var.set(f"Hazır  •  Kayıt: {self.output_dir}")
        else:
            self.status_var.set(f"{n} dosya  •  Kayıt: {self.output_dir}")

    def _convert_all(self):
        pending = [r for r in self.file_rows if r.status == "wait"]
        if not pending:
            messagebox.showinfo("Bilgi", "Dönüştürülecek dosya yok.")
            return
        self.convert_btn.set_state(False)
        threading.Thread(target=self._run_conversion, args=(pending,), daemon=True).start()

    def _run_conversion(self, rows):
        done = 0
        for row in rows:
            self.after(0, row.set_status, "loading")
            try:
                pdf_data = convert_file(row.filepath)
                out_name = Path(row.filepath).stem + ".pdf"
                out_path = os.path.join(self.output_dir, out_name)
                # Avoid overwrite
                counter = 1
                while os.path.exists(out_path):
                    out_path = os.path.join(self.output_dir,
                        f"{Path(row.filepath).stem}_{counter}.pdf")
                    counter += 1
                with open(out_path, "wb") as f:
                    f.write(pdf_data)
                self.after(0, row.set_status, "done", out_path)
                done += 1
            except Exception as e:
                msg = str(e).split("\n")[0][:80]
                self.after(0, row.set_status, "error", msg)
        self.after(0, self._on_done, done, len(rows))

    def _on_done(self, done, total):
        self.convert_btn.set_state(True)
        self.status_var.set(f"✅ {done}/{total} dosya dönüştürüldü  •  Kayıt: {self.output_dir}")
        if done > 0:
            # Open output folder
            import subprocess, platform
            if platform.system() == "Windows":
                os.startfile(self.output_dir)
            elif platform.system() == "Darwin":
                subprocess.run(["open", self.output_dir])
            else:
                subprocess.run(["xdg-open", self.output_dir])

if __name__ == "__main__":
    app = App()
    app.mainloop()
