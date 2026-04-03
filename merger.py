#!/usr/bin/env python3
"""
Libre Merger (Formerly Document Suite Pro)
A complete solution for merging, dividing, converting, encrypting, and watermarking files.
Theme: LibreOffice GitHub UI. Colors: #18A303. Clean White/Gray Mode.
"""

import os
import sys
import tempfile
import shutil
import subprocess
import threading
import tkinter as tk
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    print("Error: PyMuPDF is not installed. Please install it using: pip install PyMuPDF")
    sys.exit(1)

try:
    import customtkinter as ctk
    from tkinter import filedialog
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    print("Error: Required libraries missing.")
    sys.exit(1)

# Global Aesthetic Variables
LIBRE_GREEN = "#1AAC00"
LIBRE_HOVER = "#137D00"
LIBRE_RED   = "#E74C3C"
LIBRE_RED_H = "#C0392B"
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Keep track of all open Toplevel windows for icon refresh
_open_toplevels = []

def register_toplevel(win):
    _open_toplevels.append(win)
    win.bind("<Destroy>", lambda e, w=win: _open_toplevels.remove(w) if w in _open_toplevels else None)

def refresh_all_icons():
    ico = resource_path("icon.ico")
    for win in list(_open_toplevels):
        try:
            win.iconbitmap(ico)
        except Exception:
            pass

# ----------------------------------------------------------------------
# Backend Helpers
# ----------------------------------------------------------------------

def check_libreoffice():
    try:
        subprocess.run(["soffice", "--version"], capture_output=True, check=True)
        return True
    except (subprocess.SubprocessError, FileNotFoundError):
        return False

def is_image(ext):
    return ext in ['jpg', 'jpeg', 'png', 'tiff', 'bmp']

def is_libreoffice_supported(ext):
    return ext in ['docx', 'doc', 'pptx', 'ppt', 'xlsx', 'xls', 'csv', 'txt', 'html']

def convert_image_to_pdf(input_path, output_dir):
    try:
        img = Image.open(input_path)
        if img.mode != 'RGB': img = img.convert('RGB')
        pdf_name = Path(input_path).stem + ".pdf"
        pdf_path = Path(output_dir) / pdf_name
        img.save(str(pdf_path), "PDF", resolution=100.0)
        return str(pdf_path)
    except Exception as e:
        print(f"Error converting image {input_path}: {e}")
        return None

def convert_lo_to_pdf(input_path, output_dir):
    input_path = Path(input_path).resolve()
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    cmd = ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(output_dir), str(input_path)]
    try:
        flags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
        subprocess.run(cmd, check=True, capture_output=True, text=True, creationflags=flags)
    except subprocess.CalledProcessError as e:
        print(f"Error converting {input_path.name}: {e.stderr}")
        return None
    pdf_path = output_dir / (input_path.stem + ".pdf")
    return str(pdf_path) if pdf_path.exists() else None

def parse_page_ranges(pages_str, max_pages):
    if not pages_str or str(pages_str).strip().lower() == "all": return None
    pages = set()
    try:
        for part in str(pages_str).split(','):
            part = part.strip()
            if '-' in part:
                start_str, end_str = part.split('-')
                start_idx = max(0, int(start_str) - 1)
                end_idx = min(max_pages - 1, int(end_str) - 1)
                for i in range(start_idx, end_idx + 1): pages.add(i)
            else:
                idx = int(part) - 1
                if 0 <= idx < max_pages: pages.add(idx)
    except Exception: pass
    return sorted(list(pages)) if pages else None

def apply_watermark(doc, watermark_text):
    for page in doc:
        try:
            p = fitz.Point(page.rect.width / 4, page.rect.height * 3 / 4)
            page.insert_text(p, watermark_text, fontsize=72, fontname="helv", color=(1, 0, 0), fill_opacity=0.3, rotate=-45, overlay=True)
        except Exception as e:
            print(f"Warning: Failed to watermark {page.number}: {e}")

def create_file_icon(ext):
    ext = ext.lower()
    img = Image.new("RGB", (32, 32), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)
    colors = {
        'pdf': '#E5252A', 'docx': '#2B579A', 'doc': '#2B579A',
        'pptx': '#D24726', 'ppt': '#D24726', 'xlsx': '#217346',
        'xls': '#217346', 'csv': '#217346', 'jpg': '#8E44AD',
        'png': '#8E44AD', 'txt': '#34495E', 'html': '#E67E22'
    }
    bg_color = colors.get(ext, '#888888')
    draw.rectangle([0, 0, 32, 32], fill=bg_color)
    try: font = ImageFont.truetype("arialbd.ttf", 10)
    except IOError: font = ImageFont.load_default()
    text = ext.upper()
    try:
        bbox = font.getbbox(text)
        w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
    except AttributeError:
        w, h = draw.textsize(text, font=font)
    draw.text(((32-w)/2, (32-h)/2), text, fill="white", font=font)
    return ctk.CTkImage(light_image=img, dark_image=img, size=(24, 24))

# ----------------------------------------------------------------------
# Shared: Branded Dialog Header
# ----------------------------------------------------------------------
def _add_branded_header(dialog, title_text, subtitle=None):
    """
    Adds a green branded header with logo + title to any CTkToplevel.
    """
    header = ctk.CTkFrame(dialog, fg_color=LIBRE_GREEN, corner_radius=0, height=50)
    header.pack(fill="x", side="top")
    header.pack_propagate(False)

    img_path = resource_path("Libre Merger.png")
    if os.path.exists(img_path):
        try:
            logo_img = Image.open(img_path)
            logo_ctk = ctk.CTkImage(logo_img, size=(28, 28))
            ctk.CTkLabel(header, text="", image=logo_ctk).pack(side="left", padx=(12, 6), pady=10)
        except Exception:
            pass

    title_col_frame = ctk.CTkFrame(header, fg_color="transparent")
    title_col_frame.pack(side="left", pady=8, fill="y")
    ctk.CTkLabel(title_col_frame, text=title_text, font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
                 text_color="white").pack(anchor="w")
    if subtitle:
        ctk.CTkLabel(title_col_frame, text=subtitle, font=ctk.CTkFont(family="Segoe UI", size=10),
                     text_color="#d4f5d4").pack(anchor="w")


def _center_bottom_of_parent(dialog, master):
    """Position dialog at bottom-center of master window."""
    dialog.update_idletasks()
    try:
        mx = master.winfo_x()
        my = master.winfo_y()
        mw = master.winfo_width()
        mh = master.winfo_height()
        w  = dialog.winfo_width()
        h  = dialog.winfo_height()
        x  = mx + (mw // 2) - (w // 2)
        y  = my + mh - h - 30
        # Keep on screen
        screen_h = dialog.winfo_screenheight()
        if y + h > screen_h - 40:
            y = screen_h - h - 60
        dialog.geometry(f"+{x}+{y}")
    except Exception:
        pass


# ----------------------------------------------------------------------
# Custom Message Box
# ----------------------------------------------------------------------
class CustomMessageBox(ctk.CTkToplevel):
    def __init__(self, master, title, message, msg_type="info"):
        super().__init__(master)
        self.title(title)
        self.geometry("460x240")
        self.resizable(False, False)
        # Defer iconbitmap — CTkToplevel needs a tick to fully initialise on Windows
        self.after(50, lambda: self._set_icon())
        self.transient(master)
        self.grab_set()
        register_toplevel(self)

    def _set_icon(self):
        try: self.iconbitmap(resource_path("icon.ico"))
        except: pass

        _add_branded_header(self, "Libre Merger", title)

        color = LIBRE_GREEN if msg_type == "info" else LIBRE_RED
        icon_str = "ℹ️" if msg_type == "info" else "⚠️"

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=15)

        ctk.CTkLabel(body, text=f"{icon_str}  {title}",
                     font=ctk.CTkFont(weight="bold", size=16),
                     text_color=color).pack(pady=(5, 5))
        ctk.CTkLabel(body, text=message,
                     font=("Segoe UI", 13), wraplength=400, justify="center").pack(pady=5)
        ctk.CTkButton(body, text="  OK  ", width=120,
                      corner_radius=8,
                      fg_color=color, hover_color=LIBRE_HOVER if msg_type=="info" else LIBRE_RED_H,
                      font=ctk.CTkFont(size=13, weight="bold"),
                      command=self.destroy).pack(pady=10)

        _center_bottom_of_parent(self, master)


# ----------------------------------------------------------------------
# Page Selection Dialog
# ----------------------------------------------------------------------
class PageSelectionDialog(ctk.CTkToplevel):
    def __init__(self, master, current_value=""):
        super().__init__(master)
        self.title("Select Pages")
        self.geometry("420x230")
        self.resizable(False, False)
        self.result = None
        # Defer icon so Toplevel is fully ready
        self.after(50, lambda: self._set_icon())
        self.transient(master)
        self.grab_set()
        register_toplevel(self)

        _add_branded_header(self, "Libre Merger", "Page Range Selector")

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=12)

        ctk.CTkLabel(body, text="Extract Specific Pages",
                     font=ctk.CTkFont(weight="bold", size=15)).pack(pady=(5, 2))
        ctk.CTkLabel(body, text="e.g. 1-5, 8, 10-12   ·   Leave blank for ALL",
                     text_color="gray", font=("Segoe UI", 11)).pack()

        self.entry = ctk.CTkEntry(body, width=320, height=36,
                                  font=("Segoe UI", 13),
                                  placeholder_text="e.g. 1-3, 7, 10-12")
        self.entry.pack(pady=10)
        if current_value and current_value.lower() != "all":
            self.entry.insert(0, current_value)
        self.entry.focus()
        # Enter key submits
        self.entry.bind("<Return>", lambda e: self.apply())
        self.bind("<Return>", lambda e: self.apply())

        btn_frame = ctk.CTkFrame(body, fg_color="transparent")
        btn_frame.pack(pady=5)
        ctk.CTkButton(btn_frame, text="Cancel", width=110, corner_radius=8,
                      fg_color=("gray80", "gray30"), text_color=("gray10","gray90"),
                      hover_color=("gray70", "gray40"),
                      command=self.destroy).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="Apply", width=110, corner_radius=8,
                      fg_color=LIBRE_GREEN, hover_color=LIBRE_HOVER,
                      font=ctk.CTkFont(weight="bold"),
                      command=self.apply).pack(side="left", padx=10)

        _center_bottom_of_parent(self, master)

    def _set_icon(self):
        try: self.iconbitmap(resource_path("icon.ico"))
        except: pass

    def apply(self):
        self.result = self.entry.get()
        self.destroy()


# ----------------------------------------------------------------------
# Settings Accordion Component
# ----------------------------------------------------------------------
class SettingsAccordion(ctk.CTkFrame):
    def __init__(self, master, app_ref, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.app = app_ref
        self.is_open = False
        self.toggle_btn = ctk.CTkButton(
            self, text="▶  Advanced Formatting Settings", anchor="w",
            fg_color=("gray92", "gray18"), text_color=("gray10", "gray90"),
            hover_color=("gray85", "gray25"), corner_radius=8, height=38,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            command=self.toggle
        )
        self.toggle_btn.pack(fill="x", pady=(5, 0))
        self.content_frame = ctk.CTkFrame(
            self, fg_color=("gray97", "#1C1C1C"), corner_radius=8,
            border_width=1, border_color=("gray82", "#333333")
        )
        self.build_contents()

    def build_contents(self):
        opts_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        opts_frame.pack(fill="both", expand=True, padx=15, pady=15)

        # Left column
        ctk.CTkLabel(opts_frame, text="Watermark:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, sticky="e", pady=8, padx=10)
        self.app.watermark_entry = ctk.CTkEntry(opts_frame, placeholder_text="e.g., CONFIDENTIAL", width=180)
        self.app.watermark_entry.grid(row=0, column=1, sticky="w", pady=8)

        ctk.CTkLabel(opts_frame, text="Password:", font=ctk.CTkFont(weight="bold")).grid(row=1, column=0, sticky="e", pady=8, padx=10)
        self.app.password_entry = ctk.CTkEntry(opts_frame, placeholder_text="Encrypt Output", show="*", width=180)
        self.app.password_entry.grid(row=1, column=1, sticky="w", pady=8)

        self.app.toc_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(opts_frame, text="Generate PDF Bookmarks (Navigation Outline)",
                        variable=self.app.toc_var,
                        fg_color=LIBRE_GREEN, hover_color=LIBRE_HOVER).grid(row=2, column=0, columnspan=2, sticky="w", padx=20, pady=8)

        # Compression slider — animated, fixed-size label to prevent layout shifts
        ctk.CTkLabel(opts_frame, text="Output Size:", font=ctk.CTkFont(weight="bold")).grid(row=3, column=0, sticky="ne", pady=10, padx=10)
        comp_outer = ctk.CTkFrame(opts_frame, fg_color="transparent", width=220)
        comp_outer.grid(row=3, column=1, sticky="nw", pady=8)
        comp_outer.grid_propagate(False)  # hold fixed width so label changes don't shift neighbors

        self.app.compression_var = ctk.IntVar(value=100)

        self._comp_slider = ctk.CTkSlider(
            comp_outer, variable=self.app.compression_var,
            from_=0, to=100, number_of_steps=2,
            button_color=LIBRE_GREEN, progress_color=LIBRE_GREEN,
            button_hover_color=LIBRE_HOVER,
            width=200, height=18,
            command=self._on_slider_change
        )
        self._comp_slider.place(x=0, y=4)

        # Fixed-height label beneath slider — won't resize the grid row
        self._comp_label = ctk.CTkLabel(
            comp_outer, text="High Resolution (No Compression)",
            text_color=LIBRE_GREEN, anchor="w", width=210,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")
        )
        self._comp_label.place(x=0, y=30)

        # Separator column
        opts_frame.grid_columnconfigure(2, weight=0, minsize=20)

        # Right column
        ctk.CTkLabel(opts_frame, text="Sort Override:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=3, sticky="e", pady=8, padx=10)
        self.app.sort_var = ctk.StringVar(value="Manual Order")
        ctk.CTkOptionMenu(opts_frame,
            values=["Manual Order", "Name (A-Z)", "Name (Z-A)", "Date (Oldest first)", "Date (Newest first)"],
            variable=self.app.sort_var, command=self.app.apply_sorting,
            fg_color=LIBRE_GREEN, button_color=LIBRE_GREEN, button_hover_color=LIBRE_HOVER
        ).grid(row=0, column=4, sticky="w", pady=8)

        ctk.CTkLabel(opts_frame, text="Final Format:", font=ctk.CTkFont(weight="bold")).grid(row=1, column=3, sticky="e", pady=8, padx=10)
        self.app.output_format_var = ctk.StringVar(value="PDF")
        ctk.CTkOptionMenu(opts_frame,
            values=["PDF", "DOCX (via LibreOffice)", "Image Sequence Folder"],
            variable=self.app.output_format_var,
            fg_color=LIBRE_GREEN, button_color=LIBRE_GREEN, button_hover_color=LIBRE_HOVER
        ).grid(row=1, column=4, sticky="w", pady=8)

        self.app.open_file_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(opts_frame, text="Open file/folder when finished",
                        variable=self.app.open_file_var,
                        fg_color=LIBRE_GREEN, hover_color=LIBRE_HOVER).grid(row=2, column=3, columnspan=2, sticky="w", padx=20, pady=8)

    def _on_slider_change(self, value):
        v = int(float(value))
        if v >= 80:
            color = LIBRE_GREEN
            label = "High Resolution (No Compression)"
            self._comp_slider.configure(button_color=LIBRE_GREEN, progress_color=LIBRE_GREEN)
        elif v >= 40:
            color = "#F39C12"
            label = "Balanced (Smaller File Size)"
            self._comp_slider.configure(button_color="#F39C12", progress_color="#F39C12")
        else:
            color = LIBRE_RED
            label = "Maximum Compression (AI-Ready)"
            self._comp_slider.configure(button_color=LIBRE_RED, progress_color=LIBRE_RED)
        self._comp_label.configure(text=label, text_color=color)

    def toggle(self):
        self.is_open = not self.is_open
        arrow = "▼" if self.is_open else "▶"
        self.toggle_btn.configure(text=f"{arrow}  Advanced Formatting Settings")
        if self.is_open:
            self.content_frame.pack(fill="x", padx=10, pady=(0, 10))
        else:
            self.content_frame.pack_forget()


# ----------------------------------------------------------------------
# Main Application
# ----------------------------------------------------------------------
class LibreMergerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Libre Merger")
        self.geometry("870x660")
        self.minsize(760, 560)
        # Set icon immediately, then re-apply after idle to ensure taskbar shows it
        try: self.iconbitmap(resource_path("icon.ico"))
        except: pass
        self.after(200, self._reapply_icon)
        self.files_to_merge = []
        self.row_widgets = []
        self._icons_cache = {}
        self.build_header()
        self.build_bottom_actions()
        self.body_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.body_frame.pack(side="top", fill="both", expand=True, padx=20, pady=(10, 0))
        self.build_files_list()
        self.build_settings()

    def _reapply_icon(self):
        try: self.iconbitmap(resource_path("icon.ico"))
        except: pass

    # ------------------------------------------------------------------
    # Header
    # ------------------------------------------------------------------
    def build_header(self):
        header = ctk.CTkFrame(self, fg_color=LIBRE_GREEN, corner_radius=0, height=62)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        img_path = resource_path("Libre Merger.png")
        if os.path.exists(img_path):
            logo_img = Image.open(img_path)
            self.header_logo = ctk.CTkImage(logo_img, size=(34, 34))
            ctk.CTkLabel(header, text="", image=self.header_logo).pack(side="left", padx=(15, 8), pady=10)

        ctk.CTkLabel(header, text="Libre Merger",
                     font=ctk.CTkFont(family="Segoe UI", size=23, weight="bold"),
                     text_color="white").pack(side="left", pady=10)

        # Theme & About buttons (right)
        btn_style = dict(fg_color="transparent", border_width=1, border_color="white",
                         text_color="white", hover_color=LIBRE_HOVER, corner_radius=8, height=30)

        self.theme_btn = ctk.CTkButton(header, text="🌙 Dark", width=80,
                                       command=self.toggle_theme, **btn_style)
        self.theme_btn.pack(side="right", padx=(5, 15), pady=14)

        ctk.CTkButton(header, text="ℹ️ About", width=80,
                      command=self.show_about, **btn_style).pack(side="right", padx=5, pady=14)

        if ctk.get_appearance_mode() == "Dark":
            self.theme_btn.configure(text="☀️ Light")

    def toggle_theme(self):
        # Animate: show spinner frames while theme applies
        self.theme_btn.configure(state="disabled")
        frames = ["◐", "◓", "◑", "◒"]
        going_dark = ctk.get_appearance_mode() == "Light"

        def _spin(i=0):
            if i < len(frames) * 3:  # 3 full rotations
                self.theme_btn.configure(text=frames[i % len(frames)])
                self.after(60, lambda: _spin(i + 1))
            else:
                # Apply the actual theme now
                if going_dark:
                    ctk.set_appearance_mode("Dark")
                    final_text = "☀️ Light"
                else:
                    ctk.set_appearance_mode("Light")
                    final_text = "🌙 Dark"
                self.theme_btn.configure(text=final_text, state="normal")
                # Re-apply icon to main window AND all open children
                try: self.iconbitmap(resource_path("icon.ico"))
                except: pass
                self.after(150, refresh_all_icons)

        _spin()

    # ------------------------------------------------------------------
    # Files List
    # ------------------------------------------------------------------
    def build_files_list(self):
        self.body_frame.grid_columnconfigure(0, weight=1)
        self.body_frame.grid_rowconfigure(1, weight=10)

        top_bar = ctk.CTkFrame(self.body_frame, fg_color="transparent")
        top_bar.grid(row=0, column=0, sticky="ew", pady=(0, 5))

        ctk.CTkButton(top_bar, text="＋  Add Documents",
                      font=ctk.CTkFont(weight="bold"), corner_radius=8,
                      fg_color=LIBRE_GREEN, hover_color=LIBRE_HOVER,
                      command=self.add_files).pack(side="left")
        ctk.CTkButton(top_bar, text="⊘  Clear Queue",
                      fg_color="transparent", border_width=1, corner_radius=8,
                      text_color=("gray10", "gray90"),
                      hover_color=("gray90", "gray25"),
                      command=self.clear_files).pack(side="left", padx=10)

        self.file_count_lbl = ctk.CTkLabel(top_bar, text="0 files selected", text_color="gray")
        self.file_count_lbl.pack(side="right")

        self.scrollable_frame = ctk.CTkScrollableFrame(
            self.body_frame,
            fg_color=("white", "#1E1E1E"),
            border_width=1, border_color=("gray80", "gray25")
        )
        self.scrollable_frame.grid(row=1, column=0, sticky="nsew", pady=5)

    def build_settings(self):
        self.body_frame.grid_rowconfigure(2, weight=0)
        self.accordion = SettingsAccordion(self.body_frame, self)
        self.accordion.grid(row=2, column=0, sticky="ew", pady=(5, 0))

    # ------------------------------------------------------------------
    # Bottom Actions
    # ------------------------------------------------------------------
    def build_bottom_actions(self):
        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(side="bottom", fill="x", padx=20, pady=15)
        bottom.grid_columnconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(bottom, text="Ready", text_color="gray")
        self.status_label.grid(row=0, column=0, sticky="w", padx=5)

        self.progress_bar = ctk.CTkProgressBar(bottom, mode="determinate", progress_color=LIBRE_GREEN)
        self.progress_bar.grid(row=1, column=0, sticky="ew", padx=5, pady=(5, 10))
        self.progress_bar.set(0)

        self.merge_btn = ctk.CTkButton(
            bottom, text="🔀  Merge",
            command=self.start_main_thread,
            font=ctk.CTkFont(weight="bold", size=16),
            height=48, corner_radius=10,
            fg_color=LIBRE_GREEN, hover_color=LIBRE_HOVER
        )
        self.merge_btn.grid(row=0, column=1, rowspan=2, padx=10)

    # ------------------------------------------------------------------
    # About Dialog
    # ------------------------------------------------------------------
    def show_about(self):
        about = ctk.CTkToplevel(self)
        about.title("About Libre Merger")
        about.geometry("460x340")
        about.resizable(False, False)
        about.transient(self)
        about.grab_set()
        # Defer icon — required on Windows for CTkToplevel
        about.after(50, lambda: about.iconbitmap(resource_path("icon.ico")))
        register_toplevel(about)

        _add_branded_header(about, "Libre Merger", "About This Application")

        body = ctk.CTkFrame(about, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=15)

        ctk.CTkLabel(body, text="Libre Merger",
                     font=ctk.CTkFont(size=26, weight="bold"),
                     text_color=LIBRE_GREEN).pack(pady=(10, 5))
        ctk.CTkLabel(body,
                     text="Free & Open Source Community Tool\nDeveloped by Muhammad Uzair",
                     justify="center", font=("Segoe UI", 13)).pack(pady=8)
        ctk.CTkLabel(body,
                     text="Merge • Split • Convert • Encrypt • Watermark",
                     text_color="gray", font=("Segoe UI", 11, "italic")).pack(pady=5)
        ctk.CTkButton(body, text="Close", width=120, corner_radius=8,
                      fg_color=LIBRE_GREEN, hover_color=LIBRE_HOVER,
                      font=ctk.CTkFont(weight="bold"),
                      command=about.destroy).pack(pady=15)

        _center_bottom_of_parent(about, self)

    # ------------------------------------------------------------------
    # File Management
    # ------------------------------------------------------------------
    def get_icon(self, ext):
        if ext not in self._icons_cache:
            self._icons_cache[ext] = create_file_icon(ext)
        return self._icons_cache[ext]

    def add_files(self):
        file_paths = filedialog.askopenfilenames(
            title="Import Documents",
            filetypes=[("All Supported", "*.pdf;*.pptx;*.docx;*.jpg;*.png;*.txt;*.csv;*.xlsx;*.html"),
                       ("All files", "*.*")]
        )
        if file_paths:
            for path in file_paths:
                ext = os.path.splitext(path)[1][1:].lower()
                if ext == 'pdf' or is_libreoffice_supported(ext) or is_image(ext):
                    self.files_to_merge.append({'path': path, 'ext': ext, 'pages': 'all'})
            if self.sort_var.get() != "Manual Order":
                self.apply_sorting(self.sort_var.get())
            else:
                self.refresh_list_ui()

    def clear_files(self):
        self.files_to_merge.clear()
        self.refresh_list_ui()

    def remove_file(self, index):
        self.files_to_merge.pop(index)
        self.refresh_list_ui()

    def set_pages(self, index):
        dialog = PageSelectionDialog(self, current_value=self.files_to_merge[index]['pages'])
        self.wait_window(dialog)
        if dialog.result is not None:
            v = dialog.result.strip().lower()
            self.files_to_merge[index]['pages'] = v if v and v != "all" else "all"
            self.refresh_list_ui()

    def apply_sorting(self, choice):
        if choice == "Manual Order": return
        if choice == "Name (A-Z)":
            self.files_to_merge.sort(key=lambda x: os.path.basename(x['path']).lower())
        elif choice == "Name (Z-A)":
            self.files_to_merge.sort(key=lambda x: os.path.basename(x['path']).lower(), reverse=True)
        elif choice == "Date (Oldest first)":
            self.files_to_merge.sort(key=lambda x: os.path.getmtime(x['path']))
        elif choice == "Date (Newest first)":
            self.files_to_merge.sort(key=lambda x: os.path.getmtime(x['path']), reverse=True)
        self.refresh_list_ui()

    def open_file(self, path):
        if os.name == 'nt': os.startfile(path)

    def move_row(self, index, direction):
        new_idx = index + direction
        if 0 <= new_idx < len(self.files_to_merge):
            self.files_to_merge[index], self.files_to_merge[new_idx] = \
                self.files_to_merge[new_idx], self.files_to_merge[index]
            self.sort_var.set("Manual Order")
            self.refresh_list_ui()

    def refresh_list_ui(self):
        for w in self.row_widgets: w.destroy()
        self.row_widgets.clear()
        self.file_count_lbl.configure(text=f"{len(self.files_to_merge)} items")

        bg_col = ("gray97", "#252525")
        for idx, file_obj in enumerate(self.files_to_merge):
            row_frame = ctk.CTkFrame(
                self.scrollable_frame,
                fg_color=bg_col, corner_radius=6,
                border_width=1, border_color=("gray85", "gray30")
            )
            row_frame.pack(fill="x", pady=3, padx=5)

            ctk.CTkLabel(row_frame, text=f"{idx+1}.",
                         font=ctk.CTkFont(weight="bold"), width=28).pack(side="left", padx=(10, 4), pady=6)
            ctk.CTkLabel(row_frame, text="", image=self.get_icon(file_obj['ext'])).pack(side="left", padx=4, pady=6)

            name_lbl = ctk.CTkLabel(row_frame, text=os.path.basename(file_obj['path']), anchor="w")
            name_lbl.pack(side="left", fill="x", expand=True, padx=5, pady=6)

            # Move buttons
            ctk.CTkButton(row_frame, text="▲", width=26, height=22,
                          corner_radius=5, fg_color=("gray88","gray28"),
                          text_color=("gray30","gray70"),
                          hover_color=("gray80","gray35"),
                          command=lambda i=idx: self.move_row(i, -1)).pack(side="left", padx=2)
            ctk.CTkButton(row_frame, text="▼", width=26, height=22,
                          corner_radius=5, fg_color=("gray88","gray28"),
                          text_color=("gray30","gray70"),
                          hover_color=("gray80","gray35"),
                          command=lambda i=idx: self.move_row(i, 1)).pack(side="left", padx=(2, 8))

            # Pages button — show exact value ("All" or e.g. "1-5"), no emoji icon
            raw_pages = file_obj['pages']
            p_text = "All" if raw_pages == "all" else raw_pages
            if len(p_text) > 12: p_text = p_text[:10] + "…"
            ctk.CTkButton(row_frame, text=p_text, width=80, height=26,
                          corner_radius=6,
                          fg_color=("gray90", "gray25"),
                          text_color=("gray20", "gray85"),
                          border_width=1, border_color=("gray70", "gray45"),
                          hover_color=("gray82", "gray35"),
                          font=ctk.CTkFont(family="Segoe UI", size=12),
                          command=lambda i=idx: self.set_pages(i)).pack(side="left", padx=6, pady=6)

            # ── Refined X / Remove button ──────────────────────────────
            ctk.CTkButton(row_frame, text="✕", width=28, height=28,
                          corner_radius=14,  # fully round
                          fg_color=("gray90", "gray20"),
                          text_color=(LIBRE_RED, "#ff6b6b"),
                          border_width=1, border_color=(LIBRE_RED, "#C0392B"),
                          hover_color=(LIBRE_RED, LIBRE_RED),
                          font=ctk.CTkFont(size=12, weight="bold"),
                          command=lambda i=idx: self.remove_file(i)).pack(side="right", padx=10, pady=6)

            if file_obj['ext'] == 'pdf':
                name_lbl.bind("<Double-Button-1>", lambda e, p=file_obj['path']: self.open_file(p))
            name_lbl.configure(cursor="hand2")
            self.row_widgets.append(row_frame)

    # ------------------------------------------------------------------
    # Processing — NO extra confirmation dialog before Merge
    # ------------------------------------------------------------------
    def start_main_thread(self):
        if not self.files_to_merge:
            CustomMessageBox(self, "No Files Added",
                             "Please add at least one document\nbefore starting the merge.", msg_type="warning")
            return

        out_fmt = self.output_format_var.get()
        if out_fmt == "Image Sequence Folder":
            output_path = filedialog.askdirectory(title="Select Folder to Save Output Images")
            if not output_path: return
        elif out_fmt == "DOCX (via LibreOffice)":
            output_path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word Document", "*.docx")],
                initialfile="merged_document.docx")
            if not output_path: return
        else:
            output_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                initialfile="merged_document.pdf")
            if not output_path: return

        self.merge_btn.configure(state="disabled")
        self.progress_bar.set(0)

        args = (output_path, out_fmt,
                self.watermark_entry.get().strip(),
                self.password_entry.get(),
                self.toc_var.get(),
                self.compression_var.get())
        threading.Thread(target=self.process_documents, args=args, daemon=True).start()

    def update_status(self, text, progress=None):
        self.status_label.configure(text=text)
        if progress is not None: self.progress_bar.set(progress)

    def process_documents(self, output_path, out_fmt, watermark, password, make_toc, compression_level):
        try:
            if any(is_libreoffice_supported(f['ext']) for f in self.files_to_merge) and not check_libreoffice():
                self.update_status("Error: LibreOffice required but not found.", 0)
                self.after(0, lambda: CustomMessageBox(self, "LibreOffice Required",
                    "LibreOffice is required to process Office documents.\nPlease install it and try again.",
                    msg_type="error"))
                return

            with tempfile.TemporaryDirectory() as tmpdir:
                total = len(self.files_to_merge)
                processed_docs = []

                for idx, file_obj in enumerate(self.files_to_merge, 1):
                    src_path = file_obj['path']
                    ext = file_obj['ext']
                    filename = os.path.basename(src_path)
                    self.update_status(f"Converting ({idx}/{total}): {filename}", (idx - 0.5) / (2 * total))

                    pdf_path = src_path
                    if is_image(ext): pdf_path = convert_image_to_pdf(src_path, tmpdir)
                    elif is_libreoffice_supported(ext): pdf_path = convert_lo_to_pdf(src_path, tmpdir)

                    if pdf_path and os.path.exists(pdf_path):
                        processed_docs.append({'pdf_path': pdf_path, 'filename': filename, 'pages_str': file_obj['pages']})
                    else:
                        print(f"Failed to convert {src_path}. Skipping.")

                if not processed_docs:
                    raise Exception("No files could be processed.")

                self.update_status("Merging & Slicing Pages…", 0.6)
                master_doc = fitz.open()
                master_toc = []
                current_page_idx = 0

                for pdoc in processed_docs:
                    src_fitz = fitz.open(pdoc['pdf_path'])
                    num_pages = src_fitz.page_count
                    requested_pages = parse_page_ranges(pdoc['pages_str'], num_pages)

                    if requested_pages is None:
                        master_toc.append([1, pdoc['filename'], current_page_idx + 1])
                        master_doc.insert_pdf(src_fitz)
                        current_page_idx += num_pages
                    else:
                        master_toc.append([1, f"{pdoc['filename']} (Custom Pages)", current_page_idx + 1])
                        for p_idx in requested_pages:
                            master_doc.insert_pdf(src_fitz, from_page=p_idx, to_page=p_idx)
                            current_page_idx += 1
                    src_fitz.close()

                if watermark:
                    self.update_status("Applying Watermark…", 0.8)
                    apply_watermark(master_doc, watermark)

                if make_toc and master_toc:
                    self.update_status("Generating PDF Bookmarks…", 0.85)
                    try:
                        master_doc.set_toc(master_toc)
                    except Exception as e:
                        print(f"Warning: Failed to inject PDF bookmarks: {e}")

                self.update_status(f"Exporting to {out_fmt}…", 0.9)
                save_kwargs = {}
                if password:
                    save_kwargs.update({
                        'encryption': fitz.PDF_ENCRYPT_AES_256,
                        'owner_pw': password,
                        'user_pw': password
                    })

                if out_fmt == "PDF":
                    if compression_level >= 80:
                        master_doc.save(output_path, **save_kwargs)
                    elif compression_level >= 40:
                        master_doc.save(output_path, garbage=3, deflate=True, **save_kwargs)
                    else:
                        master_doc.save(output_path, garbage=4, deflate=True,
                                        deflate_images=True, deflate_fonts=True,
                                        clean=True, **save_kwargs)
                elif out_fmt == "Image Sequence Folder":
                    os.makedirs(output_path, exist_ok=True)
                    for i, page in enumerate(master_doc):
                        page.get_pixmap(dpi=150).save(os.path.join(output_path, f"page_{i+1}.png"))
                elif out_fmt == "DOCX (via LibreOffice)":
                    tmp_pdf_out = os.path.join(tmpdir, "final_merged_temp.pdf")
                    master_doc.save(tmp_pdf_out)
                    flags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                    subprocess.run(
                        ["soffice", "--headless", "--infilter=writer_pdf_import",
                         "--convert-to", "docx", "--outdir",
                         str(Path(output_path).parent), tmp_pdf_out],
                        check=True, creationflags=flags
                    )
                    lo_out = os.path.join(Path(output_path).parent, "final_merged_temp.docx")
                    if os.path.exists(lo_out): shutil.move(lo_out, output_path)

                master_doc.close()
                self.update_status("✅ Success!", 1.0)
                self.after(0, lambda: CustomMessageBox(self, "Merge Complete",
                    f"Processing complete!\nSaved to:\n{output_path}", msg_type="info"))
                if self.open_file_var.get() and os.name == 'nt':
                    os.startfile(output_path)

        except Exception as e:
            self.update_status("An error occurred.", 0)
            self.after(0, lambda: CustomMessageBox(self, "Error",
                f"Failed:\n{str(e)}", msg_type="error"))
        finally:
            self.merge_btn.configure(state="normal")


if __name__ == "__main__":
    app = LibreMergerApp()
    app.mainloop()