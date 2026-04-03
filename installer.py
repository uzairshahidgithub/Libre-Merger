import os
import sys
import shutil
import subprocess
import threading
import time
from pathlib import Path

try:
    import customtkinter as ctk
    from PIL import Image
except ImportError:
    sys.exit(1)

LIBRE_GREEN = "#1AAC00"
LIBRE_HOVER = "#137D00"
LIBRE_RED   = "#E74C3C"
LIBRE_ORANGE = "#F39C12"
LIBRE_ORANGE_H = "#D68910"

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def create_shortcut(target, shortcut_path):
    ps_script = f"""
    $wshell = New-Object -ComObject WScript.Shell
    $shortcut = $wshell.CreateShortcut('{shortcut_path}')
    $shortcut.TargetPath = '{target}'
    $shortcut.Save()
    """
    try:
        flags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
        subprocess.run(["powershell", "-Command", ps_script], creationflags=flags)
    except Exception as e:
        print(f"Shortcut Error: {e}")


# ----------------------------------------------------------------------
# Shared: Branded Header
# ----------------------------------------------------------------------
def _add_branded_header(container, title_text, subtitle=None):
    """Add a green branded header bar with logo + title to a frame/window."""
    header = ctk.CTkFrame(container, fg_color=LIBRE_GREEN, corner_radius=0, height=58)
    header.pack(fill="x", side="top")
    header.pack_propagate(False)

    img_path = resource_path("Libre Merger.png")
    if os.path.exists(img_path):
        try:
            logo_img = Image.open(img_path)
            logo_ctk = ctk.CTkImage(logo_img, size=(32, 32))
            ctk.CTkLabel(header, text="", image=logo_ctk).pack(side="left", padx=(14, 8), pady=10)
        except Exception:
            pass

    col = ctk.CTkFrame(header, fg_color="transparent")
    col.pack(side="left", pady=8)
    ctk.CTkLabel(col, text=title_text,
                 font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
                 text_color="white").pack(anchor="w")
    if subtitle:
        ctk.CTkLabel(col, text=subtitle,
                     font=ctk.CTkFont(family="Segoe UI", size=10),
                     text_color="#d4f5d4").pack(anchor="w")


# ----------------------------------------------------------------------
# Installer Application
# ----------------------------------------------------------------------
class InstallerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Libre Merger Setup")
        self.geometry("540x420")
        self.resizable(False, False)
        self._logo_ctk = None  # keep reference

        try:
            self.iconbitmap(resource_path("icon.ico"))
        except Exception:
            pass

        self.container = ctk.CTkFrame(self, fg_color=("white", "#1F1F1F"), corner_radius=0)
        self.container.pack(fill="both", expand=True)

        self._preload_logo()
        self.show_splash()

    def _preload_logo(self):
        img_path = resource_path("Libre Merger.png")
        if os.path.exists(img_path):
            try:
                self._logo_img = Image.open(img_path)
            except Exception:
                self._logo_img = None
        else:
            self._logo_img = None

    def clear_container(self):
        for widget in self.container.winfo_children():
            widget.destroy()

    # ----------------------------------------------------------------
    # Page: Splash
    # ----------------------------------------------------------------
    def show_splash(self):
        self.clear_container()

        _add_branded_header(self.container, "Libre Merger Setup", "Installation Wizard")

        body = ctk.CTkFrame(self.container, fg_color="transparent")
        body.pack(fill="both", expand=True)

        if self._logo_img:
            big_logo = ctk.CTkImage(self._logo_img, size=(90, 90))
            self._big_logo_ref = big_logo  # prevent GC
            ctk.CTkLabel(body, text="", image=big_logo).pack(pady=(22, 6))

        ctk.CTkLabel(body, text="Libre Merger",
                     font=("Segoe UI", 26, "bold"), text_color=LIBRE_GREEN).pack()
        ctk.CTkLabel(body, text="Developed by Muhammad Uzair",
                     font=("Segoe UI", 13), text_color="gray").pack(pady=4)

        ctk.CTkLabel(body, text="A free open-source community tool.", text_color="gray").pack(pady=(12, 2))
        ctk.CTkLabel(body,
                     text="💖  Please consider supporting open-source development!",
                     font=("Segoe UI", 12, "italic")).pack(pady=4)

        # Auto-advance after 3.5s
        self.after(3500, self.check_libreoffice)

    # ----------------------------------------------------------------
    # Page: LibreOffice Check
    # ----------------------------------------------------------------
    def check_libreoffice(self):
        self.clear_container()

        _add_branded_header(self.container, "Libre Merger Setup", "System Requirements Check")

        body = ctk.CTkFrame(self.container, fg_color="transparent")
        body.pack(fill="both", expand=True)

        ctk.CTkLabel(body, text="System Requirements",
                     font=("Segoe UI", 20, "bold")).pack(pady=(28, 16))

        has_lo = False
        try:
            flags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            res = subprocess.run(["soffice", "--version"],
                                 capture_output=True, creationflags=flags)
            if res.returncode == 0:
                has_lo = True
        except Exception:
            pass

        if has_lo:
            ctk.CTkLabel(body,
                         text="✅  LibreOffice is verified and installed!\nFull Office document conversion is active.",
                         text_color=LIBRE_GREEN, font=("Segoe UI", 14), justify="center").pack(pady=14)
            ctk.CTkButton(body, text="Continue →", fg_color=LIBRE_GREEN, hover_color=LIBRE_HOVER,
                          font=("Segoe UI", 14, "bold"), corner_radius=8, height=36,
                          command=self.show_config).pack(pady=24)
        else:
            ctk.CTkLabel(body, text="⚠️  LibreOffice Not Detected",
                         text_color=LIBRE_RED, font=("Segoe UI", 17, "bold")).pack(pady=8)
            msg = ("LibreOffice is required to convert Word, Excel, and Text documents.\n"
                   "Download and install it if you need those advanced features.\n\n"
                   "You can still install Libre Merger now for PDF & Image processing.")
            ctk.CTkLabel(body, text=msg, justify="center").pack(pady=8)

            tf = ctk.CTkFrame(body, fg_color="transparent")
            tf.pack(pady=16)
            ctk.CTkButton(tf, text="Exit Setup", fg_color="gray", hover_color="dim gray",
                          corner_radius=8, command=self.destroy).pack(side="left", padx=10)
            ctk.CTkButton(tf, text="Bypass & Install Anyway",
                          fg_color=LIBRE_GREEN, hover_color=LIBRE_HOVER,
                          corner_radius=8, font=("Segoe UI", 13, "bold"),
                          command=self.show_config).pack(side="left", padx=10)

    # ----------------------------------------------------------------
    # Page: Installation Options
    # ----------------------------------------------------------------
    def show_config(self):
        self.clear_container()

        _add_branded_header(self.container, "Libre Merger Setup", "Installation Options")

        body = ctk.CTkFrame(self.container, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=30, pady=14)

        ctk.CTkLabel(body, text="Choose Installation Options",
                     font=("Segoe UI", 18, "bold")).pack(pady=(8, 14))

        default_install_dir = os.path.join(
            os.environ.get("LOCALAPPDATA", "C:\\"), "Programs", "Libre Merger"
        )

        self.path_var = ctk.StringVar(value=default_install_dir)
        self.path_var.trace_add("write", self._on_path_change)

        path_frame = ctk.CTkFrame(body, fg_color="transparent")
        path_frame.pack(fill="x", pady=6)
        ctk.CTkLabel(path_frame, text="Install Location:",
                     font=("Segoe UI", 12, "bold")).pack(anchor="w")
        entry_row = ctk.CTkFrame(path_frame, fg_color="transparent")
        entry_row.pack(fill="x")
        self.path_entry = ctk.CTkEntry(entry_row, textvariable=self.path_var,
                                       font=("Segoe UI", 12))
        self.path_entry.pack(side="left", fill="x", expand=True, pady=4)
        ctk.CTkButton(entry_row, text="Browse", width=70, fg_color="gray",
                      hover_color="dim gray", corner_radius=6,
                      command=self._browse_dir).pack(side="left", padx=(6, 0))

        self.desktop_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(body, text="Create Desktop Shortcut",
                        variable=self.desktop_var, fg_color=LIBRE_GREEN).pack(anchor="w", pady=6)

        self.start_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(body, text="Create Start Menu Shortcut",
                        variable=self.start_var, fg_color=LIBRE_GREEN).pack(anchor="w", pady=4)

        self.install_btn = ctk.CTkButton(
            body, text="Install Now", fg_color=LIBRE_GREEN, hover_color=LIBRE_HOVER,
            font=("Segoe UI", 14, "bold"), height=38, corner_radius=9,
            command=self._check_before_install
        )
        self.install_btn.pack(pady=18)

        # Trigger initial check
        self._on_path_change()

    def _browse_dir(self):
        from tkinter import filedialog
        chosen = filedialog.askdirectory(title="Choose Install Location")
        if chosen:
            self.path_var.set(chosen.replace("/", "\\"))

    def _on_path_change(self, *args):
        """Update button color/text based on whether install already exists."""
        try:
            target = os.path.join(self.path_var.get(), "LibreMerger.exe")
            if os.path.exists(target):
                self.install_btn.configure(
                    text="⚠️  Existing Install Detected — Click to Reinstall",
                    fg_color=LIBRE_ORANGE, hover_color=LIBRE_ORANGE_H
                )
            else:
                self.install_btn.configure(
                    text="Install Now",
                    fg_color=LIBRE_GREEN, hover_color=LIBRE_HOVER
                )
        except Exception:
            pass

    def _check_before_install(self):
        """If an existing install is found, show a dedicated reinstall confirmation page."""
        target = os.path.join(self.path_var.get(), "LibreMerger.exe")
        if os.path.exists(target):
            self.show_reinstall_confirm(self.path_var.get())
        else:
            self.perform_installation(self.path_var.get(), reinstall=False)

    # ----------------------------------------------------------------
    # Page: Reinstall Confirmation
    # ----------------------------------------------------------------
    def show_reinstall_confirm(self, install_dir):
        self.clear_container()

        _add_branded_header(self.container, "Libre Merger Setup", "Existing Installation Detected")

        body = ctk.CTkFrame(self.container, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=30, pady=16)

        ctk.CTkLabel(body, text="⚠️  Existing Installation Found",
                     font=("Segoe UI", 18, "bold"), text_color=LIBRE_ORANGE).pack(pady=(10, 6))
        ctk.CTkLabel(body,
                     text=f"Libre Merger is already installed at:\n{install_dir}\n\n"
                          "Reinstalling will delete the existing files and replace them\n"
                          "with the new version in the same location.",
                     justify="center", font=("Segoe UI", 12)).pack(pady=8)

        btn_frame = ctk.CTkFrame(body, fg_color="transparent")
        btn_frame.pack(pady=18)

        ctk.CTkButton(btn_frame, text="← Cancel", width=130,
                      fg_color="gray", hover_color="dim gray", corner_radius=8,
                      command=self.show_config).pack(side="left", padx=12)

        ctk.CTkButton(btn_frame, text="🔄  Reinstall",
                      width=160, fg_color=LIBRE_ORANGE, hover_color=LIBRE_ORANGE_H,
                      corner_radius=8, font=("Segoe UI", 14, "bold"),
                      command=lambda: self.perform_installation(install_dir, reinstall=True)
                      ).pack(side="left", padx=12)

    # ----------------------------------------------------------------
    # Page: Installing (progress)
    # ----------------------------------------------------------------
    def perform_installation(self, install_dir, reinstall=False):
        self.target_dir = install_dir
        self._reinstall = reinstall
        self._desktop = self.desktop_var.get()
        self._startmenu = self.start_var.get()
        self.clear_container()

        _add_branded_header(self.container, "Libre Merger Setup",
                            "Reinstalling…" if reinstall else "Installing…")

        body = ctk.CTkFrame(self.container, fg_color="transparent")
        body.pack(fill="both", expand=True)

        ctk.CTkLabel(body,
                     text="Reinstalling Libre Merger…" if reinstall else "Installing Libre Merger…",
                     font=("Segoe UI", 20, "bold"),
                     text_color=LIBRE_ORANGE if reinstall else LIBRE_GREEN).pack(pady=(50, 16))

        self.prog = ctk.CTkProgressBar(body, mode="indeterminate",
                                       progress_color=LIBRE_ORANGE if reinstall else LIBRE_GREEN)
        self.prog.pack(fill="x", padx=60)
        self.prog.start()

        threading.Thread(target=self._install_worker, daemon=True).start()

    def _install_worker(self):
        try:
            # For reinstall: delete existing directory contents before copying
            if self._reinstall and os.path.isdir(self.target_dir):
                for item in os.listdir(self.target_dir):
                    item_path = os.path.join(self.target_dir, item)
                    try:
                        if os.path.isfile(item_path) or os.path.islink(item_path):
                            os.remove(item_path)
                        elif os.path.isdir(item_path):
                            shutil.rmtree(item_path)
                    except Exception as e:
                        print(f"Could not remove {item_path}: {e}")

            os.makedirs(self.target_dir, exist_ok=True)

            payload_exe = resource_path("LibreMerger.exe")
            payload_png = resource_path("Libre Merger.png")
            payload_ico = resource_path("icon.ico")

            dest_exe = os.path.join(self.target_dir, "LibreMerger.exe")

            time.sleep(1.5)  # Visual polish for progress bar

            if os.path.exists(payload_exe):
                shutil.copy(payload_exe, dest_exe)
            if os.path.exists(payload_png):
                shutil.copy(payload_png, os.path.join(self.target_dir, "Libre Merger.png"))
            if os.path.exists(payload_ico):
                shutil.copy(payload_ico, os.path.join(self.target_dir, "icon.ico"))

            if self._desktop and os.path.exists(dest_exe):
                desktop_path = os.path.join(os.environ["USERPROFILE"], "Desktop", "Libre Merger.lnk")
                create_shortcut(dest_exe, desktop_path)

            if self._startmenu and os.path.exists(dest_exe):
                start_path = os.path.join(
                    os.environ["APPDATA"],
                    "Microsoft", "Windows", "Start Menu", "Programs", "Libre Merger.lnk"
                )
                create_shortcut(dest_exe, start_path)

            self.after(400, self.show_success)
        except Exception as e:
            print(f"Install failed: {e}")
            self.after(400, self.show_failure)

    # ----------------------------------------------------------------
    # Page: Success
    # ----------------------------------------------------------------
    def show_success(self):
        self.clear_container()

        _add_branded_header(self.container, "Libre Merger Setup", "Installation Complete")

        body = ctk.CTkFrame(self.container, fg_color="transparent")
        body.pack(fill="both", expand=True)

        ctk.CTkLabel(body, text="🎉  Installation Complete!",
                     font=("Segoe UI", 24, "bold"), text_color=LIBRE_GREEN).pack(pady=(36, 10))
        ctk.CTkLabel(body,
                     text="Libre Merger is set up and ready on your system.",
                     font=("Segoe UI", 13)).pack()

        self.launch_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(body, text="Launch Libre Merger now",
                        variable=self.launch_var, fg_color=LIBRE_GREEN).pack(pady=18)

        ctk.CTkButton(body, text="Finish Setup ✓",
                      fg_color=LIBRE_GREEN, hover_color=LIBRE_HOVER,
                      font=("Segoe UI", 14, "bold"), height=38, corner_radius=9,
                      command=self.finish_launch).pack(pady=6)

    # ----------------------------------------------------------------
    # Page: Failure
    # ----------------------------------------------------------------
    def show_failure(self):
        self.clear_container()

        _add_branded_header(self.container, "Libre Merger Setup", "Installation Failed")

        body = ctk.CTkFrame(self.container, fg_color="transparent")
        body.pack(fill="both", expand=True)

        ctk.CTkLabel(body, text="❌  Installation Failed",
                     font=("Segoe UI", 22, "bold"), text_color=LIBRE_RED).pack(pady=(44, 12))
        ctk.CTkLabel(body,
                     text="Something went wrong during the copy phase.\n"
                          "Ensure you have permission to write to that directory.",
                     font=("Segoe UI", 12), justify="center").pack()

        bf = ctk.CTkFrame(body, fg_color="transparent")
        bf.pack(pady=22)
        ctk.CTkButton(bf, text="← Try Again", fg_color=LIBRE_ORANGE, hover_color=LIBRE_ORANGE_H,
                      corner_radius=8, command=self.show_config).pack(side="left", padx=10)
        ctk.CTkButton(bf, text="Close Setup", fg_color="gray", hover_color="dim gray",
                      corner_radius=8, command=self.destroy).pack(side="left", padx=10)

    # ----------------------------------------------------------------
    def finish_launch(self):
        if self.launch_var.get():
            try:
                dest_exe = os.path.join(self.target_dir, "LibreMerger.exe")
                if os.path.exists(dest_exe):
                    os.startfile(dest_exe)
            except Exception:
                pass
        self.destroy()


if __name__ == "__main__":
    app = InstallerApp()
    app.mainloop()
