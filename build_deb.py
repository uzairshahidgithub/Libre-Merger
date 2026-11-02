#!/usr/bin/env python3
"""
Automated Debian (.deb) Package Builder for Libre Merger on Linux Ubuntu.
Builds standalone binary with PyInstaller and packages it into a standard .deb package.
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path
from PIL import Image

VERSION = "1.0.0"
PACKAGE_NAME = "libre-merger"
ARCH = "amd64"
DEB_FILENAME = f"{PACKAGE_NAME}_{VERSION}_{ARCH}.deb"


def get_installed_size_kb(dir_path):
    total_bytes = 0
    for root, dirs, files in os.walk(dir_path):
        for f in files:
            fp = os.path.join(root, f)
            if not os.path.islink(fp):
                total_bytes += os.path.getsize(fp)
    return max(1, total_bytes // 1024)


def build_binary():
    print("\n--------------------------------------------------------------")
    print("PHASE 1: Building Standalone Linux Distribution via PyInstaller...")
    print("--------------------------------------------------------------\n")

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "merger.py",
        f"--name={PACKAGE_NAME}",
        "--onedir",
        "--windowed",
        "--add-data=Libre Merger.png:.",
        "--add-data=icon.ico:.",
        "--collect-all=PIL",
        "--collect-all=customtkinter",
        "--hidden-import=PIL._tkinter_finder",
        "--log-level=WARN",
        "--clean",
        "-y"
    ]
    subprocess.run(cmd, check=True)
    dist_app_dir = os.path.join("dist", PACKAGE_NAME)
    if not os.path.exists(dist_app_dir):
        raise RuntimeError(f"Build output failed: {dist_app_dir}")
    print(f"Standalone application distribution compiled at: {dist_app_dir}")
    return dist_app_dir


def build_deb_package(dist_app_dir):
    print("\n--------------------------------------------------------------")
    print("PHASE 2: Assembling Ubuntu .deb Package Structure...")
    print("--------------------------------------------------------------\n")

    pkg_root = Path("build") / "deb"
    if pkg_root.exists():
        shutil.rmtree(pkg_root)

    debian_dir = pkg_root / "DEBIAN"
    bin_dir = pkg_root / "usr" / "bin"
    lib_dir = pkg_root / "usr" / "lib" / PACKAGE_NAME
    apps_dir = pkg_root / "usr" / "share" / "applications"
    pixmaps_dir = pkg_root / "usr" / "share" / "pixmaps"
    icons_dir = pkg_root / "usr" / "share" / "icons" / "hicolor" / "256x256" / "apps"
    share_dir = pkg_root / "usr" / "share" / PACKAGE_NAME
    doc_dir = pkg_root / "usr" / "share" / "doc" / PACKAGE_NAME

    for d in [debian_dir, bin_dir, apps_dir, pixmaps_dir, icons_dir, share_dir, doc_dir]:
        d.mkdir(parents=True, exist_ok=True)

    # 1. Install application directory tree to /usr/lib/libre-merger
    shutil.copytree(dist_app_dir, lib_dir)
    target_exec = lib_dir / PACKAGE_NAME
    os.chmod(target_exec, 0o755)

    # 2. Install launcher script in /usr/bin/libre-merger
    launcher_file = bin_dir / PACKAGE_NAME
    launcher_content = f"""#!/bin/sh
exec /usr/lib/{PACKAGE_NAME}/{PACKAGE_NAME} "$@"
"""
    launcher_file.write_text(launcher_content, encoding="utf-8")
    os.chmod(launcher_file, 0o755)

    # 3. Resized App Icons (256x256)
    logo_src = "Libre Merger.png"
    if os.path.exists(logo_src):
        img = Image.open(logo_src)
        icon_256 = img.resize((256, 256), Image.Resampling.LANCZOS)
        icon_256.save(str(pixmaps_dir / f"{PACKAGE_NAME}.png"))
        icon_256.save(str(icons_dir / f"{PACKAGE_NAME}.png"))
        os.chmod(pixmaps_dir / f"{PACKAGE_NAME}.png", 0o644)
        os.chmod(icons_dir / f"{PACKAGE_NAME}.png", 0o644)

    # 3. Desktop Entry File
    desktop_file = apps_dir / f"{PACKAGE_NAME}.desktop"
    desktop_content = f"""[Desktop Entry]
Version=1.0
Type=Application
Name=Libre Merger
GenericName=Document Merger & Converter
Comment=Merge, split, convert, encrypt, and watermark documents offline with LibreOffice
Exec=/usr/bin/{PACKAGE_NAME} %U
Icon={PACKAGE_NAME}
Terminal=false
Categories=Office;Utility;
MimeType=application/pdf;application/vnd.openxmlformats-officedocument.wordprocessingml.document;application/msword;application/vnd.oasis.opendocument.text;application/vnd.openxmlformats-officedocument.presentationml.presentation;application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;image/png;image/jpeg;
Keywords=PDF;Document;Merger;Converter;LibreOffice;Watermark;Encrypt;
StartupNotify=true
StartupWMClass={PACKAGE_NAME}
"""
    desktop_file.write_text(desktop_content, encoding="utf-8")
    os.chmod(desktop_file, 0o644)

    # 4. App Resources & Artwork
    for asset in ["Libre Merger.png", "icon.ico", "README.md"]:
        if os.path.exists(asset):
            shutil.copy2(asset, share_dir / asset)

    # 5. Copyright / License
    copyright_file = doc_dir / "copyright"
    copyright_file.write_text("""Format: https://www.debian.org/doc/packaging-manuals/copyright-format/1.0/
Upstream-Name: Libre Merger
Upstream-Contact: Muhammad Uzair <uzairshahidgithub@users.noreply.github.com>
Source: https://github.com/uzairshahidgithub/Libre-Merger

Files: *
Copyright: 2026 Muhammad Uzair
License: MIT
 Permission is hereby granted, free of charge, to any person obtaining a copy
 of this software and associated documentation files (the "Software"), to deal
 in the Software without restriction, including without limitation the rights
 to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 copies of the Software, and to permit persons to whom the Software is
 furnished to do so, subject to the following conditions:
 .
 The above copyright notice and this permission notice shall be included in all
 copies or substantial portions of the Software.
""", encoding="utf-8")
    os.chmod(copyright_file, 0o644)

    # 6. Control File
    installed_size = get_installed_size_kb(pkg_root / "usr")
    control_file = debian_dir / "control"
    control_content = f"""Package: {PACKAGE_NAME}
Version: {VERSION}
Section: office
Priority: optional
Architecture: {ARCH}
Installed-Size: {installed_size}
Maintainer: Muhammad Uzair <uzairshahidgithub@users.noreply.github.com>
Depends: libreoffice-writer, libreoffice-calc, libreoffice-impress, xdg-utils
Homepage: https://github.com/uzairshahidgithub/Libre-Merger
Description: Free & Open-Source Automated Document Merger
 Libre Merger is a complete, offline document tool powered by LibreOffice
 and PyMuPDF. Merge, split, convert, watermark, and password-protect
 PDFs, Word documents, Excel spreadsheets, PowerPoint presentations,
 and images in any combination with a modern, responsive user interface.
"""
    control_file.write_text(control_content, encoding="utf-8")
    os.chmod(control_file, 0o644)

    # 7. Post-install and post-remove triggers
    postinst_file = debian_dir / "postinst"
    postinst_content = """#!/bin/sh
set -e
if [ "$1" = "configure" ]; then
    if which update-desktop-database >/dev/null 2>&1; then
        update-desktop-database -q /usr/share/applications || true
    fi
    if which gtk-update-icon-cache >/dev/null 2>&1; then
        gtk-update-icon-cache -q -t -f /usr/share/icons/hicolor || true
    fi
fi
exit 0
"""
    postinst_file.write_text(postinst_content, encoding="utf-8")
    os.chmod(postinst_file, 0o755)

    postrm_file = debian_dir / "postrm"
    postrm_content = """#!/bin/sh
set -e
if [ "$1" = "remove" ] || [ "$1" = "purge" ]; then
    if which update-desktop-database >/dev/null 2>&1; then
        update-desktop-database -q /usr/share/applications || true
    fi
    if which gtk-update-icon-cache >/dev/null 2>&1; then
        gtk-update-icon-cache -q -t -f /usr/share/icons/hicolor || true
    fi
fi
exit 0
"""
    postrm_file.write_text(postrm_content, encoding="utf-8")
    os.chmod(postrm_file, 0o755)

    # Build the .deb archive
    print("\n--------------------------------------------------------------")
    print(f"PHASE 3: Packaging {DEB_FILENAME} via dpkg-deb...")
    print("--------------------------------------------------------------\n")
    dist_dir = Path("dist")
    dist_dir.mkdir(parents=True, exist_ok=True)
    deb_out = dist_dir / DEB_FILENAME

    build_cmd = ["dpkg-deb", "--build", "--root-owner-group", str(pkg_root), str(deb_out)]
    subprocess.run(build_cmd, check=True)

    print("\n==============================================================")
    print("Debian Package Build Complete! ✔️")
    print(f"Output: {deb_out.resolve()}")
    print("==============================================================\n")
    return str(deb_out)


def main():
    dist_app_dir = os.path.join("dist", PACKAGE_NAME)
    if not os.path.exists(dist_app_dir) or not os.path.isdir(dist_app_dir):
        dist_app_dir = build_binary()
    deb_path = build_deb_package(dist_app_dir)
    print(f"Successfully generated: {deb_path}")


if __name__ == "__main__":
    main()
