import os
import sys
import shutil
import subprocess
from PIL import Image

def main():
    icon_path = "icon.ico"
    logo_path = "Libre Merger.png"
    
    # Check if we should regenerate the icon
    regenerate = False
    if not os.path.exists(icon_path):
        regenerate = True
    elif os.path.exists(logo_path):
        if os.path.getmtime(logo_path) > os.path.getmtime(icon_path):
            regenerate = True
            
    if regenerate and os.path.exists(logo_path):
        print(f"Pre-processing: Syncing {icon_path} with {logo_path}...")
        try:
            img = Image.open(logo_path)
            img.save(icon_path, format="ICO", sizes=[(256, 256), (128, 128), (64, 64), (32, 32)])
        except Exception as e:
            print(f"Warning: Could not sync icon: {e}")
    else:
        print(f"Pre-processing: Using existing {icon_path}")

    # Linux Build Pipeline: Standalone binary + .deb package
    if sys.platform.startswith("linux"):
        import build_deb
        print("\n--------------------------------------------------------------")
        print("Detected Linux / Ubuntu Platform — Launching Debian Build...")
        print("--------------------------------------------------------------\n")
        build_deb.main()
        return

    # Windows Build Pipeline: Core Exe + Setup.exe Wizard
    import PyInstaller.__main__
    sep = os.pathsep

    print("\n--------------------------------------------------------------")
    print("PHASE 1: Building Core Libre Merger Executable...")
    print("--------------------------------------------------------------\n")
    PyInstaller.__main__.run([
        'merger.py',
        '--name=LibreMerger',
        '--onefile',
        '--windowed',
        f'--add-data=Libre Merger.png{sep}.',
        f'--add-data=icon.ico{sep}.',
        '--icon=icon.ico',
        '--log-level=WARN',
        '--clean'
    ])
    
    print("\n--------------------------------------------------------------")
    print("PHASE 2: Building Final Setup.exe Wizard Installer...")
    print("--------------------------------------------------------------\n")
    
    core_exe_path = os.path.join("dist", "LibreMerger.exe")
    if not os.path.exists(core_exe_path):
        print("Fatal Error: Core exe failed to compile.")
        return
        
    PyInstaller.__main__.run([
        'installer.py',
        '--name=Setup',
        '--onefile',
        '--windowed',
        f'--add-data={core_exe_path}{sep}.',
        f'--add-data=Libre Merger.png{sep}.',
        f'--add-data=icon.ico{sep}.',
        '--icon=icon.ico',
        '--log-level=WARN',
        '--clean'
    ])

    print("\n==============================================================")
    print("Build Pipeline Complete! ✔️")
    print("Your Final Application Installer is securely packaged at:")
    final_path = os.path.abspath(os.path.join("dist", "Setup.exe"))
    print(f"{final_path}")
    print("==============================================================")


if __name__ == '__main__':
    main()
