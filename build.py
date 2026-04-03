import PyInstaller.__main__
import shutil
import os
from PIL import Image

def main():
    icon_path = "icon.ico"
    logo_path = "Libre Merger.png"
    
    # Check if we should regenerate the icon
    # Condition: icon missing OR logo is newer than icon (updated artwork)
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

    # Phase 1: Compile the raw Libre Merger application
    print("\n--------------------------------------------------------------")
    print("PHASE 1: Building Core Libre Merger Executable...")
    print("--------------------------------------------------------------\n")
    PyInstaller.__main__.run([
        'merger.py',
        '--name=LibreMerger',
        '--onefile',
        '--windowed',
        '--add-data=Libre Merger.png;.',
        '--add-data=icon.ico;.',
        '--icon=icon.ico',
        '--log-level=WARN',
        '--clean'
    ])
    
    # Phase 2: Compile the Installer application carrying the finished Phase-1 Exe as a payload
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
        f'--add-data={core_exe_path};.',
        '--add-data=Libre Merger.png;.',
        '--add-data=icon.ico;.',
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
    for d in ['build', 'dist']:
        if os.path.exists(d): shutil.rmtree(d, ignore_errors=True)
    main()
