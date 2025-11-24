import os
import shutil
import sys
import subprocess

# --- CONFIGURATION ---
APP_NAME = "PoE Map Tracker"
EXE_NAME = "PoE-MapTracker.exe" # New clean filename
SOURCE_EXE = os.path.join("dist", "PoE-MapTracker.exe") # Expected output from PyInstaller

def install():
    print(f"--- INSTALLING {APP_NAME} ---")

    # 1. Verify executable existence
    if not os.path.exists(SOURCE_EXE):
        print(f"[ERROR] File '{SOURCE_EXE}' not found!")
        print("You need to run the build command first:")
        print("python -m PyInstaller --noconsole --onefile PoE-MapTracker.py")
        input("Press ENTER to exit...")
        sys.exit()

    # 2. Define Paths
    install_dir = os.path.join(os.environ["LOCALAPPDATA"], "PoEMapTracker")
    target_exe = os.path.join(install_dir, EXE_NAME)
    
    start_menu = os.path.join(os.environ["APPDATA"], "Microsoft", "Windows", "Start Menu", "Programs")
    shortcut_path = os.path.join(start_menu, f"{APP_NAME}.lnk")

    # 3. Copy Files
    print(f"Installing to: {install_dir}...")
    if not os.path.exists(install_dir):
        os.makedirs(install_dir)

    try:
        shutil.copy2(SOURCE_EXE, target_exe)
        print("Files copied successfully.")
    except Exception as e:
        print(f"[ERROR] Could not copy files: {e}")
        print("Make sure the program is closed before updating.")
        input("Press ENTER to exit...")
        sys.exit()

    # 4. Create Shortcut (PowerShell)
    print("Creating Start Menu shortcut...")
    ps_script = f"""
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("{shortcut_path}")
    $Shortcut.TargetPath = "{target_exe}"
    $Shortcut.WorkingDirectory = "{install_dir}"
    $Shortcut.Description = "Automated Map Tracker for Path of Exile"
    $Shortcut.Save()
    """
    try:
        subprocess.run(["powershell", "-Command", ps_script], check=True)
        print("Shortcut created.")
    except:
        print("Warning: Could not create shortcut automatically.")

    print("\n--- SUCCESS! ---")
    print(f"You can now find '{APP_NAME}' in your Start Menu.")
    input("Press ENTER to close.")

if __name__ == "__main__":
    install()