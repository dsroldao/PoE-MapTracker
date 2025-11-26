# PoE Map Tracker
A modern, lightweight, and automated map tracking overlay for Path of Exile.
Built with Python and CustomTkinter, this tool automatically detects when you enter maps, tracks your time, counts deaths, and identifies encounter mechanics via chat logs, saving everything to a clean Excel spreadsheet.
### Screenshots
Standard mode showing map name, tier, timer and mechanics found
<p align="center">
<img src="https://raw.githubusercontent.com/dsroldao/PoE-MapTracker/refs/heads/main/screenshots/standardmode.png" alt="Standard Mode">

Compact mode, showing only the timer for minimal distraction
<p align="center">
<img src="https://raw.githubusercontent.com/dsroldao/PoE-MapTracker/refs/heads/main/screenshots/compactmode.png" alt="Compact Mode">

###  Key Features
-  **Smart Focus (Overwolf-style):** The overlay intelligently manages its visibility. It stays "Always on Top" while you are playing Path of Exile, but automatically hides when you Alt-Tab to other windows (like Chrome or Discord).
-  **Automated Excel Logging:** Every run is saved to Map History/map_history.xlsx. If the file is open in Excel when a map finishes, the tracker queues the data in memory and retries saving every 5 seconds to prevent data loss.
-  **Modern UI:** Sleek dark theme interface powered by CustomTkinter with rounded corners and Segoe UI typography.
-  **Mechanics Detection:** Automatically detects and logs special encounters based on NPC dialogue:
  Delirium, Ultimatum, Harvest (Oshabi), Blight, Expedition, Sanctum, Eagon/Memory Tear, Nameless Seer, and more.
-  **Compact Mode:** Click the "M" button to toggle between the full dashboard and a tiny, non-intrusive timer.
-  **System Tray Support**: Minimizes to the system tray area.
-  **Auto-Close:** Monitors the game process. If Path of Exile is closed, the tracker gracefully exits after 10 seconds.
###  Installation
#### Option 1: Standalone Executable (Recommended)
1. Go to the [PoE-MapTracker/releases](https://github.com/dsroldao/PoE-MapTracker/releases) page.
2. Download PoE-MapTracker.exe.
3. Place it in a dedicated folder (e.g., Desktop/PoE Tracker).
4. Run the executable. A shortcut will be automatically created in your Start Menu.
#### Option 2: Running from Source
If you prefer to run the Python script directly:
1. Install Python 3.12+.
2. Install the required dependencies:
	```pip install customtkinter openpyxl pystray Pillow```
3. Run the script:
	```python PoE-MapTracker.py ```
###  Building from Source
To compile the project into a standalone .exe (just like the Release version), this project uses Nuitka for better performance and stability compared to PyInstaller.
#### Prerequisites
- Python 3.12+
- Standard C Compiler (MinGW64 recommended for Nuitka)
#### Build Command
Run the following command in your terminal to generate the executable with the icon and all dependencies embedded:
```python -m nuitka --onefile --standalone --enable-plugin=tk-inter --windows-disable-console --include-package-data=customtkinter --include-package=openpyxl --include-package=pystray --include-package=PIL --windows-icon-from-ico=icon.ico --include-data-file=icon.ico=icon.ico --disable-ccache --output-dir=dist --output-filename=PoE-MapTracker.exe PoE-MapTracker.py ```
(Note: Ensure icon.ico is present in the root folder before compiling)
###  Data Location
The application creates the following files in the directory where it is executed:
- **config.txt:** Stores the path to your Client.txt.
- **Map History/:** Folder containing your map_history.xlsx spreadsheet.
###  Troubleshooting
**"The program says 'Log not found!'"**
- Click on the message or restart the app. A file dialog will appear asking you to manually locate your Client.txt file (usually inside Path of Exile/logs).

**"Windows Defender/Smartscreen flagged the file"**
- This is a common false positive for unsigned Python applications compiled with Nuitka/PyInstaller. The project is open source, so you can review the code and build it yourself if you prefer.

**"The overlay minimizes when I click it in-game"**
- Ensure your Path of Exile is running in "Windowed Fullscreen" or "Borderless" mode. Exclusive Fullscreen mode forces other windows to minimize.
