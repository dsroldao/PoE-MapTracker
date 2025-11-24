# PoE Map Tracker   
A simple, lightweight, and automated map tracker overlay for Path of Exile.
It detects which map you are in, tracks the time, deaths, and mechanics encountered (e.g., Delirium, Harvest, etc.) via chat logs.   
## Features   
- **Automated Tracking:** Detects map entry and exit via `Client.txt`.   
- **Excel Export:** Automatically saves runs to an `.xlsx` file with detailed statistics.   
- **Mechanic Detection:** Identifies mechanics (Delirium, Expedition, etc.) based on NPC dialogue lines.   
- **Overlay UI:** Minimalist, "Always on Top" interface.   
- **Safe Zones:** Ignores towns, hideouts, and menagerie.   
   
## Output Files (Data Location)   
The application generates files in the **same directory where the executable is located**:   
1. `map\_history.xlsx`: The Excel spreadsheet containing your run history and statistics.   
2. `config.txt`: Stores the path to your `Client.txt`.   
   
### Where is my spreadsheet?   
- **If running from source:** It will be in the project folder.   
- **If installed via script:** It is located in your local app data folder. You can access it by pasting this into your File Explorer address bar:
`%LOCALAPPDATA%\PoEMapTracker`   
   
## How to Run (Source)   
1. Install Python 3.x.   
2. Install dependencies:   
    ```
    pip install openpyxl
    
    
    ```
3. Run the script:   
    ```
    python tracker.py
    
    
    ```
   
## How to Build (EXE)   
To create a standalone executable for distribution:   
```
pip install pyinstaller
python -m PyInstaller --noconsole --onefile tracker.py


```
The executable will be in the `dist/` folder.   
## Configuration   
The app automatically searches for the PoE `Client.txt` file. If not found, it will ask you to select it manually.   
