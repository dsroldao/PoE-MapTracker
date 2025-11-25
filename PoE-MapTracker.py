import tkinter as tk
from tkinter import filedialog, messagebox
import time
import threading
import os
import sys
import json
import urllib.request
import subprocess
import shutil
from datetime import datetime
import ctypes

# Optional dependency check
try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ==========================================
#               CONFIGURATION
# ==========================================

# --- App Info (UPDATE THIS ON EVERY NEW VERSION) ---
APP_VERSION = "1.0.6"
GITHUB_REPO = "dsroldao/PoE-MapTracker" 

# --- Logic Settings ---
SAVE_DELAY = 10  # Seconds to wait before saving/finishing map

# --- Visual Settings ---
THEME = {
    "bg": "#2f3136",
    "title_bg": "#202225",
    "text": "#ffffff",
    "accent": "#d4af37",
    "timer_idle": "#b9bbbe",
    "death": "#ed4245",
    "border": "#202225"
}

FONTS = {
    "main": ("Verdana", 9),
    "timer": ("Verdana", 10, "bold"), 
    "title": ("Verdana", 7, "bold"),
    "badge": ("Verdana", 7, "bold")
}

# --- Game Logic Settings ---
# Safe Zones where the timer should pause
SAFE_ZONES = [
    "Lioneye's Watch", "The Forest Encampment", "The Sarn Encampment", 
    "Highgate", "Overseer's Tower", "The Bridge Encampment", 
    "Oriath", "Karui Shores", "Hideout", "Kingsmarch", 
    "The Rogue Harbour", "The Menagerie", "Mine Encampment", 
    "Aspirant's Plaza", "The Forbidden Sanctum", "Azurite Mine",
    "Monastery of the Keepers"
]

# Mechanics Configuration
# Format: "Trigger Phrase": {"name": "Mechanic Name", "color": "HexColor", "badge": "ShortCode"}
MECHANICS_CONFIG = {
    # Delirium
    "The Strange Voice": {"name": "Delirium",      "color": "#A9A9A9", "badge": "D"},
    "Strange Voice":     {"name": "Delirium",      "color": "#A9A9A9", "badge": "D"},
    
    # Ultimatum
    "The Trialmaster":   {"name": "Ultimatum",     "color": "#FF4444", "badge": "U"},
    "Trialmaster":       {"name": "Ultimatum",     "color": "#FF4444", "badge": "U"},
    
    # Blight
    "Sister Cassia":     {"name": "Blight",        "color": "#FFFACD", "badge": "B"},
    
    # Incursion
    "Alva":              {"name": "Incursion",     "color": "#FFA500", "badge": "A"},
    
    # Beasts
    "Einhar":            {"name": "Beasts",        "color": "#D2691E", "badge": "Ei"},
    
    # Delve
    "Niko":              {"name": "Delve",         "color": "#4488FF", "badge": "N"},
    
    # Betrayal
    "Jun":               {"name": "Betrayal",      "color": "#90EE90", "badge": "Sy"},
    "Interrogate":       {"name": "Betrayal",      "color": "#90EE90", "badge": "Sy"},
    
    # Maven
    "The Envoy":         {"name": "Maven",         "color": "#FF00FF", "badge": "M"},
    "The Maven":         {"name": "Maven",         "color": "#FF00FF", "badge": "M"},
    
    # Expedition
    "Dannig":            {"name": "Expedition",    "color": "#00FFFF", "badge": "E"},
    "Tujen":             {"name": "Expedition",    "color": "#00FFFF", "badge": "E"},
    "Rog":               {"name": "Expedition",     "color": "#00FFFF", "badge": "E"},
    "Gwennen":           {"name": "Expedition",    "color": "#00FFFF", "badge": "E"},
    
    # Sanctum
    "Divinia":           {"name": "Sanctum",       "color": "#FFD700", "badge": "S"},
    
    # Nameless Seer
    "Nameless Seer":     {"name": "Nameless Seer", "color": "#C8A2C8", "badge": "NS"},

    # Eagon (Memory Tear mechanic)
    "Eagon Caeserius":   {"name": "Eagon",         "color": "#9370DB", "badge": "Eg"}, 
    "Eagon":             {"name": "Eagon",         "color": "#9370DB", "badge": "Eg"},

    # Harvest (Oshabi)
    "Oshabi":            {"name": "Harvest",       "color": "#7FFFD4", "badge": "H"},
    "Oshabi, Avatar of the Grove": {"name": "Harvest", "color": "#7FFFD4", "badge": "H"} 
}

class PoEOverlay:
    def __init__(self, root):
        self.root = root
        self._check_dependencies()
        self._init_window()
        self._init_variables()
        self._setup_ui()
        
        # Start background tasks
        self.root.after(500, self._init_log_search)
        
        # Start Update Check (Only if compiled as EXE)
        if getattr(sys, 'frozen', False):
            self.root.after(2000, self._check_for_updates)

    def _check_dependencies(self):
        if not HAS_OPENPYXL:
            messagebox.showerror("Missing Dependency", "Please install openpyxl:\npip install openpyxl")
            sys.exit()

    def _init_window(self):
        title_text = f"PoE Map Tracker v{APP_VERSION}"
        self.root.title(title_text)
        # Initial Width (Standard Mode)
        self.root.geometry("350x40+100+100")
        self.root.overrideredirect(True) 
        self.root.wm_attributes("-topmost", True) 
        self.root.wm_attributes("-alpha", 0.95) 
        self.root.configure(bg=THEME["bg"])
        self.root.config(highlightbackground=THEME["border"], highlightcolor=THEME["border"], highlightthickness=1)

    def _init_variables(self):
        self.status = "idle" 
        self.current_map = "Starting..."
        self.start_time = 0
        self.elapsed = 0
        self.cooldown_start = 0
        self.deaths = 0
        self.mechanics_found = []
        self.tier_var = tk.StringVar(value="16")
        self.running = True
        self.is_compact = False # Variable to control display mode

        if getattr(sys, 'frozen', False):
            self.app_dir = os.path.dirname(sys.executable)
        else:
            self.app_dir = os.path.dirname(os.path.abspath(__file__))
            
        self.config_file = os.path.join(self.app_dir, "config.txt")
        self.excel_file = os.path.join(self.app_dir, "map_history.xlsx")

    # ==========================================
    #                  UI SETUP
    # ==========================================
    def _setup_ui(self):
        self.title_bar = tk.Frame(self.root, bg=THEME["title_bg"], height=16)
        self.title_bar.pack(fill="x", side="top")
        self.title_bar.pack_propagate(False)
        
        self._setup_title_bar_content()
        self._setup_main_content()
        self._make_draggable(self.title_bar)
        self._make_draggable(self.lbl_title)

    def _setup_title_bar_content(self):
        # Store label in self to hide it later in compact mode
        self.lbl_title = tk.Label(self.title_bar, text=f" PoE Map Tracker v{APP_VERSION}", bg=THEME["title_bg"], fg="#b9bbbe", font=FONTS["title"])
        self.lbl_title.pack(side="left", padx=2)
        
        # Right side button container
        btn_container = tk.Frame(self.title_bar, bg=THEME["title_bg"])
        btn_container.pack(side="right", fill="y")

        # Compact Mode Button (M)
        self.btn_mode = tk.Label(btn_container, text=" M ", bg=THEME["title_bg"], fg="#b9bbbe", font=FONTS["title"], cursor="hand2")
        self.btn_mode.pack(side="left", fill="y")
        self.btn_mode.bind("<Button-1>", lambda e: self._toggle_compact_mode())
        self.btn_mode.bind("<Enter>", lambda e: self.btn_mode.config(bg=THEME["accent"], fg=THEME["bg"]))
        self.btn_mode.bind("<Leave>", lambda e: self.btn_mode.config(bg=THEME["title_bg"], fg="#b9bbbe"))

        # Close Button (X)
        btn_close = tk.Label(btn_container, text=" X ", bg=THEME["title_bg"], fg="#b9bbbe", font=FONTS["title"], cursor="hand2")
        btn_close.pack(side="left", fill="y")
        btn_close.bind("<Button-1>", lambda e: sys.exit())
        btn_close.bind("<Enter>", lambda e: btn_close.config(bg=THEME["death"], fg=THEME["text"]))
        btn_close.bind("<Leave>", lambda e: btn_close.config(bg=THEME["title_bg"], fg="#b9bbbe"))

    def _setup_main_content(self):
        self.content_frame = tk.Frame(self.root, bg=THEME["bg"])
        self.content_frame.pack(fill="both", expand=True, padx=2)

        # Elements (Created once, managed by pack/forget)
        self.lbl_map = tk.Label(self.content_frame, text=self.current_map, bg=THEME["bg"], fg=THEME["text"], font=FONTS["main"])
        self.lbl_sep = tk.Label(self.content_frame, text="|", bg=THEME["bg"], fg="#40444b", font=FONTS["main"])
        
        self.lbl_timer = tk.Label(self.content_frame, text="00:00", bg=THEME["bg"], fg=THEME["timer_idle"], font=FONTS["timer"])
        
        self.lbl_deaths = tk.Label(self.content_frame, text="", bg=THEME["bg"], fg=THEME["death"], font=("Verdana", 9, "bold"))
        
        self.frm_mechanics = tk.Frame(self.content_frame, bg=THEME["bg"])
        
        self.tier_frame = tk.Frame(self.content_frame, bg=THEME["bg"])
        tk.Label(self.tier_frame, text="T", bg=THEME["bg"], fg="#b9bbbe", font=FONTS["main"]).pack(side="left")
        self.entry_tier = tk.Entry(self.tier_frame, textvariable=self.tier_var, width=3, bg="#40444b", fg=THEME["text"], borderwidth=0, font=FONTS["main"], justify="center")
        self.entry_tier.pack(side="left", padx=(2, 0))

        # Apply initial layout (Standard)
        self._apply_layout_standard()

    def _toggle_compact_mode(self):
        self.is_compact = not self.is_compact
        if self.is_compact:
            self._apply_layout_compact()
        else:
            self._apply_layout_standard()

    def _apply_layout_standard(self):
        # Remove all to ensure order
        self._unpack_all()
        
        # Restore Title Text
        self.lbl_title.pack(side="left", padx=2)

        # Restore Geometry - Compact Standard (350px)
        self.root.geometry("350x40")
        
        # Add in standard order
        self.lbl_map.pack(side="left", padx=(5, 5))
        self.lbl_sep.pack(side="left")
        self.lbl_timer.pack(side="left", padx=(5, 5))
        self.lbl_deaths.pack(side="left", padx=(0, 5))
        self.frm_mechanics.pack(side="left", padx=(0, 5))
        self.tier_frame.pack(side="right", padx=2)

    def _apply_layout_compact(self):
        # Remove all
        self._unpack_all()
        
        # HIDE Title Text to free up space for buttons
        self.lbl_title.pack_forget()

        # Compact Geometry (Timer Only)
        self.root.geometry("100x40")
        
        # Add timer only, centered
        self.lbl_timer.pack(side="top", pady=2) 

    def _unpack_all(self):
        self.lbl_map.pack_forget()
        self.lbl_sep.pack_forget()
        self.lbl_timer.pack_forget()
        self.lbl_deaths.pack_forget()
        self.frm_mechanics.pack_forget()
        self.tier_frame.pack_forget()

    def _make_draggable(self, widget):
        def start_move(e):
            e.widget.winfo_toplevel().x = e.x
            e.widget.winfo_toplevel().y = e.y
        def do_move(e):
            root = e.widget.winfo_toplevel()
            x = root.winfo_x() + (e.x - root.x)
            y = root.winfo_y() + (e.y - root.y)
            root.geometry(f"+{x}+{y}")
        widget.bind("<Button-1>", start_move)
        widget.bind("<B1-Motion>", do_move)

    # ==========================================
    #            AUTO UPDATER LOGIC
    # ==========================================
    def _check_for_updates(self):
        if "SEU_USUARIO" in GITHUB_REPO: return
        threading.Thread(target=self._fetch_github_version, daemon=True).start()

    def _fetch_github_version(self):
        try:
            url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
            req = urllib.request.Request(url, headers={'User-Agent': 'PoE-Tracker-App'})
            with urllib.request.urlopen(req) as response:
                data = json.loads(response.read().decode())
                latest_tag = data.get("tag_name", "").strip().replace("v", "")
                
                if latest_tag != APP_VERSION:
                    exe_url = next((a["browser_download_url"] for a in data.get("assets", []) if a["name"].endswith(".exe")), None)
                    if exe_url: self.root.after(0, lambda: self._prompt_update(latest_tag, exe_url))
        except: pass

    def _prompt_update(self, version, url):
        msg = f"New version {version} available!\nUpdate now?"
        if messagebox.askyesno("Update", msg):
            self.lbl_map.config(text="Updating...", fg=THEME["accent"])
            threading.Thread(target=self._perform_update, args=(url,), daemon=True).start()

    def _perform_update(self, url):
        try:
            new_file_path = os.path.join(self.app_dir, "new_PoE-MapTracker.exe")
            with urllib.request.urlopen(url) as response, open(new_file_path, 'wb') as out_file:
                shutil.copyfileobj(response, out_file)
            
            current_exe = sys.executable
            batch_script = os.path.join(self.app_dir, "update.bat")
            
            with open(batch_script, "w") as bat:
                bat.write(f'@echo off\ntimeout /t 2 /nobreak > NUL\ndel "{current_exe}"\nmove "{new_file_path}" "{current_exe}"\nstart "" "{current_exe}"\ndel "%~f0"\n')
            subprocess.Popen([batch_script], shell=True)
            self.root.quit()
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Update Failed", str(e)))

    # ==========================================
    #             LOG MONITORING
    # ==========================================
    def _init_log_search(self):
        self.log_path = self._get_log_path()
        if self.log_path and os.path.exists(self.log_path):
            self.current_map = "No Hideout"
            threading.Thread(target=self._logic_loop, daemon=True).start()
            threading.Thread(target=self._monitor_log_file, daemon=True).start()
            self._update_gui_loop()
        else:
            self.lbl_map.config(text="Log not found!", fg=THEME["death"])

    def _get_log_path(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r") as f:
                    saved = f.read().strip()
                if os.path.exists(saved): return saved
            except: pass

        candidates = [
            r"C:\Games\Steam\steamapps\common\Path of Exile\logs\Client.txt",
            r"C:\Program Files (x86)\Grinding Gear Games\Path of Exile\logs\Client.txt",
            r"C:\Program Files (x86)\Steam\steamapps\common\Path of Exile\logs\Client.txt",
            r"D:\SteamLibrary\steamapps\common\Path of Exile\logs\Client.txt",
            os.path.expanduser(r"~\Documents\My Games\Path of Exile\logs\Client.txt")
        ]
        for path in candidates:
            if os.path.exists(path):
                self._save_config(path)
                return path

        messagebox.showinfo("Setup", "Please select 'Client.txt' in your Path of Exile logs folder.")
        selected = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if selected: self._save_config(selected)
        return selected

    def _save_config(self, path):
        try:
            with open(self.config_file, "w") as f: f.write(path)
        except: pass

    def _monitor_log_file(self):
        try:
            f = open(self.log_path, 'r', encoding='utf-8')
            f.seek(0, 2)
            while self.running:
                line = f.readline()
                if not line:
                    time.sleep(0.1)
                    continue
                self._process_log_line(line)
        except Exception as e:
            self.current_map = "Log Error"

    def _process_log_line(self, line):
        # 1. Zone Change
        if " : You have entered " in line:
            try:
                zone = line.split(" : You have entered ")[1].strip().replace(".", "")
                self._handle_zone_change(zone)
            except: pass
        
        # 2. Deaths
        if " : You have been slain." in line and self.status == "running":
            self.deaths += 1

        # 3. Mechanics (Improved Logic)
        if self.status == "running":
            # Ignore player chat lines (@, #, $, %)
            if any(char in line for char in ["@", "#", "$", "%", "&"]):
                return

            # Search triggers in the whole line
            for trigger, data in MECHANICS_CONFIG.items():
                if trigger in line:
                    self._add_mechanic(data)

    def _handle_zone_change(self, zone):
        is_safe = any(s in zone for s in SAFE_ZONES)
        
        if is_safe:
            if self.status == "running": 
                self.status = "cooldown"
                self.cooldown_start = time.time()
        else:
            if self.status == "running": 
                if self.current_map != zone:
                    self._save_to_excel()
                    self._start_run(zone)
            elif self.status == "cooldown":
                self.status = "running"
                self.current_map = zone
            else:
                self._start_run(zone)

    def _start_run(self, zone):
        self.status = "running"
        self.current_map = zone
        self.start_time = time.time()
        self.elapsed = 0
        self.deaths = 0 
        self.mechanics_found = []
        # FIX: Clear UI on Main Thread
        self.root.after(0, self._clear_mechanics_ui)

    def _clear_mechanics_ui(self):
        for widget in self.frm_mechanics.winfo_children():
            widget.destroy()

    def _add_mechanic(self, mech_data):
        name = mech_data["name"]
        if name not in self.mechanics_found:
            self.mechanics_found.append(name)
            # FIX: Draw on Main Thread
            self.root.after(0, lambda: self._draw_mechanic_badge(mech_data))

    def _draw_mechanic_badge(self, data):
        canvas = tk.Canvas(self.frm_mechanics, width=18, height=18, bg=data["color"], highlightthickness=0)
        canvas.pack(side="left", padx=1)
        canvas.create_text(9, 9, text=data["badge"], fill="#000000", font=FONTS["badge"])

    # ==========================================
    #               DATA SAVING
    # ==========================================
    def _save_to_excel(self):
        if self.elapsed < 10: return

        try:
            if not os.path.exists(self.excel_file):
                self._create_excel_file()
            
            wb = load_workbook(self.excel_file)
            self._append_history(wb)
            self._update_statistics(wb)
            wb.save(self.excel_file)
        except: pass

    def _create_excel_file(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "History"
        headers = ["Date", "Time", "Map", "Tier", "Duration", "Deaths", "Mechanics"]
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="2F3136", end_color="2F3136", fill_type="solid")

        wb.create_sheet("Statistics")
        ws_stats = wb["Statistics"]
        headers_stats = ["Map Name", "Count", "Last Tier", "Total Deaths", "", "Mechanic", "Count"]
        ws_stats.append(headers_stats)
        for cell in ws_stats[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="D4AF37", end_color="D4AF37", fill_type="solid")
            if cell.col_idx in [6, 7]: 
                cell.fill = PatternFill(start_color="5865F2", end_color="5865F2", fill_type="solid")
        
        wb.save(self.excel_file)

    def _append_history(self, wb):
        ws = wb["History"]
        now = datetime.now()
        ws.append([
            now.strftime("%d/%m/%Y"),
            now.strftime("%H:%M:%S"),
            self.current_map,
            int(self.tier_var.get()) if self.tier_var.get().isdigit() else self.tier_var.get(),
            self._format_time(self.elapsed),
            self.deaths,
            ", ".join(self.mechanics_found)
        ])

    def _update_statistics(self, wb):
        ws_hist = wb["History"]
        ws_stats = wb["Statistics"]
        ws_stats.delete_rows(2, ws_stats.max_row + 1)
        maps_data = {}
        mechanics_data = {}

        for row in ws_hist.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 4: continue
            
            m_name, m_tier = row[2], row[3]
            m_deaths = row[5] if isinstance(row[5], int) else 0
            if m_name:
                if m_name not in maps_data:
                    maps_data[m_name] = {"count": 0, "tier": m_tier, "deaths": 0}
                maps_data[m_name]["count"] += 1
                maps_data[m_name]["deaths"] += m_deaths
                maps_data[m_name]["tier"] = m_tier

            m_mechs = row[6]
            if m_mechs:
                for mech in m_mechs.split(", "):
                    if mech: mechanics_data[mech] = mechanics_data.get(mech, 0) + 1

        sorted_maps = sorted(maps_data.items(), key=lambda x: x[1]['count'], reverse=True)
        for idx, (name, data) in enumerate(sorted_maps):
            ws_stats.cell(row=idx+2, column=1, value=name)
            ws_stats.cell(row=idx+2, column=2, value=data['count'])
            ws_stats.cell(row=idx+2, column=3, value=data['tier'])
            ws_stats.cell(row=idx+2, column=4, value=data['deaths'])

        sorted_mechs = sorted(mechanics_data.items(), key=lambda x: x[1], reverse=True)
        for idx, (name, count) in enumerate(sorted_mechs):
            ws_stats.cell(row=idx+2, column=6, value=name)
            ws_stats.cell(row=idx+2, column=7, value=count)

        ws_stats.column_dimensions['A'].width = 25
        ws_stats.column_dimensions['F'].width = 15

    def _logic_loop(self):
        while self.running:
            time.sleep(0.1)
            if self.status == "running":
                self.elapsed = time.time() - self.start_time
            elif self.status == "cooldown":
                # Use SAVE_DELAY instead of hardcoded 5
                remain = SAVE_DELAY - (time.time() - self.cooldown_start)
                if remain <= 0:
                    self._save_to_excel()
                    self.status = "idle"
                    self.elapsed = 0
                    self.current_map = "No Hideout"
                    self.deaths = 0
                    self.mechanics_found = []
                    
    def _update_gui_loop(self):
        if self.status == "cooldown":
            # Use SAVE_DELAY instead of hardcoded 5
            remain = int(SAVE_DELAY - (time.time() - self.cooldown_start))
            if self.is_compact:
                self.lbl_map.config(text="") # No text in compact mode
            else:
                self.lbl_map.config(text=f"Saving: {remain}s", fg="orange")
        else:
            if not self.is_compact:
                name = self.current_map
                if len(name) > 22: name = name[:20] + ".."
                self.lbl_map.config(text=name, fg=THEME["text"])
        
        self.lbl_timer.config(text=self._format_time(self.elapsed))
        self.lbl_deaths.config(text=f"☠ {self.deaths}" if self.deaths > 0 else "")
        self.lbl_timer.config(fg=THEME["accent"] if self.status == "running" else THEME["timer_idle"])

        if self.status == "idle" and self.frm_mechanics.winfo_children():
            for widget in self.frm_mechanics.winfo_children(): widget.destroy()

        self.root.after(100, self._update_gui_loop)

    def _format_time(self, s):
        return f"{int(s//60):02d}:{int(s%60):02d}"

if __name__ == "__main__":
    root = tk.Tk()
    PoEOverlay(root)
    root.mainloop()