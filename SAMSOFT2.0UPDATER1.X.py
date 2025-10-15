#!/usr/bin/env python3
"""
Samsoft Update Manager - Windows 11 Style
Modern Windows 11 Update interface with original backend functionality
"""

import sys
import os
import ctypes
import subprocess
import threading
import time
import json
import textwrap
import queue
from pathlib import Path
from collections import deque
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, font

# ---------- Auto-elevation ----------
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    params = " ".join([f'"{arg}"' for arg in sys.argv])
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
    sys.exit()

# ---------- Configuration ----------
REPO_DIR = os.path.join(os.getcwd(), "SamsoftRepo")
CONFIG_FILE = os.path.join(REPO_DIR, "config.json")
os.makedirs(REPO_DIR, exist_ok=True)

# Windows 11 Color Palette
W11_COLORS = {
    'bg_primary': '#f3f3f3',
    'bg_secondary': '#ffffff',
    'bg_card': '#fafafa',
    'accent': '#0067c0',
    'accent_hover': '#005a9e',
    'text_primary': '#000000',
    'text_secondary': '#605e5c',
    'border': '#e5e5e5',
    'success': '#107c10',
    'warning': '#f7630c',
    'error': '#d13438'
}

# Default configuration
DEFAULT_CONFIG = {
    "repo_path": REPO_DIR,
    "update_categories": {
        "windows": True,
        "office": True,
        "dotnet": True,
        "vcredist": False
    },
    "auto_reboot": False,
    "dark_mode": False
}

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
        except:
            return DEFAULT_CONFIG
    return DEFAULT_CONFIG

def save_config(config):
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f, indent=4)

# ---------- Windows 11 Style Update Manager ----------
class Windows11UpdateManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Windows Update")
        self.root.geometry("920x700")
        self.root.configure(bg=W11_COLORS['bg_primary'])
        
        # Remove window decorations for modern look (optional)
        # self.root.overrideredirect(True)
        
        self.config = load_config()
        self.repo_path = self.config.get("repo_path", REPO_DIR)
        self.pswindowsupdate_available = False
        
        # State variables
        self.checking_updates = False
        self.installing_updates = False
        self.updates_available = []
        self.last_check_time = "Never"
        
        # Thread control
        self.log_queue = queue.Queue()
        self.ui_update_queue = queue.Queue()
        self.running_threads = []
        self.stop_event = threading.Event()
        
        # Custom fonts (Windows 11 style)
        self.setup_fonts()
        
        # Create UI
        self.create_ui()
        self.start_ui_loop()
        
        # Initial check
        threading.Thread(target=self.check_pswindowsupdate, daemon=True).start()

    def setup_fonts(self):
        """Setup Windows 11-style fonts"""
        self.font_title = font.Font(family="Segoe UI", size=24, weight="normal")
        self.font_heading = font.Font(family="Segoe UI", size=16, weight="normal")
        self.font_body = font.Font(family="Segoe UI", size=11)
        self.font_body_bold = font.Font(family="Segoe UI", size=11, weight="bold")
        self.font_small = font.Font(family="Segoe UI", size=9)

    def create_ui(self):
        """Create Windows 11-style interface"""
        
        # Main container
        main_container = tk.Frame(self.root, bg=W11_COLORS['bg_primary'])
        main_container.pack(fill="both", expand=True)
        
        # Header section
        self.create_header(main_container)
        
        # Scrollable content area
        canvas = tk.Canvas(main_container, bg=W11_COLORS['bg_primary'], 
                          highlightthickness=0, bd=0)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        
        self.scrollable_frame = tk.Frame(canvas, bg=W11_COLORS['bg_primary'])
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True, padx=40, pady=20)
        scrollbar.pack(side="right", fill="y")
        
        # Content sections
        self.create_status_card()
        self.create_update_history_card()
        self.create_advanced_options_card()
        self.create_additional_tools_card()
        
        # Status bar at bottom
        self.create_status_bar()

    def create_header(self, parent):
        """Create Windows 11-style header"""
        header = tk.Frame(parent, bg=W11_COLORS['bg_primary'], height=80)
        header.pack(fill="x", padx=40, pady=(20, 0))
        header.pack_propagate(False)
        
        # Title
        title_label = tk.Label(
            header,
            text="Windows Update",
            font=self.font_title,
            bg=W11_COLORS['bg_primary'],
            fg=W11_COLORS['text_primary']
        )
        title_label.pack(anchor="w", pady=(10, 0))
        
        # Subtitle
        self.subtitle_label = tk.Label(
            header,
            text="You're up to date",
            font=self.font_body,
            bg=W11_COLORS['bg_primary'],
            fg=W11_COLORS['text_secondary']
        )
        self.subtitle_label.pack(anchor="w")

    def create_status_card(self):
        """Main status card - Windows 11 style"""
        card = self.create_card(self.scrollable_frame)
        
        # Status icon and text
        status_frame = tk.Frame(card, bg=W11_COLORS['bg_secondary'])
        status_frame.pack(fill="x", pady=20, padx=20)
        
        # Icon placeholder (you can add actual icon)
        icon_label = tk.Label(
            status_frame,
            text="✓",
            font=("Segoe UI", 32),
            bg=W11_COLORS['bg_secondary'],
            fg=W11_COLORS['success']
        )
        icon_label.pack(side="left", padx=(0, 15))
        
        # Status text
        text_frame = tk.Frame(status_frame, bg=W11_COLORS['bg_secondary'])
        text_frame.pack(side="left", fill="x", expand=True)
        
        self.status_title = tk.Label(
            text_frame,
            text="You're up to date",
            font=self.font_heading,
            bg=W11_COLORS['bg_secondary'],
            fg=W11_COLORS['text_primary'],
            anchor="w"
        )
        self.status_title.pack(fill="x")
        
        self.status_subtitle = tk.Label(
            text_frame,
            text="Last checked: Never",
            font=self.font_body,
            bg=W11_COLORS['bg_secondary'],
            fg=W11_COLORS['text_secondary'],
            anchor="w"
        )
        self.status_subtitle.pack(fill="x")
        
        # Progress bar (hidden by default)
        self.progress_frame = tk.Frame(card, bg=W11_COLORS['bg_secondary'])
        
        self.progress_label = tk.Label(
            self.progress_frame,
            text="Checking for updates...",
            font=self.font_body,
            bg=W11_COLORS['bg_secondary'],
            fg=W11_COLORS['text_secondary']
        )
        self.progress_label.pack(pady=(10, 5), padx=20, anchor="w")
        
        # Custom progress bar
        self.progress_canvas = tk.Canvas(
            self.progress_frame,
            height=4,
            bg=W11_COLORS['border'],
            highlightthickness=0,
            bd=0
        )
        self.progress_canvas.pack(fill="x", padx=20, pady=(0, 20))
        self.progress_bar_rect = None
        self.current_progress = 0
        
        # Buttons
        button_frame = tk.Frame(card, bg=W11_COLORS['bg_secondary'])
        button_frame.pack(fill="x", pady=(0, 20), padx=20)
        
        self.check_button = self.create_accent_button(
            button_frame,
            "Check for updates",
            self.on_check_updates
        )
        self.check_button.pack(side="left", padx=(0, 10))
        
        self.download_button = self.create_secondary_button(
            button_frame,
            "Download to repo",
            self.on_download_updates
        )
        self.download_button.pack(side="left", padx=(0, 10))
        
        self.install_button = self.create_accent_button(
            button_frame,
            "Install updates",
            self.on_install_updates
        )
        self.install_button.pack(side="left")
        self.install_button.pack_forget()  # Hidden initially

    def create_update_history_card(self):
        """Update history card"""
        card = self.create_card(self.scrollable_frame)
        
        # Card title
        title = tk.Label(
            card,
            text="Update history",
            font=self.font_body_bold,
            bg=W11_COLORS['bg_secondary'],
            fg=W11_COLORS['text_primary']
        )
        title.pack(anchor="w", pady=(20, 10), padx=20)
        
        # Log area with custom styling
        log_container = tk.Frame(card, bg=W11_COLORS['bg_card'], 
                                relief="flat", bd=1)
        log_container.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        # Text widget for logs
        self.log_text = tk.Text(
            log_container,
            wrap="word",
            bg=W11_COLORS['bg_card'],
            fg=W11_COLORS['text_primary'],
            font=self.font_small,
            relief="flat",
            bd=0,
            height=8,
            cursor="arrow"
        )
        
        log_scrollbar = ttk.Scrollbar(log_container, orient="vertical", 
                                     command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        log_scrollbar.pack(side="right", fill="y")
        
        self.log_text.config(state="disabled")

    def create_advanced_options_card(self):
        """Advanced options card"""
        card = self.create_card(self.scrollable_frame)
        
        title = tk.Label(
            card,
            text="Advanced options",
            font=self.font_body_bold,
            bg=W11_COLORS['bg_secondary'],
            fg=W11_COLORS['text_primary']
        )
        title.pack(anchor="w", pady=(20, 15), padx=20)
        
        # Options list
        self.create_option_row(card, "Install from offline repo", 
                              self.on_install_offline)
        self.create_option_row(card, "Update Office (Click-to-Run)", 
                              self.on_update_office)
        self.create_option_row(card, "Update .NET Framework", 
                              self.on_update_dotnet)
        self.create_option_row(card, "Update VC++ Redistributables", 
                              self.on_update_vcredist)
        
        # Separator
        sep = tk.Frame(card, bg=W11_COLORS['border'], height=1)
        sep.pack(fill="x", padx=20, pady=10)
        
        # Settings
        self.auto_reboot_var = tk.BooleanVar(value=self.config.get("auto_reboot", False))
        self.create_toggle_row(card, "Automatic restart", self.auto_reboot_var, 
                              self.on_toggle_auto_reboot)
        
        # Change repo path
        self.create_option_row(card, "Change repository path", 
                              self.on_change_repo, pad_bottom=20)

    def create_additional_tools_card(self):
        """Additional tools card"""
        card = self.create_card(self.scrollable_frame)
        
        title = tk.Label(
            card,
            text="Additional tools",
            font=self.font_body_bold,
            bg=W11_COLORS['bg_secondary'],
            fg=W11_COLORS['text_primary']
        )
        title.pack(anchor="w", pady=(20, 10), padx=20)
        
        desc = tk.Label(
            card,
            text="Repository path: " + self.repo_path,
            font=self.font_small,
            bg=W11_COLORS['bg_secondary'],
            fg=W11_COLORS['text_secondary'],
            wraplength=600,
            justify="left"
        )
        desc.pack(anchor="w", padx=20, pady=(0, 20))

    def create_status_bar(self):
        """Bottom status bar"""
        status_bar = tk.Frame(self.root, bg=W11_COLORS['bg_secondary'], 
                             height=40, relief="flat", bd=0)
        status_bar.pack(side="bottom", fill="x")
        status_bar.pack_propagate(False)
        
        self.status_text = tk.Label(
            status_bar,
            text="Ready",
            font=self.font_small,
            bg=W11_COLORS['bg_secondary'],
            fg=W11_COLORS['text_secondary'],
            anchor="w"
        )
        self.status_text.pack(side="left", padx=20)

    def create_card(self, parent):
        """Create a Windows 11-style card"""
        card = tk.Frame(
            parent,
            bg=W11_COLORS['bg_secondary'],
            relief="flat",
            bd=0
        )
        card.pack(fill="x", pady=(0, 20))
        
        # Add subtle border
        border = tk.Frame(card, bg=W11_COLORS['border'], height=1)
        border.pack(fill="x", side="bottom")
        
        return card

    def create_accent_button(self, parent, text, command):
        """Create Windows 11-style accent button"""
        btn = tk.Button(
            parent,
            text=text,
            font=self.font_body,
            bg=W11_COLORS['accent'],
            fg='white',
            activebackground=W11_COLORS['accent_hover'],
            activeforeground='white',
            relief="flat",
            bd=0,
            padx=20,
            pady=8,
            cursor="hand2",
            command=command
        )
        
        # Hover effects
        btn.bind("<Enter>", lambda e: btn.config(bg=W11_COLORS['accent_hover']))
        btn.bind("<Leave>", lambda e: btn.config(bg=W11_COLORS['accent']))
        
        return btn

    def create_secondary_button(self, parent, text, command):
        """Create Windows 11-style secondary button"""
        btn = tk.Button(
            parent,
            text=text,
            font=self.font_body,
            bg=W11_COLORS['bg_card'],
            fg=W11_COLORS['text_primary'],
            activebackground=W11_COLORS['border'],
            activeforeground=W11_COLORS['text_primary'],
            relief="solid",
            bd=1,
            padx=20,
            pady=8,
            cursor="hand2",
            command=command
        )
        
        btn.config(highlightbackground=W11_COLORS['border'], 
                  highlightthickness=1)
        
        # Hover effects
        btn.bind("<Enter>", lambda e: btn.config(bg=W11_COLORS['border']))
        btn.bind("<Leave>", lambda e: btn.config(bg=W11_COLORS['bg_card']))
        
        return btn

    def create_option_row(self, parent, text, command, pad_bottom=0):
        """Create an option row with chevron"""
        row = tk.Frame(parent, bg=W11_COLORS['bg_secondary'], cursor="hand2")
        row.pack(fill="x", padx=20, pady=(0, pad_bottom))
        
        label = tk.Label(
            row,
            text=text,
            font=self.font_body,
            bg=W11_COLORS['bg_secondary'],
            fg=W11_COLORS['text_primary'],
            cursor="hand2"
        )
        label.pack(side="left", pady=12)
        
        chevron = tk.Label(
            row,
            text="›",
            font=("Segoe UI", 14),
            bg=W11_COLORS['bg_secondary'],
            fg=W11_COLORS['text_secondary'],
            cursor="hand2"
        )
        chevron.pack(side="right", pady=12)
        
        # Click handlers
        row.bind("<Button-1>", lambda e: command())
        label.bind("<Button-1>", lambda e: command())
        chevron.bind("<Button-1>", lambda e: command())
        
        # Hover effects
        def on_enter(e):
            row.config(bg=W11_COLORS['bg_card'])
            label.config(bg=W11_COLORS['bg_card'])
            chevron.config(bg=W11_COLORS['bg_card'])
        
        def on_leave(e):
            row.config(bg=W11_COLORS['bg_secondary'])
            label.config(bg=W11_COLORS['bg_secondary'])
            chevron.config(bg=W11_COLORS['bg_secondary'])
        
        row.bind("<Enter>", on_enter)
        row.bind("<Leave>", on_leave)

    def create_toggle_row(self, parent, text, var, command):
        """Create a toggle switch row"""
        row = tk.Frame(parent, bg=W11_COLORS['bg_secondary'])
        row.pack(fill="x", padx=20, pady=12)
        
        label = tk.Label(
            row,
            text=text,
            font=self.font_body,
            bg=W11_COLORS['bg_secondary'],
            fg=W11_COLORS['text_primary']
        )
        label.pack(side="left")
        
        # Toggle switch (using checkbutton styled as toggle)
        toggle = tk.Checkbutton(
            row,
            variable=var,
            bg=W11_COLORS['bg_secondary'],
            activebackground=W11_COLORS['bg_secondary'],
            relief="flat",
            bd=0,
            command=command,
            cursor="hand2"
        )
        toggle.pack(side="right")

    def update_progress(self, value):
        """Update progress bar with smooth animation"""
        self.current_progress = value
        
        if value > 0:
            if not self.progress_frame.winfo_ismapped():
                self.progress_frame.pack(fill="x", after=self.status_subtitle.master.master)
            
            # Animate progress bar
            width = self.progress_canvas.winfo_width()
            if width > 1:
                bar_width = int(width * (value / 100))
                
                if self.progress_bar_rect:
                    self.progress_canvas.delete(self.progress_bar_rect)
                
                self.progress_bar_rect = self.progress_canvas.create_rectangle(
                    0, 0, bar_width, 4,
                    fill=W11_COLORS['accent'],
                    outline=""
                )
        else:
            self.progress_frame.pack_forget()

    def log(self, message, level="info"):
        """Add message to log"""
        self.log_queue.put((message, level))

    def update_log_display(self):
        """Update log text widget"""
        messages = []
        try:
            while not self.log_queue.empty():
                messages.append(self.log_queue.get_nowait())
        except queue.Empty:
            pass
        
        if messages:
            self.log_text.config(state="normal")
            for msg, level in messages:
                timestamp = time.strftime("%H:%M:%S")
                formatted = f"[{timestamp}] {msg}\n"
                self.log_text.insert("end", formatted)
            
            self.log_text.see("end")
            self.log_text.config(state="disabled")

    def start_ui_loop(self):
        """Main UI update loop"""
        self.update_log_display()
        self.root.after(50, self.start_ui_loop)

    def set_status(self, title, subtitle=None, icon="✓", color=None):
        """Update main status display"""
        def update():
            self.status_title.config(text=title)
            if subtitle:
                self.status_subtitle.config(text=subtitle)
            if color:
                icon_label = self.status_title.master.master.winfo_children()[0]
                icon_label.config(text=icon, fg=color)
        
        self.ui_update_queue.put(update)

    def run_async(self, func):
        """Run function in background thread"""
        thread = threading.Thread(target=func, daemon=True)
        self.running_threads.append(thread)
        thread.start()

    # ---------- Event Handlers ----------
    
    def on_check_updates(self):
        """Check for updates"""
        if self.checking_updates:
            return
        self.run_async(self.check_updates)

    def on_download_updates(self):
        """Download updates to repo"""
        self.run_async(self.download_updates)

    def on_install_updates(self):
        """Install updates online"""
        self.run_async(self.install_updates)

    def on_install_offline(self):
        """Install from offline repo"""
        self.run_async(self.install_offline)

    def on_update_office(self):
        """Update Office"""
        self.run_async(self.update_office)

    def on_update_dotnet(self):
        """Update .NET"""
        self.run_async(self.update_dotnet)

    def on_update_vcredist(self):
        """Update VC++ Redistributables"""
        self.run_async(self.update_vcredist)

    def on_toggle_auto_reboot(self):
        """Toggle auto reboot setting"""
        self.config["auto_reboot"] = self.auto_reboot_var.get()
        save_config(self.config)
        status = "enabled" if self.auto_reboot_var.get() else "disabled"
        self.log(f"Automatic restart {status}")

    def on_change_repo(self):
        """Change repository path"""
        new_path = filedialog.askdirectory(
            initialdir=self.repo_path,
            title="Select Repository Directory"
        )
        if new_path:
            self.repo_path = new_path
            self.config["repo_path"] = new_path
            save_config(self.config)
            self.log(f"Repository path changed to: {new_path}")

    # ---------- Backend Functions (Original Logic) ----------
    
    def run_powershell(self, command, capture_output=True):
        """Run PowerShell command"""
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            
            completed = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", 
                 "-WindowStyle", "Hidden", "-Command", command],
                capture_output=capture_output, 
                text=True, 
                timeout=3600,
                startupinfo=startupinfo
            )
            
            return completed.stdout.strip(), completed.stderr.strip(), completed.returncode
        except subprocess.TimeoutExpired:
            return "", "Command timed out", 1
        except Exception as e:
            return "", f"Error: {str(e)}", 1

    def check_pswindowsupdate(self):
        """Check if PSWindowsUpdate module is available"""
        self.log("Checking for PSWindowsUpdate module...")
        check_cmd = "Get-Module -ListAvailable -Name PSWindowsUpdate"
        out, err, code = self.run_powershell(check_cmd)
        
        if not out.strip() and not err:
            self.log("PSWindowsUpdate module not found")
            self.pswindowsupdate_available = False
        else:
            self.pswindowsupdate_available = True
            self.log("PSWindowsUpdate module is available")

    def ensure_module(self):
        """Ensure PSWindowsUpdate module is installed"""
        if self.pswindowsupdate_available:
            return True
        
        self.log("Installing PSWindowsUpdate module...")
        
        # First, try to set PSGallery as trusted
        trust_cmd = "Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue"
        self.run_powershell(trust_cmd)
        
        # Install NuGet provider
        nuget_cmd = "Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue"
        self.run_powershell(nuget_cmd)
        
        # Now install PSWindowsUpdate with improved error handling
        install_cmd = textwrap.dedent("""
            $ErrorActionPreference = 'Stop'
            try {
                if (!(Get-Module -ListAvailable -Name PSWindowsUpdate)) {
                    Install-Module PSWindowsUpdate -Force -Scope AllUsers -AllowClobber
                    Write-Output "PSWindowsUpdate installed successfully"
                } else {
                    Write-Output "PSWindowsUpdate already installed"
                }
            } catch {
                Write-Error $_.Exception.Message
                exit 1
            }
        """)
        
        out, err, code = self.run_powershell(install_cmd)
        
        if out:
            self.log(out)
        
        if code != 0 or (err and "error" in err.lower()):
            self.log(f"Failed to install module: {err if err else 'Unknown error'}", "error")
            return False
        
        self.pswindowsupdate_available = True
        self.log("PSWindowsUpdate module installed successfully")
        return True

    def check_updates(self):
        """Check for Windows updates"""
        self.checking_updates = True
        self.check_button.config(state="disabled")
        
        self.set_status(
            "Checking for updates...",
            "This might take a few minutes",
            "⟳",
            W11_COLORS['accent']
        )
        
        self.progress_label.config(text="Checking for updates...")
        self.log("Checking for updates online...")
        
        for i in range(0, 30, 5):
            self.update_progress(i)
            time.sleep(0.1)
        
        if not self.ensure_module():
            self.update_progress(0)
            self.set_status("Error", "Failed to load update module", "✕", W11_COLORS['error'])
            self.checking_updates = False
            self.check_button.config(state="normal")
            return
        
        # Improved check command with better error handling
        cmd = textwrap.dedent("""
            Import-Module PSWindowsUpdate
            $ErrorActionPreference = 'Continue'
            
            try {
                $updates = Get-WindowsUpdate -MicrosoftUpdate
                if ($updates) {
                    $updates | Select-Object Title, KB, Size, IsDownloaded | ConvertTo-Json
                } else {
                    Write-Output "[]"
                }
            } catch {
                Write-Error $_.Exception.Message
                exit 1
            }
        """)
        
        out, err, code = self.run_powershell(cmd)
        
        for i in range(30, 90, 10):
            self.update_progress(i)
            time.sleep(0.05)
        
        self.last_check_time = time.strftime("%I:%M %p, %B %d, %Y")
        
        if code != 0 or (err and "error" in err.lower() and "0x80240024" not in err):
            self.log(f"Error checking updates: {err}", "error")
            self.set_status("Error checking for updates", 
                          f"Last checked: {self.last_check_time}",
                          "✕", W11_COLORS['error'])
        elif not out.strip() or out.strip() == "[]":
            self.log("Your device is up to date")
            self.set_status("You're up to date", 
                          f"Last checked: {self.last_check_time}",
                          "✓", W11_COLORS['success'])
            self.install_button.pack_forget()
        else:
            try:
                updates = json.loads(out)
                update_count = len(updates) if isinstance(updates, list) else 1
                
                self.log(f"Found {update_count} available updates")
                
                # Log update details
                if isinstance(updates, list):
                    for update in updates[:10]:  # Show first 10
                        if isinstance(update, dict):
                            title = update.get('Title', 'Unknown')
                            kb = update.get('KB', 'N/A')
                            self.log(f"  - {title} (KB{kb})")
                else:
                    self.log(f"Update: {updates.get('Title', 'Unknown')}")
                
                self.set_status(f"{update_count} update{'s' if update_count != 1 else ''} available", 
                              f"Last checked: {self.last_check_time}",
                              "!", W11_COLORS['warning'])
                self.install_button.pack(side="left", after=self.download_button)
                
            except json.JSONDecodeError:
                self.log("Found updates but couldn't parse details")
                self.set_status("Updates available", 
                              f"Last checked: {self.last_check_time}",
                              "!", W11_COLORS['warning'])
                self.install_button.pack(side="left", after=self.download_button)
        
        self.update_progress(100)
        time.sleep(0.3)
        self.update_progress(0)
        
        self.checking_updates = False
        self.check_button.config(state="normal")

    def download_updates(self):
        """Download updates to repository"""
        self.log(f"Downloading updates to {self.repo_path}...")
        self.update_progress(10)
        
        if not self.ensure_module():
            self.update_progress(0)
            return
        
        download_dir = os.path.join(self.repo_path, "Downloads")
        os.makedirs(download_dir, exist_ok=True)
        
        self.update_progress(30)
        self.progress_label.config(text="Downloading updates...")
        
        # Improved download command with proper error handling
        cmd = textwrap.dedent(f"""
            Import-Module PSWindowsUpdate
            $ErrorActionPreference = 'Continue'
            
            try {{
                $updates = Get-WindowsUpdate -MicrosoftUpdate
                if ($updates) {{
                    $updates | ForEach-Object {{
                        Write-Output "Downloading: $($_.Title)"
                    }}
                    
                    # Download updates
                    Get-WindowsUpdate -MicrosoftUpdate -Download -AcceptAll -Verbose
                    Write-Output "Download completed successfully"
                }} else {{
                    Write-Output "No updates available to download"
                }}
            }} catch {{
                Write-Error $_.Exception.Message
                exit 1
            }}
        """)
        
        self.update_progress(50)
        out, err, code = self.run_powershell(cmd)
        
        self.update_progress(90)
        
        if out:
            for line in out.split('\n'):
                if line.strip():
                    self.log(line)
        
        if err and "error" in err.lower():
            self.log(f"Download error: {err}", "error")
        else:
            self.log("Updates downloaded successfully")
            self._create_update_manifest(download_dir)
        
        self.update_progress(100)
        time.sleep(0.3)
        self.update_progress(0)

    def _create_update_manifest(self, download_dir):
        """Create update manifest with async file writing"""
        manifest_path = os.path.join(self.repo_path, "updates_manifest.json")
        cmd = textwrap.dedent("""
            Import-Module PSWindowsUpdate
            Get-WindowsUpdate -MicrosoftUpdate | Select-Object Title, KB, Size, IsDownloaded | ConvertTo-Json
        """)
        
        out, err, code = self.run_powershell(cmd)
        if out and not err and code == 0:
            try:
                updates = json.loads(out)
                # Async file write for performance
                def write_manifest():
                    with open(manifest_path, 'w') as f:
                        json.dump(updates, f, indent=2)
                threading.Thread(target=write_manifest, daemon=True).start()
                self.log("Created update manifest")
            except Exception as e:
                self.log(f"Could not create manifest: {str(e)}")

    def install_updates(self):
        """Install updates online"""
        self.installing_updates = True
        self.log("Installing updates...")
        self.progress_label.config(text="Installing updates...")
        self.update_progress(10)
        
        if not self.ensure_module():
            self.update_progress(0)
            self.installing_updates = False
            return
        
        self.update_progress(30)
        
        check_cmd = """Import-Module PSWindowsUpdate; 
                       $updates = Get-WUList -MicrosoftUpdate; 
                       if ($updates) { $updates | ConvertTo-Json } else { '[]' }"""
        out, err, code = self.run_powershell(check_cmd)
        
        try:
            updates_list = json.loads(out) if out else []
            if not updates_list or (isinstance(updates_list, dict) and not updates_list):
                self.log("No updates available")
                self.update_progress(0)
                self.installing_updates = False
                return
        except Exception as e:
            self.log(f"Failed to check updates: {str(e)}", "error")
            self.update_progress(0)
            self.installing_updates = False
            return
        
        update_count = len(updates_list) if isinstance(updates_list, list) else 1
        self.log(f"Installing {update_count} updates...")
        self.update_progress(50)
        
        # Use correct cmdlet and parameters
        reboot_param = "-AutoReboot" if self.config.get("auto_reboot", False) else "-IgnoreReboot"
        
        # Proper PowerShell command with error handling
        cmd = textwrap.dedent(f"""
            Import-Module PSWindowsUpdate
            $ErrorActionPreference = 'Continue'
            
            try {{
                Get-WindowsUpdate -MicrosoftUpdate -Install -AcceptAll {reboot_param} -Verbose
                Write-Output "Installation completed"
            }} catch {{
                Write-Error $_.Exception.Message
                exit 1
            }}
        """)
        
        self.log("Running Windows Update installation...")
        out, err, code = self.run_powershell(cmd, capture_output=True)
        
        self.update_progress(90)
        
        if out:
            for line in out.split('\n'):
                if line.strip():
                    self.log(line)
        
        if code != 0 or (err and "error" in err.lower()):
            self.log(f"Installation failed: {err if err else 'Unknown error'}", "error")
        else:
            self.log("Updates installed successfully")
        
        self.update_progress(100)
        time.sleep(0.3)
        self.update_progress(0)
        self.installing_updates = False

    def install_offline(self):
        """Install updates from offline repository"""
        self.log(f"Installing from repository: {self.repo_path}...")
        self.progress_label.config(text="Installing offline updates...")
        self.update_progress(10)
        
        download_dir = os.path.join(self.repo_path, "Downloads")
        if not os.path.exists(download_dir) or not os.listdir(download_dir):
            self.log("No updates found in repository", "error")
            self.update_progress(0)
            return
        
        msu_files = [f for f in os.listdir(download_dir) if f.endswith('.msu')]
        
        if not msu_files:
            self.log("No .msu files found", "error")
            self.update_progress(0)
            return
        
        self.log(f"Found {len(msu_files)} update files")
        self.update_progress(30)
        
        success_count = 0
        
        for i, msu_file in enumerate(msu_files):
            if self.stop_event.is_set():
                break
            
            msu_path = os.path.join(download_dir, msu_file)
            self.log(f"Installing {msu_file}...")
            
            try:
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
                
                result = subprocess.run(
                    ["dism", "/online", "/add-package", 
                     f"/packagepath:{msu_path}", "/quiet", "/norestart"],
                    capture_output=True, text=True, timeout=600,
                    startupinfo=startupinfo
                )
                
                if result.returncode == 0:
                    self.log(f"Successfully installed {msu_file}")
                    success_count += 1
                else:
                    self.log(f"Failed to install {msu_file}", "error")
            
            except Exception as e:
                self.log(f"Error installing {msu_file}: {str(e)}", "error")
            
            progress = 30 + ((i + 1) * 60 / len(msu_files))
            self.update_progress(int(progress))
        
        self.log(f"Installed {success_count} of {len(msu_files)} updates")
        
        self.update_progress(100)
        time.sleep(0.3)
        self.update_progress(0)

    def update_office(self):
        """Update Microsoft Office"""
        self.log("Updating Office (Click-to-Run)...")
        self.progress_label.config(text="Updating Office...")
        self.update_progress(30)
        
        possible_paths = [
            r"C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe",
            r"C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
        ]
        
        office_path = None
        for path in possible_paths:
            if os.path.exists(path):
                office_path = path
                break
        
        if not office_path:
            self.log("Office Click-to-Run not found", "error")
            self.update_progress(0)
            return
        
        self.update_progress(60)
        
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            
            result = subprocess.run(
                [office_path, "/update", "user"], 
                capture_output=True, text=True, timeout=1200,
                startupinfo=startupinfo
            )
            
            self.update_progress(90)
            
            if result.returncode == 0:
                self.log("Office updated successfully")
            else:
                self.log("Office update completed with warnings")
        
        except Exception as e:
            self.log(f"Office update error: {str(e)}", "error")
        
        self.update_progress(100)
        time.sleep(0.3)
        self.update_progress(0)

    def update_dotnet(self):
        """Update .NET Framework"""
        if not self.ensure_module():
            return
        
        self.log("Updating .NET Framework...")
        self.progress_label.config(text="Updating .NET Framework...")
        self.update_progress(30)
        
        cmd = textwrap.dedent("""
            Import-Module PSWindowsUpdate
            $ErrorActionPreference = 'Continue'
            
            try {
                $updates = Get-WindowsUpdate -MicrosoftUpdate | Where-Object { $_.Title -like '*.NET*' }
                if ($updates) {
                    Get-WindowsUpdate -MicrosoftUpdate -Install -AcceptAll -IgnoreReboot -Verbose | Where-Object { $_.Title -like '*.NET*' }
                    Write-Output ".NET Framework updates installed"
                } else {
                    Write-Output "No .NET updates available"
                }
            } catch {
                Write-Error $_.Exception.Message
                exit 1
            }
        """)
        
        out, err, code = self.run_powershell(cmd)
        
        self.update_progress(90)
        
        if out:
            self.log(out)
        
        if code == 0:
            self.log(".NET Framework update completed")
        else:
            self.log(f".NET update error: {err if err else 'Unknown error'}", "error")
        
        self.update_progress(100)
        time.sleep(0.3)
        self.update_progress(0)

    def update_vcredist(self):
        """Update Visual C++ Redistributables"""
        self.log("Updating VC++ Redistributables...")
        self.progress_label.config(text="Updating VC++ Redistributables...")
        self.update_progress(30)
        
        try:
            subprocess.run(["where", "winget"], check=True, capture_output=True)
            cmd = "winget upgrade --id Microsoft.VCRedist.* --silent --accept-package-agreements"
        except:
            cmd = textwrap.dedent("""
                $urls = @(
                    'https://aka.ms/vs/17/release/vc_redist.x64.exe',
                    'https://aka.ms/vs/17/release/vc_redist.x86.exe'
                )
                foreach ($url in $urls) {
                    $file = "$env:TEMP\\" + [System.IO.Path]::GetFileName($url)
                    Invoke-WebRequest -Uri $url -OutFile $file
                    Start-Process -Wait -FilePath $file -ArgumentList "/install", "/quiet", "/norestart"
                }
            """)
        
        out, err, code = self.run_powershell(cmd)
        
        self.update_progress(90)
        
        if code == 0:
            self.log("VC++ Redistributables updated")
        else:
            self.log(f"VC++ update error: {err}", "error")
        
        self.update_progress(100)
        time.sleep(0.3)
        self.update_progress(0)

    def cleanup(self):
        """Cleanup on exit"""
        self.stop_event.set()
        for thread in self.running_threads:
            if thread.is_alive():
                thread.join(timeout=1)


# ---------- Main Entry Point ----------
if __name__ == "__main__":
    root = tk.Tk()
    app = Windows11UpdateManager(root)
    
    def on_closing():
        app.cleanup()
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()
