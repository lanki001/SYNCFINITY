import os
import sys
import shutil
import time
import json
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime, timedelta
import queue
import hashlib
import hmac
from PIL import Image
import pystray
from pystray import MenuItem as item
import win32com.client


# License configuration
LICENSE_FILE = "synfinity.lic"
LICENSE_SECRET = "SynFinity_Secret_Key_2025_v1.2"
MAX_USERS = 100
LICENSE_VALIDITY_DAYS = 365

# Character set for short license keys
CHARSET = "23456789ABCDEFGHJKLMNPQRSTUVWXYZ"

# Original configuration
SOURCE_FOLDERS = []
FILE_SERVER_BASE = r"\\192.168.2.110\shares"  
FILE_SERVER_BASE_2 = r"\\192.168.2.110\shares2" 
DEST_FOLDER = ""
DEFAULT_DEST_FOLDER = "BACKUP"  
CONFIG_FILE = "file_sync_config.json"

ALLOWED_EXTENSIONS = {'.pdf', '.docx', '.xlsx', '.pptx', '.jpeg','.txt','.cdr','.jfif','.png', '.doc','.jpg', '.xls'}
MAX_FILE_SIZE_MB = 50
MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024


class LicenseManager:
    """Manages 16-character license key validation"""

    @staticmethod
    def create_startup_shortcut():
        """Create or update startup shortcut for Windows"""
        try:
            # Get startup folder path
            startup_folder = os.path.join(
                os.environ['APPDATA'],
                r'Microsoft\Windows\Start Menu\Programs\Startup'
            )

            # Shortcut path
            shortcut_path = os.path.join(startup_folder, "SyncFinity.lnk")

            # Get current script path
            if getattr(sys, 'frozen', False):
                # Running as compiled executable
                target_path = sys.executable
            else:
                # Running as script
                target_path = os.path.abspath(sys.argv[0])

            # Create shortcut using Windows COM
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(shortcut_path)
            shortcut.TargetPath = target_path
            shortcut.Arguments = "--auto-start"
            shortcut.WorkingDirectory = os.path.dirname(target_path)
            shortcut.Description = "SyncFinity File Sync Application"
            shortcut.IconLocation = target_path
            shortcut.Save()

            return True, shortcut_path

        except Exception as e:
            return False, str(e)

    @staticmethod
    def check_startup_shortcut():
        """Check if startup shortcut exists and is valid"""
        try:
            startup_folder = os.path.join(
                os.environ['APPDATA'],
                r'Microsoft\Windows\Start Menu\Programs\Startup'
            )
            shortcut_path = os.path.join(startup_folder, "SyncFinity.lnk")

            return os.path.exists(shortcut_path), shortcut_path

        except Exception as e:
            return False, None

    @staticmethod
    def remove_startup_shortcut():
        """Remove startup shortcut"""
        try:
            startup_folder = os.path.join(
                os.environ['APPDATA'],
                r'Microsoft\Windows\Start Menu\Programs\Startup'
            )
            shortcut_path = os.path.join(startup_folder, "SyncFinity.lnk")

            if os.path.exists(shortcut_path):
                os.remove(shortcut_path)
                return True, "Shortcut removed successfully"
            else:
                return False, "Shortcut does not exist"

        except Exception as e:
            return False, str(e)
    @staticmethod
    def encode_compact_v2(user_id, expiry_timestamp):
        """
        Encode user_id and expiry timestamp into compact format
        User ID: 1-100 (7 bits)
        Timestamp: Unix timestamp / 3600 (stored as hours since epoch, 25 bits)
        Total: 32 bits = 4 bytes
        """
        if user_id < 1 or user_id > MAX_USERS:
            raise ValueError(f"User ID must be between 1 and {MAX_USERS}")
        
        hours_since_epoch = int(expiry_timestamp / 3600)
        
        if hours_since_epoch > 0x1FFFFFF:
            raise ValueError("Timestamp too far in future")
        
        packed = (user_id << 25) | hours_since_epoch
        data_bytes = packed.to_bytes(4, byteorder='big')
        
        return data_bytes
    
    @staticmethod
    def decode_compact_v2(data_bytes):
        """Decode compact format back to user_id and expiry timestamp"""
        packed = int.from_bytes(data_bytes, byteorder='big')
        
        user_id = (packed >> 25) & 0x7F
        hours_since_epoch = packed & 0x1FFFFFF
        expiry_timestamp = hours_since_epoch * 3600
        
        return user_id, expiry_timestamp
    
    @staticmethod
    def generate_checksum(data_bytes):
        """Generate 3-byte checksum using HMAC"""
        signature = hmac.new(
            LICENSE_SECRET.encode(),
            data_bytes,
            hashlib.sha256
        ).digest()
        return signature[:3]
    
    @staticmethod
    def readable_to_bytes(readable_str):
        """Convert readable string back to bytes"""
        readable_str = readable_str.replace('-', '').replace(' ', '').upper()
        
        bytes_list = []
        for i in range(0, len(readable_str), 2):
            if i + 1 < len(readable_str):
                try:
                    idx1 = CHARSET.index(readable_str[i])
                    idx2 = CHARSET.index(readable_str[i + 1])
                    byte_val = (idx1 + (idx2 << 5)) % 256
                    bytes_list.append(byte_val)
                except ValueError:
                    return None
        
        return bytes(bytes_list) if bytes_list else None
    
    @staticmethod
    def validate_license_key(license_key):
        """
        Validate a 16-character license key
        Returns: (is_valid, user_id, expiry_date, error_message)
        """
        try:
            clean_key = license_key.replace('-', '').replace(' ', '').upper()
            
            if len(clean_key) != 16:
                return False, None, None, "Invalid key length (must be 16 characters)"
            
            full_data = LicenseManager.readable_to_bytes(clean_key)
            if full_data is None or len(full_data) < 7:
                return False, None, None, "Invalid key format"
            
            data_bytes = full_data[:4]
            provided_checksum = full_data[4:7]
            
            expected_checksum = LicenseManager.generate_checksum(data_bytes)
            
            if provided_checksum != expected_checksum:
                return False, None, None, "Invalid key (checksum mismatch)"
            
            user_id, expiry_timestamp = LicenseManager.decode_compact_v2(data_bytes)
            
            if user_id < 1 or user_id > MAX_USERS:
                return False, None, None, f"Invalid user ID"
            
            expiry_date = datetime.fromtimestamp(expiry_timestamp)
            
            now = datetime.now()
            if now > expiry_date:
                days_expired = (now - expiry_date).days
                return False, user_id, expiry_date, f"License expired {days_expired} days ago"
            
            return True, user_id, expiry_date, "Valid license"
            
        except Exception as e:
            return False, None, None, f"License validation error: {str(e)}"
    
    @staticmethod
    def save_license(license_key):
        """Save license key to file"""
        try:
            with open(LICENSE_FILE, 'w') as f:
                f.write(license_key)
            return True
        except Exception as e:
            return False
    
    @staticmethod
    def load_license():
        """Load license key from file"""
        if not os.path.exists(LICENSE_FILE):
            return None
        try:
            with open(LICENSE_FILE, 'r') as f:
                return f.read().strip()
        except:
            return None


class LicenseDialog:
    """Dialog for license activation"""
    
    def __init__(self, parent):
        self.result = None
        self.dialog = tk.Toplevel(parent)
        try:
            icon_path = resource_path("infinity.ico")
            if os.path.exists(icon_path):
                self.dialog.iconbitmap(icon_path)
        except:
            pass
        
        self.dialog.title("SyncFinity - License Activation")
        self.dialog.geometry("500x350")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (250)
        y = (self.dialog.winfo_screenheight() // 2) - (175)
        self.dialog.geometry(f'500x350+{x}+{y}')
        
        self.dialog.protocol("WM_DELETE_WINDOW", self.cancel)
        
        self.setup_ui()
        
    def setup_ui(self):
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        title_label = ttk.Label(main_frame, text="License Activation Required",
                               font=('Segoe UI', 14, 'bold'))
        title_label.pack(pady=(0, 10))
        
        info_text = ("SyncFinity requires a valid license key to operate.\n"
                    "Please enter your 16-character license key below.\n"
                    "Format: XXXX-XXXX-XXXX-XXXX")
        info_label = ttk.Label(main_frame, text=info_text, justify=tk.CENTER)
        info_label.pack(pady=(0, 20))
        
        ttk.Label(main_frame, text="License Key:").pack(anchor=tk.W)
        self.license_entry = ttk.Entry(main_frame, width=60, font=('Courier', 10))
        self.license_entry.pack(fill=tk.X, pady=(5, 10))
        self.license_entry.focus()
        
        helper_label = ttk.Label(main_frame, 
                                text="Hyphens are optional (e.g., A3F29K7L2M4P5N8Q or A3F2-9K7L-2M4P-5N8Q)",
                                font=('Segoe UI', 8), foreground="gray")
        helper_label.pack(pady=(0, 10))
        
        self.status_label = ttk.Label(main_frame, text="", foreground="red")
        self.status_label.pack(pady=(0, 20))
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X)
        
        ttk.Button(btn_frame, text="Activate", 
                  command=self.activate).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(btn_frame, text="Exit", 
                  command=self.cancel).pack(side=tk.RIGHT)
        
        self.license_entry.bind('<Return>', lambda e: self.activate())
    
    def activate(self):
        license_key = self.license_entry.get().strip()
        
        if not license_key:
            self.status_label.config(text="Please enter a license key", foreground="red")
            return
        
        is_valid, user_id, expiry_date, message = LicenseManager.validate_license_key(license_key)
        
        if is_valid:
            if LicenseManager.save_license(license_key):
                self.result = (license_key, user_id, expiry_date)
                self.dialog.destroy()
            else:
                self.status_label.config(text="Failed to save license", foreground="red")
        else:
            self.status_label.config(text=message, foreground="red")
    
    def cancel(self):
        self.result = None
        self.dialog.destroy()
    
    def show(self):
        self.dialog.wait_window()
        return self.result


class FileSyncGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("SyncFinity")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)

        # System tray icon
        self.tray_icon = None
        self.window_visible = True

        # License information
        self.license_key = None
        self.user_id = None
        self.license_expiry = None

        # Check license before proceeding
        if not self.check_license():
            self.root.quit()
            return

        # Configure style
        style = ttk.Style()
        style.theme_use('clam')

        # Variables
        self.source_folders = []
        self.dest_folder = ""
        self.dest_choice = None
        self.sync_thread = None
        self.sync_running = False
        self.log_queue = queue.Queue()
        self.auto_start_enabled = True
        self.current_sync_log = []
        self.next_sync_time = None
        self.sync_interval_hours = 4


        # Colors
        self.bg_color = "#f0f0f0"
        self.accent_color = "#0078d4"
        self.success_color = "#107c10"
        self.warning_color = "#ff8c00"
        self.error_color = "#d13438"

        self.setup_gui()
        self.load_config()
        self.process_log_queue()

        # Create or update startup shortcut on first run
        self.setup_startup_shortcut()

        # Log license info
        days_remaining = (self.license_expiry - datetime.now()).days
        self.log_message(f"Licensed to User #{self.user_id} | Expires: {self.license_expiry.strftime('%Y-%m-%d')} ({days_remaining} days remaining)",
            "SUCCESS")


        # Setup system tray
        self.setup_system_tray()
        
        # Hide window instead of showing it in taskbar
        self.root.withdraw()
        self.window_visible = False
        
        # Auto-start sync after GUI is fully loaded (increased delay for network shares)
        self.root.after(3000, self.auto_start_sync)
    
    def setup_system_tray(self):
        """Setup system tray icon and menu"""
        try:
            # Load icon from file
            icon_path = resource_path("infinity.ico")
            if os.path.exists(icon_path):
                tray_image = Image.open(icon_path)
            else:
                # Create a simple default icon if file not found
                tray_image = Image.new('RGB', (64, 64), color='blue')
            
            # Create menu
            menu = (
                item('Show/Hide Window', self.toggle_window, default=True),
                item('Start Sync', self.start_sync_from_tray, enabled=lambda item: not self.sync_running),
                item('Stop Sync', self.stop_sync_from_tray, enabled=lambda item: self.sync_running),
                pystray.Menu.SEPARATOR,
                pystray.Menu.SEPARATOR,
                item('Exit', self.quit_application)
            )
            
            # Create tray icon
            self.tray_icon = pystray.Icon("SyncFinity", tray_image, "SyncFinity", menu)
            
            # Run tray icon in separate thread
            tray_thread = threading.Thread(target=self.tray_icon.run, daemon=True)
            tray_thread.start()
            
        except Exception as e:
            print(f"Error setting up system tray: {e}")
            # If tray setup fails, show window normally
            self.root.deiconify()
            self.window_visible = True
    
    def toggle_window(self, icon=None, item=None):
        """Show or hide the main window"""
        if self.window_visible:
            self.root.withdraw()
            self.window_visible = False
        else:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
            self.window_visible = True
    
    def start_sync_from_tray(self, icon=None, item=None):
        """Start sync from tray menu"""
        self.root.after(0, self.start_sync)
    
    def stop_sync_from_tray(self, icon=None, item=None):
        """Stop sync from tray menu"""
        self.root.after(0, self.stop_sync)
    
    def manual_sync_from_tray(self, icon=None, item=None):
        """Manual sync from tray menu"""
        self.root.after(0, self.manual_sync)
    
    def show_license_info_from_tray(self, icon=None, item=None):
        """Show license info from tray menu"""
        self.root.after(0, self.show_license_info)
    
    def open_settings_from_tray(self, icon=None, item=None):
        """Open settings from tray menu"""
        self.root.after(0, self.open_settings)
    
    def quit_application(self, icon=None, item=None):
        """Quit the application completely"""
        if self.sync_running:
            self.sync_running = False
        if self.tray_icon:
            self.tray_icon.stop()
        self.root.after(0, self.root.destroy)
    
    def check_license(self):
        """Check and validate license"""
        license_key = LicenseManager.load_license()
        
        if license_key:
            is_valid, user_id, expiry_date, message = LicenseManager.validate_license_key(license_key)
            
            if is_valid:
                self.license_key = license_key
                self.user_id = user_id
                self.license_expiry = expiry_date
                return True
            else:
                messagebox.showerror("Invalid License", 
                                   f"Your license is invalid:\n{message}\n\nPlease enter a valid license key.")
        
        dialog = LicenseDialog(self.root)
        result = dialog.show()
        
        if result:
            self.license_key, self.user_id, self.license_expiry = result
            return True
        else:
            messagebox.showwarning("License Required", 
                                 "SyncFinity requires a valid license to operate.")
            self.root.after(100, self.root.destroy)
            return False

    def setup_startup_shortcut(self):
        """Create or update startup shortcut on initial run"""
        try:
            # Check if shortcut exists
            exists, shortcut_path = LicenseManager.check_startup_shortcut()

            if not exists:
                # Create shortcut
                success, result = LicenseManager.create_startup_shortcut()

                if success:
                    self.log_message(f"Startup shortcut created: {result}", "SUCCESS")
                    self.log_message("SyncFinity will now start automatically with Windows", "INFO")
                else:
                    self.log_message(f"Could not create startup shortcut: {result}", "WARNING")
            else:
                # Shortcut exists, verify and update if needed
                success, result = LicenseManager.create_startup_shortcut()
                if success:
                    self.log_message("Startup shortcut verified/updated", "SUCCESS")

        except Exception as e:
            self.log_message(f"Startup shortcut setup error: {e}", "WARNING")


    def setup_gui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        title_frame = ttk.Frame(main_frame)
        title_frame.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        title_label = ttk.Label(title_frame, text="SyncFinity",
                               font=('Segoe UI', 16, 'bold'))
        title_label.pack()
        
        self.setup_source_section(main_frame)
        self.setup_destination_section(main_frame)
        self.setup_control_section(main_frame)
        self.setup_log_section(main_frame)
        self.setup_status_bar(main_frame)
    
    def setup_source_section(self, parent):
        source_frame = ttk.LabelFrame(parent, text="Source Folders", padding="10")
        source_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        source_frame.columnconfigure(0, weight=1)
        
        list_frame = ttk.Frame(source_frame)
        list_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        list_frame.columnconfigure(0, weight=1)
        
        self.source_listbox = tk.Listbox(list_frame, height=4, selectmode=tk.EXTENDED)
        self.source_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        source_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.source_listbox.yview)
        source_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.source_listbox.config(yscrollcommand=source_scrollbar.set)
        
        source_btn_frame = ttk.Frame(source_frame)
        source_btn_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        ttk.Button(source_btn_frame, text="Add Folder", 
                  command=self.add_source_folder).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(source_btn_frame, text="Remove Selected", 
                  command=self.remove_source_folder).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(source_btn_frame, text="Auto-Select (Documents/Desktop)", 
                  command=self.auto_select_sources).pack(side=tk.LEFT)
    
    def setup_destination_section(self, parent):
        dest_frame = ttk.LabelFrame(parent, text="Destination", padding="10")
        dest_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        dest_frame.columnconfigure(1, weight=1)
        
        ttk.Label(dest_frame, text="Server:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.server_var = tk.StringVar()
        server_combo = ttk.Combobox(dest_frame, textvariable=self.server_var, 
                                   values=["Primary Server (shares)", "Backup Server (shares2)"],
                                   state="readonly")
        server_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        server_combo.bind('<<ComboboxSelected>>', self.on_server_change)
        
        ttk.Button(dest_frame, text="Browse Folders", 
                  command=self.browse_destination).grid(row=0, column=2)
        
        ttk.Label(dest_frame, text="Path:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(10, 0))
        self.dest_label = ttk.Label(dest_frame, text="No destination selected", 
                                   foreground="gray", relief=tk.SUNKEN, padding="5")
        self.dest_label.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
    
    def setup_control_section(self, parent):
        control_frame = ttk.Frame(parent)
        control_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.start_btn = ttk.Button(control_frame, text="Start Sync Monitor", 
                                   command=self.start_sync, style="Accent.TButton")
        self.start_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.stop_btn = ttk.Button(control_frame, text="Stop Sync", 
                                  command=self.stop_sync, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.manual_sync_btn = ttk.Button(control_frame, text="Manual Sync", 
                                         command=self.manual_sync)
        self.manual_sync_btn.pack(side=tk.LEFT, padx=(0, 20))
        
        ttk.Button(control_frame, text="Settings", 
                  command=self.open_settings).pack(side=tk.RIGHT)
        ttk.Button(control_frame, text="License Info", 
                  command=self.show_license_info).pack(side=tk.RIGHT, padx=(0, 5))
    
    def setup_log_section(self, parent):
        log_frame = ttk.LabelFrame(parent, text="Sync Log", padding="10")
        log_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        log_btn_frame = ttk.Frame(log_frame)
        log_btn_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        ttk.Button(log_btn_frame, text="Save Log", 
                  command=self.save_log).pack(side=tk.LEFT)
    
    def setup_status_bar(self, parent):
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(parent, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        style = ttk.Style()
        style.configure("Accent.TButton", foreground="white")
        try:
            style.map("Accent.TButton", 
                     background=[('active', self.accent_color), ('!active', self.accent_color)])
        except:
            pass
    
    def show_license_info(self):
        """Show license information dialog"""
        days_remaining = (self.license_expiry - datetime.now()).days
        
        info_text = (
            f"License Information:\n\n"
            f"User ID: #{self.user_id}\n"
            f"License Key: {self.license_key}\n"
            f"Expiry Date: {self.license_expiry.strftime('%B %d, %Y')}\n"
            f"Days Remaining: {days_remaining} days\n"
            f"Status: {'Active' if days_remaining > 0 else 'Expired'}\n\n"
            f"Max Users: {MAX_USERS}"
        )
        
        messagebox.showinfo("License Information", info_text)
    
    def get_logs_folder_path(self):
        """Get the path to the LOGS folder in the destination"""
        if self.dest_folder:
            return os.path.join(self.dest_folder, "LOGS")
        return None
    
    def ensure_logs_folder_exists(self):
        """Create LOGS folder if it doesn't exist"""
        logs_path = self.get_logs_folder_path()
        if logs_path:
            try:
                os.makedirs(logs_path, exist_ok=True)
                return logs_path
            except Exception as e:
                self.log_message(f"Failed to create LOGS folder: {e}", "WARNING")
        return None
    
    def save_sync_log_to_file(self):
        """Save the current sync session log to a file"""
        if not self.current_sync_log:
            return
        
        logs_folder = self.ensure_logs_folder_exists()
        if not logs_folder:
            self.log_message("Cannot save log file - LOGS folder unavailable", "WARNING")
            return
        
        try:
            filename = "sync_log.txt"
            log_path = os.path.join(logs_folder, filename)
            
            with open(log_path, 'w', encoding='utf-8') as f:
                f.write("="*60 + "\n")
                f.write("SyncFinity Sync Log\n")
                f.write(f"User: #{self.user_id}\n")
                f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("="*60 + "\n\n")
                
                for log_entry in self.current_sync_log:
                    f.write(log_entry + "\n")
                
                f.write("\n" + "="*60 + "\n")
                f.write("End of Sync Log\n")
                f.write("="*60 + "\n")
            
            self.log_message(f"Sync log saved: {filename}", "SUCCESS")
            return log_path
            
        except Exception as e:
            self.log_message(f"Failed to save sync log: {e}", "ERROR")
            return None
    
    def is_network_path(self, path):
        """Check if path is a network path (UNC path)"""
        return path.startswith('\\\\') or path.startswith('//')
    
    def wait_for_network_path(self, path, max_retries=5, retry_delay=2):
        """
        Wait for network path to become accessible with retry logic
        Returns: (is_accessible, error_message)
        """
        if not self.is_network_path(path):
            # Local path, check immediately
            return os.path.exists(path), None
        
        # Network path - retry with delays
        for attempt in range(max_retries):
            try:
                if os.path.exists(path):
                    if attempt > 0:
                        self.log_message(f"Network share accessible after {attempt + 1} attempt(s)", "SUCCESS")
                    return True, None
            except Exception as e:
                pass
            
            if attempt < max_retries - 1:
                self.log_message(f"Network share not ready, retrying in {retry_delay}s... (Attempt {attempt + 1}/{max_retries})", "WARNING")
                time.sleep(retry_delay)
        
        return False, f"Network share not accessible after {max_retries} attempts"
    
    def auto_start_sync(self):
        """Automatically start sync if configuration is complete"""
        if not self.auto_start_enabled:
            return
        
        # Check basic configuration
        if not self.source_folders:
            self.log_message("Auto-start skipped: No source folders configured", "WARNING")
            return
        
        if not self.dest_folder:
            self.log_message("Auto-start skipped: No destination folder configured", "WARNING")
            return
        
        # For network destinations, wait for accessibility
        if self.is_network_path(self.dest_folder):
            self.log_message(f"Checking network destination accessibility: {self.dest_folder}", "INFO")
            is_accessible, error_msg = self.wait_for_network_path(self.dest_folder)
            
            if not is_accessible:
                self.log_message(f"Auto-start delayed: {error_msg}", "WARNING")
                self.log_message("Destination will be checked again when sync starts", "INFO")
                # Still start sync monitor - it will check accessibility when syncing
        
        if not self.sync_running:
            self.log_message("Auto-starting sync monitor with saved configuration...", "INFO")
            self.start_sync()
    
    def log_message(self, message, level="INFO"):
        """Add message to log queue for thread-safe logging"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_queue.put((timestamp, message, level))
        
        if hasattr(self, 'current_sync_log'):
            log_entry = f"[{timestamp}] {message}"
            self.current_sync_log.append(log_entry)
    
    def start_new_sync_log(self):
        """Start a new sync log session"""
        self.current_sync_log = []
        self.log_message("="*50, "INFO")
        self.log_message("NEW SYNC SESSION STARTED", "SUCCESS")
    
    def process_log_queue(self):
        """Process messages from log queue and display them"""
        try:
            while True:
                timestamp, message, level = self.log_queue.get_nowait()
                
                color_map = {
                    "INFO": "black",
                    "SUCCESS": self.success_color,
                    "WARNING": self.warning_color,
                    "ERROR": self.error_color
                }
                
                color = color_map.get(level, "black")
                
                self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
                
                if color != "black":
                    start_line = self.log_text.index(tk.END + "-2l linestart")
                    end_line = self.log_text.index(tk.END + "-1l lineend")
                    self.log_text.tag_add(level, start_line, end_line)
                    self.log_text.tag_config(level, foreground=color)
                
                self.log_text.see(tk.END)
                
        except queue.Empty:
            pass
        
        self.root.after(100, self.process_log_queue)
    
    def add_source_folder(self):
        folder = filedialog.askdirectory(title="Select Source Folder")
        if folder and folder not in self.source_folders:
            self.source_folders.append(folder)
            self.update_source_display()
            self.save_config()
    
    def remove_source_folder(self):
        selected_indices = self.source_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("No Selection", "Please select folders to remove.")
            return
        
        for i in reversed(selected_indices):
            del self.source_folders[i]
        
        self.update_source_display()
        self.save_config()
    
    def auto_select_sources(self):
        user_home = os.path.expanduser("~")
        auto_folders = []
        
        documents = os.path.join(user_home, "Documents")
        desktop = os.path.join(user_home, "Desktop")
        
        if os.path.exists(documents):
            auto_folders.append(documents)
        if os.path.exists(desktop):
            auto_folders.append(desktop)
        
        if auto_folders:
            added = 0
            for folder in auto_folders:
                if folder not in self.source_folders:
                    self.source_folders.append(folder)
                    added += 1
            
            self.update_source_display()
            self.save_config()
            
            if added > 0:
                self.log_message(f"Auto-selected {added} folders: Documents and/or Desktop", "SUCCESS")
            else:
                self.log_message("Documents and Desktop folders already selected", "INFO")
        else:
            messagebox.showinfo("Auto-Select", "Documents and Desktop folders not found.")
    
    def update_source_display(self):
        self.source_listbox.delete(0, tk.END)
        for folder in self.source_folders:
            display_name = os.path.basename(folder) or folder
            self.source_listbox.insert(tk.END, f"{display_name} ({folder})")
    
    def on_server_change(self, event=None):
        self.dest_folder = ""
        self.dest_choice = None
        self.update_dest_display()
    
    def browse_destination(self):
        server_text = self.server_var.get()
        if not server_text:
            messagebox.showwarning("Server Required", "Please select a server first.")
            return
        
        if "Primary" in server_text:
            base_path = FILE_SERVER_BASE
            self.dest_choice = 1
        else:
            base_path = FILE_SERVER_BASE_2
            self.dest_choice = 2
        
        if not os.path.exists(base_path):
            result = messagebox.askyesno("Server Inaccessible", 
                                       f"Server path '{base_path}' is not accessible.\n"
                                       "Do you want to select a local folder instead?")
            if result:
                folder = filedialog.askdirectory(title="Select Destination Folder")
                if folder:
                    self.dest_folder = folder
                    self.update_dest_display()
                    self.save_config()
            return
        
        folder = filedialog.askdirectory(title="Select Destination Folder", 
                                       initialdir=base_path)
        if folder:
            self.dest_folder = folder
            self.update_dest_display()
            self.save_config()
    
    def update_dest_display(self):
        if self.dest_folder:
            self.dest_label.config(text=self.dest_folder, foreground="black")
        else:
            self.dest_label.config(text="No destination selected", foreground="gray")
    
    def start_sync(self):
        if not self.source_folders:
            messagebox.showerror("Error", "Please select at least one source folder.")
            return
        
        if not self.dest_folder:
            messagebox.showerror("Error", "Please select a destination folder.")
            return
        
        global SOURCE_FOLDERS, DEST_FOLDER
        SOURCE_FOLDERS = self.source_folders.copy()
        DEST_FOLDER = self.dest_folder
        
        self.sync_running = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        
        self.next_sync_time = datetime.now()
        
        server_name = "Primary Server" if self.dest_choice == 1 else "Backup Server" if self.dest_choice == 2 else "Custom"
        
        self.start_new_sync_log()
        self.log_message("SYNC MONITOR STARTED", "SUCCESS")
        self.log_message(f"Sources: {len(self.source_folders)} folders", "INFO")
        self.log_message(f"Destination: {server_name}", "INFO")
        self.log_message(f"Path: {self.dest_folder}", "INFO")
        self.log_message(f"Sync interval: {self.sync_interval_hours} hours", "INFO")
        self.log_message("First sync starting immediately...", "INFO")
        self.log_message("="*50, "INFO")
        
        self.status_var.set("Sync monitor running...")
        
        self.sync_thread = threading.Thread(target=self.sync_worker, daemon=True)
        self.sync_thread.start()
        
        self.update_status_timer()
    
    def stop_sync(self):
        self.sync_running = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        
        self.log_message("Sync monitor stopped by user", "WARNING")
        self.status_var.set("Ready")
        self.next_sync_time = None
    
    def manual_sync(self):
        if not self.source_folders:
            messagebox.showerror("Error", "Please select at least one source folder.")
            return
        
        if not self.dest_folder:
            messagebox.showerror("Error", "Please select a destination folder.")
            return
        
        global SOURCE_FOLDERS, DEST_FOLDER
        SOURCE_FOLDERS = self.source_folders.copy()
        DEST_FOLDER = self.dest_folder
        
        self.start_new_sync_log()
        self.log_message("Starting manual sync...", "INFO")
        self.status_var.set("Manual sync in progress...")
        
        self.manual_sync_btn.config(state=tk.DISABLED)
        
        threading.Thread(target=self.manual_sync_worker, daemon=True).start()
    
    def manual_sync_worker(self):
        success = self.sync_files()
        
        self.root.after(0, lambda: self.manual_sync_btn.config(state=tk.NORMAL))
        self.root.after(0, lambda: self.status_var.set("Ready"))
        
        if success:
            self.log_message("Manual sync completed", "SUCCESS")
            self.root.after(0, self.save_sync_log_to_file)
        else:
            self.log_message("Manual sync failed", "ERROR")
    
    def update_status_timer(self):
        """Update status bar with countdown timer"""
        if not self.sync_running or not self.next_sync_time:
            return
        
        now = datetime.now()
        if now >= self.next_sync_time:
            return
        
        time_remaining = self.next_sync_time - now
        hours = int(time_remaining.total_seconds() // 3600)
        minutes = int((time_remaining.total_seconds() % 3600) // 60)
        
        if hours > 0:
            status_text = f"Next sync in {hours}h {minutes}m"
        elif minutes > 0:
            status_text = f"Next sync in {minutes}m"
        else:
            status_text = "Next sync starting soon..."
        
        self.status_var.set(status_text)
        
        self.root.after(30000, self.update_status_timer)
    
    def sync_worker(self):
        """Worker thread for continuous sync monitoring"""
        while self.sync_running:
            try:
                now = datetime.now()
                
                if now >= self.next_sync_time:
                    self.log_message(f"Starting scheduled sync at {now.strftime('%H:%M:%S')}", "INFO")
                    success = self.sync_files()
                    
                    if success:
                        self.log_message("Sync cycle completed successfully", "SUCCESS")
                        self.root.after(0, self.save_sync_log_to_file)
                        
                        self.next_sync_time = now + timedelta(hours=self.sync_interval_hours)
                        next_sync_str = self.next_sync_time.strftime("%I:%M %p")
                        self.log_message(f"Next sync scheduled for {next_sync_str}", "INFO")
                        
                    else:
                        self.log_message("Sync cycle failed, retrying in 1 hour", "ERROR")
                        self.next_sync_time = now + timedelta(hours=1)
                        retry_time_str = self.next_sync_time.strftime("%I:%M %p")
                        self.log_message(f"Retry scheduled for {retry_time_str}", "WARNING")
                    
                    self.root.after(0, self.update_status_timer)
                
                time.sleep(60)
                
            except Exception as e:
                self.log_message(f"Sync worker error: {e}", "ERROR")
                time.sleep(60)
    
    def sync_files(self):
        """Sync files from source to destination"""
        if not SOURCE_FOLDERS or not DEST_FOLDER:
            return False
        
        # Check destination accessibility with retry for network paths
        if self.is_network_path(DEST_FOLDER):
            self.log_message(f"Verifying network destination: {DEST_FOLDER}", "INFO")
            is_accessible, error_msg = self.wait_for_network_path(DEST_FOLDER, max_retries=3, retry_delay=3)
            
            if not is_accessible:
                self.log_message(f"Destination not accessible: {error_msg}", "ERROR")
                return False
        elif not os.path.exists(DEST_FOLDER):
            self.log_message(f"Destination folder not accessible: {DEST_FOLDER}", "ERROR")
            return False
        
        self.log_message(f"Syncing {len(SOURCE_FOLDERS)} source folder(s)", "INFO")
        self.log_message(f"Allowed: {', '.join(ALLOWED_EXTENSIONS)} | Max: {MAX_FILE_SIZE_MB}MB", "INFO")
        
        total_processed = 0
        total_filtered = 0
        
        for source_path in SOURCE_FOLDERS:
            if not os.path.exists(source_path):
                self.log_message(f"Source not found, skipping: {source_path}", "WARNING")
                continue
            
            source_name = os.path.basename(source_path) or "Root"
            dest_subfolder = os.path.join(DEST_FOLDER, source_name)
            
            try:
                os.makedirs(dest_subfolder, exist_ok=True)
            except Exception as e:
                self.log_message(f"Failed to create destination for {source_name}: {e}", "ERROR")
                continue
            
            self.log_message(f"Processing: {source_name}", "INFO")
            processed, filtered = self.sync_folder_recursively(source_path, dest_subfolder, source_name)
            
            total_processed += processed
            total_filtered += filtered
            
            self.log_message(f"Completed {source_name}: {processed} files, {filtered} filtered", "SUCCESS")
        
        self.log_message(f"TOTAL: {total_processed} files processed, {total_filtered} filtered", "SUCCESS")
        return True
    
    def sync_folder_recursively(self, src_folder, dest_folder, source_name):
        """Recursively sync folder contents"""
        files_processed = 0
        files_filtered = 0
        
        try:
            for item_name in os.listdir(src_folder):
                src_path = os.path.join(src_folder, item_name)
                dest_path = os.path.join(dest_folder, item_name)
                
                if os.path.isfile(src_path):
                    is_allowed, reason = self.is_file_allowed(src_path)
                    if not is_allowed:
                        files_filtered += 1
                        rel_path = os.path.relpath(src_path, src_folder)
                        self.log_message(f"[{source_name}] Filtered: {rel_path} - {reason}", "WARNING")
                        continue
                    
                    files_processed += 1
                    try:
                        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                        file_exists = os.path.exists(dest_path)
                        
                        shutil.copy2(src_path, dest_path)
                        file_size_mb = os.path.getsize(src_path) / (1024 * 1024)
                        rel_path = os.path.relpath(src_path, src_folder)
                        
                        action = "Replaced" if file_exists else "Uploaded"
                        self.log_message(f"[{source_name}] {action}: {rel_path} ({file_size_mb:.1f}MB)", "SUCCESS")
                        
                    except Exception as e:
                        rel_path = os.path.relpath(src_path, src_folder)
                        self.log_message(f"[{source_name}] Failed: {rel_path} - {e}", "ERROR")
                
                elif os.path.isdir(src_path):
                    try:
                        os.makedirs(dest_path, exist_ok=True)
                        sub_processed, sub_filtered = self.sync_folder_recursively(src_path, dest_path, source_name)
                        files_processed += sub_processed
                        files_filtered += sub_filtered
                    except Exception as e:
                        rel_path = os.path.relpath(src_path, src_folder)
                        self.log_message(f"[{source_name}] Folder error: {rel_path} - {e}", "ERROR")
        
        except Exception as e:
            self.log_message(f"Error processing {src_folder}: {e}", "ERROR")
        
        return files_processed, files_filtered
    
    def is_file_allowed(self, file_path):
        """Check if file meets filtering criteria"""
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext not in ALLOWED_EXTENSIONS:
            return False, f"Extension '{file_ext}' not allowed"
        
        try:
            file_size = os.path.getsize(file_path)
            if file_size > MAX_FILE_SIZE_BYTES:
                size_mb = file_size / (1024 * 1024)
                return False, f"Too large ({size_mb:.1f}MB > {MAX_FILE_SIZE_MB}MB)"
        except OSError as e:
            return False, f"Cannot check size: {e}"
        
        return True, "Allowed"
    
    def open_settings(self):
        """Open settings dialog"""
        settings_window = tk.Toplevel(self.root)
        settings_window.title("Settings")
        settings_window.geometry("500x550")
        settings_window.transient(self.root)
        settings_window.grab_set()
        
        try:
            icon_path = resource_path("infinity.ico")
            if os.path.exists(icon_path):
                settings_window.iconbitmap(icon_path)
        except:
            pass
        
        # Ensure window is visible
        if not self.window_visible:
            self.root.deiconify()
            self.window_visible = True
        
        auto_frame = ttk.LabelFrame(settings_window, text="Auto-Start Options", padding="10")
        auto_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.auto_start_var = tk.BooleanVar(value=self.auto_start_enabled)
        auto_check = ttk.Checkbutton(auto_frame, 
                                   text="Auto-start sync monitor on launch (when configured)",
                                   variable=self.auto_start_var)
        auto_check.pack(anchor=tk.W)
        # Startup shortcut section
        startup_frame = ttk.LabelFrame(settings_window, text="Windows Startup", padding="10")
        startup_frame.pack(fill=tk.X, padx=10, pady=10)

        exists, shortcut_path = LicenseManager.check_startup_shortcut()
        status_text = "Enabled - SyncFinity will start with Windows" if exists else "Disabled"
        status_color = "green" if exists else "gray"

        status_label = ttk.Label(startup_frame, text=f"Status: {status_text}",
                                 foreground=status_color)
        status_label.pack(anchor=tk.W, pady=(0, 5))

        startup_btn_frame = ttk.Frame(startup_frame)
        startup_btn_frame.pack(fill=tk.X)

        def create_shortcut():
            success, result = LicenseManager.create_startup_shortcut()
            if success:
                messagebox.showinfo("Success", f"Startup shortcut created:\n{result}")
                status_label.config(text="Status: Enabled - SyncFinity will start with Windows",
                                    foreground="green")
                self.log_message("Startup shortcut created", "SUCCESS")
            else:
                messagebox.showerror("Error", f"Failed to create shortcut:\n{result}")

        def remove_shortcut():
            success, message = LicenseManager.remove_startup_shortcut()
            if success:
                messagebox.showinfo("Success", message)
                status_label.config(text="Status: Disabled", foreground="gray")
                self.log_message("Startup shortcut removed", "INFO")
            else:
                messagebox.showerror("Error", f"Failed to remove shortcut:\n{message}")

        ttk.Button(startup_btn_frame, text="Enable Startup",
                   command=create_shortcut).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(startup_btn_frame, text="Disable Startup",
                   command=remove_shortcut).pack(side=tk.LEFT)
        interval_frame = ttk.LabelFrame(settings_window, text="Sync Interval", padding="10")
        interval_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.interval_var = tk.StringVar(value=str(self.sync_interval_hours))
        ttk.Label(interval_frame, text="Sync every X hours:").pack(anchor=tk.W)
        interval_entry = ttk.Entry(interval_frame, textvariable=self.interval_var, width=10)
        interval_entry.pack(anchor=tk.W, pady=5)
        ttk.Label(interval_frame, text="(Recommended: 4 hours)", font=('Arial', 8)).pack(anchor=tk.W)
        
        ext_frame = ttk.LabelFrame(settings_window, text="Allowed File Extensions", padding="10")
        ext_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.ext_var = tk.StringVar(value=', '.join(ALLOWED_EXTENSIONS))
        ttk.Label(ext_frame, text="Extensions (comma-separated):").pack(anchor=tk.W)
        ext_entry = ttk.Entry(ext_frame, textvariable=self.ext_var, width=50)
        ext_entry.pack(fill=tk.X, pady=5)
        
        size_frame = ttk.LabelFrame(settings_window, text="Maximum File Size", padding="10")
        size_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.size_var = tk.StringVar(value=str(MAX_FILE_SIZE_MB))
        ttk.Label(size_frame, text="Maximum file size (MB):").pack(anchor=tk.W)
        size_entry = ttk.Entry(size_frame, textvariable=self.size_var, width=10)
        size_entry.pack(anchor=tk.W, pady=5)
        
        btn_frame = ttk.Frame(settings_window)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(btn_frame, text="Save", 
                  command=lambda: self.save_settings(settings_window)).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="Cancel", 
                  command=settings_window.destroy).pack(side=tk.RIGHT)
    
    def save_settings(self, window):
        """Save settings and close dialog"""
        
        global ALLOWED_EXTENSIONS, MAX_FILE_SIZE_MB, MAX_FILE_SIZE_BYTES
        global FILE_SERVER_BASE, FILE_SERVER_BASE_2
        
        try:
            self.auto_start_enabled = self.auto_start_var.get()
            
            interval_text = self.interval_var.get().strip()
            if interval_text:
                new_interval = int(interval_text)
                if new_interval > 0:
                    self.sync_interval_hours = new_interval
                    self.log_message(f"Sync interval updated to {new_interval} hours", "SUCCESS")
                else:
                    raise ValueError("Interval must be greater than 0")
            
            ext_text = self.ext_var.get().strip()
            if ext_text:
                new_extensions = set()
                for ext in ext_text.split(','):
                    ext = ext.strip()
                    if ext and not ext.startswith('.'):
                        ext = '.' + ext
                    if ext:
                        new_extensions.add(ext.lower())
                ALLOWED_EXTENSIONS = new_extensions
            
            size_text = self.size_var.get().strip()
            if size_text:
                MAX_FILE_SIZE_MB = int(size_text)
                MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024
            
            self.save_config()
            
            self.log_message("Settings updated successfully", "SUCCESS")
            window.destroy()
            
        except ValueError as e:
            messagebox.showerror("Settings Error", f"Invalid value: {e}")
        except Exception as e:
            messagebox.showerror("Settings Error", f"Error saving settings: {e}")
    
    def save_log(self):
        filename = filedialog.asksaveasfilename(
            title="Save Log File",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.get(1.0, tk.END))
                self.log_message(f"Log saved to: {filename}", "SUCCESS")
            except Exception as e:
                messagebox.showerror("Save Error", f"Failed to save log: {e}")
    
    def save_config(self):
        """Save current configuration"""
        config = {
            "source_folders": self.source_folders,
            "dest_folder": self.dest_folder,
            "dest_choice": self.dest_choice,
            "auto_start_enabled": self.auto_start_enabled,
            "sync_interval_hours": self.sync_interval_hours,
            "last_updated": time.strftime('%Y-%m-%d %H:%M:%S')
        }
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(config, f, indent=2)
        except Exception as e:
            self.log_message(f"Warning: Could not save configuration: {e}", "WARNING")
    
    def load_config(self):
        """Load saved configuration"""
        if not os.path.exists(CONFIG_FILE):
            return
        
        try:
            with open(CONFIG_FILE, 'r') as f:
                config = json.load(f)
            
            self.auto_start_enabled = config.get("auto_start_enabled", True)
            self.sync_interval_hours = config.get("sync_interval_hours", 4)
            
            source_folders = config.get("source_folders", [])
            valid_sources = [folder for folder in source_folders if os.path.exists(folder)]
            if valid_sources:
                self.source_folders = valid_sources
                self.update_source_display()
            
            dest_folder = config.get("dest_folder", "")
            dest_choice = config.get("dest_choice", None)
            
            # OPTION 1 FIX: Trust saved config without checking existence
            # Network shares may not be immediately available at startup
            if dest_folder and dest_choice is not None:
                self.dest_folder = dest_folder
                self.dest_choice = dest_choice
                
                if dest_choice == 1:
                    self.server_var.set("Primary Server (shares)")
                elif dest_choice == 2:
                    self.server_var.set("Backup Server (shares2)")
                
                self.update_dest_display()
                self.log_message(f"Destination loaded from config: {dest_folder}", "INFO")
            
            last_updated = config.get("last_updated", "Unknown")
            self.log_message(f"Configuration loaded (updated: {last_updated})", "SUCCESS")
            self.log_message(f"Sync interval: {self.sync_interval_hours} hours", "INFO")
            
        except Exception as e:
            self.log_message(f"Warning: Could not load configuration: {e}", "WARNING")

    def on_closing(self):
        """Handle application closing - minimize to tray instead"""
        if messagebox.askyesno("Exit", "Do you want to exit SyncFinity completely?\n\nClick 'No' to minimize to system tray."):
            self.quit_application()
        else:
            self.root.withdraw()
            self.window_visible = False


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath("C:\\Users\\Akonwele Jeffery\\Downloads\\iconTrial")
    return os.path.join(base_path, relative_path)


def main():
    """Main function to run the GUI application"""
    force_auto_start = "--auto-start" in sys.argv
    
    root = tk.Tk()
    
    # Try to set icon
    try:
        icon_path = resource_path("infinity.ico")
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
    except:
        pass
    
    app = FileSyncGUI(root)
    
    if force_auto_start:
        app.auto_start_enabled = True
        root.after(1000, app.auto_start_sync)
    
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()


if __name__ == "__main__":
    main()