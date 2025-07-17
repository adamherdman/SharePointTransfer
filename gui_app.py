import customtkinter as ctk
import threading
import queue
import os
import re
import json
import subprocess
import sys
from tkinter import messagebox, filedialog

from PIL import Image
from customtkinter import CTkImage

# Import all logic modules
from discovery_logic import discover_data_folders, discover_sub_folders
from download_logic import perform_download
from upload_logic import perform_upload

# --- ROBUST HELPER FUNCTIONS FOR PATHS ---
def get_base_path():
    """
    Get the base path for the application, which works for both a script
    and a PyInstaller-bundled executable.
    """
    if getattr(sys, 'frozen', False):
        # We are running in a bundle (e.g., a PyInstaller .exe)
        return os.path.dirname(sys.executable)
    else:
        # We are running in a normal Python environment (as a .py script)
        return os.path.dirname(os.path.abspath(__file__))

def get_asset_path(relative_path):
    """
    Get the absolute path to a resource (like an image), which works for
    both a script and a PyInstaller bundle.
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # Not running in a bundle, so use the script's directory
        base_path = get_base_path()
    return os.path.join(base_path, relative_path)

# --- DIALOG CLASSES (No changes needed in these) ---
# ... (IdentifierInputDialog, PassphraseDialog, SharePointFolderExplorerDialog, ManifestNameDialog classes remain unchanged from the previous version) ...
class IdentifierInputDialog(ctk.CTkToplevel):
    def __init__(self, parent, folder_name):
        super().__init__(parent)
        self.title("Enter Identifier")
        self.geometry("450x200")
        self.transient(parent)
        self.grab_set()
        self.identifier = None
        self.label_info = ctk.CTkLabel(self, text=f"The selected SharePoint folder is:\n'{folder_name}'", wraplength=400)
        self.label_info.pack(pady=(10,0), padx=20)
        self.label_prompt = ctk.CTkLabel(self, text="This name doesn't contain a 'Dxxxx' ID.\nPlease enter the local identifier (e.g., D12345):")
        self.label_prompt.pack(pady=(5,5), padx=20)
        self.entry = ctk.CTkEntry(self)
        self.entry.pack(pady=5, padx=20, fill="x")
        self.entry.focus()
        self.button_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.button_frame.pack(pady=(15,10), padx=20, fill="x")
        self.button_frame.grid_columnconfigure((0,1), weight=1)
        self.ok_button = ctk.CTkButton(self.button_frame, text="OK", command=self.on_ok)
        self.ok_button.grid(row=0, column=0, padx=(0,5), sticky="ew")
        self.cancel_button = ctk.CTkButton(self.button_frame, text="Cancel", command=self.on_cancel, fg_color="gray")
        self.cancel_button.grid(row=0, column=1, padx=(5,0), sticky="ew")
        self.entry.bind("<Return>", self.on_ok)
        self.bind("<Escape>", self.on_cancel)
    def on_ok(self, event=None):
        entered_text = self.entry.get().strip()
        if not re.fullmatch(r'D\d+', entered_text):
            messagebox.showerror("Invalid Input", "Identifier must be in the format 'D' followed by one or more numbers (e.g., D12345).", parent=self)
            self.entry.focus()
            return
        self.identifier = entered_text
        self.destroy()
    def on_cancel(self, event=None):
        self.identifier = None
        self.destroy()
    def get_identifier(self):
        self.master.wait_window(self)
        return self.identifier

class ConfigDialog(ctk.CTkToplevel):
    def __init__(self, parent, config_path):
        super().__init__(parent)
        self.title("Configuration")
        self.geometry("650x430")
        self.transient(parent)
        self.grab_set()
        self.config_path = config_path
        self.entries = {}
        # --- MODIFICATION: Default data path uses the robust helper ---
        self.default_data_path = os.path.join(get_base_path(), "Data")
        self.LABEL_MAP = {
            "DATA_FOLDER_PATH": "Local Data Folder Path",
            "APP_USERNAME": "Email for SP Download",
            "APP_PASSWORD": "Password for SP Download",
            "SFTP_HOSTNAME": "SFTP Hostname",
            "SFTP_PORT": "SFTP Port",
            "SFTP_USERNAME": "SFTP Username",
            "SFTP_PRIVATE_KEY_PATH": "SFTP Private Key Path"
        }
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)
        self.main_frame.grid_columnconfigure(1, weight=1)
        for i, (key, label_text) in enumerate(self.LABEL_MAP.items()):
            label = ctk.CTkLabel(self.main_frame, text=f"{label_text}:")
            label.grid(row=i, column=0, padx=10, pady=8, sticky="w")
            if key == "DATA_FOLDER_PATH":
                entry = ctk.CTkEntry(self.main_frame, placeholder_text=self.default_data_path)
            else:
                entry = ctk.CTkEntry(self.main_frame)
            entry.grid(row=i, column=1, padx=10, pady=8, sticky="ew")
            self.entries[key] = entry
            if key == "SFTP_PRIVATE_KEY_PATH":
                browse_button = ctk.CTkButton(self.main_frame, text="Browse...", width=80, command=self.browse_for_key_file)
                browse_button.grid(row=i, column=2, padx=(0, 10), pady=8)
            elif key == "DATA_FOLDER_PATH":
                browse_button = ctk.CTkButton(self.main_frame, text="Browse...", width=80, command=self.browse_for_data_folder)
                browse_button.grid(row=i, column=2, padx=(0, 10), pady=8)
        self.button_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.button_frame.pack(pady=10, padx=20, fill="x")
        self.button_frame.grid_columnconfigure((0, 1), weight=1)
        self.save_button = ctk.CTkButton(self.button_frame, text="Save", command=self.save_config)
        self.save_button.grid(row=0, column=0, padx=5, sticky="ew")
        self.cancel_button = ctk.CTkButton(self.button_frame, text="Cancel", command=self.destroy, fg_color="gray")
        self.cancel_button.grid(row=0, column=1, padx=5, sticky="ew")
        self.load_config()
    def browse_for_data_folder(self):
        dir_path = filedialog.askdirectory(title="Select Data Folder")
        if dir_path:
            entry = self.entries["DATA_FOLDER_PATH"]
            entry.delete(0, "end")
            entry.insert(0, dir_path)
    def browse_for_key_file(self):
        filepath = filedialog.askopenfilename(title="Select Private Key File",filetypes=(("All files", "*.*"), ("No extension (id_rsa)", "*"), ("PPK files", "*.ppk")))
        if filepath:
            entry = self.entries["SFTP_PRIVATE_KEY_PATH"]
            entry.delete(0, "end")
            entry.insert(0, filepath)
    def load_config(self):
        try:
            with open(self.config_path, 'r') as f:
                config_data = json.load(f)
            for key, entry in self.entries.items():
                entry.insert(0, config_data.get(key, ""))
        except (FileNotFoundError, json.JSONDecodeError): pass
    def save_config(self):
        new_config = {}
        for key, entry in self.entries.items():
            new_config[key] = entry.get()
        try:
            with open(self.config_path, 'w') as f:
                json.dump(new_config, f, indent=2)
            messagebox.showinfo("Success", "Configuration saved successfully.")
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save configuration:\n{e}")

class PassphraseDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Passphrase Required")
        self.geometry("350x150")
        self.transient(parent)
        self.grab_set()
        self.passphrase = None
        self.label = ctk.CTkLabel(self, text="Enter SFTP key passphrase (leave blank if none):")
        self.label.pack(pady=10, padx=20)
        self.entry = ctk.CTkEntry(self, show="*")
        self.entry.pack(pady=5, padx=20, fill="x")
        self.entry.focus()
        self.ok_button = ctk.CTkButton(self, text="OK", command=self.on_ok)
        self.ok_button.pack(pady=20)
        self.entry.bind("<Return>", self.on_ok)
    def on_ok(self, event=None):
        self.passphrase = self.entry.get()
        self.destroy()
    def get_passphrase(self):
        self.master.wait_window(self)
        return self.passphrase

class SharePointFolderExplorerDialog(ctk.CTkToplevel):
    def __init__(self, parent, sharepoint_url, top_level_folders):
        super().__init__(parent)
        self.parent_app = parent
        self.sharepoint_url = sharepoint_url
        self.selected_folder_path = None
        self.selected_folder_name = None
        self.path_stack = [("Shared Documents", "")]
        self.sub_folder_queue = queue.Queue()
        self.title("Select SharePoint Folder")
        self.geometry("500x550")
        self.transient(parent)
        self.grab_set()
        self.nav_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.nav_frame.pack(padx=20, pady=(10, 5), fill="x")
        self.nav_frame.grid_columnconfigure(1, weight=1)
        self.back_button = ctk.CTkButton(self.nav_frame, text="Up", command=self._navigate_back, width=60, state="disabled")
        self.back_button.grid(row=0, column=0, sticky="w")
        self.path_label = ctk.CTkLabel(self.nav_frame, text=self._get_current_path_display(), anchor="w", wraplength=380)
        self.path_label.grid(row=0, column=1, padx=10, sticky="ew")
        self.scroll_frame = ctk.CTkScrollableFrame(self)
        self.scroll_frame.pack(padx=20, pady=5, fill="both", expand=True)
        self.radio_var = ctk.StringVar()
        self.folder_widgets = []
        self.action_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.action_frame.pack(padx=20, pady=(5, 15), fill="x")
        self.action_frame.grid_columnconfigure((0,1), weight=1)
        self.download_button = ctk.CTkButton(self.action_frame, text="Select and Continue", command=self._on_select)
        self.download_button.grid(row=0, column=0, padx=(0,5), sticky="ew")
        self.cancel_button = ctk.CTkButton(self.action_frame, text="Cancel", command=self._on_cancel, fg_color="gray")
        self.cancel_button.grid(row=0, column=1, padx=(5,0), sticky="ew")
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self._populate_folder_list(top_level_folders)
        self._check_sub_folder_queue()
    def _get_current_path_display(self):
        return "/".join([item[0] for item in self.path_stack])
    def _get_current_path_relative_url(self):
        return "/".join([item[1] for item in self.path_stack if item[1]])
    def _clear_folder_list(self):
        for widget in self.folder_widgets:
            widget.destroy()
        self.folder_widgets = []
        self.radio_var.set(None)
    def _populate_folder_list(self, folders):
        self._clear_folder_list()
        if not folders:
            label = ctk.CTkLabel(self.scroll_frame, text="This folder is empty.")
            self.folder_widgets.append(label)
            label.pack(anchor="w", padx=10, pady=5)
            return
        for folder_name in sorted(folders, key=str.lower):
            frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
            frame.pack(fill="x", pady=2, padx=5)
            radio = ctk.CTkRadioButton(frame, text="", variable=self.radio_var, value=folder_name)
            radio.pack(side="left")
            nav_button = ctk.CTkButton(frame, text=folder_name, anchor="w", fg_color="transparent", text_color=ctk.ThemeManager.theme["CTkLabel"]["text_color"],
                                       command=lambda name=folder_name: self._navigate_to(name))
            nav_button.pack(side="left", fill="x", expand=True)
            self.folder_widgets.append(frame)
    def _navigate_to(self, folder_name):
        self.path_stack.append((folder_name, folder_name))
        self._update_view_for_navigation()
    def _navigate_back(self):
        if len(self.path_stack) > 1:
            self.path_stack.pop()
            self._update_view_for_navigation()
    def _update_view_for_navigation(self):
        self._clear_folder_list()
        loading_label = ctk.CTkLabel(self.scroll_frame, text="Loading...")
        loading_label.pack(pady=20)
        self.folder_widgets.append(loading_label)
        self.path_label.configure(text=self._get_current_path_display())
        self.back_button.configure(state="normal" if len(self.path_stack) > 1 else "disabled")
        site_relative_url = self.parent_app.web_properties['ServerRelativeUrl']
        current_relative_path = self._get_current_path_relative_url()
        parent_folder_url = f"{site_relative_url.rstrip('/')}/Shared Documents"
        if current_relative_path:
            parent_folder_url += f"/{current_relative_path}"
        
        config_path = self.parent_app.get_config_path()
        threading.Thread(target=discover_sub_folders, 
                         args=(self.sharepoint_url, parent_folder_url, self.sub_folder_queue, config_path), 
                         daemon=True).start()
    def _check_sub_folder_queue(self):
        try:
            msg_type, msg_data = self.sub_folder_queue.get_nowait()
            if msg_type == "sub_folders_found":
                self._populate_folder_list(msg_data)
            elif msg_type == "error":
                self._clear_folder_list()
                error_label = ctk.CTkLabel(self.scroll_frame, text=f"Error: {msg_data}", text_color="red", wraplength=400)
                error_label.pack(pady=20)
                self.folder_widgets.append(error_label)
        except queue.Empty:
            pass
        finally:
            self.after(100, self._check_sub_folder_queue)
    def _on_select(self):
        selected_name = self.radio_var.get()
        if not selected_name:
            messagebox.showwarning("No Selection", "Please select a folder to download.", parent=self)
            return
        current_relative_path = self._get_current_path_relative_url()
        self.selected_folder_path = f"{current_relative_path}/{selected_name}" if current_relative_path else selected_name
        self.selected_folder_name = selected_name
        self.destroy()
    def _on_cancel(self):
        self.selected_folder_path = None
        self.selected_folder_name = None
        self.destroy()
    def get_selection(self):
        self.master.wait_window(self)
        return self.selected_folder_path, self.selected_folder_name

class ManifestNameDialog(ctk.CTkToplevel):
    def __init__(self, parent, folder_path):
        super().__init__(parent)
        self.title("Enter Manifest Filename")
        self.geometry("450x200")
        self.transient(parent)
        self.grab_set()
        self.manifest_name = None
        self.label_info = ctk.CTkLabel(self, text=f"Selected folder:\n'{folder_path}'", wraplength=400)
        self.label_info.pack(pady=(10,0), padx=20)
        self.label_prompt = ctk.CTkLabel(self, text="Please enter the exact name of the .csv manifest file:")
        self.label_prompt.pack(pady=(5,5), padx=20)
        self.entry = ctk.CTkEntry(self, placeholder_text="e.g., manifest.csv or index.csv")
        self.entry.pack(pady=5, padx=20, fill="x")
        self.entry.focus()
        self.button_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.button_frame.pack(pady=(15,10), padx=20, fill="x")
        self.button_frame.grid_columnconfigure((0,1), weight=1)
        self.ok_button = ctk.CTkButton(self.button_frame, text="OK", command=self.on_ok)
        self.ok_button.grid(row=0, column=0, padx=(0,5), sticky="ew")
        self.cancel_button = ctk.CTkButton(self.button_frame, text="Cancel", command=self.on_cancel, fg_color="gray")
        self.cancel_button.grid(row=0, column=1, padx=(5,0), sticky="ew")
        self.entry.bind("<Return>", self.on_ok)
        self.bind("<Escape>", self.on_cancel)
    def on_ok(self, event=None):
        entered_text = self.entry.get().strip()
        if not entered_text.lower().endswith('.csv'):
            messagebox.showerror("Invalid Input", "Filename must end with .csv", parent=self)
            self.entry.focus()
            return
        if not entered_text:
            messagebox.showerror("Invalid Input", "Filename cannot be empty.", parent=self)
            self.entry.focus()
            return
        self.manifest_name = entered_text
        self.destroy()
    def on_cancel(self, event=None):
        self.manifest_name = None
        self.destroy()
    def get_manifest_name(self):
        self.master.wait_window(self)
        return self.manifest_name


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.withdraw()
        self.create_splash_screen()
        
        self.title("Data Transfer Hub")
        self.geometry("700x620")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        
        self.grid_columnconfigure(0, weight=1)
        
        try:
            icon_path = get_asset_path(os.path.join("Images", "di3.ico"))
            self.iconbitmap(icon_path)
        except Exception as e:
            print(f"Could not load icon: {e}")
            
        self.process_queue = queue.Queue()
        self.stop_event = threading.Event()
        self.download_folder_path = None
        self.web_properties = None

        self.url_label = ctk.CTkLabel(self, text="SharePoint Site URL:")
        self.url_label.grid(row=0, column=0, padx=20, pady=(20, 5), sticky="w")
        
        self.url_entry = ctk.CTkEntry(self, placeholder_text="https://your-tenant.sharepoint.com/sites/YourSite")
        self.url_entry.grid(row=1, column=0, padx=20, pady=5, sticky="ew")

        self.config_path_label = ctk.CTkLabel(self, text="Configuration File Path:")
        self.config_path_label.grid(row=2, column=0, padx=20, pady=(10, 5), sticky="w")
        
        self.config_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.config_frame.grid(row=3, column=0, padx=20, pady=(0,5), sticky="ew")
        self.config_frame.grid_columnconfigure(0, weight=1)
        
        self.config_path_entry = ctk.CTkEntry(self, placeholder_text="Default: config.json next to the app")
        self.config_path_entry.grid(row=0, column=0, in_=self.config_frame, sticky="ew")
        
        self.config_browse_button = ctk.CTkButton(self.config_frame, text="Browse...", width=100, command=self._browse_for_config)
        self.config_browse_button.grid(row=0, column=1, padx=(10, 0))
        
        self.top_button_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.top_button_frame.grid(row=4, column=0, padx=20, pady=(10, 5), sticky="ew")
        self.top_button_frame.grid_columnconfigure((0, 1), weight=1)
        
        self.config_button = ctk.CTkButton(self.top_button_frame, text="Edit Configuration", command=self.open_config_window)
        self.config_button.grid(row=0, column=0, padx=(0, 5), sticky="ew")
        
        self.open_folder_button = ctk.CTkButton(self.top_button_frame, text="Open Data Folder", command=self.open_download_folder, state="normal")
        self.open_folder_button.grid(row=0, column=1, padx=(5, 0), sticky="ew")
        
        self.action_button_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.action_button_frame.grid(row=5, column=0, padx=20, pady=5, sticky="ew")
        self.action_button_frame.grid_columnconfigure((0, 1), weight=1)
        
        self.download_button = ctk.CTkButton(self.action_button_frame, text="Start Download", command=self.start_discovery_process)
        self.download_button.grid(row=0, column=0, padx=(0, 5), sticky="ew")
        self.original_button_color = self.download_button.cget("fg_color")
        self.original_hover_color = self.download_button.cget("hover_color")
        
        self.upload_button = ctk.CTkButton(self.action_button_frame, text="Start Upload", command=self.start_upload_process)
        self.upload_button.grid(row=0, column=1, padx=(5, 0), sticky="ew")
        
        self.log_box = ctk.CTkTextbox(self, state="disabled", wrap="word")
        self.log_box.grid(row=6, column=0, padx=20, pady=5, sticky="nsew")
        self.grid_rowconfigure(6, weight=1)
        self.log("Welcome to the Data Transfer Hub!\n\nThis application transfers large datasets from SharePoint to an SFTP server.\n- Please specify your SharePoint URL and configuration file path.\n- The download process requires a manifest .csv file inside the selected SharePoint folder.\n- The upload process allows you to select a downloaded data folder to send to the SFTP server.\n--------------------------------------------------\nCreator: Adam Herdman\nTitle: Cloud Infrastructure Engineer, NWICB\nContact: adam.herdman@nhs.net\n--------------------------------------------------")
        
        self.progress_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.progress_frame.grid(row=7, column=0, padx=20, pady=(5, 20), sticky="ew")
        self.progress_frame.grid_columnconfigure(0, weight=1)
        
        self.status_label = ctk.CTkLabel(self.progress_frame, text="Status: Idle", anchor="w")
        self.status_label.grid(row=0, column=0, sticky="ew")
        
        self.filename_label = ctk.CTkLabel(self.progress_frame, text="", anchor="w", text_color="gray")
        self.filename_label.grid(row=1, column=0, sticky="ew")
        
        self.progress_bar = ctk.CTkProgressBar(self.progress_frame)
        self.progress_bar.set(0)
        self.progress_bar.grid(row=2, column=0, pady=(5,0), sticky="ew")
        
        self.check_queue()

    def _browse_for_config(self):
        filepath = filedialog.askopenfilename(
            title="Select Configuration File",
            filetypes=(("JSON files", "*.json"), ("All files", "*.*"))
        )
        if filepath:
            self.config_path_entry.delete(0, "end")
            self.config_path_entry.insert(0, filepath)
            self.log(f"Configuration file set to: {filepath}")

    def get_config_path(self):
        path = self.config_path_entry.get().strip()
        if path and os.path.isfile(path):
            return path
        return os.path.join(get_base_path(), "config.json")

    def create_splash_screen(self):
        splash_win = ctk.CTkToplevel(self)
        splash_win.overrideredirect(True)
        transparent_color = '#abcdef'
        splash_win.configure(fg_color=transparent_color)
        splash_win.wm_attributes("-transparentcolor", transparent_color)
        try:
            image_path = get_asset_path(os.path.join("Images", "di3.png"))
            pil_image = Image.open(image_path)
            width, height = pil_image.size
            splash_image = CTkImage(light_image=pil_image, size=(width, height))
            screen_width = splash_win.winfo_screenwidth()
            screen_height = splash_win.winfo_screenheight()
            x = (screen_width / 2) - (width / 2)
            y = (screen_height / 2) - (height / 2)
            splash_win.geometry(f'{width}x{height}+{int(x)}+{int(y)}')
            splash_label = ctk.CTkLabel(splash_win, image=splash_image, text="", fg_color=transparent_color)
            splash_label.pack()
            self.splash_win = splash_win
            splash_win.after(4000, self.show_main_window)
        except Exception as e:
            print(f"Splash screen image not found: {e}. Skipping.")
            if hasattr(self, 'splash_win'): splash_win.destroy()
            self.show_main_window()
            
    def show_main_window(self):
        if hasattr(self, 'splash_win') and self.splash_win.winfo_exists():
            self.splash_win.destroy()
        self.deiconify()
        
    def open_config_window(self):
        config_path = self.get_config_path()
        if not os.path.exists(config_path):
             self.log(f"Info: Config file not found at '{config_path}'. A new one will be created upon saving.")
        ConfigDialog(self, config_path=config_path)
        
    def get_data_folder_path(self):
        """
        Gets the data folder path. Prefers the path from config.json,
        but falls back to a 'Data' folder next to the application.
        """
        try:
            config_path = self.get_config_path()
            with open(config_path, 'r') as f:
                config = json.load(f)
            
            path = config.get("DATA_FOLDER_PATH", "").strip()
            
            if path and os.path.isdir(path):
                return path
        except (FileNotFoundError, json.JSONDecodeError):
            pass 
        
        # Fallback to a 'Data' folder next to the application's executable.
        return os.path.join(get_base_path(), "Data")
        
    def set_ui_for_processing(self, is_uploading=False, is_discovery=False):
        self.stop_event.clear()
        self.progress_bar.set(0)
        if not is_discovery:
            self.log_box.configure(state="normal")
            self.log_box.delete("1.0", "end")
            self.log_box.configure(state="disabled")
        self.download_button.configure(state="disabled")
        self.upload_button.configure(state="disabled")
        self.config_button.configure(state="disabled")
        self.open_folder_button.configure(state="disabled")
        if is_discovery:
            self.download_button.configure(text="Discovering...")
            return
        if is_uploading:
            self.upload_button.configure(text="Stop", command=self.stop_process, state="normal", fg_color="#D32F2F", hover_color="#B71C1C")
        else:
            self.download_button.configure(text="Stop", command=self.stop_process, state="normal", fg_color="#D32F2F", hover_color="#B71C1C")
            
    def reset_ui_from_processing(self):
        self.download_button.configure(text="Start Download", command=self.start_discovery_process, state="normal", fg_color=self.original_button_color, hover_color=self.original_hover_color)
        self.upload_button.configure(text="Start Upload", command=self.start_upload_process, state="normal", fg_color=self.original_button_color, hover_color=self.original_hover_color)
        self.config_button.configure(state="normal")
        self.open_folder_button.configure(state="normal")
        
    def stop_process(self):
        self.log("Sending stop signal...")
        self.status_label.configure(text="Status: Stopping...")
        self.download_button.configure(state="disabled")
        self.upload_button.configure(state="disabled")
        self.stop_event.set()
        
    def start_discovery_process(self):
        url = self.url_entry.get().strip()
        if not url:
            self.log("Error: SharePoint URL cannot be empty.")
            messagebox.showerror("Input Error", "Please enter a SharePoint Site URL.", parent=self)
            return
            
        config_path = self.get_config_path()
        if not os.path.exists(config_path):
            messagebox.showerror("Configuration Error", f"Configuration file not found at:\n{config_path}\n\nPlease create one using the 'Edit Configuration' button or specify a valid path.", parent=self)
            return

        self.set_ui_for_processing(is_discovery=True)
        self.log("Starting discovery process...")
        threading.Thread(target=discover_data_folders, args=(url, self.process_queue, config_path), daemon=True).start()

    def show_folder_explorer_dialog(self, top_level_folders):
        if not self.web_properties:
            self.log("❌ CRITICAL ERROR: Could not get SharePoint site properties before showing folder explorer.")
            messagebox.showerror("Error", "Could not get SharePoint site properties. Cannot proceed.", parent=self)
            self.reset_ui_from_processing()
            return

        dialog = SharePointFolderExplorerDialog(self, self.url_entry.get(), top_level_folders)
        selected_folder_path, selected_folder_name = dialog.get_selection()

        if selected_folder_path and selected_folder_name:
            self.log(f"User selected SharePoint folder: '{selected_folder_path}'.")
            
            manifest_dialog = ManifestNameDialog(self, selected_folder_path)
            manifest_name = manifest_dialog.get_manifest_name()

            if manifest_name:
                self.log(f"User provided manifest filename: '{manifest_name}'.")
                self._get_identifier_and_start_download(selected_folder_path, selected_folder_name, manifest_name)
            else:
                self.log("Manifest name input cancelled. Download aborted.")
                self.reset_ui_from_processing()
        else:
            self.log("Folder selection cancelled. Download aborted.")
            self.reset_ui_from_processing()

    def _get_identifier_and_start_download(self, folder_path, folder_name, manifest_filename):
        local_folder_id = None
        match = re.search(r'(D\d+)', folder_name)
        if match:
            local_folder_id = match.group(1)
            self.log(f"Using identifier '{local_folder_id}' found in folder name '{folder_name}'.")
            self.start_download_thread(folder_path, manifest_filename, local_folder_id)
        else:
            self.log(f"SharePoint folder name '{folder_name}' does not contain a 'Dxxxx' identifier.")
            self.log("Prompting user for local identifier...")
            id_dialog = IdentifierInputDialog(self, folder_name)
            manual_id = id_dialog.get_identifier()
            if manual_id:
                local_folder_id = manual_id
                self.log(f"User entered local identifier: '{local_folder_id}'.")
                self.start_download_thread(folder_path, manifest_filename, local_folder_id)
            else:
                self.log("Identifier input cancelled. Download aborted.")
                self.reset_ui_from_processing()

    def start_download_thread(self, sharepoint_folder_relative_path, manifest_filename, local_folder_id):
        url = self.url_entry.get().strip()
        data_folder_path = self.get_data_folder_path()
        config_path = self.get_config_path()
        output_dir = get_base_path()
        
        self.set_ui_for_processing(is_uploading=False)
        self.log(f"Starting download for SharePoint folder '{sharepoint_folder_relative_path}'.")
        self.log(f"Using manifest file: '{manifest_filename}'.")
        self.log(f"Local data will be saved to a subfolder named '{local_folder_id}' within: {data_folder_path}")
        threading.Thread(target=perform_download,
                         args=(url, sharepoint_folder_relative_path, manifest_filename, local_folder_id, data_folder_path, self.process_queue, self.stop_event, config_path, output_dir),
                         daemon=True).start()
    
    def start_upload_process(self):
        config_path = self.get_config_path()
        if not os.path.exists(config_path):
            messagebox.showerror("Configuration Error", f"Configuration file not found at:\n{config_path}\n\nPlease specify a valid configuration before uploading.", parent=self)
            return

        self.log("Scanning for local data folders...")
        data_dir = self.get_data_folder_path()
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)
            self.log(f"Created data directory: {data_dir}")
        
        d_folders = [d for d in os.listdir(data_dir) if os.path.isdir(os.path.join(data_dir, d))]
        if not d_folders:
            messagebox.showinfo("No Data Found", f"No data folders found in '{data_dir}' to upload.")
            return
        self.show_directory_selection_dialog(d_folders)

    def show_directory_selection_dialog(self, folders):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Select Directory to Upload")
        dialog.geometry("400x400")
        dialog.transient(self)
        dialog.grab_set()
        label = ctk.CTkLabel(dialog, text="Choose a directory to upload:")
        label.pack(pady=10)
        radio_var = ctk.StringVar(value=folders[0])
        scroll_frame = ctk.CTkScrollableFrame(dialog)
        scroll_frame.pack(pady=5, padx=20, fill="both", expand=True)
        for folder in folders:
            ctk.CTkRadioButton(scroll_frame, text=folder, variable=radio_var, value=folder).pack(anchor="w", padx=10, pady=5)
        def on_select():
            selected_folder = radio_var.get()
            dialog.destroy()
            self.get_passphrase_and_run_upload(selected_folder)
        def on_cancel():
            self.log("Selection cancelled by user.")
            self.reset_ui_from_processing()
            dialog.destroy()
        button = ctk.CTkButton(dialog, text="Select and Continue", command=on_select)
        button.pack(pady=20)
        dialog.protocol("WM_DELETE_WINDOW", on_cancel)
        
    def get_passphrase_and_run_upload(self, folder_name):
        dialog = PassphraseDialog(self)
        passphrase = dialog.get_passphrase()
        if passphrase is not None:
            data_folder_path = self.get_data_folder_path()
            local_path = os.path.join(data_folder_path, folder_name)
            config_path = self.get_config_path()
            output_dir = get_base_path()
            
            self.set_ui_for_processing(is_uploading=True)
            self.log(f"Starting upload for '{folder_name}'...")
            threading.Thread(target=perform_upload, args=(local_path, self.process_queue, self.stop_event, passphrase, config_path, output_dir), daemon=True).start()
            
    def log(self, message):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.configure(state="disabled")
        self.log_box.see("end")
        
    def show_completion_popup(self, title, error_count):
        if error_count == 0:
            message = f"The {title.lower()} process completed successfully with no unresolved errors."
        else:
            plural = "s" if error_count > 1 else ""
            message = f"The {title.lower()} process finished with {error_count} unresolved error{plural}.\n\nPlease check '{title.lower()}_errors.txt' for details."
        messagebox.showinfo(f"{title} Finished", message)

    def check_queue(self):
        try:
            while True:
                msg_type, msg_data = self.process_queue.get_nowait()
                if msg_type == "web_props":
                    self.web_properties = msg_data
                elif msg_type == "folders_found":
                    self.log("✅ Discovery complete. Please select a folder.")
                    self.status_label.configure(text="Status: Awaiting selection...")
                    self.reset_ui_from_processing()
                    self.show_folder_explorer_dialog(msg_data)
                elif msg_type == "status":
                    self.status_label.configure(text=f"Status: {msg_data}")
                    self.log(f"Status: {msg_data}")
                elif msg_type == "filename":
                    self.filename_label.configure(text=msg_data)
                elif msg_type == "progress":
                    current, total = msg_data
                    self.progress_bar.set(current / total if total > 0 else 0)
                    self.status_label.configure(text=f"Status: Processing... ({current}/{total})")
                elif msg_type == "done" or msg_type == "stopped":
                    is_upload = "upload" in self.status_label.cget("text").lower() or (self.filename_label.cget("text") and "upload" in self.filename_label.cget("text").lower())
                    title = "Upload" if is_upload else "Download"
                    self.download_folder_path, error_count = msg_data
                    self.reset_ui_from_processing()
                    self.status_label.configure(text=f"Status: {title} Complete!")
                    self.show_completion_popup(title, error_count)
                elif msg_type == "file_info":
                    self.log(f"ℹ️ {msg_data}")
                elif msg_type == "file_error":
                    self.log(f"⚠️ {msg_data}")
                elif msg_type == "error":
                    self.log(f"❌ CRITICAL ERROR: {msg_data}")
                    self.status_label.configure(text="Status: Critical Error!")
                    self.reset_ui_from_processing()
        except queue.Empty:
            pass
        finally:
            self.after(100, self.check_queue)
            
    def open_download_folder(self):
        path_to_open = self.get_data_folder_path()
        if not os.path.exists(path_to_open):
            try:
                os.makedirs(path_to_open)
                self.log(f"Created missing data directory: {path_to_open}")
            except Exception as e:
                self.log(f"Error: Could not create data directory at {path_to_open}: {e}")
                messagebox.showerror("Error", f"The data directory does not exist and could not be created:\n{path_to_open}")
                return

        self.log(f"Opening folder: {path_to_open}")
        try:
            if sys.platform == "win32":
                os.startfile(os.path.realpath(path_to_open))
            elif sys.platform == "darwin":
                subprocess.run(["open", path_to_open])
            else:
                subprocess.run(["xdg-open", path_to_open])
        except Exception as e:
            self.log(f"Error opening folder: {e}")
            messagebox.showerror("Error", f"Could not open the folder:\n{e}")


if __name__ == "__main__":
    app = App()
    app.mainloop()