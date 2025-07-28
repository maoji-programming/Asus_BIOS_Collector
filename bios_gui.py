from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import zipfile
import os
import logging
import json
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, scrolledtext, ttk
import re
import queue
from datetime import datetime

# Load configuration
CONFIG_JSON_PATH = 'config.json'
if os.path.exists(CONFIG_JSON_PATH):
    with open(CONFIG_JSON_PATH, 'r') as f:
        config = json.load(f)
else:
    config = {}

CHROMEDRIVER_PATH = config.get('chromedriver', ".\\chromedriver\\chromedriver.exe")
MODEL_LIST_PATH = config.get('model_list', ".\\model_list.txt")
LOGS_PATH = config.get('logs', ".\\logs")
DOWNLOAD_PATH = config.get('download_path', "E:\\BIOS")

# Configure logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(os.path.join(LOGS_PATH, "bios.log")),
        logging.StreamHandler()
    ]
)


def unzip_file(zip_file_path, extract_to_path):
    os.makedirs(extract_to_path, exist_ok=True)
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to_path)
        logging.info(f"Extracted all files to: {extract_to_path}")

def retrieve_model_list(file_path: str):
    if not os.path.exists(file_path):
        logging.error(f"File not found: {file_path}")
        return []
    models = set()
    with open(file_path, 'r') as file:
        for line in file:
            parts = line.strip().split()
            if len(parts) > 1:
                models.add(parts[1])
    unique_models = list(models)
    logging.info(f"Retrieved {len(unique_models)} unique models from the file.")
    return unique_models

# Global variable for success/fail count (for debug/logging)
GLOBAL_SUCCESS_COUNT = 0
GLOBAL_FAIL_COUNT = 0

def download_asus_bios(model: str, bios_download_path: str):
    global GLOBAL_SUCCESS_COUNT, GLOBAL_FAIL_COUNT
    os.makedirs(bios_download_path, exist_ok=True)
    current_version = get_bios_version_for_model(model, bios_download_path)
    if current_version is None:
        current_version = 0

    options = webdriver.ChromeOptions()
    chrome_driver_path = CHROMEDRIVER_PATH
    prefs = {"download.default_directory": bios_download_path}
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(executable_path=chrome_driver_path, options=options)

    website_url = f"https://www.asus.com/supportonly/{model.lower()}/helpdesk_bios/"
    logging.info(f"Navigating to: {website_url}")
    driver.get(website_url)
    driver.maximize_window()

    try:
        div_elements = driver.find_elements(By.CSS_SELECTOR, "div[class*='ProductSupportDriverBIOS__contentLeft'] div")
        print( d.text for d in div_elements) #test output
        for div in div_elements:
            if "BIOS for ASUS EZ Flash Utility" in div.text:
                logging.info("Locating the wording")
                next_div = div.find_elements(By.XPATH, "following-sibling::div/div")
                for div2_ in next_div:
                    if "Version" in div2_.text:
                        version = div2_.text.replace("Version", "").strip()
                        if int(version) > current_version:
                            logging.info(f"New BIOS version available: {version}")
                            try:
                                download_link = div.find_element(
                                    By.XPATH,
                                    "../../div[contains(@class,'ProductSupportDriverBIOS__contentRight')]/div[contains(@class,'ProductSupportDriverBIOS__downloadBtn__')]"
                                )
                                logging.info(f"Downloading BIOS of Model: {model} with version {version}...")
                                driver.execute_script("arguments[0].scrollIntoView();", download_link)
                                download_link.click()

                                # Wait for the zip file to appear (timeout: 60s)
                                downloaded_zip = os.path.join(bios_download_path, f"{model.upper()}AS{version}.zip")
                                max_wait = 60
                                waited = 0
                                while not os.path.exists(downloaded_zip) and waited < max_wait:
                                    time.sleep(1)
                                    waited += 1
                                if not os.path.exists(downloaded_zip):
                                    raise FileNotFoundError(f"Downloaded zip not found: {downloaded_zip}")

                                logging.info(f"unzip file: {downloaded_zip}")
                                unzip_file(downloaded_zip, bios_download_path)
                                logging.info("BIOS downloaded and unzipped successfully.")
                                if os.path.exists(downloaded_zip):
                                    os.remove(downloaded_zip)
                                    logging.info(f"Removed original zip file: {downloaded_zip}")
                                GLOBAL_SUCCESS_COUNT += 1
                            except Exception as e:
                                logging.error(f"Failed to download or unzip BIOS for {model}: {e}")
                                GLOBAL_FAIL_COUNT += 1
                        else:
                            logging.info(f"No new BIOS version available. Current version: {current_version}, Latest version: {version}")
                            GLOBAL_FAIL_COUNT += 1
    finally:
        driver.quit()
        logging.info("Browser closed.")

def extract_bios_version_from_filename(filename: str):
    """
    Check if the BIOS filename is in the format XXXXXX.NNN,
    where NNN is the version number, and return the version as int.
    Returns None if not matched.
    """
    match = re.match(r"^[A-Za-z0-9]+\.([0-9]{3})$", filename)
    if match:
        return int(match.group(1))
    return None

def get_bios_version_for_model(model: str, bios_folder: str):
    """
    Given a model and a folder, find the BIOS filename in BIOS format (XXXXXX.NNN),
    and return the version number (NNN as int). Returns None if not found.
    """
    # Only use bios_folder directly, do not join with "BIOS"
    if not os.path.isdir(bios_folder):
        logging.warning(f"BIOS folder not found: {bios_folder}")
        return None
    for filename in os.listdir(bios_folder):
        if filename.lower().startswith(model.lower()):
            version = extract_bios_version_from_filename(filename)
            if version is not None:
                return version
    return None

def execute(download_path=None, log_callback=None):
    if download_path is None:
        download_path = DOWNLOAD_PATH
    models = retrieve_model_list(MODEL_LIST_PATH)
    logging.info(f"Models to process: {', '.join(models)}")
    success_count = 0
    fail_count = 0

    for m in models:
        logging.info(f"Processing model: {m}")
        try:
            before_version = get_bios_version_for_model(m, download_path)
            download_asus_bios(model=m, bios_download_path=download_path)
            after_version = get_bios_version_for_model(m, download_path)
            if after_version is not None and (before_version is None or after_version > before_version):
                success_count += 1
                if log_callback:
                    log_callback(f"SUCCESS: {m}")
            else:
                fail_count += 1
                if log_callback:
                    log_callback(f"FAILED: {m} (No new BIOS downloaded)")
        except Exception as e:
            logging.error(f"Failed to process {m}: {e}")
            fail_count += 1
            if log_callback:
                log_callback(f"FAILED: {m} ({e})")

    logging.info(f"Completed. Success: {success_count}, Failed: {fail_count}")
    return success_count, fail_count

class TkinterLogHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text_widget.configure(state='normal')
            self.text_widget.insert('end', msg + '\n')
            self.text_widget.see('end')
            self.text_widget.configure(state='disabled')
        self.text_widget.after(0, append)

class BIOSDownloaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("ASUS BIOS Downloader")
        self.root.geometry("700x700")
        self.root.minsize(700, 700)

        # Variables
        self.config = config
        self.model_list = []
        self.is_running = False
        self.progress_queue = queue.Queue()
        self.stats = {
            'processed': 0,
            'success': 0,
            'failed': 0,
        }

        self.setup_styles()
        self.create_widgets()
        self.load_initial_config()
        self.monitor_progress()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Title.TLabel', font=('Arial', 16, 'bold'))
        style.configure('Success.TLabel', foreground='green', font=('Arial', 9, 'bold'))
        style.configure('Error.TLabel', foreground='red', font=('Arial', 9, 'bold'))
        style.configure('Warning.TLabel', foreground='orange', font=('Arial', 9, 'bold'))
        style.configure('Info.TLabel', foreground='blue', font=('Arial', 9, 'bold'))

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill="both", expand=True)

        # Title
        ttk.Label(main_frame, text="ASUS BIOS Downloader", style='Title.TLabel').pack(pady=(0, 10))

        # Config section
        self.create_config_section(main_frame)

        # Model list section
        self.create_model_list_section(main_frame)

        # Control section
        self.create_control_section(main_frame)

        # Progress section
        self.create_progress_section(main_frame)

        # Log section
        self.create_log_section(main_frame)

        # Status bar
        self.create_status_bar()

    def create_config_section(self, parent):
        config_frame = ttk.LabelFrame(parent, text="Configuration", padding="10")
        config_frame.pack(fill="x", pady=(0, 10))

        config_frame.columnconfigure(1, weight=1)
        config_frame.columnconfigure(3, weight=1)

        ttk.Label(config_frame, text="ChromeDriver Path:", font=('Arial', 9, 'bold')).grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
        self.driver_path_var = tk.StringVar(value=CHROMEDRIVER_PATH)
        ttk.Entry(config_frame, textvariable=self.driver_path_var).grid(row=0, column=1, columnspan=2, sticky="ew", padx=(0, 10), pady=5)
        ttk.Button(config_frame, text="Browse", command=self.browse_driver_path).grid(row=0, column=3, sticky="w", pady=5)

        ttk.Label(config_frame, text="Model List File:", font=('Arial', 9, 'bold')).grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
        self.model_file_var = tk.StringVar(value=MODEL_LIST_PATH)
        ttk.Entry(config_frame, textvariable=self.model_file_var).grid(row=1, column=1, columnspan=2, sticky="ew", padx=(0, 10), pady=5)
        ttk.Button(config_frame, text="Browse", command=self.browse_model_file).grid(row=1, column=3, sticky="w", pady=5)

        ttk.Label(config_frame, text="Download Path:", font=('Arial', 9, 'bold')).grid(row=2, column=0, sticky="w", padx=(0, 10), pady=5)
        self.download_path_var = tk.StringVar(value=DOWNLOAD_PATH)
        ttk.Entry(config_frame, textvariable=self.download_path_var).grid(row=2, column=1, columnspan=2, sticky="ew", padx=(0, 10), pady=5)
        ttk.Button(config_frame, text="Browse", command=self.browse_download_path).grid(row=2, column=3, sticky="w", pady=5)

    def create_model_list_section(self, parent):
        model_frame = ttk.LabelFrame(parent, text="Model List", padding="10")
        model_frame.pack(fill="both", expand=True, pady=(0, 10))

        header_frame = ttk.Frame(model_frame)
        header_frame.pack(fill="x", pady=(0, 10))

        ttk.Button(header_frame, text="Load Model List", command=self.load_model_file).pack(side="left", padx=(0, 10))
        ttk.Button(header_frame, text="Clear All", command=self.clear_model_list).pack(side="left")

        self.model_count_var = tk.StringVar(value="Models: 0")
        ttk.Label(header_frame, textvariable=self.model_count_var, style='Info.TLabel').pack(side="right")

        list_frame = ttk.Frame(model_frame)
        list_frame.pack(fill="both", expand=True)

        self.model_listbox = tk.Listbox(list_frame, height=6, selectmode="extended", font=('Consolas', 9))
        model_scrollbar_y = ttk.Scrollbar(list_frame, orient="vertical", command=self.model_listbox.yview)
        model_scrollbar_x = ttk.Scrollbar(list_frame, orient="horizontal", command=self.model_listbox.xview)
        self.model_listbox.configure(yscrollcommand=model_scrollbar_y.set, xscrollcommand=model_scrollbar_x.set)
        self.model_listbox.grid(row=0, column=0, sticky="nsew")
        model_scrollbar_y.grid(row=0, column=1, sticky="ns")
        model_scrollbar_x.grid(row=1, column=0, sticky="ew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

    def create_control_section(self, parent):
        control_frame = ttk.Frame(parent)
        control_frame.pack(pady=(0, 10))

        self.start_button = ttk.Button(control_frame, text="Start Download", command=self.start_download)
        self.start_button.pack(side="left", padx=(0, 15))

        self.stop_button = ttk.Button(control_frame, text="Stop", command=self.stop_download, state="disabled")
        self.stop_button.pack(side="left", padx=(0, 15))

        ttk.Button(control_frame, text="Open Download Folder", command=self.open_download_folder).pack(side="left", padx=(0, 15))
        ttk.Button(control_frame, text="Clear Logs", command=self.clear_logs).pack(side="left")

    def create_progress_section(self, parent):
        progress_frame = ttk.LabelFrame(parent, text="Progress", padding="10")
        progress_frame.pack(fill="x", pady=(0, 10))

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100, length=400)
        self.progress_bar.pack(fill="x", pady=(0, 10))

        self.progress_label_var = tk.StringVar(value="Ready to start")
        ttk.Label(progress_frame, textvariable=self.progress_label_var, font=('Arial', 10, 'bold')).pack(pady=(0, 10))

        stats_frame = ttk.Frame(progress_frame)
        stats_frame.pack()

        self.processed_var = tk.StringVar(value="Processed: 0")
        self.success_var = tk.StringVar(value="Success: 0")
        self.failed_var = tk.StringVar(value="Failed: 0")

        ttk.Label(stats_frame, textvariable=self.processed_var, style='Info.TLabel').grid(row=0, column=0, padx=20, pady=5)
        ttk.Label(stats_frame, textvariable=self.success_var, style='Success.TLabel').grid(row=0, column=1, padx=20, pady=5)
        ttk.Label(stats_frame, textvariable=self.failed_var, style='Error.TLabel').grid(row=0, column=2, padx=20, pady=5)

    def create_log_section(self, parent):
        log_frame = ttk.LabelFrame(parent, text="Activity Logs", padding="10")
        log_frame.pack(fill="both", expand=True, pady=(0, 10))

        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, wrap="word", font=('Consolas', 9))
        self.log_text.pack(fill="both", expand=True)
        self.log_text.tag_configure("success", foreground="green")
        self.log_text.tag_configure("error", foreground="red")
        self.log_text.tag_configure("warning", foreground="orange")
        self.log_text.tag_configure("info", foreground="blue")

    def create_status_bar(self):
        status_frame = ttk.Frame(self.root)
        status_frame.pack(side="bottom", fill="x")
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(status_frame, textvariable=self.status_var, relief="sunken", anchor="w", padding="5")
        status_bar.pack(side="left", fill="x", expand=True)
        version_label = ttk.Label(status_frame, text="v1.0", relief="sunken", padding="5")
        version_label.pack(side="right")

    def browse_driver_path(self):
        filename = filedialog.askopenfilename(
            title="Select ChromeDriver Executable",
            filetypes=[("Executable files", "*.exe"), ("All files", "*.*")]
        )
        if filename:
            self.driver_path_var.set(filename)

    def browse_model_file(self):
        filename = filedialog.askopenfilename(
            title="Select Model List File",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        if filename:
            self.model_file_var.set(filename)

    def browse_download_path(self):
        folder = filedialog.askdirectory(title="Select Download Folder")
        if folder:
            self.download_path_var.set(folder)

    def load_initial_config(self):
        self.driver_path_var.set(self.config.get('chromedriver', CHROMEDRIVER_PATH))
        self.model_file_var.set(self.config.get('model_list', MODEL_LIST_PATH))
        self.download_path_var.set(self.config.get('download_path', DOWNLOAD_PATH))
        self.log_message("Configuration loaded", "success")

    def load_model_file(self):
        file_path = self.model_file_var.get()
        self.model_list = retrieve_model_list(file_path)
        self.update_model_listbox()
        self.log_message(f"Loaded {len(self.model_list)} models from {os.path.basename(file_path)}", "success")

    def clear_model_list(self):
        self.model_list = []
        self.update_model_listbox()
        self.log_message("Model list cleared", "info")

    def update_model_listbox(self):
        self.model_listbox.delete(0, "end")
        for i, model in enumerate(self.model_list, 1):
            self.model_listbox.insert("end", f"{i:3d}. {model}")
        self.model_count_var.set(f"Models: {len(self.model_list)}")

    def open_download_folder(self):
        folder = self.download_path_var.get()
        if os.path.exists(folder):
            os.startfile(folder)
            self.log_message("Opened download folder", "info")
        else:
            messagebox.showinfo("Info", "Download folder does not exist yet.")

    def clear_logs(self):
        self.log_text.delete(1.0, "end")
        self.log_message("Logs cleared", "info")

    def log_message(self, message, tag="normal"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        self.log_text.insert("end", formatted_message, tag)
        self.log_text.see("end")
        self.status_var.set(message)
        try:
            with open("bios_download_log.txt", "a", encoding="utf-8") as f:
                f.write(formatted_message)
        except:
            pass

    def start_download(self):
        if not self.model_list:
            messagebox.showwarning("Warning", "No models loaded!")
            return
        if not self.driver_path_var.get() or not os.path.exists(self.driver_path_var.get()):
            messagebox.showwarning("Warning", "Please specify a valid ChromeDriver path!")
            return
        if not self.download_path_var.get():
            messagebox.showwarning("Warning", "Please specify a download path!")
            return

        self.is_running = True
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.stats = {'processed': 0, 'success': 0, 'failed': 0}
        self.update_statistics()
        self.progress_var.set(0)
        self.progress_label_var.set("Starting download process...")
        self.log_message("Starting BIOS download session", "info")
        self.log_message(f"Total models to process: {len(self.model_list)}", "info")
        self.download_worker()

    def stop_download(self):
        self.is_running = False
        self.log_message("Stopping download process...", "warning")
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.progress_label_var.set("Stopped by user")

    def update_statistics(self):
        self.processed_var.set(f"Processed: {self.stats['processed']}")
        self.success_var.set(f"Success: {self.stats['success']}")
        self.failed_var.set(f"Failed: {self.stats['failed']}")

    def download_worker(self):
        total_models = len(self.model_list)
        processed = 0
        success = 0
        failed = 0

        for idx, model in enumerate(self.model_list):
            if not self.is_running:
                break
            processed += 1
            self.stats['processed'] = processed
            self.progress_var.set((processed / total_models) * 100)
            self.progress_label_var.set(f"Processing {processed}/{total_models}: {model}")
            self.update_statistics()
            self.log_message(f"[{processed}/{total_models}] Processing: {model}", "info")
            try:
                before_version = get_bios_version_for_model(model, self.download_path_var.get())
                download_asus_bios(model=model, bios_download_path=self.download_path_var.get())
                after_version = get_bios_version_for_model(model, self.download_path_var.get())
                if after_version is not None and (before_version is None or after_version > before_version):
                    success += 1
                    self.stats['success'] = success
                    self.log_message(f"[{processed}/{total_models}] Success: {model}", "success")
                else:
                    failed += 1
                    self.stats['failed'] = failed
                    self.log_message(f"[{processed}/{total_models}] Failed: {model} (No new BIOS downloaded)", "error")
            except Exception as e:
                failed += 1
                self.stats['failed'] = failed
                self.log_message(f"[{processed}/{total_models}] Failed: {model} ({e})", "error")
            self.update_statistics()
            self.root.update()
        self.progress_var.set(100)
        self.progress_label_var.set("Processing complete")
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.is_running = False
        summary = (f"Processing complete!\n"
                   f"Success: {success}\n"
                   f"Failed: {failed}")
        self.log_message(summary, "info")

    def monitor_progress(self):
        # Placeholder for future threaded/async progress
        self.root.after(100, self.monitor_progress)

def main():
    root = tk.Tk()
    app = BIOSDownloaderGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()