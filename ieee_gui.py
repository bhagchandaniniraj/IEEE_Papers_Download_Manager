import customtkinter as ctk
from tkinter import messagebox, filedialog
import requests
import browser_cookie3
import csv
import os
import threading
import re
import time
import sys
import ctypes
import psutil
import random
from urllib.parse import urlparse, parse_qs
from concurrent.futures import ThreadPoolExecutor

# Configuration
MAX_RETRIES = 2
MAX_WORKERS = 5
DELAY_RANGE = (3, 8)  # Random delay between downloads in seconds
REQUIRED_COOKIES = {'JSESSIONID', 'ERIGHTS', 'xpluserinfo', 'ipCheck'}
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
}

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def kill_browser_processes():
    browsers = ['msedge', 'chrome', 'firefox', 'iexplore']
    for proc in psutil.process_iter():
        try:
            if any(browser in proc.name().lower() for browser in browsers):
                proc.kill()
        except psutil.AccessDenied:
            pass

class DownloadManager(ctk.CTk):
    def __init__(self):
        if not is_admin():
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            sys.exit()
            
        super().__init__()
        self.title("IEEE Paper Download Manager")
        self.geometry("1200x800")
        self.current_process = None
        self.stats = {
            'total': 0, 'success': 0, 'failed': 0, 'skipped': 0,
            'total_size': 0, 'success_size': 0, 'skipped_size': 0
        }
        self.selected_csv = ""
        self.base_path = ""
        self.cookies = None
        self.paused = threading.Event()
        self.stopped = threading.Event()
        self.download_queue = []
        self.executor = ThreadPoolExecutor(max_workers=MAX_WORKERS)
        self.next_document = {"title": "", "url": ""}
        self.countdown = 0
        self.process = psutil.Process(os.getpid())
        self._setup_ui()
        self.reset_all()
        self.start_memory_monitor()

    def _setup_ui(self):
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Main Frame
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # Browser and Cookie Section
        browser_frame = ctk.CTkFrame(self.main_frame)
        browser_frame.pack(pady=5, fill="x")
        
        ctk.CTkButton(browser_frame, text="Fetch Institutional Cookies", 
                     command=self.fetch_cookies, width=220).pack(side="left", padx=5)
        
        self.cookie_text = ctk.CTkTextbox(browser_frame, height=80, wrap="word")
        self.cookie_text.pack(side="left", fill="x", expand=True, padx=5)
        self.cookie_text.insert("1.0", "No cookies fetched")

        # File Selection Section
        file_frame = ctk.CTkFrame(self.main_frame)
        file_frame.pack(pady=10, fill="x")
        
        self.csv_entry = ctk.CTkEntry(file_frame, width=500)
        self.csv_entry.pack(side="left", padx=5)
        
        ctk.CTkButton(file_frame, text="Browse CSV", width=100,
                     command=self.browse_csv).pack(side="left", padx=5)

        # Base Path Display
        self.base_path_label = ctk.CTkLabel(self.main_frame, text="Base Path: Not Set", 
                                          anchor="w", font=("Arial", 12))
        self.base_path_label.pack(fill="x", padx=20, pady=5)

        # Next Download Info
        self.next_doc_frame = ctk.CTkFrame(self.main_frame)
        self.next_doc_frame.pack(pady=5, fill="x")
        
        ctk.CTkLabel(self.next_doc_frame, text="Next Download:", 
                    font=("Arial", 12, "bold")).pack(side="left", padx=5)
        self.next_doc_label = ctk.CTkLabel(self.next_doc_frame, text="No pending downloads",
                                         wraplength=1000, justify="left")
        self.next_doc_label.pack(side="left", fill="x", expand=True)
        self.countdown_label = ctk.CTkLabel(self.next_doc_frame, text="")
        self.countdown_label.pack(side="right", padx=10)

        # Progress Section
        self.progress_label = ctk.CTkLabel(self.main_frame, text="Ready", font=("Arial", 14))
        self.progress_label.pack(pady=5)
        self.progress_bar = ctk.CTkProgressBar(self.main_frame)
        self.progress_bar.pack(fill="x", padx=20, pady=5)

        # Stats Panel
        stats_frame = ctk.CTkFrame(self.main_frame)
        stats_frame.pack(pady=10)
        self.stats_labels = {
            'total': ctk.CTkLabel(stats_frame, text="Total: 0"),
            'success': ctk.CTkLabel(stats_frame, text="Success: 0 (0 MB)", text_color="green"),
            'failed': ctk.CTkLabel(stats_frame, text="Failed: 0", text_color="red"),
            'skipped': ctk.CTkLabel(stats_frame, text="Skipped: 0 (0 MB)", text_color="orange")
        }
        for label in self.stats_labels.values():
            label.pack(side="left", padx=10)

        # Content Tabs
        self.tab_view = ctk.CTkTabview(self.main_frame)
        self.tab_view.pack(fill="both", expand=True, padx=10, pady=10)
        self.tabs = {}
        for name in ["Downloaded", "Failed", "Skipped"]:
            tab = self.tab_view.add(name)
            scroll_frame = ctk.CTkScrollableFrame(tab)
            scroll_frame.pack(fill="both", expand=True)
            self.tabs[name.lower()] = scroll_frame

        # Control Buttons
        control_frame = ctk.CTkFrame(self.main_frame)
        control_frame.pack(pady=10)
        self.start_btn = ctk.CTkButton(control_frame, text="Start Download", command=self.start_process)
        self.start_btn.pack(side="left", padx=10)
        self.pause_btn = ctk.CTkButton(control_frame, text="Pause", command=self.toggle_pause, state="disabled")
        self.pause_btn.pack(side="left", padx=10)
        self.stop_btn = ctk.CTkButton(control_frame, text="Stop", command=self.stop_process, state="disabled")
        self.stop_btn.pack(side="left", padx=10)

        # Status Bar
        self.status_bar = ctk.CTkFrame(self, height=30)
        self.status_bar.grid(row=1, column=0, sticky="ew")

        self.ram_label = ctk.CTkLabel(
            self.status_bar, 
            text="App RAM: 0 MB | System RAM: 0 MB", 
            font=("Arial", 12)
        )
        self.ram_label.pack(side="left", padx=10)

        self.credit_label = ctk.CTkLabel(
            self.status_bar, 
            text="Application Developed by Niraj Bhagchandani", 
            font=("Arial", 12, "italic")
        )
        self.credit_label.pack(side="right", padx=10)

    def start_memory_monitor(self):
        def monitor():
            while True:
                mem = self.process.memory_info().rss
                total_mem = psutil.virtual_memory().total
                self.after(0, lambda: self.ram_label.configure(
                    text=f"App RAM: {mem//1024//1024} MB | System RAM: {total_mem//1024//1024} MB"
                ))
                time.sleep(1)
        threading.Thread(target=monitor, daemon=True).start()

    def fetch_cookies(self):
        try:
            kill_browser_processes()
            time.sleep(2)
            cookies = browser_cookie3.load(domain_name='ieeexplore.ieee.org')
            found = {c.name for c in cookies}
            missing = REQUIRED_COOKIES - found
            
            if missing:
                self.cookie_text.delete("1.0", "end")
                self.cookie_text.insert("1.0", f"Missing cookies: {', '.join(missing)}")
                self.cookies = None
            else:
                self.cookies = {c.name: c.value for c in cookies}
                self.cookie_text.delete("1.0", "end")
                self.cookie_text.insert("1.0", "Cookies fetched successfully!")
        except Exception as e:
            messagebox.showerror("Cookie Error", f"{str(e)}\n\nTry manual cookie export if persists")

    def browse_csv(self):
        try:
            file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
            if file_path:
                self.selected_csv = file_path
                self.csv_entry.delete(0, "end")
                self.csv_entry.insert(0, file_path)
                filename = os.path.basename(file_path)
                base_name = os.path.splitext(filename)[0]
                self.base_path = re.sub(r'[^a-zA-Z0-9_\- ]', '_', base_name)
                self.base_path_label.configure(text=f"Base Path: {self.base_path}")
        except Exception as e:
            messagebox.showerror("File Error", f"Invalid CSV file:\n{str(e)}")

    def reset_all(self):
        self.stats = {'total':0,'success':0,'failed':0,'skipped':0,'total_size':0,'success_size':0,'skipped_size':0}
        self.download_queue = []
        for tab in self.tabs.values():
            for widget in tab.winfo_children():
                widget.destroy()
        self.progress_bar.set(0)
        self.progress_label.configure(text="Ready")
        self.next_doc_label.configure(text="No pending downloads")
        self.countdown_label.configure(text="")
        self.update_stats()

    def create_card(self, tab_name, title, status, url, size=0):
        parent = self.tabs[tab_name]
        frame = ctk.CTkFrame(parent, corner_radius=10)
        status_color = {'success': "#2ecc71", 'failed': "#e74c3c", 'skipped': "#f1c40f"}[status]
        
        ctk.CTkLabel(frame, text="●", text_color=status_color, font=("Arial", 24)).pack(side="left", padx=10)
        
        text_frame = ctk.CTkFrame(frame, fg_color="transparent")
        ctk.CTkLabel(text_frame, text=title, font=("Arial", 12, "bold")).pack(anchor="w")
        ctk.CTkLabel(text_frame, text=f"{url}\nSize: {size/1024/1024:.2f} MB", 
                    font=("Arial", 10)).pack(anchor="w")
        text_frame.pack(side="left", fill="x", expand=True)
        
        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        if status == 'failed':
            ctk.CTkButton(btn_frame, text="Retry", width=80,
                         command=lambda: self.retry_download(url)).pack(pady=2)
            ctk.CTkButton(btn_frame, text="Open Link", width=80,
                         command=lambda: webbrowser.open(url)).pack(pady=2)
        else:
            ctk.CTkButton(btn_frame, text="Open", width=80,
                         command=lambda: os.startfile(url)).pack(pady=2)
        
        btn_frame.pack(side="right", padx=10)
        frame.pack(fill="x", padx=5, pady=2)

    def update_stats(self):
        self.stats_labels['total'].configure(text=f"Total: {self.stats['total']}")
        self.stats_labels['success'].configure(
            text=f"Success: {self.stats['success']} ({self.stats['success_size']//1024//1024} MB)")
        self.stats_labels['failed'].configure(text=f"Failed: {self.stats['failed']}")
        self.stats_labels['skipped'].configure(
            text=f"Skipped: {self.stats['skipped']} ({self.stats['skipped_size']//1024//1024} MB)")

    def start_process(self):
        if not self.executor._shutdown:
            self.reset_all()
            self.start_btn.configure(state="disabled")
            self.pause_btn.configure(state="normal")
            self.stop_btn.configure(state="normal")
            threading.Thread(target=self.process_csv, daemon=True).start()

    def stop_process(self):
        self.stopped.set()
        self.executor.shutdown(wait=False)
        self.start_btn.configure(state="normal")
        self.pause_btn.configure(state="disabled")
        self.stop_btn.configure(state="disabled")

    def toggle_pause(self):
        if not self.paused.is_set():
            self.paused.set()
            self.pause_btn.configure(text="Resume")
        else:
            self.paused.clear()
            self.pause_btn.configure(text="Pause")

    def process_csv(self):
        try:
            with open(self.selected_csv, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                os.makedirs(self.base_path, exist_ok=True)
                
                self.download_queue = []
                for row in reader:
                    if self.stopped.is_set():
                        break
                    entry = {
                        'title': row['Document Title'].strip(),
                        'url': row['PDF Link'].strip(),
                        'category': row['Document Identifier'].strip(),
                        'processed': False,
                        'status': 'pending'
                    }
                    self.download_queue.append(entry)
                    self.stats['total'] += 1

                for index, entry in enumerate(self.download_queue):
                    if self.stopped.is_set():
                        break

                    output_path = self.get_output_path(entry)
                    
                    # Skip existing files immediately
                    if os.path.exists(output_path):
                        self.stats['skipped'] += 1
                        self.stats['skipped_size'] += os.path.getsize(output_path)
                        entry['status'] = 'skipped'
                        entry['processed'] = True
                        self.after(0, lambda e=entry, p=output_path: self.create_card(
                            "skipped", e['title'], 'skipped', p, os.path.getsize(p)))
                        self.after(0, self.update_progress)
                        continue

                    # Show next download preview
                    next_info = "No more pending downloads"
                    if index + 1 < len(self.download_queue):
                        next_entry = self.download_queue[index + 1]
                        next_info = f"Next: {next_entry['title']}\nURL: {next_entry['url']}"
                    self.after(0, lambda: self.next_doc_label.configure(text=next_info))

                    # Apply random delay
                    delay = random.randint(*DELAY_RANGE)
                    for sec in range(delay, 0, -1):
                        if self.stopped.is_set():
                            break
                        while self.paused.is_set():
                            time.sleep(0.1)
                        self.after(0, lambda s=sec: self.countdown_label.configure(
                            text=f"Starting in: {s}s"
                        ))
                        time.sleep(1)
                    self.after(0, lambda: self.countdown_label.configure(text=""))

                    # Process download
                    self.download_paper(entry)
                    entry['processed'] = True
                    self.after(0, self.update_progress)

        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.after(0, self.stop_process)

    def get_output_path(self, entry):
        folder = os.path.join(self.base_path, entry['category'])
        os.makedirs(folder, exist_ok=True)
        filename = re.sub(r'[^a-zA-Z0-9_\- ]', '_', entry['title'])[:100] + ".pdf"
        return os.path.join(folder, filename)

    def download_paper(self, entry):
        for attempt in range(MAX_RETRIES):
            try:
                pdf_url = self.get_pdf_url(entry['url'])
                if not pdf_url:
                    raise ValueError("PDF URL not found")
                
                response = requests.get(
                    pdf_url,
                    headers=HEADERS,
                    cookies=self.cookies,
                    timeout=30,
                    stream=True
                )
                response.raise_for_status()
                
                if not self.is_valid_pdf(response.content):
                    raise ValueError("Invalid PDF content")
                
                output_path = self.get_output_path(entry)
                file_size = int(response.headers.get('content-length', 0))
                
                with open(output_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        if self.stopped.is_set():
                            return
                        f.write(chunk)
                
                self.stats['success'] += 1
                self.stats['success_size'] += file_size
                self.after(0, lambda e=entry, p=output_path: self.create_card(
                    "downloaded", e['title'], 'success', p, file_size))
                return

            except Exception as e:
                print(f"Attempt {attempt+1} failed: {str(e)}")
        
        self.stats['failed'] += 1
        self.after(0, lambda e=entry: self.create_card("failed", e['title'], 'failed', e['url']))

    def get_pdf_url(self, stamp_url):
        try:
            parsed = urlparse(stamp_url)
            params = parse_qs(parsed.query)
            arnumber = params.get('arnumber', [''])[0]
            return f"https://ieeexplore.ieee.org/stampPDF/getPDF.jsp?arnumber={arnumber}" if arnumber else None
        except:
            return None

    def is_valid_pdf(self, content):
        return content.startswith(b'%PDF') and b'%%EOF' in content[-1024:]

    def retry_download(self, url):
        entry = next((e for e in self.download_queue if e['url'] == url), None)
        if entry:
            entry['processed'] = False
            entry['status'] = 'pending'
            self.executor.submit(self.download_paper, entry)
            self.update_progress()

    def update_progress(self):
        processed = sum(1 for e in self.download_queue if e['processed'])
        total = len(self.download_queue)
        self.progress_bar.set(processed/total if total > 0 else 0)
        self.progress_label.configure(text=f"Processed {processed}/{total}")
        self.update_stats()

if __name__ == "__main__":
    try:
        app = DownloadManager()
        app.mainloop()
    except Exception as e:
        messagebox.showerror("Fatal Error", f"Application failed to start:\n{str(e)}")
