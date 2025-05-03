import customtkinter as ctk
from tkinter import messagebox, filedialog
import requests
import browser_cookie3
import csv
import os
import threading
import re
from urllib.parse import urlparse, parse_qs
import time
import winreg
import sys
import ctypes

# Configuration
MAX_RETRIES = 2
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

def get_default_browser():
    try:
        reg_path = r'Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path) as key:
            prog_id = winreg.QueryValueEx(key, 'ProgId')[0]
        
        cmd_reg_path = fr'{prog_id}\shell\open\command'
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, cmd_reg_path) as key:
            cmd = winreg.QueryValueEx(key, None)[0]
        
        if 'chrome' in cmd.lower():
            return "Google Chrome"
        elif 'firefox' in cmd.lower():
            return "Mozilla Firefox"
        elif 'edge' in cmd.lower():
            return "Microsoft Edge"
        return "Unknown Browser"
    except Exception:
        return "Browser Detection Failed"

class DownloadManager(ctk.CTk):
    def __init__(self):
        super().__init__()
        if not is_admin():
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            sys.exit()
            
        self.title("IEEE Paper Download Manager")
        self.geometry("1200x800")
        self.current_process = None
        self.stats = {'total': 0, 'success': 0, 'failed': 0, 'skipped': 0}
        self.selected_csv = ""
        self.base_path = ""
        self.cookies = None
        self.paused = threading.Event()
        self.stopped = threading.Event()
        self._setup_ui()
        self.reset_all()

    def _setup_ui(self):
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)
        self.grid_columnconfigure(0, weight=1)

        # Main Frame
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # Browser and Cookie Section
        browser_frame = ctk.CTkFrame(self.main_frame)
        browser_frame.pack(pady=5, fill="x")
        
        ctk.CTkButton(browser_frame, text="Fetch Institutional Cookies", 
                     command=self.fetch_cookies, width=200).pack(side="left", padx=5)
        
        self.browser_status = ctk.CTkLabel(browser_frame, text="Status: Not checked", 
                                         font=("Arial", 12), anchor="w")
        self.browser_status.pack(side="left", fill="x", expand=True)

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

        # Progress Section
        self.progress_label = ctk.CTkLabel(self.main_frame, text="Ready", font=("Arial", 14))
        self.progress_label.pack(pady=5)
        
        self.progress_bar = ctk.CTkProgressBar(self.main_frame)
        self.progress_bar.pack(fill="x", padx=20, pady=5)
        self.progress_bar.set(0)

        # Stats Panel
        stats_frame = ctk.CTkFrame(self.main_frame)
        stats_frame.pack(pady=10)
        
        self.stats_labels = {
            'total': ctk.CTkLabel(stats_frame, text="Total: 0"),
            'success': ctk.CTkLabel(stats_frame, text="Success: 0", text_color="green"),
            'failed': ctk.CTkLabel(stats_frame, text="Failed: 0", text_color="red"),
            'skipped': ctk.CTkLabel(stats_frame, text="Skipped: 0", text_color="orange")
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
        
        self.start_btn = ctk.CTkButton(control_frame, text="Start Download", 
                                      command=self.start_process, width=120)
        self.start_btn.pack(side="left", padx=10)
        
        self.pause_btn = ctk.CTkButton(control_frame, text="Pause", 
                                      command=self.toggle_pause, state="disabled", width=80)
        self.pause_btn.pack(side="left", padx=10)
        
        self.stop_btn = ctk.CTkButton(control_frame, text="Stop", 
                                     command=self.stop_process, state="disabled", width=80)
        self.stop_btn.pack(side="left", padx=10)

        # Status Bar
        status_bar = ctk.CTkFrame(self, height=30, corner_radius=0)
        status_bar.grid(row=1, column=0, sticky="ew")
        ctk.CTkLabel(status_bar, text="Developed by Niraj Bhagchandani", 
                    font=("Arial", 12, "italic")).pack(side="right", padx=10)

    def fetch_cookies(self):
        try:
            browser = get_default_browser()
            required_cookies = {'JSESSIONID', 'ERIGHTS', 'xpluserinfo', 'ipCheck'}
            
            try:
                cookies = browser_cookie3.load(domain_name='ieeexplore.ieee.org')
                found_cookies = {c.name for c in cookies}
                missing = required_cookies - found_cookies
                
                if missing:
                    status_text = f"❌ Missing cookies: {', '.join(missing)}"
                    status_color = "red"
                    self.cookies = None
                else:
                    status_text = "✅ Institutional access verified"
                    status_color = "green"
                    self.cookies = {c.name: c.value for c in cookies}
                
            except PermissionError:
                status_text = "⚠️ Close browser and run as admin!"
                status_color = "orange"
                self.cookies = None
                messagebox.showwarning("Permission Needed", 
                    "1. Close all browsers completely\n"
                    "2. Right-click this app\n"
                    "3. Select 'Run as administrator'")
            
            self.browser_status.configure(
                text=f"{browser} Status: {status_text}",
                text_color=status_color
            )
            
        except Exception as e:
            self.browser_status.configure(
                text=f"Error: {str(e)}",
                text_color="red"
            )

    def browse_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
            self.selected_csv = file_path
            self.csv_entry.delete(0, "end")
            self.csv_entry.insert(0, file_path)
            
            filename = os.path.basename(file_path)
            base_name = os.path.splitext(filename)[0]
            sanitized_name = re.sub(r'[^a-zA-Z0-9_\- ]', '_', base_name)
            self.base_path = sanitized_name.replace(' ', '_')
            self.base_path_label.configure(text=f"Base Path: {self.base_path}")

    def reset_all(self):
        self.stats = {'total': 0, 'success': 0, 'failed': 0, 'skipped': 0}
        self.queue = []
        self.progress_bar.set(0)
        self.progress_label.configure(text="Ready")
        for label in self.stats_labels.values():
            label.configure(text=f"{label._text.split(':')[0]}: 0")
        for tab in self.tabs.values():
            for widget in tab.winfo_children():
                widget.destroy()
        self.paused.clear()
        self.stopped.clear()

    def create_card(self, tab_name, title, status, url):
        parent = self.tabs[tab_name]
        frame = ctk.CTkFrame(parent, corner_radius=10)
        status_color = {
            'success': "green",
            'failed': "red",
            'skipped': "orange"
        }[status]
        
        ctk.CTkLabel(frame, text="●", text_color=status_color, 
                    font=("Arial", 24)).pack(side="left", padx=10)
        
        text_frame = ctk.CTkFrame(frame, fg_color="transparent")
        ctk.CTkLabel(text_frame, text=title, font=("Arial", 12, "bold")).pack(anchor="w")
        ctk.CTkLabel(text_frame, text=url, font=("Arial", 10)).pack(anchor="w")
        text_frame.pack(side="left", fill="x", expand=True)
        
        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        if status == 'failed':
            ctk.CTkButton(btn_frame, text="Retry", width=80,
                         command=lambda: self.retry_download(url)).pack(pady=2)
            ctk.CTkButton(btn_frame, text="Open Link", width=80,
                         command=lambda: self.open_link(url)).pack(pady=2)
        else:
            ctk.CTkButton(btn_frame, text="Open", width=80,
                         command=lambda: self.open_file(url)).pack(pady=2)
        
        btn_frame.pack(side="right", padx=10)
        frame.pack(fill="x", padx=5, pady=2)

    def update_progress(self):
        processed = len([item for item in self.queue if item['processed']])
        total = len(self.queue)
        self.progress_bar.set(processed/total if total > 0 else 0)
        self.progress_label.configure(text=f"Processed {processed}/{total}")
        self.update_stats()

    def update_stats(self):
        for key, label in self.stats_labels.items():
            label.configure(text=f"{key.capitalize()}: {self.stats[key]}")

    def start_process(self):
        if not self.current_process or not self.current_process.is_alive():
            self.start_btn.configure(state="disabled")
            self.pause_btn.configure(state="normal", text="Pause")
            self.stop_btn.configure(state="normal")
            self.reset_all()
            self.current_process = threading.Thread(target=self.process_csv)
            self.current_process.start()

    def stop_process(self):
        self.stopped.set()
        self.pause_btn.configure(state="disabled")
        self.stop_btn.configure(state="disabled")
        self.start_btn.configure(state="normal")

    def toggle_pause(self):
        if not self.paused.is_set():
            self.paused.set()
            self.pause_btn.configure(text="Resume")
        else:
            self.paused.clear()
            self.pause_btn.configure(text="Pause")

    def process_csv(self):
        if not self.selected_csv:
            messagebox.showerror("Error", "No CSV file selected!")
            self.after(0, lambda: self.start_btn.configure(state="normal"))
            return

        self.queue = []
        try:
            with open(self.selected_csv, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                headers = next(reader)
                url_index = headers.index('PDF Link')
                title_index = headers.index('Document Title')
                category_index = headers.index('Document Identifier')

                os.makedirs(self.base_path, exist_ok=True)

                for row in reader:
                    if len(row) < max(url_index, title_index, category_index):
                        continue
                    
                    entry = {
                        'title': row[title_index].strip(),
                        'url': row[url_index].strip(),
                        'category': row[category_index].strip(),
                        'processed': False,
                        'status': 'pending'
                    }
                    self.queue.append(entry)
                    self.stats['total'] += 1

            self.after(0, self.update_stats)
            
            for entry in self.queue:
                if self.stopped.is_set():
                    break
                while self.paused.is_set():
                    time.sleep(0.1)
                if not entry['url']:
                    continue
                
                output_path = self.get_output_path(entry)
                
                if os.path.exists(output_path):
                    self.stats['skipped'] += 1
                    entry['status'] = 'skipped'
                    entry['processed'] = True
                    self.after(0, lambda e=entry, p=output_path: 
                             self.create_card("skipped", e['title'], 'skipped', p))
                    self.after(0, self.update_progress)
                    continue
                
                self.download_paper(entry)
                entry['processed'] = True
                self.after(0, self.update_progress)

        except Exception as e:
            self.handle_error(str(e))
        finally:
            self.after(0, lambda: self.start_btn.configure(state="normal"))
            self.after(0, lambda: self.pause_btn.configure(state="disabled"))
            self.after(0, lambda: self.stop_btn.configure(state="disabled"))

    def get_output_path(self, entry):
        folder = os.path.join(self.base_path, entry['category'])
        os.makedirs(folder, exist_ok=True)
        filename = re.sub(r'[^a-zA-Z0-9_\- ]', '_', entry['title'])[:100]
        return os.path.join(folder, f"{filename}.pdf")

    def download_paper(self, entry):
        for attempt in range(MAX_RETRIES):
            try:
                pdf_url = self.get_pdf_url(entry['url'])
                if not pdf_url:
                    raise ValueError("PDF URL not found")
                
                # Use institutional cookies if available
                cookies = self.cookies if self.cookies else {}
                
                response = requests.get(
                    pdf_url,
                    headers=HEADERS,
                    cookies=cookies,
                    timeout=30
                )
                response.raise_for_status()
                
                if not self.is_valid_pdf(response.content):
                    raise ValueError("Invalid PDF content")
                
                output_path = self.get_output_path(entry)
                
                with open(output_path, 'wb') as f:
                    f.write(response.content)
                
                self.stats['success'] += 1
                self.after(0, lambda e=entry, p=output_path: 
                         self.create_card("downloaded", e['title'], 'success', p))
                return

            except Exception as e:
                print(f"Attempt {attempt+1} failed: {str(e)}")
        
        self.stats['failed'] += 1
        self.after(0, lambda e=entry: 
                 self.create_card("failed", e['title'], 'failed', e['url']))

    def get_pdf_url(self, stamp_url):
        try:
            parsed = urlparse(stamp_url)
            params = parse_qs(parsed.query)
            arnumber = params.get('arnumber', [''])[0]
            if arnumber:
                return f"https://ieeexplore.ieee.org/stampPDF/getPDF.jsp?arnumber={arnumber}"
        except Exception as e:
            print(f"URL parsing error: {str(e)}")
        return None

    def is_valid_pdf(self, content):
        if not content.startswith(b'%PDF'):
            return False
        if b'%%EOF' not in content[-1024:]:
            return False
        error_indicators = [
            b'<title>Access Denied</title>',
            b'<h1>403 Forbidden</h1>',
            b'CAPTCHA_CHALLENGE'
        ]
        return not any(indicator in content for indicator in error_indicators)

    def open_file(self, path):
        if os.path.exists(path):
            os.startfile(path)
        else:
            messagebox.showerror("Error", "File not found")

    def open_link(self, url):
        import webbrowser
        webbrowser.open(url)

    def retry_download(self, url):
        entry = next((item for item in self.queue if item['url'] == url), None)
        if entry:
            entry['processed'] = False
            entry['status'] = 'pending'
            self.download_paper(entry)
            self.update_progress()

    def handle_error(self, message):
        self.after(0, lambda: messagebox.showerror("Error", message))

if __name__ == "__main__":
    app = DownloadManager()
    app.mainloop()
