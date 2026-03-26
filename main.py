import subprocess
import os
import sys
import socket
import threading
import time
from datetime import datetime
import schedule
import win32event
import win32api
from winerror import ERROR_ALREADY_EXISTS
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
from pystray import MenuItem as item, Icon
from PIL import Image, ImageDraw
import configparser

# --- Single Instance Lock using a Mutex ---
class SingleInstance:
    """Ensures only one instance of the application can run."""
    def __init__(self, name):
        self.mutex_name = name
        self.mutex = win32event.CreateMutex(None, 1, self.mutex_name)
        self.last_error = win32api.GetLastError()

    def is_running(self):
        return self.last_error == ERROR_ALREADY_EXISTS

    def __del__(self):
        if self.mutex:
            win32api.CloseHandle(self.mutex)

# --- Main Application Class ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("API Connection Monitor")
        self.root.geometry("600x450")
        self.root.resizable(False, False)
        
        self.style = ttk.Style(self.root)
        self.style.theme_use('clam')

        self.scheduler_thread = None
        self.stop_scheduler = threading.Event()
        self.icon = None
        
        # --- Get Documents folder path ---
        self.documents_path = os.path.join(os.path.expanduser('~'), 'Documents')
        self.default_log_path = os.path.join(self.documents_path, 'API Latency Logs')
        
        # --- Load Configuration ---
        self.load_config()

        # --- UI Elements ---
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Configuration Section
        config_frame = ttk.LabelFrame(self.main_frame, text="Configuration (Loaded from config.ini)", padding="10")
        config_frame.pack(fill=tk.X, pady=5)
        config_frame.columnconfigure(1, weight=1)

        ttk.Label(config_frame, text="API Endpoint:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.host_entry = ttk.Entry(config_frame)
        self.host_entry.insert(0, self.config.get('Settings', 'endpoint', fallback='tmgposapi.themall.co.th'))
        self.host_entry.grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=5)

        ttk.Label(config_frame, text="Schedule Times (HH:MM):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        
        time_frame = ttk.Frame(config_frame)
        time_frame.grid(row=1, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        
        # Modified fallback from 3 times to 5 times
        times = [t.strip() for t in self.config.get('Settings', 'schedule_times', fallback='08:00,12:00,15:00,17:00,19:00').split(',')]
        
        # Time Entry 1
        self.time1_entry = ttk.Entry(time_frame, width=8)
        self.time1_entry.insert(0, times[0] if len(times) > 0 else "")
        self.time1_entry.pack(side=tk.LEFT, padx=(0, 3))
        # Time Entry 2
        self.time2_entry = ttk.Entry(time_frame, width=8)
        self.time2_entry.insert(0, times[1] if len(times) > 1 else "")
        self.time2_entry.pack(side=tk.LEFT, padx=3)
        # Time Entry 3
        self.time3_entry = ttk.Entry(time_frame, width=8)
        self.time3_entry.insert(0, times[2] if len(times) > 2 else "")
        self.time3_entry.pack(side=tk.LEFT, padx=3)
        # Time Entry 4 (Added)
        self.time4_entry = ttk.Entry(time_frame, width=8)
        self.time4_entry.insert(0, times[3] if len(times) > 3 else "")
        self.time4_entry.pack(side=tk.LEFT, padx=3)
        # Time Entry 5 (Added)
        self.time5_entry = ttk.Entry(time_frame, width=8)
        self.time5_entry.insert(0, times[4] if len(times) > 4 else "")
        self.time5_entry.pack(side=tk.LEFT, padx=3)
        
        ttk.Label(config_frame, text="Log File Path:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.log_path_entry = ttk.Entry(config_frame)
        self.log_path_entry.insert(0, self.config.get('Settings', 'log_path', fallback=self.default_log_path))
        self.log_path_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        
        self.browse_button = ttk.Button(config_frame, text="Browse...", command=self.select_log_folder)
        self.browse_button.grid(row=2, column=2, padx=5, pady=5)

        # Control Section
        control_frame = ttk.Frame(self.main_frame, padding="10")
        control_frame.pack(fill=tk.X, pady=5)

        self.start_button = ttk.Button(control_frame, text="Start Monitoring", command=self.start_monitoring)
        self.start_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        self.stop_button = ttk.Button(control_frame, text="Stop Monitoring", command=self.stop_monitoring, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        
        # Log Section
        log_frame = ttk.LabelFrame(self.main_frame, text="Status Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10, relief="flat")
        self.log_area.pack(fill=tk.BOTH, expand=True)
        
        self.log("API Connection Monitor Initialized.")
        self.log("Reading settings from config.ini...")
        self.log(f"Default log path set to Documents folder.")


        self.root.protocol("WM_DELETE_WINDOW", self.hide_to_tray)
        self.setup_tray_icon_thread()

    def load_config(self):
        """Reads settings from the config.ini file."""
        self.config = configparser.ConfigParser()
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))
        
        config_path = os.path.join(application_path, 'config.ini')
        
        if not os.path.exists(config_path):
            self.log(f"Warning: config.ini not found. Using default values.")
        else:
            self.config.read(config_path)

    def select_log_folder(self):
        folder_selected = filedialog.askdirectory(initialdir=self.documents_path)
        if folder_selected:
            self.log_path_entry.delete(0, tk.END)
            self.log_path_entry.insert(0, folder_selected)

    def log(self, message):
        self.root.after(0, self._log_message, message)

    def _log_message(self, message):
        now = datetime.now().strftime("%H:%M:%S")
        self.log_area.insert(tk.END, f"[{now}] {message}\n")
        self.log_area.see(tk.END)

    def start_monitoring(self):
        self.host = self.host_entry.get()
        self.log_folder = self.log_path_entry.get()
        
        schedule_times_str = []
        # Added time4_entry and time5_entry
        time_entries = [self.time1_entry.get(), self.time2_entry.get(), self.time3_entry.get(), self.time4_entry.get(), self.time5_entry.get()]

        for t in time_entries:
            if t.strip():
                try:
                    time.strptime(t.strip(), '%H:%M')
                    schedule_times_str.append(t.strip())
                except ValueError:
                    self.log(f"Error: Invalid time format '{t}'. Please use HH:MM (24-hour format).")
                    return

        if not schedule_times_str:
            self.log("Error: Please enter at least one valid schedule time.")
            return

        # Added self.time4_entry, self.time5_entry to disabled widget list
        for widget in [self.start_button, self.host_entry, self.time1_entry, self.time2_entry, self.time3_entry, self.time4_entry, self.time5_entry, self.log_path_entry, self.browse_button]:
            widget.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)

        self.log(f"Monitoring started for {self.host}. Scheduled times: {', '.join(schedule_times_str)}")
        
        schedule.clear()
        for scheduled_time in schedule_times_str:
            schedule.every().day.at(scheduled_time).do(self.run_diagnostics_thread)
        
        self.stop_scheduler.clear()
        self.scheduler_thread = threading.Thread(target=self.run_scheduler, daemon=True)
        self.scheduler_thread.start()

        threading.Timer(2.0, self.run_diagnostics_thread).start()

    def stop_monitoring(self):
        self.stop_scheduler.set()
        schedule.clear()
        
        # Added self.time4_entry, self.time5_entry to normal widget list
        for widget in [self.start_button, self.host_entry, self.time1_entry, self.time2_entry, self.time3_entry, self.time4_entry, self.time5_entry, self.log_path_entry, self.browse_button]:
            widget.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        
        self.log("Monitoring stopped by user.")

    def run_scheduler(self):
        while not self.stop_scheduler.is_set():
            schedule.run_pending()
            time.sleep(1)

    def run_diagnostics_thread(self):
        threading.Thread(target=self.run_diagnostics, daemon=True).start()

    def run_diagnostics(self):
        """
        Runs diagnostics by executing commands directly and capturing their output within Python.
        """
        self.log(f"Running diagnostics for {self.host}...")
        computer_name = socket.gethostname()
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        file_name = f"{computer_name}_{timestamp}.txt"
        
        try:
            os.makedirs(self.log_folder, exist_ok=True)
        except OSError as e:
            self.log(f"CRITICAL ERROR: Could not create log directory '{self.log_folder}'. Please check permissions. Error: {e}")
            return
            
        full_path = os.path.join(self.log_folder, file_name)

        try:
            with open(full_path, 'w', encoding='utf-8', errors='ignore') as f:
                f.write(f"COMPREHENSIVE NETWORK DIAGNOSTIC REPORT\n")
                f.write(f"=================================================\n")
                f.write(f"Report generated on: {datetime.now().strftime('%Y-%m-%d at %H:%M:%S')}\n")
                f.write(f"Target Host: {self.host}\n\n")

                # --- 1. Ping Local DNS Servers ---
                f.write("\n===== 1. PING LOCAL DNS SERVERS =====\n\n")
                f.write("--- Finding and Pinging configured DNS servers ---\n")
                f.flush()
                ps_get_dns_cmd = "(Get-DnsClientServerAddress -AddressFamily IPv4).ServerAddresses"
                dns_process = subprocess.run(
                    ['powershell', '-Command', ps_get_dns_cmd],
                    capture_output=True, text=True, encoding='utf-8', errors='ignore',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                if dns_process.stdout:
                    dns_servers = [s.strip() for s in dns_process.stdout.strip().split('\n') if s.strip()]
                    f.write(f"Found DNS Servers: {', '.join(dns_servers)}\n\n")
                    for server in dns_servers:
                        f.write(f"--- Pinging DNS Server: {server} (4 packets) ---\n")
                        f.flush()
                        ping_dns_process = subprocess.run(
                            ['ping', '-n', '4', server],
                            capture_output=True, text=True, encoding='utf-8', errors='ignore',
                            creationflags=subprocess.CREATE_NO_WINDOW
                        )
                        f.write(ping_dns_process.stdout)
                        f.write(ping_dns_process.stderr)
                        f.write("\n")
                else:
                    f.write("Could not automatically determine DNS servers.\n")
                    f.write(dns_process.stderr)


                # --- 2. Traceroute ---
                f.write("\n\n===== 2. TRACEROUTE TO VIEW NETWORK PATH =====\n\n")
                f.flush()
                process = subprocess.run(
                    ['tracert', self.host], 
                    capture_output=True, text=True, encoding='utf-8', errors='ignore',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                f.write(process.stdout)
                f.write(process.stderr)

                # --- 3. DNS Test & Ping Destination ---
                f.write("\n\n===== 3. DNS LATENCY & RESOLUTION TEST =====\n\n")
                f.write("--- Measuring time to resolve DNS name ---\n")
                f.flush()
                ps_command = f"Measure-Command {{Resolve-DnsName {self.host} -Type A -ErrorAction SilentlyContinue}}"
                process = subprocess.run(
                    ['powershell', '-ExecutionPolicy', 'Bypass', '-Command', ps_command],
                    capture_output=True, text=True, encoding='utf-8', errors='ignore',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                f.write(process.stdout)
                f.write(process.stderr)
                
                f.write("\n--- Pinging destination IP to measure latency (4 packets) ---\n")
                f.flush()
                process = subprocess.run(
                    ['ping', '-n', '4', self.host],
                    capture_output=True, text=True, encoding='utf-8', errors='ignore',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                f.write(process.stdout)
                f.write(process.stderr)

                # --- 4. Curl Test ---
                f.write("\n\n===== 4. CURL API CONNECTION TIMING =====\n\n")
                f.flush()
                curl_format = "DNS Lookup:      %{time_namelookup}s\\nTCP Connection:  %{time_connect}s\\nSSL Handshake:   %{time_appconnect}s\\nTTFB:              %{time_starttransfer}s\\nTotal Time:      %{time_total}s\\n"
                process = subprocess.run(
                    ['curl', '-s', '-w', curl_format, '-o', 'nul', f"https://{self.host}"],
                    capture_output=True, text=True, encoding='utf-8', errors='ignore',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                f.write(process.stdout)
                f.write(process.stderr)

            self.log(f"Diagnostics complete. Log saved to {full_path}")

        except Exception as e:
            self.log(f"A critical error occurred while writing the log file: {e}")


    # --- System Tray Logic ---
    def create_icon_image(self):
        width = 64
        height = 64
        color1 = (20, 20, 120)
        color2 = (80, 80, 220)
        image = Image.new('RGB', (width, height), color1)
        dc = ImageDraw.Draw(image)
        dc.rectangle((width // 2, 0, width, height // 2), fill=color2)
        dc.rectangle((0, height // 2, width // 2, height), fill=color2)
        return image
        
    def setup_tray_icon_thread(self):
        image = self.create_icon_image()
        menu = (item('Show', self.show_from_tray, default=True), item('Exit', self.exit_app))
        self.icon = Icon("API_Monitor", image, "API Connection Monitor", menu)
        self.icon.visible = False
        
        tray_thread = threading.Thread(target=self.icon.run, daemon=True)
        tray_thread.start()

    def hide_to_tray(self):
        self.log("Minimizing to system tray...")
        self.root.withdraw()
        self.icon.visible = True

    def show_from_tray(self):
        self.icon.visible = False
        self.root.after(0, self.root.deiconify)

    def exit_app(self):
        self.icon.stop()
        self.stop_monitoring()
        self.root.quit()

# --- Main Application Execution ---
if __name__ == '__main__':
    instance_name = "Global\\API_Monitor_UI_Mutex_v11" # Incremented version
    instance = SingleInstance(instance_name)
    if instance.is_running():
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Application Already Running", 
                               "An instance of the API Connection Monitor is already running.")
        root.destroy()
        sys.exit(1)
        
    root = tk.Tk()
    app = App(root)
    
    root.withdraw()
    app.icon.visible = True
    
    root.after(2000, app.start_monitoring)
    
    root.mainloop()
