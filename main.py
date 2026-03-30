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
import re

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
        self.root.title("API Connection Monitor & MTR Dashboard")
        self.root.geometry("700x550")
        self.root.resizable(False, False)
        
        self.style = ttk.Style(self.root)
        self.style.theme_use('clam')

        self.scheduler_thread = None
        self.stop_scheduler = threading.Event()
        self.icon = None
        
        # MTR Variables
        self.mtr_running = False
        self.mtr_data = {}
        self.mtr_threads = []
        self.ping_history = []
        self.max_history = 60
        
        # --- Get Documents folder path ---
        self.documents_path = os.path.join(os.path.expanduser('~'), 'Documents')
        self.default_log_path = os.path.join(self.documents_path, 'API Latency Logs')
        
        # --- Load Configuration ---
        self.load_config()

        # --- Setup Notebook (Tabs) ---
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.tab1 = ttk.Frame(self.notebook)
        self.tab2 = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab1, text="📅 Scheduled Monitor")
        self.notebook.add(self.tab2, text="📈 Live MTR Dashboard")

        self.setup_tab1()
        self.setup_tab2()

        self.root.protocol("WM_DELETE_WINDOW", self.hide_to_tray)
        self.setup_tray_icon_thread()

    def setup_tab1(self):
        # Configuration Section
        config_frame = ttk.LabelFrame(self.tab1, text="Configuration (Loaded from config.ini)", padding="10")
        config_frame.pack(fill=tk.X, pady=5)
        config_frame.columnconfigure(1, weight=1)

        ttk.Label(config_frame, text="API Endpoint:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.host_entry = ttk.Entry(config_frame)
        self.host_entry.insert(0, self.config.get('Settings', 'endpoint', fallback='tmgposapi.themall.co.th'))
        self.host_entry.grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=5)

        ttk.Label(config_frame, text="Schedule Times:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        
        time_frame = ttk.Frame(config_frame)
        time_frame.grid(row=1, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        
        times = [t.strip() for t in self.config.get('Settings', 'schedule_times', fallback='08:00,12:00,15:00,17:00,19:00').split(',')]
        
        self.time1_entry = ttk.Entry(time_frame, width=8); self.time1_entry.insert(0, times[0] if len(times) > 0 else ""); self.time1_entry.pack(side=tk.LEFT, padx=(0, 3))
        self.time2_entry = ttk.Entry(time_frame, width=8); self.time2_entry.insert(0, times[1] if len(times) > 1 else ""); self.time2_entry.pack(side=tk.LEFT, padx=3)
        self.time3_entry = ttk.Entry(time_frame, width=8); self.time3_entry.insert(0, times[2] if len(times) > 2 else ""); self.time3_entry.pack(side=tk.LEFT, padx=3)
        self.time4_entry = ttk.Entry(time_frame, width=8); self.time4_entry.insert(0, times[3] if len(times) > 3 else ""); self.time4_entry.pack(side=tk.LEFT, padx=3)
        self.time5_entry = ttk.Entry(time_frame, width=8); self.time5_entry.insert(0, times[4] if len(times) > 4 else ""); self.time5_entry.pack(side=tk.LEFT, padx=3)
        
        ttk.Label(config_frame, text="Log File Path:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.log_path_entry = ttk.Entry(config_frame)
        self.log_path_entry.insert(0, self.config.get('Settings', 'log_path', fallback=self.default_log_path))
        self.log_path_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        
        self.browse_button = ttk.Button(config_frame, text="Browse...", command=self.select_log_folder)
        self.browse_button.grid(row=2, column=2, padx=5, pady=5)

        # Control Section
        control_frame = ttk.Frame(self.tab1, padding="10")
        control_frame.pack(fill=tk.X, pady=5)

        self.start_button = ttk.Button(control_frame, text="Start Scheduled Monitor", command=self.start_monitoring)
        self.start_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        self.stop_button = ttk.Button(control_frame, text="Stop Scheduled Monitor", command=self.stop_monitoring, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        
        # Log Section
        log_frame = ttk.LabelFrame(self.tab1, text="Status Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10, relief="flat")
        self.log_area.pack(fill=tk.BOTH, expand=True)
        
        self.log("API Connection Monitor Initialized.")

    def setup_tab2(self):
        # MTR Controls
        ctrl_frame = ttk.Frame(self.tab2, padding="5")
        ctrl_frame.pack(fill=tk.X)
        
        ttk.Label(ctrl_frame, text="Target Host:").pack(side=tk.LEFT, padx=5)
        self.mtr_target = ttk.Entry(ctrl_frame, width=30)
        self.mtr_target.insert(0, self.config.get('Settings', 'endpoint', fallback='tmgposapi.themall.co.th'))
        self.mtr_target.pack(side=tk.LEFT, padx=5)
        
        self.mtr_start_btn = ttk.Button(ctrl_frame, text="Start MTR", command=self.start_mtr)
        self.mtr_start_btn.pack(side=tk.LEFT, padx=5)
        self.mtr_stop_btn = ttk.Button(ctrl_frame, text="Stop MTR", command=self.stop_mtr, state=tk.DISABLED)
        self.mtr_stop_btn.pack(side=tk.LEFT, padx=5)
        
        self.mtr_status_lbl = ttk.Label(ctrl_frame, text="Ready", foreground="gray")
        self.mtr_status_lbl.pack(side=tk.RIGHT, padx=10)

        # Graph Section (Canvas)
        graph_frame = ttk.LabelFrame(self.tab2, text="Destination Latency Graph (Live)", padding="5")
        graph_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.graph_canvas = tk.Canvas(graph_frame, height=120, bg="#1e1e1e", highlightthickness=0)
        self.graph_canvas.pack(fill=tk.X, padx=5, pady=5)
        self.draw_graph_grid()

        # Table Section (Treeview)
        table_frame = ttk.Frame(self.tab2, padding="5")
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ("Hop", "IP Address", "Loss%", "Sent", "Recv", "Best", "Avrg", "Worst", "Last")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=10)
        
        # Setup Column widths
        self.tree.column("Hop", width=40, anchor="center")
        self.tree.column("IP Address", width=180, anchor="w")
        self.tree.column("Loss%", width=60, anchor="center")
        self.tree.column("Sent", width=60, anchor="center")
        self.tree.column("Recv", width=60, anchor="center")
        self.tree.column("Best", width=60, anchor="center")
        self.tree.column("Avrg", width=60, anchor="center")
        self.tree.column("Worst", width=60, anchor="center")
        self.tree.column("Last", width=60, anchor="center")
        
        for col in columns:
            self.tree.heading(col, text=col)
            
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # ================= MTR & GRAPH LOGIC =================
    def draw_graph_grid(self):
        self.graph_canvas.delete("all")
        w = 660  # Approx width
        h = 120
        # Draw horizontal grid lines
        for y in range(0, h, 30):
            self.graph_canvas.create_line(0, y, w, y, fill="#333333", dash=(2, 2))
        self.graph_canvas.create_text(10, 10, text="ms", fill="#888888", anchor="w")

    def update_graph(self):
        self.draw_graph_grid()
        if not self.ping_history:
            return
            
        w = 660
        h = 120
        max_val = max(100, max(self.ping_history) * 1.2) # Minimum scale is 100ms
        
        x_step = w / (self.max_history - 1)
        points = []
        
        for i, val in enumerate(self.ping_history):
            x = i * x_step
            # Invert Y axis (0 is at bottom)
            y = h - ((val / max_val) * h)
            # Clip y to stay inside canvas
            y = max(0, min(h, y))
            points.append((x, y))
            
        if len(points) > 1:
            # Draw line
            for i in range(len(points) - 1):
                x1, y1 = points[i]
                x2, y2 = points[i+1]
                # Color code: Green if < 100, Yellow if < 200, Red if high
                color = "#00ff00"
                if self.ping_history[i+1] > 200: color = "#ff4444"
                elif self.ping_history[i+1] > 100: color = "#ffcc00"
                
                self.graph_canvas.create_line(x1, y1, x2, y2, fill=color, width=2)
            
            # Show latest value text
            last_val = self.ping_history[-1]
            self.graph_canvas.create_text(w - 30, points[-1][1] - 10, text=f"{last_val}ms", fill="#ffffff")

    def start_mtr(self):
        self.mtr_target_host = self.mtr_target.get().strip()
        if not self.mtr_target_host:
            return
            
        self.mtr_start_btn.config(state=tk.DISABLED)
        self.mtr_stop_btn.config(state=tk.NORMAL)
        self.mtr_target.config(state=tk.DISABLED)
        
        self.mtr_running = True
        self.mtr_data = {}
        self.ping_history = []
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        self.mtr_status_lbl.config(text="Tracing route (approx 10-15s)...", foreground="#d97706")
        threading.Thread(target=self.run_mtr_trace, daemon=True).start()

    def stop_mtr(self):
        self.mtr_running = False
        self.mtr_start_btn.config(state=tk.NORMAL)
        self.mtr_stop_btn.config(state=tk.DISABLED)
        self.mtr_target.config(state=tk.NORMAL)
        self.mtr_status_lbl.config(text="Stopped", foreground="gray")

    def run_mtr_trace(self):
        # Run Tracert (-d to skip DNS resolution for speed)
        process = subprocess.Popen(['tracert', '-d', '-h', '30', self.mtr_target_host], 
                                   stdout=subprocess.PIPE, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
        
        hop_regex = re.compile(r'^\s*(\d+)\s+.*\s+(\d+\.\d+\.\d+\.\d+)')
        
        while True:
            line = process.stdout.readline()
            if not line: break
            
            match = hop_regex.search(line)
            if match:
                hop_num = int(match.group(1))
                ip = match.group(2)
                self.mtr_data[hop_num] = {'ip': ip, 'sent': 0, 'recv': 0, 'best': 9999, 'worst': 0, 'sum': 0, 'last': 0}
                
                # Insert into Treeview safely
                self.root.after(0, lambda h=hop_num, i=ip: self.tree.insert("", "end", iid=str(h), values=(h, i, "0", "0", "0", "0", "0", "0", "0")))

        process.wait()
        
        if not self.mtr_running: return # Was stopped during trace
        
        self.root.after(0, lambda: self.mtr_status_lbl.config(text="Pinging continuous...", foreground="#16a34a"))
        threading.Thread(target=self.mtr_ping_loop, daemon=True).start()

    def mtr_ping_loop(self):
        while self.mtr_running:
            threads = []
            for hop_num, data in self.mtr_data.items():
                t = threading.Thread(target=self.ping_single_hop, args=(hop_num, data['ip']))
                t.start()
                threads.append(t)
                
            for t in threads:
                t.join() # Wait for all ping requests in this cycle to finish
                
            if not self.mtr_running: break
            
            self.root.after(0, self.update_mtr_ui)
            time.sleep(1) # Interval

    def ping_single_hop(self, hop_num, ip):
        # Ping with 1 count, 1000ms timeout
        process = subprocess.run(['ping', '-n', '1', '-w', '1000', ip], 
                                 capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
        
        output = process.stdout
        self.mtr_data[hop_num]['sent'] += 1
        
        ms_match = re.search(r'time[=<](\d+)ms', output)
        if ms_match:
            ms = int(ms_match.group(1))
            self.mtr_data[hop_num]['recv'] += 1
            self.mtr_data[hop_num]['last'] = ms
            self.mtr_data[hop_num]['sum'] += ms
            if ms < self.mtr_data[hop_num]['best']: self.mtr_data[hop_num]['best'] = ms
            if ms > self.mtr_data[hop_num]['worst']: self.mtr_data[hop_num]['worst'] = ms
            
            # If this is the final destination, add to graph history
            if hop_num == max(self.mtr_data.keys()):
                self.ping_history.append(ms)
                if len(self.ping_history) > self.max_history:
                    self.ping_history.pop(0)
        else:
            # Timeout
            self.mtr_data[hop_num]['last'] = "ERR"
            if hop_num == max(self.mtr_data.keys()):
                # Represent timeout as spike in graph
                self.ping_history.append(999)
                if len(self.ping_history) > self.max_history:
                    self.ping_history.pop(0)

    def update_mtr_ui(self):
        for hop_num, data in self.mtr_data.items():
            sent = data['sent']
            recv = data['recv']
            loss = int(((sent - recv) / sent) * 100) if sent > 0 else 0
            best = data['best'] if data['best'] != 9999 else 0
            avg = int(data['sum'] / recv) if recv > 0 else 0
            worst = data['worst']
            last = data['last']
            
            # Update Treeview row
            if self.tree.exists(str(hop_num)):
                self.tree.item(str(hop_num), values=(hop_num, data['ip'], f"{loss}%", sent, recv, best, avg, worst, last))
                
        self.update_graph()

    # ================= ORIGINAL SCHEDULED LOGIC =================
    def load_config(self):
        self.config = configparser.ConfigParser()
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(application_path, 'config.ini')
        if not os.path.exists(config_path): pass
        else: self.config.read(config_path)

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
        self.log(f"Running scheduled diagnostics for {self.host}...")
        computer_name = socket.gethostname()
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        file_name = f"{computer_name}_{timestamp}.txt"
        try: os.makedirs(self.log_folder, exist_ok=True)
        except OSError as e: return self.log(f"CRITICAL ERROR: Could not create log dir: {e}")
            
        full_path = os.path.join(self.log_folder, file_name)
        try:
            with open(full_path, 'w', encoding='utf-8', errors='ignore') as f:
                f.write(f"COMPREHENSIVE NETWORK DIAGNOSTIC REPORT\n=================================================\n")
                f.write(f"Report generated on: {datetime.now().strftime('%Y-%m-%d at %H:%M:%S')}\nTarget Host: {self.host}\n\n")

                f.write("\n===== 1. PING LOCAL DNS SERVERS =====\n\n--- Finding and Pinging configured DNS servers ---\n")
                f.flush()
                ps_get_dns_cmd = "(Get-DnsClientServerAddress -AddressFamily IPv4).ServerAddresses"
                dns_process = subprocess.run(['powershell', '-Command', ps_get_dns_cmd], capture_output=True, text=True, errors='ignore', creationflags=subprocess.CREATE_NO_WINDOW)
                if dns_process.stdout:
                    dns_servers = [s.strip() for s in dns_process.stdout.strip().split('\n') if s.strip()]
                    f.write(f"Found DNS Servers: {', '.join(dns_servers)}\n\n")
                    for server in dns_servers:
                        f.write(f"--- Pinging DNS Server: {server} (4 packets) ---\n")
                        f.flush()
                        ping_dns_process = subprocess.run(['ping', '-n', '4', server], capture_output=True, text=True, errors='ignore', creationflags=subprocess.CREATE_NO_WINDOW)
                        f.write(ping_dns_process.stdout + "\n")
                
                f.write("\n\n===== 2. TRACEROUTE TO VIEW NETWORK PATH =====\n\n")
                f.flush()
                process = subprocess.run(['tracert', self.host], capture_output=True, text=True, errors='ignore', creationflags=subprocess.CREATE_NO_WINDOW)
                f.write(process.stdout)

                f.write("\n\n===== 3. DNS LATENCY & RESOLUTION TEST =====\n\n--- Measuring time to resolve DNS name ---\n")
                f.flush()
                ps_command = f"Measure-Command {{Resolve-DnsName {self.host} -Type A -ErrorAction SilentlyContinue}}"
                process = subprocess.run(['powershell', '-ExecutionPolicy', 'Bypass', '-Command', ps_command], capture_output=True, text=True, errors='ignore', creationflags=subprocess.CREATE_NO_WINDOW)
                f.write(process.stdout)
                
                f.write("\n--- Pinging destination IP to measure latency (4 packets) ---\n")
                f.flush()
                process = subprocess.run(['ping', '-n', '4', self.host], capture_output=True, text=True, errors='ignore', creationflags=subprocess.CREATE_NO_WINDOW)
                f.write(process.stdout)

                f.write("\n\n===== 4. CURL API CONNECTION TIMING =====\n\n")
                f.flush()
                curl_format = "DNS Lookup:      %{time_namelookup}s\\nTCP Connection:  %{time_connect}s\\nSSL Handshake:   %{time_appconnect}s\\nTTFB:              %{time_starttransfer}s\\nTotal Time:      %{time_total}s\\n"
                process = subprocess.run(['curl', '-s', '-w', curl_format, '-o', 'nul', f"https://{self.host}"], capture_output=True, text=True, errors='ignore', creationflags=subprocess.CREATE_NO_WINDOW)
                f.write(process.stdout)

            self.log(f"Diagnostics complete. Log saved to {full_path}")
        except Exception as e:
            self.log(f"A critical error occurred while writing the log file: {e}")

    # --- System Tray Logic ---
    def create_icon_image(self):
        width, height = 64, 64
        image = Image.new('RGB', (width, height), (20, 20, 120))
        dc = ImageDraw.Draw(image)
        dc.rectangle((width // 2, 0, width, height // 2), fill=(80, 80, 220))
        dc.rectangle((0, height // 2, width // 2, height), fill=(80, 80, 220))
        return image
        
    def setup_tray_icon_thread(self):
        image = self.create_icon_image()
        menu = (item('Show', self.show_from_tray, default=True), item('Exit', self.exit_app))
        self.icon = Icon("API_Monitor", image, "API Connection Monitor", menu)
        self.icon.visible = False
        threading.Thread(target=self.icon.run, daemon=True).start()

    def hide_to_tray(self):
        self.root.withdraw()
        self.icon.visible = True

    def show_from_tray(self):
        self.icon.visible = False
        self.root.after(0, self.root.deiconify)

    def exit_app(self):
        self.icon.stop()
        self.stop_monitoring()
        self.stop_mtr()
        self.root.quit()

# --- Main Application Execution ---
if __name__ == '__main__':
    instance_name = "Global\\API_Monitor_UI_Mutex_v12" 
    instance = SingleInstance(instance_name)
    if instance.is_running():
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Application Already Running", "An instance of the Monitor is already running.")
        root.destroy()
        sys.exit(1)
        
    root = tk.Tk()
    app = App(root)
    root.withdraw()
    app.icon.visible = True
    root.after(2000, app.start_monitoring)
    root.mainloop()
