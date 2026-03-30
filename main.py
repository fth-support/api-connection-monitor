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
        self.root.title("API Connection Monitor & Live MTR")
        self.root.geometry("800x650") # ขยายหน้าจอให้พอดีกับ Header ใหม่
        self.root.resizable(False, False)
        
        self.style = ttk.Style(self.root)
        self.style.theme_use('clam')

        self.scheduler_thread = None
        self.stop_scheduler = threading.Event()
        self.icon = None
        
        # State Tracking
        self.scheduled_running = False
        self.mtr_running = False
        
        # MTR Variables
        self.mtr_data = {}
        self.time_history = []
        self.hop_history = {}
        self.max_history = 60
        
        self.documents_path = os.path.join(os.path.expanduser('~'), 'Documents')
        self.default_log_path = os.path.join(self.documents_path, 'API Latency Logs')
        self.load_config()

        # ================= GLOBAL HEADER (API ENDPOINT) =================
        header_frame = ttk.Frame(root, padding="10")
        header_frame.pack(fill=tk.X)
        
        ttk.Label(header_frame, text="🎯 Target API Endpoint:", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        self.host_entry = ttk.Entry(header_frame, width=50, font=("Arial", 10))
        self.host_entry.insert(0, self.config.get('Settings', 'endpoint', fallback='tmgposapi.themall.co.th'))
        self.host_entry.pack(side=tk.LEFT, padx=5)

        # ================= NOTEBOOK (TABS) =================
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        self.tab1 = ttk.Frame(self.notebook)
        self.tab2 = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab1, text="📅 Scheduled Monitor")
        self.notebook.add(self.tab2, text="📈 Live MTR Dashboard")

        self.setup_tab1()
        self.setup_tab2()

        self.root.protocol("WM_DELETE_WINDOW", self.hide_to_tray)
        self.setup_tray_icon_thread()

    def update_endpoint_state(self):
        """ ล็อกช่อง Endpoint ถ้ามีเครื่องมือใดเครื่องมือหนึ่งกำลังทำงานอยู่ """
        if self.scheduled_running or self.mtr_running:
            self.host_entry.config(state=tk.DISABLED)
        else:
            self.host_entry.config(state=tk.NORMAL)

    def setup_tab1(self):
        config_frame = ttk.LabelFrame(self.tab1, text="Schedule Configuration", padding="10")
        config_frame.pack(fill=tk.X, pady=5)
        config_frame.columnconfigure(1, weight=1)

        ttk.Label(config_frame, text="Schedule Times:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        time_frame = ttk.Frame(config_frame)
        time_frame.grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        
        times = [t.strip() for t in self.config.get('Settings', 'schedule_times', fallback='08:00,12:00,15:00,17:00,19:00').split(',')]
        self.time1_entry = ttk.Entry(time_frame, width=8); self.time1_entry.insert(0, times[0] if len(times) > 0 else ""); self.time1_entry.pack(side=tk.LEFT, padx=(0, 3))
        self.time2_entry = ttk.Entry(time_frame, width=8); self.time2_entry.insert(0, times[1] if len(times) > 1 else ""); self.time2_entry.pack(side=tk.LEFT, padx=3)
        self.time3_entry = ttk.Entry(time_frame, width=8); self.time3_entry.insert(0, times[2] if len(times) > 2 else ""); self.time3_entry.pack(side=tk.LEFT, padx=3)
        self.time4_entry = ttk.Entry(time_frame, width=8); self.time4_entry.insert(0, times[3] if len(times) > 3 else ""); self.time4_entry.pack(side=tk.LEFT, padx=3)
        self.time5_entry = ttk.Entry(time_frame, width=8); self.time5_entry.insert(0, times[4] if len(times) > 4 else ""); self.time5_entry.pack(side=tk.LEFT, padx=3)
        
        ttk.Label(config_frame, text="Log File Path:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.log_path_entry = ttk.Entry(config_frame)
        self.log_path_entry.insert(0, self.config.get('Settings', 'log_path', fallback=self.default_log_path))
        self.log_path_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        
        self.browse_button = ttk.Button(config_frame, text="Browse...", command=self.select_log_folder)
        self.browse_button.grid(row=1, column=2, padx=5, pady=5)

        control_frame = ttk.Frame(self.tab1, padding="10")
        control_frame.pack(fill=tk.X, pady=5)
        self.start_button = ttk.Button(control_frame, text="Start Scheduled Monitor", command=self.start_monitoring)
        self.start_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        self.stop_button = ttk.Button(control_frame, text="Stop Scheduled Monitor", command=self.stop_monitoring, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        
        log_frame = ttk.LabelFrame(self.tab1, text="Status Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10, relief="flat")
        self.log_area.pack(fill=tk.BOTH, expand=True)
        self.log("API Connection Monitor Initialized.")

    def setup_tab2(self):
        ctrl_frame = ttk.Frame(self.tab2, padding="5")
        ctrl_frame.pack(fill=tk.X)
        
        self.mtr_start_btn = ttk.Button(ctrl_frame, text="▶ Start Live MTR", command=self.start_mtr)
        self.mtr_start_btn.pack(side=tk.LEFT, padx=5)
        self.mtr_stop_btn = ttk.Button(ctrl_frame, text="⏹ Stop MTR", command=self.stop_mtr, state=tk.DISABLED)
        self.mtr_stop_btn.pack(side=tk.LEFT, padx=5)
        
        self.mtr_status_lbl = ttk.Label(ctrl_frame, text="Ready", foreground="gray")
        self.mtr_status_lbl.pack(side=tk.LEFT, padx=10)
        
        # ตัวเว้นระยะ
        ttk.Frame(ctrl_frame).pack(side=tk.LEFT, expand=True, fill=tk.X)
        
        # นาฬิกา Real-time สำหรับเป็นหลักฐานการ Capture
        self.clock_lbl = ttk.Label(ctrl_frame, text="--/--/---- --:--:--", font=("Arial", 10, "bold"), foreground="#0284c7")
        self.clock_lbl.pack(side=tk.RIGHT, padx=10)
        self.update_clock()

        # Graph Section
        graph_frame = ttk.LabelFrame(self.tab2, text="Multi-Hop Latency Graph & Timeline", padding="5")
        graph_frame.pack(fill=tk.X, padx=5, pady=5)
        self.graph_canvas = tk.Canvas(graph_frame, height=160, bg="#1e1e1e", highlightthickness=0)
        self.graph_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.draw_graph_bg()

        # Table Section (ระบุหน่วย ms)
        table_frame = ttk.Frame(self.tab2, padding="5")
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ("Hop", "IP Address", "Loss%", "Sent", "Recv", "Best (ms)", "Avrg (ms)", "Worst (ms)", "Last (ms)")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=9)
        
        self.tree.column("Hop", width=40, anchor="center")
        self.tree.column("IP Address", width=180, anchor="w")
        self.tree.column("Loss%", width=60, anchor="center")
        self.tree.column("Sent", width=60, anchor="center")
        self.tree.column("Recv", width=60, anchor="center")
        self.tree.column("Best (ms)", width=70, anchor="center")
        self.tree.column("Avrg (ms)", width=70, anchor="center")
        self.tree.column("Worst (ms)", width=70, anchor="center")
        self.tree.column("Last (ms)", width=70, anchor="center")
        
        for col in columns: self.tree.heading(col, text=col)
            
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def update_clock(self):
        """ อัปเดตนาฬิกาให้เดินตลอดเวลา """
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.clock_lbl.config(text=now)
        self.root.after(1000, self.update_clock)

    # ================= MTR & GRAPH LOGIC =================
    def draw_graph_bg(self):
        self.graph_canvas.delete("all")
        self.graph_canvas.create_text(10, 10, text="Waiting for data...", fill="#888888", anchor="w")

    def update_graph(self):
        self.graph_canvas.delete("all")
        w = int(self.graph_canvas.winfo_width())
        if w < 10: w = 700
        h = int(self.graph_canvas.winfo_height())
        if h < 10: h = 160
        
        bottom_margin = 25
        right_margin = 40
        graph_w = w - right_margin
        graph_h = h - bottom_margin
        
        if not self.time_history: return
            
        max_val = 100
        for hop, hist in self.hop_history.items():
            if hist:
                valid_ms = [x for x in hist if x < 999]
                if valid_ms and max(valid_ms) > max_val: max_val = max(valid_ms)
        max_val = max_val * 1.2
        
        for i in range(5):
            y = graph_h - (i * (graph_h / 4))
            val = int((i / 4) * max_val)
            self.graph_canvas.create_line(0, y, graph_w, y, fill="#333333", dash=(2, 2))
            self.graph_canvas.create_text(5, y - 8, text=f"{val}ms", fill="#888888", anchor="w", font=("Arial", 8))
            
        x_step = graph_w / (self.max_history - 1) if self.max_history > 1 else graph_w
        for i, tm in enumerate(self.time_history):
            x = i * x_step
            if i % 10 == 0 or i == len(self.time_history) - 1:
                self.graph_canvas.create_line(x, 0, x, graph_h, fill="#444444", dash=(1, 4))
                self.graph_canvas.create_text(x, graph_h + 10, text=tm, fill="#aaaaaa", font=("Arial", 8))
        
        for hop_num, hist in self.hop_history.items():
            if len(hist) < 2: continue
            points = []
            for i, val in enumerate(hist):
                x = i * x_step
                plot_val = max_val if val >= 999 else val
                y = graph_h - ((plot_val / max_val) * graph_h)
                points.append((x, y))
                
            for i in range(len(points) - 1):
                x1, y1 = points[i]; x2, y2 = points[i+1]
                val2 = hist[i+1]
                if val2 >= 999: color = "#ff00ff"
                elif val2 > 200: color = "#ff4444"
                elif val2 > 100: color = "#ffcc00"
                else: color = "#00ff00"
                
                line_w = 2 if hop_num == max(self.hop_history.keys()) else 1
                self.graph_canvas.create_line(x1, y1, x2, y2, fill=color, width=line_w)
            
            last_y = points[-1][1]
            self.graph_canvas.create_text(graph_w + 5, last_y, text=f"H{hop_num}", fill="#ffffff", anchor="w", font=("Arial", 8, "bold"))

    def start_mtr(self):
        self.host = self.host_entry.get().strip()
        if not self.host: return
            
        self.mtr_running = True
        self.update_endpoint_state()
        self.mtr_start_btn.config(state=tk.DISABLED)
        self.mtr_stop_btn.config(state=tk.NORMAL)
        
        self.mtr_data = {}
        self.time_history = []
        self.hop_history = {}
        for item in self.tree.get_children(): self.tree.delete(item)
            
        self.mtr_status_lbl.config(text="Tracing route (approx 10-15s)...", foreground="#d97706")
        threading.Thread(target=self.run_mtr_trace, daemon=True).start()

    def stop_mtr(self):
        self.mtr_running = False
        self.update_endpoint_state()
        self.mtr_start_btn.config(state=tk.NORMAL)
        self.mtr_stop_btn.config(state=tk.DISABLED)
        self.mtr_status_lbl.config(text="Stopped", foreground="gray")

    def run_mtr_trace(self):
        process = subprocess.Popen(['tracert', '-d', '-h', '30', self.host], stdout=subprocess.PIPE, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
        hop_regex = re.compile(r'^\s*(\d+)\s+.*\s+(\d+\.\d+\.\d+\.\d+)')
        while True:
            line = process.stdout.readline()
            if not line: break
            match = hop_regex.search(line)
            if match:
                hop_num = int(match.group(1)); ip = match.group(2)
                self.mtr_data[hop_num] = {'ip': ip, 'sent': 0, 'recv': 0, 'best': 9999, 'worst': 0, 'sum': 0, 'last': 0}
                self.root.after(0, lambda h=hop_num, i=ip: self.tree.insert("", "end", iid=str(h), values=(h, i, "0", "0", "0", "0", "0", "0", "0")))
        process.wait()
        if not self.mtr_running: return 
        self.root.after(0, lambda: self.mtr_status_lbl.config(text="Monitoring Live...", foreground="#16a34a"))
        threading.Thread(target=self.mtr_ping_loop, daemon=True).start()

    def mtr_ping_loop(self):
        while self.mtr_running:
            cycle_results = {}
            threads = []
            for hop_num, data in self.mtr_data.items():
                t = threading.Thread(target=self.ping_single_hop_cycle, args=(hop_num, data['ip'], cycle_results))
                t.start(); threads.append(t)
            for t in threads: t.join() 
            if not self.mtr_running: break
            
            current_time = datetime.now().strftime("%H:%M:%S")
            self.time_history.append(current_time)
            if len(self.time_history) > self.max_history: self.time_history.pop(0)
                
            for hop_num in self.mtr_data.keys():
                if hop_num not in self.hop_history: self.hop_history[hop_num] = []
                ms = cycle_results.get(hop_num, 999)
                self.hop_history[hop_num].append(ms)
                if len(self.hop_history[hop_num]) > self.max_history: self.hop_history[hop_num].pop(0)
                    
            self.root.after(0, self.update_mtr_ui)
            time.sleep(1)

    def ping_single_hop_cycle(self, hop_num, ip, results_dict):
        process = subprocess.run(['ping', '-n', '1', '-w', '1000', ip], capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
        self.mtr_data[hop_num]['sent'] += 1
        ms_match = re.search(r'time[=<](\d+)ms', process.stdout)
        if ms_match:
            ms = int(ms_match.group(1))
            self.mtr_data[hop_num]['recv'] += 1
            self.mtr_data[hop_num]['last'] = ms
            self.mtr_data[hop_num]['sum'] += ms
            if ms < self.mtr_data[hop_num]['best']: self.mtr_data[hop_num]['best'] = ms
            if ms > self.mtr_data[hop_num]['worst']: self.mtr_data[hop_num]['worst'] = ms
            results_dict[hop_num] = ms
        else:
            self.mtr_data[hop_num]['last'] = "ERR"
            results_dict[hop_num] = 999

    def update_mtr_ui(self):
        for hop_num, data in self.mtr_data.items():
            sent = data['sent']; recv = data['recv']
            loss = int(((sent - recv) / sent) * 100) if sent > 0 else 0
            best = data['best'] if data['best'] != 9999 else 0
            avg = int(data['sum'] / recv) if recv > 0 else 0
            worst = data['worst']; last = data['last']
            if self.tree.exists(str(hop_num)):
                self.tree.item(str(hop_num), values=(hop_num, data['ip'], f"{loss}%", sent, recv, best, avg, worst, last))
        self.update_graph()

    # ================= ORIGINAL SCHEDULED LOGIC =================
    def load_config(self):
        self.config = configparser.ConfigParser()
        if getattr(sys, 'frozen', False): application_path = os.path.dirname(sys.executable)
        else: application_path = os.path.dirname(os.path.abspath(__file__))
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
        self.host = self.host_entry.get().strip()
        if not self.host: return
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
        if not schedule_times_str: return
        
        self.scheduled_running = True
        self.update_endpoint_state()
        for widget in [self.start_button, self.time1_entry, self.time2_entry, self.time3_entry, self.time4_entry, self.time5_entry, self.log_path_entry, self.browse_button]:
            widget.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        
        self.log(f"Monitoring started for {self.host}. Scheduled times: {', '.join(schedule_times_str)}")
        schedule.clear()
        for scheduled_time in schedule_times_str: schedule.every().day.at(scheduled_time).do(self.run_diagnostics_thread)
        self.stop_scheduler.clear()
        self.scheduler_thread = threading.Thread(target=self.run_scheduler, daemon=True)
        self.scheduler_thread.start()
        threading.Timer(2.0, self.run_diagnostics_thread).start()

    def stop_monitoring(self):
        self.scheduled_running = False
        self.stop_scheduler.set()
        schedule.clear()
        self.update_endpoint_state()
        for widget in [self.start_button, self.time1_entry, self.time2_entry, self.time3_entry, self.time4_entry, self.time5_entry, self.log_path_entry, self.browse_button]:
            widget.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.log("Monitoring stopped by user.")

    def run_scheduler(self):
        while not self.stop_scheduler.is_set():
            schedule.run_pending()
            time.sleep(1)

    def run_diagnostics_thread(self): threading.Thread(target=self.run_diagnostics, daemon=True).start()

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

if __name__ == '__main__':
    instance_name = "Global\\API_Monitor_UI_Mutex_v14" 
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
