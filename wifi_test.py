import tkinter as tk
from tkinter import ttk
import pywifi
from pywifi import const
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import time
import threading
import socket
import subprocess
import math
import sys

try:
    import pythoncom
except ImportError:
    pythoncom = None

class WifiAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Wi-Fi Environment Analyzer (White Mode)")
        self.root.geometry("1100x850")
        self.is_scanning = False
        self.auto_scan_active = False
        self.target_band = "2.4GHz"
        self.current_results = []
        self.current_ssid = ""

        self.info_frame = ttk.Frame(root, padding=5, relief="groove")
        self.info_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        self.info_label = ttk.Label(self.info_frame, text="接続情報取得中...", font=("MS Gothic", 12, "bold"))
        self.info_label.pack()

        self.ctrl_frame = ttk.Frame(root, padding=5)
        self.ctrl_frame.pack(side=tk.TOP, fill=tk.X)

        self.scan_btn = ttk.Button(self.ctrl_frame, text="手動スキャン", command=self.start_manual_scan)
        self.scan_btn.pack(side=tk.LEFT, padx=5)

        self.auto_btn = ttk.Button(self.ctrl_frame, text="自動更新: OFF", command=self.toggle_auto_scan)
        self.auto_btn.pack(side=tk.LEFT, padx=5)

        ttk.Label(self.ctrl_frame, text=" |  周波数帯: ").pack(side=tk.LEFT, padx=5)
        self.band_var = tk.StringVar(value="2.4GHz")
        ttk.Radiobutton(self.ctrl_frame, text="2.4GHz", variable=self.band_var, value="2.4GHz", command=self.refresh_graph_only).pack(side=tk.LEFT)
        ttk.Radiobutton(self.ctrl_frame, text="5GHz", variable=self.band_var, value="5GHz", command=self.refresh_graph_only).pack(side=tk.LEFT)

        self.status_label = ttk.Label(self.ctrl_frame, text="準備完了", foreground="blue")
        self.status_label.pack(side=tk.RIGHT, padx=20)

        self.graph_frame = ttk.Frame(root)
        self.graph_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=5)
        plt.style.use("default")
        plt.rcParams['font.family'] = 'MS Gothic'
        self.fig, self.ax = plt.subplots(figsize=(10, 5))
        self.fig.subplots_adjust(bottom=0.15)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.graph_frame)
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.log_frame = ttk.LabelFrame(root, text="検出ログ", padding=5)
        self.log_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
        self.log_text = tk.Text(self.log_frame, height=5, state='disabled', font=("MS Gothic", 9))
        self.log_text.pack(side=tk.LEFT, fill=tk.X, expand=True)
        scrollbar = ttk.Scrollbar(self.log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text['yscrollcommand'] = scrollbar.set

        self.init_wifi()
        self.update_connection_info()

    def init_wifi(self):
        self.wifi = pywifi.PyWiFi()
        if len(self.wifi.interfaces()) == 0:
            self.log_message("エラー: Wi-Fiインターフェースが見つかりません。")
            self.scan_btn.config(state='disabled')
            self.iface = None
        else:
            self.iface = self.wifi.interfaces()[0]
            self.log_message(f"初期化完了: {self.iface.name()}")

    def get_current_connection_info(self):
        ssid = "未接続"
        ip = "取得不可"
        try:
            hostname = socket.gethostname()
            ip = socket.gethostbyname(hostname)
        except: pass
        try:
            output = subprocess.check_output("netsh wlan show interfaces", shell=True).decode('cp932', errors='ignore')
            for line in output.split('\n'):
                if "SSID" in line and "BSSID" not in line:
                    parts = line.split(':')
                    if len(parts) > 1:
                        ssid = parts[1].strip()
                        break
        except: pass
        return ssid, ip

    def update_connection_info(self):
        ssid, ip = self.get_current_connection_info()
        self.current_ssid = ssid
        self.info_label.config(text=f"現在の接続: SSID [{ssid}]  /  IP [{ip}]")

    def log_message(self, msg):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        print(f"[LOG] {msg}")

    def frequency_to_channel(self, freq_value):
        if freq_value is None: return None, None

        freq_mhz = freq_value
        if freq_value > 100_000_000:
            freq_mhz = freq_value / 1_000_000
        elif freq_value > 100_000:
            freq_mhz = freq_value / 1_000

        if 2412 <= freq_mhz <= 2484:
            ch = 14 if abs(freq_mhz - 2484) < 1 else int((freq_mhz - 2412) // 5 + 1)
            return ch, "2.4GHz"
        elif 5170 <= freq_mhz <= 5895:
            ch = int((freq_mhz - 5180) / 5) + 36
            return ch, "5GHz"
        return None, None

    def start_manual_scan(self):
        if not self.is_scanning:
            self.update_connection_info()
            threading.Thread(target=self.scan_process, daemon=True).start()

    def toggle_auto_scan(self):
        if self.auto_scan_active:
            self.auto_scan_active = False
            self.auto_btn.config(text="自動更新: OFF")
            self.status_label.config(text="自動更新停止")
        else:
            self.auto_scan_active = True
            self.auto_btn.config(text="自動更新: ON")
            self.start_manual_scan()

    def refresh_graph_only(self):
        if self.current_results:
            self.process_results(self.current_results)

    def scan_process(self):
        if pythoncom:
            pythoncom.CoInitialize()

        self.is_scanning = True
        self.root.after(0, lambda: self.status_label.config(text="スキャン中...", foreground="red"))
        self.root.after(0, lambda: self.scan_btn.config(state='disabled'))

        try:
            print("スキャン開始...")
            self.iface.scan()
            time.sleep(1)  # スキャン結果待ち時間を短縮（元は4秒）
            results = self.iface.scan_results()

            print(f"スキャン完了: 生データ {len(results)} 件")
            self.current_results = results
            self.process_results(results)

        except Exception as e:
            err_msg = f"スキャンエラー: {e}"
            print(err_msg)
            self.root.after(0, lambda: self.log_message(err_msg))

        finally:
            self.is_scanning = False
            self.root.after(0, lambda: self.scan_btn.config(state='normal'))
            self.root.after(0, lambda: self.status_label.config(text="完了", foreground="green"))

            if self.auto_scan_active:
                time.sleep(0.1)  # 次スキャンまでの待機を短縮（元は2秒）
                threading.Thread(target=self.scan_process, daemon=True).start()

    def process_results(self, results):
        target_band = self.band_var.get()
        data = []
        unique_check = set()

        for network in results:
            ssid = getattr(network, "ssid", "")
            if not ssid: continue

            freq_val = getattr(network, "freq", None) or getattr(network, "frequency", None)
            signal = getattr(network, "signal", -100)

            channel, band = self.frequency_to_channel(freq_val)

            if band != target_band or channel is None:
                continue

            key = f"{ssid}_{channel}"
            if key in unique_check: continue
            unique_check.add(key)

            data.append({"channel": channel, "signal": signal, "ssid": ssid, "band": band})

        print(f"{target_band}帯の有効データ: {len(data)} 件")
        df = pd.DataFrame(data)
        self.root.after(0, lambda: self.update_graph(df, target_band))

    def _channel_axis(self, band):
        if band == "2.4GHz":
            return list(np_linspace(1, 14, 200)), list(range(1, 15)), (1, 14)
        return list(np_linspace(34, 179, 500)), list(range(34, 180, 8)), (34, 179)

    def _curve(self, x_axis, center_ch, peak_dbm, band):
        base = -100
        spread = 2.5 if band == "2.4GHz" else 5.0
        y_vals = []
        for x in x_axis:
            delta = abs(x - center_ch)
            if delta > spread * 2:
                y = base
            else:
                y = base + (peak_dbm - base) * math.exp(-0.5 * (delta / (spread/2.5)) ** 2)
            y_vals.append(max(y, base))
        return y_vals

    def update_graph(self, df, band):
        # ★白背景設定
        plt.style.use("default")
        plt.rcParams['font.family'] = 'MS Gothic'

        self.ax.clear()

        # 軸データの生成
        x_min, x_max = (1, 14) if band == "2.4GHz" else (34, 179)
        step = (x_max - x_min) / 400
        x_axis = [x_min + i * step for i in range(401)]

        # グリッド設定（見やすいグレー）
        self.ax.grid(True, linestyle="--", alpha=0.5, color="#999999")
        self.ax.set_ylim(-100, -20)
        self.ax.set_xlim(x_min, x_max)

        if df.empty:
            self.ax.text((x_min+x_max)/2, -60, "No Wi-Fi Found", ha="center", color="#333", fontsize=14)
            self.canvas.draw()
            return

        connected_ssid = (self.current_ssid or "").strip()
        colors = plt.cm.tab10.colors

        if not df.empty:
            df = df.sort_values(by="signal", ascending=True)

        for idx, row in df.iterrows():
            is_connected = (row["ssid"] == connected_ssid)
            color = colors[idx % len(colors)]

            if is_connected:
                color = "#D32F2F"

            y_curve = self._curve(x_axis, row["channel"], row["signal"], band)

            z = 10 if is_connected else 5

            self.ax.fill_between(x_axis, y_curve, -100, color=color, alpha=0.3, zorder=z)
            self.ax.plot(x_axis, y_curve, color=color, linewidth=2.5 if is_connected else 1.5, zorder=z+1)

            font_weight = "bold" if is_connected else "normal"
            font_size = 11 if is_connected else 9

            self.ax.text(row["channel"], row["signal"] + 2, row["ssid"],
                         color="black",
                         fontsize=font_size,
                         fontweight=font_weight,
                         ha="center", va="bottom",
                         zorder=z+5,
                         rotation=0,
                         bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="none", alpha=0.7))

        if band == "2.4GHz":
            self.ax.set_xticks(range(1, 15))
        else:
            self.ax.set_xticks(range(36, 180, 8))

        self.ax.set_xlabel("チャンネル", fontsize=12)
        self.ax.set_ylabel("信号強度 (dBm)", fontsize=12)
        self.ax.set_title(f"Wi-Fi スペクトラム ({band})", fontsize=14, fontweight="bold")

        self.canvas.draw()
        msg = f"{band}帯: {len(df)} 件検出"
        self.log_message(msg)

def np_linspace(start, stop, num):
    step = (stop - start) / (num - 1)
    return [start + step * i for i in range(num)]

if __name__ == "__main__":
    root = tk.Tk()
    app = WifiAnalyzerApp(root)
    root.mainloop()