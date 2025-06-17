import tkinter as tk
import subprocess

from rode.devices.classic.nt_usb_5g import NTUSB5thGenAppDevice
from rode.devices.utils.device_detection import DeviceDetectionUtils
from rode.devices.utils.versions import Version

dashboard_py = r"C:\Users\ate\Documents\nt1_gui\mira_dashboard2.py"
try:
    subprocess.Popen(['python', dashboard_py], shell=True)
except Exception as e:
    print(f"Could not launch dashboard: {e}")

from tkinter import messagebox
import subprocess
import threading
import os
from pathlib import Path
from soundcheck_tcpip.soundcheck.installation import construct_installation
from soundcheck_tcpip.soundcheck.controller import SCControlTCPIP
from PIL import Image, ImageTk
import sqlite3
import datetime
import json

# --- LAUNCH DASHBOARD BATCH FILE ---
dashboard_bat = os.path.join(os.path.dirname(__file__), "start_dashboard.bat")
try:
    # Use start to open in a new window, hide terminal popups
    subprocess.Popen(['start', '', dashboard_bat], shell=True)
except Exception as e:
    print(f"Could not launch dashboard: {e}")

SCRIPT_DIR = Path(__file__).parent
SEQUENCE_PATH = SCRIPT_DIR / "Mira - NT1 Gen 5 Production Sequence.sqc"
HARDWARE_CONFIG_PATH = Path(r"C:\SoundCheck 22\Steps\Hardware\NT1 Hardware.har")
LOGO_PATH = Path(r"C:\Users\ate\Documents\nt1_gui\rode_logo.png")
DEVICE_CLEANUP_EXE = Path(r"C:\Users\ate\Documents\DeviceCleanupCmd\x64\DeviceCleanupCmd.exe")
DB_PATH = SCRIPT_DIR / "mira_test_results.db"

DUT_RESULTS = {
    "DUT 1": {
        "usb_fr": "USB 1 FR Check",
        "usb_sens": "USB 1 Sensitivity Check",
        "amp": "Amp Test 1 check",
        "xlr_fr": "XLR 1 FR",
        "xlr_sens": "XLR 1 Sensitivity Check",
        "usb_noise": [f"USB 1 Band {i} Noise check" for i in range(1, 6)] + ["USB 1 Noise Check"],
        "xlr_noise": [f"XLR 1 Band {i} Noise check" for i in range(1, 6)] + ["XLR 1 Noise Check"],
        "usb_fr_curve": "USB 1 FR Curve",
        "xlr_fr_curve": "XLR 1 FR Curve",
        "xlr_smoothed_curve": "XLR DUT 1 Smoothed",
        "usb_smoothed_curve": "DUT 1 Smoothed (USB)",
        "xlr_noise_curve": "XLR 1 Noise A weighted",
        "usb_noise_curve": "USB 1 Noise A weighted",
    },
    "DUT 2": {
        "usb_fr": "USB 2 FR Check",
        "usb_sens": "USB 2 Sensitivity Check",
        "amp": "Amp Test 2 Check",
        "xlr_fr": "XLR 2 FR",
        "xlr_sens": "XLR 2 Sensitivity Check",
        "usb_noise": [f"USB 2 Band {i} Noise check" for i in range(1, 6)] + ["USB 2 Noise Check"],
        "xlr_noise": [f"XLR 2 Band {i} Noise check" for i in range(1, 6)] + ["XLR 2 Noise Check"],
        "usb_fr_curve": "USB 2 FR Curve",
        "xlr_fr_curve": "XLR 2 FR Curve",
        "xlr_smoothed_curve": "XLR DUT 2 Smoothed",
        "usb_smoothed_curve": "DUT 2 Smoothed (USB)",
        "xlr_noise_curve": "XLR 2 Noise A weighted",
        "usb_noise_curve": "USB 2 Noise A weighted",
    },
    "DUT 3": {
        "usb_fr": "USB 3 FR Check",
        "usb_sens": "USB 3 Sensitivity Check",
        "amp": "Amp Test 3 Check",
        "xlr_fr": "XLR 3 FR",
        "xlr_sens": "XLR 3 Sensitivity Check",
        "usb_noise": [f"USB 3 Band {i} Noise check" for i in range(1, 6)] + ["USB 3 Noise Check"],
        "xlr_noise": [f"XLR 3 Band {i} Noise check" for i in range(1, 6)] + ["XLR 3 Noise Check"],
        "usb_fr_curve": "USB 3 FR Curve",
        "xlr_fr_curve": "XLR 3 FR Curve",
        "xlr_smoothed_curve": "XLR DUT 3 Smoothed",
        "usb_smoothed_curve": "DUT 3 Smoothed (USB)",
        "xlr_noise_curve": "XLR 3 Noise A weighted",
        "usb_noise_curve": "USB 3 Noise A weighted",
    }
}

FAIL_MODE_LEGEND = {
    "FM1": "XLR FR Fail",
    "FM2": "XLR Sensitivity Fail",
    "FM3": "USB FR Fail",
    "FM4": "USB Sensitivity Fail",
    "FM5": "XLR FR + Sens Fail",
    "FM6": "USB FR + Sens Fail",
    "FM8": "Amp Test Fail",
    "FM9": "Noise Test Fail",
    "FM12": "USB All (FR + Sens + Amp) Fail",
    "FM13": "All USB + All XLR + Amp Fail",
    "FM16": "All XLR + Noise Fail"
}

def verify_firmware(desired_firmware: Version | str) -> bool:
    if isinstance(desired_firmware, str):
        desired_firmware = Version(desired_firmware)

    connected_devices = [
        device
        for device in DeviceDetectionUtils().get_connected_devices()
        if isinstance(device, NTUSB5thGenAppDevice)
    ]

    for device in connected_devices:
        if device.get_version() != desired_firmware:
            return False

    return True

def determine_fail_mode(usb_fr, usb_sens, amp, xlr_fr, xlr_sens, noise):
    if usb_fr and usb_sens and amp and xlr_fr and xlr_sens:
        return "FM13"
    if xlr_fr and xlr_sens and noise:
        return "FM16"
    if usb_fr and usb_sens and amp:
        return "FM12"
    if xlr_fr and xlr_sens:
        return "FM5"
    if usb_fr and usb_sens:
        return "FM6"
    if amp:
        return "FM8"
    if xlr_fr:
        return "FM1"
    if xlr_sens:
        return "FM2"
    if usb_fr:
        return "FM3"
    if usb_sens:
        return "FM4"
    if noise:
        return "FM9"
    return None

def is_failed_result(val):
    if isinstance(val, str):
        return val.strip().lower() in ["fail", "nan"]
    if isinstance(val, dict):
        return not val.get("Passed", True)
    return False

def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        c = conn.cursor()
        c.execute("""
            CREATE TABLE IF NOT EXISTS test_results (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT,
                employee_number TEXT,
                dut TEXT,
                status TEXT,
                fail_mode TEXT,
                failed_checks TEXT,
                memlist_json TEXT,
                xlr_fr_curve TEXT,
                usb_fr_curve TEXT,
                xlr_smoothed_curve TEXT,
                usb_smoothed_curve TEXT,
                xlr_noise_curve TEXT,
                usb_noise_curve TEXT
            )
        """)
        conn.commit()

def save_test_result(
    dut, status, fail_mode, failed, memlist_json, employee_number,
    xlr_fr_curve=None, usb_fr_curve=None,
    xlr_smoothed_curve=None, usb_smoothed_curve=None,
    xlr_noise_curve=None, usb_noise_curve=None
):
    with sqlite3.connect(DB_PATH) as conn:
        c = conn.cursor()
        c.execute("""
            INSERT INTO test_results (
                timestamp, employee_number, dut, status, fail_mode, failed_checks, memlist_json,
                xlr_fr_curve, usb_fr_curve,
                xlr_smoothed_curve, usb_smoothed_curve,
                xlr_noise_curve, usb_noise_curve
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            datetime.datetime.now().isoformat(), employee_number, dut, status, fail_mode, failed, memlist_json,
            xlr_fr_curve, usb_fr_curve,
            xlr_smoothed_curve, usb_smoothed_curve,
            xlr_noise_curve, usb_noise_curve
        ))
        conn.commit()

IDLE_TIMEOUT_MS = 10 * 60 * 1000  # 10 minutes

class OperatorPrompt(tk.Toplevel):
    def __init__(self, master, initial=False):
        super().__init__(master)
        self.result = None
        self.title("Operator Login" if initial else "Re-authentication")
        self.geometry("480x360+400+200")
        self.configure(bg="#181c25")
        self.resizable(False, False)
        self.grab_set()
        self.focus_force()
        self.lift()
        self.transient(master)
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)

        font_big = ("Segoe UI", 28, "bold")
        font_med = ("Segoe UI", 18)

        tk.Label(self, text="ENTER OPERATOR NUMBER", bg="#181c25", fg="#F7C325", font=font_big, pady=30).pack()
        self.entry_var = tk.StringVar()
        entry = tk.Entry(self, textvariable=self.entry_var, font=font_big, width=10, justify="center")
        entry.pack(pady=10)
        entry.focus_set()
        def on_validate(P): return P.isdigit() or P == ""
        vcmd = (self.register(on_validate), "%P")
        entry.config(validate="key", validatecommand=vcmd)
        btn_frame = tk.Frame(self, bg="#181c25")
        btn_frame.pack(pady=28)
        tk.Button(btn_frame, text="OK", font=font_med, width=10, height=2, bg="#22bb22", fg="white",
                  command=self.on_ok).pack(side=tk.LEFT, padx=20)
        tk.Button(btn_frame, text="Cancel", font=font_med, width=10, height=2, bg="#C50F1F", fg="white",
                  command=self.on_cancel).pack(side=tk.LEFT, padx=20)
        self.bind('<Return>', lambda event: self.on_ok())
        self.bind('<Escape>', lambda event: self.on_cancel())
    def on_ok(self):
        val = self.entry_var.get().strip()
        if val:
            self.result = val
            self.grab_release()
            self.destroy()
    def on_cancel(self):
        self.result = None
        self.grab_release()
        self.destroy()

class SoundCheckGUI:
    def __init__(self, master):
        init_db()
        self.master = master
        self.master.title("M.I.R.A – NT1 Gen 5")
        self.master.configure(bg="black")
        self.master.attributes("-topmost", True)
        self.master.state('zoomed')
        self.employee_number = None
        self.last_activity = datetime.datetime.now()

        self.pass_count = 0
        self.fail_count = 0

        self.main_layout = tk.Frame(master, bg="black")
        self.main_layout.pack(expand=True, fill="both")

        sidebar = tk.Frame(self.main_layout, bg="#181c25", width=220)
        sidebar.pack(side="left", fill="y", padx=(0, 12))
        sidebar.grid_propagate(False)

        tk.Label(sidebar, text="Test Stats", font=("Segoe UI", 20, "bold"), bg="#181c25", fg="#F7C325").pack(pady=(38, 18))
        self.pass_label = tk.Label(
            sidebar, text="PASS\n0", font=("Segoe UI", 32, "bold"),
            bg="#181c25", fg="#16C60C", width=7, height=2, relief="groove", bd=3
        )
        self.pass_label.pack(pady=(0, 18))
        self.fail_label = tk.Label(
            sidebar, text="FAIL\n0", font=("Segoe UI", 32, "bold"),
            bg="#181c25", fg="#C50F1F", width=7, height=2, relief="groove", bd=3
        )
        self.fail_label.pack(pady=(0, 16))

        tk.Label(sidebar, text="OPERATOR", font=("Segoe UI", 20, "bold"), bg="#181c25", fg="#F7C325").pack(pady=(16, 0))
        self.operator_number_label = tk.Label(
            sidebar, text="—", font=("Segoe UI", 38, "bold"), bg="#23294a", fg="#F7C325", width=7, height=1, pady=8, relief="sunken", bd=2
        )
        self.operator_number_label.pack(pady=(3, 20))
        tk.Button(sidebar, text="⎋ Logout", command=self.logout_employee,
                  bg="#C50F1F", fg="white", font=("Segoe UI", 18, "bold"), width=12, height=2).pack(pady=(20, 0))

        top_frame = tk.Frame(self.main_layout, bg="black")
        top_frame.pack(side="top", pady=(10, 5))

        tk.Label(top_frame, text="M.I.R.A – NT1 Gen 5", font=("Segoe UI", 28, "bold"),
                 bg="black", fg="#F7C325").pack(pady=(0, 10))

        try:
            logo = Image.open(LOGO_PATH)
            logo.thumbnail((500, 180), Image.LANCZOS)
            self.logo_img = ImageTk.PhotoImage(logo)
            tk.Label(top_frame, image=self.logo_img, bg="black").pack()
        except Exception:
            tk.Label(top_frame, text="RØDE", font=("Segoe UI", 50, "bold"), bg="black", fg="#F7C325").pack()

        mid_frame = tk.Frame(self.main_layout, bg="black")
        mid_frame.pack(expand=True)
        self.frames, self.labels, self.fm_labels = {}, {}, {}

        for i, dut in enumerate(DUT_RESULTS):
            frame = tk.Frame(mid_frame, bg="#666600", width=480, height=400, relief=tk.RIDGE, bd=5)
            frame.grid(row=0, column=i, padx=30)
            frame.grid_propagate(False)
            title = tk.Label(frame, text=dut, font=("Segoe UI", 18, "bold"), bg="#666600", fg="white")
            title.pack(pady=(10, 5))
            result = tk.Label(frame, text="", font=("Consolas", 11), bg="#666600", fg="white", justify="left")
            result.pack(pady=5)
            failmode = tk.Label(frame, text="", font=("Segoe UI", 24, "bold"), bg="#666600", fg="white")
            failmode.pack(pady=5)
            self.frames[dut] = (frame, title)
            self.labels[dut] = result
            self.fm_labels[dut] = failmode

        bot_frame = tk.Frame(self.main_layout, bg="black")
        bot_frame.pack(side="bottom", pady=30)
        button_font = ("Segoe UI", 20, "bold")
        self.run_button = tk.Button(bot_frame, text="▶ Run Sequence", command=self.run_sequence_trigger,
                  bg="green", fg="white", font=button_font, height=2, width=15)
        self.run_button.pack(side=tk.LEFT, padx=40, ipadx=30, ipady=20)
        tk.Button(bot_frame, text="🔁 Reset USB Devices", command=self.reset_usb,
                  bg="orange", fg="white", font=button_font, height=2, width=15).pack(
            side=tk.LEFT, padx=40, ipadx=30, ipady=20)
        tk.Button(bot_frame, text="❔ Fail Mode Legend", command=self.show_fail_mode_legend,
                  bg="#444", fg="white", font=button_font, height=2, width=15).pack(
            side=tk.LEFT, padx=40, ipadx=30, ipady=20)
        self.run_anim_label = tk.Label(self.main_layout, text="", font=("Segoe UI", 18),
                                       bg="black", fg="white")
        self.run_anim_label.pack(side="bottom", pady=(0, 20))
        self.anim_running = False
        self.anim_stage = 0
        self.schedule_auto_usb_cleanup()

        self.reset_idle_timer()
        self.master.bind_all("<Any-KeyPress>", self.reset_idle_timer)
        self.master.bind_all("<Any-Button>", self.reset_idle_timer)
        self.master.bind_all("<Motion>", self.reset_idle_timer)
        self.prompt_employee_number(initial=True)

    def prompt_employee_number(self, initial=False):
        self.employee_number = None
        self.operator_number_label.config(text="—")
        self.disable_test_buttons()
        self.master.attributes("-topmost", False)
        self.master.update()
        while not self.employee_number:
            prompt = OperatorPrompt(self.master, initial=initial)
            self.master.wait_window(prompt)
            emp = prompt.result
            if emp is None:
                if initial:
                    self.master.quit()
                    return
                else:
                    continue
            emp = emp.strip()
            if emp:
                self.employee_number = emp
                self.operator_number_label.config(text=emp)
                self.enable_test_buttons()
                self.last_activity = datetime.datetime.now()
                break
        self.master.attributes("-topmost", True)
        self.master.update()

    def reset_idle_timer(self, event=None):
        self.last_activity = datetime.datetime.now()
        if hasattr(self, "idle_timer_id"):
            self.master.after_cancel(self.idle_timer_id)
        self.idle_timer_id = self.master.after(IDLE_TIMEOUT_MS, self.check_idle)

    def check_idle(self):
        now = datetime.datetime.now()
        delta = (now - self.last_activity).total_seconds()
        if delta >= 600:
            self.logout_employee(idle=True)
        else:
            self.reset_idle_timer()

    def logout_employee(self, idle=False):
        self.employee_number = None
        self.operator_number_label.config(text="—")
        self.disable_test_buttons()
        if idle:
            messagebox.showinfo("Session Timeout", "No activity detected for 10 minutes.\nPlease re-enter your operator number.")
        self.prompt_employee_number(initial=False)

    def enable_test_buttons(self):
        self.run_button.config(state=tk.NORMAL)

    def disable_test_buttons(self):
        self.run_button.config(state=tk.DISABLED)

    def run_sequence_trigger(self):
        if not self.employee_number:
            self.prompt_employee_number(initial=False)
        else:
            self.run_sequence_threaded()

    def animate_running(self):
        if self.anim_running:
            dots = "." * (self.anim_stage % 4)
            self.run_anim_label.config(text=f"🔄 Running{dots}")
            self.anim_stage += 1
            self.master.after(500, self.animate_running)

    def schedule_auto_usb_cleanup(self):
        self.reset_usb(auto=True)
        self.master.after(3000, self.schedule_auto_usb_cleanup)

    def reset_usb(self, auto=False):
        try:
            result = subprocess.run(
                [str(DEVICE_CLEANUP_EXE), "-s", "*VID_19F7*"],
                capture_output=True, text=True
            )
            removed_count = result.stdout.count("removing device")
            if not auto:
                msg = f"✅ Removed {removed_count} non-present RØDE NT1 USB devices."
                if removed_count == 0:
                    msg = "No non-present NT1 USB devices found to remove."
                messagebox.showinfo("USB Cleanup", msg + f"\n\nDetails:\n{result.stdout}")
        except Exception as e:
            if not auto:
                messagebox.showerror("USB Cleanup Error", str(e))

    def run_sequence_threaded(self):
        self.anim_running = True
        self.anim_stage = 0
        self.animate_running()
        threading.Thread(target=self.run_sequence, daemon=True).start()

    def run_sequence(self):
        try:
            install = construct_installation([22, 0], Path(r"C:\SoundCheck 22"))
            install.import_har(str(HARDWARE_CONFIG_PATH))
            sc = SCControlTCPIP(install)
            sc.launch()
            sc.open_sequence(SEQUENCE_PATH)
            sc.run_sequence()

            result_names = sc.get_memlist_names("Results")
            curve_names = sc.get_memlist_names("Curves")

            dut_pass_this_run = 0
            dut_fail_this_run = 0

            for dut, tests in DUT_RESULTS.items():
                failed = []
                usb_fr = usb_sens = amp = xlr_fr = xlr_sens = noise = False
                for key, test in tests.items():
                    if key in ["usb_fr_curve", "xlr_fr_curve", "xlr_smoothed_curve", "usb_smoothed_curve", "xlr_noise_curve", "usb_noise_curve"]:
                        continue
                    if isinstance(test, list):
                        noise_failed = False
                        for item in test:
                            val = sc.get_result(item) if item in result_names else None
                            if val is None or is_failed_result(val):
                                noise_failed = True
                        if noise_failed:
                            if "usb_noise" in key:
                                failed.append("- USB Noise Check ❌")
                            elif "xlr_noise" in key:
                                failed.append("- XLR Noise Check ❌")
                            noise = True
                    else:
                        val = sc.get_result(test) if test in result_names else None
                        if val is None or is_failed_result(val):
                            failed.append(f"- {test} ❌")
                            if key == "usb_fr": usb_fr = True
                            if key == "usb_sens": usb_sens = True
                            if key == "amp": amp = True
                            if key == "xlr_fr": xlr_fr = True
                            if key == "xlr_sens": xlr_sens = True

                has_nan = any("nan" in f.lower() or "missing" in f.lower() for f in failed)
                fm = determine_fail_mode(usb_fr, usb_sens, amp, xlr_fr, xlr_sens, noise)
                frame, title = self.frames[dut]
                result_label = self.labels[dut]
                fm_label = self.fm_labels[dut]

                if has_nan or fm:
                    color = "#8B0000" if fm else "#800000"
                    title.config(text=f"{dut} – ❌ FAILED", bg=color)
                    result_label.config(text="\n".join(failed), bg=color)
                    fm_label.config(text=fm or "", bg=color)
                    frame.config(bg=color)
                    dut_fail_this_run += 1
                    status = "FAILED"
                else:
                    color = "#006400"
                    title.config(text=f"{dut} – ✅ PASSED", bg=color)
                    result_label.config(text="All checks passed.", bg=color)
                    fm_label.config(text="", bg=color)
                    frame.config(bg=color)
                    dut_pass_this_run += 1
                    status = "PASSED"

                memlist_dict = {}
                for name in result_names:
                    val = sc.get_result(name)
                    if hasattr(val, "tolist"):
                        val = val.tolist()
                    memlist_dict[name] = val

                def get_curve_json(curve_name):
                    if curve_name and curve_name in curve_names:
                        curve = sc.get_curve(curve_name)
                        if curve is not None:
                            x_data = curve.get("XData") if curve.get("XData") is not None else []
                            y_data = curve.get("YData") if curve.get("YData") is not None else []
                            return json.dumps({"X": list(x_data), "Y": list(y_data)})
                    return None

                xlr_fr_curve = get_curve_json(tests.get("xlr_fr_curve"))
                usb_fr_curve = get_curve_json(tests.get("usb_fr_curve"))
                xlr_smoothed_curve = get_curve_json(tests.get("xlr_smoothed_curve"))
                usb_smoothed_curve = get_curve_json(tests.get("usb_smoothed_curve"))
                xlr_noise_curve = get_curve_json(tests.get("xlr_noise_curve"))
                usb_noise_curve = get_curve_json(tests.get("usb_noise_curve"))

                save_test_result(
                    dut=dut,
                    status=status,
                    fail_mode=fm or "",
                    failed="; ".join(failed),
                    memlist_json=json.dumps(memlist_dict),
                    employee_number=self.employee_number,
                    xlr_fr_curve=xlr_fr_curve,
                    usb_fr_curve=usb_fr_curve,
                    xlr_smoothed_curve=xlr_smoothed_curve,
                    usb_smoothed_curve=usb_smoothed_curve,
                    xlr_noise_curve=xlr_noise_curve,
                    usb_noise_curve=usb_noise_curve
                )

            self.pass_count += dut_pass_this_run
            self.fail_count += dut_fail_this_run
            self.update_counters()

        except Exception as e:
            messagebox.showerror("Run Sequence Error", str(e))
        finally:
            self.anim_running = False
            self.run_anim_label.config(text="")

    def update_counters(self):
        self.pass_label.config(text=f"PASS\n{self.pass_count}")
        self.fail_label.config(text=f"FAIL\n{self.fail_count}")

    def show_fail_mode_legend(self):
        legend = "\n".join(f"{k}: {v}" for k, v in FAIL_MODE_LEGEND.items())
        messagebox.showinfo("Fail Mode Legend", legend)

if __name__ == "__main__":
    root = tk.Tk()
    app = SoundCheckGUI(root)
    root.mainloop()
