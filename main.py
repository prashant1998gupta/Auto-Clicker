import ctypes
import threading
import time
import tkinter as tk
from tkinter import ttk


user32 = ctypes.WinDLL("user32", use_last_error=True)

MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP = 0x0010

SHORTCUT_MAP = {
    "F6": 0x75,
    "F7": 0x76,
    "F8": 0x77,
    "F9": 0x78,
}


class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


class OtherApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Other App")
        self.root.geometry("420x620")
        self.root.minsize(420, 560)
        self.root.resizable(True, True)

        self.shutdown = threading.Event()
        self.active_event = threading.Event()
        self.state_lock = threading.Lock()

        self.interval_var = tk.StringVar(value="1")
        self.input_var = tk.StringVar(value="left")
        self.shortcut_var = tk.StringVar(value="F6")
        self.delay_var = tk.StringVar(value="0")
        self.position_mode_var = tk.StringVar(value="Current cursor")
        self.saved_x_var = tk.StringVar(value="0")
        self.saved_y_var = tk.StringVar(value="0")
        self.status_var = tk.StringVar(value="Idle")
        self.shortcut_hint_var = tk.StringVar()
        self.activity_var = tk.StringVar(value="0 actions")
        self.elapsed_var = tk.StringVar(value="0.0s elapsed")
        self.location_var = tk.StringVar(value="Current cursor")

        self.shortcut_vk = SHORTCUT_MAP[self.shortcut_var.get()]
        self.pending_start_token = 0
        self.session_count = 0
        self.session_started_at = None
        self.run_config = {
            "interval": 1.0,
            "input_mode": "left",
            "position_mode": "Current cursor",
            "saved_point": None,
        }

        self.shortcut_var.trace_add("write", self._on_shortcut_change)

        self._build_ui()
        self._refresh_shortcut_hint()
        self._refresh_position_summary()
        self._schedule_stats_refresh()

        self.action_thread = threading.Thread(target=self._action_loop, daemon=True)
        self.shortcut_thread = threading.Thread(target=self._shortcut_loop, daemon=True)
        self.action_thread.start()
        self.shortcut_thread.start()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)

        ttk.Label(frame, text="Action interval (s)").pack(anchor="w")
        ttk.Entry(frame, textvariable=self.interval_var).pack(fill="x", pady=(6, 10))

        ttk.Label(frame, text="Mode").pack(anchor="w")
        ttk.Combobox(
            frame,
            textvariable=self.input_var,
            values=("left", "right"),
            state="readonly",
        ).pack(fill="x", pady=(6, 10))

        shortcut_row = ttk.Frame(frame)
        shortcut_row.pack(fill="x", pady=(0, 10))
        shortcut_row.columnconfigure(0, weight=1)
        shortcut_row.columnconfigure(1, weight=1)

        shortcut_left = ttk.Frame(shortcut_row)
        shortcut_left.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Label(shortcut_left, text="Shortcut").pack(anchor="w")
        ttk.Combobox(
            shortcut_left,
            textvariable=self.shortcut_var,
            values=tuple(SHORTCUT_MAP.keys()),
            state="readonly",
        ).pack(fill="x", pady=(6, 0))

        delay_right = ttk.Frame(shortcut_row)
        delay_right.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        ttk.Label(delay_right, text="Start delay (s)").pack(anchor="w")
        ttk.Entry(delay_right, textvariable=self.delay_var).pack(fill="x", pady=(6, 0))

        target_box = ttk.LabelFrame(frame, text="Target", padding=12)
        target_box.pack(fill="x", pady=(0, 10))
        target_box.columnconfigure(0, weight=1)

        ttk.Label(target_box, text="Position mode").pack(anchor="w")
        ttk.Combobox(
            target_box,
            textvariable=self.position_mode_var,
            values=("Current cursor", "Saved point"),
            state="readonly",
        ).pack(fill="x", pady=(6, 10))

        coords = ttk.Frame(target_box)
        coords.pack(fill="x", pady=(0, 10))
        coords.columnconfigure(0, weight=1)
        coords.columnconfigure(1, weight=1)

        x_frame = ttk.Frame(coords)
        x_frame.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Label(x_frame, text="X").pack(anchor="w")
        ttk.Entry(x_frame, textvariable=self.saved_x_var).pack(fill="x", pady=(6, 0))

        y_frame = ttk.Frame(coords)
        y_frame.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        ttk.Label(y_frame, text="Y").pack(anchor="w")
        ttk.Entry(y_frame, textvariable=self.saved_y_var).pack(fill="x", pady=(6, 0))

        target_actions = ttk.Frame(target_box)
        target_actions.pack(anchor="w")

        ttk.Button(
            target_actions, text="Capture current point", command=self.capture_point
        ).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(target_actions, text="Reset point", command=self.reset_point).grid(
            row=0, column=1
        )
        ttk.Label(target_box, textvariable=self.location_var).pack(anchor="w", pady=(10, 0))

        status_box = ttk.LabelFrame(frame, text="Status", padding=12)
        status_box.pack(fill="x", pady=(0, 10))
        ttk.Label(status_box, textvariable=self.status_var, font=("Segoe UI", 11, "bold")).pack(
            anchor="w"
        )
        ttk.Label(status_box, textvariable=self.shortcut_hint_var).pack(anchor="w", pady=(4, 0))
        ttk.Label(status_box, textvariable=self.activity_var).pack(anchor="w", pady=(8, 0))
        ttk.Label(status_box, textvariable=self.elapsed_var).pack(anchor="w", pady=(2, 0))

        controls = ttk.Frame(frame)
        controls.pack(fill="x")
        controls.columnconfigure(0, weight=1)
        controls.columnconfigure(1, weight=1)

        ttk.Button(controls, text="Activate", command=self.start).grid(
            row=0, column=0, sticky="ew", padx=(0, 6)
        )
        ttk.Button(controls, text="Deactivate", command=self.stop).grid(
            row=0, column=1, sticky="ew", padx=(6, 0)
        )

    def _action_loop(self) -> None:
        while not self.shutdown.is_set():
            if not self.active_event.is_set():
                time.sleep(0.05)
                continue

            with self.state_lock:
                config = dict(self.run_config)

            if config["position_mode"] == "Saved point" and config["saved_point"] is not None:
                user32.SetCursorPos(*config["saved_point"])

            self._perform_action(config["input_mode"])

            with self.state_lock:
                self.session_count += 1

            time.sleep(config["interval"])

    def _shortcut_loop(self) -> None:
        was_pressed = False
        while not self.shutdown.is_set():
            shortcut_vk = self.shortcut_vk
            pressed = bool(user32.GetAsyncKeyState(shortcut_vk) & 0x8000)
            if pressed and not was_pressed:
                self.root.after(0, self.toggle)
            was_pressed = pressed
            time.sleep(0.05)

    def _perform_action(self, input_mode: str) -> None:
        if input_mode == "right":
            user32.mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
            user32.mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
            return

        user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

    def _get_interval(self) -> float:
        try:
            interval = float(self.interval_var.get())
        except ValueError:
            interval = 1.0

        return max(0.1, min(interval, 3600.0))

    def _get_delay(self) -> float:
        try:
            delay = float(self.delay_var.get())
        except ValueError:
            delay = 0.0

        return max(0.0, min(delay, 60.0))

    def _get_saved_point(self) -> tuple[int, int]:
        try:
            x = int(float(self.saved_x_var.get()))
        except ValueError:
            x = 0

        try:
            y = int(float(self.saved_y_var.get()))
        except ValueError:
            y = 0

        return (x, y)

    def _start_after_delay(self, token: int, delay: float) -> None:
        time.sleep(delay)
        if self.shutdown.is_set() or token != self.pending_start_token:
            return
        self.root.after(0, lambda: self._finish_start(token))

    def _finish_start(self, token: int) -> None:
        if token != self.pending_start_token or self.shutdown.is_set():
            return

        self.active_event.set()
        self.session_started_at = time.time()
        self.status_var.set(self._build_active_status())

    def _build_active_status(self) -> str:
        with self.state_lock:
            config = dict(self.run_config)

        target = "current cursor"
        if config["position_mode"] == "Saved point" and config["saved_point"] is not None:
            x, y = config["saved_point"]
            target = f"saved point {x}, {y}"

        return f"Active ({config['input_mode']} mode, every {config['interval']:.1f}s, {target})"

    def _refresh_shortcut_hint(self) -> None:
        self.shortcut_hint_var.set(f"Press {self.shortcut_var.get()} from anywhere to switch")

    def _refresh_position_summary(self) -> None:
        if self.position_mode_var.get() == "Saved point":
            point = self._get_saved_point()
            self.location_var.set(f"Saved point: X={point[0]}, Y={point[1]}")
        else:
            self.location_var.set("Current cursor")

    def _schedule_stats_refresh(self) -> None:
        self._refresh_stats()
        self.root.after(200, self._schedule_stats_refresh)

    def _refresh_stats(self) -> None:
        with self.state_lock:
            count = self.session_count

        self.activity_var.set(f"{count} actions")

        if self.session_started_at is None:
            self.elapsed_var.set("0.0s elapsed")
            return

        elapsed = max(0.0, time.time() - self.session_started_at)
        self.elapsed_var.set(f"{elapsed:.1f}s elapsed")

    def _on_shortcut_change(self, *_args: object) -> None:
        self.shortcut_vk = SHORTCUT_MAP.get(self.shortcut_var.get(), SHORTCUT_MAP["F6"])
        self._refresh_shortcut_hint()

    def capture_point(self) -> None:
        point = POINT()
        user32.GetCursorPos(ctypes.byref(point))
        self.saved_x_var.set(str(point.x))
        self.saved_y_var.set(str(point.y))
        self.position_mode_var.set("Saved point")
        self._refresh_position_summary()
        self.status_var.set(f"Point captured ({point.x}, {point.y})")

    def reset_point(self) -> None:
        self.saved_x_var.set("0")
        self.saved_y_var.set("0")
        self.position_mode_var.set("Current cursor")
        self._refresh_position_summary()
        self.status_var.set("Point reset")

    def start(self) -> None:
        self.pending_start_token += 1
        self.active_event.clear()

        interval = self._get_interval()
        input_mode = self.input_var.get()
        position_mode = self.position_mode_var.get()
        saved_point = self._get_saved_point() if position_mode == "Saved point" else None

        with self.state_lock:
            self.run_config = {
                "interval": interval,
                "input_mode": input_mode,
                "position_mode": position_mode,
                "saved_point": saved_point,
            }
            self.session_count = 0

        self.session_started_at = None
        self._refresh_position_summary()

        delay = self._get_delay()
        if delay > 0:
            self.status_var.set(f"Waiting {delay:.1f}s before activation")
            token = self.pending_start_token
            threading.Thread(
                target=self._start_after_delay,
                args=(token, delay),
                daemon=True,
            ).start()
            return

        self._finish_start(self.pending_start_token)

    def stop(self) -> None:
        self.pending_start_token += 1
        self.active_event.clear()
        self.session_started_at = None
        self.status_var.set("Idle")

    def toggle(self) -> None:
        if self.active_event.is_set():
            self.stop()
            return

        self.start()

    def _on_close(self) -> None:
        self.shutdown.set()
        self.active_event.clear()
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    ttk.Style().theme_use("vista")
    OtherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
