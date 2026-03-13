import ctypes
import threading
import time
import tkinter as tk
from tkinter import ttk


user32 = ctypes.WinDLL("user32", use_last_error=True)

VK_F6 = 0x75
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP = 0x0010


class AutoClickerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Auto Clicker")
        self.root.geometry("380x340")
        self.root.minsize(380, 340)
        self.root.resizable(False, False)

        self.running = False
        self.shutdown = threading.Event()

        self.cps_var = tk.StringVar(value="10")
        self.button_var = tk.StringVar(value="left")
        self.status_var = tk.StringVar(value="Stopped")
        self.hotkey_var = tk.StringVar(value="Press F6 to toggle")

        self._build_ui()

        self.click_thread = threading.Thread(target=self._click_loop, daemon=True)
        self.hotkey_thread = threading.Thread(target=self._hotkey_loop, daemon=True)
        self.click_thread.start()
        self.hotkey_thread.start()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)

        ttk.Label(frame, text="Clicks per second").pack(anchor="w")
        ttk.Entry(frame, textvariable=self.cps_var).pack(fill="x", pady=(6, 14))

        ttk.Label(frame, text="Mouse button").pack(anchor="w")
        button_picker = ttk.Combobox(
            frame,
            textvariable=self.button_var,
            values=("left", "right"),
            state="readonly",
        )
        button_picker.pack(fill="x", pady=(6, 14))

        status_box = ttk.LabelFrame(frame, text="Status", padding=12)
        status_box.pack(fill="x", pady=(0, 14))
        ttk.Label(status_box, textvariable=self.status_var, font=("Segoe UI", 11, "bold")).pack(
            anchor="w"
        )
        ttk.Label(status_box, textvariable=self.hotkey_var).pack(anchor="w", pady=(4, 0))

        controls = ttk.Frame(frame)
        controls.pack(fill="x")
        controls.columnconfigure(0, weight=1)
        controls.columnconfigure(1, weight=1)

        ttk.Button(controls, text="Start", command=self.start).grid(
            row=0, column=0, sticky="ew", padx=(0, 6)
        )
        ttk.Button(controls, text="Stop", command=self.stop).grid(
            row=0, column=1, sticky="ew", padx=(6, 0)
        )

        ttk.Label(
            frame,
            text="The clicker uses the current mouse cursor position.\nUse F6 from anywhere to start or stop.",
            justify="left",
        ).pack(anchor="w", pady=(16, 0))

    def _click_loop(self) -> None:
        while not self.shutdown.is_set():
            if self.running:
                interval = self._get_interval()
                self._perform_click()
                time.sleep(interval)
            else:
                time.sleep(0.05)

    def _hotkey_loop(self) -> None:
        was_pressed = False
        while not self.shutdown.is_set():
            pressed = bool(user32.GetAsyncKeyState(VK_F6) & 0x8000)
            if pressed and not was_pressed:
                self.root.after(0, self.toggle)
            was_pressed = pressed
            time.sleep(0.05)

    def _get_interval(self) -> float:
        try:
            cps = float(self.cps_var.get())
        except ValueError:
            cps = 10.0

        cps = max(0.1, min(cps, 1000.0))
        return 1.0 / cps

    def _perform_click(self) -> None:
        if self.button_var.get() == "right":
            user32.mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
            user32.mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
            return

        user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

    def start(self) -> None:
        self.running = True
        self.status_var.set(f"Running ({self.button_var.get()} click, {self.cps_var.get()} CPS)")

    def stop(self) -> None:
        self.running = False
        self.status_var.set("Stopped")

    def toggle(self) -> None:
        if self.running:
            self.stop()
        else:
            self.start()

    def _on_close(self) -> None:
        self.shutdown.set()
        self.running = False
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    ttk.Style().theme_use("vista")
    AutoClickerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
