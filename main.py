import json
import math
import sys
import time
import tkinter as tk
from pathlib import Path
from tkinter import font as tkfont
from tkinter import ttk

from input_actions import NativeInput
from session import RunConfig, Session


APP_DIR = (Path(sys.executable) if getattr(sys, "frozen", False) else Path(__file__)).parent
SETTINGS_PATH = APP_DIR / "auto_clicker_settings.json"
LEGACY_PROFILE_PATH = APP_DIR / "other_app_profiles.json"
VK_F6 = 0x75
VK_ESCAPE = 0x1B


class AutoClickerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Auto Clicker")
        self.native = NativeInput()
        self.session = Session(self.native)
        self.closed = False
        self.was_shortcut_pressed = self.native.is_pressed(VK_F6)
        self.was_escape_pressed = self.native.is_pressed(VK_ESCAPE)

        interval, button = self._load_preferences()
        self.interval_var = tk.StringVar(value=interval)
        self.button_var = tk.StringVar(value=button)
        self.status_var = tk.StringVar(value="Stopped")
        self.count_var = tk.StringVar(value="0 clicks")
        self._build_ui()

        # Let Tk measure the controls so Windows text scaling cannot clip buttons.
        self.root.update_idletasks()
        self.root.minsize(max(400, self.root.winfo_reqwidth()), self.root.winfo_reqheight())
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._poll()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Auto Clicker", font=("Segoe UI", 18, "bold")).pack(anchor="w")
        ttk.Label(
            frame, text="Set the interval, place your cursor, and press F6.", wraplength=340
        ).pack(anchor="w", pady=(6, 20))

        interval_row = ttk.Frame(frame)
        interval_row.pack(fill="x", pady=(0, 14))
        ttk.Label(interval_row, text="Click every").pack(side="left")
        self.interval_entry = ttk.Entry(interval_row, textvariable=self.interval_var, width=10)
        self.interval_entry.pack(side="left", padx=10)
        ttk.Label(interval_row, text="seconds").pack(side="left")

        button_row = ttk.Frame(frame)
        button_row.pack(fill="x", pady=(0, 20))
        ttk.Label(button_row, text="Mouse button").pack(side="left", padx=(0, 12))
        self.button_combo = ttk.Combobox(
            button_row, textvariable=self.button_var, values=("left", "right"),
            state="readonly", width=10,
        )
        self.button_combo.pack(side="left")

        ttk.Separator(frame).pack(fill="x")
        self.status_font = tkfont.Font(family="Segoe UI", size=11, weight="bold")
        status_box = ttk.Frame(frame, height=self.status_font.metrics("linespace") * 3)
        status_box.pack(fill="x", pady=(16, 4))
        status_box.pack_propagate(False)
        ttk.Label(
            status_box, textvariable=self.status_var, wraplength=340,
            font=self.status_font,
        ).pack(anchor="w")
        ttk.Label(frame, textvariable=self.count_var).pack(anchor="w", pady=(0, 16))

        controls = ttk.Frame(frame)
        controls.pack(fill="x")
        controls.columnconfigure((0, 1), weight=1)
        self.start_button = ttk.Button(controls, text="Start (3s)", command=self.start)
        self.start_button.grid(row=0, column=0, sticky="ew", padx=(0, 6), ipady=5)
        self.stop_button = ttk.Button(controls, text="Stop", command=self.stop)
        self.stop_button.grid(row=0, column=1, sticky="ew", padx=(6, 0), ipady=5)
        ttk.Label(frame, text="F6: Start / Stop     Esc: Stop").pack(anchor="w", pady=(16, 0))
        self._sync_controls()

    @staticmethod
    def _valid_interval(value: object) -> float | None:
        try:
            interval = float(value)
        except (TypeError, ValueError, OverflowError):
            return None
        return interval if math.isfinite(interval) and 0.1 <= interval <= 3600 else None

    def _load_preferences(self) -> tuple[str, str]:
        path = SETTINGS_PATH if SETTINGS_PATH.exists() else LEGACY_PROFILE_PATH
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if path == LEGACY_PROFILE_PATH:
                data = data["profiles"][data.get("last_profile", "Default")]
            interval = self._valid_interval(data.get("interval"))
            button = data.get("button", data.get("input_mode", "left"))
            return (
                f"{interval:g}" if interval is not None else "1",
                button if button in ("left", "right") else "left",
            )
        except (OSError, ValueError, TypeError, KeyError, AttributeError):
            return ("1", "left")

    def _save_preferences(self) -> None:
        interval = self._valid_interval(self.interval_var.get())
        button = self.button_var.get()
        if interval is None or button not in ("left", "right"):
            return
        try:
            SETTINGS_PATH.write_text(
                json.dumps({"interval": f"{interval:g}", "button": button}, indent=2),
                encoding="utf-8",
            )
        except OSError:
            pass

    def _sync_controls(self) -> None:
        active = self.session.active
        self.interval_entry.configure(state="disabled" if active else "normal")
        self.button_combo.configure(state="disabled" if active else "readonly")
        self.start_button.configure(state="disabled" if active else "normal")
        self.stop_button.configure(state="normal" if active else "disabled")

    def _poll(self) -> None:
        if self.closed:
            return
        pressed = self.native.is_pressed(VK_F6)
        escape = self.native.is_pressed(VK_ESCAPE)
        if escape and not self.was_escape_pressed:
            self.stop()
        elif pressed and not self.was_shortcut_pressed and not escape:
            self.toggle()
        self.was_shortcut_pressed = pressed
        self.was_escape_pressed = escape

        active = self.session.active
        now = time.monotonic()
        self.session.tick(now)
        if active:
            if self.session.active and self.session.started_at is None:
                self.status_var.set(f"Starting in {self.session.activate_at - now:.1f}s...")
            elif self.session.status.startswith("Active:"):
                self.status_var.set("Clicking...")
            else:
                self.status_var.set(self.session.status)
        self.count_var.set(f"{self.session.count} clicks")
        self.session.events.clear()
        self._sync_controls()
        self.poll_id = self.root.after(50, self._poll)

    def start(self, delay: float = 3.0) -> None:
        if self.session.active:
            return
        interval = self._valid_interval(self.interval_var.get())
        if interval is None:
            self.status_var.set("Enter an interval from 0.1 to 3600 seconds.")
            return
        if self.button_var.get() not in ("left", "right"):
            self.status_var.set("Choose left or right mouse button.")
            return
        self._save_preferences()
        self.session.start(
            RunConfig(interval=interval, input_mode=self.button_var.get(), delay=delay),
            time.monotonic(),
        )
        self.status_var.set(f"Starting in {delay:.0f}s..." if delay else "Clicking...")
        self.count_var.set("0 clicks")
        self._sync_controls()

    def stop(self) -> None:
        self.session.stop("Stopped", time.monotonic())
        self.status_var.set("Stopped")
        self._sync_controls()

    def toggle(self) -> None:
        if self.session.active:
            self.stop()
        else:
            self.start(delay=0)

    def _on_close(self) -> None:
        if self.closed:
            return
        self.stop()
        self._save_preferences()
        self.closed = True
        self.root.after_cancel(self.poll_id)
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    ttk.Style().theme_use("vista")
    AutoClickerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
