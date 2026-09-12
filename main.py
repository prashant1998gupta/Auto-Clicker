import ctypes
import json
import math
import sys
import time
import tkinter as tk
from tkinter import ttk
from pathlib import Path

from input_actions import NativeInput, UNITY_PRESETS, parse_shortcut, user32
from session import RunConfig, Session


SHORTCUT_MAP = {
    "F6": 0x75,
    "F7": 0x76,
    "F8": 0x77,
    "F9": 0x78,
}

PROFILE_PATH = (
    Path(sys.executable) if getattr(sys, "frozen", False) else Path(__file__)
).with_name("other_app_profiles.json")


class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


class AutoClickerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("AutoClicker - Unity Tools")
        self.root.geometry("650x850")
        self.root.minsize(540, 640)
        self.root.resizable(True, True)

        self.native = NativeInput()
        self.session = Session(self.native)
        self.closed = False
        self.was_shortcut_pressed = False
        self.was_escape_pressed = False
        self.capture_deadline = None

        self.interval_var = tk.StringVar(value="1")
        self.input_var = tk.StringVar(value="left")
        self.key_var = tk.StringVar(value="SPACE")
        self.unity_only_var = tk.BooleanVar(value=False)
        self.preview_var = tk.BooleanVar(value=False)
        self.preset_var = tk.StringVar(value="Play / Stop")
        self.shortcut_var = tk.StringVar(value="F6")
        self.delay_var = tk.StringVar(value="0")
        self.position_mode_var = tk.StringVar(value="Current cursor")
        self.profile_var = tk.StringVar(value="Default")
        self.limit_mode_var = tk.StringVar(value="Unlimited")
        self.limit_value_var = tk.StringVar(value="10")
        self.saved_x_var = tk.StringVar(value="0")
        self.saved_y_var = tk.StringVar(value="0")
        self.status_var = tk.StringVar(value="Idle")
        self.shortcut_hint_var = tk.StringVar()
        self.activity_var = tk.StringVar(value="0 actions")
        self.elapsed_var = tk.StringVar(value="0.0s elapsed")
        self.location_var = tk.StringVar(value="Current cursor")
        self.limit_hint_var = tk.StringVar(value="No session limit")

        self.shortcut_vk = SHORTCUT_MAP[self.shortcut_var.get()]
        self.profile_store = self._load_profile_store()

        self.shortcut_var.trace_add("write", self._on_shortcut_change)
        self.limit_mode_var.trace_add("write", self._on_limit_change)
        self.limit_value_var.trace_add("write", self._on_limit_change)

        self._build_ui()
        self._sync_profile_choices()
        self._load_selected_profile()
        self._refresh_shortcut_hint()
        self._refresh_position_summary()
        self._refresh_limit_hint()
        self.input_var.trace_add("write", self._sync_input_fields)
        for variable in (
            self.interval_var, self.input_var, self.key_var, self.delay_var,
            self.position_mode_var, self.saved_x_var, self.saved_y_var,
            self.unity_only_var, self.preview_var,
        ):
            variable.trace_add("write", self._on_settings_change)
        self._sync_input_fields()
        self._poll()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        style = ttk.Style(self.root)
        style.configure("Title.TLabel", font=("Segoe UI", 20, "bold"))
        style.configure("Hint.TLabel", foreground="#526272")
        style.configure("TButton", padding=(10, 6))
        header = ttk.Frame(self.root, padding=(20, 16, 20, 10))
        header.pack(fill="x")
        ttk.Label(header, text="AutoClicker", style="Title.TLabel").pack(anchor="w")
        ttk.Label(header, text="Unity tools  /  Mouse and keyboard automation", style="Hint.TLabel").pack(anchor="w", pady=(4, 0))

        footer = ttk.Frame(self.root, padding=(20, 10, 20, 16))
        footer.pack(side="bottom", fill="x")
        ttk.Label(footer, textvariable=self.status_var, wraplength=490, font=("Segoe UI", 10, "bold")).pack(anchor="w")
        stats = ttk.Frame(footer)
        stats.pack(fill="x", pady=(6, 10))
        ttk.Label(stats, textvariable=self.activity_var).pack(side="left")
        ttk.Label(stats, textvariable=self.elapsed_var).pack(side="right")
        controls = ttk.Frame(footer)
        controls.pack(fill="x")
        controls.columnconfigure((0, 1), weight=1)
        ttk.Button(controls, text="Start", command=self.start).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(controls, text="Stop / Esc", command=self.stop).grid(row=0, column=1, sticky="ew", padx=(6, 0))
        ttk.Label(footer, textvariable=self.shortcut_hint_var, style="Hint.TLabel").pack(anchor="w", pady=(8, 0))

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=18)
        action_page = ttk.Frame(self.notebook)
        profile_page = ttk.Frame(self.notebook, padding=16)
        self.notebook.add(action_page, text="Actions & Unity")
        self.notebook.add(profile_page, text="Profiles & limits")

        container = ttk.Frame(action_page)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        canvas = tk.Canvas(container, highlightthickness=0)
        canvas.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=scrollbar.set)

        frame = ttk.Frame(canvas, padding=18)
        frame.columnconfigure(0, weight=1)

        frame_window = canvas.create_window((0, 0), window=frame, anchor="nw")

        def _sync_scroll_region(_event: tk.Event) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _sync_frame_width(event: tk.Event) -> None:
            canvas.itemconfigure(frame_window, width=event.width)

        def _on_mousewheel(event: tk.Event) -> None:
            delta = -1 * int(event.delta / 120) if event.delta else 0
            if delta:
                canvas.yview_scroll(delta, "units")

        frame.bind("<Configure>", _sync_scroll_region)
        canvas.bind("<Configure>", _sync_frame_width)
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        action_frame = frame
        frame = profile_page
        profile_box = ttk.LabelFrame(frame, text="Profiles", padding=12)
        profile_box.pack(fill="x", pady=(0, 10))
        profile_box.columnconfigure(0, weight=1)

        profile_row = ttk.Frame(profile_box)
        profile_row.pack(fill="x")
        profile_row.columnconfigure(0, weight=1)

        self.profile_combo = ttk.Combobox(
            profile_row,
            textvariable=self.profile_var,
            state="normal",
        )
        self.profile_combo.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(profile_row, text="Load", command=self.load_profile).grid(row=0, column=1)
        ttk.Button(profile_row, text="Save", command=self.save_profile).grid(
            row=0, column=2, padx=8
        )
        ttk.Button(profile_row, text="Delete", command=self.delete_profile).grid(row=0, column=3)

        frame = action_frame
        unity_box = ttk.LabelFrame(frame, text="Unity shortcut presets", padding=12)
        unity_box.pack(fill="x", pady=(0, 12))
        preset_row = ttk.Frame(unity_box)
        preset_row.pack(fill="x")
        ttk.Combobox(preset_row, textvariable=self.preset_var, values=tuple(UNITY_PRESETS), state="readonly").pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(preset_row, text="Use preset", command=self.use_unity_preset).pack(side="right")
        ttk.Label(unity_box, text="Loads one action with a 3-second delay. Check your bindings in Unity: Edit > Shortcuts.", wraplength=450, style="Hint.TLabel").pack(anchor="w", pady=(8, 0))

        ttk.Label(frame, text="Action interval (s)").pack(anchor="w")
        ttk.Entry(frame, textvariable=self.interval_var).pack(fill="x", pady=(6, 10))

        ttk.Label(frame, text="Mode").pack(anchor="w")
        ttk.Combobox(
            frame,
            textvariable=self.input_var,
            values=("left", "right", "keyboard"),
            state="readonly",
        ).pack(fill="x", pady=(6, 10))

        ttk.Label(frame, text="Key or combination").pack(anchor="w")
        self.key_combo = ttk.Combobox(frame, textvariable=self.key_var, values=("SPACE", "ENTER", "W", "A", "S", "D", "CTRL+P", "CTRL+SHIFT+P", "CTRL+ALT+P"))
        self.key_combo.pack(fill="x", pady=(6, 4))
        ttk.Label(frame, text="Examples: SPACE, W, CTRL+SHIFT+P. Esc always stops.", style="Hint.TLabel").pack(anchor="w", pady=(0, 10))
        ttk.Checkbutton(frame, text="Stop if Unity Editor is not focused", variable=self.unity_only_var).pack(anchor="w")
        ttk.Checkbutton(frame, text="Preview only (log actions without sending input)", variable=self.preview_var).pack(anchor="w", pady=(4, 12))
        ttk.Label(frame, text="Changing settings stops the current session. Start to apply them.", style="Hint.TLabel", wraplength=450).pack(anchor="w", pady=(0, 10))

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

        limit_box = ttk.LabelFrame(profile_page, text="Session", padding=12)
        limit_box.pack(fill="x", pady=(0, 10))

        limit_row = ttk.Frame(limit_box)
        limit_row.pack(fill="x")
        limit_row.columnconfigure(0, weight=1)
        limit_row.columnconfigure(1, weight=1)

        limit_mode = ttk.Frame(limit_row)
        limit_mode.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Label(limit_mode, text="Limit mode").pack(anchor="w")
        ttk.Combobox(
            limit_mode,
            textvariable=self.limit_mode_var,
            values=("Unlimited", "By count", "By duration"),
            state="readonly",
        ).pack(fill="x", pady=(6, 0))

        limit_value = ttk.Frame(limit_row)
        limit_value.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        ttk.Label(limit_value, text="Limit value").pack(anchor="w")
        ttk.Entry(limit_value, textvariable=self.limit_value_var).pack(fill="x", pady=(6, 0))

        ttk.Label(limit_box, textvariable=self.limit_hint_var).pack(anchor="w", pady=(10, 0))

        target_box = ttk.LabelFrame(frame, text="Mouse position (mouse modes only)", padding=12)
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
            target_actions, text="Capture point in 3s", command=self.capture_point
        ).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(target_actions, text="Reset point", command=self.reset_point).grid(
            row=0, column=1
        )
        ttk.Label(target_box, textvariable=self.location_var).pack(anchor="w", pady=(10, 0))

        log_box = ttk.LabelFrame(profile_page, text="Recent actions", padding=12)
        log_box.pack(fill="both", expand=True)
        self.action_log = tk.Text(log_box, height=9, width=40, state="disabled", wrap="word", font=("Consolas", 10))
        self.action_log.pack(fill="both", expand=True)
        ttk.Label(log_box, text="Shows the last 100 actions for this launch.", style="Hint.TLabel").pack(anchor="w", pady=(6, 0))

    def _poll(self) -> None:
        if self.closed:
            return
        now = time.monotonic()
        pressed = self.native.is_pressed(self.shortcut_vk)
        escape = self.native.is_pressed(0x1B)
        if escape and not self.was_escape_pressed:
            self.stop()
        elif pressed and not self.was_shortcut_pressed and not escape:
            self.toggle()
        self.was_shortcut_pressed = pressed
        self.was_escape_pressed = escape
        was_active = self.session.active
        self.session.tick(now)
        if was_active:
            self.status_var.set(self.session.status)
        for event in self.session.events:
            self.action_log.configure(state="normal")
            self.action_log.insert("end", f"{time.strftime('%H:%M:%S')}  {event}\n")
            if int(self.action_log.index("end-1c").split(".")[0]) > 101:
                self.action_log.delete("1.0", "2.0")
            self.action_log.see("end")
            self.action_log.configure(state="disabled")
        self.session.events.clear()
        if self.capture_deadline is not None and now >= self.capture_deadline:
            self.capture_deadline = None
            self._capture_point_now()
        self._refresh_stats()
        self.poll_id = self.root.after(50, self._poll)

    def _sync_input_fields(self, *_args: object) -> None:
        self.key_combo.configure(state="normal" if self.input_var.get() == "keyboard" else "disabled")

    def _on_settings_change(self, *_args: object) -> None:
        if self.session.active:
            self.stop()
            self.status_var.set("Settings changed. Press Start to run with the new settings.")

    def use_unity_preset(self) -> None:
        self.stop()
        self.input_var.set("keyboard")
        self.key_var.set(UNITY_PRESETS[self.preset_var.get()])
        self.unity_only_var.set(True)
        self.delay_var.set("3")
        self.limit_mode_var.set("By count")
        self.limit_value_var.set("1")
        self.position_mode_var.set("Current cursor")
        self._refresh_position_summary()
        self.status_var.set(f"Loaded {self.preset_var.get()}. Press Start, then switch to Unity.")

    def _get_interval(self) -> float:
        try:
            interval = float(self.interval_var.get())
        except ValueError:
            interval = 1.0

        return max(0.1, min(interval, 3600.0)) if math.isfinite(interval) else 1.0

    def _get_delay(self) -> float:
        try:
            delay = float(self.delay_var.get())
        except ValueError:
            delay = 0.0

        return max(0.0, min(delay, 60.0)) if math.isfinite(delay) else 0.0

    def _get_saved_point(self) -> tuple[int, int]:
        try:
            x = int(float(self.saved_x_var.get()))
        except (ValueError, OverflowError):
            x = 0

        try:
            y = int(float(self.saved_y_var.get()))
        except (ValueError, OverflowError):
            y = 0

        return (x, y)

    def _get_limit_config(self) -> tuple[str, int | float | None]:
        mode = self.limit_mode_var.get()
        if mode == "Unlimited":
            return (mode, None)

        try:
            raw_value = float(self.limit_value_var.get())
        except ValueError:
            raw_value = 10.0

        if not math.isfinite(raw_value):
            raw_value = 10.0

        if mode == "By count":
            return (mode, max(1, int(raw_value)))

        return (mode, max(1.0, min(raw_value, 86400.0)))

    def _refresh_shortcut_hint(self) -> None:
        self.shortcut_hint_var.set(f"{self.shortcut_var.get()}: Start/Stop   |   Esc: Stop (including countdown)")

    def _refresh_position_summary(self) -> None:
        if self.position_mode_var.get() == "Saved point":
            point = self._get_saved_point()
            self.location_var.set(f"Saved point: X={point[0]}, Y={point[1]}")
        else:
            self.location_var.set("Current cursor")

    def _refresh_limit_hint(self) -> None:
        mode, value = self._get_limit_config()
        if mode == "Unlimited":
            self.limit_hint_var.set("No session limit")
        elif mode == "By count":
            self.limit_hint_var.set(f"Stops automatically after {int(value)} actions")
        else:
            self.limit_hint_var.set(f"Stops automatically after {float(value):.1f} seconds")

    def _refresh_stats(self) -> None:
        noun = "previews" if self.session.config.preview else "actions"
        self.activity_var.set(f"{self.session.count} {noun}")
        self.elapsed_var.set(f"{self.session.elapsed:.1f}s elapsed")

    def _on_shortcut_change(self, *_args: object) -> None:
        if self.session.active:
            self.stop()
        self.shortcut_vk = SHORTCUT_MAP.get(self.shortcut_var.get(), SHORTCUT_MAP["F6"])
        self._refresh_shortcut_hint()

    def _on_limit_change(self, *_args: object) -> None:
        self._on_settings_change()
        self._refresh_limit_hint()

    def capture_point(self) -> None:
        self.stop()
        self.capture_deadline = time.monotonic() + 3
        self.status_var.set("Move the cursor to the target. Capturing in 3s; Esc cancels.")

    def _capture_point_now(self) -> None:
        point = POINT()
        if not user32.GetCursorPos(ctypes.byref(point)):
            self.status_var.set("Could not capture the cursor position")
            return
        self.saved_x_var.set(str(point.x))
        self.saved_y_var.set(str(point.y))
        self.position_mode_var.set("Saved point")
        self._refresh_position_summary()
        self.status_var.set(f"Point captured ({point.x}, {point.y})")

    def reset_point(self) -> None:
        self.capture_deadline = None
        self.saved_x_var.set("0")
        self.saved_y_var.set("0")
        self.position_mode_var.set("Current cursor")
        self._refresh_position_summary()
        self.status_var.set("Point reset")

    def _collect_current_profile(self) -> dict[str, object]:
        return {
            "interval": self.interval_var.get(),
            "input_mode": self.input_var.get(),
            "keyboard_key": self.key_var.get(),
            "unity_only": self.unity_only_var.get(),
            "preview": self.preview_var.get(),
            "shortcut": self.shortcut_var.get(),
            "delay": self.delay_var.get(),
            "position_mode": self.position_mode_var.get(),
            "saved_x": self.saved_x_var.get(),
            "saved_y": self.saved_y_var.get(),
            "limit_mode": self.limit_mode_var.get(),
            "limit_value": self.limit_value_var.get(),
        }

    def _apply_profile(self, data: dict[str, object]) -> None:
        self.stop()
        self.interval_var.set(str(data.get("interval", "1")))
        self.input_var.set(str(data.get("input_mode", "left")))
        self.key_var.set(str(data.get("keyboard_key", "SPACE")))
        self.unity_only_var.set(data.get("unity_only") is True)
        self.preview_var.set(data.get("preview") is True)
        self.shortcut_var.set(str(data.get("shortcut", "F6")))
        self.delay_var.set(str(data.get("delay", "0")))
        self.position_mode_var.set(str(data.get("position_mode", "Current cursor")))
        self.saved_x_var.set(str(data.get("saved_x", "0")))
        self.saved_y_var.set(str(data.get("saved_y", "0")))
        self.limit_mode_var.set(str(data.get("limit_mode", "Unlimited")))
        self.limit_value_var.set(str(data.get("limit_value", "10")))
        self._refresh_position_summary()
        self._refresh_limit_hint()

    def _load_profile_store(self) -> dict[str, object]:
        default_store = {
            "last_profile": "Default",
            "profiles": {
                "Default": {
                    "interval": "1",
                    "input_mode": "left",
                    "shortcut": "F6",
                    "delay": "0",
                    "position_mode": "Current cursor",
                    "saved_x": "0",
                    "saved_y": "0",
                    "limit_mode": "Unlimited",
                    "limit_value": "10",
                }
            },
        }

        if not PROFILE_PATH.exists():
            return default_store

        try:
            loaded = json.loads(PROFILE_PATH.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return default_store

        if not isinstance(loaded, dict) or not isinstance(loaded.get("profiles"), dict):
            return default_store

        profiles = {name: data for name, data in loaded["profiles"].items() if isinstance(data, dict)}
        if "Default" not in profiles:
            profiles["Default"] = default_store["profiles"]["Default"]

        return {
            "last_profile": str(loaded.get("last_profile", "Default")),
            "profiles": profiles,
        }

    def _save_profile_store(self) -> None:
        try:
            PROFILE_PATH.write_text(
                json.dumps(self.profile_store, indent=2),
                encoding="utf-8",
            )
        except OSError:
            self.status_var.set("Could not save profiles")

    def _sync_profile_choices(self) -> None:
        profiles = sorted(self.profile_store["profiles"].keys())
        self.profile_combo["values"] = profiles

    def _load_selected_profile(self) -> None:
        profiles = self.profile_store["profiles"]
        selected = self.profile_store.get("last_profile", "Default")
        if selected not in profiles:
            selected = "Default"
        self.profile_var.set(selected)
        self._apply_profile(profiles[selected])

    def load_profile(self) -> None:
        name = self.profile_var.get().strip()
        profile = self.profile_store["profiles"].get(name)
        if profile is None:
            self.status_var.set("Profile not found")
            return

        self.profile_store["last_profile"] = name
        self._apply_profile(profile)
        self._save_profile_store()
        self.status_var.set(f"Profile loaded: {name}")

    def save_profile(self) -> None:
        name = self.profile_var.get().strip() or "Default"
        self.profile_var.set(name)
        self.profile_store["profiles"][name] = self._collect_current_profile()
        self.profile_store["last_profile"] = name
        self._sync_profile_choices()
        self._save_profile_store()
        self.status_var.set(f"Profile saved: {name}")

    def delete_profile(self) -> None:
        name = self.profile_var.get().strip()
        if name in ("", "Default"):
            self.status_var.set("Default profile cannot be deleted")
            return

        if name not in self.profile_store["profiles"]:
            self.status_var.set("Profile not found")
            return

        del self.profile_store["profiles"][name]
        self.profile_store["last_profile"] = "Default"
        self._sync_profile_choices()
        self._load_selected_profile()
        self._save_profile_store()
        self.status_var.set(f"Profile deleted: {name}")

    def start(self) -> None:
        self.stop()
        interval = self._get_interval()
        limit_mode, limit_value = self._get_limit_config()
        input_mode = self.input_var.get()
        if input_mode not in ("left", "right", "keyboard"):
            self.status_var.set("Select left, right or keyboard mode")
            return
        try:
            keys = parse_shortcut(self.key_var.get(), self.shortcut_vk) if input_mode == "keyboard" else ()
        except ValueError as error:
            self.status_var.set(str(error))
            return
        position_mode = self.position_mode_var.get()
        saved_point = self._get_saved_point() if input_mode != "keyboard" and position_mode == "Saved point" else None
        config = RunConfig(
            interval=interval, input_mode=input_mode, keys=keys,
            key_label=self.key_var.get().strip().upper(), saved_point=saved_point,
            max_actions=int(limit_value) if limit_mode == "By count" else None,
            max_duration=float(limit_value) if limit_mode == "By duration" else None,
            delay=self._get_delay(), unity_only=self.unity_only_var.get(),
            preview=self.preview_var.get(),
        )
        self.session.start(config, time.monotonic())
        self.status_var.set(self.session.status)
        self._refresh_position_summary()
        self._refresh_limit_hint()

    def stop(self) -> None:
        self.capture_deadline = None
        self.session.stop("Idle", time.monotonic())
        self.status_var.set(self.session.status)

    def toggle(self) -> None:
        if self.session.active:
            self.stop()
            return

        self.start()

    def _on_close(self) -> None:
        self.stop()
        name = self.profile_var.get().strip() or "Default"
        self.profile_store["profiles"][name] = self._collect_current_profile()
        self.profile_store["last_profile"] = name
        self._save_profile_store()
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
