import ctypes
import json
import tempfile
import time
import tkinter as tk
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import input_actions
import main
from input_actions import NativeInput, parse_shortcut
from session import RunConfig, Session


class SessionTests(unittest.TestCase):
    def setUp(self):
        self.native = Mock()
        self.native.foreground_executable.return_value = "unity.exe"
        self.native.keys_held.return_value = False
        self.native.mouse_target_is_unity.return_value = True
        self.session = Session(self.native)

    def test_count_limit_is_exact_and_keeps_final_statistics(self):
        self.session.start(RunConfig(max_actions=2, interval=1), 0)
        for now in (0, 0.5, 1, 2):
            self.session.tick(now)
        self.assertEqual(self.native.perform.call_count, 2)
        self.assertEqual(self.session.count, 2)
        self.assertEqual(self.session.status, "Completed count limit")
        self.session.stop("Idle", 100)
        self.assertEqual(self.session.elapsed, 1)

    def test_duration_expires_between_actions_without_an_extra_click(self):
        self.session.start(RunConfig(interval=10, max_duration=2), 0)
        self.session.tick(0)
        self.session.tick(2)
        self.session.tick(10)
        self.assertEqual(self.native.perform.call_count, 1)
        self.assertFalse(self.session.active)
        self.assertEqual(self.session.status, "Completed duration limit")

    def test_cancelled_countdown_never_sends_input(self):
        self.session.start(RunConfig(delay=3), 0)
        self.session.tick(1)
        self.session.stop("Idle", 2)
        self.session.tick(5)
        self.native.perform.assert_not_called()

    def test_focus_loss_stops_without_automatic_resume(self):
        self.session.start(RunConfig(unity_only=True), 0)
        self.session.tick(0)
        self.native.foreground_executable.return_value = "notepad.exe"
        self.session.tick(0.1)
        self.native.foreground_executable.return_value = "unity.exe"
        self.session.tick(2)
        self.assertEqual(self.native.perform.call_count, 1)
        self.assertFalse(self.session.active)

    def test_unknown_focus_and_unity_hub_are_rejected(self):
        for executable in ("", "unity hub.exe"):
            with self.subTest(executable=executable):
                self.native.foreground_executable.return_value = executable
                self.session.start(RunConfig(unity_only=True), 0)
                self.session.tick(0)
                self.native.perform.assert_not_called()

    def test_click_outside_unity_is_rejected_even_when_unity_has_focus(self):
        self.native.mouse_target_is_unity.return_value = False
        self.session.start(RunConfig(unity_only=True, saved_point=(50, 50)), 0)
        self.session.tick(0)
        self.native.perform.assert_not_called()
        self.assertFalse(self.session.active)

    def test_preview_logs_without_touching_windows(self):
        self.session.start(RunConfig(preview=True, unity_only=True, max_actions=1), 0)
        self.session.tick(0)
        self.assertEqual(self.native.mock_calls, [])
        self.assertEqual(self.session.events, ["Preview #1: left click"])

    def test_held_keys_wait_without_counting_an_action(self):
        self.native.keys_held.return_value = True
        self.session.start(RunConfig(max_actions=1, input_mode="keyboard", keys=(0x11, 0x50)), 0)
        self.session.tick(0)
        self.assertEqual(self.session.count, 0)
        self.native.keys_held.return_value = False
        self.session.tick(1)
        self.assertEqual(self.session.count, 1)

    def test_send_failure_stops_and_does_not_increment_count(self):
        self.native.perform.side_effect = OSError("Blocked input")
        self.session.start(RunConfig(), 0)
        self.session.tick(0)
        self.assertFalse(self.session.active)
        self.assertEqual(self.session.status, "Blocked input")
        self.assertEqual(self.session.count, 0)


class KeyboardTests(unittest.TestCase):
    def test_chords_and_invalid_input(self):
        self.assertEqual(parse_shortcut(" ctrl + shift + p ", 0x75), (0x11, 0x10, 0x50))
        for invalid in ("", "ESC", "CTRL", "CTRL+F6", "CTRL+CTRL+P", "A+B", "CTRL++", "HELLO"):
            with self.subTest(invalid=invalid), self.assertRaises(ValueError):
                parse_shortcut(invalid, 0x75)

    def test_key_chord_orders_modifiers_and_releases_in_reverse(self):
        received = []
        def send(count, events, size):
            self.assertEqual(size, 40 if ctypes.sizeof(ctypes.c_void_p) == 8 else 28)
            received.extend((events[i].ki.wVk, events[i].ki.dwFlags) for i in range(count))
            return count
        with patch.object(input_actions.user32, "SendInput", side_effect=send):
            NativeInput().perform("keyboard", (0x11, 0x10, 0x50), None)
        self.assertEqual(received, [(0x11, 0), (0x10, 0), (0x50, 0), (0x50, 2), (0x10, 2), (0x11, 2)])

    def test_partial_send_releases_only_pressed_modifiers(self):
        calls = []
        def send(count, events, size):
            calls.append([(events[i].ki.wVk, events[i].ki.dwFlags) for i in range(count)])
            return 1 if len(calls) == 1 else count
        with patch.object(input_actions.user32, "SendInput", side_effect=send), self.assertRaises(OSError):
            NativeInput().perform("keyboard", (0x11, 0x50), None)
        self.assertEqual(calls[1], [(0x11, 2)])


class AppTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.profile_path = Path(self.temp.name) / "profiles.json"
        self.path_patch = patch.object(main, "PROFILE_PATH", self.profile_path)
        self.path_patch.start()
        self.addCleanup(self.path_patch.stop)
        self.native_patch = patch.object(main, "NativeInput")
        self.native = self.native_patch.start().return_value
        self.native.is_pressed.return_value = False
        self.addCleanup(self.native_patch.stop)
        self.root = tk.Tk()
        self.root.withdraw()
        self.app = main.AutoClickerApp(self.root)
        self.addCleanup(self.app._on_close)

    def test_preset_preview_and_profile_round_trip(self):
        self.app.use_unity_preset()
        self.assertTrue(self.app.unity_only_var.get())
        self.assertEqual(self.app.limit_value_var.get(), "1")
        self.assertEqual(self.app.delay_var.get(), "3")
        self.app.preview_var.set(True)
        self.app.delay_var.set("0")
        self.app.start()
        self.app.session.tick(time.monotonic())
        self.assertEqual(self.app.session.count, 1)
        self.native.perform.assert_not_called()
        self.app.profile_var.set("Unity test")
        self.app.save_profile()
        self.app.key_var.set("SPACE")
        self.app.load_profile()
        self.assertEqual(self.app.key_var.get(), "CTRL+P")
        stored = json.loads(self.profile_path.read_text())
        self.assertTrue(stored["profiles"]["Unity test"]["preview"])

    def test_toggle_cancels_pending_start_and_escape_cancels_capture(self):
        self.app.delay_var.set("3")
        self.app.start()
        self.app.toggle()
        self.assertFalse(self.app.session.active)
        self.app.capture_point()
        self.native.is_pressed.side_effect = lambda key: key == 0x1B
        self.root.after_cancel(self.app.poll_id)
        self.app._poll()
        self.assertIsNone(self.app.capture_deadline)

    def test_invalid_key_does_not_start(self):
        self.app.input_var.set("keyboard")
        self.app.key_var.set("F6")
        self.app.start()
        self.assertFalse(self.app.session.active)
        self.assertIn("Start/Stop", self.app.status_var.get())

    def test_legacy_profile_loads_with_defaults(self):
        self.app._apply_profile({"input_mode": "right", "interval": "2"})
        self.assertEqual(self.app.input_var.get(), "right")
        self.assertEqual(self.app.key_var.get(), "SPACE")
        self.assertFalse(self.app.unity_only_var.get())

    def test_changing_preview_stops_a_running_session(self):
        self.app.start()
        self.app.preview_var.set(True)
        self.assertFalse(self.app.session.active)
        self.assertIn("Settings changed", self.app.status_var.get())

    def test_layout_keeps_start_and_stop_in_footer(self):
        self.root.update_idletasks()
        buttons = []
        def visit(widget):
            if isinstance(widget, main.ttk.Button) and widget.cget("text") in ("Start", "Stop / Esc"):
                buttons.append(widget)
            for child in widget.winfo_children():
                visit(child)
        visit(self.root)
        self.assertEqual(len(buttons), 2)
        for button in buttons:
            self.assertEqual(button.master.master.master, self.root)
        self.assertEqual(len(self.app.notebook.tabs()), 2)


if __name__ == "__main__":
    unittest.main()
