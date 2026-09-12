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
        self.settings_path = Path(self.temp.name) / "settings.json"
        self.legacy_path = Path(self.temp.name) / "profiles.json"
        for name, value in (("SETTINGS_PATH", self.settings_path), ("LEGACY_PROFILE_PATH", self.legacy_path)):
            path_patch = patch.object(main, name, value)
            path_patch.start()
            self.addCleanup(path_patch.stop)
        self.native_patch = patch.object(main, "NativeInput")
        self.native = self.native_patch.start().return_value
        self.native.is_pressed.return_value = False
        self.native.keys_held.return_value = False
        self.addCleanup(self.native_patch.stop)
        self.root = tk.Tk()
        self.root.withdraw()
        self.app = main.AutoClickerApp(self.root)
        self.addCleanup(self.app._on_close)

    def poll_once(self):
        self.root.after_cancel(self.app.poll_id)
        self.app._poll()

    def test_repeats_selected_mouse_button_until_stopped(self):
        self.app.interval_var.set("0.5")
        self.app.button_var.set("right")
        self.app.start(delay=0)
        began = self.app.session.activate_at
        for offset in (0, 0.25, 0.5):
            self.app.session.tick(began + offset)
        self.assertEqual(self.native.perform.call_count, 2)
        self.native.perform.assert_called_with("right", (), None)
        self.app.stop()
        self.app.session.tick(began + 5)
        self.assertEqual(self.native.perform.call_count, 2)

    def test_f6_starts_immediately_and_does_not_repeat_while_held(self):
        self.native.is_pressed.side_effect = lambda key: key == main.VK_F6
        self.poll_once()
        self.assertTrue(self.app.session.active)
        self.assertEqual(self.app.session.config.delay, 0)
        self.poll_once()
        self.assertTrue(self.app.session.active)
        self.native.is_pressed.side_effect = lambda key: False
        self.poll_once()
        self.native.is_pressed.side_effect = lambda key: key == main.VK_F6
        self.poll_once()
        self.assertFalse(self.app.session.active)

    def test_escape_cancels_start_button_countdown(self):
        self.app.start()
        self.assertEqual(self.app.session.config.delay, 3)
        self.native.is_pressed.side_effect = lambda key: key == main.VK_ESCAPE
        self.poll_once()
        self.app.session.tick(time.monotonic() + 10)
        self.assertFalse(self.app.session.active)
        self.native.perform.assert_not_called()

    def test_invalid_intervals_do_not_start(self):
        for value in ("", "hello", "0", "-1", "nan", "inf", "0.01", "3601"):
            with self.subTest(value=value):
                self.app.interval_var.set(value)
                self.app.start()
                self.assertFalse(self.app.session.active)
                self.assertIn("interval", self.app.status_var.get())

    def test_settings_are_saved_and_loaded_without_a_profile_screen(self):
        self.app.interval_var.set("0.25")
        self.app.button_var.set("right")
        self.app._save_preferences()
        self.assertEqual(self.app._load_preferences(), ("0.25", "right"))
        self.assertEqual(json.loads(self.settings_path.read_text()), {"interval": "0.25", "button": "right"})

    def test_old_unity_profile_cannot_restrict_mouse_clicking(self):
        legacy = {"last_profile": "Unity", "profiles": {"Unity": {
            "interval": "2", "input_mode": "keyboard", "keyboard_key": "CTRL+P",
            "unity_only": True, "preview": True, "limit_mode": "By count",
            "limit_value": "1", "position_mode": "Saved point", "saved_x": "50",
        }}}
        self.legacy_path.write_text(json.dumps(legacy), encoding="utf-8")
        interval, button = self.app._load_preferences()
        self.assertEqual((interval, button), ("2", "left"))
        self.app.interval_var.set(interval)
        self.app.button_var.set(button)
        self.app.start(delay=0)
        config = self.app.session.config
        self.assertFalse(config.unity_only)
        self.assertFalse(config.preview)
        self.assertIsNone(config.max_actions)
        self.assertIsNone(config.saved_point)
        self.assertEqual(config.input_mode, "left")
        self.assertEqual(json.loads(self.legacy_path.read_text()), legacy)

    def test_settings_lock_while_clicking_and_unlock_on_stop(self):
        self.app.start()
        self.assertEqual(str(self.app.interval_entry.cget("state")), "disabled")
        self.assertEqual(str(self.app.start_button.cget("state")), "disabled")
        self.app.stop()
        self.assertEqual(str(self.app.interval_entry.cget("state")), "normal")
        self.assertEqual(str(self.app.button_combo.cget("state")), "readonly")


if __name__ == "__main__":
    unittest.main()
