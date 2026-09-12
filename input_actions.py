import ctypes
from ctypes import wintypes
from pathlib import PureWindowsPath


user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

MODIFIERS = {"CTRL": 0x11, "SHIFT": 0x10, "ALT": 0x12}
KEYS = {
    "SPACE": 0x20, "ENTER": 0x0D, "TAB": 0x09,
    "UP": 0x26, "DOWN": 0x28, "LEFT": 0x25, "RIGHT": 0x27,
    "HOME": 0x24, "END": 0x23, "PAGEUP": 0x21, "PAGEDOWN": 0x22,
    "BACKSPACE": 0x08, "DELETE": 0x2E, "INSERT": 0x2D,
    **{f"F{i}": 0x6F + i for i in range(1, 13)},
    **{key: ord(key) for key in "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"},
}
EXTENDED_KEYS = {0x21, 0x22, 0x23, 0x24, 0x25, 0x26, 0x27, 0x28, 0x2D, 0x2E}
UNITY_PRESETS = {
    "Play / Stop": "CTRL+P",
    "Pause / Resume": "CTRL+SHIFT+P",
    "Step frame": "CTRL+ALT+P",
    "Frame selected": "F",
    "Move tool": "W",
    "Rotate tool": "E",
    "Scale tool": "R",
}


def parse_shortcut(value: str, reserved_key: int) -> tuple[int, ...]:
    parts = [part.strip().upper() for part in value.split("+")]
    if not parts or parts[-1] not in KEYS:
        raise ValueError("Choose a key such as SPACE, W, or CTRL+SHIFT+P. Esc is reserved for Stop.")
    if any(part not in MODIFIERS for part in parts[:-1]):
        raise ValueError("Use Ctrl, Shift or Alt before a single key, for example CTRL+P.")
    if len(set(parts)) != len(parts):
        raise ValueError("A shortcut cannot contain the same key twice.")
    keys = tuple(MODIFIERS[part] for part in parts[:-1]) + (KEYS[parts[-1]],)
    if reserved_key in keys:
        raise ValueError("The action key cannot also be the Start/Stop hotkey.")
    return keys


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG),
                ("mouseData", wintypes.DWORD), ("dwFlags", wintypes.DWORD),
                ("time", wintypes.DWORD), ("dwExtraInfo", ctypes.c_size_t)]


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD),
                ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD),
                ("dwExtraInfo", ctypes.c_size_t)]


class INPUTUNION(ctypes.Union):
    _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT)]


class INPUT(ctypes.Structure):
    _anonymous_ = ("data",)
    _fields_ = [("type", wintypes.DWORD), ("data", INPUTUNION)]


user32.SendInput.argtypes = [wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int]
user32.SendInput.restype = wintypes.UINT
user32.GetAsyncKeyState.argtypes = [ctypes.c_int]
user32.GetAsyncKeyState.restype = ctypes.c_short
user32.GetForegroundWindow.restype = wintypes.HWND
user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
user32.WindowFromPoint.argtypes = [wintypes.POINT]
user32.WindowFromPoint.restype = wintypes.HWND
kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
kernel32.OpenProcess.restype = wintypes.HANDLE
kernel32.QueryFullProcessImageNameW.argtypes = [
    wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)
]
kernel32.CloseHandle.argtypes = [wintypes.HANDLE]


class NativeInput:
    def is_pressed(self, key: int) -> bool:
        return bool(user32.GetAsyncKeyState(key) & 0x8000)

    def keys_held(self, keys: tuple[int, ...]) -> bool:
        return any(self.is_pressed(key) for key in (0x10, 0x11, 0x12, 0x5B, 0x5C, *keys))

    def foreground_executable(self) -> str:
        window = user32.GetForegroundWindow()
        if not window:
            return ""
        process_id = wintypes.DWORD()
        user32.GetWindowThreadProcessId(window, ctypes.byref(process_id))
        handle = kernel32.OpenProcess(0x1000, False, process_id.value)
        if not handle:
            return ""
        try:
            buffer = ctypes.create_unicode_buffer(32768)
            size = wintypes.DWORD(len(buffer))
            if not kernel32.QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(size)):
                return ""
            if user32.GetForegroundWindow() != window:
                return ""
            return PureWindowsPath(buffer.value).name.lower()
        finally:
            kernel32.CloseHandle(handle)

    def mouse_target_is_unity(self, point: tuple[int, int] | None) -> bool:
        position = wintypes.POINT(*point) if point is not None else wintypes.POINT()
        if point is None and not user32.GetCursorPos(ctypes.byref(position)):
            return False
        foreground = user32.GetForegroundWindow()
        target = user32.WindowFromPoint(position)
        if not foreground or not target:
            return False
        foreground_pid, target_pid = wintypes.DWORD(), wintypes.DWORD()
        user32.GetWindowThreadProcessId(foreground, ctypes.byref(foreground_pid))
        user32.GetWindowThreadProcessId(target, ctypes.byref(target_pid))
        return foreground_pid.value != 0 and foreground_pid.value == target_pid.value and self.foreground_executable() == "unity.exe"

    def perform(self, mode: str, keys: tuple[int, ...], point: tuple[int, int] | None) -> None:
        events = []
        if mode == "keyboard":
            for key in keys:
                flags = 1 if key in EXTENDED_KEYS else 0
                events.append(INPUT(type=1, ki=KEYBDINPUT(wVk=key, dwFlags=flags)))
            for key in reversed(keys):
                flags = 2 | (1 if key in EXTENDED_KEYS else 0)
                events.append(INPUT(type=1, ki=KEYBDINPUT(wVk=key, dwFlags=flags)))
        else:
            if point is not None and not user32.SetCursorPos(*point):
                raise OSError("Could not move the cursor to the saved point.")
            down, up = (0x08, 0x10) if mode == "right" else (0x02, 0x04)
            events = [INPUT(type=0, mi=MOUSEINPUT(dwFlags=flag)) for flag in (down, up)]

        batch = (INPUT * len(events))(*events)
        sent = user32.SendInput(len(batch), batch, ctypes.sizeof(INPUT))
        if sent != len(batch):
            # Release only keys/buttons whose down event may have reached Windows.
            if sent:
                if mode == "keyboard":
                    pressed = []
                    for event in events[:sent]:
                        if event.ki.dwFlags & 2:
                            pressed.remove(event.ki.wVk)
                        else:
                            pressed.append(event.ki.wVk)
                    releases = [INPUT(type=1, ki=KEYBDINPUT(wVk=key, dwFlags=2 | (1 if key in EXTENDED_KEYS else 0))) for key in reversed(pressed)]
                else:
                    releases = events[1:]
                release_batch = (INPUT * len(releases))(*releases)
                user32.SendInput(len(release_batch), release_batch, ctypes.sizeof(INPUT))
            raise OSError("Windows could not send the action. Check the target application's permissions.")
