<div align="center">

# 🖱️ Auto-Clicker
**A Simple, Powerful, and Feature-Rich Windows Desktop Utility built with Python & Tkinter.**

[![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)]()
[![Tkinter](https://img.shields.io/badge/GUI-Tkinter-E34F26?style=for-the-badge&logo=python&logoColor=white)]()
[![Windows](https://img.shields.io/badge/OS-Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)]()
[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20a%20Coffee-ffdd00?style=for-the-badge&logo=buy-me-a-coffee&logoColor=black)](https://buymeacoffee.com/logic_builder)

---

</div>

## 🚀 Overview

Auto-Clicker provides a seamless and easy way to automate mouse clicks on Windows. With features ranging from global hotkeys to session limits, it is perfect for any task requiring repetitive clicking. **No third-party packages required!**

## 💡 What is the use of this? (Use Cases)

This tool is designed to automate repetitive mouse clicks, saving you time and physical strain. Some common use cases include:
- **Gaming**: Automate repetitive clicking in idle, clicker, or incremental games.
- **Data Entry**: Speed up repetitive form submissions or UI navigation tasks.
- **Software Testing**: Rapidly trigger UI elements to stress-test your applications.
- **Keeping Sessions Active**: Prevent system idle timeouts or screen locking by simulating user activity.
- **Shopping & Ticketing**: Ensure you can click the "Buy" button instantly during high-demand sales or ticket drops.

## ✨ Features

- **Toggle On/Off**: Easily activate and deactivate right from the window or via keyboard shortcuts.
- **Global Shortcuts**: Support for custom hotkeys (`F6` to `F9`) from any window.
- **Saved Profiles**: Local persistence for all your custom click settings.
- **Adjustable Intervals**: Set specific click delays in seconds.
- **Start Delay**: Optional delayed start time before clicking begins.
- **Session Limits**: Stop automatically by limit count or total duration.
- **Targeting Modes**: Click at your current cursor or define a specific `X`/`Y` screen location.
- **Input Modes**: Left/right mouse clicks or keyboard combinations, including `CTRL+SHIFT+P`.
- **Unity Presets**: Play/Stop, Pause/Resume, Step frame, Frame selected, Move, Rotate, and Scale.
- **Unity Focus Check**: Stop when the foreground process is no longer `Unity.exe`; mouse targets must also belong to that Unity process.
- **Preview Mode**: Log the intended actions without sending mouse or keyboard input.
- **Emergency Stop**: `Esc` cancels both running sessions and countdowns.
- **Organized UI**: Separate action and profile tabs, with Start/Stop always outside the scrolling area.
- **Live Status Tracking**: View live activity actions and total elapsed time.

## Using Unity Tools

1. Open **Actions & Unity**, choose a Unity preset, and click **Use preset**.
2. Check the **Key or combination** field against your Unity bindings. Presets use common Windows defaults; custom bindings and contextual commands depend on your Unity setup. See Unity's [Windows shortcut reference](https://docs.unity3d.com/2017.4/Documentation/uploads/Main/Unity_HotKeys_Win.pdf) and [Shortcuts Manager](https://docs.unity3d.com/6000.0/Documentation/Manual/ShortcutsManager.html).
3. Enable **Preview only** for a dry run. View the result in **Profiles & limits > Recent actions**.
4. For real input, turn Preview off, click **Start**, then focus Unity during the three-second countdown. Focus the Scene view for tool commands. Presets run once; set a count or duration in **Profiles & limits** for a repeat test.
5. Press `Esc` to stop at any time. To reuse your configuration, name and save a profile.

The Unity check reads the foreground process; it does not bring Unity forward or inspect whether a particular panel, text field, or modal dialog has focus. Focus changes stop the session instead of automatically resuming it. Preview mode works without Unity running. Changing action settings stops the session; press Start to apply the new settings.

Keyboard combinations support Ctrl, Shift, Alt, A-Z, 0-9, F1-F12, Space, Enter, Tab, arrows, and navigation keys. Esc and the selected Start/Stop hotkey are reserved. Keys represent key presses, not case-sensitive text. The app waits for held modifiers/action keys to be released before sending input.

**Capture point in 3s** gives you time to move the cursor to a mouse target; Esc cancels capture. Saved mouse coordinates are ignored in keyboard mode.

The recent-action log stays in memory for this launch (last 100 actions). Profiles remain on disk beside the EXE. Existing profiles load with their original mouse settings.

## 🛠️ Run & Installation

### Download the Windows build

Download [AutoClicker.exe](../../raw/main/dist/AutoClicker.exe) and double-click it. Python is not required. The executable is backed up in this repository under `dist/AutoClicker.exe`.

Keep it in a writable folder: saved profiles are stored beside the executable in `other_app_profiles.json`.

### Rebuild on Windows

With Python installed, run `powershell -ExecutionPolicy Bypass -File .\build.ps1` from this folder. This installs the pinned build dependency in `.venv-build`, runs the automation tests, and creates `dist/AutoClicker.exe`.

Run tests separately with `python -m unittest -v test_automation`. Tests mock Windows input so they do not send clicks or keys to other applications.

After rebuilding, commit and push `dist/AutoClicker.exe` along with any source changes to back up the new version.

Ensure you have Python installed on your Windows machine. No external dependencies are needed (it uses standard libraries like `ctypes` and `tkinter`).

```powershell
# Run the app
python main.py
```

## 💖 Support The Project

If you find this tool helpful or it saves you time, please consider buying me a coffee! Your support keeps this project active and loved by devs.

<a href="https://buymeacoffee.com/logic_builder" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/v2/default-yellow.png" alt="Buy Me A Coffee" style="height: 60px !important;width: 217px !important;" ></a>

**Link:** [https://buymeacoffee.com/logic_builder](https://buymeacoffee.com/logic_builder)

<br>

<div align="center">
Made with ❤️ for Developers
</div>
