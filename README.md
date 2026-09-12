# Auto Clicker

A simple Windows auto clicker with one small window.

## Download

[Download AutoClicker.exe](https://github.com/prashant1998gupta/Auto-Clicker/raw/refs/heads/main/dist/AutoClicker.exe). Double-click it to run; Python is not required.

## Use

1. Set **Click every** in seconds. For example, `1` clicks once per second; `0.1` clicks about ten times per second.
2. Choose **left** or **right** mouse button.
3. Place the cursor over your target and press **F6** to start.
4. Press **F6** again or **Esc** to stop.

The **Start (3s)** button gives you three seconds to position the cursor before clicking. **Stop** and **Esc** also cancel the countdown. Clicking follows the current cursor and continues until stopped.

The status and click count appear on the same screen. Stop before changing settings. Intervals from 0.1 to 3600 seconds are supported.

## Settings

The interval and mouse button are remembered in `auto_clicker_settings.json` beside the executable when the folder is writable.

Existing `other_app_profiles.json` files are left untouched. On first launch, only the last profile's mouse button and interval are read. Old Unity, keyboard, preview, position, and session-limit settings do not apply.

## Run From Source

Python on Windows is required. No third-party runtime packages are needed.

```powershell
python main.py
```

## Rebuild

```powershell
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

The script installs the pinned build dependency in `.venv-build`, runs the tests, and creates `dist/AutoClicker.exe`. Commit and push the EXE together with source changes to keep a GitHub backup.

Run tests separately with `python -m unittest -v test_automation`. Windows input is mocked, so the tests do not click or type into other applications.

## Support

[Buy me a coffee](https://buymeacoffee.com/logic_builder)
