# Attendance Generator - Build & Distribute

Turns this into a standalone `.exe` (or Mac/Linux equivalent) that other
people can double-click to run -- no Python installed on their machines.

## Files in this folder

| File                     | What it is                                              |
|---------------------------|----------------------------------------------------------|
| `attendance_core.py`      | The generation logic (parsing, styling, formulas)        |
| `attendance_gui.py`       | The window (imports `attendance_core.py`)                |
| `build_windows.bat`       | Builds `AttendanceGenerator.exe` on Windows               |
| `build_mac_linux.sh`      | Builds the equivalent on macOS or Linux                   |

All four files (well, the two build scripts are OS-specific -- you only
need the one matching your build machine) must sit **in the same folder**.

## Important: build separately for each operating system

PyInstaller does not create a cross-platform file. A `.exe` built on
Windows only runs on Windows; a binary built on a Mac only runs on a Mac.
There is no single file that works everywhere.

**If everyone you're distributing to uses Windows** (the common case),
you only need to do the Windows build once, on one Windows PC.

## One-time setup (only on the machine doing the build)

You need Python installed **once**, only on the computer used to build the
`.exe`. The people who *receive* the finished `.exe` do not need Python at
all.

1. If Python isn't already installed: download it from
   https://www.python.org/downloads/ and during setup, check
   **"Add python.exe to PATH"**.

## Building on Windows

1. Put `attendance_core.py`, `attendance_gui.py`, and `build_windows.bat`
   in the same folder.
2. Double-click `build_windows.bat` (or right-click it -> Run, or open a
   command prompt in that folder and run it).
3. It installs `openpyxl` and `pyinstaller`, then builds. Takes under a
   minute.
4. Find `AttendanceGenerator.exe` inside the new `dist` folder.
5. Copy that single `.exe` file to give to anyone who needs it. That's the
   whole app -- no installer, no dependencies.

## Building on macOS / Linux

1. Put `attendance_core.py`, `attendance_gui.py`, and `build_mac_linux.sh`
   in the same folder.
2. Open Terminal in that folder and run:
   ```
   chmod +x build_mac_linux.sh
   ./build_mac_linux.sh
   ```
3. Find the built app inside the new `dist` folder.

## First-run security warnings (expected, safe to bypass)

Since the `.exe` isn't digitally code-signed (that costs money and isn't
needed for an internal tool), the OS will flag it the first time:

- **Windows**: "Windows protected your PC" (SmartScreen). Click
  **More info** -> **Run anyway**.
- **macOS**: "cannot be opened because the developer cannot be verified"
  (Gatekeeper). Right-click the app -> **Open** -> confirm **Open**. Only
  needs to be done once per machine.

This is normal for any app that isn't purchased from an app store or
signed with a paid developer certificate -- it doesn't mean anything is
wrong with the file.

## Changing holidays / standard hours each month

The first time anyone runs the `.exe`, it creates a file called
`attendance_settings.json` in the same folder as the `.exe`, e.g.:

```json
{
  "target_hours": 9,
  "weekly_off": "Friday",
  "holidays": {
    "2026-08-05": "Holiday",
    "2026-08-26": "Holiday"
  }
}
```

To update holidays for a new month, open that file in any text editor,
change the dates (format `YYYY-MM-DD`), and save. **No rebuild needed** --
just re-run the `.exe`. This file travels with the `.exe`, so if you copy
the `.exe` to someone else's machine, copy `attendance_settings.json`
alongside it if you want them to start with the same holiday list (or let
it auto-create with defaults on their first run and they can edit it
there).

## Updating the tool later

If you change `attendance_core.py` or `attendance_gui.py` (say, to fix a
bug or add a feature), you need to re-run the build script and
redistribute the new `.exe` -- the `.exe` is a frozen snapshot of the code
at build time. Settings (holidays, hours) are the only thing that updates
without a rebuild.
