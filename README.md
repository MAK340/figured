# SumSelect

Highlight numbers anywhere in Windows, press a hotkey, and get the maths — sum, subtract,
multiply, divide, or a formula you build yourself, one operator at a time.

Built for accounting work: it understands `1,250.50`, accounting negatives like `(450)`,
Arabic-Indic digits (`١٢٣٫٥`), and written calculations like `1,250 × 3 − (400 / 2)`.

![tray icon](icon.png)

## What it does

* **Hotkey** (default `Ctrl+Alt+S`) — reads whatever is highlighted, in almost any app.
* **Formula builder** — every number becomes a tag with a clickable operator between each
  pair. Click to cycle `+ − × ÷`, or use the keyboard. The result updates live and follows
  normal order of operations (`×` and `÷` before `+` and `−`), so it matches Excel.
* **Drop a number** — click a tag to exclude it (a stray year, a "VAT 15%" rate, a line number).
* **Tally** — add results from different documents into one running total, kept across restarts.
* **Rounding** — full precision, 2 decimals, or 0 decimals, rounded half-up.
* **History** — the last 10 results in the tray menu; click one to copy it.
* **Excel export** — `Shift+Enter` copies the calculation as a formula, e.g. `=1250.5+2300*-450+99.75`.
* **Auto-on-copy** (optional) — any `Ctrl+C` containing two or more numbers pops the result up
  without stealing focus from what you are typing in.

## Controls

| Action | Mouse | Keyboard |
| --- | --- | --- |
| Change an operator | click it | `←` `→` to pick, `↑` `↓` to cycle, or type `+ - * /` |
| Drop / restore a number | click the tag | — |
| Copy the result | click the result | `Enter` |
| Copy as an Excel formula | — | `Shift+Enter` |
| Add to tally | "+ Add to tally" | `T` |
| Close | — | `Esc` |

`Set all: + − × ÷` applies one operator to every gap in one click.

## Why it is light

Measured on a 16-core laptop: **~16 MB** resident, **~0.05 % CPU** idle, **~140 ms of CPU**
per calculation.

* Win32 `RegisterHotKey` instead of a global keyboard hook, so it costs nothing while you type.
* Reads the selection through **UI Automation** first — no clipboard round-trip. It falls back
  to a synthetic `Ctrl+C` only when an app does not expose its selection, and restores whatever
  was on your clipboard afterwards.
* `AddClipboardFormatListener` for auto-on-copy: event driven, no polling.
* One popup window, reused and re-rendered rather than recreated.

## Install

Requires Windows and Python 3.12+.

```powershell
git clone https://github.com/MAK340/sumselect.git
cd sumselect
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install comtypes pystray pillow pyinstaller
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

`build.ps1` pre-generates the UI Automation COM wrapper, runs the self-test, builds a
one-folder PyInstaller bundle, installs it to `%LOCALAPPDATA%\SumSelect\app`, registers it to
start with Windows, and launches it.

To run from source without building: `.\.venv\Scripts\pythonw.exe sumselect.py`

## Settings

`%APPDATA%\SumSelect\config.json` — hotkey, popup seconds, rounding, auto-on-copy, tally and
history. Edit it, then use **Reload settings** in the tray menu. The file is read as UTF-8
with or without a byte-order mark, so editing it in Notepad is safe.

```json
{ "hotkey": "ctrl+alt+s", "auto_copy": false, "popup_seconds": 6, "decimals": null }
```

Hotkey syntax: modifiers `ctrl` `alt` `shift` `win` plus a letter, digit, `f1`–`f24`, or a
named key. If the combination is already taken by another app, the tray icon says so.

## Tests

```powershell
.\.venv\Scripts\python.exe sumselect.py --test
```

Runs the number parser, the operator seeding, the precedence evaluator and the formatters
over a set of awkward inputs (accounting negatives, Arabic-Indic digits, dates, division by
zero, thousands separators).

## The icon

Generated locally with ComfyUI (Qwen-Image) for the tile, with the Σ composited on top in
code — image models draw letters unreliably, and the glyph has to stay crisp at 16 px.
`comfy_icon.py` and `compose_icon.py` reproduce it; `build.ps1` bakes the result into the exe.

## Layout

| File | |
| --- | --- |
| `sumselect.py` | the whole app — parser, evaluator, Win32 plumbing, Tk popup, tray |
| `build.ps1` | build and install |
| `comfy_icon.py`, `compose_icon.py` | icon generation |
| `icon.png`, `icon.ico` | the generated icon |

## Licence

MIT
