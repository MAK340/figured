"""SumSelect v2 - highlight numbers anywhere on Windows, get Sum / - / x / / instantly.

Speed design:
  * Win32 RegisterHotKey (no global keyboard hook -> zero typing overhead)
  * Reads the selection through UI Automation first (no clipboard, no Ctrl+C wait);
    falls back to a synthetic Ctrl+C only when the app doesn't expose its selection
  * AddClipboardFormatListener for auto-on-copy (event driven, no polling)
  * One persistent popup window, shown without stealing focus in auto mode
Features: tally across selections, tick/untick individual numbers, rounding, history.
"""
import ctypes
import json
import os
import queue
import re
import sys
import threading
import time
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation, getcontext

getcontext().prec = 34
APP = "SumSelect"
CFG_DIR = os.path.join(os.environ.get("APPDATA", "."), APP)
CFG_PATH = os.path.join(CFG_DIR, "config.json")
DEFAULT_CFG = {"hotkey": "ctrl+alt+s", "auto_copy": False, "popup_seconds": 6,
               "decimals": None, "history": [], "tally": []}

# ======================================================================== parsing
_TRANS = {ord(a): str(i) for i, a in enumerate("٠١٢٣٤٥٦٧٨٩")}
_TRANS.update({ord(a): str(i) for i, a in enumerate("۰۱۲۳۴۵۶۷۸۹")})
_TRANS.update({ord("٫"): ".", ord("٬"): ",", ord("−"): "-", ord("–"): "-", ord("—"): "-",
               ord("×"): "*", ord("÷"): "/", ord(" "): " ", ord(" "): " "})

NUM_RE = re.compile(
    r"(?<![A-Za-z\d.])"
    r"(?P<open>\()?\s*"
    r"(?P<sign>(?<![\w)])[-+])?"
    r"(?P<num>\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?|\.\d+)"
    r"(?P<close>\s*\))?"
)


def normalize(text):
    return text.translate(_TRANS)


def extract_numbers(text):
    """List of (Decimal value, source text, start, end)."""
    out = []
    for m in NUM_RE.finditer(text):
        try:
            v = Decimal(m.group("num").replace(",", ""))
        except InvalidOperation:
            continue
        paren = bool(m.group("open") and m.group("close"))
        if m.group("sign") == "-" or paren:
            v = -v
        src = m.group("num")
        if paren:
            src = f"({src})"
        elif m.group("sign") == "-":
            src = "-" + src
        out.append((v, src, m.start(), m.end()))
    return out


OP_CHARS = {"*": "*", "/": "/", "+": "+", "-": "-"}


def seed_ops(text, nums, is_expr=False):
    """Operator for each gap between consecutive numbers, read from the text itself.

    A bare symbol glued between digits (2026-09-19) is not treated as an operator
    unless the whole selection reads as a calculation.
    """
    ops = []
    for a, b in zip(nums, nums[1:]):
        gap = text[a[3]:b[2]]
        bare = gap.strip()
        op = "+"
        if bare and (is_expr or gap != bare or len(bare) > 1):
            if bare.lower() == "x":
                op = "*"
            else:
                for ch in reversed(bare):
                    if ch in OP_CHARS:
                        op = OP_CHARS[ch]
                        break
        ops.append(op)
    return ops


def build(values, ops):
    """Evaluate values joined by ops using * / before + -. values/ops already filtered."""
    if not values:
        return None
    terms, cur, pending = [], values[0], None
    for op, v in zip(ops, values[1:]):
        if op == "*":
            cur = cur * v
        elif op == "/":
            if v == 0:
                return None
            cur = cur / v
        else:
            terms.append((pending, cur))
            pending, cur = op, v
    terms.append((pending, cur))
    total = Decimal(0)
    for sign, term in terms:
        total = term if sign is None else (total + term if sign == "+" else total - term)
    return total


def formula_text(srcs, ops, excel=False):
    parts = []
    for i, src in enumerate(srcs):
        if i:
            parts.append({"*": "\u00d7", "/": "\u00f7", "+": "+", "-": "\u2212"}[ops[i - 1]]
                         if not excel else ops[i - 1])
        parts.append(src)
    return (" ".join(parts)) if not excel else ("=" + "".join(parts))


EXPR_OK = re.compile(r"^[\d\s.,+\-*/x()^%=]+$", re.I)
TOK_RE = re.compile(r"\s*(\d+(?:\.\d+)?|\.\d+|[-+*/()^%])")


def _tokens(s):
    s = s.strip().rstrip("=").strip()
    s = re.sub(r"(?<=\d),(?=\d{3}(?!\d))", "", s)
    s = re.sub(r"(?<=[\d)\s])[xX](?=[\s\d(])", "*", s)
    out, pos = [], 0
    while pos < len(s):
        m = TOK_RE.match(s, pos)
        if not m:
            if s[pos:].strip() == "":
                break
            raise ValueError("bad token")
        out.append(m.group(1))
        pos = m.end()
    return out


def evaluate(expr):
    toks = _tokens(expr)
    n_nums = sum(1 for t in toks if t[0].isdigit() or t[0] == ".")
    if n_nums < 2 or not any(t in "+-*/^" for t in toks[1:]):
        raise ValueError("not an expression")
    i = 0

    def peek():
        return toks[i] if i < len(toks) else None

    def take():
        nonlocal i
        i += 1
        return toks[i - 1]

    def postfix(v):
        while peek() == "%":
            take()
            v = v / 100
        return v

    def atom():
        t = peek()
        if t in ("-", "+"):
            take()
            v = atom()
            return -v if t == "-" else v
        if t == "(":
            take()
            v = add()
            if peek() != ")":
                raise ValueError(")")
            take()
            return postfix(v)
        if t is None or not (t[0].isdigit() or t[0] == "."):
            raise ValueError("unexpected")
        take()
        return postfix(Decimal(t))

    def power():
        v = atom()
        if peek() == "^":
            take()
            v = v ** power()
        return v

    def mul():
        v = power()
        while peek() in ("*", "/"):
            op = take()
            r = power()
            v = v * r if op == "*" else v / r
        return v

    def add():
        v = mul()
        while peek() in ("+", "-"):
            op = take()
            r = mul()
            v = v + r if op == "+" else v - r
        return v

    v = add()
    if i != len(toks):
        raise ValueError("trailing")
    return v


def analyze(raw):
    text = normalize(raw or "")
    nums = extract_numbers(text)
    expr_val = None
    stripped = text.strip()
    is_date = bool(re.fullmatch(r"\d{1,4}[-/]\d{1,2}[-/]\d{1,4}", stripped))
    looks_expr = bool(stripped) and not is_date and len(stripped) < 500 and bool(EXPR_OK.match(stripped))
    ops = seed_ops(text, nums, looks_expr)
    if looks_expr:
        try:
            expr_val = evaluate(stripped)
        except (ValueError, InvalidOperation, ArithmeticError):
            expr_val = None
    return nums, ops, expr_val


def _round(d, decimals):
    if decimals is None:
        return d
    return d.quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP)


def fmt(d, decimals=None):
    if d is None:
        return "—"
    if decimals is not None:
        return f"{_round(d, decimals):,.{decimals}f}"
    if d == d.to_integral_value():
        return f"{d:,.0f}"
    s = f"{d.quantize(Decimal('1.0000000000')).normalize():,f}"
    return s.rstrip("0").rstrip(".") if "." in s else s


def plain(d, decimals=None):
    return fmt(d, decimals).replace(",", "")


def parse_hotkey(spec):
    """'ctrl+alt+s' -> (modifier flags, virtual key)."""
    mods, vk = 0, None
    table = {"ctrl": 0x2, "control": 0x2, "alt": 0x1, "shift": 0x4, "win": 0x8, "windows": 0x8}
    named = {"space": 0x20, "enter": 0x0D, "tab": 0x09, "insert": 0x2D, "home": 0x24,
             "end": 0x23, "pageup": 0x21, "pagedown": 0x22, "`": 0xC0, "=": 0xBB, "-": 0xBD}
    for part in spec.lower().replace(" ", "").split("+"):
        if part in table:
            mods |= table[part]
        elif len(part) == 1 and part.isalnum():
            vk = ord(part.upper())
        elif re.fullmatch(r"f([1-9]|1\d|2[0-4])", part):
            vk = 0x6F + int(part[1:])
        elif part in named:
            vk = named[part]
        else:
            raise ValueError(f"unknown key '{part}'")
    if vk is None:
        raise ValueError("hotkey needs a non-modifier key")
    return mods, vk


if "--test" in sys.argv:
    cases = ["1,250.50\n(300)\n-49.5", "Total SAR 1,000 and 2,500.75 (Q3 2025)",
             "1,250 \u00d7 3 \u2212 (400 / 2)", "\u0661\u0662\u0663\u066b\u0665 + \u0666", "100 / 8",
             "2026-09-19", "5 / 0", "nothing", "100 + 50 * 2", "12 x 4", "1500 2750.25 (250.25) 600"]
    for c in cases:
        nums, ops, e = analyze(c)
        vals = [n[0] for n in nums]
        srcs = [n[1] for n in nums]
        print(repr(c), "->", formula_text(srcs, ops), "=", fmt(build(vals, ops)),
              "| as written:", e, "| excel:", formula_text([plain(v) for v in vals], ops, excel=True))
    v = [Decimal(x) for x in ("100", "50", "2")]
    assert build(v, ["+", "*"]) == 200 and build(v, ["*", "+"]) == 5002
    assert build(v, ["-", "/"]) == 75 and build(v, ["/", "/"]) == 1
    assert build([Decimal(5), Decimal(0)], ["/"]) is None
    assert build([Decimal("1250.5")], []) == Decimal("1250.5")
    print("build() assertions ok")
    sys.exit(0)

# ======================================================================== Windows
import tkinter as tk
import tkinter.font as tkfont
import winreg
from ctypes import wintypes

import pystray
from PIL import Image, ImageDraw, ImageFont

u32 = ctypes.WinDLL("user32", use_last_error=True)
k32 = ctypes.WinDLL("kernel32", use_last_error=True)
LRESULT = ctypes.c_ssize_t
WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)


class WNDCLASSW(ctypes.Structure):
    _fields_ = [("style", wintypes.UINT), ("lpfnWndProc", WNDPROC), ("cbClsExtra", ctypes.c_int),
                ("cbWndExtra", ctypes.c_int), ("hInstance", wintypes.HINSTANCE),
                ("hIcon", wintypes.HICON), ("hCursor", wintypes.HANDLE),
                ("hbrBackground", wintypes.HBRUSH), ("lpszMenuName", wintypes.LPCWSTR),
                ("lpszClassName", wintypes.LPCWSTR)]


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD), ("dwFlags", wintypes.DWORD),
                ("time", wintypes.DWORD), ("dwExtraInfo", ctypes.c_size_t)]


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG), ("mouseData", wintypes.DWORD),
                ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD), ("dwExtraInfo", ctypes.c_size_t)]


class _INPUTU(ctypes.Union):
    _fields_ = [("ki", KEYBDINPUT), ("mi", MOUSEINPUT)]


class INPUT(ctypes.Structure):
    _fields_ = [("type", wintypes.DWORD), ("u", _INPUTU)]


u32.DefWindowProcW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
u32.DefWindowProcW.restype = LRESULT
u32.CreateWindowExW.argtypes = [wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD,
                                ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                wintypes.HWND, wintypes.HMENU, wintypes.HINSTANCE, wintypes.LPVOID]
u32.CreateWindowExW.restype = wintypes.HWND
u32.RegisterClassW.argtypes = [ctypes.POINTER(WNDCLASSW)]
u32.RegisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int, wintypes.UINT, wintypes.UINT]
u32.UnregisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int]
u32.AddClipboardFormatListener.argtypes = [wintypes.HWND]
u32.GetMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT]
u32.TranslateMessage.argtypes = [ctypes.POINTER(wintypes.MSG)]
u32.DispatchMessageW.argtypes = [ctypes.POINTER(wintypes.MSG)]
u32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
u32.SendInput.argtypes = [wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int]
u32.GetAsyncKeyState.argtypes = [ctypes.c_int]
u32.GetAsyncKeyState.restype = ctypes.c_short
u32.OpenClipboard.argtypes = [wintypes.HWND]
u32.GetClipboardData.argtypes = [wintypes.UINT]
u32.GetClipboardData.restype = wintypes.HANDLE
u32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
u32.SetClipboardData.restype = wintypes.HANDLE
u32.GetClipboardSequenceNumber.restype = wintypes.DWORD
u32.GetParent.argtypes = [wintypes.HWND]
u32.GetParent.restype = wintypes.HWND
u32.GetForegroundWindow.restype = wintypes.HWND
u32.SetForegroundWindow.argtypes = [wintypes.HWND]
u32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
u32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int,
                             ctypes.c_int, ctypes.c_int, wintypes.UINT]
u32.GetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
u32.GetWindowLongPtrW.restype = ctypes.c_ssize_t
u32.SetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_ssize_t]
u32.SetWindowLongPtrW.restype = ctypes.c_ssize_t
k32.GlobalLock.argtypes = [wintypes.HGLOBAL]
k32.GlobalLock.restype = ctypes.c_void_p
k32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
k32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
k32.GlobalAlloc.restype = wintypes.HGLOBAL
k32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
k32.GetModuleHandleW.restype = wintypes.HMODULE
k32.CreateMutexW.restype = wintypes.HANDLE

CF_UNICODETEXT = 13
WM_HOTKEY, WM_CLIPBOARDUPDATE, WM_APP_REHOTKEY = 0x0312, 0x031D, 0x8001
MOD_NOREPEAT = 0x4000
HOTKEY_ID = 1
VK_CTRL, VK_ALT, VK_SHIFT, VK_LWIN, VK_RWIN, VK_C = 0x11, 0x12, 0x10, 0x5B, 0x5C, 0x43


# ---------------------------------------------------------------- clipboard / keys
def clip_seq():
    return u32.GetClipboardSequenceNumber()


def _open_clip():
    for _ in range(25):
        if u32.OpenClipboard(None):
            return True
        time.sleep(0.01)
    return False


def clip_get():
    if not _open_clip():
        return None
    try:
        h = u32.GetClipboardData(CF_UNICODETEXT)
        if not h:
            return None
        p = k32.GlobalLock(h)
        if not p:
            return None
        try:
            return ctypes.wstring_at(p)
        finally:
            k32.GlobalUnlock(h)
    finally:
        u32.CloseClipboard()


def clip_set(text):
    buf = ctypes.create_unicode_buffer(text)
    size = ctypes.sizeof(buf)
    if not _open_clip():
        return False
    try:
        u32.EmptyClipboard()
        h = k32.GlobalAlloc(0x0002, size)
        p = k32.GlobalLock(h)
        ctypes.memmove(p, buf, size)
        k32.GlobalUnlock(h)
        u32.SetClipboardData(CF_UNICODETEXT, h)
        return True
    finally:
        u32.CloseClipboard()


def send_keys(seq):
    arr = (INPUT * len(seq))()
    for i, (vk, up) in enumerate(seq):
        arr[i].type = 1
        arr[i].u.ki = KEYBDINPUT(vk, 0, 0x2 if up else 0, 0, 0)
    u32.SendInput(len(seq), arr, ctypes.sizeof(INPUT))


def key_down(vk):
    return bool(u32.GetAsyncKeyState(vk) & 0x8000)


u32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
u32.GetWindowThreadProcessId.restype = wintypes.DWORD
u32.AttachThreadInput.argtypes = [wintypes.DWORD, wintypes.DWORD, wintypes.BOOL]
u32.BringWindowToTop.argtypes = [wintypes.HWND]
u32.SetFocus.argtypes = [wintypes.HWND]


def force_foreground(hwnd):
    """Windows only lets the foreground thread hand over focus, so borrow its input queue."""
    fg = u32.GetForegroundWindow()
    me = k32.GetCurrentThreadId()
    other = u32.GetWindowThreadProcessId(fg, None) if fg else 0
    attached = bool(other and other != me and u32.AttachThreadInput(me, other, True))
    try:
        u32.BringWindowToTop(hwnd)
        u32.SetForegroundWindow(hwnd)
        u32.SetFocus(hwnd)
    finally:
        if attached:
            u32.AttachThreadInput(me, other, False)


# ---------------------------------------------------------------- UI Automation reader
class SelectionReader:
    """Reads the focused control's selected text via UI Automation on a dedicated COM thread."""

    def __init__(self):
        self.req, self.res = queue.Queue(), queue.Queue()
        self.ok = True
        self.thread = None
        self._start()

    def _start(self):
        self.req, self.res = queue.Queue(), queue.Queue()
        self.thread = threading.Thread(target=self._run, daemon=True)
        self.thread.start()

    def _run(self):
        try:
            import comtypes
            import comtypes.client
            comtypes.CoInitializeEx(comtypes.COINIT_MULTITHREADED)
            try:
                from comtypes.gen import UIAutomationClient as UIA
            except ImportError:
                comtypes.client.GetModule("UIAutomationCore.dll")
                from comtypes.gen import UIAutomationClient as UIA
            auto = comtypes.client.CreateObject(UIA.CUIAutomation, interface=UIA.IUIAutomation)
        except Exception:
            self.ok = False
            return
        while True:
            token = self.req.get()
            text = None
            try:
                el = auto.GetFocusedElement()
                pat = el.GetCurrentPattern(10014)  # UIA_TextPatternId
                if pat:
                    tp = pat.QueryInterface(UIA.IUIAutomationTextPattern)
                    sel = tp.GetSelection()
                    parts = [sel.GetElement(i).GetText(200000) for i in range(sel.Length)]
                    text = "\n".join(p for p in parts if p)
            except Exception:
                text = None
            self.res.put((token, text))

    def read(self, timeout=0.35):
        if not self.ok:
            return None
        token = object()
        self.req.put(token)
        t_end = time.time() + timeout
        while True:
            try:
                tok, text = self.res.get(timeout=max(0.001, t_end - time.time()))
            except queue.Empty:
                self._start()  # app hung the COM call: abandon that thread
                return None
            if tok is token:
                return text


# ---------------------------------------------------------------- config / autostart
def load_cfg():
    cfg = json.loads(json.dumps(DEFAULT_CFG))
    try:
        with open(CFG_PATH, encoding="utf-8-sig") as f:
            cfg.update(json.load(f))
    except Exception:
        pass
    return cfg


def save_cfg(cfg):
    try:
        os.makedirs(CFG_DIR, exist_ok=True)
        tmp = CFG_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
        os.replace(tmp, CFG_PATH)
    except Exception:
        pass


RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"


def _exe_cmd():
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"'
    pyw = sys.executable.replace("python.exe", "pythonw.exe")
    return f'"{pyw}" "{os.path.abspath(__file__)}"'


def autostart_enabled():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as k:
            winreg.QueryValueEx(k, APP)
            return True
    except OSError:
        return False


def set_autostart(on):
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
        if on:
            winreg.SetValueEx(k, APP, 0, winreg.REG_SZ, _exe_cmd())
        else:
            try:
                winreg.DeleteValue(k, APP)
            except OSError:
                pass


# ---------------------------------------------------------------- popup
BG, FG, MUTED, DIM, ACCENT, TAB_BG, BORDER = ("#1f2329", "#f2f4f7", "#9aa3ad", "#6b737d",
                                              "#4cc38a", "#2c323a", "#3a414b")
CHIP_ON, CHIP_OFF = ("#2f3b35", "#262a30")
MAX_CHIPS = 60
OP_SYM = {"+": "+", "-": "\u2212", "*": "\u00d7", "/": "\u00f7"}
OP_CYCLE = ["+", "-", "*", "/"]
GWL_EXSTYLE, WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TOPMOST = -20, 0x08000000, 0x80, 0x8
SW_HIDE, SW_SHOWNOACTIVATE, SW_SHOW = 0, 4, 5
HWND_TOPMOST = wintypes.HWND(-1)
SWP_NOSIZE, SWP_NOACTIVATE, SWP_SHOWWINDOW = 0x1, 0x10, 0x40


class Popup:
    """Formula builder: numbers as chips, a clickable operator between each pair."""

    def __init__(self, app):
        self.app = app
        w = self.win = tk.Toplevel(app.root)
        w.overrideredirect(True)
        w.configure(bg=BG, highlightthickness=1, highlightbackground=BORDER)
        w.attributes("-alpha", 0.0)
        self.body = tk.Frame(w, bg=BG, padx=14, pady=10)
        self.body.pack()
        self.f_big = tkfont.Font(family="Segoe UI Semibold", size=22)
        self.f_small = tkfont.Font(family="Segoe UI", size=9)
        self.f_hint = tkfont.Font(family="Segoe UI", size=8)
        self.f_chip = tkfont.Font(family="Consolas", size=10)
        self.f_chip_off = tkfont.Font(family="Consolas", size=10, overstrike=1)
        self.f_op = tkfont.Font(family="Segoe UI Semibold", size=10)
        w.bind("<Enter>", lambda e: setattr(self, "hover", True))
        w.bind("<Leave>", lambda e: setattr(self, "hover", False))
        w.bind("<Key>", self.on_key)
        w.update_idletasks()
        self.hwnd = u32.GetParent(w.winfo_id()) or w.winfo_id()
        ex = u32.GetWindowLongPtrW(self.hwnd, GWL_EXSTYLE)
        u32.SetWindowLongPtrW(self.hwnd, GWL_EXSTYLE, ex | WS_EX_TOOLWINDOW | WS_EX_TOPMOST)
        u32.ShowWindow(self.hwnd, SW_HIDE)
        self.visible = False
        self.hover = False
        self.focused_mode = False
        self.prev_fg = None
        self.deadline = 0
        self._tick_id = None
        self.size = (0, 0)
        self.nums, self.ops, self.included, self.sel = [], [], [], 0

    # ---------------------------------------------------------------- state
    def show(self, nums, ops, expr_val, focus):
        self.nums = nums
        self.ops = list(ops)
        self.included = [True] * len(nums)
        self.expr_val = expr_val
        self.sel = 0
        self.hover = False
        self.size = (0, 0)
        self.render()
        self.present(focus)

    def live(self):
        """Included (value, src) pairs and the operators that join them."""
        idx = [i for i, on in enumerate(self.included) if on]
        vals = [self.nums[i][0] for i in idx]
        srcs = [self.nums[i][1] for i in idx]
        ops = [self.ops[a] for a, b in zip(idx, idx[1:])]
        return vals, srcs, ops, idx

    def result(self):
        vals, _, ops, _ = self.live()
        return build(vals, ops)

    def op_slots(self):
        """Indices into self.ops that are live (between two included numbers)."""
        idx = [i for i, on in enumerate(self.included) if on]
        return [a for a, b in zip(idx, idx[1:])]

    # ---------------------------------------------------------------- drawing
    def render(self):
        for c in self.body.winfo_children():
            c.destroy()
        dec = self.app.cfg.get("decimals")
        if not self.nums:
            tk.Label(self.body, text="No numbers found", bg=BG, fg=FG, font=("Segoe UI", 12)).pack()
            self._tally_row(dec)
            self.hint = tk.Label(self.body, text="Esc to close", bg=BG, fg=DIM, font=self.f_hint)
            self.hint.pack(anchor="w", pady=(4, 0))
            return

        self._formula_row()
        total = self.result()
        self.value = tk.Label(self.body, text=fmt(total, dec) if total is not None else "\u2014",
                              bg=BG, fg=FG, cursor="hand2", font=self.f_big)
        self.value.pack(anchor="w", pady=(8, 0))
        self.value.bind("<Button-1>", lambda e: self.copy())

        vals, srcs, ops, _ = self.live()
        n_all, n_on = len(self.nums), len(vals)
        info = f"{n_on} number{'s' if n_on != 1 else ''}" if n_on == n_all else f"{n_on} of {n_all} numbers"
        if vals:
            info += (f"  \u00b7  sum {fmt(sum(vals, Decimal(0)), dec)}"
                     f"  \u00b7  avg {fmt(sum(vals, Decimal(0)) / len(vals), dec if dec is not None else 2)}"
                     f"  \u00b7  min {fmt(min(vals), dec)}  \u00b7  max {fmt(max(vals), dec)}")
        tk.Label(self.body, text=info, bg=BG, fg=MUTED, font=self.f_small).pack(anchor="w")
        if self.expr_val is not None and total is not None and self.expr_val != total:
            tk.Label(self.body, text=f"as written (with brackets): {fmt(self.expr_val, dec)}",
                     bg=BG, fg=MUTED, font=self.f_small).pack(anchor="w")

        setall = tk.Frame(self.body, bg=BG)
        setall.pack(anchor="w", pady=(6, 0))
        tk.Label(setall, text="Set all:", bg=BG, fg=DIM, font=self.f_hint).pack(side="left", padx=(0, 4))
        for op in OP_CYCLE:
            b = tk.Label(setall, text=OP_SYM[op], bg=TAB_BG, fg=MUTED, padx=8, pady=1,
                         font=self.f_op, cursor="hand2")
            b.pack(side="left", padx=(0, 3))
            b.bind("<Button-1>", lambda e, op=op: self.set_all(op))
        self._tally_row(dec, parent=setall)
        self.hint = tk.Label(self.body, text="click an operator or a number \u00b7 \u2190\u2192 pick \u00b7 "
                                             "\u2191\u2193 cycle + \u2212 \u00d7 \u00f7 \u00b7 "
                                             "Enter copy \u00b7 Shift+Enter Excel \u00b7 Esc",
                             bg=BG, fg=DIM, font=self.f_hint)
        self.hint.pack(anchor="w", pady=(5, 0))

    def _formula_row(self):
        wrap = tk.Frame(self.body, bg=BG)
        wrap.pack(anchor="w")
        slots = self.op_slots()
        row, width, max_w = None, 0, 470
        shown = min(len(self.nums), MAX_CHIPS)

        def new_row():
            r = tk.Frame(wrap, bg=BG)
            r.pack(anchor="w", pady=1)
            return r

        row = new_row()
        for i in range(shown):
            val, src = self.nums[i][0], self.nums[i][1]
            w_px = self.f_chip.measure(src) + 20
            if width + w_px > max_w:
                row, width = new_row(), 0
            on = self.included[i]
            c = tk.Label(row, text=src, bg=CHIP_ON if on else CHIP_OFF, fg=FG if on else DIM,
                         font=self.f_chip if on else self.f_chip_off, padx=6, pady=2, cursor="hand2")
            c.pack(side="left")
            c.bind("<Button-1>", lambda e, i=i: self.toggle(i))
            width += w_px
            if i in slots:
                selected = slots and slots[self.sel % len(slots)] == i
                o = tk.Label(row, text=OP_SYM[self.ops[i]], bg=ACCENT if selected else TAB_BG,
                             fg="#0d1a12" if selected else FG, font=self.f_op, padx=6, pady=2,
                             cursor="hand2")
                o.pack(side="left", padx=3)
                o.bind("<Button-1>", lambda e, i=i: self.cycle(i))
                width += self.f_op.measure(OP_SYM[self.ops[i]]) + 18
        if len(self.nums) > MAX_CHIPS:
            tk.Label(wrap, text=f"+{len(self.nums) - MAX_CHIPS} more (added)", bg=BG, fg=DIM,
                     font=self.f_hint).pack(anchor="w")

    def _tally_row(self, dec, parent=None):
        t = self.app.tally_values()
        row = parent if parent is not None else tk.Frame(self.body, bg=BG)
        if parent is None:
            row.pack(anchor="w", pady=(6, 0), fill="x")
        add = tk.Label(row, text="+ Add to tally", bg=TAB_BG, fg=MUTED, padx=6, pady=1,
                       font=self.f_hint, cursor="hand2")
        add.pack(side="left", padx=(12, 0))
        add.bind("<Button-1>", lambda e: self.add_tally())
        if t:
            tk.Label(row, text=f"Tally {fmt(sum(t, Decimal(0)), dec)} ({len(t)})", bg=BG, fg=ACCENT,
                     font=self.f_hint).pack(side="left", padx=(8, 0))
            for label, fn in (("Copy", self.copy_tally), ("Clear", self.clear_tally)):
                b = tk.Label(row, text=label, bg=TAB_BG, fg=MUTED, padx=6, pady=1,
                             font=self.f_hint, cursor="hand2")
                b.pack(side="left", padx=(4, 0))
                b.bind("<Button-1>", lambda e, fn=fn: fn())

    def refresh(self):
        self.render()
        self._place(keep=True)
        self.bump()
        if self.focused_mode and self.visible:
            for d in (1, 40, 120):
                self.win.after(d, self._refocus)

    def _refocus(self):
        if self.visible and u32.GetForegroundWindow() != self.hwnd:
            force_foreground(self.hwnd)
        if self.visible:
            self.win.focus_force()

    # ---------------------------------------------------------------- edits
    def cycle_sel(self, step):
        """Up/Down cycle the highlighted operator through + - x /."""
        slots = self.op_slots()
        if not slots:
            return
        i = slots[self.sel % len(slots)]
        self.ops[i] = OP_CYCLE[(OP_CYCLE.index(self.ops[i]) + step) % 4]
        self.refresh()

    def cycle(self, i):
        self.ops[i] = OP_CYCLE[(OP_CYCLE.index(self.ops[i]) + 1) % 4]
        slots = self.op_slots()
        if i in slots:
            self.sel = slots.index(i)
        self.refresh()

    def set_op(self, op):
        slots = self.op_slots()
        if not slots:
            return
        self.ops[slots[self.sel % len(slots)]] = op
        self.sel = min(self.sel + 1, len(slots) - 1)
        self.refresh()

    def set_all(self, op):
        self.ops = [op] * len(self.ops)
        self.refresh()

    def toggle(self, i):
        self.included[i] = not self.included[i]
        self.refresh()

    def move(self, d):
        slots = self.op_slots()
        if slots:
            self.sel = (self.sel + d) % len(slots)
            self.refresh()

    # ---------------------------------------------------------------- actions
    def on_key(self, e):
        k, ch = e.keysym, e.char
        if k == "Escape":
            return self.hide(restore_focus=True)
        if k == "Left":
            return self.move(-1)
        if k == "Right":
            return self.move(1)
        if k == "Up":
            return self.cycle_sel(1)
        if k == "Down":
            return self.cycle_sel(-1)
        if k == "Return":
            return self.copy(excel=bool(e.state & 0x1))
        if ch in ("+", "-", "*", "/"):
            return self.set_op(ch)
        if k in ("KP_Add", "plus"):
            return self.set_op("+")
        if k in ("KP_Subtract", "minus"):
            return self.set_op("-")
        if k in ("KP_Multiply", "asterisk", "x", "X"):
            return self.set_op("*")
        if k in ("KP_Divide", "slash"):
            return self.set_op("/")
        if k.lower() == "t":
            return self.add_tally()

    def copy(self, excel=False):
        vals, srcs, ops, _ = self.live()
        dec = self.app.cfg.get("decimals")
        if excel:
            cells = [("(" + plain(v) + ")" if v < 0 else plain(v)) for v in vals]
            s = formula_text(cells, ops, excel=True)
        else:
            total = self.result()
            if total is None:
                return
            s = plain(total, dec)
        self.app.copy_text(s)
        total = self.result()
        if total is not None:
            self.app.add_history(total, formula_text(srcs, ops)[:40])
        self.value.configure(fg=ACCENT)
        self.hint.configure(text=("copied Excel formula: " + s) if excel else f"copied {s}")
        self.bump(1.5)

    def add_tally(self):
        total = self.result()
        if total is None:
            return
        self.app.tally_add(total)
        self.refresh()
        self.hint.configure(text=f"added {fmt(total, self.app.cfg.get('decimals'))} to tally")

    def copy_tally(self):
        t = self.app.tally_values()
        if t:
            self.app.copy_text(plain(sum(t, Decimal(0)), self.app.cfg.get("decimals")))
            self.hint.configure(text="tally copied")
            self.bump(1.5)

    def clear_tally(self):
        self.app.tally_clear()
        self.refresh()

    # ---------------------------------------------------------------- window
    def _place(self, keep=False):
        w = self.win
        w.update_idletasks()
        ww, wh = w.winfo_reqwidth(), w.winfo_reqheight()
        # Resizing an active popup makes Windows drop its focus, so while it is open the
        # window only ever grows, in steps, and never shrinks.
        ww = max(360, -(-ww // 40) * 40)
        if keep:
            ww, wh = max(ww, self.size[0]), max(wh, self.size[1])
            if (ww, wh) == self.size:
                return
            x, y = self.pos
        else:
            pt = wintypes.POINT()
            u32.GetCursorPos(ctypes.byref(pt))
            sw, sh = w.winfo_screenwidth(), w.winfo_screenheight()
            x = min(max(0, pt.x + 16), sw - ww - 8)
            y = min(max(0, pt.y + 20), sh - wh - 48)
            self.pos = (x, y)
        sw, sh = w.winfo_screenwidth(), w.winfo_screenheight()
        x = min(x, max(0, sw - ww - 8))
        y = min(y, max(0, sh - wh - 48))
        self.pos = (x, y)
        self.size = (ww, wh)
        w.geometry(f"{ww}x{wh}+{x}+{y}")
        u32.SetWindowPos(self.hwnd, HWND_TOPMOST, x, y, ww, wh, SWP_NOACTIVATE)

    def present(self, focus):
        self.focused_mode = focus
        ex = u32.GetWindowLongPtrW(self.hwnd, GWL_EXSTYLE)
        ex = (ex & ~WS_EX_NOACTIVATE) if focus else (ex | WS_EX_NOACTIVATE)
        u32.SetWindowLongPtrW(self.hwnd, GWL_EXSTYLE, ex)
        self._place()
        self.win.attributes("-alpha", 0.0)
        if focus:
            self.prev_fg = u32.GetForegroundWindow()
            u32.ShowWindow(self.hwnd, SW_SHOW)
            force_foreground(self.hwnd)
            self.win.focus_force()
            self.win.after(30, self.win.focus_force)
        else:
            u32.ShowWindow(self.hwnd, SW_SHOWNOACTIVATE)
        self.visible = True
        self._fade(0)
        self.bump()
        if self._tick_id is None:
            self._tick()

    def _fade(self, step):
        steps = (0.35, 0.7, 0.97)
        if step < len(steps) and self.visible:
            self.win.attributes("-alpha", steps[step])
            self.win.after(12, self._fade, step + 1)

    def bump(self, secs=None):
        self.deadline = time.time() + (secs or self.app.cfg.get("popup_seconds", 6))

    def _tick(self):
        self._tick_id = None
        if not self.visible:
            return
        if self.hover:
            self.bump(1.5)
        if self.focused_mode and u32.GetForegroundWindow() != self.hwnd:
            self._refocus()
        if time.time() >= self.deadline:
            self.hide()
            return
        self._tick_id = self.win.after(200, self._tick)

    def hide(self, restore_focus=False):
        if not self.visible:
            return
        self.visible = False
        self.win.attributes("-alpha", 0.0)
        u32.ShowWindow(self.hwnd, SW_HIDE)
        if restore_focus and self.focused_mode and self.prev_fg:
            u32.SetForegroundWindow(self.prev_fg)



# ---------------------------------------------------------------- app
class App:
    def __init__(self):
        self.cfg = load_cfg()
        save_cfg(self.cfg)
        self.q = queue.Queue()
        self.busy = False
        self.ignore_until = 0.0
        self.ignore_seq = None
        self.root = tk.Tk()
        self.root.withdraw()
        self.popup = Popup(self)
        self.reader = SelectionReader()
        self.hotkey_ok = False
        self.icon = pystray.Icon(APP, make_icon(), self.tooltip(), menu=self.menu())
        threading.Thread(target=self.icon.run, daemon=True).start()
        threading.Thread(target=self.win32_loop, daemon=True).start()
        self.root.after(15, self.pump)

    # ---- state helpers (tk thread)
    def tally_values(self):
        return [Decimal(x) for x in self.cfg["tally"]]

    def tally_add(self, v):
        self.cfg["tally"].append(str(v))
        self._changed()

    def tally_clear(self):
        self.cfg["tally"] = []
        self._changed()

    def add_history(self, v, label):
        h = self.cfg["history"]
        entry = {"v": str(v), "label": label, "t": time.strftime("%H:%M")}
        if h and h[0]["v"] == entry["v"] and h[0]["label"] == label:
            h[0] = entry
        else:
            h.insert(0, entry)
        del h[10:]
        self._changed()

    def _changed(self):
        save_cfg(self.cfg)
        try:
            self.icon.title = self.tooltip()
            self.icon.update_menu()
        except Exception:
            pass

    def copy_text(self, s):
        self.ignore_until = time.time() + 0.5
        clip_set(s)

    def tooltip(self):
        t = self.tally_values()
        s = f"{APP} — {self.cfg['hotkey'].title()}"
        if t:
            s += f"\nTally: {fmt(sum(t, Decimal(0)), self.cfg.get('decimals'))}"
        return s

    # ---- tray
    def menu(self):
        M, I = pystray.Menu, pystray.MenuItem

        def history_items():
            h = self.cfg["history"]
            if not h:
                yield I("(empty)", None, enabled=False)
                return
            dec = self.cfg.get("decimals")
            for e in h:
                yield I(f"{fmt(Decimal(e['v']), dec)}    {e['label']} · {e['t']}", copier(e["v"]))
            yield M.SEPARATOR
            yield I("Clear history", lambda: self.q.put(("clear_history",)))

        def tally_items():
            t = self.tally_values()
            dec = self.cfg.get("decimals")
            yield I(f"Total: {fmt(sum(t, Decimal(0)), dec)}  ({len(t)} items)", None, enabled=False)
            yield I("Copy total", lambda: self.q.put(("copy_tally",)), enabled=bool(t))
            yield I("Clear", lambda: self.q.put(("clear_tally",)), enabled=bool(t))

        def copier(v):
            return lambda icon, item: self.q.put(("copy", v))

        def set_dec(d):
            return lambda icon, item: self.q.put(("decimals", d))

        return M(
            I(lambda item: f"Calculate selection  ({self.cfg['hotkey'].title()})",
              lambda: self.q.put(("grab",)), default=True),
            I("Tally", M(tally_items)),
            I("History", M(history_items)),
            I("Rounding", M(
                I("Full precision", set_dec(None), radio=True,
                  checked=lambda item: self.cfg.get("decimals") is None),
                I("2 decimals", set_dec(2), radio=True, checked=lambda item: self.cfg.get("decimals") == 2),
                I("0 decimals", set_dec(0), radio=True, checked=lambda item: self.cfg.get("decimals") == 0),
            )),
            M.SEPARATOR,
            I("Auto-calculate on copy", lambda: self.q.put(("toggle_auto",)),
              checked=lambda item: self.cfg["auto_copy"]),
            I("Start with Windows", lambda: set_autostart(not autostart_enabled()),
              checked=lambda item: autostart_enabled()),
            I("Settings file (change hotkey)", lambda: os.startfile(CFG_PATH)),
            I("Reload settings", lambda: self.q.put(("reload",))),
            M.SEPARATOR,
            I("Quit", lambda: self.q.put(("quit",))),
        )

    # ---- Win32 message thread: hotkey + clipboard listener
    def win32_loop(self):
        self._wndproc = WNDPROC(self._proc)
        hinst = k32.GetModuleHandleW(None)
        wc = WNDCLASSW()
        wc.lpfnWndProc = self._wndproc
        wc.hInstance = hinst
        wc.lpszClassName = "SumSelectMsgWnd"
        u32.RegisterClassW(ctypes.byref(wc))
        self.msg_hwnd = u32.CreateWindowExW(0, "SumSelectMsgWnd", APP, 0, 0, 0, 0, 0,
                                            wintypes.HWND(-3), None, hinst, None)
        u32.AddClipboardFormatListener(self.msg_hwnd)
        self._register_hotkey()
        msg = wintypes.MSG()
        while u32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            u32.TranslateMessage(ctypes.byref(msg))
            u32.DispatchMessageW(ctypes.byref(msg))

    def _register_hotkey(self):
        u32.UnregisterHotKey(self.msg_hwnd, HOTKEY_ID)
        try:
            mods, vk = parse_hotkey(self.cfg["hotkey"])
            self.hotkey_ok = bool(u32.RegisterHotKey(self.msg_hwnd, HOTKEY_ID, mods | MOD_NOREPEAT, vk))
        except ValueError:
            self.hotkey_ok = False
        if not self.hotkey_ok:
            self.q.put(("notify", f"Hotkey {self.cfg['hotkey']} is unavailable (used by another app). "
                                  "Change it in the settings file, then Reload settings."))

    def _proc(self, hwnd, msg, wparam, lparam):
        if msg == WM_HOTKEY:
            threading.Thread(target=self.grab_selection, daemon=True).start()
            return 0
        if msg == WM_CLIPBOARDUPDATE:
            self.on_clipboard_update()
            return 0
        if msg == WM_APP_REHOTKEY:
            self._register_hotkey()
            return 0
        return u32.DefWindowProcW(hwnd, msg, wparam, lparam)

    def on_clipboard_update(self):
        if not self.cfg["auto_copy"] or self.busy or time.time() < self.ignore_until:
            return
        text = clip_get()
        if text and len(text) < 50000:
            nums, ops, expr_val = analyze(text)
            if len(nums) >= 2:
                self.q.put(("show", nums, ops, expr_val, False))

    # ---- selection grab (worker thread)
    def grab_selection(self):
        if self.busy:
            return
        self.busy = True
        try:
            text = self.reader.read()
            if not text or not text.strip():
                text = self._grab_via_clipboard()
            nums, ops, expr_val = analyze(text or "")
            self.q.put(("show", nums, ops, expr_val, True))
        finally:
            self.ignore_until = time.time() + 0.3
            self.busy = False

    def _grab_via_clipboard(self):
        t0 = time.time()
        mods = (VK_CTRL, VK_ALT, VK_SHIFT, VK_LWIN, VK_RWIN)
        while time.time() - t0 < 1.0 and any(key_down(m) for m in mods):
            time.sleep(0.008)
        old = clip_get()
        seq = clip_seq()
        send_keys([(VK_CTRL, 0), (VK_C, 0), (VK_C, 1), (VK_CTRL, 1)])
        t0 = time.time()
        while clip_seq() == seq and time.time() - t0 < 0.6:
            time.sleep(0.005)
        if clip_seq() == seq:
            return None
        time.sleep(0.01)
        text = clip_get()
        if old is not None and text is not None and old != text:
            clip_set(old)
        return text

    # ---- tk-thread dispatcher
    def pump(self):
        try:
            while True:
                msg = self.q.get_nowait()
                kind = msg[0]
                if kind == "show":
                    self.popup.show(msg[1], msg[2], msg[3], focus=msg[4])
                elif kind == "grab":
                    threading.Thread(target=self.grab_selection, daemon=True).start()
                elif kind == "copy":
                    self.copy_text(plain(Decimal(msg[1]), self.cfg.get("decimals")))
                elif kind == "copy_tally":
                    t = self.tally_values()
                    if t:
                        self.copy_text(plain(sum(t, Decimal(0)), self.cfg.get("decimals")))
                elif kind == "clear_tally":
                    self.tally_clear()
                    if self.popup.visible:
                        self.popup.refresh()
                elif kind == "clear_history":
                    self.cfg["history"] = []
                    self._changed()
                elif kind == "decimals":
                    self.cfg["decimals"] = msg[1]
                    self._changed()
                elif kind == "toggle_auto":
                    self.cfg["auto_copy"] = not self.cfg["auto_copy"]
                    self._changed()
                elif kind == "reload":
                    keep = {k: self.cfg[k] for k in ("history", "tally")}
                    self.cfg = load_cfg()
                    self.cfg.update(keep)
                    u32.PostMessageW(self.msg_hwnd, WM_APP_REHOTKEY, 0, 0)
                    self._changed()
                elif kind == "notify":
                    try:
                        self.icon.notify(msg[1], APP)
                    except Exception:
                        pass
                elif kind == "quit":
                    self.icon.stop()
                    self.root.destroy()
                    os._exit(0)
        except queue.Empty:
            pass
        # Poll fast only while something is happening; idle ticks cost CPU for nothing.
        self.root.after(15 if (self.busy or self.popup.visible) else 50, self.pump)

    def run(self):
        self.root.mainloop()


def make_icon():
    """Use the generated icon.png when it ships with the app; fall back to a drawn glyph."""
    for base in (getattr(sys, "_MEIPASS", None), os.path.dirname(os.path.abspath(__file__))):
        if not base:
            continue
        path = os.path.join(base, "icon.png")
        if os.path.exists(path):
            try:
                return Image.open(path).convert("RGBA")
            except Exception:
                pass
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((2, 2, 62, 62), radius=14, fill=(76, 195, 138, 255))
    try:
        font = ImageFont.truetype("seguisb.ttf", 46)
    except OSError:
        font = ImageFont.load_default()
    d.text((32, 30), "Σ", font=font, fill=(13, 26, 18, 255), anchor="mm")
    return img


if __name__ == "__main__":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        pass
    k32.CreateMutexW(None, False, "Local\\SumSelectSingleton")
    if ctypes.get_last_error() == 183:
        sys.exit(0)
    App().run()
