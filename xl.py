"""Drive Microsoft Excel on macOS through AppleScript, and capture its window.

Everything here is Apple Events plus `screencapture`. Nothing is faked: the
pixels come from the real Excel window, chrome and all.
"""

from __future__ import annotations

import os
import re
import subprocess
import tempfile
import time

APP = "Microsoft Excel"


class ScriptError(RuntimeError):
    pass


def osa(script: str, timeout: float = 120.0) -> str:
    r = subprocess.run(["osascript", "-e", script], capture_output=True,
                       text=True, timeout=timeout)
    if r.returncode != 0:
        raise ScriptError(r.stderr.strip() or f"osascript failed: {script[:120]}")
    return r.stdout.strip()


def tell(body: str, timeout: float = 120.0) -> str:
    return osa(f'tell application "{APP}"\n{body}\nend tell', timeout)


def available() -> str | None:
    try:
        return tell("get version", timeout=30)
    except Exception:
        return None


def warm_up(timeout: float = 300.0) -> float:
    """Absorb Excel's first-document-Apple-Event stall.

    Measured on Excel 16.112: the FIRST scripting command that touches the
    document model after launch takes ~108 seconds while Excel builds its
    scripting bridge. Every command after it takes ~0.1s. Without this the first
    real call looks like a hang and every caller has to guess a timeout.
    """
    t0 = time.time()
    try:
        tell("count of workbooks", timeout=timeout)
    except Exception:
        pass
    return time.time() - t0


def open_book(path: str, visible: bool = True) -> str:
    """Open a workbook and return its name (Excel addresses workbooks by name).

    Always closes an already-open copy first. Excel addresses workbooks by NAME,
    so a stale copy left open from an earlier run satisfies a "is it open?"
    check and the file just rewritten on disk is never read - which silently
    measures and captures the previous geometry.
    """
    path = os.path.abspath(path)
    name = os.path.basename(path)
    tell(f'''
        try
            close workbook "{name}" saving no
        end try
        open POSIX file "{path}"
        {"activate" if visible else ""}
    ''', timeout=300)
    return name


def close_book(name: str) -> None:
    try:
        tell(f'close workbook "{name}" saving no', timeout=60)
    except Exception:
        pass


def set_calc_auto() -> None:
    tell("set calculation to calculation automatic")


def set_frame(name: str, n: int, probe: str) -> str:
    """Set n, then read a grid cell back. The read only returns once Excel has
    finished evaluating, so it doubles as a recalculation barrier."""
    return tell(f'''
        tell workbook "{name}"
            set value of range "A1" of worksheet "Grid" to {n}
            get value of range "{probe}" of worksheet "Grid"
        end tell
    ''')


def window_bounds(name: str) -> tuple[int, int, int, int]:
    """(x, y, w, h) of the Excel document window, in screen points.

    Excel's window has `bounds` as {left, top, right, bottom} - it does not
    answer the standard `position`/`size` pair, which return `missing value`.
    """
    out = tell(f'''
        set b to bounds of window 1 of workbook "{name}"
        set AppleScript's text item delimiters to ","
        return b as text
    ''')
    l, t, r, b = (int(float(v)) for v in out.split(","))
    return l, t, r - l, b - t


def set_window(name: str, x: int, y: int, w: int, h: int) -> None:
    """A zoomed window ignores a new size, so un-zoom it first."""
    tell(f'''
        tell window 1 of workbook "{name}"
            set zoomed to false
            set bounds to {{{x}, {y}, {x + w}, {y + h}}}
        end tell
    ''')


def set_zoom(name: str, pct: int) -> None:
    tell(f'set zoom of window 1 of workbook "{name}" to {pct}')


def select(name: str, ref: str, sheet: str = "Grid") -> None:
    """Select a cell so the formula bar shows a real formula."""
    tell(f'''
        activate object worksheet "{sheet}" of workbook "{name}"
        select range "{ref}" of worksheet "{sheet}" of workbook "{name}"
    ''')


def normal_view(name: str) -> None:
    tell(f'set view of window 1 of workbook "{name}" to normal view')


def scroll_home(name: str) -> None:
    """Put A1 at the top-left so the picture starts where the sheet does."""
    tell(f'''
        tell window 1 of workbook "{name}"
            set scroll row to 1
            set scroll column to 1
        end tell
    ''')


def screen_size() -> tuple[int, int]:
    """Usable screen area in points.

    Asked of Excel itself. Finder and System Events both need their own
    Automation grant, and without it the query hangs rather than failing.
    """
    try:
        out = tell("get {usable width, usable height}", timeout=20)
        w, h = (int(float(v)) for v in out.split(", ")[:2])
        if w > 200 and h > 200:
            return w, h
    except Exception:
        pass
    r = subprocess.run(["system_profiler", "SPDisplaysDataType"],
                       capture_output=True, text=True)
    m = re.search(r"Resolution: (\d+) x (\d+)(.*)", r.stdout)
    if not m:
        return 1440, 900
    w, h = int(m.group(1)), int(m.group(2))
    if "Retina" in m.group(3):
        w, h = w // 2, h // 2
    return w, h - 50                       # leave the menu bar out of it


LAST_GRAB_ERR = ""


def grab(rect: tuple[int, int, int, int], path: str) -> bool:
    """screencapture a screen rect. Needs Screen Recording permission.

    Two traps, both of which fail silently:

    - screencapture EXITS 0 when it did not write the file, so the return code
      proves nothing. The file has to be checked.
    - it refuses any destination whose filename starts with a dot -
      "cannot write file to intended destination" - so a tidy hidden scratch
      file like `.poll.png` never appears and looks exactly like a denied
      Screen Recording permission.
    """
    global LAST_GRAB_ERR
    x, y, w, h = rect
    r = subprocess.run(["screencapture", "-x", "-o", "-t", "png",
                        "-R", f"{x},{y},{w},{h}", path], capture_output=True,
                       text=True)
    ok = os.path.exists(path) and os.path.getsize(path) > 0
    LAST_GRAB_ERR = "" if ok else (r.stderr.strip() or "no file written")
    return ok


def can_capture() -> tuple[bool, str]:
    """Probe the Screen Recording permission with a small grab."""
    p = os.path.join(tempfile.gettempdir(), "xlcap_probe.png")
    ok = grab((0, 0, 64, 64), p)
    err = LAST_GRAB_ERR
    if os.path.exists(p):
        os.remove(p)
    return ok, ("" if ok else
                f"screencapture failed ({err}) - if it mentions permission, "
                "grant Screen Recording to the app running this (System "
                "Settings > Privacy & Security), then relaunch that app")
