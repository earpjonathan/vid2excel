"""Screenshot the real Excel window once per frame.

The whole difficulty is knowing WHEN to press the shutter. Excel recalculates
and repaints asynchronously, so grabbing straight after setting `n` returns the
previous frame - which is how the desmos-video project silently ended up with
88% duplicate frames and an animation that looked like it ran at a fraction of
its real rate. No output-fps setting repairs that, because the frames genuinely
are identical.

Waiting for the window to be "stable" is equally wrong: a window that has not
repainted yet is perfectly stable and passes the test instantly. So the wait is
always: change from a pre-update baseline FIRST, then settle.

The one frame where that cannot work is a frame that draws the same picture as
the last one - held frames are normal in cartoons. Those are known in advance
from the frame map, so they skip the change requirement instead of burning the
timeout. They are still captured for real, never copied, so the duplicate-frame
measurement at the end stays honest.
"""

from __future__ import annotations

import math
import os
import sys
import time

import cv2
import numpy as np

import xl


def fingerprint(img: np.ndarray, box) -> bytes:
    x0, y0, x1, y1 = box
    return cv2.cvtColor(img[y0:y1, x0:x1], cv2.COLOR_BGR2GRAY).tobytes()


def _longest_run(profile: np.ndarray, thresh: float, gap: int = 4) -> tuple[int, int]:
    """Start/end of the longest stretch above `thresh`, bridging small gaps."""
    on = profile >= thresh
    best = (0, 0, len(on))
    i = 0
    while i < len(on):
        if not on[i]:
            i += 1
            continue
        j, last = i, i
        while j < len(on):
            if on[j]:
                last = j
            elif j - last > gap:
                break
            j += 1
        if last + 1 - i > best[0]:
            best = (last + 1 - i, i, last + 1)
        i = j
    return best[1], best[2]


def _bands(profile: np.ndarray, thresh: float, gap: int = 4) -> list:
    """Every stretch above `thresh`, longest first."""
    on = profile >= thresh
    out, i = [], 0
    while i < len(on):
        if not on[i]:
            i += 1
            continue
        j, last = i, i
        while j < len(on):
            if on[j]:
                last = j
            elif j - last > gap:
                break
            j += 1
        out.append((i, last + 1))
        i = j
    return sorted(out, key=lambda b: b[0] - b[1])


def calibrate(name: str, rect, tmp: str, focus: str, probes, aspect=None) -> tuple:
    """Find the picture inside the window, to fingerprint that and not the chrome.

    Guessing at chrome heights is fragile - the ribbon can be collapsed, the
    formula bar can be two lines high, and the status bar ticks over during a
    recalculation, which would read as "still changing" forever.

    Use the painted block itself. An earlier version diffed two frames and took
    what moved, which is wrong for exactly the footage this is good at: two
    adjacent frames of a cartoon differ in a handful of cells, so the box came
    out as that handful and every repaint elsewhere was invisible - measured at
    73% distinct captured against 100% distinct source. The whole grid is the
    right region, whether or not this particular pair of frames moved in it.
    """
    xl.set_frame(name, probes[0], focus)
    time.sleep(0.8)
    if not xl.grab(rect, tmp):
        raise SystemExit("screencapture failed during calibration")
    box = picture_box(cv2.imread(tmp), aspect, tol=2.5)
    if box is None:                       # not a colourful frame - fall back
        shots = []
        for n in probes:
            xl.set_frame(name, n, focus)
            time.sleep(0.8)
            xl.grab(rect, tmp)
            shots.append(cv2.imread(tmp))
        d = cv2.cvtColor(cv2.absdiff(shots[0], shots[1]), cv2.COLOR_BGR2GRAY) > 8
        h, w = d.shape
        if d.sum() < 100:
            return int(w * .2), int(h * .2), int(w * .8), int(h * .8)
        rows = d.sum(1).astype(float)
        y0, y1 = _longest_run(rows, max(4.0, rows.max() * 0.05))
        cols = d[y0:y1].sum(0).astype(float)
        x0, x1 = _longest_run(cols, max(4.0, cols.max() * 0.05))
        box = (x0, y0, x1, y1)
    x0, y0, x1, y1 = box
    return max(0, x0 - 2), max(0, y0 - 2), x1 + 2, y1 + 2


def picture_box(img: np.ndarray, aspect: float | None = None,
                tol: float = 1.6) -> tuple | None:
    """Bounding box of the painted grid.

    A painted cell is anything that is not the sheet background: saturated
    colour, OR dark. Keying on saturation alone breaks on a dark frame - an
    underwater shot is nearly all low-saturation navy, so the box came out as a
    fragment and the measured cell aspect with it.

    Taking the LONGEST band is not enough either. In a short 16:9 window the
    ribbon is a full-width painted band of its own, and once the grid no longer
    towers over it the ribbon wins: measured as a 2904x346 box, which the fit
    step turned into a 1.94x column-width "correction".

    So when the grid's aspect is known, pick the band that matches it. A real
    cell-squareness error is a few percent - the ECMA width formula is close,
    just not exact - while the chrome bands are off by multiples, so a +-60%
    window separates them without assuming the answer to the finer measurement.
    `tol` is that window: strict when measuring squareness, loose when all that
    is wanted is a region that covers the grid.
    """
    v = img.max(2).astype(int)
    sat = v - img.min(2).astype(int)
    mask = (v < 225) | (sat > 25)
    rows = mask.sum(1).astype(float)
    if rows.max() < 20:
        return None
    best, best_score = None, 1e9
    # 0.20, not 0.50. Empty sheet measures 0.000 here - Excel's gridlines are
    # too light to trip the mask at all - while a band of picture that happens
    # to be mostly white still measures 0.48. The gap is wide, so the threshold
    # belongs near the bottom of it; at 0.50 a white background inside the shot
    # split the grid in two and only the right-hand piece was measured.
    for y0, y1 in _bands(rows, rows.max() * 0.20):
        if y1 - y0 < 20:
            continue
        cols = mask[y0:y1].sum(0).astype(float)
        if cols.max() < 20:
            continue
        x0, x1 = _longest_run(cols, cols.max() * 0.20)
        if x1 - x0 < 20:
            continue
        if aspect is None:
            score = -float((y1 - y0) * (x1 - x0))
        else:
            score = abs(math.log(((x1 - x0) / (y1 - y0)) / aspect))
        if score < best_score:
            best, best_score = (x0, y0, x1, y1), score
    if aspect is not None and best_score > math.log(tol):
        return None
    return best


def fit_cells(path: str, grid, cell_w_pt: float, focus: str, zoom=None,
              probe: int = 1, aspect: float = 1080 / 960) -> float:
    """How much to multiply the column width by to make the cells square.

    openpyxl's width unit is characters of the Normal font, which does not
    predict the on-screen aspect, and the ECMA width formula describes it only
    approximately. So measure what Excel actually drew rather than trusting the
    arithmetic - but only believe a small correction. A large one means the
    grid was not the thing that got measured, which has happened twice now, and
    applying it wrecks the geometry that was very nearly right to begin with.
    """
    w, h = grid
    name, rect = open_for_capture(path, grid, cell_w_pt, focus, zoom=zoom,
                                  aspect=aspect)
    tmp = os.path.join(os.path.dirname(os.path.abspath(path)), "fit.png")
    try:
        xl.set_frame(name, probe, focus)
        time.sleep(1.0)
        if not xl.grab(rect, tmp):
            return 1.0
        box = picture_box(cv2.imread(tmp), w / h, tol=1.6)
        if box is None:
            print("  ! could not find the grid to measure - leaving the column "
                  "width alone", file=sys.stderr)
            return 1.0
        x0, y0, x1, y1 = box
        scale = ((y1 - y0) / h) / ((x1 - x0) / w)
        if abs(scale - 1) > MAX_FIT:
            print(f"  ! measured a {scale:.2f}x width correction; the width "
                  f"formula is never that far out, so ignoring it",
                  file=sys.stderr)
            return 1.0
        return scale
    finally:
        xl.close_book(name)
        if os.path.exists(tmp):
            os.remove(tmp)


MAX_FIT = 0.20         # the ECMA width formula is never wronger than this
FIT_MARGIN = 0.92      # keep visible empty rows under the grid
CHROME_PT = 254        # measured: title bar + ribbon + formula bar + column
                       # header above the grid, sheet tabs + status bar below


def plan(screen: tuple[int, int], grid: tuple[int, int], cell_w_pt: float,
         aspect: float = 1080 / 960, margin: int = 60) -> tuple[int, int, int]:
    """(window width, window height, zoom %) that fits the grid on this screen.

    The workbook keeps human-sized cells; the capture shrinks them with Excel's
    own zoom instead, so the .xlsx someone opens is still pleasant to scrub.

    Both dimensions have to be fitted. A 9:8 window is tall enough that width
    always binds, so the first version only fitted width - but a 16:9 window is
    short, and the same zoom then runs the grid off the bottom of the screen.
    """
    sw, sh = screen
    win_h = max(400, sh - margin)
    win_w = int(round(win_h * aspect))
    if win_w > sw - margin:
        win_w = sw - margin
        win_h = int(round(win_w / aspect))
    w, h = grid
    # Leave the grid short of the available height on purpose. Filling it
    # exactly leaves no empty rows between the picture and the sheet tabs,
    # which both looks wrong and removes the gap that separates the grid from
    # the chrome when the box is detected.
    fit_w = (win_w - 46) / (w * cell_w_pt)          # row headers plus slack
    fit_h = (win_h - CHROME_PT) * FIT_MARGIN / (h * cell_w_pt)
    zoom = int(max(10, min(400, min(fit_w, fit_h) * 100)))
    return win_w, win_h, zoom


def open_for_capture(path: str, grid, cell_w_pt: float, focus: str,
                     zoom=None, aspect: float = 1080 / 960) -> tuple[str, tuple]:
    """Open the workbook, size the window, and return the rect to grab."""
    stall = xl.warm_up()
    if stall > 5:
        print(f"\r    (Excel scripting bridge warmed up in {stall:.0f}s)")
    name = xl.open_book(path)
    xl.set_calc_auto()
    win_w, win_h, auto_zoom = plan(xl.screen_size(), grid, cell_w_pt, aspect)
    xl.set_window(name, 0, 25, win_w, win_h)  # Excel clamps below the menu bar
    xl.normal_view(name)
    xl.set_zoom(name, zoom or auto_zoom)
    xl.scroll_home(name)
    xl.select(name, focus)                  # formula bar shows a real formula
    time.sleep(1.0)
    return name, xl.window_bounds(name)


def capture_frames(path: str, dst: str, order: list[int], grid, cell_w_pt: float,
                   focus: str, max_wait: float = 4.0, settle: int = 2,
                   zoom=None, probe: int = 1, aspect: float = 1080 / 960,
                   progress=None) -> dict:
    """Step n through every frame and screenshot the window each time."""
    ok, why = xl.can_capture()
    if not ok:
        raise SystemExit(why)
    os.makedirs(dst, exist_ok=True)
    name, rect = open_for_capture(path, grid, cell_w_pt, focus, zoom=zoom,
                                  aspect=aspect)
    tmp = os.path.join(dst, "poll.png")   # NOT a dotfile: see xl.grab
    box = calibrate(name, rect, tmp, focus, (probe, max(1, probe % len(order) + 1)))
    stats = {"rect": rect, "box": box, "missed": [], "polls": 0, "waits": []}

    loop0 = time.time()
    try:
        for i, di in enumerate(order):
            n = i + 1
            expect = i == 0 or di != order[i - 1]
            if not xl.grab(rect, tmp):
                raise SystemExit("screencapture failed mid-run")
            base = fingerprint(cv2.imread(tmp), box)
            xl.set_frame(name, n, focus)

            t0 = time.time()
            moved, last, stable, img = not expect, None, 0, None
            while True:
                if not xl.grab(rect, tmp):
                    raise SystemExit("screencapture failed mid-run")
                img = cv2.imread(tmp)
                stats["polls"] += 1
                fp = fingerprint(img, box)
                if fp != base:
                    moved = True
                if moved:
                    if fp == last:
                        stable += 1
                        if stable >= settle - 1:
                            break
                    else:
                        stable, last = 0, fp
                if time.time() - t0 > max_wait:
                    break
            stats["waits"].append(time.time() - t0)
            if expect and not moved:
                stats["missed"].append(n)   # repaint never arrived: a real miss
            cv2.imwrite(os.path.join(dst, f"f{n:06d}.png"), img)
            if progress and (n % 10 == 0 or n == len(order)):
                progress(n, len(order))
    finally:
        stats["loop_s"] = time.time() - loop0   # excludes Excel's warm-up
        if os.path.exists(tmp):
            os.remove(tmp)
        xl.close_book(name)
    return stats


def measure(dst: str, order: list[int], box) -> dict:
    """Duplicate rate of the CAPTURE against the duplicate rate of the SOURCE.

    Eyeballing the animation cannot tell a slow capture from slow source
    material - both look like a low frame rate. Comparing the two ratios can:
    if the capture repeats far more often than the workbook does, the wait logic
    is broken, not the clip.

    Compared over the calibrated picture box, so chrome that happens to differ
    (or not) between two shots cannot flatter or spoil the number.
    """
    n = len(order)
    src_changes = sum(1 for i in range(1, n) if order[i] != order[i - 1])
    cap_changes, missed, prev = 0, [], None
    for i in range(n):
        img = cv2.imread(os.path.join(dst, f"f{i + 1:06d}.png"))
        fp = None if img is None else fingerprint(img, box)
        if i and fp is not None and prev is not None:
            if fp != prev:
                cap_changes += 1
            elif order[i] != order[i - 1]:
                missed.append(i + 1)
        prev = fp
    den = max(1, n - 1)
    return {"frames": n,
            "source_distinct": src_changes / den,
            "capture_distinct": cap_changes / den,
            "source_changes": src_changes, "capture_changes": cap_changes,
            "missed_frames": missed}
