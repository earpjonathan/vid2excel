#!/usr/bin/env python3
"""Turn a short clip into a colour animation that really runs inside Excel.

    decode -> autocrop -> downscale to the cell grid -> quantise -> emit

One command writes the .xlsx, screenshots real Excel one frame at a time, and
stacks the result over the original clip as a 1080x1920 reel.
"""

from __future__ import annotations

import argparse
import os
import shutil
import sys
import tempfile
import time

import numpy as np

import book
import capture
import reel
import video

DUR_LO, DUR_HI = 8.0, 15.0        # measured retention window, see README


def parse_grid(s: str) -> tuple[int, int]:
    try:
        w, h = (int(v) for v in s.lower().split("x"))
    except ValueError:
        raise SystemExit(f"--grid wants WxH, got {s!r}")
    return w, h


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("video")
    ap.add_argument("-o", "--out", required=True, help="output base path")
    ap.add_argument("--start", type=float, default=0.0)
    ap.add_argument("--duration", type=float, default=12.0)
    ap.add_argument("--fps", type=float, default=15.0)
    ap.add_argument("--grid", default="120x68", help="cells, WxH (default 120x68)")
    ap.add_argument("--cell", type=int, default=16, help="cell size in points")
    ap.add_argument("--palette", type=int, default=32, help="colours = CF rules")
    ap.add_argument("--tau", type=float, default=1.0,
                    help="palette hysteresis in Lab units, 0 disables")
    ap.add_argument("--no-crop", action="store_true")
    ap.add_argument("--no-reel", action="store_true", help="write the .xlsx only")
    ap.add_argument("--layout", choices=("reel", "excel"), default="reel",
                    help="reel = 1080x1920 Excel over the clip (default); "
                         "excel = 1920x1080 of the Excel window alone")
    ap.add_argument("--no-preview", action="store_true")
    ap.add_argument("--zoom", type=int, help="force Excel zoom %% for the capture")
    ap.add_argument("--max-wait", type=float, default=4.0,
                    help="seconds to wait for Excel to repaint a frame")
    ap.add_argument("--keep-frames", help="keep the raw captures in this folder")
    ap.add_argument("--limit", type=int, help="only this many frames - smoke test")
    ap.add_argument("--no-fit", action="store_true",
                    help="skip the measured square-cell correction")
    args = ap.parse_args()

    src = os.path.abspath(args.video)
    if not os.path.exists(src):
        raise SystemExit(f"no such file: {src}")
    base = os.path.abspath(args.out)
    os.makedirs(os.path.dirname(base) or ".", exist_ok=True)
    w, h = parse_grid(args.grid)
    aspect = 16 / 9 if args.layout == "excel" else 1080 / 960

    src_dur, src_fps = video.video_info(src)
    duration = args.duration if args.duration > 0 else src_dur
    duration = min(duration, max(0.0, src_dur - args.start))
    if args.layout == "reel" and not DUR_LO <= duration <= DUR_HI:
        print(f"  ! {duration:.1f}s is outside the {DUR_LO:g}-{DUR_HI:g}s window "
              f"that retains viewers - see README", file=sys.stderr)

    print(f"  source   {os.path.basename(src)}  {src_dur:.1f}s @ {src_fps:g}fps")
    crop = None if args.no_crop else video.detect_crop(src)
    print(f"  crop     {crop if crop else 'none needed'}")

    t0 = time.time()
    frames = list(video.sample_frames(src, args.fps, (w, h), args.start,
                                      duration, crop))
    if not frames:
        raise SystemExit("no frames sampled - check --start/--duration")
    if args.limit:
        frames = frames[:args.limit]
        duration = len(frames) / args.fps
    palette_bgr, centres = video.build_palette(frames, args.palette)
    idx = video.quantise(frames, centres, args.tau)
    distinct, order = video.dedupe(idx)
    print(f"  frames   {len(order)} at {args.fps:g}fps, {len(distinct)} distinct "
          f"({len(distinct) / len(order):.0%})")
    print(f"  palette  {args.palette} colours -> {args.palette} CF rules")

    xlsx = base + ".xlsx"
    info = book.write_workbook(xlsx, distinct, order, palette_bgr, args.cell,
                               title=os.path.basename(base))
    print(f"  workbook {xlsx}  ({os.path.getsize(xlsx) / 1e6:.1f} MB)")
    print(f"           {info['cells']:,} formula cells, "
          f"{info['cells']:,}x{info['distinct']} data cells, "
          f"cell {info['cell_px'][0]}x{info['cell_px'][1]}pt "
          f"(width {info['col_width']}, height {info['row_height']}pt)")

    if not args.no_preview:
        pv = base + ".preview.mp4"
        if reel.write_preview(pv, distinct, order, palette_bgr, args.fps,
                              src=src, start=args.start, duration=duration):
            print(f"  preview  {pv}")

    print(f"  image    {time.time() - t0:.1f}s")
    if args.no_reel:
        return 0

    # Calibrate on the most strongly painted frame, not on frame 1. Frame 1 is
    # whatever the clip happens to open on - here a dark car interior, which
    # measured as a fragment of the grid and produced a 38% column-width
    # "correction" that was not needed.
    ink = [float(np.abs(255 - palette_bgr[f].astype(int)).mean()) for f in distinct]
    best = int(np.argmax(ink))
    probe = order.index(best) + 1
    cell_w = float(args.cell)
    if not args.no_fit:
        scale = capture.fit_cells(xlsx, (w, h), cell_w, info["focus"],
                                  zoom=args.zoom, probe=probe, aspect=aspect)
        if abs(scale - 1) > 0.02:
            cell_w = args.cell * scale
            info = book.write_workbook(xlsx, distinct, order, palette_bgr,
                                       args.cell, title=os.path.basename(base),
                                       width_scale=scale)
            check = capture.fit_cells(xlsx, (w, h), cell_w, info["focus"],
                                      zoom=args.zoom, probe=probe, aspect=aspect)
            print(f"  fit      cells were {1 / scale:.3f} wide per tall - "
                  f"width x{scale:.3f}, now off by {abs(1 - check):.1%}")
        else:
            print(f"  fit      cells already square (off by {abs(1 - scale):.1%})")

    tmp = args.keep_frames or tempfile.mkdtemp(prefix="xlcap_")
    os.makedirs(tmp, exist_ok=True)
    try:
        def prog(n, total):
            print(f"\r    capturing {n}/{total}", end="", flush=True)

        t1 = time.time()
        stats = capture.capture_frames(xlsx, tmp, order, (w, h), cell_w,
                                       info["focus"], max_wait=args.max_wait,
                                       zoom=args.zoom, probe=probe, aspect=aspect,
                                       progress=prog)
        print()
        m = capture.measure(tmp, order, stats["box"])
        rate = stats["loop_s"] / max(1, len(order))
        print(f"  capture  {stats['rect']}  {rate * 1000:.0f}ms/frame, "
              f"{stats['polls'] / max(1, len(order)):.1f} polls/frame")
        print(f"  distinct source {m['source_distinct']:.1%}  "
              f"capture {m['capture_distinct']:.1%}"
              + (f"   MISSED {len(m['missed_frames'])}" if m["missed_frames"] else ""))
        if m["capture_distinct"] < m["source_distinct"] * 0.9:
            print("  ! the capture repeats more than the workbook does - the "
                  "repaint wait is too short. Raise --max-wait.", file=sys.stderr)

        if args.layout == "excel":
            out = base + ".excel.mp4"
            ok = reel.build_flat(tmp, src, args.start, duration, args.fps, out)
        else:
            out = base + ".reel.mp4"
            ok = reel.build_reel(tmp, src, args.start, duration, args.fps, out,
                                 crop=crop)
        print(f"  video    {out}" if ok else "  ! render failed")
    finally:
        if not args.keep_frames:
            shutil.rmtree(tmp, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
