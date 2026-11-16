#!/usr/bin/env python3
"""How long does Excel take to repaint the grid when `n` changes?

Conditional formatting is the one part of this design with no obvious ceiling:
every rule is evaluated against every cell in the range on every recalculation,
so 8,160 cells x 32 rules is 261k rule evaluations per frame. This measures it
before the pipeline commits to those numbers.

    ./.venv/bin/python bench_cf.py                 # 0 / 16 / 32 rules
    ./.venv/bin/python bench_cf.py --grid 96x54    # a smaller grid

The first document Apple Event after Excel launches can take ~2 minutes while
Excel builds its scripting bridge; everything after it is ~0.1s. That warm-up is
absorbed here rather than being mistaken for a slow recalculation.
"""

from __future__ import annotations

import argparse
import os
import statistics
import subprocess
import tempfile
import time

import numpy as np

import book
import xl


def timed(script: str, timeout: float = 300) -> float:
    t0 = time.time()
    subprocess.run(["osascript", "-e", script], capture_output=True, timeout=timeout)
    return (time.time() - t0) * 1000


def bench(path: str, focus: str, steps: int = 40) -> dict:
    name = xl.open_book(path)
    xl.set_calc_auto()
    E = 'tell application "Microsoft Excel"'
    for n in (2, 3, 4):
        timed(f'{E}\nset value of range "A1" of worksheet "Grid" of workbook '
              f'"{name}" to {n}\nend tell')
    t = [timed(f'{E}\nset value of range "A1" of worksheet "Grid" of workbook '
               f'"{name}" to {n}\nget value of range "{focus}" of worksheet '
               f'"Grid" of workbook "{name}"\nend tell')
         for n in range(10, 10 + steps)]
    o = [timed(f'{E}\nget value of range "{focus}" of worksheet "Grid" of '
               f'workbook "{name}"\nend tell') for _ in range(15)]
    xl.close_book(name)
    return {"total": statistics.median(t), "overhead": statistics.median(o),
            "p90": sorted(t)[int(0.9 * len(t))]}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--grid", default="120x68")
    ap.add_argument("--frames", type=int, default=120)
    ap.add_argument("--rules", default="0,16,32")
    ap.add_argument("--cell", type=int, default=16)
    args = ap.parse_args()
    w, h = (int(v) for v in args.grid.lower().split("x"))

    if not xl.available():
        raise SystemExit("Microsoft Excel is not responding to AppleScript")
    rng = np.random.default_rng(0)
    tmp = tempfile.mkdtemp(prefix="xlbench_")
    print(f"  grid {w}x{h} = {w * h:,} formula cells, {args.frames} frames")
    print(f"  {'rules':>6} {'total':>9} {'overhead':>9} {'recalc+CF':>10} "
          f"{'p90':>8}   fps")
    try:
        for r in (int(v) for v in args.rules.split(",")):
            k = max(1, r)
            frames = [rng.integers(0, k, (h, w), dtype=np.uint8)
                      for _ in range(args.frames)]
            pal = rng.integers(0, 255, (k, 3), dtype=np.uint8)
            p = os.path.join(tmp, f"bench{r}.xlsx")
            info = book.write_workbook(p, frames, list(range(args.frames)), pal,
                                       args.cell)
            if r == 0:                      # rewrite with the rules stripped out
                _strip_cf(p)
            m = bench(p, info["focus"])
            net = m["total"] - m["overhead"]
            print(f"  {r:>6} {m['total']:8.0f}ms {m['overhead']:8.0f}ms "
                  f"{net:9.0f}ms {m['p90']:7.0f}ms {1000 / max(net, 1):6.1f}")
    finally:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)
    return 0


def _strip_cf(path: str) -> None:
    """Drop every conditional formatting block, keeping the formulas."""
    import re
    import shutil
    import zipfile
    src = path + ".tmp"
    shutil.move(path, src)
    zin, zout = zipfile.ZipFile(src), zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        data = zin.read(it.filename)
        if it.filename.endswith("sheet1.xml"):
            data = re.sub(rb"<conditionalFormatting.*?</conditionalFormatting>",
                          b"", data, flags=re.S)
        zout.writestr(it, data)
    zin.close()
    zout.close()
    os.remove(src)


if __name__ == "__main__":
    raise SystemExit(main())
