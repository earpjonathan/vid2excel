# vid2excel

The Excel sibling of [P.E.T.E.R](https://github.com/earpjonathan/P.E.T.E.R).
Write-up: https://projects.jonathanearp.xyz/desmos/#excel

Turn a short video clip into a **colour** animation that really runs inside an
Excel workbook, then composite it into a 9:16 reel.

```bash
./.venv/bin/python vid2excel.py clip.mp4 -o out/clip --start 35 --duration 11
```

Three files come out:

| file | what it is |
| --- | --- |
| `clip.xlsx` | a real workbook - open it, type a number in `A1`, watch it repaint |
| `clip.reel.mp4` | 1080x1920, real Excel on top, the clip below, original audio |
| `clip.preview.mp4` | the quantised grid as a plain video, for judging the palette |

`--no-reel` writes only the workbook. `--limit N` captures only N frames, which
is the fast way to check framing before committing to a full run.

## How it works

There is no edge detection, no tracing and no curve fitting - Excel fills a cell
with any colour, so the whole image pipeline is:

    decode -> autocrop -> downscale to the cell grid -> quantise -> emit

The workbook has three sheets:

| sheet | contents |
| --- | --- |
| `Grid` | 120x68 cells, one per pixel, each holding `=INDEX(Data!<its own row>,1,k)` |
| `Data` | one row per grid cell, one column per **distinct** frame |
| `Map` | one row: output frame `n` -> the `Data` column that draws it |

`A1` holds `n` and `A2` holds `k = INDEX(Map!…,n)`. Type a new `n` and 8,160
formulas recalculate, each pulling its own palette index out of `Data`.

A formula cannot set a fill, so **colour comes from conditional formatting**:
one `cellIs equal` rule per palette entry over the whole grid range, each with
`stopIfTrue`. The cells' own digits are hidden with the `;;;` number format (and
a matching font colour in each rule, as a second line of defence).

## Measured findings

Everything below was measured on this machine - Excel 16.112.1, M5 Pro, macOS 26
- not assumed.

### Square cells are not what the spec says

Excel measures column width in characters of the Normal font's widest digit and
row height in points, which do not share a unit. The documented conversion
(ECMA-376 18.3.1.13, `Truncate((px-5)/MDW*100+0.5)/100` with `MDW=7`) is a
**Windows 96-dpi convention and is wrong here**. Using it produced 9x12pt cells
- a picture stretched 33% vertically.

Measured instead, by writing known widths and asking Excel for the rendered size:

| width attribute | 1.0 | 1.5 | 2.0 | 3.0 | 4.0 | 6.0 | 8.0 | 12.0 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| rendered points | 6 | 9 | 12 | 18 | 24 | 36 | 48 | 72 |

Exactly `rendered = 6 * width`, straight through the origin, while row height in
points renders 1:1. So a square cell of `P` points is `width = P/6`, `height = P`.

Two further traps here. Excel's own `Range.Width` reports the cell as exactly as
wide as it is tall while it renders 10.4 x 12.0 device pixels, so **asking Excel
does not settle it**. And measuring the picture in a screenshot by taking the
extent of every saturated pixel does not either - the ribbon icons and the green
Share button are saturated too, and they stretch the box across the whole
window. The grid is the one large *contiguous* block, so both `picture_box()`
and `calibrate()` take the longest run of saturated rows, then the longest run
of columns *within those rows*.

Because that constant is only true for one Excel build, font and display, the
pipeline does not trust it: `fit_cells()` captures one frame, measures the
picture Excel actually drew, and rewrites the workbook if the cells are more
than 2% off square. On this machine it reports 0.2% and changes nothing.

### Conditional formatting is the expensive part

Time for `n` to change and the grid to finish recalculating, 8,160 formula
cells, measured as an Apple Event round trip with the ~108ms event overhead
subtracted:

| CF rules | recalc + CF | live fps in Excel |
| --- | --- | --- |
| 0 (formulas only) | 35 ms | 28.9 |
| 16 | 136 ms | 7.4 |
| 32 | 211 ms | 4.7 |

CF costs about **5.5 ms per rule** over this range and dominates the formulas by
5x. On real footage it is faster - 153 ms with 32 rules - because neighbouring
cells share palette indices and `stopIfTrue` lets most cells exit on an early
rule; random data hits the worst case.

**32 colours is still the default.** The capture waits for each frame, so rule
count does not affect the reel at all, and it only costs the person scrubbing
the file by hand the difference between 6.5 and 7.4 fps. What it buys is the
picture: at 16 colours a face loses its eyes and mouth to the skin tone.

| palette | mean abs error | |
| --- | --- | --- |
| 16 | 9.30 / 255 | Peter's eyes and mouth wash out |
| 32 | 6.66 / 255 | features survive |

Run it yourself: `./.venv/bin/python bench_cf.py`.

### One palette for the whole clip

A per-frame palette is the obvious implementation and it strobes: the cluster
centres drift a little every frame, so flat regions shimmer even where nothing
moved. `build_palette()` fits k-means once, in Lab, over pixels sampled from
every frame.

`--tau` then adds hysteresis in Lab units - a pixel keeps last frame's index
unless another entry beats it by more than tau. Separating changes that persist
into the next frame (motion) from changes that revert (flicker):

| tau | motion kept | flicker kept |
| --- | --- | --- |
| 0 | 100% | 100% |
| **1 (default)** | **96.8%** | **81%** |
| 2 | 94.4% | 71.5% |
| 4 | 90.8% | 59.4% |

The default sits just under the Lab just-noticeable difference (~2.3), because
the expensive mistake is suppressing real motion, not leaving some sparkle in.

### Held frames depend entirely on the shot

Cartoons animate on twos and threes, so `dedupe()` collapses repeated index
images and several frame numbers point at one `Data` column. How much that saves
is a property of the footage, not the code - measured on the same clip:

| window | changed cells per frame (median of 8,160) | distinct |
| --- | --- | --- |
| 0-12 s, car drives into the lake, camera pans | 1037 | 180/181 |
| 35-46 s, two characters on the shore | 47 | 163/166 |

A panning camera changes every cell, so nothing holds. Do not read a high
distinct count as a bug.

### Capture rate

The 11 s test reel, 166 frames at 15 fps: **655 ms/frame, 2.2 screenshots per
frame, source distinct 98.2% against capture distinct 98.2%, zero missed
frames.** `screencapture` itself costs ~87 ms for the window and ~57 ms for a
small rect, so the poll loop is cheap; what makes the run take two minutes is
waiting for Excel to actually repaint.

## Gotchas

These cost real time:

- **`screencapture` refuses any destination whose filename starts with a dot**
  ("cannot write file to intended destination") **and exits 0 while doing it.**
  A tidy hidden scratch file like `.poll.png` therefore never appears, the
  return code says success, and the whole thing looks exactly like a denied
  Screen Recording permission. Always check the file, never the exit code.
- **Excel addresses workbooks by NAME.** A stale copy left open from an earlier
  run satisfies an "is it already open?" test, so the file just rewritten on
  disk is never read - here that measured the previous geometry, applied a 20%
  correction that was not needed, and captured a clipped grid. `open_book()`
  now closes any same-named workbook before opening.
- **Excel's first document Apple Event after launch takes ~108 seconds.**
  `get version` and `activate` answer instantly the whole time, and the process
  sits at 0% CPU with no dialog on screen, so it looks exactly like a hang. It
  is Excel building its scripting bridge; every command after it takes ~0.1 s.
  `xl.warm_up()` absorbs it deliberately instead of letting each caller guess a
  timeout.
- **Excel windows do not answer `position` or `size`** - both return
  `missing value`. Use `bounds`, which is `{left, top, right, bottom}`. And a
  `zoomed` window silently ignores a new size, so un-zoom it first.
- **Excel rewrites the formulas with `@`** - `=@INDEX(Data!$A$4140:$FG$4140,1,k)`
  - on load. That is the implicit intersection operator, added because the file
  does not declare dynamic-array support. It is harmless: `INDEX` with explicit
  row and column returns a scalar anyway.
- **Screenshotting before the app repaints** is the single most expensive bug in
  the sibling project - it silently produced ~88% duplicate frames. Waiting for
  the window to be "stable" is also wrong, because a window that has not
  repainted yet is perfectly stable and passes instantly. `capture.py` always
  waits for the pixels to CHANGE from a pre-update baseline first, then settle.
- **A 16:9 window needs the grid fitted in BOTH dimensions.** The 9:8 reel
  window is tall enough that width always binds, so the first `plan()` only
  fitted width. Ask for 16:9 and the same zoom runs the grid off the bottom of
  the screen.
- **Pick the band that matches the grid's aspect, not the longest one.** In a
  short 16:9 window the ribbon is a full-width painted band in its own right,
  and once the grid no longer towers over it the ribbon wins - measured as a
  2904x346 box, which the fit step turned into a 1.94x column-width
  "correction". A real squareness error is a few percent, so a +-60% window
  separates chrome from grid without assuming the answer.
- **Refuse a large fit correction outright.** The ECMA width formula measures
  within ~1% every time it has been checked here. Anything bigger means the
  grid was not what got measured, and applying it wrecks geometry that was
  already right. `MAX_FIT` caps it at 20%.
- **The column threshold belongs near the bottom of the gap, not the middle.**
  Empty sheet measures 0.000 - Excel's gridlines are too light to trip the mask
  - while a stretch of picture that is mostly white still measures 0.48. At
  0.50 a white studio background split the grid in two and only the right-hand
  piece was measured.
- **Terminal with no window swallows keystrokes silently.** The launcher drives
  Terminal through System Events; if the last window was closed, `count
  windows` is 0, the keystrokes go nowhere and the run simply never starts.
  Check the count before typing.
- **Do not calibrate on frame 1, and do not detect the picture by saturation
  alone.** Both assume the clip opens on a bright, fully painted frame. Run the
  same pipeline over a whole episode-length clip that opens on a dark car
  interior and the box comes out as a fragment: the fit step then "corrected"
  the column width by 38% and the capture dropped to 96.6% distinct against
  98.9% source, 16 missed frames. A painted cell is *saturated OR dark*, the run
  threshold has to be a majority of the row (gridlines cover ~1/8 of an empty
  one), and the probe frame is chosen as the most strongly painted frame in the
  clip rather than whichever one is first.
- **Do not calibrate the watched region on two adjacent frames.** An earlier
  version diffed frames 1 and 2 and watched whatever moved - but two adjacent
  frames of a cartoon differ in a handful of cells, so it watched that handful
  and every repaint elsewhere was invisible. Measured 73% distinct captured
  against 100% distinct source. Watch the whole painted grid instead.
- **Fingerprint the picture, not the window.** The status bar ticks over during
  a recalculation, which would read as "still changing" forever.
- Frames that are genuinely held never move, so they would burn the whole
  timeout. They are known in advance from the frame map and skip the change
  requirement - but they are still captured for real, never copied, so the
  duplicate measurement stays honest.
- **Measure the capture rate, don't eyeball it.** `capture.measure()` reports the
  fraction of distinct captured frames against the fraction of distinct source
  frames. A slow capture and slow source material look identical on screen.

## Duration

Target **8-15 seconds**, enforced with a warning. Measured across 45 posted
reels in the sibling project: the two best performers both looped (>100%
retention) and both were under 16 s, and retention fits
`1.58 - 0.72*log10(duration_s)` with cross-validated R^2 = 0.47.

## Requirements

```bash
uv venv .venv && uv pip install --python .venv/bin/python \
    openpyxl numpy opencv-python-headless
```

Plus ffmpeg, and Microsoft Excel for macOS. Two macOS permissions, both granted
to **the app that runs this** - Terminal, iTerm, whichever - not to Excel:

- **Automation** -> Microsoft Excel (prompts on first use)
- **Screen Recording** (System Settings > Privacy & Security). The app must be
  quit and reopened after granting it, and the permission only takes effect on
  the next launch.

Run it from a Terminal that holds those permissions.

## Limits

`check_limits()` asserts rather than assuming: `Data` needs `w*h` rows against
Excel's 1,048,576, and one column per distinct frame against 16,384. At 120x68
that is 8,160 rows, so the real ceiling is frames, not cells.
