# vid2excel

The Excel version of [P.E.T.E.R](https://github.com/earpjonathan/P.E.T.E.R).
Write-up: https://projects.jonathanearp.xyz/desmos/#excel

Turns a short video clip into a colour animation that runs inside an Excel
workbook, then puts it into a 9:16 reel.

```bash
./.venv/bin/python vid2excel.py clip.mp4 -o out/clip --start 35 --duration 11
```

You get three files:

| file | what it is |
| --- | --- |
| `clip.xlsx` | the workbook. Open it, type a number in `A1` and it repaints |
| `clip.reel.mp4` | 1080x1920, real Excel on top, the clip underneath, original audio |
| `clip.preview.mp4` | the colour-reduced grid as a plain video, for checking the palette |

`--no-reel` only writes the workbook. `--limit N` only captures N frames, which
is the quick way to check the framing before a full run.

## How it works

Excel can fill a cell with any colour, so there's no edge detection, tracing or
curve fitting. The whole image pipeline is:

    decode -> autocrop -> downscale to the cell grid -> quantise -> write

The workbook has three sheets:

| sheet | contents |
| --- | --- |
| `Grid` | 120x68 cells, one per pixel, each holding `=INDEX(Data!<its own row>,1,k)` |
| `Data` | one row per grid cell, one column per distinct frame |
| `Map` | one row: output frame `n` -> the `Data` column that draws it |

`A1` holds `n` and `A2` holds `k = INDEX(Map!…,n)`. Type a new `n` and all 8,160
formulas recalculate, each one pulling its palette index out of `Data`.

A formula can't set a fill colour, so the colour comes from conditional
formatting: one `cellIs equal` rule per palette colour across the whole grid,
each with `stopIfTrue`. The numbers in the cells are hidden with the `;;;` number
format, and each rule also sets a matching font colour in case that fails.

## What I measured

All of this was measured on my machine: Excel 16.112.1, M5 Pro, macOS 26.

### Square cells don't match the spec

Excel measures column width in characters of the Normal font's widest digit, and
row height in points, so the two don't share a unit. The documented conversion
(ECMA-376 18.3.1.13, `Truncate((px-5)/MDW*100+0.5)/100` with `MDW=7`) is a
Windows 96-dpi convention, and it's wrong on a Mac. Using it gave 9x12 pt cells,
so the picture came out stretched 33% tall.

So I wrote known widths into a sheet and asked Excel for the rendered size:

| width attribute | 1.0 | 1.5 | 2.0 | 3.0 | 4.0 | 6.0 | 8.0 | 12.0 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| rendered points | 6 | 9 | 12 | 18 | 24 | 36 | 48 | 72 |

That's exactly `rendered = 6 * width`, through zero, while row height in points
renders 1:1. So a square cell of `P` points is `width = P/6`, `height = P`.

Two more problems came up here. Excel's own `Range.Width` says the cell is
exactly as wide as it is tall while it's drawing 10.4 x 12.0 device pixels, so
asking Excel doesn't settle it. Measuring the picture in a screenshot by taking
the extent of every saturated pixel doesn't work either, because the ribbon
icons and the green Share button are saturated too and they stretch the box
across the whole window. The grid is the one large unbroken block, so both
`picture_box()` and `calibrate()` take the longest run of saturated rows, then
the longest run of columns within those rows.

The 6x constant only holds for one Excel build, font and display, so the
pipeline doesn't rely on it. `fit_cells()` captures one frame, measures the
picture Excel actually drew, and rewrites the workbook if the cells are more than
2% off square. On my machine it measures 0.2% and changes nothing.

### Conditional formatting is the slow part

Time for `n` to change and the grid to finish recalculating, 8,160 formula cells,
measured as an Apple Event round trip with the ~108 ms event overhead taken off:

| CF rules | recalc + CF | live fps in Excel |
| --- | --- | --- |
| 0 (formulas only) | 35 ms | 28.9 |
| 16 | 136 ms | 7.4 |
| 32 | 211 ms | 4.7 |

Conditional formatting costs about 5.5 ms per rule over this range, five times
what the formulas cost. Real footage is faster (153 ms with 32 rules) because
neighbouring cells share palette indices and `stopIfTrue` lets most cells stop at
an early rule. Random data is the worst case.

The default is still 32 colours. The capture waits for every frame, so the rule
count doesn't change the reel at all, and someone scrubbing the file by hand only
loses the difference between 6.5 and 7.4 fps. The extra colours make the picture
better. At 16 colours a face loses its eyes and mouth to the skin tone.

| palette | mean abs error | |
| --- | --- | --- |
| 16 | 9.30 / 255 | Peter's eyes and mouth wash out |
| 32 | 6.66 / 255 | you can still see them |

To run it yourself: `./.venv/bin/python bench_cf.py`.

### One palette for the whole clip

A new palette for each frame is the obvious way to do it, and it strobes. The colour
clusters drift a little every frame, so flat areas shimmer even where nothing
moved. `build_palette()` fits k-means once, in Lab colour space, over pixels
sampled from every frame.

`--tau` adds hysteresis in Lab units: a pixel keeps last frame's colour unless
another one beats it by more than tau. That separates changes that stay into the
next frame (motion) from changes that flip back (flicker):

| tau | motion kept | flicker kept |
| --- | --- | --- |
| 0 | 100% | 100% |
| 1 (default) | 96.8% | 81% |
| 2 | 94.4% | 71.5% |
| 4 | 90.8% | 59.4% |

The default sits just under the smallest colour difference people can see in Lab
(about 2.3), because losing real motion is a worse mistake than leaving a little
sparkle in.

### Held frames depend on the shot

Cartoons animate on twos and threes, so `dedupe()` collapses repeated frames and
several frame numbers point at one `Data` column. How much that saves depends on
the footage. On the same clip:

| window | changed cells per frame (median of 8,160) | distinct |
| --- | --- | --- |
| 0-12 s, car drives into the lake, camera pans | 1037 | 180/181 |
| 35-46 s, two characters on the shore | 47 | 163/166 |

A panning camera changes every cell, so no frames get held. A high distinct count
isn't a bug.

### Capture rate

On the 11 s test reel (166 frames at 15 fps) it took 655 ms a frame and 2.2
screenshots per frame. 98.2% of source frames were distinct and 98.2% of captured
frames were distinct, with no missed frames. `screencapture` itself takes about
87 ms for the window and 57 ms for a small rectangle, so polling is cheap. The
run takes two minutes because it's waiting for Excel to repaint.

## Gotchas

These cost me a lot of time.

- `screencapture` won't write to a file whose name starts with a dot ("cannot
  write file to intended destination"), and it exits 0 anyway. So a hidden
  scratch file like `.poll.png` never appears, the return code says it worked,
  and it looks like a missing Screen Recording permission. Check that the file
  exists instead of trusting the exit code.
- Excel finds open workbooks by name. A stale copy left open from an earlier run
  passes an "is it already open?" check, so the file that was just rewritten on
  disk never gets read. When that happened it measured the old geometry, applied
  a 20% correction that wasn't needed, and captured a clipped grid. `open_book()`
  now closes any workbook with the same name before opening.
- Excel's first document Apple Event after launch takes about 108 seconds. The
  whole time, `get version` and `activate` answer instantly and the process sits
  at 0% CPU with no dialog, so it looks like a hang. It's Excel building its
  scripting bridge, and every command after that takes about 0.1 s.
  `xl.warm_up()` waits for it once so nothing else has to guess a timeout.
- Excel windows don't answer `position` or `size`. Both return `missing value`.
  Use `bounds`, which is `{left, top, right, bottom}`. A `zoomed` window also
  ignores a new size without saying so, so un-zoom it first.
- Excel rewrites the formulas with `@` on load, e.g.
  `=@INDEX(Data!$A$4140:$FG$4140,1,k)`. That's the implicit intersection
  operator, added because the file doesn't declare dynamic-array support. It
  doesn't matter here, because `INDEX` with an explicit row and column returns a
  single value anyway.
- Screenshotting before the app repaints was the most expensive bug in P.E.T.E.R.
  It quietly produced about 88% duplicate frames. Waiting for the window to be
  "stable" doesn't fix it, because a window that hasn't repainted yet is
  perfectly stable and passes straight away. `capture.py` waits for the pixels to
  change from a snapshot taken before the update, and then waits for them to
  settle.
- A 16:9 window needs the grid fitted in both directions. The 9:8 reel window is
  tall enough that width is always the limit, so the first `plan()` only fitted
  width. With 16:9, the same zoom runs the grid off the bottom of the screen.
- Pick the band that matches the grid's shape instead of the longest one. In a
  short 16:9 window the ribbon is a full-width painted band too, and once the grid
  isn't much taller than it, the ribbon wins. It measured as a 2904x346 box, and
  the fit step turned that into a 1.94x column-width "correction". A real
  squareness error is a few percent, so a +-60% window separates the ribbon from
  the grid without assuming the answer.
- Refuse large fit corrections. The ECMA width formula has been within about 1%
  every time I've checked it. Anything bigger means something other than the grid
  got measured, and applying it breaks geometry that was already right. `MAX_FIT`
  caps it at 20%.
- Put the column threshold near the bottom of the gap. An empty sheet measures
  0.000 (Excel's gridlines are too light to trip the mask), while a stretch of
  picture that's mostly white still measures 0.48. At 0.50, a white studio
  background split the grid in two and only the right-hand piece was measured.
- Terminal with no window open drops keystrokes. The launcher drives Terminal
  through System Events, and if the last window was closed, `count windows` is 0,
  the keystrokes go nowhere and the run never starts. Check the count before
  typing.
- Don't calibrate on frame 1, and don't find the picture by saturation alone.
  Both assume the clip opens on a bright, fully painted frame. I ran it on an
  episode-length clip that opens on a dark car interior, and the box came out as a
  fragment. The fit step then "corrected" the column width by 38%, and the
  capture dropped to 96.6% distinct against 98.9% in the source, with 16 missed
  frames. Now a painted cell counts if it's saturated or dark, a run has to cover
  most of the row (gridlines cover about 1/8 of an empty row), and the probe frame
  is the most strongly painted frame in the clip.
- Don't pick the watched region from two adjacent frames. An earlier version
  diffed frames 1 and 2 and watched whatever moved, but two adjacent frames of a
  cartoon only differ in a few cells, so it watched those few and missed every
  repaint anywhere else. It captured 73% distinct frames against 100% in the
  source. It watches the whole painted grid now.
- Fingerprint the grid instead of the whole window. The status bar ticks over
  during a recalculation, which would look like it was still changing forever.
- Frames that really are held never change, so waiting for a change would use up
  the whole timeout. The frame map already says which ones they are, so they skip
  the change check. They're still captured for real and never copied, so the
  duplicate count stays accurate.
- Measure the capture rate instead of judging it by eye. `capture.measure()`
  reports the share of distinct captured frames against the share of distinct
  source frames. On screen, a slow capture and a slow-moving clip look the same.

## Duration

Aim for 8-15 seconds. Anything outside that gets a warning. Across 45 posted
reels from P.E.T.E.R, the two that did best both looped (over 100% retention) and
both were under 16 s, and retention fits `1.58 - 0.72*log10(duration_s)` with a
cross-validated R^2 of 0.47.

## Requirements

```bash
uv venv .venv && uv pip install --python .venv/bin/python -r requirements.txt
```

You also need ffmpeg and Microsoft Excel for macOS. It needs two macOS
permissions, and both go to the app you run it from (Terminal, iTerm or
whatever), not to Excel:

- Automation > Microsoft Excel (it asks the first time)
- Screen Recording (System Settings > Privacy & Security). Quit and reopen the
  app after granting it, because it only takes effect on the next launch.

## Limits

`check_limits()` asserts both of these: `Data` needs `w*h` rows,
against Excel's limit of 1,048,576, and one column per distinct frame, against
16,384. At 120x68 that's 8,160 rows, so the number of distinct frames runs out
long before the rows do.
