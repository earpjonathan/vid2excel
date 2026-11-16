"""decode -> autocrop -> downscale to the cell grid -> quantise to one palette.

That is the whole image pipeline. Excel fills a cell with any colour, so there
is no edge detection, no tracing and no curve fitting here.
"""

from __future__ import annotations

import re
import shutil
import subprocess

import cv2
import numpy as np

VIDEO_EXT = (".mp4", ".mov", ".mkv", ".webm", ".m4v", ".avi")


def ffmpeg() -> str:
    return shutil.which("ffmpeg") or "/opt/homebrew/bin/ffmpeg"


def video_info(path: str) -> tuple[float, float]:
    """(duration seconds, source fps)."""
    cap = cv2.VideoCapture(path)
    fps = cap.get(cv2.CAP_PROP_FPS) or 24.0
    n = cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0
    cap.release()
    if not np.isfinite(fps) or fps <= 0:
        fps = 24.0
    return (float(n) / fps if fps > 0 else 0.0), fps


def detect_crop(path: str, samples: int = 8):
    """Find pillarbox/letterbox bars. Returns (w, h, x, y) or None.

    Lifted from the desmos-video project. Cropping matters more here than it did
    there: a bar does not just waste grid columns, it quantises into a solid
    black band that eats palette entries the picture needed.

    cropdetect reports the bounding box of non-black pixels, so a dark scene
    alone would report a box far smaller than the real picture. Taking the UNION
    over several samples avoids that - bars are constant, dark scenes are not,
    and a union can only ever crop less than the truth.
    """
    cap = cv2.VideoCapture(path)
    W = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
    H = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
    cap.release()
    dur, _ = video_info(path)
    if not (W and H and dur):
        return None

    x0, y0, x1, y1, found = W, H, 0, 0, False
    for i in range(samples):
        t = dur * (i + 0.5) / samples
        r = subprocess.run([ffmpeg(), "-hide_banner", "-ss", f"{t:g}", "-t", "0.8",
                            "-i", path, "-vf", "cropdetect=24:2:0", "-f", "null", "-"],
                           capture_output=True, text=True)
        m = re.findall(r"crop=(\d+):(\d+):(\d+):(\d+)", r.stderr)
        if not m:
            continue
        w, h, x, y = (int(v) for v in m[-1])
        x0, y0 = min(x0, x), min(y0, y)
        x1, y1 = max(x1, x + w), max(y1, y + h)
        found = True
    if not found:
        return None

    w, h = x1 - x0, y1 - y0
    if w >= 0.98 * W and h >= 0.98 * H:
        return None
    if abs((w / h) - 4 / 3) < 0.05:            # snap a near-4:3 result to exact
        w = int(round(h * 4 / 3))
        x0 = max(0, (W - w) // 2)
    w -= w % 2
    h -= h % 2
    return (w, h, x0, y0)


def sample_frames(path: str, fps: float, size: tuple[int, int], start: float,
                  duration: float, crop=None):
    """Yield BGR frames at `fps`, cropped and resized to the cell grid."""
    cap = cv2.VideoCapture(path)
    if not cap.isOpened():
        raise SystemExit(f"could not open video: {path}")
    src_fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
    if not np.isfinite(src_fps) or src_fps <= 0:
        src_fps = 30.0
    if start > 0:
        cap.set(cv2.CAP_PROP_POS_MSEC, start * 1000.0)

    step, next_at, i = src_fps / fps, 0.0, 0
    while True:
        ok, frame = cap.read()
        if not ok:
            break
        if duration > 0 and i / src_fps > duration:
            break
        if i >= next_at:
            next_at += step
            if crop:
                cw, ch, cx, cy = crop
                frame = frame[cy:cy + ch, cx:cx + cw]
            yield cv2.resize(frame, size, interpolation=cv2.INTER_AREA)
        i += 1
    cap.release()


def build_palette(frames: list[np.ndarray], k: int, sample: int = 200_000,
                  seed: int = 0) -> np.ndarray:
    """ONE palette for the whole clip, k-means in Lab. Returns (k,3) BGR uint8.

    A per-frame palette is the obvious implementation and it strobes: the
    cluster centres drift a little every frame, so every flat region shimmers
    even where nothing moved. Fitting once over pixels drawn from the entire
    clip makes a frame's colours depend only on that frame's content.
    """
    lab = np.concatenate([
        cv2.cvtColor(f, cv2.COLOR_BGR2LAB).reshape(-1, 3) for f in frames])
    if len(lab) > sample:
        rng = np.random.default_rng(seed)
        lab = lab[rng.choice(len(lab), sample, replace=False)]
    crit = (cv2.TERM_CRITERIA_EPS + cv2.TERM_CRITERIA_MAX_ITER, 30, 0.5)
    cv2.setRNGSeed(seed)
    _, _, centres = cv2.kmeans(lab.astype(np.float32), k, None, crit, 4,
                               cv2.KMEANS_PP_CENTERS)
    bgr = cv2.cvtColor(np.clip(centres, 0, 255).astype(np.uint8)[None, :, :],
                       cv2.COLOR_LAB2BGR)[0]
    order = np.argsort(centres[:, 0])          # dark -> light, purely cosmetic
    return bgr[order], centres[order]


def quantise(frames: list[np.ndarray], centres: np.ndarray,
             tau: float = 1.0) -> list[np.ndarray]:
    """Map every frame through the shared palette. Returns (h,w) uint8 indices.

    `tau` is a hysteresis in Lab units: a pixel keeps the index it had last
    frame unless some other entry beats it by more than tau. Without it, a pixel
    sitting on the boundary between two palette entries flips back and forth on
    compression noise alone, which reads as sparkle across flat areas - and it
    also splits frames that were otherwise identical, costing Data columns.

    The default is deliberately below the Lab just-noticeable difference (~2.3),
    because the expensive mistake here is suppressing real motion, not leaving
    some sparkle in. Measured on the test clip, separating changes that persist
    into the next frame (motion) from changes that revert (flicker):

        tau      motion kept    flicker kept
        0            100%           100%
        1           96.8%            81%
        2           94.4%          71.5%
        4           90.8%          59.4%

    Every extra unit buys less flicker for more motion, and motion is what the
    animation is made of.
    """
    out, prev = [], None
    for f in frames:
        lab = cv2.cvtColor(f, cv2.COLOR_BGR2LAB).reshape(-1, 3).astype(np.float32)
        d = np.sqrt(((lab[:, None, :] - centres[None, :, :]) ** 2).sum(-1))
        best = d.argmin(1)
        if prev is not None and tau > 0:
            rows = np.arange(len(best))
            keep = d[rows, prev] <= d[rows, best] + tau
            best = np.where(keep, prev, best)
        prev = best
        out.append(best.astype(np.uint8).reshape(f.shape[:2]))
    return out


def dedupe(frames: list[np.ndarray]) -> tuple[list[np.ndarray], list[int]]:
    """Collapse repeated index images. Returns (distinct frames, n -> index).

    Cartoons animate on twos and threes, so adjacent frames are frequently
    identical after quantising. A held frame then costs one number in Map
    instead of 8,160 numbers in Data.
    """
    distinct, order, seen = [], [], {}
    for f in frames:
        key = f.tobytes()
        if key not in seen:
            seen[key] = len(distinct)
            distinct.append(f)
        order.append(seen[key])
    return distinct, order
