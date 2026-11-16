"""Stack the Excel capture over the original clip into 1080x1920.

build_reel() is taken from the desmos-video project unchanged in shape - the
composite step does not care what produced the top half.
"""

from __future__ import annotations

import os
import shutil
import subprocess


def ffmpeg() -> str:
    return shutil.which("ffmpeg") or "/opt/homebrew/bin/ffmpeg"


def build_reel(frames_dir: str, src: str, start: float, duration: float,
               fps: float, out: str, width: int = 1080, height: int = 1920,
               crf: int = 18, crop=None) -> bool:
    """Excel on top, the clip below on a blurred copy of itself, its own audio."""
    half = height // 2
    seq = os.path.join(frames_dir, "f%06d.png")

    cmd = [ffmpeg(), "-y", "-v", "error", "-framerate", f"{fps:g}", "-i", seq]
    if start > 0:
        cmd += ["-ss", f"{start:g}"]
    cmd += ["-i", src]
    if duration > 0:
        cmd += ["-t", f"{duration:g}"]

    # The capture is taken at the top half's aspect, so it scales in exactly.
    # The clip is 16:9 and its half is not, so it sits on a blurred, zoomed copy
    # of itself rather than dead letterbox bars.
    top = f"scale={width}:{half}:force_original_aspect_ratio=decrease," \
          f"pad={width}:{half}:(ow-iw)/2:(oh-ih)/2:color=0xF2F2F2,setsar=1[top]"
    pre, srcv = "", "[1:v]"
    if crop:
        cw, ch, cx, cy = crop
        pre = f"[1:v]crop={cw}:{ch}:{cx}:{cy},split=2[c1][c2];"
        srcv = "[c1]"
    bot = (f"{srcv}scale={width}:{half}:force_original_aspect_ratio=increase,"
           f"crop={width}:{half},gblur=sigma=24[bg];"
           f"{'[c2]' if crop else '[1:v]'}scale={width}:{half}:"
           f"force_original_aspect_ratio=decrease[fg];"
           f"[bg][fg]overlay=(W-w)/2:(H-h)/2,setsar=1[bot]")
    cmd += ["-filter_complex",
            f"[0:v]{top};{pre}{bot};[top][bot]vstack=inputs=2[v]",
            "-map", "[v]", "-map", "1:a?",
            "-c:v", "libx264", "-preset", "slow", "-crf", str(crf),
            "-pix_fmt", "yuv420p", "-r", f"{fps:g}",
            "-c:a", "aac", "-b:a", "128k", "-shortest", out]
    r = subprocess.run(cmd, capture_output=True)
    if r.returncode != 0:
        print(r.stderr.decode()[:800])
        return False
    return True


def build_flat(frames_dir: str, src: str, start: float, duration: float,
               fps: float, out: str, width: int = 1920, height: int = 1080,
               crf: int = 18) -> bool:
    """Just the Excel window, at 16:9, with the clip's audio.

    The window is sized to the output aspect already, so the pad is normally
    zero; it is Excel's own chrome grey so any residual edge blends in rather
    than reading as a letterbox bar.
    """
    seq = os.path.join(frames_dir, "f%06d.png")
    cmd = [ffmpeg(), "-y", "-v", "error", "-framerate", f"{fps:g}", "-i", seq]
    if start > 0:
        cmd += ["-ss", f"{start:g}"]
    cmd += ["-i", src]
    if duration > 0:
        cmd += ["-t", f"{duration:g}"]
    cmd += ["-filter_complex",
            f"[0:v]scale={width}:{height}:force_original_aspect_ratio=decrease,"
            f"pad={width}:{height}:(ow-iw)/2:(oh-ih)/2:color=0xF2F2F2,setsar=1[v]",
            "-map", "[v]", "-map", "1:a?",
            "-c:v", "libx264", "-preset", "slow", "-crf", str(crf),
            "-pix_fmt", "yuv420p", "-r", f"{fps:g}",
            "-c:a", "aac", "-b:a", "128k", "-shortest", out]
    r = subprocess.run(cmd, capture_output=True)
    if r.returncode != 0:
        print(r.stderr.decode()[:800])
        return False
    return True


def write_preview(path: str, frames, order, palette_bgr, fps: float, scale: int = 8,
                  src: str | None = None, start: float = 0.0,
                  duration: float = 0.0) -> bool:
    """The quantised grid as a plain video - what Excel is being asked to draw.

    Not part of the deliverable; it is how you judge the palette and the grid
    size without waiting on a capture.
    """
    import numpy as np
    h, w = frames[0].shape
    pal = np.asarray(palette_bgr, dtype=np.uint8)
    cmd = [ffmpeg(), "-y", "-v", "error", "-f", "rawvideo", "-pix_fmt", "bgr24",
           "-s", f"{w * scale}x{h * scale}", "-framerate", f"{fps:g}", "-i", "-"]
    if src:
        if start > 0:
            cmd += ["-ss", f"{start:g}"]
        cmd += ["-i", src]
        if duration > 0:
            cmd += ["-t", f"{duration:g}"]
        cmd += ["-map", "0:v", "-map", "1:a?", "-c:a", "aac", "-b:a", "128k",
                "-shortest"]
    cmd += ["-c:v", "libx264", "-preset", "veryfast", "-crf", "18",
            "-pix_fmt", "yuv420p", path]
    p = subprocess.Popen(cmd, stdin=subprocess.PIPE, stderr=subprocess.PIPE)
    for i in order:
        img = pal[frames[i]]
        p.stdin.write(np.repeat(np.repeat(img, scale, 0), scale, 1).tobytes())
    p.stdin.close()
    err = p.stderr.read()
    if p.wait() != 0:
        print(err.decode()[:600])
        return False
    return True
