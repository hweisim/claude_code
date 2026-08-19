#!/usr/bin/env python3
"""Natural portrait retouch: equalize eye openness, reduce puffiness.

Pipeline
--------
1. Detect 478 face landmarks (MediaPipe FaceLandmarker, Tasks API).
2. Measure each eye's aperture/width ratio and warp only the narrower eye, so
   the pair matches. Only a fraction of the gap is closed by default -- perfect
   symmetry is what makes a retouch look retouched.
3. Reduce under-eye puffiness by flattening the low-frequency lightness and
   colour of the eye-bag region while leaving skin texture (high frequency)
   completely untouched.
4. Optionally take a little width off the lower cheeks/jowls and flatten their
   shading, which is most of what reads as a "puffy" face.

Every distance is derived from the landmarks, so the numbers scale with the
photo's resolution and framing instead of being hard-coded pixel offsets.
"""

from __future__ import annotations

import argparse
import os
import sys
import urllib.request
from dataclasses import dataclass

import cv2
import numpy as np
from PIL import Image

import mediapipe as mp
from mediapipe.tasks import python as mp_python
from mediapipe.tasks.python import vision as mp_vision

MODEL_URL = ("https://storage.googleapis.com/mediapipe-models/face_landmarker/"
             "face_landmarker/float16/1/face_landmarker.task")
DEFAULT_MODEL = os.path.join(os.path.expanduser("~"), ".cache", "mediapipe",
                             "face_landmarker.task")

# --- canonical face-mesh indices (MediaPipe "right" == the subject's right) ---
R_UPPER = [246, 161, 160, 159, 158, 157, 173]
R_LOWER = [7, 163, 144, 145, 153, 154, 155]
R_OUTER, R_INNER = 33, 133
L_UPPER = [466, 388, 387, 386, 385, 384, 398]
L_LOWER = [249, 390, 373, 374, 380, 381, 382]
L_OUTER, L_INNER = 263, 362

R_EYE_RING = [33, 7, 163, 144, 145, 153, 154, 155, 133, 173, 157, 158, 159, 160, 161, 246]
L_EYE_RING = [362, 382, 381, 380, 374, 373, 390, 249, 263, 466, 388, 387, 386, 385, 384, 398]
R_BROW = [70, 63, 105, 66, 107, 55, 65, 52, 53, 46]
L_BROW = [300, 293, 334, 296, 336, 285, 295, 282, 283, 276]
LIPS = [61, 146, 91, 181, 84, 17, 314, 405, 321, 375, 291, 409, 270, 269, 267,
        0, 37, 39, 40, 185]
NOSE = [168, 6, 197, 195, 5, 4, 45, 220, 115, 48, 64, 98, 97, 2, 326, 327, 278,
        294, 344, 440, 275]

# jaw/cheek contour, temple -> chin, per side, with a "how much to pull" profile
R_JAW = [234, 93, 132, 58, 172, 136, 150, 149]
L_JAW = [454, 323, 361, 288, 397, 365, 379, 378]
JAW_PROFILE = [0.30, 0.60, 0.85, 1.00, 1.00, 0.85, 0.55, 0.30]

BROW_TOP, CHIN = 10, 152


def unit(v: np.ndarray) -> np.ndarray:
    n = float(np.linalg.norm(v))
    return v / n if n > 1e-9 else np.zeros_like(v)


@dataclass
class Eye:
    name: str
    upper: np.ndarray
    lower: np.ndarray
    outer: np.ndarray
    inner: np.ndarray

    @property
    def width(self) -> float:
        return float(np.linalg.norm(self.outer - self.inner))

    @property
    def center(self) -> np.ndarray:
        return (self.upper.mean(axis=0) + self.lower.mean(axis=0)) / 2.0

    @property
    def axis(self) -> np.ndarray:
        """Unit vector along the eye, inner corner -> outer corner."""
        return unit(self.outer - self.inner)

    @property
    def up(self) -> np.ndarray:
        """Unit vector perpendicular to the eye, pointing at the upper lid."""
        ax = self.axis
        perp = np.array([-ax[1], ax[0]], np.float32)
        if float(np.dot(self.upper.mean(axis=0) - self.center, perp)) < 0:
            perp = -perp
        return perp

    def apex_pair(self) -> tuple[np.ndarray, np.ndarray, float]:
        """Widest-open upper/lower lid pair, and that opening in pixels."""
        up = self.up
        gaps = [float(np.dot(u - l, up)) for u, l in zip(self.upper, self.lower)]
        i = int(np.argmax(gaps))
        return self.upper[i], self.lower[i], gaps[i]

    @property
    def aperture(self) -> float:
        return self.apex_pair()[2]

    @property
    def openness(self) -> float:
        """Aperture as a fraction of eye width -- comparable between eyes."""
        return self.aperture / max(self.width, 1e-6)


class Warp:
    """Accumulates a smooth displacement field, then resamples the image once."""

    def __init__(self, shape: tuple[int, int]):
        h, w = shape
        self.h, self.w = h, w
        self.dx = np.zeros((h, w), np.float32)
        self.dy = np.zeros((h, w), np.float32)

    def _box(self, center: np.ndarray, radius: float):
        cx, cy = float(center[0]), float(center[1])
        x0 = max(0, int(cx - radius)); x1 = min(self.w, int(cx + radius) + 1)
        y0 = max(0, int(cy - radius)); y1 = min(self.h, int(cy + radius) + 1)
        if x1 <= x0 or y1 <= y0:
            return None
        ys, xs = np.mgrid[y0:y1, x0:x1]
        return x0, x1, y0, y1, xs.astype(np.float32) - cx, ys.astype(np.float32) - cy

    def translate(self, center, delta, sigma) -> None:
        """Move content near `center` by `delta`, feathered over `sigma`."""
        box = self._box(center, 3.0 * sigma)
        if box is None:
            return
        x0, x1, y0, y1, u, v = box
        wt = np.exp(-(u * u + v * v) / (2.0 * sigma * sigma))
        self.dx[y0:y1, x0:x1] += float(delta[0]) * wt
        self.dy[y0:y1, x0:x1] += float(delta[1]) * wt

    def scale(self, center, axis, up, gain_axis, gain_up, sig_axis, sig_up) -> None:
        """Stretch content around `center` in the eye's own frame of reference."""
        radius = 3.0 * max(sig_axis, sig_up)
        box = self._box(center, radius)
        if box is None:
            return
        x0, x1, y0, y1, dxg, dyg = box
        a = dxg * axis[0] + dyg * axis[1]          # along the eye
        b = dxg * up[0] + dyg * up[1]              # across the eye
        wt = np.exp(-0.5 * ((a / sig_axis) ** 2 + (b / sig_up) ** 2))
        mag_a = gain_axis * a * wt
        mag_b = gain_up * b * wt
        self.dx[y0:y1, x0:x1] += mag_a * axis[0] + mag_b * up[0]
        self.dy[y0:y1, x0:x1] += mag_a * axis[1] + mag_b * up[1]

    def weight_at(self, center, axis, up, sig_axis, sig_up, point) -> float:
        d = np.asarray(point, np.float32) - np.asarray(center, np.float32)
        a = float(np.dot(d, axis)); b = float(np.dot(d, up))
        return float(np.exp(-0.5 * ((a / sig_axis) ** 2 + (b / sig_up) ** 2)))

    def apply(self, img: np.ndarray) -> np.ndarray:
        ys, xs = np.mgrid[0:self.h, 0:self.w].astype(np.float32)
        # inverse mapping: content at q lands at q + d(q)
        map_x = xs - self.dx
        map_y = ys - self.dy
        return cv2.remap(img, map_x, map_y, cv2.INTER_LANCZOS4,
                         borderMode=cv2.BORDER_REFLECT_101)


def load_landmarker(model_path: str):
    if not os.path.exists(model_path):
        os.makedirs(os.path.dirname(model_path) or ".", exist_ok=True)
        sys.stderr.write("downloading face landmark model...\n")
        urllib.request.urlretrieve(MODEL_URL, model_path)
    opts = mp_vision.FaceLandmarkerOptions(
        base_options=mp_python.BaseOptions(model_asset_path=model_path),
        num_faces=1,
        min_face_detection_confidence=0.3,
    )
    return mp_vision.FaceLandmarker.create_from_options(opts)


def landmarks(landmarker, bgr: np.ndarray) -> np.ndarray:
    h, w = bgr.shape[:2]
    rgb = cv2.cvtColor(bgr, cv2.COLOR_BGR2RGB)
    result = landmarker.detect(mp.Image(image_format=mp.ImageFormat.SRGB, data=rgb))
    if not result.face_landmarks:
        raise SystemExit("no face found in the image")
    return np.array([[lm.x * w, lm.y * h] for lm in result.face_landmarks[0]], np.float32)


def eyes_of(pts: np.ndarray) -> tuple[Eye, Eye]:
    right = Eye("right", pts[R_UPPER], pts[R_LOWER], pts[R_OUTER], pts[R_INNER])
    left = Eye("left", pts[L_UPPER], pts[L_LOWER], pts[L_OUTER], pts[L_INNER])
    return right, left


# --------------------------------------------------------------------------- #
# geometry: eye equalization + face slimming
# --------------------------------------------------------------------------- #

def eye_scale_params(eye: Eye) -> tuple[float, float]:
    """Feathering of the eye warp: wide and shallow, like the eye itself."""
    return 0.55 * eye.width, 0.32 * eye.width


def add_eye_stretch(warp: Warp, eye: Eye, gain: float) -> None:
    sig_axis, sig_up = eye_scale_params(eye)
    warp.scale(eye.center, eye.axis, eye.up, gain_axis=0.0, gain_up=gain,
               sig_axis=sig_axis, sig_up=sig_up)


def first_guess_gain(warp: Warp, eye: Eye, target_aperture: float) -> float:
    """Analytic gain from where the warp will put the lid landmarks."""
    sig_axis, sig_up = eye_scale_params(eye)
    c, ax, up = eye.center, eye.axis, eye.up
    p_up, p_lo, aperture = eye.apex_pair()
    v_up = float(np.dot(p_up - c, up))
    v_lo = float(np.dot(p_lo - c, up))
    w_up = warp.weight_at(c, ax, up, sig_axis, sig_up, p_up)
    w_lo = warp.weight_at(c, ax, up, sig_axis, sig_up, p_lo)
    denom = v_up * w_up - v_lo * w_lo            # aperture(g) = aperture + g*denom
    return 0.0 if abs(denom) < 1e-6 else (target_aperture - aperture) / denom


def solve_eye_gain(bgr: np.ndarray, landmarker, pts: np.ndarray, fraction: float,
                   max_gain: float, verbose: bool) -> tuple[str, float]:
    """Find the stretch that actually lands the narrower eye on its target.

    Resampling softens lid edges, so the analytic gain undershoots what the
    detector measures back. Re-measure on the warped pixels and correct, a
    couple of rounds, which is cheap and self-correcting at any resolution.
    """
    right, left = eyes_of(pts)
    narrow, wide = (right, left) if right.openness < left.openness else (left, right)
    start = narrow.openness
    target = start + (wide.openness - start) * fraction
    if verbose:
        print(f"  eyes: openness right={right.openness:.4f} left={left.openness:.4f} "
              f"(differ by {abs(right.openness - left.openness) / wide.openness:.1%})")
        print(f"  opening the {narrow.name} eye: {start:.4f} -> {target:.4f} target")
    if target - start < 1e-4 or max_gain <= 0:
        return narrow.name, 0.0

    probe = Warp(bgr.shape[:2])
    gain = float(np.clip(first_guess_gain(probe, narrow, target * narrow.width),
                         0.0, max_gain))
    best_gain, best_err = gain, float("inf")
    for _ in range(4):
        if gain <= 1e-4:
            break
        trial = Warp(bgr.shape[:2])
        add_eye_stretch(trial, narrow, gain)
        got = pick_eye(eyes_of(landmarks(landmarker, trial.apply(bgr))), narrow.name)
        err = abs(got.openness - target)
        if err < best_err:
            best_gain, best_err = gain, err
        if verbose:
            print(f"    gain {gain:.3f} -> openness {got.openness:.4f}")
        if err <= 0.01 * target:
            break
        slope = (got.openness - start) / gain           # openness is ~linear in gain
        if slope <= 1e-6:
            break
        nxt = float(np.clip((target - start) / slope, 0.0, max_gain))
        if abs(nxt - gain) < 1e-4:
            break
        gain = nxt
    return narrow.name, best_gain


def pick_eye(pair: tuple[Eye, Eye], name: str) -> Eye:
    right, left = pair
    return right if name == "right" else left


def slim_face(warp: Warp, pts: np.ndarray, amount: float, verbose: bool) -> None:
    if amount <= 0:
        return
    face_w = float(np.linalg.norm(pts[454] - pts[234]))
    top, chin = pts[BROW_TOP], pts[CHIN]
    axis = unit(chin - top)
    sigma = 0.11 * face_w
    peak = amount * face_w / 2.0

    for side in (R_JAW, L_JAW):
        for idx, weight in zip(side, JAW_PROFILE):
            p = pts[idx]
            on_axis = top + axis * float(np.dot(p - top, axis))
            warp.translate(p, unit(on_axis - p) * peak * weight, sigma)
    if verbose:
        print(f"  jawline: pulling in up to {peak:.1f}px "
              f"({amount:.1%} of half face width {face_w / 2:.0f}px)")


# --------------------------------------------------------------------------- #
# tone: flatten the swelling without touching skin texture
# --------------------------------------------------------------------------- #

def polygon_mask(shape, polygons, feather: float) -> np.ndarray:
    m = np.zeros(shape[:2], np.float32)
    for poly in polygons:
        cv2.fillPoly(m, [np.round(poly).astype(np.int32)], 1.0)
    if feather > 0:
        m = cv2.GaussianBlur(m, (0, 0), feather)
    return m


def under_eye_mask(shape, eye: Eye, depth: float, feather: float) -> np.ndarray:
    """Band hugging the lower lid and extending onto the cheek."""
    down = -eye.up
    lid = np.vstack([eye.outer[None, :], eye.lower, eye.inner[None, :]])
    outward = np.vstack([eye.outer[None, :] + down * depth * 0.45,
                         eye.lower + down * depth,
                         eye.inner[None, :] + down * depth * 0.55])
    poly = np.vstack([lid, outward[::-1]])
    return polygon_mask(shape, [poly], feather)


def cheek_mask(shape, pts: np.ndarray, feather: float) -> np.ndarray:
    face_w = float(np.linalg.norm(pts[454] - pts[234]))
    top, axis = pts[BROW_TOP], unit(pts[CHIN] - pts[BROW_TOP])
    polys = []
    for side in (R_JAW, L_JAW):
        outer = np.array([pts[i] for i in side], np.float32)
        inner = []
        for p in outer:
            on_axis = top + axis * float(np.dot(p - top, axis))
            inner.append(p + unit(on_axis - p) * 0.42 * face_w)
        polys.append(np.vstack([outer, np.array(inner, np.float32)[::-1]]))
    return polygon_mask(shape, polys, feather)


def protect_mask(shape, pts: np.ndarray, eye_w: float) -> np.ndarray:
    """Features that must keep their contrast: eyes, brows, lips, nostrils."""
    polys = [np.array([pts[i] for i in ring], np.float32)
             for ring in (R_EYE_RING, L_EYE_RING, R_BROW, L_BROW, LIPS, NOSE)]
    m = polygon_mask(shape, polys, 0.0)
    k = max(3, int(0.10 * eye_w) | 1)
    m = cv2.dilate(m, cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (k, k)))
    return cv2.GaussianBlur(m, (0, 0), 0.10 * eye_w)


def flatten(bgr: np.ndarray, mask: np.ndarray, alpha: float,
            detail_sigma: float, broad_sigma: float) -> np.ndarray:
    """Pull local lightness/colour toward the surrounding average.

    Splits the image into texture (detail) and modelling (base), softens only
    the base inside `mask`, and puts the original texture back on top -- so
    swelling and shadow fade while pores and lashes survive intact.
    """
    if alpha <= 0 or float(mask.max()) <= 0:
        return bgr
    lab = cv2.cvtColor(bgr, cv2.COLOR_BGR2LAB).astype(np.float32)
    out = lab.copy()
    for ch, gain in ((0, 1.0), (1, 0.6), (2, 0.6)):
        plane = lab[:, :, ch]
        base = cv2.GaussianBlur(plane, (0, 0), detail_sigma)
        detail = plane - base
        broad = cv2.GaussianBlur(base, (0, 0), broad_sigma)
        base = base + (broad - base) * (alpha * gain * mask)
        out[:, :, ch] = base + detail
    out[:, :, 0] = np.clip(out[:, :, 0], 0, 255)
    out[:, :, 1:] = np.clip(out[:, :, 1:], 0, 255)
    return cv2.cvtColor(out.astype(np.uint8), cv2.COLOR_LAB2BGR)


# --------------------------------------------------------------------------- #

def retouch(bgr: np.ndarray, landmarker, args) -> np.ndarray:
    s = args.strength
    pts = landmarks(landmarker, bgr)
    right, left = eyes_of(pts)
    eye_w = (right.width + left.width) / 2.0

    if args.verbose:
        print("geometry:")

    warp = Warp(bgr.shape[:2])
    if args.equalize > 0:
        name, gain = solve_eye_gain(bgr, landmarker, pts,
                                    min(args.equalize * s, 1.0),
                                    args.max_eye_gain, args.verbose)
        if gain > 0:
            add_eye_stretch(warp, pick_eye((right, left), name), gain)
    if args.puff_lift > 0:
        # centred well clear of the lid, so the eye can never be pinched shut
        for eye in (right, left):
            lift = args.puff_lift * s * eye.width
            for p in eye.lower:
                warp.translate(p - eye.up * 1.10 * eye.width,
                               eye.up * lift, 0.35 * eye.width)
        if args.verbose:
            print(f"  under-eye: easing the fullness up by "
                  f"{args.puff_lift * s * eye_w:.1f}px")
    slim_face(warp, pts, args.slim * s, args.verbose)
    out = warp.apply(bgr)

    # landmarks shifted with the warp, so re-detect before the tonal pass
    pts = landmarks(landmarker, out)
    right, left = eyes_of(pts)
    keep = protect_mask(out.shape, pts, eye_w)

    if args.verbose:
        print("tone:")
    if args.bags > 0:
        m = np.maximum(under_eye_mask(out.shape, right, 0.85 * right.width, 0.22 * eye_w),
                       under_eye_mask(out.shape, left, 0.85 * left.width, 0.22 * eye_w))
        m = np.clip(m * (1.0 - keep), 0, 1)
        out = flatten(out, m, args.bags * s, 0.030 * eye_w, 0.50 * eye_w)
        if args.verbose:
            print(f"  under-eye bags: flattening at {args.bags * s:.2f}")
    if args.cheeks > 0:
        m = np.clip(cheek_mask(out.shape, pts, 0.5 * eye_w) * (1.0 - keep), 0, 1)
        out = flatten(out, m, args.cheeks * s, 0.045 * eye_w, 0.90 * eye_w)
        if args.verbose:
            print(f"  cheeks/jowls: flattening at {args.cheeks * s:.2f}")
    return out


def report(landmarker, before: np.ndarray, after: np.ndarray) -> None:
    for label, img in (("before", before), ("after", after)):
        r, l = eyes_of(landmarks(landmarker, img))
        gap = abs(r.openness - l.openness) / max(r.openness, l.openness)
        print(f"  {label:6s} openness  right={r.openness:.4f}  left={l.openness:.4f}"
              f"   difference={gap:5.1%}")


def read_image(path: str) -> tuple[np.ndarray, Image.Image]:
    pil = Image.open(path)
    exif = pil.getexif()
    if exif.get(274, 1) != 1:                       # bake in EXIF orientation
        pil = Image.open(path)
        from PIL import ImageOps
        pil = ImageOps.exif_transpose(pil)
    rgb = np.array(pil.convert("RGB"))
    return cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR), Image.open(path)


def write_image(path: str, bgr: np.ndarray, source: Image.Image, quality: int) -> None:
    rgb = cv2.cvtColor(bgr, cv2.COLOR_BGR2RGB)
    out = Image.fromarray(rgb)
    exif = source.getexif()
    exif.pop(274, None)                             # orientation is baked in now
    params = {}
    if path.lower().endswith((".jpg", ".jpeg")):
        params = {"quality": quality, "subsampling": 0}
    out.save(path, exif=exif.tobytes() if len(exif) else None, **params)


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("image")
    p.add_argument("-o", "--out", help="default: <name>_retouched.jpg")
    p.add_argument("--compare", help="also write a before/after side-by-side here")
    p.add_argument("--strength", type=float, default=1.0,
                   help="global multiplier on every effect (default 1.0)")
    p.add_argument("--equalize", type=float, default=0.75,
                   help="fraction of the eye-openness gap to close (default 0.75)")
    p.add_argument("--max-eye-gain", type=float, default=0.30,
                   help="hard cap on the eye stretch (default 0.30)")
    p.add_argument("--bags", type=float, default=0.55,
                   help="under-eye puffiness flattening, 0-1 (default 0.55)")
    p.add_argument("--puff-lift", type=float, default=0.0,
                   help="optional geometric lift of the under-eye fullness, in "
                        "eye widths; the tonal pass usually suffices (default 0)")
    p.add_argument("--cheeks", type=float, default=0.35,
                   help="cheek/jowl flattening, 0-1 (default 0.35)")
    p.add_argument("--slim", type=float, default=0.022,
                   help="jawline pull-in, in half face widths (default 0.022)")
    p.add_argument("--model", default=DEFAULT_MODEL)
    p.add_argument("--quality", type=int, default=96)
    p.add_argument("-q", "--quiet", dest="verbose", action="store_false")
    args = p.parse_args()

    out_path = args.out or f"{os.path.splitext(args.image)[0]}_retouched.jpg"
    landmarker = load_landmarker(args.model)
    bgr, source = read_image(args.image)
    result = retouch(bgr, landmarker, args)

    write_image(out_path, result, source, args.quality)
    if args.verbose:
        print("result:")
        report(landmarker, bgr, result)
        print(f"wrote {out_path}")

    if args.compare:
        pair = np.hstack([bgr, result])
        cv2.imwrite(args.compare, pair,
                    [cv2.IMWRITE_JPEG_QUALITY, args.quality])
        if args.verbose:
            print(f"wrote {args.compare}")


if __name__ == "__main__":
    main()
