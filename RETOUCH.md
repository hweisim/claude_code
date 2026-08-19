# retouch_face.py

Natural-looking portrait retouch: makes a pair of unequal eyes match, and takes
the puffiness out of the under-eye area and lower cheeks — without the plastic
look that comes from blurring skin.

## How it works

| Step | What it does |
| --- | --- |
| Landmarks | 478-point face mesh (MediaPipe FaceLandmarker) locates lids, jaw, brows, lips. |
| Eye equalization | Measures each eye's aperture ÷ width, then stretches **only the narrower eye** vertically until it matches. A closed loop re-measures the warped pixels and corrects, so it lands on target at any resolution. By default it closes 75% of the gap — full symmetry is what makes a retouch look retouched. |
| Under-eye puffiness | Frequency separation: the image is split into texture (high frequency) and modelling (low frequency). Only the *modelling* is flattened inside a feathered under-eye mask, so the bag's shadow and bulge fade while pores and lashes stay untouched. Colour channels are flattened at 60% of the lightness amount, which also evens out dark-circle discolouration. |
| Cheeks / jawline | The same tonal flattening at a lower amount over the lower cheeks and jowls, plus a ~2% inward pull on the jaw contour — this is most of what reads as a "puffy" face. |
| Protection mask | Eyes, brows, lips and nostrils are masked out of every tonal pass so they keep their contrast. |

All distances are derived from the landmarks, so the numbers scale with the
photo instead of being hard-coded pixel offsets. The image is resampled exactly
once (Lanczos), and EXIF is carried over.

## Install

```bash
pip install pillow numpy opencv-python-headless mediapipe
```

On a bare Linux box you may also need EGL for MediaPipe:
`apt-get install -y libegl1 libgles2`. The landmark model (~3.7 MB) downloads
on first run to `~/.cache/mediapipe/`.

## Use

```bash
# defaults: subtle, natural
python3 retouch_face.py photo.jpg

# see it next to the original
python3 retouch_face.py photo.jpg -o retouched.jpg --compare before_after.jpg

# dial the whole thing up or down
python3 retouch_face.py photo.jpg --strength 0.6      # gentler
python3 retouch_face.py photo.jpg --strength 1.4      # stronger

# or tune one thing at a time
python3 retouch_face.py photo.jpg --equalize 1.0 --bags 0.7 --cheeks 0.2 --slim 0
```

It prints what it measured and what it did, e.g.

```
  eyes: openness right=0.2304 left=0.2101 (differ by 8.8%)
  opening the left eye: 0.2101 -> 0.2253 target
  before openness  right=0.2304  left=0.2101   difference= 8.8%
  after  openness  right=0.2261  left=0.2324   difference= 2.7%
```

`--help` lists every knob. "Right" and "left" are the subject's own right and
left, not the viewer's.
