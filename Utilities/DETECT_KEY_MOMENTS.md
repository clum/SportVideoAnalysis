# Step 3 detector notes

`detect_key_moments.py` currently uses deterministic optical-flow motion and
frame-difference cues, with separate sampling, event-window, threshold, and
suppression defaults for soccer, basketball, and hockey. The annotation times
calibrate the score threshold and suppress detections near known events. The
detector is a candidate finder: its `Confidence (0-1)` value is a normalized
heuristic score, not a probability or event label. Inspect candidate times
against the video before using them downstream.

This baseline uses OpenCV and has no pretrained model or model weights. If a
future detector adds a trained model, keep its weights outside the repository
and document the official download/release, version, and expected checksum
here. Do not download weights automatically during a scan.

Install Python dependencies with `python -m pip install -r Utilities/requirements.txt`.
Install FFmpeg separately; `ffprobe` must be on `PATH`. Run `python
Utilities/detect_key_moments.py --help` for the current CLI.
