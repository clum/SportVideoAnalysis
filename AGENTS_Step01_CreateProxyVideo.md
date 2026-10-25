# Project guidance:  Step 1 - Create Proxy Video File

## Current video-processing goal

Build a Python command-line tool that takes a high-resolution source video (for example, MP4) and creates a smaller, lower-frame-rate proxy video for downstream sports-video analysis. Keep the original file unchanged. The proxy should remain temporally useful for analysis and have predictable, documented properties.

## Deliverables

- A primary Python script that performs the video-processing task: probe the source video, create a lower-resolution and lower-frame-rate proxy, and report the resulting properties. Structure the conversion logic so it can be called as a reusable function as well as from the command line.
- An example Python script in a Jupyter notebook that demonstrates how to call the primary function or invoke the primary script, including the input path, output path, and any commonly configured resolution or frame-rate options. Make it clear which paths are placeholders and how to run the example.
- Keep the example aligned with the primary script's actual interface so it can be adapted or run without guessing argument names or defaults.

## Implementation expectations

- Accept an input video path and an output path, with configurable target resolution and frame rate. Choose sensible defaults, but make them visible in `--help` and record the effective settings in the run output.
- Use FFmpeg/ffprobe through subprocesses for probing and transcoding unless the project establishes a different video backend. Check that required executables are available, pass arguments as a list (not through a shell), and report actionable errors including relevant FFmpeg output.
- Probe the source before conversion. Capture at least its dimensions, nominal/average frame rate when available, duration, and whether it has audio. Handle missing or unusual metadata without crashing where practical.
- Scale proportionally without stretching. Never enlarge the source to meet a target size. Preserve display orientation and account for rotation metadata. Ensure encoded dimensions satisfy the selected codec's constraints (for example, even width and height for common H.264 encoders).
- Make frame-rate conversion explicit and deterministic. Support lowering the frame rate while preserving the clip's timeline and duration; do not speed up or slow down the action by changing timestamps. Document how variable-frame-rate inputs are handled.
- Use a broadly compatible proxy format and codec (MP4 with H.264 is a reasonable default). Decide and document audio behavior; make audio optional if it is not needed for analysis. Keep encoding settings configurable when downstream requirements justify it.
- Validate input paths and supported media, create output directories when appropriate, and avoid silently overwriting an existing output. Write to a temporary file and rename it into place only after FFmpeg succeeds and the result can be probed, so interrupted runs do not leave a seemingly complete proxy.
- Print a concise success summary with source and output properties, including actual resolution, frame rate, duration, and output path. Return a nonzero exit status on failure.
- Keep the script usable on Windows and other supported platforms. Use `pathlib` for paths, avoid hard-coded machine-specific locations, and include a short usage example in the script documentation or project README.

## Downstream-analysis considerations

- A proxy is for analysis, so preserve temporal ordering and enough spatial detail to identify the actions or objects the analysis needs. Do not assume a particular resolution or frame rate is correct for every sport or analysis task; expose these as parameters.
- Avoid adding overlays, cropping, stabilization, or other visual changes unless specifically requested. Such transformations can alter measurements or interfere with computer-vision workflows.
- Keep conversion logs and errors clear enough to reproduce a run. When practical, report both requested and actual output properties rather than assuming the encoder produced exactly the requested values.
