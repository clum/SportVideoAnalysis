# Project guidance: Step 3 - Detect Additional Key Moments

## Current goal

Build a Python workflow that scans the Step 2 trimmed clip for additional key moments. It takes exactly three user inputs: a video file from `output/clips/`, a sport value from a supported `SportType` enumeration, and an `.xlsx` workbook containing known key-moment times and labels. Return a spreadsheet of additional detected times. Do not classify or assign labels to new detections in this step.

All times in the annotation workbook and output are elapsed times relative to the supplied Step 2 clip, not the original recording or the full Step 1 proxy.

## Deliverables

- A reusable Python utility in `Utilities/` that probes the clip, reads the sport and annotation inputs, finds additional key-moment candidates, and writes a results workbook. Provide a callable function and a command-line interface.
- A `SportType` enum (or reuse an existing project enum) as the single source of supported sport values. Accept only declared enum values; do not infer the sport from the filename or workbook.
- A Jupyter notebook in `Test/` demonstrating the three inputs with the Step 2 output clip. Find the repository root by searching the current directory and its parents.
- If no suitable input workbook exists, add a small key-moment workbook template under `Test/` so the required sheet and columns are clear. Keep it separate from the Step 2 range workbook template because their row meanings differ.
- Keep the function API, CLI flags, workbook examples, and notebook consistent. Keep Python package dependencies in `Utilities/requirements.txt`; keep large model weights outside the repository and document how they are obtained when needed.

## Input workbook contract

- Require a worksheet named `KeyMoments` with columns `Time (seconds)` and `Label`. Each nonblank row is a user-provided known key moment; require a valid time and nonempty label. Ignore fully blank rows and report worksheet row numbers for malformed entries.
- Accept nonnegative decimal seconds, `HH:MM:SS[.mmm]` text, and native Excel duration/time cells. Treat numeric cells as seconds, not Excel date serials. Require literal values rather than formulas.
- Validate every annotated time against the probed clip duration. Reject a missing sheet/header, duplicate or unusable annotations, unsupported time values, or a workbook with no valid examples. Never silently skip a bad row.
- Treat supplied labels as annotations for the user's known examples. They may guide example grouping or candidate similarity, but the detector must not emit predicted labels for new moments.

## Detection expectations

- Use the selected `SportType` to select a sport-aware detection profile while keeping the scanning pipeline reusable across sports. Keep sport-specific cues/configuration isolated and extensible; unsupported sports should fail with a list of supported enum values.
- Use the supplied annotated moments as positive temporal examples and scan the whole clip for additional candidate moments using visual and temporal cues appropriate to the selected sport. Detection is a candidate-finding task: do not add classification or label prediction in this step.
- Do not return the supplied annotations as newly detected results. Suppress duplicate or near-duplicate candidates around known moments and merge overlapping detections into one timestamp. Make the event-window, score/threshold, and temporal suppression settings explicit and reproducible.
- Preserve the original clip and annotation workbook. Do not alter speed, crop, stabilize, overlay, or re-encode the video as part of scanning.
- Use ffprobe for video properties and process frames in timestamp order. Handle decode failures, missing streams, unsupported codecs, and unavailable model/runtime dependencies with actionable errors.
- Avoid hard-coded absolute paths, model locations, or assumptions that every sport uses the same frame rate or visual cues. Report which sport profile and detection settings were used.

## Output workbook contract

- Write a separate `.xlsx` output by default beside the input clip as `<clip-stem>_key_moments.xlsx`; do not overwrite the input workbook or clip.
- Include a `Detected Moments` worksheet containing only additional detections, sorted chronologically. The required field is `Time (seconds)`; include a numeric confidence/score column only when its meaning and range are documented. Do not include a predicted event label.
- Keep times relative to the supplied trimmed clip. Document timestamp precision and any frame-boundary rounding. When no additional moments are found, write a valid workbook with the required headers and zero result rows.
- Reject an existing output unless overwrite is explicit. Write to a temporary file, verify it can be reopened and contains the expected sheet/headers, then publish it. Report the video, selected sport, provided example count, detection count, settings, and output path.

## Example usage and validation

- Document a CLI invocation with the three required inputs, for example: `python Utilities/detect_key_moments.py output/clips/example_selected.mp4 --sport <SUPPORTED_SPORT> --examples Test/key_moments_example.xlsx --output output/clips/example_selected_key_moments.xlsx`.
- Validate a workbook with several known moments, and confirm returned detections are within clip duration, chronologically sorted, and not duplicates of supplied moments or one another.
- Cover an empty detection result, invalid/empty annotations, labels missing, bad timestamps, times beyond duration, malformed workbook, unsupported sport value, unreadable video, missing model/runtime, and existing output behavior.
- Check that changing the sport enum selects the corresponding profile, while output rows remain unlabeled. Verify input video and annotation workbook hashes are unchanged after a run.
- Use the same fixed clip and settings to make a scan reproducible; keep thresholds and relevant run metadata available for debugging without bloating the results sheet.
