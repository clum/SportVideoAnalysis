# Project guidance: Step 2 - Trim a Proxy Video by Spreadsheet Time Ranges

## Current goal

Build a Python workflow that takes a proxy video created in Step 1 and an `.xlsx` workbook containing one or more start/stop time ranges. Keep only those ranges, in chronological order, and join them into one shorter output video. Do not modify the source proxy.

Treat the proxy timeline as the reference for all times. The spreadsheet values are elapsed times from the beginning of that proxy, not time-of-day values.

## Deliverables

- A reusable Python utility in `Utilities/` that reads and validates the range workbook, trims the requested portions of the proxy, concatenates them into one output, and reports the selected ranges and output properties. Provide both a callable function and a command-line interface.
- A Jupyter notebook in `Test/` demonstrating the utility with the Step 1 proxy and an example `.xlsx` range workbook. Resolve the repository root by searching from the current directory and its parents; do not assume the notebook server was launched from the repository root.
- If no suitable workbook example exists, include a small example/template workbook under `Test/` so the required columns and time entry formats are clear.
- Keep the notebook, workbook example, CLI flags, and Python function interface consistent.

## Workbook contract

- Read `.xlsx` workbooks with a Python library suited to reading Excel files (for example, `openpyxl`). Keep Python package requirements in `Utilities/requirements.txt`; FFmpeg and ffprobe remain separate system prerequisites.
- Use a worksheet named `Segments` with required columns `Start Time (seconds)` and `Stop Time (seconds)`. An optional `Label` column may identify a segment for reporting. Ignore completely blank rows; reject partially filled rows with the worksheet row number in the error.
- Accept nonnegative decimal seconds and elapsed-time text in `HH:MM:SS` or `HH:MM:SS.mmm` form. Also accept native Excel duration/time cells when they represent elapsed time. Interpret numeric cells as seconds, not Excel date serials.
- Require literal input values rather than formulas, so results do not depend on stale formula caches. Reject missing sheets or headers, invalid time values, and workbooks with no valid segment rows.
- Spreadsheet row order defines output order and must be chronological. Require each start to be earlier than its stop, each stop to be within the probed proxy duration, and each range to start at or after the prior range ends. Adjacent segments are allowed; overlapping or out-of-order ranges are errors rather than silently sorted or merged.
- Treat each interval as start-inclusive and stop-exclusive. Convert boundaries to the nearest valid frame boundary for the proxy frame rate, report the effective boundaries, and document the rounding behavior.

## Video-processing expectations

- Probe the proxy with ffprobe before processing. Require a readable video stream and a known, usable duration. Detect audio and preserve it by default when present, with an option to remove it.
- Use FFmpeg/ffprobe via `subprocess` with argument lists (never shell interpolation). Use an accurate trim/re-encode and concat workflow so segment boundaries do not depend on keyframe locations. Keep playback speed, frame rate, orientation, and audio/video synchronization intact. The output should retain the proxy's display dimensions and frame rate unless an explicit option says otherwise.
- Validate that input video, workbook, and output paths are distinct and usable. Create the output directory when appropriate, reject an existing output unless overwrite is explicitly requested, and write to a temporary file before publishing only a successful, probeable result.
- Report the input proxy, workbook, each requested/effective segment range, total kept duration, output dimensions, frame rate, audio presence, and output path. Return a nonzero process exit code for failures and include actionable diagnostics.
- Do not add overlays, crop, stabilize, or otherwise alter the selected frames. Do not silently skip invalid spreadsheet rows or change their order.

## Example usage and validation

- Document an example CLI invocation and function call in the utility or notebook. Make clear that times refer to the proxy, show the workbook column names, and identify example paths that users should replace.
- Validate a workbook with one range and another with several ranges; verify the output duration is approximately the sum of the kept intervals, adjusted only for documented frame-boundary rounding.
- Check audio presence and synchronization when the proxy has audio; also cover a silent proxy and the option to drop audio.
- Check malformed workbooks, missing headers/sheet, blank rows, invalid or negative times, zero-length ranges, ranges beyond duration, overlap/out-of-order rows, and existing output behavior. Confirm failures leave no partial final output and never modify the input proxy or workbook.
- Confirm the notebook's repository-root discovery works when launched from `Test/` and from the repository root.
