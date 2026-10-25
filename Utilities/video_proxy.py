"""Create lower-resolution, lower-frame-rate MP4 proxies with FFmpeg.

FFmpeg and ffprobe must be installed and available on PATH.

Python API::

    result = create_proxy("input.mp4", "output/proxy.mp4")

Command line::

    python Utilities/video_proxy.py input.mp4 output/proxy.mp4
"""

from __future__ import annotations

import argparse
import json
import math
import os
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from fractions import Fraction
from pathlib import Path
from typing import Any


class VideoProcessingError(RuntimeError):
    """Raised when a source cannot be probed or a proxy cannot be created."""


@dataclass(frozen=True)
class VideoInfo:
    """Properties reported by ffprobe for a video file."""

    width: int
    height: int
    display_width: int
    display_height: int
    frame_rate: float | None
    duration_seconds: float | None
    has_audio: bool
    video_codec: str | None


@dataclass(frozen=True)
class ProxyResult:
    """Source and output details returned by :func:`create_proxy`."""

    source: VideoInfo
    output: VideoInfo
    output_path: Path
    max_width: int
    max_height: int
    target_fps: float
    include_audio: bool


def _parse_number(value: Any) -> float | None:
    if value in (None, "", "N/A"):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError, ZeroDivisionError):
        try:
            number = float(Fraction(str(value)))
        except (ValueError, ZeroDivisionError):
            return None
    return number if math.isfinite(number) else None


def _rotation_degrees(stream: dict[str, Any]) -> int:
    rotation = _parse_number((stream.get("tags") or {}).get("rotate"))
    for side_data in stream.get("side_data_list") or []:
        candidate = _parse_number(side_data.get("rotation"))
        if candidate is not None:
            rotation = candidate
            break
    return int(round(rotation or 0)) % 360


def _probe(path: Path, ffprobe: str) -> VideoInfo:
    command = [
        ffprobe,
        "-v", "error",
        "-show_entries",
        "stream=codec_type,codec_name,width,height,avg_frame_rate,r_frame_rate,duration:stream_tags=rotate:stream_side_data=rotation:format=duration",
        "-of", "json",
        str(path),
    ]
    try:
        completed = subprocess.run(
            command, check=True, capture_output=True, text=True, encoding="utf-8", errors="replace"
        )
    except FileNotFoundError as exc:
        raise VideoProcessingError(
            "ffprobe was not found. Install FFmpeg (including ffprobe) and add it to PATH."
        ) from exc
    except subprocess.CalledProcessError as exc:
        details = (exc.stderr or "").strip()
        raise VideoProcessingError(
            f"Could not read video metadata for '{path}'.{(' ' + details) if details else ''}"
        ) from exc
    except OSError as exc:
        raise VideoProcessingError(f"Could not run ffprobe: {exc}") from exc

    try:
        metadata = json.loads(completed.stdout)
    except json.JSONDecodeError as exc:
        raise VideoProcessingError(f"ffprobe returned invalid metadata for '{path}'.") from exc
    streams = metadata.get("streams") or []
    video = next((stream for stream in streams if stream.get("codec_type") == "video"), None)
    if video is None:
        raise VideoProcessingError(f"No video stream was found in '{path}'.")
    try:
        width = int(video["width"])
        height = int(video["height"])
    except (KeyError, TypeError, ValueError) as exc:
        raise VideoProcessingError(f"Video dimensions are missing or invalid in '{path}'.") from exc

    rotation = _rotation_degrees(video)
    display_width, display_height = (height, width) if rotation in (90, 270) else (width, height)
    frame_rate = _parse_number(video.get("avg_frame_rate"))
    if not frame_rate:
        frame_rate = _parse_number(video.get("r_frame_rate"))
    duration = _parse_number(video.get("duration"))
    if duration is None:
        duration = _parse_number((metadata.get("format") or {}).get("duration"))

    return VideoInfo(
        width=width,
        height=height,
        display_width=display_width,
        display_height=display_height,
        frame_rate=frame_rate,
        duration_seconds=duration,
        has_audio=any(stream.get("codec_type") == "audio" for stream in streams),
        video_codec=video.get("codec_name"),
    )


def _format_rate(value: float | None) -> str:
    return "unknown" if value is None else f"{value:.3f} fps"


def _format_duration(value: float | None) -> str:
    return "unknown" if value is None else f"{value:.3f} s"


def create_proxy(
    input_path: str | Path,
    output_path: str | Path,
    *,
    max_width: int = 1280,
    max_height: int = 720,
    target_fps: float = 15,
    include_audio: bool = True,
    overwrite: bool = False,
) -> ProxyResult:
    """Create a CFR H.264 MP4 proxy and return source/output properties.

    The video is scaled proportionally to fit within ``max_width`` by
    ``max_height``. Smaller sources are not enlarged. FFmpeg's automatic
    rotation handling is retained, and the ``fps`` filter changes frame
    sampling without changing playback speed.

    Existing outputs are rejected unless ``overwrite=True``. FFmpeg and
    ffprobe must be available on PATH.
    """
    source_path = Path(input_path).expanduser()
    destination = Path(output_path).expanduser()
    if not source_path.is_file():
        raise VideoProcessingError(f"Input file does not exist or is not a file: '{source_path}'.")
    try:
        source_resolved = source_path.resolve(strict=True)
        destination_resolved = destination.resolve(strict=False)
    except OSError as exc:
        raise VideoProcessingError(f"Could not resolve an input or output path: {exc}") from exc
    if source_resolved == destination_resolved:
        raise VideoProcessingError("Input and output paths must refer to different files.")
    if destination.suffix.lower() != ".mp4":
        raise VideoProcessingError("Output path must use the .mp4 extension.")
    if not isinstance(max_width, int) or isinstance(max_width, bool) or max_width < 2:
        raise VideoProcessingError("Maximum width must be an integer of at least 2 pixels.")
    if not isinstance(max_height, int) or isinstance(max_height, bool) or max_height < 2:
        raise VideoProcessingError("Maximum height must be an integer of at least 2 pixels.")
    try:
        target_fps = float(target_fps)
    except (TypeError, ValueError) as exc:
        raise VideoProcessingError("Target frame rate must be a positive number.") from exc
    if not math.isfinite(target_fps) or target_fps <= 0:
        raise VideoProcessingError("Target frame rate must be a positive, finite number.")
    if destination.exists() and not overwrite:
        raise VideoProcessingError(
            f"Output already exists: '{destination}'. Choose another path or set overwrite=True."
        )

    ffmpeg = shutil.which("ffmpeg")
    ffprobe = shutil.which("ffprobe")
    missing = [name for name, path in (("ffmpeg", ffmpeg), ("ffprobe", ffprobe)) if path is None]
    if missing:
        raise VideoProcessingError(
            f"Required executable(s) not found: {', '.join(missing)}. "
            "Install FFmpeg (including ffprobe) and add both tools to PATH."
        )
    assert ffmpeg is not None and ffprobe is not None

    source_info = _probe(source_path, ffprobe)
    try:
        destination.parent.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise VideoProcessingError(f"Could not create output directory '{destination.parent}': {exc}") from exc

    temporary_path: Path | None = None
    try:
        handle = tempfile.NamedTemporaryFile(
            prefix=f".{destination.stem}-", suffix=".tmp.mp4", dir=destination.parent, delete=False
        )
        temporary_path = Path(handle.name)
        handle.close()
        temporary_path.unlink()

        scale_filter = (
            f"scale=w='min({max_width},iw)':h='min({max_height},ih)':"
            "force_original_aspect_ratio=decrease:force_divisible_by=2"
        )
        video_filter = f"{scale_filter},fps=fps={target_fps:.12g}:round=near"
        command = [
            ffmpeg,
            "-hide_banner", "-v", "error", "-nostdin", "-y",
            "-i", str(source_path),
            "-map", "0:v:0",
        ]
        if include_audio:
            command += ["-map", "0:a?"]
        command += [
            "-vf", video_filter,
            "-c:v", "libx264", "-preset", "fast", "-crf", "23",
            "-pix_fmt", "yuv420p", "-fps_mode", "cfr",
            "-map_metadata", "-1", "-metadata:s:v:0", "rotate=0",
        ]
        if include_audio:
            command += ["-c:a", "aac", "-b:a", "128k"]
        else:
            command += ["-an"]
        command += ["-movflags", "+faststart", str(temporary_path)]

        try:
            subprocess.run(
                command, check=True, capture_output=True, text=True, encoding="utf-8", errors="replace"
            )
        except FileNotFoundError as exc:
            raise VideoProcessingError("ffmpeg was not found. Install FFmpeg and add it to PATH.") from exc
        except subprocess.CalledProcessError as exc:
            details = (exc.stderr or "").strip()
            raise VideoProcessingError(
                f"FFmpeg could not create the proxy.{(' ' + details) if details else ''}"
            ) from exc
        except OSError as exc:
            raise VideoProcessingError(f"Could not run ffmpeg: {exc}") from exc

        if not temporary_path.is_file() or temporary_path.stat().st_size == 0:
            raise VideoProcessingError("FFmpeg completed without creating a usable output file.")
        output_info = _probe(temporary_path, ffprobe)
        if output_info.width <= 0 or output_info.height <= 0:
            raise VideoProcessingError("The generated proxy has invalid video dimensions.")
        if include_audio and source_info.has_audio and not output_info.has_audio:
            raise VideoProcessingError("The source has audio, but the generated proxy does not.")

        try:
            if overwrite:
                os.replace(temporary_path, destination)
            else:
                # Same-directory hard links atomically refuse a newly-created destination.
                os.link(temporary_path, destination)
                temporary_path.unlink()
        except FileExistsError as exc:
            raise VideoProcessingError(
                f"Output already exists: '{destination}'. Choose another path or set overwrite=True."
            ) from exc
        except OSError as exc:
            raise VideoProcessingError(f"Could not publish proxy to '{destination}': {exc}") from exc
        temporary_path = None
    finally:
        if temporary_path is not None:
            try:
                temporary_path.unlink(missing_ok=True)
            except OSError:
                pass

    return ProxyResult(
        source=source_info,
        output=output_info,
        output_path=destination,
        max_width=max_width,
        max_height=max_height,
        target_fps=target_fps,
        include_audio=include_audio,
    )


def _print_result(result: ProxyResult) -> None:
    src, out = result.source, result.output
    print(f"Source: {src.display_width}x{src.display_height}, {_format_rate(src.frame_rate)}, "
          f"duration {_format_duration(src.duration_seconds)}, audio {'yes' if src.has_audio else 'no'}")
    print(f"Proxy:  {out.display_width}x{out.display_height}, {_format_rate(out.frame_rate)}, "
          f"duration {_format_duration(out.duration_seconds)}, audio {'yes' if out.has_audio else 'no'}")
    print(f"Settings: max {result.max_width}x{result.max_height}, {result.target_fps:g} fps, "
          f"audio {'included' if result.include_audio else 'removed'}")
    print(f"Output: {result.output_path}")


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Create a lower-resolution, lower-frame-rate H.264 MP4 video proxy."
    )
    parser.add_argument("input", type=Path, help="source video file")
    parser.add_argument("output", type=Path, help="proxy output path (.mp4)")
    parser.add_argument("--max-width", type=int, default=1280, help="maximum output width (default: 1280)")
    parser.add_argument("--max-height", type=int, default=720, help="maximum output height (default: 720)")
    parser.add_argument("--fps", type=float, default=15, help="constant output frame rate (default: 15)")
    parser.add_argument("--no-audio", action="store_true", help="remove audio from the proxy")
    parser.add_argument("--overwrite", action="store_true", help="replace an existing output file")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    try:
        result = create_proxy(
            args.input,
            args.output,
            max_width=args.max_width,
            max_height=args.max_height,
            target_fps=args.fps,
            include_audio=not args.no_audio,
            overwrite=args.overwrite,
        )
    except VideoProcessingError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    _print_result(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
