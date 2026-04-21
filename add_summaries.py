#!/usr/bin/env python3

import argparse
import json
import re
from pathlib import Path


YOUTUBE_PATTERNS = [
    re.compile(r"^https?://(?:www\.)?youtube\.com/watch\?v=([A-Za-z0-9_-]{11})(?:[&?].*)?$"),
    re.compile(r"^https?://youtu\.be/([A-Za-z0-9_-]{11})(?:[?].*)?$"),
]

START_MARKER_TEMPLATE = "<!-- video-summary:start {video_id} -->"
END_MARKER_TEMPLATE = "<!-- video-summary:end {video_id} -->"


def extract_video_id(url: str) -> str | None:
    url = url.strip()
    for pattern in YOUTUBE_PATTERNS:
        m = pattern.match(url)
        if m:
            return m.group(1)
    return None


def build_video_to_stem_map(transcripts_dir: Path) -> dict[str, str]:
    """
    Map YouTube video_id -> file stem by reading transcript JSON files.
    Example:
      transcripts/0001.txt -> {"video_id":"LPZh9BOjkQs", ...}
      returns {"LPZh9BOjkQs": "0001"}
    """
    mapping: dict[str, str] = {}

    for path in sorted(transcripts_dir.iterdir()):
        if not path.is_file():
            continue

        try:
            text = path.read_text(encoding="utf-8").strip()
            payload = json.loads(text)
            video_id = payload.get("video_id")
            if isinstance(video_id, str) and video_id:
                mapping[video_id] = path.stem
        except Exception as e:
            print(f"Warning: could not parse transcript file {path}: {e}")

    return mapping


def find_summary_file(summaries_dir: Path, stem: str) -> Path | None:
    """
    Prefer:
      <stem>.simple_summary.md
    Fallback:
      any markdown file starting with '<stem>.'
      or exactly '<stem>.md'
    """
    preferred = summaries_dir / f"{stem}.simple_summary.md"
    if preferred.exists():
        return preferred

    exact = summaries_dir / f"{stem}.md"
    if exact.exists():
        return exact

    candidates = sorted(summaries_dir.glob(f"{stem}*.md"))
    return candidates[0] if candidates else None


def strip_existing_inserted_block(lines: list[str], start_index: int, video_id: str) -> int:
    """
    If a previously inserted summary block exists immediately after the URL,
    return the index just after that block. Otherwise return start_index.
    """
    start_marker = START_MARKER_TEMPLATE.format(video_id=video_id)
    end_marker = END_MARKER_TEMPLATE.format(video_id=video_id)

    i = start_index

    # Skip blank lines after the URL
    while i < len(lines) and lines[i].strip() == "":
        i += 1

    if i < len(lines) and lines[i].strip() == start_marker:
        i += 1
        while i < len(lines) and lines[i].strip() != end_marker:
            i += 1
        if i < len(lines) and lines[i].strip() == end_marker:
            i += 1

        # Also consume trailing blank lines after the marker block
        while i < len(lines) and lines[i].strip() == "":
            i += 1

        return i

    return start_index


def insert_summaries_into_markdown(
    markdown_path: Path,
    transcripts_dir: Path,
    summaries_dir: Path,
) -> str:
    video_to_stem = build_video_to_stem_map(transcripts_dir)
    original = markdown_path.read_text(encoding="utf-8")
    lines = original.splitlines(keepends=True)

    output: list[str] = []
    i = 0

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        video_id = extract_video_id(stripped)

        output.append(line)
        i += 1

        if not video_id:
            continue

        stem = video_to_stem.get(video_id)
        if not stem:
            print(f"Warning: no transcript file found for video_id={video_id}")
            continue

        summary_file = find_summary_file(summaries_dir, stem)
        if not summary_file:
            print(f"Warning: no summary file found for stem={stem} video_id={video_id}")
            continue

        # Remove any previously inserted block directly after the URL
        new_i = strip_existing_inserted_block(lines, i, video_id)
        if new_i != i:
            i = new_i

        summary_text = summary_file.read_text(encoding="utf-8").rstrip()

        start_marker = START_MARKER_TEMPLATE.format(video_id=video_id)
        end_marker = END_MARKER_TEMPLATE.format(video_id=video_id)

        block = (
            "\n"
            f"{start_marker}\n\n"
            f"{summary_text}\n\n"
            f"{end_marker}\n\n"
        )
        output.append(block)

    return "".join(output)


def main():
    parser = argparse.ArgumentParser(
        description="Insert video summaries below YouTube URLs in a markdown notes file."
    )
    parser.add_argument("markdown_file", help="Path to the markdown notes file")
    parser.add_argument(
        "--transcripts-dir",
        default="./transcripts",
        help="Directory containing transcript JSON files",
    )
    parser.add_argument(
        "--summaries-dir",
        default="./summaries",
        help="Directory containing summary markdown files",
    )
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="Overwrite the original markdown file",
    )
    parser.add_argument(
        "--output",
        help="Write output to a different file instead of stdout",
    )

    args = parser.parse_args()

    markdown_path = Path(args.markdown_file)
    transcripts_dir = Path(args.transcripts_dir)
    summaries_dir = Path(args.summaries_dir)

    new_content = insert_summaries_into_markdown(
        markdown_path=markdown_path,
        transcripts_dir=transcripts_dir,
        summaries_dir=summaries_dir,
    )

    if args.in_place:
        markdown_path.write_text(new_content, encoding="utf-8")
    elif args.output:
        Path(args.output).write_text(new_content, encoding="utf-8")
    else:
        print(new_content, end="")


if __name__ == "__main__":
    main()
