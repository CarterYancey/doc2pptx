#!/usr/bin/env python3

import os
import json
import time
from pathlib import Path
from typing import Any, Dict

import requests

OPENAI_API_KEY = os.environ["OPENAI_API_KEY"]

INPUT_DIR = Path("transcripts")
OUTPUT_DIR = Path("summaries")
OUTPUT_DIR.mkdir(exist_ok=True)

MODEL = "gpt-5.4"
API_URL = "https://api.openai.com/v1/chat/completions"

REQUEST_TIMEOUT = 300
SLEEP_BETWEEN_REQUESTS = 1.0

DEVELOPER_PROMPT = """You are summarizing a YouTube video from an automatically extracted transcript.

The transcript may contain:
- speech recognition errors
- missing punctuation
- incorrect capitalization
- malformed names, brands, products, or technical terms
- duplicated phrases
- broken sentence boundaries
- omitted context around speaker transitions

Your job:
1. Read the transcript generously and infer the intended meaning when it is reasonably clear.
2. Silently correct obvious transcript noise in your understanding before summarizing.
3. Do not invent facts that are not supported by the transcript.
4. If something is ambiguous, represent it cautiously.
5. Preserve the speaker's actual arguments, sequence, and emphasis.
6. Produce a detailed summary that captures the video's structure, main ideas, supporting details, examples, and conclusions.
7. Prefer concrete wording over vague wording.
8. When the speaker appears to mention a proper noun or technical term unclearly, normalize it only if the intended reference is very likely; otherwise note uncertainty.
9. Focus on content extraction, not style critique.

Output valid JSON only, matching the provided schema.
"""

USER_PROMPT_TEMPLATE = """Summarize this transcript in a way that is useful for creating a detailed PowerPoint presentation later.

Requirements:
- Write for an analyst who has not watched the video.
- Capture the main thesis, major sections, and supporting arguments.
- Include important examples, frameworks, processes, and recommendations.
- Reconstruct the likely intended meaning where transcript noise is obvious.
- Do not mention every small aside unless it helps explain the core message.
- If the video is argumentative, identify the claims and evidence.
- If the transcript is unclear in places, say so briefly and move on.

Transcript:
\"\"\"
{transcript}
\"\"\"
"""

RESPONSE_FORMAT = {
    "type": "json_schema",
    "json_schema": {
        "name": "video_summary",
        "schema": {
            "type": "object",
            "additionalProperties": False,
            "properties": {
                "title": {"type": "string"},
                "executive_summary": {"type": "string"},
                "detailed_summary": {
                    "type": "array",
                    "items": {"type": "string"}
                },
                "key_points": {
                    "type": "array",
                    "items": {"type": "string"}
                },
            },
            "required": [
                "title",
                "executive_summary",
                "detailed_summary",
                "key_points",
            ]
        }
    }
}

def call_openai(transcript_text: str) -> Dict[str, Any]:
    headers = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json",
    }

    payload = {
        "model": MODEL,
        "messages": [
            {"role": "developer", "content": DEVELOPER_PROMPT},
            {
                "role": "user",
                "content": USER_PROMPT_TEMPLATE.format(transcript=transcript_text)
            },
        ],
        "response_format": RESPONSE_FORMAT,
        "temperature": 0.2,
    }

    resp = requests.post(API_URL, headers=headers, json=payload, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    data = resp.json()

    content = data["choices"][0]["message"]["content"]
    return json.loads(content)

def to_markdown(summary: Dict[str, Any]) -> str:
    lines = []
    lines.append(f"# {summary['title']}\n")
    lines.append("## Executive Summary\n")
    lines.append(summary["executive_summary"] + "\n")

    def add_list_section(title: str, items: list[str]):
        lines.append(f"## {title}\n")
        if not items:
            lines.append("- None\n")
            return
        for item in items:
            lines.append(f"- {item}")
        lines.append("")

    add_list_section("Detailed Summary", summary["detailed_summary"])
    add_list_section("Key Points", summary["key_points"])

    return "\n".join(lines)

def summarize_file(path: Path) -> None:
    transcript = path.read_text(encoding="utf-8", errors="replace").strip()
    if not transcript:
        print(f"[SKIP] Empty file: {path.name}")
        return

    print(f"[START] {path.name}")

    try:
        summary = call_openai(transcript)

        json_path = OUTPUT_DIR / f"{path.stem}.simple_summary.json"
        md_path = OUTPUT_DIR / f"{path.stem}.simple_summary.md"

        json_path.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
        md_path.write_text(to_markdown(summary), encoding="utf-8")

        print(f"[OK] {path.name} -> {json_path.name}, {md_path.name}")

    except Exception as e:
        print(f"[ERROR] {path.name}: {e}")

def main():
    txt_files = sorted(INPUT_DIR.glob("*.txt"))
    txt_files = [f for f in txt_files if not f.name.endswith("1.txt")]
    print("txt_files: ",txt_files)
    if not txt_files:
        print(f"No .txt files found in {INPUT_DIR.resolve()}")
        return

    for path in txt_files:
        summarize_file(path)
        time.sleep(SLEEP_BETWEEN_REQUESTS)

if __name__ == "__main__":
    main()
