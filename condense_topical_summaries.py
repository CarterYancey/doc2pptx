#!/usr/bin/env python3

import os
import json
import time
from pathlib import Path
from typing import Any, Dict

import requests

OPENAI_API_KEY = os.environ["OPENAI_API_KEY"]

INPUT_DIR = Path("TopicalSummaries")
OUTPUT_DIR = Path("TopicalSummaries_condensed")
OUTPUT_DIR.mkdir(exist_ok=True)

MODEL = "gpt-5.4"
API_URL = "https://api.openai.com/v1/chat/completions"

REQUEST_TIMEOUT = 300
SLEEP_BETWEEN_REQUESTS = 1.0

DEVELOPER_PROMPT = """You are compiling summaries of YouTube lectures into a single instruction guide. The summaries all cover related topics, but they may make overlapping and/or distinct points. Your job is to create a single, extremely detailed informational document based on the lectures. Your document should not read like a video summary but rather like a single lecture. The lecture is designed to be part of a intensive bootcamp on AI for technical managers of AI teams.

The summaries may contain:
- repeated points
- different perspectives
- 3rd person tone
- duplicated phrases

Your job:
1. Read the summaries generously 
2. Do not invent facts that are not supported by the summaries.
3. Preserve the actual arguments, supporting details, examples, and emphases.
4. Prefer concrete wording over vague wording.

Output valid JSON only, matching the provided schema.
"""

USER_PROMPT_TEMPLATE = """Compile these summaries in a way that is useful for creating a detailed presentation later.

Requirements:
- Write for an analyst who has not watched the videos.
- Capture the main thesis, major sections, and supporting arguments of all videos mentioned.
- Include important examples, frameworks, processes, and recommendations.
- Do not repeat points or include asides not relevent to the topic.
- Do not write it like a summary itself, but as a standalone, highly informative lecture.

Summaries:
\"\"\"
{transcript}
\"\"\"
"""

RESPONSE_FORMAT = {
    "type": "json_schema",
    "json_schema": {
        "name": "Lecture",
        "schema": {
            "type": "object",
            "additionalProperties": False,
            "properties": {
                "title": {"type": "string"},
                "executive_summary": {"type": "string"},
                "key_points": {
                    "type": "array",
                    "items": {"type": "string"}
                },
                "detailed_information": {
                    "type": "array",
                    "items": {"type": "string"}
                },
            },
            "required": [
                "title",
                "executive_summary",
                "key_points",
                "detailed_information",
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

    add_list_section("Key Points", summary["key_points"])
    add_list_section("Details", summary["detailed_information"])

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
    txt_files = sorted(INPUT_DIR.glob("*.md"))
    if not txt_files:
        print(f"No .txt files found in {INPUT_DIR.resolve()}")
        return

    for path in txt_files:
        summarize_file(path)
        time.sleep(SLEEP_BETWEEN_REQUESTS)

if __name__ == "__main__":
    main()
