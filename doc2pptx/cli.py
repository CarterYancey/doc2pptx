"""Command-line entry point for doc2pptx."""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from .llm import DEFAULT_OLLAMA_HOST, DEFAULT_OLLAMA_MODEL
from .logging_config import _default_log_path, configure_logging, logger
from .pipeline import generate_pptx
from .readers import MAX_TABLE_ROWS_PER_SLIDE


def main():
    parser = argparse.ArgumentParser(
        description="Convert documents (txt, md, docx, pdf, html, xlsx, csv) to PowerPoint presentations.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python -m doc2pptx notes.md -o presentation.pptx
  python -m doc2pptx report.pdf -o deck.pptx --template brand.pptx
  python -m doc2pptx article.docx -o slides.pptx --title "Q4 Report"
  python -m doc2pptx page.html -o slides.pptx --max-bullets 5
  python -m doc2pptx data.xlsx -o tables.pptx
  python -m doc2pptx data.csv -o tables.pptx --max-table-rows 10
  python -m doc2pptx report.pdf -o deck.pptx --no-llm
  python -m doc2pptx report.pdf -o deck.pptx --ollama-model qwen2.5
        """,
    )
    parser.add_argument("input", help="Input document path (.txt, .md, .docx, .pdf, .html, .xlsx, .csv, .tsv)")
    parser.add_argument("-o", "--output", default="output.pptx", help="Output .pptx path (default: output.pptx)")
    parser.add_argument("-t", "--template", default=None, help="Template .pptx to use for styling")
    parser.add_argument("--title", default=None, help="Override the presentation title")
    parser.add_argument("--max-bullets", type=int, default=7, help="Max bullet points per slide before splitting (default: 7)")
    parser.add_argument("--max-table-rows", type=int, default=MAX_TABLE_ROWS_PER_SLIDE,
                        help=f"Max data rows per table slide (default: {MAX_TABLE_ROWS_PER_SLIDE})")
    parser.add_argument("--no-llm", action="store_true",
                        help="Skip the Ollama rewrite step (LLM rewrite is on by default).")
    parser.add_argument("--ollama-host", default=DEFAULT_OLLAMA_HOST,
                        help=f"Ollama server base URL (default: {DEFAULT_OLLAMA_HOST}, env: OLLAMA_HOST)")
    parser.add_argument("--ollama-model", default=DEFAULT_OLLAMA_MODEL,
                        help=f"Ollama model name (default: {DEFAULT_OLLAMA_MODEL}, env: OLLAMA_MODEL)")
    parser.add_argument("--prompt-file", default=None,
                        help="Path to a text file containing a custom system prompt for the rewrite step.")
    parser.add_argument("--max-chunk-chars", type=int, default=6000,
                        help="Max characters per LLM chunk; large docs are split on headings/paragraphs (default: 6000).")
    parser.add_argument("--chunk-overlap-sentences", type=int, default=2,
                        help="Trailing sentences from the previous chunk to prepend to the next as context (default: 2).")
    parser.add_argument("--log-file", default=None,
                        help="Path for the verbose run log. If omitted, a timestamped log is written under ./logs/. "
                             "Full LLM prompts, inputs, and responses are captured at DEBUG level.")
    parser.add_argument("--no-log-file", action="store_true",
                        help="Disable writing a run log file (terminal logging still runs).")
    parser.add_argument("-v", "--verbose", action="store_true",
                        help="Show DEBUG-level output in the terminal (chunk previews, prompts, responses).")
    parser.add_argument("-q", "--quiet", action="store_true",
                        help="Only show WARNING-level output in the terminal (the log file still captures everything).")

    args = parser.parse_args()

    if args.verbose and args.quiet:
        print("Error: --verbose and --quiet are mutually exclusive.", file=sys.stderr)
        sys.exit(2)

    log_file_path: Path | None = None
    if not args.no_log_file:
        log_file_path = Path(args.log_file) if args.log_file else _default_log_path(args.input)
    resolved_log = configure_logging(
        log_file=log_file_path,
        verbose=args.verbose,
        quiet=args.quiet,
    )
    if resolved_log is not None:
        logger.info("📝 Writing run log to %s", resolved_log)

    if not os.path.exists(args.input):
        logger.error("Input file not found: %s", args.input)
        sys.exit(1)

    if args.template and not os.path.exists(args.template):
        logger.error("Template file not found: %s", args.template)
        sys.exit(1)

    if args.prompt_file and not os.path.exists(args.prompt_file):
        logger.error("Prompt file not found: %s", args.prompt_file)
        sys.exit(1)

    generate_pptx(
        input_path=args.input,
        output_path=args.output,
        template_path=args.template,
        title=args.title,
        max_bullets=args.max_bullets,
        max_table_rows=args.max_table_rows,
        use_llm=not args.no_llm,
        ollama_host=args.ollama_host,
        ollama_model=args.ollama_model,
        llm_prompt_file=args.prompt_file,
        max_chunk_chars=args.max_chunk_chars,
        chunk_overlap_sentences=args.chunk_overlap_sentences,
    )
