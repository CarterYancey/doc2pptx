# doc2pptx

Convert documents (txt, md, docx, pdf, html, xlsx, csv) into PowerPoint presentations.

## Features

- Converts multiple document formats to PPTX
- Supports optional `.pptx` templates to inherit styles, themes, fonts, and slide layouts
- Parses headings, body text, and tables
- Auto-chunks long slides and splits large tables across multiple slides
- Generates formatted PowerPoint presentations

## Installation

Requires Python ≥3.12 and [uv](https://docs.astral.sh/uv/). Install dependencies with:

```bash
uv sync
```

## Usage

Basic conversion:

```bash
uv run doc2pptx.py input.md -o output.pptx
```

With a style template:

```bash
uv run doc2pptx.py report.pdf -o output.pptx --template brand.pptx
```

With custom title:

```bash
uv run doc2pptx.py notes.txt -o output.pptx --template brand.pptx --title "My Deck"
```

Limit bullets per slide:

```bash
uv run doc2pptx.py article.docx -o slides.pptx --max-bullets 5
```

Spreadsheet with custom rows per table slide:

```bash
uv run doc2pptx.py data.csv -o tables.pptx --max-table-rows 10
```

## Supported Formats

| Format | Extensions |
|--------|------------|
| Text | `.txt` |
| Markdown | `.md`, `.markdown` |
| Word | `.docx` |
| PDF | `.pdf` |
| HTML | `.html`, `.htm` |
| Spreadsheet | `.xlsx`, `.xls`, `.xlsm`, `.csv`, `.tsv` |

## Options

| Option | Default | Description |
|--------|---------|-------------|
| `-o`, `--output` | `output.pptx` | Output file path |
| `-t`, `--template` | — | Template `.pptx` for styling |
| `--title` | *(from document)* | Override presentation title |
| `--max-bullets` | `7` | Bullet points per slide before auto-splitting |
| `--max-table-rows` | `15` | Data rows per table slide (spreadsheets only) |
| `--no-llm` | *(off)* | Skip the Ollama rewrite step |
| `--ollama-host` | `http://localhost:11434` | Ollama server base URL (env: `OLLAMA_HOST`) |
| `--ollama-model` | `llama3.2` | Ollama model name (env: `OLLAMA_MODEL`) |
| `--prompt-file` | — | Path to a text file with a custom system prompt |

## LLM rewrite (optional)

When a local [Ollama](https://ollama.com) server is running, doc2pptx will
reshape the extracted text into PPT-friendly markdown (concise bullets, clear
slide headings) before building slides. This usually produces tighter decks
from prose-heavy sources like PDFs and long Word documents.

- The rewrite is **on by default**; pass `--no-llm` to skip it.
- If the server is unreachable, doc2pptx prints a warning and falls back to
  the raw extracted text — generation always continues.
- Spreadsheet inputs (`.xlsx`, `.csv`, …) bypass the LLM step entirely.

Pull a model once, then run as usual:

```bash
ollama pull llama3.2
uv run doc2pptx.py report.pdf -o deck.pptx
```

Override the model, host, or prompt:

```bash
uv run doc2pptx.py report.pdf -o deck.pptx \
    --ollama-model qwen2.5 \
    --ollama-host http://localhost:11434 \
    --prompt-file my_prompt.txt
```

## Templates

Pass an existing PowerPoint file with `--template` to inherit its design. The tool will apply backgrounds, color themes, fonts, and slide layouts from the template.

## Spreadsheets

Each worksheet becomes one or more table slides. Sheets with more rows than `--max-table-rows` are automatically split across multiple slides. Columns beyond 10 are truncated.
