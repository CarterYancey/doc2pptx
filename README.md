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

## Templates

Pass an existing PowerPoint file with `--template` to inherit its design. The tool will apply backgrounds, color themes, fonts, and slide layouts from the template.

## Spreadsheets

Each worksheet becomes one or more table slides. Sheets with more rows than `--max-table-rows` are automatically split across multiple slides. Columns beyond 10 are truncated.
