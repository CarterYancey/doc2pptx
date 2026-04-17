# doc2pptx

Convert documents (txt, md, docx, pdf) into PowerPoint presentations.

## Features

- Converts multiple document formats to PPTX
- Supports optional `.pptx` templates to inherit styles, themes, fonts, and slide layouts
- Parses headings, body text, and tables
- Generates formatted PowerPoint presentations

## Installation

Requires Python ≥3.12. Install dependencies with:

```bash
uv install
```

Or with pip:

```bash
pip install beautifulsoup4 markdown pdfplumber pypdf python-docx python-pptx
```

## Usage

Basic conversion:

```bash
python doc2pptx.py input.md -o output.pptx
```

With a style template:

```bash
python doc2pptx.py report.pdf -o output.pptx --template brand.pptx
```

With custom title:

```bash
python doc2pptx.py notes.txt -o output.pptx --template brand.pptx --title "My Deck"
```

## Supported Formats

- **Text**: `.txt`
- **Markdown**: `.md`
- **Word**: `.docx`
- **PDF**: `.pdf`

## Templates

Pass an existing PowerPoint file with `--template` to inherit its design. The tool will apply backgrounds, color themes, fonts, and slide layouts from the template.
