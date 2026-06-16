# DocSkill

A Python tool that reads documents in various formats, applies professional formatting (Times New Roman, standardised sizes), correctly preserves mathematical equations, numbers figure/table captions, and optionally generates plots from numeric tables.

## Features

- **Multi-format input**: DOCX, Markdown, HTML, LaTeX, TXT, RST, and more (via Pandoc)
- **Professional formatting**: Times New Roman with Title (18pt bold), Headings (12pt bold), Body (11pt)
- **Equation preservation**: Native Word OMML equations are never corrupted during formatting
- **Caption numbering**: Automatically standardises `Figure N:` and `Table N:` labels
- **Figure generation**: Optionally creates bar charts from numeric tables and inserts them into the document
- **Interactive mode**: Just double-click `run_docskill.bat` — no command-line knowledge needed

## Setup

1. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Pandoc is required for non-DOCX formats. If not installed, DocSkill will download it automatically on first use. You can also install it manually from [pandoc.org](https://pandoc.org/installing.html).

## Usage

### One-Click (Windows)

Double-click **`run_docskill.bat`**. An interactive menu will guide you through the options.

### Interactive Mode

```bash
python main.py
```

You'll be asked:
1. Where to read files from (default: `input_docs/`)
2. Where to save results (default: `output_docs/`)
3. Whether to auto-generate figures from tables

### Command Line

```bash
# Format all docs in input_docs/ → output_docs/
python main.py

# Format a specific file
python main.py -i my_paper.docx -o formatted_paper.docx

# Format and auto-generate figures
python main.py --generate-figures
```

## Directory Structure

```
DocSkill/
├── input_docs/          ← Place your documents here
├── output_docs/         ← Formatted documents appear here
├── output_figures/      ← Generated chart images
├── main.py              ← CLI entry point
├── doc_parser.py        ← Document parsing & format conversion
├── formatter.py         ← Professional formatting engine
├── figure_generator.py  ← Auto chart generation from tables
├── run_docskill.bat     ← One-click Windows launcher
├── requirements.txt     ← Python dependencies
└── .gitignore
```

## Supported Formats

| Extension | Notes |
|-----------|-------|
| `.docx`   | Native — highest fidelity |
| `.md`     | Full LaTeX math support (`$...$`, `$$...$$`) |
| `.tex`    | LaTeX documents |
| `.html`   | HTML with MathML or LaTeX math |
| `.txt`    | Plain text |
| `.rst`    | reStructuredText |
| `.odt`    | OpenDocument |

## License

MIT
