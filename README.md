# Journal Formatting Tool

A desktop GUI tool that reformats `.docx` manuscripts into journal-specific submission formats. Select your manuscript, pick a target journal format, and save the formatted output — no manual reformatting needed.

Currently supports **MDPI** and **Elsevier** formats. New formats can be added by dropping a single Python file into the `formats/` directory.

## Why This Exists

Preparing a manuscript for journal submission means hours of tedious reformatting: adjusting fonts, margins, heading styles, table borders, reference formatting, and more. Each journal has its own template and requirements. If your paper gets rejected and you resubmit elsewhere, you start the formatting over again.

This tool automates that. It reads your manuscript once, classifies every element (title, abstract, headings, tables, figures, equations, references), and rebuilds the document in the target journal's format — complete with correct typography, spacing, three-line tables, and placeholder fields for authors and affiliations.

## Supported Formats

| Format | Font | Page Layout | Key Features |
|--------|------|-------------|-------------|
| **MDPI** | Palatino Linotype 10pt | A4, asymmetric margins (2.5/1.6/1.27/1.27 cm) | Article type header, 130.4pt left indent, outline-level headings, "at least" line spacing |
| **Elsevier** | Times New Roman 12pt | A4, 2.5cm uniform margins, 1.5x line spacing | Highlights section, graphical abstract placeholder, hanging-indent references, first-line indent after headings |

Both formats produce three-line tables (top border, header-bottom border, table-bottom border) and preserve inline formatting (bold, italic, superscript, subscript) from your source document.

## Quick Start

### Requirements

- Python 3.9+
- A `.docx` manuscript with standard heading styles (Heading 1, Heading 2, Heading 3)

### Installation

```bash
git clone https://github.com/g-pachakis/journal_formatting.git
cd journal_formatting
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate

pip install -r requirements.txt
```

### Run

```bash
python manuscript_formatter.py
```

A window opens with three steps:
1. **Open File** — select your `.docx` manuscript
2. **Choose format** — MDPI or Elsevier (radio buttons)
3. **Format Manuscript** — processes and opens a Save As dialog

## How It Works

```
manuscript.docx
       |
       v
  ┌─────────┐     Classifies every paragraph, table, and element
  │ reader.py│     into semantic types (heading, abstract, reference, etc.)
  └────┬─────┘     Returns structured data — not raw XML
       |
       v
  ┌──────────────┐
  │ formats/*.py  │   Each plugin builds a .docx from scratch
  │  mdpi.py      │   using python-docx with journal-specific
  │  elsevier.py  │   styles, spacing, and layout
  └──────┬────────┘
         |
         v
   formatted_output.docx
```

### What the Reader Detects

| Element | Detection Method |
|---------|-----------------|
| Title | First paragraph before any heading |
| Abstract | Paragraph(s) after "Abstract" heading |
| Keywords | Text starting with "Keywords:" |
| Headings (H1/H2/H3) | Word heading styles |
| Body paragraphs | Default classification |
| Table captions | Text matching "Table N." |
| Tables | Full cell data with merged cell support |
| Table footnotes | Text starting with `*` |
| Figure captions | Text matching "Figure N" |
| Equations | Text containing "(Eq. N)" |
| References | `[N]` pattern or Bibliography style |

## Manuscript Preparation Tips

For best results, your source `.docx` should use:

- **Heading 1** style for main sections (Introduction, Methods, Results, etc.)
- **Heading 2** / **Heading 3** for subsections
- Standard Word tables (not images of tables)
- `Keywords:` on its own paragraph
- References as individual paragraphs starting with `[1]`, `[2]`, etc.

The tool handles numbered headings (e.g., "1. Introduction") and unnumbered ones equally well.

## Adding a New Journal Format

Create a file in `formats/` — for example `formats/springer.py`:

```python
"""Springer Nature format plugin."""

import re
from docx import Document as DocxDocument
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

FORMAT_NAME = 'Springer'       # Shown in the GUI
FORMAT_SUFFIX = '_Springer'    # Default output filename suffix

def build(items, output_path):
    """Build formatted document from reader items.

    Args:
        items: List of dicts from reader.read_manuscript().
               Each has 'type', 'text', 'runs', and optionally 'rows'.
        output_path: Where to save the .docx

    Returns:
        output_path
    """
    doc = DocxDocument()

    # Set up page layout, styles, etc.
    # Loop through items and build the document
    # Each item['type'] tells you what it is:
    #   'heading1', 'heading2', 'heading3',
    #   'abstract_text', 'keywords', 'paragraph',
    #   'table_caption', 'table', 'table_footer',
    #   'figure_placeholder', 'equation',
    #   'references_heading', 'reference'

    for item in items:
        if item['type'] == 'paragraph':
            para = doc.add_paragraph()
            for run_data in item['runs']:
                run = para.add_run(run_data['text'])
                run.bold = run_data['bold']
                run.italic = run_data['italic']
                run.font.superscript = run_data['superscript']
                run.font.subscript = run_data['subscript']
        # ... handle other types

    doc.save(output_path)
    return output_path
```

Save the file, restart the tool, and your new format appears in the GUI automatically. No registration or configuration needed.

### Item Schema Reference

**Paragraph items** (`type` != `'table'`):
```python
{
    'type': str,        # Content classification
    'text': str,        # Full plain text
    'runs': [           # Formatted text segments
        {
            'text': str,
            'bold': bool,
            'italic': bool,
            'superscript': bool,
            'subscript': bool,
        }
    ]
}
```

**Table items** (`type` == `'table'`):
```python
{
    'type': 'table',
    'text': '',
    'runs': [],
    'rows': [           # List of rows, each a list of cells
        [
            {
                'text': str,
                'gridspan': int,            # Column span (1 = normal)
                'runs': [{ ... }],          # Same format as above
                'vmerge_continue': bool,    # True = vertically merged continuation
            }
        ]
    ]
}
```

## Project Structure

```
journal_formatting/
├── manuscript_formatter.py     # Tkinter GUI
├── reader.py                   # Manuscript reader & classifier
├── formats/
│   ├── __init__.py             # Plugin auto-discovery
│   ├── elsevier.py             # Elsevier format builder
│   └── mdpi.py                 # MDPI format builder
├── tests/
│   ├── conftest.py             # Test fixtures (sample manuscripts)
│   ├── test_reader.py          # Reader tests (14 tests)
│   ├── test_registry.py        # Plugin registry tests (4 tests)
│   ├── test_elsevier.py        # Elsevier builder tests (8 tests)
│   ├── test_mdpi.py            # MDPI builder tests (10 tests)
│   └── test_integration.py     # End-to-end pipeline tests (4 tests)
├── requirements.txt
└── .gitignore
```

## Running Tests

```bash
# Activate venv first, then:
python -m pytest tests/ -v
```

40 tests cover the reader, plugin registry, both format builders, and full pipeline integration.

## Dependencies

| Package | Purpose |
|---------|---------|
| [python-docx](https://python-docx.readthedocs.io/) | Read and write `.docx` files |
| [lxml](https://lxml.de/) | XML manipulation for table borders and advanced formatting |
| [pytest](https://docs.pytest.org/) | Testing (dev only) |

Tkinter is included with Python's standard library.

## Limitations

- **Images are not transferred** — the tool processes text, tables, and structure. Figures need to be re-inserted manually.
- **Complex equations** (MathType, Equation Editor objects) are read as text only.
- **Custom styles** in your source document may not be detected — use standard Word heading styles for best results.
- **Multi-column layouts** are not applied (both MDPI and Elsevier submission formats use single-column).

## License

MIT

---

## Feedback & Contributions

This tool was built to solve a real pain point in academic publishing. If you use it, your feedback would be genuinely valuable:

**As a researcher:**
- Which journal formats would you like to see added next? (Springer, IEEE, ACS, RSC, Wiley, ...)
- Does the output match what you'd expect from your target journal's template?
- Are there manuscript elements the tool misses or misclassifies?
- Would you prefer a command-line interface in addition to the GUI?

**As a developer:**
- See a bug or edge case? [Open an issue](https://github.com/g-pachakis/journal_formatting/issues).
- Want to contribute a new format plugin? PRs are welcome — the plugin API is documented above.
- Have ideas for the architecture? The reader/plugin split is designed to be extended.

**Get in touch:** Open an issue on [GitHub](https://github.com/g-pachakis/journal_formatting/issues) or submit a pull request. All contributions — bug reports, format requests, code, documentation — are appreciated.
