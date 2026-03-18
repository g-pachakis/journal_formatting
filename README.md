# Journal Formatting Tool

A desktop GUI tool that reformats `.docx` manuscripts into journal-specific submission formats. Select your manuscript, pick a target journal format, and save the formatted output — no manual reformatting needed.

Currently supports **MDPI** and **Elsevier** formats. New formats can be added by dropping a single Python file into the `formats/` directory.

## Why This Exists

Preparing a manuscript for journal submission means hours of tedious reformatting: adjusting fonts, margins, heading styles, table borders, reference formatting, and more. Each journal has its own template and requirements. If your paper gets rejected and you resubmit elsewhere, you start the formatting over again.

This tool automates that. It reads your manuscript once, classifies every element (title, abstract, headings, tables, figures, equations, references), and rebuilds the document in the target journal's format — complete with correct typography, spacing, three-line tables, and placeholder fields for authors and affiliations.

## Key Features

- **Two journal formats** — MDPI and Elsevier, with more coming
- **Word-recognizable captions** — Table and Figure captions use SEQ fields, so Word's cross-references, "Insert Table of Figures", and auto-numbering all work
- **RIS bibliography import** — Upload a `.ris` file exported from Zotero, Mendeley, or any reference manager. The tool cross-checks your hardwritten citations against the RIS metadata and reformats them in the correct journal style (MDPI numbered format)
- **Zotero field code embedding** — Optionally wrap references in Zotero `ADDIN ZOTERO_ITEM CSL_CITATION` field codes, so Zotero can recognize and manage them after formatting
- **Plugin architecture** — Drop a new `.py` file in `formats/` to add a journal format

## Supported Formats

| Format | Font | Page Layout | Reference Style |
|--------|------|-------------|-----------------|
| **MDPI** | Palatino Linotype 10pt | A4, asymmetric margins (2.5/1.6/1.27/1.27 cm) | Numbered `[N]` — `Lastname, F.M.; ... Title. Journal Year, Vol, Pages.` |
| **Elsevier** | Times New Roman 12pt | A4, 2.5cm uniform margins, 1.5x line spacing | Preserved from source (hanging indent, 11pt) |

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

A window opens with these controls:

1. **Manuscript (.docx)** — Select your manuscript file
2. **Bibliography (.ris)** — Optionally load a `.ris` file for reference matching
3. **Format** — Choose MDPI or Elsevier
4. **Options** — Check "Embed Zotero field codes" if you want Zotero integration
5. **Format Manuscript** — Processes and opens a Save As dialog

## RIS Bibliography Workflow

If you manage references with Zotero, Mendeley, EndNote, or any citation manager:

1. **Export your library** as a `.ris` file from your reference manager
2. **Load the `.ris` file** in the tool alongside your manuscript
3. The tool matches each hardwritten reference (e.g., `[1] Shannon, C.E. ...`) against the RIS metadata using author names, year, title, and DOI
4. **Matched references** are reformatted in the target journal's style with correct author formatting, punctuation, and DOI links
5. **Unmatched references** are preserved as-is from your manuscript

### Zotero Field Codes

When the "Embed Zotero field codes" option is checked:

- Each matched reference is wrapped in a `ADDIN ZOTERO_ITEM CSL_CITATION` Word field code
- The field contains complete CSL JSON bibliographic metadata
- When you open the output in Word with Zotero installed, Zotero can recognize and manage the citations
- You can then use Zotero's "Refresh" to update the bibliography or switch citation styles

This is useful when you want to continue editing the formatted manuscript while keeping Zotero citation management active.

## How It Works

```
manuscript.docx ─────┐
                      v
                 ┌─────────┐     Classifies every paragraph, table, and element
                 │ reader.py│     into semantic types (heading, abstract, etc.)
                 └────┬─────┘     Returns structured data — not raw XML
                      │
refs.ris ────┐        │
             v        v
        ┌──────────────────┐
        │  ris_parser.py   │     Matches [N] citations to RIS records
        │  citation_fmt.py │     Reformats matched refs in journal style
        └────────┬─────────┘
                 │
                 v
        ┌──────────────────┐
        │  formats/*.py    │     Builds .docx with journal-specific styles,
        │  caption_fields  │     SEQ field captions, optional Zotero codes
        └────────┬─────────┘
                 │
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

## MDPI Reference Format

MDPI uses a numbered citation style. When a `.ris` file is provided, references are automatically formatted as:

**Journal article:**
```
Shannon, C.E.; Weaver, W. A Mathematical Theory of Communication. Bell Syst. Tech. J. 1948, 27, 379-423. https://doi.org/10.1002/j.1538-7305.1948.tb01338.x.
```

**Book:**
```
Smith, J.D. Introduction to Chemical Engineering, 3rd; Wiley: New York, 2020.
```

**Book chapter:**
```
Jones, R.A. Absorption in Packed Columns. In Handbook of Chemical Engineering; Smith, J.D., Ed.; McGraw-Hill: New York, 2019; pp. 145-198.
```

Author names are formatted as `Lastname, F.M.` with semicolons between authors. DOIs are appended as full URLs.

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
from caption_fields import add_caption_with_seq

FORMAT_NAME = 'Springer'       # Shown in the GUI
FORMAT_SUFFIX = '_Springer'    # Default output filename suffix

def build(items, output_path, ris_data=None, zotero_enabled=False):
    """Build formatted document from reader items.

    Args:
        items: List of dicts from reader.read_manuscript()
        output_path: Where to save the .docx
        ris_data: Optional list of RIS records (from ris_parser.parse_ris)
        zotero_enabled: Whether to embed Zotero field codes

    Returns:
        output_path
    """
    doc = DocxDocument()

    for item in items:
        if item['type'] == 'table_caption':
            # Use SEQ field for Word-recognizable caption
            import re as _re
            m = _re.match(r'^Table\s+(\d+)\.?\s*(.*)', item['text'])
            if m:
                para = doc.add_paragraph()
                add_caption_with_seq(para, 'Table', m.group(1),
                                     description=m.group(2))
        elif item['type'] == 'paragraph':
            para = doc.add_paragraph()
            for run_data in item['runs']:
                run = para.add_run(run_data['text'])
                run.bold = run_data['bold']
                run.italic = run_data['italic']
        # ... handle other types

    doc.save(output_path)
    return output_path
```

Save the file, restart the tool, and your new format appears automatically.

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
├── ris_parser.py               # RIS bibliography file parser
├── citation_formatter.py       # MDPI reference formatter
├── caption_fields.py           # Word SEQ field captions & Zotero field codes
├── formats/
│   ├── __init__.py             # Plugin auto-discovery
│   ├── elsevier.py             # Elsevier format builder
│   └── mdpi.py                 # MDPI format builder
├── tests/                      # 53 tests
│   ├── conftest.py             # Test fixtures (sample manuscripts)
│   ├── test_reader.py          # Reader tests (14)
│   ├── test_registry.py        # Plugin registry tests (4)
│   ├── test_elsevier.py        # Elsevier builder tests (8)
│   ├── test_mdpi.py            # MDPI builder tests (10)
│   ├── test_integration.py     # End-to-end pipeline tests (4)
│   ├── test_ris_parser.py      # RIS parser & formatter tests (9)
│   └── test_caption_fields.py  # Caption & Zotero field tests (4)
├── requirements.txt
└── .gitignore
```

## Running Tests

```bash
# Activate venv first, then:
python -m pytest tests/ -v
```

53 tests cover the reader, plugin registry, both format builders, RIS parsing, citation formatting, caption fields, Zotero integration, and full pipeline.

## Dependencies

| Package | Purpose |
|---------|---------|
| [python-docx](https://python-docx.readthedocs.io/) | Read and write `.docx` files |
| [lxml](https://lxml.de/) | XML manipulation for table borders, SEQ fields, and Zotero field codes |
| [pytest](https://docs.pytest.org/) | Testing (dev only) |

Tkinter is included with Python's standard library. No external citation libraries needed — RIS parsing and formatting are built in.

## Limitations

- **Images are not transferred** — the tool processes text, tables, and structure. Figures need to be re-inserted manually.
- **Complex equations** (MathType, Equation Editor objects) are read as text only.
- **Custom styles** in your source document may not be detected — use standard Word heading styles for best results.
- **Multi-column layouts** are not applied (both MDPI and Elsevier submission formats use single-column).
- **RIS matching** uses author names, year, and title fragments — unusual author name formats may not match perfectly.
- **Zotero field codes** require Zotero's Word plugin to be functional after opening the document.

## License

MIT

---

## Feedback & Contributions

This tool was built to solve a real pain point in academic publishing. If you use it, your feedback would be genuinely valuable:

**As a researcher:**
- Which journal formats would you like to see added next? (Springer, IEEE, ACS, RSC, Wiley, ...)
- Does the output match what you'd expect from your target journal's template?
- Are there manuscript elements the tool misses or misclassifies?
- How well does the RIS matching work with your reference library?
- Would you prefer a command-line interface in addition to the GUI?

**As a developer:**
- See a bug or edge case? [Open an issue](https://github.com/g-pachakis/journal_formatting/issues).
- Want to contribute a new format plugin? PRs are welcome — the plugin API is documented above.
- Have ideas for the architecture? The reader/plugin split is designed to be extended.
- Want to add a new citation style? See `citation_formatter.py` for the pattern.

**Get in touch:** Open an issue on [GitHub](https://github.com/g-pachakis/journal_formatting/issues) or submit a pull request. All contributions — bug reports, format requests, code, documentation — are appreciated.
