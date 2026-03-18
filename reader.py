"""
Shared manuscript reader.

Reads a .docx manuscript, classifies each element (paragraph, table, etc.)
into a semantic type, and returns structured items for format plugins.
"""

import re
from docx import Document as DocxDocument
from docx.oxml.ns import qn as docx_qn


def get_style(p_elem):
    """Get Word style ID from a <w:p> element."""
    ppr = p_elem.find(docx_qn('w:pPr'))
    if ppr is not None:
        ps = ppr.find(docx_qn('w:pStyle'))
        if ps is not None:
            return ps.get(docx_qn('w:val'))
    return ''


def get_text(elem):
    """Extract all text from an element."""
    return ''.join(t.text or '' for t in elem.iter(docx_qn('w:t')))


def get_runs(p_elem):
    """Extract runs with formatting info from a paragraph element.

    Returns list of dicts: [{text, bold, italic, superscript, subscript}, ...]
    """
    runs = []
    for r in p_elem.findall(docx_qn('w:r')):
        rpr = r.find(docx_qn('w:rPr'))
        text = ''.join(t.text or '' for t in r.findall(docx_qn('w:t')))

        bold = False
        italic = False
        superscript = False
        subscript = False

        if rpr is not None:
            bold = rpr.find(docx_qn('w:b')) is not None
            italic = rpr.find(docx_qn('w:i')) is not None
            va = rpr.find(docx_qn('w:vertAlign'))
            if va is not None:
                val = va.get(docx_qn('w:val'))
                superscript = val == 'superscript'
                subscript = val == 'subscript'

        runs.append({
            'text': text,
            'bold': bold,
            'italic': italic,
            'superscript': superscript,
            'subscript': subscript,
        })
    return runs


def classify_paragraph(p_elem, is_after_refs, prev_classification):
    """Classify a manuscript paragraph into a semantic type."""
    style = get_style(p_elem)
    text = get_text(p_elem).strip()

    if not text:
        return 'empty'

    # Style-based classification
    if style in ('Heading1', 'Heading 1'):
        text_lower = text.lower()
        # Strip leading section number: "1. Introduction" -> "introduction"
        stripped = re.sub(r'^\d+\.\s*', '', text_lower)
        if stripped == 'abstract':
            return 'abstract_heading'
        if 'references' in stripped:
            return 'references_heading'
        return 'heading1'

    if style in ('Heading2', 'Heading 2'):
        return 'heading2'

    if style in ('Heading3', 'Heading 3'):
        return 'heading3'

    if style == 'Bibliography':
        return 'reference'

    # Text pattern-based
    if text.startswith('Keywords:') or text.startswith('Keywords :'):
        return 'keywords'

    if re.match(r'^Table\s+\d+\.', text):
        return 'table_caption'

    if re.match(r'^\[?\s*Figure\s+\d+', text, re.IGNORECASE):
        return 'figure_placeholder'

    if text.startswith('*') and len(text) < 400:
        return 'table_footer'

    if is_after_refs and re.match(r'^\[?\d+\]', text):
        return 'reference'

    # Equation detection
    if re.search(r'\(Eq\.\s*\d+\)', text):
        return 'equation'

    # Context-based: abstract text follows abstract heading
    # (supports multi-paragraph abstracts until keywords or next heading)
    if prev_classification in ('abstract_heading', 'abstract_text'):
        return 'abstract_text'

    return 'paragraph'


def read_table(tbl_elem):
    """Read a table element into a list of rows.

    Each row is a list of cell dicts:
    {text: str, gridspan: int, runs: list[dict], vmerge_continue: bool}
    """
    rows = []
    for tr in tbl_elem.findall(docx_qn('w:tr')):
        cells = []
        for tc in tr.findall(docx_qn('w:tc')):
            tcpr = tc.find(docx_qn('w:tcPr'))
            gridspan = 1
            if tcpr is not None:
                gs = tcpr.find(docx_qn('w:gridSpan'))
                if gs is not None:
                    gridspan = int(gs.get(docx_qn('w:val'), '1'))
                vm = tcpr.find(docx_qn('w:vMerge'))
                if vm is not None and vm.get(docx_qn('w:val'), '') != 'restart':
                    cells.append({
                        'text': '',
                        'gridspan': gridspan,
                        'runs': [],
                        'vmerge_continue': True,
                    })
                    continue

            cell_text = get_text(tc)
            cell_runs = []
            for p in tc.findall(docx_qn('w:p')):
                cell_runs.extend(get_runs(p))

            cells.append({
                'text': cell_text,
                'gridspan': gridspan,
                'runs': cell_runs,
                'vmerge_continue': False,
            })
        rows.append(cells)
    return rows


def read_manuscript(path):
    """Read a .docx manuscript and return classified content items in document order."""
    doc = DocxDocument(path)
    body = doc.element.body

    elements = []
    is_after_refs = False
    prev_class = None

    for child in body:
        tag = child.tag.split('}')[-1]

        if tag == 'p':
            classification = classify_paragraph(child, is_after_refs, prev_class)

            if classification == 'references_heading':
                is_after_refs = True

            if classification == 'empty':
                prev_class = classification
                continue

            text = get_text(child).strip()
            runs = get_runs(child)

            elements.append({
                'type': classification,
                'text': text,
                'runs': runs,
            })
            prev_class = classification

        elif tag == 'tbl':
            table_data = read_table(child)
            elements.append({
                'type': 'table',
                'text': '',
                'runs': [],
                'rows': table_data,
            })
            prev_class = 'table'

        elif tag == 'sectPr':
            continue

    return elements
