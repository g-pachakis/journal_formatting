"""
MDPI journal format plugin.

Builds a formatted .docx using a styles template extracted from a published
MDPI paper. All 39 MDPI named styles are loaded from the template, so the
output document has identical style definitions to editor-approved papers.
"""

import os
import re
from docx import Document as DocxDocument
from docx.oxml.ns import qn as docx_qn
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from lxml import etree as lxml_etree

from caption_fields import add_caption_with_seq, add_zotero_citation_field, ris_to_csl_json
from reference_engine import resolve_references

FORMAT_NAME = 'MDPI'
FORMAT_SUFFIX = '_MDPI'

FONT_NAME = 'Palatino Linotype'
FONT_COLOR = RGBColor(0x00, 0x00, 0x00)
DEFAULT_SIZE = Pt(10)

# Path to the styles template (extracted from benchmark)
_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'mdpi_styles.docx')


def _set_run_font(run, size=None, bold=None, italic=None,
                  superscript=False, subscript=False):
    """Apply font formatting to a run. Only sets explicit overrides."""
    run.font.name = FONT_NAME
    run.font.color.rgb = FONT_COLOR
    if size:
        run.font.size = size
    if bold:
        run.font.bold = True
    if italic:
        run.font.italic = True
    if superscript:
        run.font.superscript = True
    if subscript:
        run.font.subscript = True


def _styled_para(doc, style_name):
    """Create a paragraph with a named MDPI style."""
    para = doc.add_paragraph()
    try:
        para.style = doc.styles[style_name]
    except KeyError:
        pass
    return para


def _add_text_para(doc, style_name, text='', bold_prefix=None, font_size=None):
    """Create a styled paragraph with optional bold prefix + text."""
    para = _styled_para(doc, style_name)
    fs = font_size or DEFAULT_SIZE
    if bold_prefix:
        run = para.add_run(bold_prefix)
        _set_run_font(run, size=fs, bold=True)
    if text:
        run = para.add_run(text)
        _set_run_font(run, size=fs)
    return para


def _add_runs_para(doc, style_name, runs_data, bold_prefix=None, font_size=None):
    """Create a styled paragraph preserving inline formatting from runs."""
    para = _styled_para(doc, style_name)
    fs = font_size or DEFAULT_SIZE
    if bold_prefix:
        run = para.add_run(bold_prefix)
        _set_run_font(run, size=fs, bold=True)
    for rd in runs_data:
        if not rd['text']:
            continue
        run = para.add_run(rd['text'])
        _set_run_font(run, size=fs, bold=rd.get('bold'), italic=rd.get('italic'),
                      superscript=rd.get('superscript', False),
                      subscript=rd.get('subscript', False))
    return para


def _build_author_placeholder(doc):
    """Build MDPI author names paragraph with superscript affiliations."""
    para = _styled_para(doc, 'MDPI_1.3_authornames')
    parts = [
        ('Firstname Lastname ', {'bold': True}),
        ('1', {'superscript': True}),
        (', Firstname Lastname ', {'bold': True}),
        ('2', {'superscript': True}),
        (' and Firstname Lastname ', {'bold': True}),
        ('2,', {'superscript': True}),
        ('*', {'bold': True}),
    ]
    for text, fmt in parts:
        run = para.add_run(text)
        _set_run_font(run, size=DEFAULT_SIZE, **fmt)
    return para


def _add_three_line_table(doc, table_data):
    """Create an MDPI three-line table using the MDPI_4.1_three_line_table style."""
    if not table_data or not table_data[0]:
        return None
    ncols = sum(c.get('gridspan', 1) for c in table_data[0])
    if ncols == 0:
        return None
    nrows = len(table_data)

    table = doc.add_table(rows=nrows, cols=ncols)

    # Apply the MDPI three-line table style if available
    try:
        table.style = doc.styles['MDPI_4.1_three_line_table']
    except KeyError:
        # Fallback: set borders manually
        tbl_pr = table._tbl.find(docx_qn('w:tblPr'))
        if tbl_pr is None:
            tbl_pr = lxml_etree.SubElement(table._tbl, docx_qn('w:tblPr'))
        borders = lxml_etree.SubElement(tbl_pr, docx_qn('w:tblBorders'))
        for bname in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            el = lxml_etree.SubElement(borders, docx_qn(f'w:{bname}'))
            el.set(docx_qn('w:val'), 'none')
            el.set(docx_qn('w:sz'), '0')
            el.set(docx_qn('w:space'), '0')
            el.set(docx_qn('w:color'), 'auto')
        borders.find(docx_qn('w:top')).set(docx_qn('w:val'), 'single')
        borders.find(docx_qn('w:top')).set(docx_qn('w:sz'), '8')
        borders.find(docx_qn('w:bottom')).set(docx_qn('w:val'), 'single')
        borders.find(docx_qn('w:bottom')).set(docx_qn('w:sz'), '8')

    # Table indent for MDPI left-column gutter
    tbl_pr = table._tbl.find(docx_qn('w:tblPr'))
    ti = tbl_pr.find(docx_qn('w:tblInd'))
    if ti is None:
        ti = lxml_etree.SubElement(tbl_pr, docx_qn('w:tblInd'))
    ti.set(docx_qn('w:w'), '2608')
    ti.set(docx_qn('w:type'), 'dxa')

    # Fill cells
    for ri, row_data in enumerate(table_data):
        col_idx = 0
        for cell_data in row_data:
            if col_idx >= ncols:
                break
            if cell_data.get('vmerge_continue'):
                col_idx += cell_data.get('gridspan', 1)
                continue

            cell = table.cell(ri, col_idx)
            for p in cell.paragraphs:
                p.clear()
            para = cell.paragraphs[0]
            try:
                para.style = doc.styles['MDPI_4.2_table_body']
            except KeyError:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Header row bottom border
            if ri == 0:
                tc_pr = cell._tc.find(docx_qn('w:tcPr'))
                if tc_pr is None:
                    tc_pr = lxml_etree.SubElement(cell._tc, docx_qn('w:tcPr'))
                tc_borders = lxml_etree.SubElement(tc_pr, docx_qn('w:tcBorders'))
                bb = lxml_etree.SubElement(tc_borders, docx_qn('w:bottom'))
                bb.set(docx_qn('w:val'), 'single')
                bb.set(docx_qn('w:sz'), '4')
                bb.set(docx_qn('w:space'), '0')
                bb.set(docx_qn('w:color'), 'auto')

            if cell_data.get('runs'):
                for rd in cell_data['runs']:
                    if not rd['text']:
                        continue
                    run = para.add_run(rd['text'])
                    b = rd['bold'] or (ri == 0)
                    _set_run_font(run, size=DEFAULT_SIZE, bold=b,
                                  italic=rd['italic'],
                                  superscript=rd.get('superscript', False),
                                  subscript=rd.get('subscript', False))
            else:
                run = para.add_run(cell_data.get('text', '').strip())
                _set_run_font(run, size=DEFAULT_SIZE, bold=(ri == 0))

            gs = cell_data.get('gridspan', 1)
            if gs > 1 and col_idx + gs <= ncols:
                cell.merge(table.cell(ri, col_idx + gs - 1))
            col_idx += gs

    return table


def build(items, output_path, ris_data=None, zotero_enabled=False,
          use_crossref=False, progress_callback=None):
    """Build MDPI-formatted document from reader items.

    Uses a styles template extracted from a published MDPI paper,
    so all named styles match editor-approved formatting exactly.
    """
    # Load from styles template — all 39 MDPI styles are pre-loaded
    if os.path.isfile(_TEMPLATE_PATH):
        doc = DocxDocument(_TEMPLATE_PATH)
    else:
        doc = DocxDocument()

    # Page setup
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.50)
    section.bottom_margin = Cm(1.60)
    section.left_margin = Cm(1.27)
    section.right_margin = Cm(1.27)

    # Article type
    para = _styled_para(doc, 'MDPI_1.1_article_type')
    run = para.add_run('Article')
    _set_run_font(run, size=DEFAULT_SIZE, italic=True)

    # Find title
    title_elem = None
    for elem in items:
        if elem['type'] in ('abstract_heading', 'heading1'):
            break
        if elem['type'] == 'paragraph':
            if not title_elem:
                title_elem = elem

    # Title
    title_text = title_elem['text'] if title_elem else '[Title]'
    para = _styled_para(doc, 'MDPI_1.2_title')
    run = para.add_run(title_text)
    _set_run_font(run, size=Pt(18), bold=True)

    # Author placeholder
    _build_author_placeholder(doc)

    # Resolve references
    ref_items = [item for item in items if item['type'] == 'reference']
    resolved_refs = resolve_references(
        ref_items, ris_data=ris_data, use_crossref=use_crossref,
        progress_callback=progress_callback)
    ref_index = 0

    # Main content
    after_heading = False
    for item in items:
        itype = item['type']

        if item is title_elem:
            continue
        if itype == 'abstract_heading':
            continue

        if itype == 'abstract_text':
            _add_runs_para(doc, 'MDPI_1.7_abstract', item['runs'],
                           bold_prefix='Abstract: ')
            after_heading = False
            continue

        if itype == 'keywords':
            kw_text = item['text'].replace('Keywords:', '').replace(
                'Keywords :', '').strip()
            _add_text_para(doc, 'MDPI_1.8_keywords', kw_text,
                           bold_prefix='Keywords: ')
            after_heading = False
            continue

        if itype == 'heading1':
            para = _styled_para(doc, 'MDPI_2.1_heading1')
            run = para.add_run(item['text'])
            _set_run_font(run, bold=True)
            after_heading = True
            continue

        if itype == 'heading2':
            para = _styled_para(doc, 'MDPI_2.2_heading2')
            run = para.add_run(item['text'])
            _set_run_font(run, italic=True)
            after_heading = True
            continue

        if itype == 'heading3':
            para = _styled_para(doc, 'MDPI_2.3_heading3')
            run = para.add_run(item['text'])
            _set_run_font(run, size=DEFAULT_SIZE)
            after_heading = True
            continue

        if itype == 'references_heading':
            para = _styled_para(doc, 'MDPI_2.1_heading1')
            run = para.add_run(item['text'])
            _set_run_font(run, bold=True)
            after_heading = True
            continue

        if itype == 'paragraph':
            style = 'MDPI_3.2_text_no_indent' if after_heading else 'MDPI_3.1_text'
            _add_runs_para(doc, style, item['runs'])
            after_heading = False
            continue

        if itype == 'equation':
            _add_runs_para(doc, 'MDPI_3.9_equation', item['runs'])
            after_heading = False
            continue

        if itype == 'table_caption':
            text = item['text']
            match = re.match(r'^Table\s+(\d+)\.?\s*(.*)', text)
            if match:
                para = _styled_para(doc, 'MDPI_4.1_table_caption')
                add_caption_with_seq(para, 'Table', match.group(1),
                                     description=match.group(2),
                                     font_name=FONT_NAME, font_size=Pt(9),
                                     font_color=FONT_COLOR, bold_label=True)
            else:
                _add_text_para(doc, 'MDPI_4.1_table_caption', text, font_size=Pt(9))
            after_heading = False
            continue

        if itype == 'table':
            _add_three_line_table(doc, item['rows'])
            after_heading = False
            continue

        if itype == 'table_footer':
            _add_text_para(doc, 'MDPI_4.3_table_footer', item['text'], font_size=Pt(9))
            after_heading = False
            continue

        if itype == 'figure_placeholder':
            text = item['text']
            match = re.match(r'^\[?\s*Figure\s+(\d+)\]?\.?\s*(.*)', text, re.IGNORECASE)
            if match:
                para = _styled_para(doc, 'MDPI_5.1_figure_caption')
                add_caption_with_seq(para, 'Figure', match.group(1),
                                     description=match.group(2),
                                     font_name=FONT_NAME, font_size=Pt(9),
                                     font_color=FONT_COLOR, bold_label=True)
            else:
                _add_runs_para(doc, 'MDPI_5.1_figure_caption', item['runs'],
                               font_size=Pt(9))
            after_heading = False
            continue

        if itype == 'reference':
            if ref_index < len(resolved_refs):
                rr = resolved_refs[ref_index]
                ref_index += 1

                para = _styled_para(doc, 'MDPI_8.1_references')

                if zotero_enabled and rr.metadata:
                    csl_json = ris_to_csl_json(rr.metadata, rr.index)
                    visible = ''.join(r['text'] for r in rr.formatted_runs)
                    add_zotero_citation_field(para, csl_json, visible,
                                              font_name=FONT_NAME,
                                              font_size=DEFAULT_SIZE)
                else:
                    for rd in rr.formatted_runs:
                        run = para.add_run(rd['text'])
                        _set_run_font(run, size=DEFAULT_SIZE,
                                      bold=rd.get('bold'),
                                      italic=rd.get('italic'),
                                      superscript=rd.get('superscript', False),
                                      subscript=rd.get('subscript', False))
            after_heading = False
            continue

    doc.save(output_path)
    return output_path
