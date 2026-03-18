"""
MDPI journal format plugin.

Builds a formatted .docx from scratch using python-docx.
All MDPI styles are defined programmatically — no template file required.
"""

import re
from docx import Document as DocxDocument
from docx.oxml.ns import qn as docx_qn
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from lxml import etree as lxml_etree
from caption_fields import add_caption_with_seq, add_zotero_citation_field, ris_to_csl_json
from ris_parser import match_citation_to_ris
from citation_formatter import format_reference_mdpi

FORMAT_NAME = 'MDPI'
FORMAT_SUFFIX = '_MDPI'

FONT_NAME = 'Palatino Linotype'
FONT_COLOR = RGBColor(0x00, 0x00, 0x00)
DEFAULT_SIZE = Pt(10)
LEFT_INDENT = Pt(130.4)

TABLE_WIDTH = 7857
TABLE_INDENT = 2608

# Style definitions: (font_size, bold, italic, alignment, space_before,
#   space_after, line_spacing_pt, left_indent, first_line_indent, outline_level)
STYLES = {
    'articletype':   (Pt(10),  False, True,  WD_ALIGN_PARAGRAPH.LEFT,    Pt(12), None,   None,   None,        None,      None),
    'title':         (Pt(18),  True,  False, WD_ALIGN_PARAGRAPH.LEFT,    None,   Pt(12), Pt(12), None,        None,      None),
    'authornames':   (Pt(10),  True,  False, WD_ALIGN_PARAGRAPH.LEFT,    None,   Pt(18), Pt(13), None,        None,      None),
    'abstract':      (Pt(10),  False, False, WD_ALIGN_PARAGRAPH.JUSTIFY, Pt(12), Pt(6),  Pt(14), LEFT_INDENT, None,      None),
    'keywords':      (Pt(10),  False, False, WD_ALIGN_PARAGRAPH.JUSTIFY, Pt(12), None,   Pt(14), LEFT_INDENT, None,      None),
    'heading1':      (Pt(12),  True,  False, WD_ALIGN_PARAGRAPH.LEFT,    Pt(12), Pt(3),  Pt(14), LEFT_INDENT, None,      0),
    'heading2':      (Pt(10),  False, True,  WD_ALIGN_PARAGRAPH.LEFT,    Pt(3),  Pt(3),  Pt(14), LEFT_INDENT, None,      1),
    'heading3':      (Pt(10),  False, False, WD_ALIGN_PARAGRAPH.LEFT,    Pt(3),  Pt(3),  Pt(14), LEFT_INDENT, None,      2),
    'text':          (Pt(10),  False, False, WD_ALIGN_PARAGRAPH.JUSTIFY, None,   None,   Pt(14), LEFT_INDENT, Pt(21.25), None),
    'textnoindent':  (Pt(10),  False, False, WD_ALIGN_PARAGRAPH.JUSTIFY, None,   None,   Pt(14), LEFT_INDENT, Pt(0),     None),
    'tablecaption':  (Pt(9),   False, False, WD_ALIGN_PARAGRAPH.LEFT,    Pt(12), Pt(6),  Pt(14), LEFT_INDENT, None,      None),
    'tablebody':     (Pt(10),  False, False, WD_ALIGN_PARAGRAPH.CENTER,  None,   None,   Pt(13), None,        None,      None),
    'tablefooter':   (Pt(9),   False, False, WD_ALIGN_PARAGRAPH.JUSTIFY, None,   None,   Pt(14), LEFT_INDENT, None,      None),
    'figurecaption': (Pt(9),   False, False, WD_ALIGN_PARAGRAPH.JUSTIFY, Pt(6),  Pt(12), Pt(14), LEFT_INDENT, None,      None),
}


def _set_run_font(run, size=None, bold=None, italic=None,
                  superscript=False, subscript=False):
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


def _apply_style(para, style_key):
    s = STYLES[style_key]
    font_size, bold, italic, alignment, space_before, space_after, \
        line_spacing_pt, left_indent, first_line_indent, outline_level = s

    para.alignment = alignment
    pf = para.paragraph_format

    if space_before is not None:
        pf.space_before = space_before
    if space_after is not None:
        pf.space_after = space_after
    if line_spacing_pt is not None:
        pf.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
        pf.line_spacing = line_spacing_pt
    if left_indent is not None:
        pf.left_indent = left_indent
    if first_line_indent is not None:
        pf.first_line_indent = first_line_indent

    if outline_level is not None:
        ppr = para._p.find(docx_qn('w:pPr'))
        if ppr is None:
            ppr = lxml_etree.SubElement(para._p, docx_qn('w:pPr'))
        ol = lxml_etree.SubElement(ppr, docx_qn('w:outlineLvl'))
        ol.set(docx_qn('w:val'), str(outline_level))

    return font_size, bold, italic


def _add_para(doc, style_key, text='', bold_prefix=None):
    para = doc.add_paragraph()
    font_size, bold, italic = _apply_style(para, style_key)
    if bold_prefix:
        run = para.add_run(bold_prefix)
        _set_run_font(run, size=font_size, bold=True)
    if text:
        run = para.add_run(text)
        _set_run_font(run, size=font_size, bold=bold, italic=italic)
    return para


def _add_runs_para(doc, style_key, runs_data, bold_prefix=None,
                   size_override=None):
    para = doc.add_paragraph()
    font_size, _, _ = _apply_style(para, style_key)
    size = size_override or font_size
    if bold_prefix:
        run = para.add_run(bold_prefix)
        _set_run_font(run, size=size, bold=True)
    for rd in runs_data:
        if not rd['text']:
            continue
        run = para.add_run(rd['text'])
        _set_run_font(run, size=size, bold=rd['bold'], italic=rd['italic'],
                      superscript=rd.get('superscript', False),
                      subscript=rd.get('subscript', False))
    return para


def _build_author_placeholder(doc):
    para = doc.add_paragraph()
    _apply_style(para, 'authornames')
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
    if not table_data or not table_data[0]:
        return None
    ncols = sum(c.get('gridspan', 1) for c in table_data[0])
    if ncols == 0:
        return None
    nrows = len(table_data)
    table = doc.add_table(rows=nrows, cols=ncols)

    tbl_pr = table._tbl.find(docx_qn('w:tblPr'))
    if tbl_pr is None:
        tbl_pr = lxml_etree.SubElement(table._tbl, docx_qn('w:tblPr'))

    tw = lxml_etree.SubElement(tbl_pr, docx_qn('w:tblW'))
    tw.set(docx_qn('w:w'), str(TABLE_WIDTH))
    tw.set(docx_qn('w:type'), 'dxa')
    ti = lxml_etree.SubElement(tbl_pr, docx_qn('w:tblInd'))
    ti.set(docx_qn('w:w'), str(TABLE_INDENT))
    ti.set(docx_qn('w:type'), 'dxa')

    borders = lxml_etree.SubElement(tbl_pr, docx_qn('w:tblBorders'))
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        el = lxml_etree.SubElement(borders, docx_qn(f'w:{border_name}'))
        el.set(docx_qn('w:val'), 'none')
        el.set(docx_qn('w:sz'), '0')
        el.set(docx_qn('w:space'), '0')
        el.set(docx_qn('w:color'), 'auto')

    top_border = borders.find(docx_qn('w:top'))
    top_border.set(docx_qn('w:val'), 'single')
    top_border.set(docx_qn('w:sz'), '8')

    bottom_border = borders.find(docx_qn('w:bottom'))
    bottom_border.set(docx_qn('w:val'), 'single')
    bottom_border.set(docx_qn('w:sz'), '8')

    tl = lxml_etree.SubElement(tbl_pr, docx_qn('w:tblLayout'))
    tl.set(docx_qn('w:type'), 'fixed')
    tcm = lxml_etree.SubElement(tbl_pr, docx_qn('w:tblCellMar'))
    for side in ['left', 'right']:
        s = lxml_etree.SubElement(tcm, docx_qn(f'w:{side}'))
        s.set(docx_qn('w:w'), '0')
        s.set(docx_qn('w:type'), 'dxa')

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
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf = para.paragraph_format
            pf.line_spacing = 1.0

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


def build(items, output_path, ris_data=None, zotero_enabled=False):
    doc = DocxDocument()

    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.50)
    section.bottom_margin = Cm(1.60)
    section.left_margin = Cm(1.27)
    section.right_margin = Cm(1.27)

    style = doc.styles['Normal']
    style.font.name = FONT_NAME
    style.font.size = DEFAULT_SIZE
    style.font.color.rgb = FONT_COLOR

    _add_para(doc, 'articletype', text='Article')

    title_elem = None
    for elem in items:
        if elem['type'] in ('abstract_heading', 'heading1'):
            break
        if elem['type'] == 'paragraph':
            if not title_elem:
                title_elem = elem

    title_text = title_elem['text'] if title_elem else '[Title]'
    _add_para(doc, 'title', text=title_text)
    _build_author_placeholder(doc)

    after_heading = False

    for item in items:
        itype = item['type']

        if item is title_elem:
            continue
        if itype == 'abstract_heading':
            continue

        if itype == 'abstract_text':
            _add_runs_para(doc, 'abstract', item['runs'],
                           bold_prefix='Abstract: ')
            after_heading = False
            continue
        if itype == 'keywords':
            kw_text = item['text'].replace('Keywords:', '').replace(
                'Keywords :', '').strip()
            _add_para(doc, 'keywords', kw_text, bold_prefix='Keywords: ')
            after_heading = False
            continue
        if itype == 'heading1':
            _add_para(doc, 'heading1', text=item['text'])
            after_heading = True
            continue
        if itype == 'heading2':
            _add_para(doc, 'heading2', text=item['text'])
            after_heading = True
            continue
        if itype == 'heading3':
            _add_para(doc, 'heading3', text=item['text'])
            after_heading = True
            continue
        if itype == 'references_heading':
            _add_para(doc, 'heading1', text=item['text'])
            after_heading = True
            continue
        if itype == 'paragraph':
            style_key = 'textnoindent' if after_heading else 'text'
            _add_runs_para(doc, style_key, item['runs'])
            after_heading = False
            continue
        if itype == 'equation':
            para = _add_runs_para(doc, 'textnoindent', item['runs'])
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            after_heading = False
            continue
        if itype == 'table_caption':
            text = item['text']
            match = re.match(r'^Table\s+(\d+)\.?\s*(.*)', text)
            if match:
                num = match.group(1)
                desc = match.group(2)
                para = doc.add_paragraph()
                _apply_style(para, 'tablecaption')
                add_caption_with_seq(para, 'Table', num, description=desc,
                                     font_name=FONT_NAME, font_size=Pt(9),
                                     font_color=FONT_COLOR, bold_label=True)
            else:
                _add_para(doc, 'tablecaption', text)
            after_heading = False
            continue
        if itype == 'table':
            _add_three_line_table(doc, item['rows'])
            after_heading = False
            continue
        if itype == 'table_footer':
            _add_para(doc, 'tablefooter', text=item['text'])
            after_heading = False
            continue
        if itype == 'figure_placeholder':
            text = item['text']
            match = re.match(r'^\[?\s*Figure\s+(\d+)\]?\.?\s*(.*)', text, re.IGNORECASE)
            if match:
                num = match.group(1)
                desc = match.group(2)
                para = doc.add_paragraph()
                _apply_style(para, 'figurecaption')
                add_caption_with_seq(para, 'Figure', num, description=desc,
                                     font_name=FONT_NAME, font_size=Pt(9),
                                     font_color=FONT_COLOR, bold_label=True)
            else:
                _add_runs_para(doc, 'figurecaption', item['runs'])
            after_heading = False
            continue
        if itype == 'reference':
            ref_text = item['text']
            matched_ris = None
            if ris_data:
                matched_ris = match_citation_to_ris(ref_text, ris_data)

            if matched_ris:
                # Reformat using RIS metadata in MDPI style
                formatted = format_reference_mdpi(matched_ris)
                # Extract reference number from original
                num_match = re.match(r'^\[?(\d+)\]?\s*', ref_text)
                if num_match:
                    formatted = f'{num_match.group(0)}{formatted}'

                para = doc.add_paragraph()
                _apply_style(para, 'textnoindent')

                if zotero_enabled:
                    csl_json = ris_to_csl_json(matched_ris,
                                                int(num_match.group(1)) if num_match else 1)
                    add_zotero_citation_field(para, csl_json, formatted,
                                              font_name=FONT_NAME, font_size=Pt(9))
                else:
                    run = para.add_run(formatted)
                    _set_run_font(run, size=Pt(9))
            else:
                # No RIS match — preserve original text
                _add_runs_para(doc, 'textnoindent', item['runs'],
                               size_override=Pt(9))
            after_heading = False
            continue

    doc.save(output_path)
    return output_path
