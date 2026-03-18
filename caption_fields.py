"""
Word caption field helpers.

Creates proper Word SEQ field-based captions with bookmarks that Word
recognizes for cross-references, table of figures, and auto-numbering.
Also provides Zotero field code embedding and CSL JSON conversion.
"""

import json
import uuid
from docx.oxml.ns import qn as docx_qn
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from lxml import etree as lxml_etree

_bookmark_id_counter = 0


def _next_bookmark_id():
    global _bookmark_id_counter
    _bookmark_id_counter += 1
    return _bookmark_id_counter


def register_mdpi_styles(doc):
    """Register MDPI named paragraph styles in the document.

    Creates actual Word styles so captions and references are
    discoverable by Word's Table of Figures and style-based features.
    """
    styles = doc.styles
    left_indent = Pt(130.4)
    font_name = 'Palatino Linotype'
    font_color = RGBColor(0x00, 0x00, 0x00)

    def _make_style(name, font_size, bold=False, italic=False,
                    alignment=WD_ALIGN_PARAGRAPH.LEFT,
                    space_before=None, space_after=None,
                    line_spacing_pt=None, l_indent=None,
                    first_indent=None):
        if name in [s.name for s in styles]:
            return styles[name]
        st = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        st.font.name = font_name
        st.font.size = font_size
        st.font.color.rgb = font_color
        if bold:
            st.font.bold = True
        if italic:
            st.font.italic = True
        pf = st.paragraph_format
        pf.alignment = alignment
        if space_before is not None:
            pf.space_before = space_before
        if space_after is not None:
            pf.space_after = space_after
        if line_spacing_pt is not None:
            pf.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
            pf.line_spacing = line_spacing_pt
        if l_indent is not None:
            pf.left_indent = l_indent
        if first_indent is not None:
            pf.first_line_indent = first_indent
        return st

    _make_style('MDPI41tablecaption', Pt(9),
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
                space_before=Pt(12), space_after=Pt(6),
                line_spacing_pt=Pt(14), l_indent=left_indent)

    _make_style('MDPI51figurecaption', Pt(9),
                alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                space_before=Pt(6), space_after=Pt(12),
                line_spacing_pt=Pt(14), l_indent=left_indent)

    _make_style('MDPI43tablefooter', Pt(9),
                alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                line_spacing_pt=Pt(14), l_indent=left_indent)

    _make_style('MDPIBibliography', Pt(10),
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
                line_spacing_pt=Pt(12),
                l_indent=Pt(25.2), first_indent=Pt(-25.2))


def register_elsevier_caption_styles(doc):
    """Register Elsevier caption styles in the document."""
    styles = doc.styles
    font_name = 'Times New Roman'

    def _make_style(name, font_size, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                    space_before=None, space_after=None):
        if name in [s.name for s in styles]:
            return styles[name]
        st = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        st.font.name = font_name
        st.font.size = font_size
        pf = st.paragraph_format
        pf.alignment = alignment
        pf.line_spacing = 1.5
        if space_before is not None:
            pf.space_before = space_before
        if space_after is not None:
            pf.space_after = space_after
        return st

    _make_style('ElsevierTableCaption', Pt(10),
                space_before=Pt(12), space_after=Pt(4))
    _make_style('ElsevierFigureCaption', Pt(10),
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
                space_before=Pt(12), space_after=Pt(12))


def add_caption_with_seq(para, seq_name, number, description='',
                         font_name='Palatino Linotype', font_size=None,
                         font_color=None, bold_label=True):
    """Add a Word-recognizable caption with SEQ field and bookmark.

    Creates: bookmarkStart + "Label " + SEQ N + bookmarkEnd + ". " + description
    Word can cross-reference via the bookmark and find captions via SEQ field.
    """
    p = para._p
    bm_id = _next_bookmark_id()
    bm_name = f'_Ref{seq_name}{number}_{uuid.uuid4().hex[:6]}'

    def _rpr(bold=False):
        rpr = lxml_etree.Element(docx_qn('w:rPr'))
        rfonts = lxml_etree.SubElement(rpr, docx_qn('w:rFonts'))
        rfonts.set(docx_qn('w:ascii'), font_name)
        rfonts.set(docx_qn('w:hAnsi'), font_name)
        if font_size:
            sz = lxml_etree.SubElement(rpr, docx_qn('w:sz'))
            sz.set(docx_qn('w:val'), str(int(font_size.pt * 2)))
            szcs = lxml_etree.SubElement(rpr, docx_qn('w:szCs'))
            szcs.set(docx_qn('w:val'), str(int(font_size.pt * 2)))
        if font_color:
            c = lxml_etree.SubElement(rpr, docx_qn('w:color'))
            c.set(docx_qn('w:val'), str(font_color))
        if bold:
            lxml_etree.SubElement(rpr, docx_qn('w:b'))
        return rpr

    def make_run(text, bold=False, no_proof=False):
        r = lxml_etree.SubElement(p, docx_qn('w:r'))
        rpr = _rpr(bold)
        if no_proof:
            lxml_etree.SubElement(rpr, docx_qn('w:noProof'))
        r.append(rpr)
        t = lxml_etree.SubElement(r, docx_qn('w:t'))
        t.text = text
        if text and (text[0] == ' ' or text[-1] == ' '):
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        return r

    def make_fld_char(fld_type, bold=False):
        r = lxml_etree.SubElement(p, docx_qn('w:r'))
        r.append(_rpr(bold))
        fc = lxml_etree.SubElement(r, docx_qn('w:fldChar'))
        fc.set(docx_qn('w:fldCharType'), fld_type)
        return r

    def make_instr_text(instruction, bold=False):
        r = lxml_etree.SubElement(p, docx_qn('w:r'))
        r.append(_rpr(bold))
        it = lxml_etree.SubElement(r, docx_qn('w:instrText'))
        it.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        it.text = instruction
        return r

    # bookmarkStart
    bms = lxml_etree.SubElement(p, docx_qn('w:bookmarkStart'))
    bms.set(docx_qn('w:id'), str(bm_id))
    bms.set(docx_qn('w:name'), bm_name)

    # "Table " or "Figure " label (bold)
    make_run(f'{seq_name} ', bold=bold_label)

    # SEQ field
    make_fld_char('begin', bold=bold_label)
    make_instr_text(f' SEQ {seq_name} \\* ARABIC ', bold=bold_label)
    make_fld_char('separate', bold=bold_label)
    make_run(str(number), bold=bold_label, no_proof=True)
    make_fld_char('end', bold=bold_label)

    # bookmarkEnd
    bme = lxml_etree.SubElement(p, docx_qn('w:bookmarkEnd'))
    bme.set(docx_qn('w:id'), str(bm_id))

    # ". " after number
    make_run('. ', bold=bold_label)

    # Description text (not bold)
    if description:
        make_run(description, bold=False)

    return bm_name


def add_zotero_citation_field(para, citation_json, visible_text,
                               font_name='Palatino Linotype', font_size=None):
    """Add a Zotero-compatible citation field code to a paragraph.

    All runs include w:rPr for Word/Zotero compatibility, and the begin
    fldChar includes w:dirty="true" so Word processes the field on open.
    """
    p = para._p

    def _add_rpr(r):
        rpr = lxml_etree.SubElement(r, docx_qn('w:rPr'))
        rfonts = lxml_etree.SubElement(rpr, docx_qn('w:rFonts'))
        rfonts.set(docx_qn('w:ascii'), font_name)
        rfonts.set(docx_qn('w:hAnsi'), font_name)
        if font_size:
            sz = lxml_etree.SubElement(rpr, docx_qn('w:sz'))
            sz.set(docx_qn('w:val'), str(int(font_size.pt * 2)))
            szcs = lxml_etree.SubElement(rpr, docx_qn('w:szCs'))
            szcs.set(docx_qn('w:val'), str(int(font_size.pt * 2)))
        return rpr

    def make_run(text):
        r = lxml_etree.SubElement(p, docx_qn('w:r'))
        _add_rpr(r)
        t = lxml_etree.SubElement(r, docx_qn('w:t'))
        t.text = text
        if text and (text[0] == ' ' or text[-1] == ' '):
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    def make_fld_char(fld_type, dirty=False):
        r = lxml_etree.SubElement(p, docx_qn('w:r'))
        _add_rpr(r)
        fc = lxml_etree.SubElement(r, docx_qn('w:fldChar'))
        fc.set(docx_qn('w:fldCharType'), fld_type)
        if dirty:
            fc.set(docx_qn('w:dirty'), 'true')

    def make_instr_text(instruction):
        r = lxml_etree.SubElement(p, docx_qn('w:r'))
        _add_rpr(r)
        it = lxml_etree.SubElement(r, docx_qn('w:instrText'))
        it.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        it.text = instruction

    make_fld_char('begin', dirty=True)
    make_instr_text(f' ADDIN ZOTERO_ITEM CSL_CITATION {citation_json} ')
    make_fld_char('separate')
    make_run(visible_text)
    make_fld_char('end')


def ris_to_csl_json(ref, index):
    """Convert a RIS reference dict to Zotero CSL JSON citation format."""
    type_map = {
        'JOUR': 'article-journal', 'BOOK': 'book', 'CHAP': 'chapter',
        'CONF': 'paper-conference', 'THES': 'thesis', 'RPRT': 'report',
    }

    authors_csl = []
    for a in ref.get('authors', []):
        if ',' in a:
            parts = a.split(',', 1)
            authors_csl.append({'family': parts[0].strip(), 'given': parts[1].strip()})
        else:
            parts = a.rsplit(' ', 1)
            if len(parts) == 2:
                authors_csl.append({'family': parts[1].strip(), 'given': parts[0].strip()})
            else:
                authors_csl.append({'family': a, 'given': ''})

    item_data = {
        'id': index,
        'type': type_map.get(ref.get('type', ''), 'article'),
        'title': ref.get('title', ''),
        'author': authors_csl,
    }
    if ref.get('journal'):
        item_data['container-title'] = ref['journal']
    if ref.get('volume'):
        item_data['volume'] = ref['volume']
    if ref.get('issue'):
        item_data['issue'] = ref['issue']
    if ref.get('start_page'):
        pages = ref['start_page']
        if ref.get('end_page'):
            pages += '-' + ref['end_page']
        item_data['page'] = pages
    if ref.get('doi'):
        item_data['DOI'] = ref['doi']
    if ref.get('year'):
        item_data['issued'] = {'date-parts': [[ref['year']]]}
    if ref.get('publisher'):
        item_data['publisher'] = ref['publisher']
    if ref.get('place'):
        item_data['publisher-place'] = ref['place']

    citation = {
        'citationID': uuid.uuid4().hex[:8],
        'properties': {
            'formattedCitation': f'[{index}]',
            'plainCitation': f'[{index}]',
            'noteIndex': 0,
        },
        'citationItems': [{'id': index, 'itemData': item_data}],
        'schema': 'https://github.com/citation-style-language/schema/raw/master/csl-citation.json',
    }
    return json.dumps(citation)
