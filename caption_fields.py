"""
Word caption field helpers.

Creates proper Word SEQ field-based captions that Word recognizes
for cross-references, table of figures, and auto-numbering.
"""

from docx.oxml.ns import qn as docx_qn
from lxml import etree as lxml_etree


def add_caption_with_seq(para, seq_name, number, description='',
                         font_name='Palatino Linotype', font_size=None,
                         font_color=None, bold_label=True):
    """Add a Word-recognizable caption with SEQ field to a paragraph.

    Creates: "<seq_name> N. description" where N is a SEQ field.

    Args:
        para: python-docx Paragraph object
        seq_name: 'Table' or 'Figure'
        number: The caption number (int or str)
        description: Caption text after "Table N. "
        font_name: Font for caption text
        font_size: docx.shared.Pt value
        font_color: docx.shared.RGBColor value
        bold_label: Whether "Table N." is bold
    """
    p = para._p

    def make_run(text, bold=False, italic=False):
        """Create a w:r element with text and formatting."""
        r = lxml_etree.SubElement(p, docx_qn('w:r'))
        rpr = lxml_etree.SubElement(r, docx_qn('w:rPr'))
        # Font
        rfonts = lxml_etree.SubElement(rpr, docx_qn('w:rFonts'))
        rfonts.set(docx_qn('w:ascii'), font_name)
        rfonts.set(docx_qn('w:hAnsi'), font_name)
        if font_size:
            sz = lxml_etree.SubElement(rpr, docx_qn('w:sz'))
            sz.set(docx_qn('w:val'), str(int(font_size.pt * 2)))
            szcs = lxml_etree.SubElement(rpr, docx_qn('w:szCs'))
            szcs.set(docx_qn('w:val'), str(int(font_size.pt * 2)))
        if font_color:
            color_el = lxml_etree.SubElement(rpr, docx_qn('w:color'))
            color_el.set(docx_qn('w:val'), str(font_color))
        if bold:
            lxml_etree.SubElement(rpr, docx_qn('w:b'))
        if italic:
            lxml_etree.SubElement(rpr, docx_qn('w:i'))
        t = lxml_etree.SubElement(r, docx_qn('w:t'))
        t.text = text
        if text and (text[0] == ' ' or text[-1] == ' '):
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        return r

    def make_fld_char(fld_type):
        """Create a w:fldChar element (begin/separate/end)."""
        r = lxml_etree.SubElement(p, docx_qn('w:r'))
        rpr = lxml_etree.SubElement(r, docx_qn('w:rPr'))
        if font_size:
            sz = lxml_etree.SubElement(rpr, docx_qn('w:sz'))
            sz.set(docx_qn('w:val'), str(int(font_size.pt * 2)))
        fc = lxml_etree.SubElement(r, docx_qn('w:fldChar'))
        fc.set(docx_qn('w:fldCharType'), fld_type)
        return r

    def make_instr_text(instruction):
        """Create a w:instrText element."""
        r = lxml_etree.SubElement(p, docx_qn('w:r'))
        rpr = lxml_etree.SubElement(r, docx_qn('w:rPr'))
        if font_size:
            sz = lxml_etree.SubElement(rpr, docx_qn('w:sz'))
            sz.set(docx_qn('w:val'), str(int(font_size.pt * 2)))
        it = lxml_etree.SubElement(r, docx_qn('w:instrText'))
        it.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        it.text = instruction
        return r

    # "Table " or "Figure " label (bold)
    make_run(f'{seq_name} ', bold=bold_label)

    # SEQ field: begin
    make_fld_char('begin')
    # SEQ instruction
    make_instr_text(f' SEQ {seq_name} \\* ARABIC ')
    # SEQ field: separate
    make_fld_char('separate')
    # Visible number (placeholder until Word updates fields)
    make_run(str(number), bold=bold_label)
    # SEQ field: end
    make_fld_char('end')

    # Period and space after number
    make_run('. ', bold=bold_label)

    # Description text (not bold)
    if description:
        make_run(description, bold=False)


def add_zotero_citation_field(para, citation_json, visible_text,
                               font_name='Palatino Linotype', font_size=None):
    """Add a Zotero-compatible citation field code to a paragraph.

    Creates ADDIN ZOTERO_ITEM CSL_CITATION field with the JSON payload.
    All runs include w:rPr for Word/Zotero compatibility, and the begin
    fldChar includes w:dirty="true" so Word processes the field on open.

    Args:
        para: python-docx Paragraph object
        citation_json: JSON string with CSL citation data
        visible_text: The visible citation text (e.g., '[1]')
        font_name: Font name
        font_size: docx.shared.Pt value
    """
    p = para._p

    def _add_rpr(r):
        """Add standard run properties to a run element."""
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
    """Convert a RIS reference dict to Zotero CSL JSON citation format.

    Args:
        ref: dict from ris_parser
        index: citation number (1-based)

    Returns:
        JSON string for use in Zotero field codes
    """
    import json
    import uuid

    type_map = {
        'JOUR': 'article-journal',
        'BOOK': 'book',
        'CHAP': 'chapter',
        'CONF': 'paper-conference',
        'THES': 'thesis',
        'RPRT': 'report',
    }

    authors_csl = []
    for a in ref.get('authors', []):
        if ',' in a:
            parts = a.split(',', 1)
            authors_csl.append({
                'family': parts[0].strip(),
                'given': parts[1].strip(),
            })
        else:
            parts = a.rsplit(' ', 1)
            if len(parts) == 2:
                authors_csl.append({
                    'family': parts[1].strip(),
                    'given': parts[0].strip(),
                })
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
        'citationItems': [{
            'id': index,
            'itemData': item_data,
        }],
        'schema': 'https://github.com/citation-style-language/schema/raw/master/csl-citation.json',
    }

    return json.dumps(citation)
