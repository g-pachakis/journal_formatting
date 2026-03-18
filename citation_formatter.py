"""
Citation formatter.

Formats reference data (from RIS records or CrossRef) into journal-specific
citation styles. Returns structured runs with formatting flags so journal
titles can be italic.

MDPI reference format:
    N.<TAB>Author1, F.M.; Author2, F.M. Title of Article. *Journal* Year, Vol, Pages. DOI
"""

import re


def format_author_mdpi(author_str):
    """Format a single author name for MDPI style.
    Input: 'Shannon, Claude E.' or 'Claude E. Shannon'
    Output: 'Shannon, C.E.'
    """
    author_str = author_str.strip()
    if not author_str:
        return ''
    if ',' in author_str:
        parts = author_str.split(',', 1)
        last = parts[0].strip()
        given = parts[1].strip()
    else:
        parts = author_str.rsplit(' ', 1)
        if len(parts) == 2:
            given = parts[0].strip()
            last = parts[1].strip()
        else:
            return author_str
    initials = []
    for name in given.split():
        name = name.strip().rstrip('.')
        if name:
            initials.append(name[0].upper() + '.')
    return f"{last}, {''.join(initials)}"


def format_reference_mdpi_runs(ref, index):
    """Format a reference dict into MDPI numbered style as structured runs.

    Returns list of run dicts: [{text, bold, italic, superscript, subscript}, ...]
    These can be directly consumed by python-docx paragraph.add_run().

    MDPI format:
        N.\\tLastname, F.M.; Lastname, F.M. Title. *Journal* Year, Vol, Pages. DOI
    """
    ref_type = ref.get('type', 'JOUR')
    runs = []

    def _run(text, bold=False, italic=False):
        return {'text': text, 'bold': bold, 'italic': italic,
                'superscript': False, 'subscript': False}

    # Number + tab
    runs.append(_run(f'{index}.\t'))

    # Authors
    authors = ref.get('authors', [])
    if authors:
        formatted = [format_author_mdpi(a) for a in authors]
        runs.append(_run('; '.join(formatted) + '. '))

    # Title (not italic, not in quotes)
    title = ref.get('title', '')
    if title:
        title = title.rstrip('.')
        runs.append(_run(title + '. '))

    if ref_type in ('JOUR', 'EJOUR', 'JFULL'):
        journal = ref.get('journal', '')
        year = ref.get('year', '')
        volume = ref.get('volume', '')
        sp = ref.get('start_page', '')
        ep = ref.get('end_page', '')

        if journal:
            # Journal name is ITALIC
            runs.append(_run(journal.rstrip('.'), italic=True))

            detail = ''
            if year:
                detail += f' {year}'
            if volume:
                detail += f', {volume}'
            if sp:
                pages = sp
                if ep and ep != sp:
                    pages += f'\u2013{ep}'
                detail += f', {pages}'
            detail += '.'
            runs.append(_run(detail))

    elif ref_type in ('BOOK', 'EBOOK'):
        # Book title was already added above (as the title)
        edition = ref.get('edition', '')
        publisher = ref.get('publisher', '')
        place = ref.get('place', '')
        year = ref.get('year', '')

        detail = ''
        if edition:
            detail += f', {edition}'
        if publisher or place:
            detail += '; '
            if publisher:
                detail += publisher
            if place:
                detail += f': {place}'
        if year:
            detail += f', {year}'
        detail += '.'
        # Remove the trailing ". " from the title run and append
        if runs and runs[-1]['text'].endswith('. '):
            runs[-1]['text'] = runs[-1]['text'][:-2]
        runs.append(_run(detail))

    elif ref_type in ('CHAP', 'ECHAP'):
        book_title = ref.get('journal', '')
        editors = ref.get('editors', [])
        publisher = ref.get('publisher', '')
        place = ref.get('place', '')
        year = ref.get('year', '')
        sp = ref.get('start_page', '')
        ep = ref.get('end_page', '')

        if book_title:
            runs.append(_run('In '))
            runs.append(_run(book_title.rstrip('.'), italic=True))
            detail = ''
            if editors:
                ed_formatted = [format_author_mdpi(e) for e in editors]
                ed_suffix = 'Ed.' if len(editors) == 1 else 'Eds.'
                detail += f'; {", ".join(ed_formatted)}, {ed_suffix}'
            if publisher:
                detail += f'; {publisher}'
            if place:
                detail += f': {place}'
            if year:
                detail += f', {year}'
            if sp:
                pages = sp
                if ep:
                    pages += f'\u2013{ep}'
                detail += f'; pp. {pages}'
            detail += '.'
            runs.append(_run(detail))

    elif ref_type == 'CONF':
        conf_name = ref.get('journal', '')
        place = ref.get('place', '')
        year = ref.get('year', '')

        if conf_name:
            runs.append(_run('In Proceedings of the '))
            runs.append(_run(conf_name.rstrip('.'), italic=True))
            detail = ''
            if place:
                detail += f', {place}'
            if year:
                detail += f', {year}'
            detail += '.'
            runs.append(_run(detail))
    else:
        year = ref.get('year', '')
        if year:
            runs.append(_run(f'{year}.'))

    # DOI
    doi = ref.get('doi', '')
    if doi:
        if not doi.startswith('http'):
            doi = f'https://doi.org/{doi}'
        runs.append(_run(f' {doi}'))

    return runs


def format_reference_mdpi(ref):
    """Format a reference as a flat string (legacy, for simple use).
    Delegates to format_reference_mdpi_runs and joins text.
    """
    runs = format_reference_mdpi_runs(ref, 1)
    return ''.join(r['text'] for r in runs).lstrip('1.\t')


def format_references_mdpi(ris_records):
    """Format all RIS records into MDPI numbered references as flat strings."""
    formatted = []
    for i, ref in enumerate(ris_records, 1):
        runs = format_reference_mdpi_runs(ref, i)
        text = ''.join(r['text'] for r in runs)
        formatted.append(text)
    return formatted
