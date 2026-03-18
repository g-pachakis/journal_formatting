"""
Citation formatter.

Formats reference data (from RIS records) into journal-specific
citation styles. Currently supports MDPI numbered format.
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

    # Already in "Last, First" format
    if ',' in author_str:
        parts = author_str.split(',', 1)
        last = parts[0].strip()
        given = parts[1].strip()
    else:
        # "First Last" format
        parts = author_str.rsplit(' ', 1)
        if len(parts) == 2:
            given = parts[0].strip()
            last = parts[1].strip()
        else:
            return author_str

    # Convert given names to initials
    initials = []
    for name in given.split():
        name = name.strip().rstrip('.')
        if name:
            initials.append(name[0].upper() + '.')

    return f"{last}, {''.join(initials)}"


def format_reference_mdpi(ref):
    """Format a reference dict (from RIS) into MDPI numbered style.

    Returns a formatted reference string (without the [N] number prefix).

    MDPI format for journal articles:
        Lastname, F.M.; Lastname, F.M. Title of Article. Journal Year, Volume, Pages.
    """
    ref_type = ref.get('type', 'JOUR')
    parts = []

    # Authors
    authors = ref.get('authors', [])
    if authors:
        formatted_authors = [format_author_mdpi(a) for a in authors]
        parts.append('; '.join(formatted_authors) + '.')

    # Title
    title = ref.get('title', '')
    if title:
        # Ensure title ends with period
        title = title.rstrip('.')
        parts.append(title + '.')

    if ref_type in ('JOUR', 'EJOUR', 'JFULL'):
        # Journal article
        journal = ref.get('journal', '')
        year = ref.get('year', '')
        volume = ref.get('volume', '')
        sp = ref.get('start_page', '')
        ep = ref.get('end_page', '')

        if journal:
            journal_part = journal.rstrip('.')
            if year:
                journal_part += f' {year}'
            if volume:
                journal_part += f', {volume}'
            if sp:
                pages = sp
                if ep and ep != sp:
                    pages += f'-{ep}'
                journal_part += f', {pages}'
            journal_part += '.'
            parts.append(journal_part)

    elif ref_type in ('BOOK', 'EBOOK'):
        # Book
        edition = ref.get('edition', '')
        publisher = ref.get('publisher', '')
        place = ref.get('place', '')
        year = ref.get('year', '')

        book_info = ''
        if edition:
            book_info += f', {edition}'
        if publisher or place:
            book_info += '; '
            if publisher:
                book_info += publisher
            if place:
                book_info += f': {place}'
        if year:
            book_info += f', {year}'
        if book_info:
            parts[-1] = parts[-1].rstrip('.') + book_info + '.'

    elif ref_type in ('CHAP', 'ECHAP'):
        # Book chapter
        book_title = ref.get('journal', '')  # T2 maps to journal in our parser
        editors = ref.get('editors', [])
        publisher = ref.get('publisher', '')
        place = ref.get('place', '')
        year = ref.get('year', '')
        sp = ref.get('start_page', '')
        ep = ref.get('end_page', '')

        if book_title:
            chap_part = f'In {book_title}'
            if editors:
                ed_formatted = [format_author_mdpi(e) for e in editors]
                ed_suffix = 'Ed.' if len(editors) == 1 else 'Eds.'
                chap_part += f'; {", ".join(ed_formatted)}, {ed_suffix}'
            if publisher:
                chap_part += f'; {publisher}'
            if place:
                chap_part += f': {place}'
            if year:
                chap_part += f', {year}'
            if sp:
                pages = sp
                if ep:
                    pages += f'-{ep}'
                chap_part += f'; pp. {pages}'
            chap_part += '.'
            parts.append(chap_part)

    elif ref_type == 'CONF':
        # Conference paper
        conf_name = ref.get('journal', '')
        place = ref.get('place', '')
        year = ref.get('year', '')

        if conf_name:
            conf_part = f'In Proceedings of the {conf_name}'
            if place:
                conf_part += f', {place}'
            if year:
                conf_part += f', {year}'
            conf_part += '.'
            parts.append(conf_part)

    else:
        # Generic fallback
        year = ref.get('year', '')
        if year:
            parts.append(f'{year}.')

    # DOI
    doi = ref.get('doi', '')
    if doi:
        if not doi.startswith('http'):
            doi = f'https://doi.org/{doi}'
        parts.append(doi + '.')

    return ' '.join(parts)


def format_references_mdpi(ris_records):
    """Format all RIS records into MDPI numbered references.

    Returns list of formatted reference strings with [N] prefixes.
    """
    formatted = []
    for i, ref in enumerate(ris_records, 1):
        text = format_reference_mdpi(ref)
        formatted.append(f'[{i}]\t{text}')
    return formatted
