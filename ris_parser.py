"""
RIS file parser.

Parses .ris bibliography files into structured reference dicts
for citation matching and formatting.
"""

import re


def parse_ris(path):
    """Parse a .ris file and return a list of reference dicts.

    Each dict contains:
        type: str (JOUR, BOOK, CHAP, CONF, etc.)
        authors: list[str] (e.g., ['Shannon, Claude E.', 'Weaver, Warren'])
        title: str
        journal: str (for JOUR) or book_title for CHAP
        year: str
        volume: str
        issue: str
        start_page: str
        end_page: str
        doi: str
        publisher: str
        place: str
        edition: str
        editors: list[str]
        url: str
        keywords: list[str]
    """
    with open(path, 'r', encoding='utf-8-sig') as f:
        content = f.read()

    records = []
    current = None

    for line in content.splitlines():
        line = line.rstrip()
        if not line:
            continue

        # Match RIS tag: two uppercase letters, two spaces, hyphen, space, value
        m = re.match(r'^([A-Z][A-Z0-9])\s{2}-\s?(.*)', line)
        if not m:
            # Continuation line — append to last field if possible
            if current and current.get('_last_tag'):
                tag = current['_last_tag']
                if tag in current:
                    if isinstance(current[tag], list):
                        current[tag][-1] += ' ' + line.strip()
                    else:
                        current[tag] += ' ' + line.strip()
            continue

        tag = m.group(1)
        value = m.group(2).strip()

        if tag == 'TY':
            current = {'TY': value, '_last_tag': 'TY'}
            continue

        if tag == 'ER':
            if current:
                del current['_last_tag']
                records.append(_normalize_record(current))
                current = None
            continue

        if current is None:
            continue

        current['_last_tag'] = tag

        # Repeatable fields stored as lists
        if tag in ('AU', 'A1', 'A2', 'A3', 'A4', 'ED', 'KW'):
            current.setdefault(tag, [])
            current[tag].append(value)
        else:
            current[tag] = value

    # Handle file without final ER
    if current:
        if '_last_tag' in current:
            del current['_last_tag']
        records.append(_normalize_record(current))

    return records


def _normalize_record(raw):
    """Convert raw RIS tag dict to a clean reference dict."""
    ref = {
        'type': raw.get('TY', 'GEN'),
        'authors': raw.get('AU', raw.get('A1', [])),
        'editors': raw.get('A2', raw.get('ED', [])),
        'title': raw.get('TI', raw.get('T1', '')),
        'journal': raw.get('T2', raw.get('JO', raw.get('JF', raw.get('JA', '')))),
        'year': '',
        'volume': raw.get('VL', ''),
        'issue': raw.get('IS', ''),
        'start_page': raw.get('SP', ''),
        'end_page': raw.get('EP', ''),
        'doi': raw.get('DO', ''),
        'publisher': raw.get('PB', ''),
        'place': raw.get('CY', raw.get('PP', '')),
        'edition': raw.get('ET', ''),
        'url': raw.get('UR', ''),
        'keywords': raw.get('KW', []),
        'isbn': raw.get('SN', ''),
        'abstract': raw.get('AB', raw.get('N2', '')),
    }

    # Extract year from PY or Y1 (format: YYYY/MM/DD/other or just YYYY)
    py = raw.get('PY', raw.get('Y1', raw.get('DA', '')))
    if py:
        year_match = re.match(r'(\d{4})', py)
        if year_match:
            ref['year'] = year_match.group(1)

    # For book types, title might be in BT
    if not ref['title'] and 'BT' in raw:
        ref['title'] = raw['BT']

    return ref


def match_citation_to_ris(citation_text, ris_records):
    """Match a hardwritten citation string to a RIS record.

    Uses fuzzy matching on author last names and year.
    Returns the matched RIS record or None.
    """
    citation_text = citation_text.strip()
    # Strip leading [N] number
    citation_text = re.sub(r'^\[?\d+\]?\s*', '', citation_text)

    for rec in ris_records:
        score = 0
        total = 0

        # Match year
        if rec['year']:
            total += 2
            if rec['year'] in citation_text:
                score += 2

        # Match author last names
        for author in rec['authors'][:3]:  # Check first 3 authors
            last_name = author.split(',')[0].strip()
            if last_name and len(last_name) > 2:
                total += 1
                if last_name in citation_text:
                    score += 1

        # Match DOI
        if rec['doi'] and rec['doi'] in citation_text:
            return rec

        # Match title words (at least 3 consecutive words)
        if rec['title']:
            title_words = rec['title'].split()
            if len(title_words) >= 3:
                # Check if 3+ consecutive title words appear in citation
                for i in range(len(title_words) - 2):
                    snippet = ' '.join(title_words[i:i+3]).lower()
                    if snippet.lower() in citation_text.lower():
                        score += 2
                        total += 2
                        break
                else:
                    total += 2

        # Need at least 60% match with minimum 2 points
        if total > 0 and score >= 2 and score / total >= 0.5:
            return rec

    return None
