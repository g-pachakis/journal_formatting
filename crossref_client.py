"""
CrossRef API client for reference metadata lookup.

Uses the free CrossRef REST API (no authentication required).
Looks up references by DOI or bibliographic text search.
"""

import re
import json
import urllib.request
import urllib.parse
import urllib.error

CROSSREF_API = 'https://api.crossref.org'
USER_AGENT = 'ManuscriptFormatter/1.0 (https://github.com/g-pachakis/journal_formatting; mailto:gpachakis@upatras.gr)'


def extract_dois(text):
    """Extract DOIs from a text string.

    Returns list of DOI strings found (e.g., ['10.1002/abc.123']).
    """
    pattern = r'(10\.\d{4,9}/[^\s,;\]>]+)'
    matches = re.findall(pattern, text)
    # Clean trailing punctuation
    cleaned = []
    for doi in matches:
        doi = doi.rstrip('.')
        cleaned.append(doi)
    return cleaned


def lookup_doi(doi):
    """Look up a single DOI via CrossRef API.

    Returns a reference dict compatible with our ris_parser format, or None.
    """
    url = f'{CROSSREF_API}/works/{urllib.parse.quote(doi, safe="")}'
    try:
        req = urllib.request.Request(url, headers={'User-Agent': USER_AGENT})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode('utf-8'))
        return _crossref_to_ref(data.get('message', {}))
    except (urllib.error.URLError, urllib.error.HTTPError, json.JSONDecodeError,
            TimeoutError, OSError):
        return None


def search_reference(text, rows=1):
    """Search CrossRef for a reference using bibliographic text.

    Returns a list of reference dicts (up to `rows` results).
    """
    # Clean the text for search
    clean = re.sub(r'^\[?\d+\]?\s*', '', text)  # strip [N] prefix
    clean = re.sub(r'https?://\S+', '', clean)   # strip URLs
    clean = clean[:300]  # limit length

    params = urllib.parse.urlencode({
        'query.bibliographic': clean,
        'rows': rows,
    })
    url = f'{CROSSREF_API}/works?{params}'
    try:
        req = urllib.request.Request(url, headers={'User-Agent': USER_AGENT})
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read().decode('utf-8'))
        items = data.get('message', {}).get('items', [])
        return [_crossref_to_ref(item) for item in items]
    except (urllib.error.URLError, urllib.error.HTTPError, json.JSONDecodeError,
            TimeoutError, OSError):
        return []


def _crossref_to_ref(item):
    """Convert a CrossRef work item to our standard reference dict."""
    ref = {
        'type': _map_type(item.get('type', '')),
        'authors': [],
        'editors': [],
        'title': '',
        'journal': '',
        'year': '',
        'volume': item.get('volume', ''),
        'issue': item.get('issue', ''),
        'start_page': '',
        'end_page': '',
        'doi': item.get('DOI', ''),
        'publisher': item.get('publisher', ''),
        'place': '',
        'edition': '',
        'url': item.get('URL', ''),
        'keywords': [],
        'isbn': '',
        'abstract': '',
    }

    # Title
    titles = item.get('title', [])
    if titles:
        ref['title'] = titles[0]

    # Authors
    for author in item.get('author', []):
        family = author.get('family', '')
        given = author.get('given', '')
        if family:
            ref['authors'].append(f'{family}, {given}' if given else family)

    # Editors
    for editor in item.get('editor', []):
        family = editor.get('family', '')
        given = editor.get('given', '')
        if family:
            ref['editors'].append(f'{family}, {given}' if given else family)

    # Journal / container title
    container = item.get('container-title', [])
    if container:
        ref['journal'] = container[0]

    # Year
    issued = item.get('issued', {})
    date_parts = issued.get('date-parts', [[]])
    if date_parts and date_parts[0]:
        ref['year'] = str(date_parts[0][0])

    # Pages
    page = item.get('page', '')
    if page and '-' in page:
        parts = page.split('-', 1)
        ref['start_page'] = parts[0].strip()
        ref['end_page'] = parts[1].strip()
    elif page:
        ref['start_page'] = page.strip()

    # ISBN
    isbn_list = item.get('ISBN', [])
    if isbn_list:
        ref['isbn'] = isbn_list[0]

    return ref


def _map_type(crossref_type):
    """Map CrossRef type to RIS-compatible type."""
    mapping = {
        'journal-article': 'JOUR',
        'book': 'BOOK',
        'book-chapter': 'CHAP',
        'proceedings-article': 'CONF',
        'monograph': 'BOOK',
        'edited-book': 'BOOK',
        'reference-book': 'BOOK',
        'dissertation': 'THES',
        'report': 'RPRT',
    }
    return mapping.get(crossref_type, 'GEN')
