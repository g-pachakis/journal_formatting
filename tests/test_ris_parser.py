import os
import pytest


@pytest.fixture
def sample_ris(tmp_path):
    content = """TY  - JOUR
AU  - Shannon, Claude E.
AU  - Weaver, Warren
TI  - A Mathematical Theory of Communication
T2  - Bell System Technical Journal
VL  - 27
SP  - 379
EP  - 423
PY  - 1948
DO  - 10.1002/j.1538-7305.1948.tb01338.x
ER  -
TY  - BOOK
AU  - Smith, John D.
TI  - Introduction to Chemical Engineering
PB  - Wiley
CY  - New York
PY  - 2020
ET  - 3rd
ER  -
TY  - CHAP
AU  - Jones, Robert A.
TI  - Absorption in Packed Columns
T2  - Handbook of Chemical Engineering
A2  - Smith, John D.
PB  - McGraw-Hill
CY  - New York
PY  - 2019
SP  - 145
EP  - 198
ER  -
"""
    path = tmp_path / 'test_refs.ris'
    path.write_text(content, encoding='utf-8')
    return str(path)


def test_parse_ris_returns_list(sample_ris):
    from ris_parser import parse_ris
    records = parse_ris(sample_ris)
    assert isinstance(records, list)
    assert len(records) == 3


def test_parse_ris_journal_article(sample_ris):
    from ris_parser import parse_ris
    records = parse_ris(sample_ris)
    jour = records[0]
    assert jour['type'] == 'JOUR'
    assert jour['authors'] == ['Shannon, Claude E.', 'Weaver, Warren']
    assert jour['title'] == 'A Mathematical Theory of Communication'
    assert jour['journal'] == 'Bell System Technical Journal'
    assert jour['year'] == '1948'
    assert jour['volume'] == '27'
    assert jour['start_page'] == '379'
    assert jour['end_page'] == '423'
    assert jour['doi'] == '10.1002/j.1538-7305.1948.tb01338.x'


def test_parse_ris_book(sample_ris):
    from ris_parser import parse_ris
    records = parse_ris(sample_ris)
    book = records[1]
    assert book['type'] == 'BOOK'
    assert book['authors'] == ['Smith, John D.']
    assert book['title'] == 'Introduction to Chemical Engineering'
    assert book['publisher'] == 'Wiley'
    assert book['year'] == '2020'


def test_parse_ris_chapter(sample_ris):
    from ris_parser import parse_ris
    records = parse_ris(sample_ris)
    chap = records[2]
    assert chap['type'] == 'CHAP'
    assert chap['title'] == 'Absorption in Packed Columns'
    assert chap['journal'] == 'Handbook of Chemical Engineering'
    assert chap['editors'] == ['Smith, John D.']


def test_match_citation_to_ris(sample_ris):
    from ris_parser import parse_ris, match_citation_to_ris
    records = parse_ris(sample_ris)
    citation = '[1] Shannon, C.E.; Weaver, W. A Mathematical Theory of Communication. Bell Syst. Tech. J. 1948, 27, 379-423.'
    match = match_citation_to_ris(citation, records)
    assert match is not None
    assert match['authors'][0] == 'Shannon, Claude E.'


def test_match_citation_no_match(sample_ris):
    from ris_parser import parse_ris, match_citation_to_ris
    records = parse_ris(sample_ris)
    citation = '[99] Unknown, A. Totally Different Paper. Some Journal 2099, 1, 1-2.'
    match = match_citation_to_ris(citation, records)
    assert match is None


def test_format_author_mdpi():
    from citation_formatter import format_author_mdpi
    assert format_author_mdpi('Shannon, Claude E.') == 'Shannon, C.E.'
    assert format_author_mdpi('Smith, John D.') == 'Smith, J.D.'
    assert format_author_mdpi('Claude E. Shannon') == 'Shannon, C.E.'


def test_format_reference_mdpi_journal():
    from citation_formatter import format_reference_mdpi
    ref = {
        'type': 'JOUR',
        'authors': ['Shannon, Claude E.', 'Weaver, Warren'],
        'title': 'A Mathematical Theory of Communication',
        'journal': 'Bell Syst. Tech. J.',
        'year': '1948',
        'volume': '27',
        'start_page': '379',
        'end_page': '423',
        'doi': '10.1002/j.1538-7305.1948.tb01338.x',
        'issue': '', 'publisher': '', 'place': '', 'edition': '',
        'editors': [], 'url': '', 'keywords': [], 'isbn': '', 'abstract': '',
    }
    result = format_reference_mdpi(ref)
    assert 'Shannon, C.E.' in result
    assert 'Weaver, W.' in result
    assert 'Mathematical Theory' in result
    assert '1948' in result
    assert '27' in result
    assert '379\u2013423' in result


def test_format_reference_mdpi_book():
    from citation_formatter import format_reference_mdpi
    ref = {
        'type': 'BOOK',
        'authors': ['Smith, John D.'],
        'title': 'Introduction to Chemical Engineering',
        'journal': '', 'year': '2020', 'volume': '', 'issue': '',
        'start_page': '', 'end_page': '',
        'doi': '', 'publisher': 'Wiley', 'place': 'New York',
        'edition': '3rd', 'editors': [], 'url': '', 'keywords': [],
        'isbn': '', 'abstract': '',
    }
    result = format_reference_mdpi(ref)
    assert 'Smith, J.D.' in result
    assert 'Introduction to Chemical Engineering' in result
    assert 'Wiley' in result
    assert '2020' in result


def test_format_reference_mdpi_runs_has_italic_journal():
    from citation_formatter import format_reference_mdpi_runs
    ref = {
        'type': 'JOUR',
        'authors': ['Shannon, Claude E.'],
        'title': 'A Mathematical Theory of Communication',
        'journal': 'Bell Syst. Tech. J.',
        'year': '1948', 'volume': '27',
        'start_page': '379', 'end_page': '423',
        'doi': '10.1002/test', 'issue': '', 'publisher': '', 'place': '',
        'edition': '', 'editors': [], 'url': '', 'keywords': [], 'isbn': '', 'abstract': '',
    }
    runs = format_reference_mdpi_runs(ref, 1)
    italic_runs = [r for r in runs if r['italic']]
    assert len(italic_runs) >= 1
    assert any('Bell Syst' in r['text'] for r in italic_runs)


def test_format_reference_mdpi_runs_numbering():
    from citation_formatter import format_reference_mdpi_runs
    ref = {
        'type': 'JOUR', 'authors': ['Test, A.'],
        'title': 'Test', 'journal': 'J. Test',
        'year': '2020', 'volume': '1', 'start_page': '1', 'end_page': '2',
        'doi': '', 'issue': '', 'publisher': '', 'place': '',
        'edition': '', 'editors': [], 'url': '', 'keywords': [], 'isbn': '', 'abstract': '',
    }
    runs = format_reference_mdpi_runs(ref, 5)
    assert runs[0]['text'].startswith('5.\t')


def test_format_reference_mdpi_runs_endash_pages():
    from citation_formatter import format_reference_mdpi_runs
    ref = {
        'type': 'JOUR', 'authors': ['Test, A.'],
        'title': 'Test', 'journal': 'J. Test',
        'year': '2020', 'volume': '1', 'start_page': '100', 'end_page': '200',
        'doi': '', 'issue': '', 'publisher': '', 'place': '',
        'edition': '', 'editors': [], 'url': '', 'keywords': [], 'isbn': '', 'abstract': '',
    }
    runs = format_reference_mdpi_runs(ref, 1)
    all_text = ''.join(r['text'] for r in runs)
    assert '\u2013' in all_text  # en-dash between pages
