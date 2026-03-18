import pytest


def test_extract_dois_from_text():
    from crossref_client import extract_dois
    text = 'See Shannon (1948) doi: 10.1002/j.1538-7305.1948.tb01338.x and also 10.3390/en1234.'
    dois = extract_dois(text)
    assert '10.1002/j.1538-7305.1948.tb01338.x' in dois
    assert '10.3390/en1234' in dois


def test_extract_dois_no_match():
    from crossref_client import extract_dois
    text = 'No DOIs here, just text.'
    assert extract_dois(text) == []


def test_extract_dois_strips_trailing_period():
    from crossref_client import extract_dois
    text = 'The DOI is 10.1234/test.abc.'
    dois = extract_dois(text)
    assert dois[0] == '10.1234/test.abc'


def test_crossref_to_ref_structure():
    from crossref_client import _crossref_to_ref
    item = {
        'type': 'journal-article',
        'title': ['Test Article'],
        'author': [{'family': 'Smith', 'given': 'John'}],
        'container-title': ['Test Journal'],
        'volume': '42',
        'issue': '3',
        'page': '100-200',
        'DOI': '10.1234/test',
        'issued': {'date-parts': [[2020, 5]]},
        'publisher': 'Test Publisher',
    }
    ref = _crossref_to_ref(item)
    assert ref['type'] == 'JOUR'
    assert ref['title'] == 'Test Article'
    assert ref['authors'] == ['Smith, John']
    assert ref['journal'] == 'Test Journal'
    assert ref['year'] == '2020'
    assert ref['volume'] == '42'
    assert ref['start_page'] == '100'
    assert ref['end_page'] == '200'
    assert ref['doi'] == '10.1234/test'


def test_map_type():
    from crossref_client import _map_type
    assert _map_type('journal-article') == 'JOUR'
    assert _map_type('book') == 'BOOK'
    assert _map_type('book-chapter') == 'CHAP'
    assert _map_type('unknown-type') == 'GEN'
