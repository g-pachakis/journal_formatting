import pytest


@pytest.fixture
def sample_ref_items():
    return [
        {'type': 'reference', 'text': '[1] Shannon, C.E. A Mathematical Theory of Communication. Bell Syst. Tech. J. 1948, 27, 379-423. https://doi.org/10.1002/test', 'runs': []},
        {'type': 'reference', 'text': '[2] Smith, J.D. Unknown Paper. Some Journal 2099, 1, 1-2.', 'runs': []},
        {'type': 'reference', 'text': '[3] Jones, R. Third reference without DOI. Another J. 2020, 5, 10-15.', 'runs': []},
    ]


@pytest.fixture
def sample_ris_data():
    return [
        {
            'type': 'JOUR', 'authors': ['Shannon, Claude E.', 'Weaver, Warren'],
            'title': 'A Mathematical Theory of Communication',
            'journal': 'Bell System Technical Journal',
            'year': '1948', 'volume': '27', 'start_page': '379', 'end_page': '423',
            'doi': '10.1002/test', 'issue': '', 'publisher': '', 'place': '',
            'edition': '', 'editors': [], 'url': '', 'keywords': [], 'isbn': '', 'abstract': '',
        },
    ]


def test_extract_ref_number():
    from reference_engine import extract_ref_number
    assert extract_ref_number('[1] Some text') == 1
    assert extract_ref_number('[42] Some text') == 42
    assert extract_ref_number('5. Some text') == 5
    assert extract_ref_number('No number') is None


def test_resolve_references_with_ris(sample_ref_items, sample_ris_data):
    from reference_engine import resolve_references
    results = resolve_references(sample_ref_items, ris_data=sample_ris_data, use_crossref=False)
    assert len(results) == 3
    # First should match via DOI
    assert results[0].source == 'ris'
    assert results[0].doi == '10.1002/test'
    assert len(results[0].formatted_runs) > 0
    # Second and third should be unmatched
    assert results[1].source == 'unmatched'
    assert results[2].source == 'unmatched'


def test_resolve_references_no_ris(sample_ref_items):
    from reference_engine import resolve_references
    results = resolve_references(sample_ref_items, ris_data=None, use_crossref=False)
    assert all(r.source == 'unmatched' for r in results)
    # Unmatched should still have formatted_runs with N.\t prefix
    assert results[0].formatted_runs[0]['text'].startswith('1.\t')


def test_resolved_reference_formatting(sample_ref_items, sample_ris_data):
    from reference_engine import resolve_references
    results = resolve_references(sample_ref_items, ris_data=sample_ris_data, use_crossref=False)
    matched = results[0]
    # Should have italic journal name
    italic_runs = [r for r in matched.formatted_runs if r['italic']]
    assert len(italic_runs) >= 1


def test_get_resolution_stats(sample_ref_items, sample_ris_data):
    from reference_engine import resolve_references, get_resolution_stats
    results = resolve_references(sample_ref_items, ris_data=sample_ris_data, use_crossref=False)
    stats = get_resolution_stats(results)
    assert stats['total'] == 3
    assert stats['ris'] == 1
    assert stats['unmatched'] == 2
    assert 2 in stats['unmatched_indices']
    assert 3 in stats['unmatched_indices']
