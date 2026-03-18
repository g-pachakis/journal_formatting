import pytest


def test_read_manuscript_returns_list(sample_manuscript):
    from reader import read_manuscript
    items = read_manuscript(sample_manuscript)
    assert isinstance(items, list)
    assert len(items) > 0


def test_all_items_have_type_and_text(sample_manuscript):
    from reader import read_manuscript
    items = read_manuscript(sample_manuscript)
    for item in items:
        assert 'type' in item
        if item['type'] != 'table':
            assert 'text' in item
            assert 'runs' in item


def test_table_items_have_rows(sample_manuscript):
    from reader import read_manuscript
    items = read_manuscript(sample_manuscript)
    tables = [i for i in items if i['type'] == 'table']
    assert len(tables) == 1
    assert 'rows' in tables[0]
    assert len(tables[0]['rows']) == 3


def test_classify_title(sample_manuscript):
    from reader import read_manuscript
    items = read_manuscript(sample_manuscript)
    assert items[0]['type'] == 'paragraph'
    assert items[0]['text'] == 'Test Manuscript Title'


def test_classify_abstract(sample_manuscript):
    from reader import read_manuscript
    items = read_manuscript(sample_manuscript)
    types = [i['type'] for i in items]
    assert 'abstract_heading' in types
    assert 'abstract_text' in types


def test_classify_keywords(sample_manuscript):
    from reader import read_manuscript
    items = read_manuscript(sample_manuscript)
    kw = [i for i in items if i['type'] == 'keywords']
    assert len(kw) == 1
    assert 'keyword1' in kw[0]['text']


def test_classify_headings(sample_manuscript):
    from reader import read_manuscript
    items = read_manuscript(sample_manuscript)
    h1 = [i for i in items if i['type'] == 'heading1']
    h2 = [i for i in items if i['type'] == 'heading2']
    h3 = [i for i in items if i['type'] == 'heading3']
    assert len(h1) >= 1
    assert len(h2) == 1
    assert len(h3) == 1


def test_classify_table_caption(sample_manuscript):
    from reader import read_manuscript
    items = read_manuscript(sample_manuscript)
    captions = [i for i in items if i['type'] == 'table_caption']
    assert len(captions) == 1
    assert 'Table 1' in captions[0]['text']


def test_classify_table_footer(sample_manuscript):
    from reader import read_manuscript
    items = read_manuscript(sample_manuscript)
    footers = [i for i in items if i['type'] == 'table_footer']
    assert len(footers) == 1
    assert footers[0]['text'].startswith('*')


def test_classify_figure_placeholder(sample_manuscript):
    from reader import read_manuscript
    items = read_manuscript(sample_manuscript)
    figs = [i for i in items if i['type'] == 'figure_placeholder']
    assert len(figs) == 1


def test_classify_equation(sample_manuscript):
    from reader import read_manuscript
    items = read_manuscript(sample_manuscript)
    eqs = [i for i in items if i['type'] == 'equation']
    assert len(eqs) == 1
    assert '(Eq. 1)' in eqs[0]['text']


def test_classify_references(sample_manuscript):
    from reader import read_manuscript
    items = read_manuscript(sample_manuscript)
    refs_heading = [i for i in items if i['type'] == 'references_heading']
    refs = [i for i in items if i['type'] == 'reference']
    assert len(refs_heading) == 1
    assert len(refs) == 2


def test_runs_preserve_formatting(sample_manuscript_with_formatting):
    from reader import read_manuscript
    items = read_manuscript(sample_manuscript_with_formatting)
    abstract = [i for i in items if i['type'] == 'abstract_text']
    assert len(abstract) == 1
    runs = abstract[0]['runs']
    bold_runs = [r for r in runs if r['bold']]
    italic_runs = [r for r in runs if r['italic']]
    sub_runs = [r for r in runs if r['subscript']]
    sup_runs = [r for r in runs if r['superscript']]
    assert len(bold_runs) >= 1
    assert len(italic_runs) >= 1
    assert len(sub_runs) >= 1
    assert len(sup_runs) >= 1


def test_table_cell_schema(sample_manuscript):
    from reader import read_manuscript
    items = read_manuscript(sample_manuscript)
    table = [i for i in items if i['type'] == 'table'][0]
    cell = table['rows'][0][0]
    assert 'text' in cell
    assert 'gridspan' in cell
    assert 'runs' in cell
    assert 'vmerge_continue' in cell
    assert cell['gridspan'] == 1
    assert cell['vmerge_continue'] is False
