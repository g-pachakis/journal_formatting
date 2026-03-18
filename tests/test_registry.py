import pytest


def test_get_formats_returns_dict():
    from formats import get_formats
    result = get_formats()
    assert isinstance(result, dict)


def test_get_formats_finds_elsevier():
    from formats import get_formats
    fmts = get_formats()
    assert 'Elsevier' in fmts


def test_get_formats_finds_mdpi():
    from formats import get_formats
    fmts = get_formats()
    assert 'MDPI' in fmts


def test_format_modules_have_required_attributes():
    from formats import get_formats
    fmts = get_formats()
    for name, mod in fmts.items():
        assert hasattr(mod, 'FORMAT_NAME'), f'{name} missing FORMAT_NAME'
        assert hasattr(mod, 'FORMAT_SUFFIX'), f'{name} missing FORMAT_SUFFIX'
        assert hasattr(mod, 'build'), f'{name} missing build()'
        assert callable(mod.build), f'{name}.build is not callable'
