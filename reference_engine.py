"""
Reference resolution engine.

Orchestrates the full pipeline: extract DOIs, match against RIS,
optionally look up via CrossRef, and format in the target style.
"""

import re
from dataclasses import dataclass, field

from crossref_client import extract_dois, lookup_doi, search_reference
from ris_parser import match_citation_to_ris
from citation_formatter import format_reference_mdpi_runs


@dataclass
class ResolvedReference:
    """A reference with resolved metadata and formatted output."""
    index: int
    original_text: str
    metadata: dict = field(default_factory=dict)
    source: str = 'unmatched'  # 'ris', 'crossref_doi', 'crossref_search', 'unmatched'
    doi: str = ''
    formatted_runs: list = field(default_factory=list)


def extract_ref_number(text):
    """Extract the reference number from text like '[1]' or '1.' at the start."""
    m = re.match(r'^\[?(\d+)\]?\.?\s*', text)
    return int(m.group(1)) if m else None


def resolve_references(ref_items, ris_data=None, use_crossref=False,
                       progress_callback=None):
    """Resolve all references through the matching pipeline.

    Args:
        ref_items: list of reader items where type == 'reference'
        ris_data: optional list of RIS records from ris_parser.parse_ris()
        use_crossref: whether to use CrossRef API for unmatched references
        progress_callback: optional callable(current, total, message) for GUI

    Returns:
        list of ResolvedReference objects
    """
    results = []
    total = len(ref_items)

    for i, item in enumerate(ref_items):
        text = item['text']
        ref_num = extract_ref_number(text) or (i + 1)

        resolved = ResolvedReference(
            index=ref_num,
            original_text=text,
        )

        # Step 1: Extract DOIs from reference text
        dois = extract_dois(text)
        if dois:
            resolved.doi = dois[0]

        # Step 2: Match against RIS data (if provided)
        if ris_data:
            # Try DOI match first (most reliable)
            if resolved.doi:
                for rec in ris_data:
                    if rec.get('doi') and rec['doi'].lower() == resolved.doi.lower():
                        resolved.metadata = rec
                        resolved.source = 'ris'
                        break

            # Fall back to fuzzy text matching
            if not resolved.metadata:
                match = match_citation_to_ris(text, ris_data)
                if match:
                    resolved.metadata = match
                    resolved.source = 'ris'

        # Step 3: CrossRef lookup (if enabled and still unmatched)
        if use_crossref and not resolved.metadata:
            if progress_callback:
                progress_callback(i + 1, total, f'Looking up reference {ref_num}...')

            # Try DOI lookup first
            if resolved.doi:
                cr_ref = lookup_doi(resolved.doi)
                if cr_ref:
                    resolved.metadata = cr_ref
                    resolved.source = 'crossref_doi'

            # Fall back to text search
            if not resolved.metadata:
                cr_results = search_reference(text, rows=1)
                if cr_results:
                    resolved.metadata = cr_results[0]
                    resolved.source = 'crossref_search'
                    if not resolved.doi and resolved.metadata.get('doi'):
                        resolved.doi = resolved.metadata['doi']

        # Step 4: Format the reference
        if resolved.metadata:
            resolved.formatted_runs = format_reference_mdpi_runs(
                resolved.metadata, resolved.index)
        else:
            # Unmatched — preserve original as-is but fix numbering to N.\t format
            clean_text = re.sub(r'^\[?\d+\]?\s*', '', text)
            resolved.formatted_runs = [
                {'text': f'{resolved.index}.\t{clean_text}',
                 'bold': False, 'italic': False,
                 'superscript': False, 'subscript': False}
            ]

        results.append(resolved)

        if progress_callback:
            progress_callback(i + 1, total, f'Resolved {i + 1}/{total} references')

    return results


def get_resolution_stats(results):
    """Get statistics about reference resolution.

    Returns dict with counts per source and list of unmatched indices.
    """
    stats = {
        'total': len(results),
        'ris': 0,
        'crossref_doi': 0,
        'crossref_search': 0,
        'unmatched': 0,
        'unmatched_indices': [],
    }
    for r in results:
        stats[r.source] += 1
        if r.source == 'unmatched':
            stats['unmatched_indices'].append(r.index)
    return stats
