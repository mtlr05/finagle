import json
import os

import pytest

from finagle.agent import run_case, validate_case
from finagle.agent.schema import iter_leaves

ROOT = os.path.dirname(os.path.dirname(__file__))
CASE = os.path.join(ROOT, 'examples', 'article_case.json')
ARTICLE = os.path.join(ROOT, 'examples', 'sample_article.md')


def _load():
    with open(CASE, encoding='utf-8') as handle:
        return json.load(handle)


def test_every_article_quote_is_in_the_article():
    with open(ARTICLE, encoding='utf-8') as handle:
        text = handle.read()
    case = _load()
    quoted = [
        (path, node['evidence'])
        for path, node in iter_leaves(case)
        if node['source'] == 'article' and path != 'article.url'
    ]
    assert quoted
    for path, evidence in quoted:
        assert evidence in text, path


def test_example_validates_with_only_the_price_date_warning():
    report = validate_case(_load())
    assert report['errors'] == []
    assert report['missing'] == []
    assert report['warnings'] == ['article price is older than the article date']


def test_example_runs_and_the_buyback_program_ends(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    result = run_case(_load())
    assert result['ok'], {'error': result.get('error'), 'log': result['log']}
    assert result['log'] == []
    assert result['sanity_flags'] == []
    assert result['value_per_share'] > 0
    assert result['value_per_share_DDM'] > 0
    assert result['article_price_target'] == 55

    shares = [row['shares'] for row in result['years']]
    assert shares[4] < shares[0]
    assert shares[4:] == pytest.approx([shares[4]] * len(shares[4:]))

    acquire = next(step for step in result['steps'] if step['method'] == 'fcf_to_acquire')
    assert acquire['args']['ebitda_frac'] == pytest.approx(10 / (240 * 1.08))

    namespace = {}
    exec(result['notebook_code'], namespace)
    assert namespace['model'].fin['value_per_share'].iloc[0] == pytest.approx(result['value_per_share'])
    assert list(tmp_path.iterdir()) == [tmp_path / 'XMPL.log']
