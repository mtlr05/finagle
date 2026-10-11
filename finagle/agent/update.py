'''Change a finished case and rerun it.

The documents, the SEC, and yfinance are not read again. Derived
assumptions whose inputs changed are recomputed from the numbers already
in the case.
'''

import copy
import re

from finagle.agent.defaults import apply_defaults
from finagle.agent.runner import run_case
from finagle.agent.schema import iter_leaves
from finagle.agent.validate import validate_case

_PATH = re.compile(r'([^.\[]+)(?:\[(\d+)\])?')


def update_case(case, updates, run=True):
    '''Apply ``updates`` to ``case``, refill derived assumptions, and rerun.

    ``updates`` is a partial case of leaves. Each leaf should use source
    ``prompt``. Returns the run result, including the updated case. Pass
    ``run=False`` to validate without running the model.
    '''
    if not isinstance(case, dict):
        raise ValueError('case must be an object')
    updated = copy.deepcopy(case)
    changed = _merge(updated, updates or {})
    _drop_stale(updated, changed)
    currency = _currency(updated)
    apply_defaults(updated, currency=currency, fetch_yield=False)
    if not run:
        return {'ok': True, 'case': updated, 'validation': validate_case(updated)}
    result = run_case(updated)
    result['case'] = updated
    return result


def _merge(case, updates):
    changed = set()
    for path, node in iter_leaves(updates):
        _assign(case, path, copy.deepcopy(node))
        changed.add(path)
    return changed


def _drop_stale(case, changed):
    changed = set(changed)
    while True:
        stale = [
            path for path, node in iter_leaves(case)
            if set(node.get('derived_from') or []) & changed
        ]
        if not stale:
            return
        for path in stale:
            _delete(case, path)
            changed.add(path)


def _assign(case, path, node):
    parent, key = _parent(case, path, create=True)
    parent[key] = node


def _delete(case, path):
    parent, key = _parent(case, path, create=False)
    if isinstance(parent, dict) and key in parent:
        parent.pop(key)
    elif isinstance(parent, list) and isinstance(key, int) and 0 <= key < len(parent):
        parent.pop(key)


def _parent(case, path, create):
    cursor = case
    parts = _parts(path)
    for kind, key in parts[:-1]:
        cursor = _step(cursor, kind, key, create)
    kind, key = parts[-1]
    if kind == 'index':
        return cursor, key
    if create and not isinstance(cursor, dict):
        raise ValueError('cannot set %s' % path)
    return cursor, key


def _step(cursor, kind, key, create):
    if kind == 'key':
        if not isinstance(cursor, dict):
            raise ValueError('cannot follow %s' % key)
        if key not in cursor or cursor[key] is None:
            if not create:
                return {}
            cursor[key] = {}
        return cursor[key]
    if not isinstance(cursor, list):
        raise ValueError('cannot index a %s' % type(cursor).__name__)
    return cursor[key]


def _parts(path):
    parts = []
    for match in _PATH.finditer(path):
        parts.append(('key', match.group(1)))
        if match.group(2) is not None:
            parts.append(('index', int(match.group(2))))
    if not parts:
        raise ValueError('empty path')
    return parts


def _currency(case):
    baseline = case.get('baseline') or {}
    for node in baseline.values():
        if isinstance(node, dict) and node.get('currency'):
            return str(node['currency']).upper()
    return None
