"""Display-only map wording and saved-layout restoration."""
import json
from pathlib import Path
from .boundary_map import MapError


def display_text(settings, key, original):
    value = settings.get('text_overrides', {}).get(key, original)
    if not isinstance(value, str) or not value.strip():
        raise MapError('Map text cannot be blank. Use Reset map text to restore the original wording.')
    if len(value) > 200 or any(ord(c) < 32 for c in value):
        raise MapError('Use one line of text, up to 200 characters.')
    return value.strip()


def saved_layout(path, data, defaults):
    """Restore only the selected map for this exact workbook and wizard."""
    if not path:
        return defaults
    try:
        meta = json.loads(Path(path).with_suffix('.boundary.json').read_text(encoding='utf-8'))
        old, current = meta['case'], data['case']
        if (old['workbook_hash'], old['wizard']) != (current['workbook_hash'], current['wizard']):
            return defaults
        saved = meta['settings']
        return {key: saved.get(key, value) for key, value in defaults.items()}
    except (OSError, ValueError, KeyError, TypeError, AttributeError):
        return defaults
