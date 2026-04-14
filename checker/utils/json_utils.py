"""JSON parsing and sanitization for LLM responses."""
import json
import re
from typing import Any


def parse_json_response(raw: str) -> Any:
    """
    Parse JSON from LLM output. Handles:
    1. Direct JSON
    2. Code-fenced JSON (```json ... ```)
    3. JSON embedded in text
    Raises ValueError if unparseable.
    """
    raw = raw.strip()

    # Try direct parse
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass

    # Strip code fences
    cleaned = re.sub(r'^```(?:json)?\s*', '', raw)
    cleaned = re.sub(r'\s*```\s*$', '', cleaned)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass

    # Extract JSON array or object from text
    m = re.search(r'(\[[\s\S]*\]|\{[\s\S]*\})', cleaned)
    if m:
        try:
            return json.loads(m.group(1))
        except json.JSONDecodeError:
            pass

    raise ValueError(f"Cannot parse JSON from response: {raw[:300]}")


def save_json(data: Any, path: str | Any, *, indent: int = 2):
    """Save data as JSON file."""
    from pathlib import Path as P
    p = P(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=indent, default=str)


def load_json(path: str | Any) -> Any:
    """Load JSON from file. Returns None if file doesn't exist."""
    from pathlib import Path as P
    p = P(path)
    if not p.exists():
        return None
    with open(p, encoding="utf-8") as f:
        return json.load(f)


_XML_ILLEGAL_RE = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x84\x86-\x9f]')


def sanitize_text(text: str) -> str:
    """Remove XML-illegal characters from text (for Word output)."""
    if not text:
        return ""
    return _XML_ILLEGAL_RE.sub('', text)
