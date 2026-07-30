"""Turn a changed fingerprint into something a reader can act on.

Formatting is compared as a bag of named properties, so a difference is
always attributable: not "the format changed" but "numfmt '#,##0' ->
'#,##0,'".  The person approving a report has to be able to tell a harmless
tweak from text quietly set to 4pt white, and "None -> None" tells them
nothing at all.
"""

from __future__ import annotations

from typing import Any, Mapping

# Reports are read, not parsed; a wall of tuples helps nobody.
_MAX_FIELDS = 6
_MAX_VALUE = 60


def as_mapping(value: Any) -> dict[str, Any]:
    """Read a fingerprint as a name -> value mapping.

    Fingerprints are either dicts or sequences of ``(name, value)`` pairs,
    the latter so they stay hashable where they have to live in a set.
    """
    if value is None:
        return {}
    if isinstance(value, Mapping):
        return dict(value)
    if isinstance(value, (list, tuple)):
        pairs = {}
        for item in value:
            if isinstance(item, (list, tuple)) and len(item) == 2:
                pairs[str(item[0])] = item[1]
        return pairs
    return {}


def field_delta(old: Any, new: Any) -> dict[str, tuple[Any, Any]]:
    """Which named fields differ, and how."""
    old_map, new_map = as_mapping(old), as_mapping(new)
    missing = object()
    delta: dict[str, tuple[Any, Any]] = {}
    for name in sorted(old_map.keys() | new_map.keys()):
        before = old_map.get(name, missing)
        after = new_map.get(name, missing)
        if before != after:
            delta[name] = (None if before is missing else before,
                           None if after is missing else after)
    return delta


def _short(value: Any) -> str:
    if value is None:
        return "none"
    text = value if isinstance(value, str) else repr(value)
    return text if len(text) <= _MAX_VALUE else text[:_MAX_VALUE] + "..."


def describe_delta(delta: Mapping[str, tuple[Any, Any]], subject: str) -> str:
    """"cell format changed: numfmt '#,##0' -> '#,##0,'"."""
    if not delta:
        return subject
    names = list(delta)
    shown = names[:_MAX_FIELDS]
    parts = [f"{name} {_short(delta[name][0])} -> {_short(delta[name][1])}"
             for name in shown]
    if len(names) > len(shown):
        parts.append(f"and {len(names) - len(shown)} more")
    return f"{subject}: " + ", ".join(parts)


def delta_values(delta: Mapping[str, tuple[Any, Any]]) -> tuple[dict, dict]:
    """The changed fields only, as old/new maps fit to put on a Change."""
    old = {name: before for name, (before, _) in delta.items()}
    new = {name: after for name, (_, after) in delta.items()}
    return old, new
