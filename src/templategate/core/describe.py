"""Turn a changed fingerprint into something a reader can act on.

Formatting is compared as a bag of named properties, so a difference is
always attributable: not "the format changed" but "numfmt '#,##0' ->
'#,##0,'".  The person approving a report has to be able to tell a harmless
tweak from text quietly set to 4pt white, and "None -> None" tells them
nothing at all.
"""

from __future__ import annotations

from typing import Any, Mapping

from .messages import DeltaDetail

# Reports are read, not parsed; a wall of tuples helps nobody.
_MAX_FIELDS = 6


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
    delta: dict[str, tuple[Any, Any]] = {}
    for name in sorted(old_map.keys() | new_map.keys()):
        # Absent and explicitly-None mean the same thing: not set.  Treating
        # them as different is what made a cell that merely gained a style
        # report twenty rows of "none -> none".
        before, after = old_map.get(name), new_map.get(name)
        if before != after:
            delta[name] = (before, after)
    return delta


def describe_delta(delta: Mapping[str, tuple[Any, Any]],
                   subject_key: str) -> DeltaDetail:
    """"cell format changed: numfmt '#,##0' -> '#,##0,'".

    Returns the pieces rather than the sentence: the same delta has to come
    out as English for the JSON contract and as whatever the reader asked for
    in the report, and only the caller's locale decides which.
    """
    names = list(delta)
    shown = names[:_MAX_FIELDS]
    fields = [(name, delta[name][0], delta[name][1]) for name in shown]
    return DeltaDetail(subject_key, fields, len(names) - len(shown))


def delta_values(delta: Mapping[str, tuple[Any, Any]]) -> tuple[dict, dict]:
    """The changed fields only, as old/new maps fit to put on a Change."""
    old = {name: before for name, (before, _) in delta.items()}
    new = {name: after for name, (_, after) in delta.items()}
    return old, new
