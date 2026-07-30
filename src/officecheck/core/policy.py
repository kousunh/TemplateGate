"""Policy loading and validation.

A policy is the *trusted* description of what an edit is allowed to change.
It must be authored or approved by a human (or pinned in CI) — never by the
agent that produced the candidate document.  OfficeCheck only reads it.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import yaml

VALID_TARGETS = ("excel", "word", "auto")
VALID_MODES = ("review_only", "normal_input", "page_extension", "row_extension")
VALID_SEMANTIC_MODES = ("off", "review", "gate")


class PolicyError(ValueError):
    pass


@dataclass
class Rule:
    selector: str = "*"
    attributes: list[str] = field(default_factory=lambda: ["*"])

    @classmethod
    def from_obj(cls, obj, *, context: str) -> "Rule":
        if isinstance(obj, str):
            return cls(selector=obj)
        if not isinstance(obj, dict):
            raise PolicyError(f"{context}: each rule must be a string or a mapping")
        attrs = obj.get("attributes", ["*"])
        if isinstance(attrs, str):
            attrs = [attrs]
        if not isinstance(attrs, list) or not all(isinstance(a, str) for a in attrs):
            raise PolicyError(f"{context}: 'attributes' must be a list of strings")
        return cls(selector=str(obj.get("selector", "*")), attributes=attrs)


@dataclass
class SemanticConfig:
    mode: str = "off"
    provider: str = "command"
    command: str = ""
    model: str = ""
    checks: list[str] = field(default_factory=list)


@dataclass
class Policy:
    version: int = 1
    target: str = "auto"
    mode: str = "normal_input"
    allow: list[Rule] = field(default_factory=list)
    protect: list[Rule] = field(default_factory=list)
    structural: dict = field(default_factory=dict)
    semantic: SemanticConfig = field(default_factory=SemanticConfig)
    source_path: str = ""

    def structural_setting(self, key: str) -> str:
        """'strict' (default) or 'ignore' for a structural category."""
        value = str(self.structural.get(key, "strict")).lower()
        return value if value in ("strict", "ignore") else "strict"


def load_policy(path: str | Path) -> Policy:
    path = Path(path)
    try:
        raw = yaml.safe_load(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise PolicyError(f"policy file not found: {path}")
    except yaml.YAMLError as exc:
        raise PolicyError(f"invalid YAML in {path}: {exc}")
    if not isinstance(raw, dict):
        raise PolicyError(f"{path}: policy root must be a mapping")
    return parse_policy(raw, source_path=str(path))


def parse_policy(raw: dict, *, source_path: str = "") -> Policy:
    target = str(raw.get("target", "auto")).lower()
    if target not in VALID_TARGETS:
        raise PolicyError(f"target must be one of {VALID_TARGETS}, got {target!r}")
    mode = str(raw.get("mode", "normal_input"))
    if mode not in VALID_MODES:
        raise PolicyError(f"mode must be one of {VALID_MODES}, got {mode!r}")

    allow = [Rule.from_obj(o, context="allow") for o in _as_list(raw.get("allow"), "allow")]
    protect = [Rule.from_obj(o, context="protect") for o in _as_list(raw.get("protect"), "protect")]

    structural = raw.get("structural", {}) or {}
    if not isinstance(structural, dict):
        raise PolicyError("'structural' must be a mapping")

    sem_raw = raw.get("semantic", {}) or {}
    if not isinstance(sem_raw, dict):
        raise PolicyError("'semantic' must be a mapping")
    sem_mode = str(sem_raw.get("mode", "off")).lower()
    if sem_mode not in VALID_SEMANTIC_MODES:
        raise PolicyError(f"semantic.mode must be one of {VALID_SEMANTIC_MODES}, got {sem_mode!r}")
    checks = _as_list(sem_raw.get("checks"), "semantic.checks")
    norm_checks: list[str] = []
    for c in checks:
        if isinstance(c, str):
            norm_checks.append(c)
        elif isinstance(c, dict) and "instruction" in c:
            norm_checks.append(str(c["instruction"]))
        else:
            raise PolicyError("semantic.checks entries must be strings or {instruction: ...}")
    semantic = SemanticConfig(
        mode=sem_mode,
        provider=str(sem_raw.get("provider", "command")),
        command=str(sem_raw.get("command", "")),
        model=str(sem_raw.get("model", "")),
        checks=norm_checks,
    )

    return Policy(
        version=int(raw.get("version", 1)),
        target=target,
        mode=mode,
        allow=allow,
        protect=protect,
        structural=structural,
        semantic=semantic,
        source_path=source_path,
    )


def _as_list(value, name: str) -> list:
    if value is None:
        return []
    if not isinstance(value, list):
        raise PolicyError(f"'{name}' must be a list")
    return value


SAMPLE_POLICY_EXCEL = """\
# OfficeCheck policy — trusted acceptance rules for an Excel edit.
# This file must be authored/approved by a human or pinned in CI,
# never by the agent that edited the document.
version: 1
target: excel
mode: normal_input        # review_only | normal_input | page_extension | row_extension

# Default deny: any change NOT matched below is a violation.
allow:
  - selector: "Sheet1!B2:B100"
    attributes: [value]

# Always-fail guards (redundant with default deny, but explicit and
# they win even if an allow rule overlaps).
protect:
  - selector: "*"
    attributes: [formula, format, merge, conditional_formatting,
                 data_validation, print_settings, header_footer, vba]

structural:
  sheets: strict           # strict | ignore
  images: strict
  defined_names: strict

semantic:
  mode: "off"              # off | review | gate
  provider: command
  command: ""              # e.g. "claude -p" — receives the prompt on stdin
  model: ""
  checks: []
"""

SAMPLE_POLICY_WORD = """\
# OfficeCheck policy — trusted acceptance rules for a Word edit.
version: 1
target: word
mode: normal_input

allow:
  - selector: "p1-20"
    attributes: [text]

protect:
  - selector: "*"
    attributes: [style, section, header_footer]

structural:
  images: strict
  tables: strict

semantic:
  mode: "off"
  provider: command
  command: ""
  model: ""
  checks: []
"""
