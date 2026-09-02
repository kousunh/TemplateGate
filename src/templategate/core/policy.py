"""Policy loading and validation.

A policy is the *trusted* description of what an edit is allowed to change.
It must be authored or approved by a human (or pinned in CI) — never by the
agent that produced the candidate document.  TemplateGate only reads it.
"""

from __future__ import annotations

import difflib
from dataclasses import dataclass, field
from pathlib import Path

import yaml

from .messages import message
from .model import ALL_ATTRIBUTES, STRUCTURAL_ATTRIBUTES

VALID_TARGETS = ("excel", "word", "auto")

VALID_KEYS = ("version", "target", "mode", "allow", "protect",
              "structural", "recalculation", "semantic")
VALID_RULE_KEYS = ("selector", "attributes")
VALID_SEMANTIC_KEYS = ("mode", "provider", "command", "model", "checks")
VALID_STRUCTURAL_VALUES = ("strict", "ignore")
# What to do with a cached formula result that moved while its formula did
# not.  "ignore" (the default) treats it as the arithmetic of whichever tool
# saved the file; "strict" judges it like any other value change, for a
# workflow where the stored answers themselves are the deliverable.
VALID_RECALCULATION_VALUES = ("ignore", "strict")

# review_only    — violations are reported as warnings and never fail the run.
# normal_input   — the default: every violation is an error.
# page_extension — Word only; paragraphs are aligned by content so inserting
#                  one does not report every following paragraph as changed.
VALID_MODES = ("review_only", "normal_input", "page_extension")

VALID_SEMANTIC_MODES = ("off", "review", "gate")

MODE_REVIEW_ONLY = "review_only"
MODE_PAGE_EXTENSION = "page_extension"


class PolicyError(ValueError):
    pass


def _did_you_mean(word: str, known) -> str:
    """" — did you mean 'allow'?", when there is an obvious candidate.

    A misspelled key is not a small problem here: an unknown key is ignored,
    and a policy whose protect rule is spelled `protekt` protects nothing
    while looking like it protects everything.  Naming the likely intent is
    the difference between a two-second fix and a false sense of safety.
    """
    close = difflib.get_close_matches(word, sorted(known), n=1, cutoff=0.7)
    if close:
        return message("policy.did_you_mean", suggestion=close[0])
    return message("policy.valid_are", options=", ".join(sorted(known)))


def _check_known(word: str, known, *, what: str, context: str = "") -> None:
    if word in known:
        return
    raise PolicyError(message("policy.unknown",
                              where=f"{context}: " if context else "",
                              what=message(f"policy.word.{what}"),
                              word=word, hint=_did_you_mean(word, known)))


@dataclass
class Rule:
    selector: str = "*"
    attributes: list[str] = field(default_factory=lambda: ["*"])

    @classmethod
    def from_obj(cls, obj, *, context: str) -> "Rule":
        if isinstance(obj, str):
            return cls(selector=obj)
        if not isinstance(obj, dict):
            raise PolicyError(message("policy.rule_shape", context=context))
        for key in obj:
            _check_known(str(key), VALID_RULE_KEYS, what="key", context=context)
        attrs = obj.get("attributes", ["*"])
        if isinstance(attrs, str):
            attrs = [attrs]
        if not isinstance(attrs, list) or not all(isinstance(a, str) for a in attrs):
            raise PolicyError(message("policy.attributes_shape", context=context))
        for attribute in attrs:
            if attribute == "*":
                continue
            _check_known(attribute, ALL_ATTRIBUTES, what="attribute", context=context)
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
    recalculation: str = "ignore"
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
        raise PolicyError(message("policy.not_found", path=path))
    except yaml.YAMLError as exc:
        raise PolicyError(message("policy.bad_yaml", path=path, reason=exc))
    if not isinstance(raw, dict):
        raise PolicyError(message("policy.root_shape", path=path))
    return parse_policy(raw, source_path=str(path))


def parse_policy(raw: dict, *, source_path: str = "") -> Policy:
    for key in raw:
        _check_known(str(key), VALID_KEYS, what="policy_key")

    target = str(raw.get("target", "auto")).lower()
    if target not in VALID_TARGETS:
        raise PolicyError(message("policy.bad_target", value=target,
                                  hint=_did_you_mean(target, VALID_TARGETS)))
    mode = str(raw.get("mode", "normal_input"))
    if mode not in VALID_MODES:
        raise PolicyError(message("policy.bad_mode", value=mode,
                                  hint=_did_you_mean(mode, VALID_MODES)))

    allow = [Rule.from_obj(o, context="allow") for o in _as_list(raw.get("allow"), "allow")]
    protect = [Rule.from_obj(o, context="protect") for o in _as_list(raw.get("protect"), "protect")]

    structural = raw.get("structural", {}) or {}
    if not isinstance(structural, dict):
        raise PolicyError(message("policy.structural_shape"))
    for key, value in structural.items():
        _check_known(str(key), STRUCTURAL_ATTRIBUTES, what="structural_key")
        if str(value).lower() not in VALID_STRUCTURAL_VALUES:
            raise PolicyError(message(
                "policy.bad_structural", key=key, value=value,
                hint=_did_you_mean(str(value).lower(),
                                   VALID_STRUCTURAL_VALUES)))

    recalculation = str(raw.get("recalculation", "ignore")).lower()
    if recalculation not in VALID_RECALCULATION_VALUES:
        raise PolicyError(message(
            "policy.bad_recalculation", value=raw.get("recalculation"),
            hint=_did_you_mean(recalculation, VALID_RECALCULATION_VALUES)))

    sem_raw = raw.get("semantic", {}) or {}
    if not isinstance(sem_raw, dict):
        raise PolicyError(message("policy.semantic_shape"))
    for key in sem_raw:
        _check_known(str(key), VALID_SEMANTIC_KEYS, what="key", context="semantic")
    sem_mode = str(sem_raw.get("mode", "off")).lower()
    if sem_mode not in VALID_SEMANTIC_MODES:
        raise PolicyError(message(
            "policy.bad_semantic_mode", value=sem_mode,
            hint=_did_you_mean(sem_mode, VALID_SEMANTIC_MODES)))
    checks = _as_list(sem_raw.get("checks"), "semantic.checks")
    norm_checks: list[str] = []
    for c in checks:
        if isinstance(c, str):
            norm_checks.append(c)
        elif isinstance(c, dict) and "instruction" in c:
            norm_checks.append(str(c["instruction"]))
        else:
            raise PolicyError(message("policy.bad_check"))
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
        recalculation=recalculation,
        semantic=semantic,
        source_path=source_path,
    )


def _as_list(value, name: str) -> list:
    if value is None:
        return []
    if not isinstance(value, list):
        raise PolicyError(message("policy.must_be_list", name=name))
    return value


SAMPLE_POLICY_EXCEL = """\
# TemplateGate policy — trusted acceptance rules for an Excel edit.
# This file must be authored/approved by a human or pinned in CI,
# never by the agent that edited the document.
version: 1
target: excel
mode: normal_input        # review_only | normal_input | page_extension

# Default deny: any change NOT matched below is a violation.
allow:
  - selector: "Sheet1!B2:B100"
    attributes: [value]

# Always-fail guards (redundant with default deny, but explicit and
# they win even if an allow rule overlaps).
protect:
  - selector: "*"
    attributes: [formula, format, merge, conditional_formatting,
                 data_validation, print_settings, header_footer, vba,
                 protection, sheet_settings, layout]

structural:
  sheets: strict           # strict | ignore
  images: strict
  defined_names: strict
  # OOXML parts that editing libraries drop on save without warning.
  charts: strict
  pivot_tables: strict
  drawings: strict         # textboxes, shapes and chart frames
  comments: strict
  embedded: strict         # embedded/OLE objects
  custom_xml: strict
  parts: strict            # every other package part, including unknown ones
  links: strict            # external hyperlink and reference targets

# What to do when a formula's cached result moves but the formula itself does
# not — Excel recomputing on save, or a library discarding the answers.
#   ignore (default) — not a change to the document's logic, so not a
#                      violation.  `formula` still catches a formula replaced
#                      by a literal.
#   strict           — judge those values like any other, for a workbook whose
#                      stored answers are themselves the deliverable.
recalculation: ignore      # ignore | strict

semantic:
  mode: "off"              # off | review | gate
  provider: command
  command: ""              # e.g. "claude -p" — receives the prompt on stdin
  model: ""
  checks: []
"""

SAMPLE_POLICY_WORD = """\
# TemplateGate policy — trusted acceptance rules for a Word edit.
version: 1
target: word
mode: normal_input        # review_only | normal_input | page_extension

allow:
  - selector: "p1-20"
    attributes: [text]

protect:
  # "body"/"p<N>" cover body paragraphs only; content controls and text boxes
  # are addressed as sdt<N> / textbox<N>, so they are protected by "*" here.
  - selector: "*"
    attributes: [style, format, paragraph_format, section, header_footer,
                 field, bookmark, revision, content_control, table, moved,
                 markup]

structural:
  images: strict
  tables: strict
  # OOXML parts that editing libraries drop on save without warning.
  charts: strict
  comments: strict
  embedded: strict         # embedded/OLE objects
  custom_xml: strict
  parts: strict            # styles, numbering, footnotes, settings, theme...
  links: strict            # external hyperlink targets
  # No pivot_tables or drawings here: those are Excel part families.  Word
  # keeps its shapes and text boxes inside document.xml, where they are
  # covered by the text, markup and images attributes instead.

semantic:
  mode: "off"
  provider: command
  command: ""
  model: ""
  checks: []
"""
