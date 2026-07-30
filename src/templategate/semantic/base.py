"""Semantic (AI-based) checks — strictly optional.

Three modes, chosen in the policy (or overridden on the CLI):

  off    — the default.  Document content is never sent to any AI provider;
           only the deterministic structural checks run.
  review — semantic findings are reported as warnings and never affect
           the PASS/FAIL outcome.
  gate   — semantic findings with verdict "fail" cause the check to FAIL.

Providers are pluggable so users choose their own model/vendor; TemplateGate
does not pin one.
"""

from __future__ import annotations

from abc import ABC, abstractmethod

from ..core.model import SemanticFinding
from ..core.policy import SemanticConfig


class SemanticProvider(ABC):
    name = "base"

    @abstractmethod
    def run_checks(
        self,
        config: SemanticConfig,
        baseline_text: str,
        candidate_text: str,
    ) -> list[SemanticFinding]:
        """Run the policy's checks and return findings."""


def get_provider(name: str) -> SemanticProvider:
    from .command import CommandProvider

    providers = {"command": CommandProvider}
    if name not in providers:
        raise ValueError(
            f"unknown semantic provider {name!r} (available: {sorted(providers)})"
        )
    return providers[name]()


def extract_text(snapshot: dict) -> str:
    """Flatten a snapshot into plain text for semantic review."""
    lines: list[str] = []
    if snapshot["target"] == "excel":
        for sheet_name, sheet in snapshot["sheets"].items():
            lines.append(f"# Sheet: {sheet_name}")
            for coord, cell in sheet["cells"].items():
                if cell["value"] is not None:
                    lines.append(f"{coord}: {cell['value']}")
    else:
        for i, para in enumerate(snapshot["paragraphs"], 1):
            if para["text"]:
                lines.append(f"p{i}: {para['text']}")
        for t, table in enumerate(snapshot["tables"], 1):
            for r, row in enumerate(table["rows"], 1):
                for c, text in enumerate(row, 1):
                    if text:
                        lines.append(f"table{t}!r{r}c{c}: {text}")
    return "\n".join(lines)
