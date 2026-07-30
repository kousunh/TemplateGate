"""Semantic provider that shells out to a user-configured command.

The policy sets e.g.::

    semantic:
      mode: review
      provider: command
      command: "claude -p"      # any CLI that reads stdin, prints JSON
      model: ""                  # appended as $OFFICEGUARD_MODEL env var

The command receives a prompt on stdin and must print a JSON array:
``[{"check": "...", "verdict": "pass|fail|warning", "message": "..."}]``.
This keeps the model/vendor choice entirely in the user's hands.
"""

from __future__ import annotations

import json
import os
import shlex
import subprocess

from ..core.model import SemanticFinding
from ..core.policy import SemanticConfig
from .base import SemanticProvider

_PROMPT_TEMPLATE = """\
You are a document reviewer. Compare the baseline and candidate text of an
Office document and evaluate ONLY the checks listed below. Respond with a
JSON array only, no prose: [{{"check": "<check>", "verdict": "pass|fail|warning",
"message": "<short reason>"}}]

## Checks
{checks}

## Baseline text
{baseline}

## Candidate text
{candidate}
"""


class CommandProvider(SemanticProvider):
    name = "command"

    def run_checks(
        self,
        config: SemanticConfig,
        baseline_text: str,
        candidate_text: str,
    ) -> list[SemanticFinding]:
        if not config.command:
            return [SemanticFinding(
                check="(configuration)",
                verdict="error",
                message="semantic.command is empty; set a command or use mode: off",
            )]
        prompt = _PROMPT_TEMPLATE.format(
            checks="\n".join(f"- {c}" for c in config.checks) or "- (none)",
            baseline=baseline_text,
            candidate=candidate_text,
        )
        env = dict(os.environ)
        if config.model:
            env["OFFICEGUARD_MODEL"] = config.model
        try:
            proc = subprocess.run(
                shlex.split(config.command),
                input=prompt,
                capture_output=True,
                text=True,
                encoding="utf-8",
                env=env,
                timeout=600,
            )
        except (OSError, subprocess.TimeoutExpired) as exc:
            return [SemanticFinding(check="(execution)", verdict="error",
                                    message=f"semantic command failed: {exc}")]
        if proc.returncode != 0:
            return [SemanticFinding(
                check="(execution)", verdict="error",
                message=f"semantic command exited {proc.returncode}: {proc.stderr[:500]}",
            )]
        return _parse_findings(proc.stdout)


def _parse_findings(stdout: str) -> list[SemanticFinding]:
    text = stdout.strip()
    start, end = text.find("["), text.rfind("]")
    if start == -1 or end == -1:
        return [SemanticFinding(check="(parse)", verdict="error",
                                message=f"no JSON array in provider output: {text[:200]}")]
    try:
        items = json.loads(text[start:end + 1])
    except json.JSONDecodeError as exc:
        return [SemanticFinding(check="(parse)", verdict="error",
                                message=f"invalid JSON from provider: {exc}")]
    findings = []
    for item in items:
        if not isinstance(item, dict):
            continue
        verdict = str(item.get("verdict", "warning")).lower()
        if verdict not in ("pass", "fail", "warning"):
            verdict = "warning"
        findings.append(SemanticFinding(
            check=str(item.get("check", "")),
            verdict=verdict,
            message=str(item.get("message", "")),
        ))
    return findings
