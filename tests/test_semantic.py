"""Semantic provider: command splitting, environment and output parsing."""

import json
import sys

from templategate.core.policy import SemanticConfig
from templategate.semantic.command import CommandProvider, split_command


def test_windows_split_keeps_backslashes():
    """POSIX splitting would turn this into 'C:toolsclaude.exe'."""
    assert split_command(r"C:\tools\claude.exe -p", windows=True) == [
        r"C:\tools\claude.exe", "-p",
    ]


def test_windows_split_strips_surrounding_quotes():
    assert split_command(r'"C:\Program Files\cli\claude.exe" -p', windows=True) == [
        r"C:\Program Files\cli\claude.exe", "-p",
    ]


def test_posix_split_is_unchanged():
    assert split_command("claude -p --json", windows=False) == [
        "claude", "-p", "--json",
    ]


def test_blank_command_reports_configuration_error():
    """A whitespace-only command splits to argv [] and used to raise."""
    findings = CommandProvider().run_checks(
        SemanticConfig(mode="review", command="   "), "base", "cand"
    )
    assert [f.verdict for f in findings] == ["error"]
    assert findings[0].check == "(configuration)"


def _provider_script(tmp_path, body):
    script = tmp_path / "provider.py"
    script.write_text("import json, os, sys\nsys.stdin.read()\n" + body,
                      encoding="utf-8")
    return f'"{sys.executable}" "{script}"'


def test_model_is_exported_as_templategate_model(tmp_path):
    command = _provider_script(
        tmp_path,
        "print(json.dumps([{'check': os.environ.get('TEMPLATEGATE_MODEL', '(unset)'),"
        " 'verdict': 'pass'}]))\n",
    )
    findings = CommandProvider().run_checks(
        SemanticConfig(mode="review", command=command, model="my-model"),
        "base", "cand",
    )
    assert [f.check for f in findings] == ["my-model"]


def test_provider_findings_are_parsed(tmp_path):
    command = _provider_script(
        tmp_path,
        "print(json.dumps([{'check': 'dates', 'verdict': 'fail', 'message': 'bad'}]))\n",
    )
    findings = CommandProvider().run_checks(
        SemanticConfig(mode="gate", command=command), "base", "cand"
    )
    assert [(f.check, f.verdict, f.message) for f in findings] == [
        ("dates", "fail", "bad")
    ]


def test_nonzero_exit_is_reported_as_error(tmp_path):
    command = _provider_script(tmp_path, "sys.exit(3)\n")
    findings = CommandProvider().run_checks(
        SemanticConfig(mode="review", command=command), "base", "cand"
    )
    assert findings[0].verdict == "error"
    assert "exited 3" in findings[0].message


def test_gate_mode_fails_the_check(fixtures, tmp_path):
    """A failing semantic verdict flips PASS to FAIL under mode: gate."""
    from templategate import check
    from templategate.core.policy import parse_policy

    command = _provider_script(
        tmp_path,
        "print(json.dumps([{'check': 'c', 'verdict': 'fail', 'message': 'no'}]))\n",
    )
    policy = parse_policy({
        "target": "excel",
        "allow": [{"selector": "*", "attributes": ["*"]}],
        "semantic": {"mode": "gate", "command": command},
    })
    result = check(fixtures["excel_baseline"], fixtures["excel_good"], policy)
    assert not result.violations
    assert not result.passed


def test_review_only_policy_keeps_semantic_gate_non_blocking(fixtures, tmp_path):
    from templategate import check
    from templategate.core.policy import parse_policy

    command = _provider_script(
        tmp_path,
        "print(json.dumps([{'check': 'c', 'verdict': 'fail', 'message': 'no'}]))\n",
    )
    policy = parse_policy({
        "target": "excel",
        "mode": "review_only",
        "allow": [{"selector": "*", "attributes": ["*"]}],
        "semantic": {"mode": "gate", "command": command},
    })
    result = check(fixtures["excel_baseline"], fixtures["excel_good"], policy)
    assert result.passed
    assert [f.verdict for f in result.semantic_findings] == ["fail"]


def test_document_text_is_not_passed_on_argv(tmp_path):
    """Content must reach the provider on stdin, never as a command argument."""
    command = _provider_script(
        tmp_path,
        "print(json.dumps([{'check': ' '.join(sys.argv[1:]), 'verdict': 'pass'}]))\n",
    )
    findings = CommandProvider().run_checks(
        SemanticConfig(mode="review", command=command), "SECRET-BASE", "SECRET-CAND"
    )
    assert "SECRET" not in findings[0].check


def test_parse_ignores_prose_around_the_json_array(tmp_path):
    command = _provider_script(
        tmp_path,
        "print('here you go:')\n"
        "print(json.dumps([{'check': 'c', 'verdict': 'warning'}]))\n"
        "print('hope that helps')\n",
    )
    findings = CommandProvider().run_checks(
        SemanticConfig(mode="review", command=command), "base", "cand"
    )
    assert [f.verdict for f in findings] == ["warning"]


def test_unparseable_output_is_an_error(tmp_path):
    command = _provider_script(tmp_path, "print('no json here')\n")
    findings = CommandProvider().run_checks(
        SemanticConfig(mode="review", command=command), "base", "cand"
    )
    assert findings[0].verdict == "error"
    assert findings[0].check == "(parse)"


def test_unknown_verdict_is_downgraded_to_warning(tmp_path):
    command = _provider_script(
        tmp_path,
        "print(json.dumps([{'check': 'c', 'verdict': 'CATASTROPHE'}]))\n",
    )
    findings = CommandProvider().run_checks(
        SemanticConfig(mode="review", command=command), "base", "cand"
    )
    assert [f.verdict for f in findings] == ["warning"]
    assert json.loads(json.dumps(findings[0].to_dict()))["verdict"] == "warning"
