import json

from docgate.cli import main


def test_check_pass_exit_code(fixtures, capsys):
    code = main([
        "check",
        "--baseline", str(fixtures["excel_baseline"]),
        "--candidate", str(fixtures["excel_good"]),
        "--policy", str(fixtures["excel_policy"]),
    ])
    assert code == 0
    assert "PASS" in capsys.readouterr().out


def test_check_fail_exit_code_and_json_report(fixtures, capsys):
    code = main([
        "check",
        "--baseline", str(fixtures["excel_baseline"]),
        "--candidate", str(fixtures["excel_bad"]),
        "--policy", str(fixtures["excel_policy"]),
        "--report", "json",
    ])
    assert code == 1
    data = json.loads(capsys.readouterr().out)
    assert data["passed"] is False


def test_diff_json(fixtures, capsys):
    code = main([
        "diff",
        "--baseline", str(fixtures["word_baseline"]),
        "--candidate", str(fixtures["word_bad"]),
        "--json",
    ])
    assert code == 0
    changes = json.loads(capsys.readouterr().out)
    assert any(c["location"] == "table1!r2c2" for c in changes)


def test_missing_file_is_execution_error(fixtures, capsys):
    code = main([
        "check",
        "--baseline", "nope.xlsx",
        "--candidate", "nope2.xlsx",
        "--policy", str(fixtures["excel_policy"]),
    ])
    assert code == 2


def test_init_writes_sample_policy(tmp_path, capsys):
    out = tmp_path / "p.yaml"
    assert main(["init", "--target", "word", "--output", str(out)]) == 0
    assert "target: word" in out.read_text(encoding="utf-8")
    # refuses to overwrite
    assert main(["init", "--target", "word", "--output", str(out)]) == 2
