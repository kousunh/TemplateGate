import json

from templategate.cli import main


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


def test_corrupt_document_is_execution_error(fixtures, tmp_path, capsys):
    """A file that is not a readable workbook is an error (2), not a FAIL (1)."""
    broken = tmp_path / "broken.xlsx"
    broken.write_bytes(b"this is not an office file")
    code = main([
        "check",
        "--baseline", str(broken),
        "--candidate", str(fixtures["excel_baseline"]),
        "--policy", str(fixtures["excel_policy"]),
    ])
    assert code == 2
    assert "cannot read" in capsys.readouterr().err


def test_truncated_document_is_execution_error(fixtures, tmp_path, capsys):
    data = fixtures["excel_baseline"].read_bytes()
    truncated = tmp_path / "truncated.xlsx"
    truncated.write_bytes(data[: len(data) // 3])
    code = main([
        "check",
        "--baseline", str(truncated),
        "--candidate", str(fixtures["excel_baseline"]),
        "--policy", str(fixtures["excel_policy"]),
    ])
    assert code == 2


def test_review_only_policy_exits_zero(fixtures, tmp_path, capsys):
    policy = tmp_path / "review.yaml"
    policy.write_text("version: 1\ntarget: excel\nmode: review_only\nallow: []\n",
                      encoding="utf-8")
    code = main([
        "check",
        "--baseline", str(fixtures["excel_baseline"]),
        "--candidate", str(fixtures["excel_bad"]),
        "--policy", str(policy),
        "--report", "json",
    ])
    assert code == 0
    data = json.loads(capsys.readouterr().out)
    assert data["passed"] is True
    assert data["summary"]["violations"] > 0
    assert data["summary"]["errors"] == 0
    assert data["summary"]["warnings"] == data["summary"]["violations"]


def test_init_writes_sample_policy(tmp_path, capsys):
    out = tmp_path / "p.yaml"
    assert main(["init", "--target", "word", "--output", str(out)]) == 0
    assert "target: word" in out.read_text(encoding="utf-8")
    # refuses to overwrite
    assert main(["init", "--target", "word", "--output", str(out)]) == 2
