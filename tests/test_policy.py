import pytest

from templategate.core.evaluator import evaluate
from templategate.core.model import Change
from templategate.core.policy import PolicyError, parse_policy


def test_parse_minimal_policy():
    policy = parse_policy({"target": "excel"})
    assert policy.target == "excel"
    assert policy.semantic.mode == "off"
    assert policy.allow == []


def test_invalid_semantic_mode_rejected():
    with pytest.raises(PolicyError):
        parse_policy({"semantic": {"mode": "always"}})


def test_invalid_target_rejected():
    with pytest.raises(PolicyError):
        parse_policy({"target": "powerpoint"})


def test_default_deny():
    policy = parse_policy({
        "target": "excel",
        "allow": [{"selector": "Sheet1!B2:B10", "attributes": ["value"]}],
    })
    changes = [
        Change("Sheet1!B3", "value", old=1, new=2),
        Change("Sheet1!C3", "value", old=1, new=2),
        Change("Sheet1!B3", "format"),
    ]
    allowed, violations = evaluate(changes, policy)
    assert [c.location for c in allowed] == ["Sheet1!B3"]
    assert {(v.change.location, v.change.attribute) for v in violations} == {
        ("Sheet1!C3", "value"),
        ("Sheet1!B3", "format"),
    }
    assert all(v.rule == "not_allowed" for v in violations)


def test_protect_beats_allow():
    policy = parse_policy({
        "target": "excel",
        "allow": [{"selector": "*", "attributes": ["*"]}],
        "protect": [{"selector": "*", "attributes": ["formula"]}],
    })
    changes = [Change("Sheet1!D8", "formula", old="=SUM(A1:A2)", new=None)]
    allowed, violations = evaluate(changes, policy)
    assert not allowed
    assert violations[0].rule == "protected"


def test_structural_ignore_drops_changes():
    policy = parse_policy({
        "target": "excel",
        "structural": {"sheets": "ignore"},
    })
    changes = [Change("sheet:Temp", "sheet_structure", detail="sheet added")]
    allowed, violations = evaluate(changes, policy)
    assert not allowed and not violations
