"""DocGate command-line interface.

Exit codes: 0 = PASS, 1 = FAIL, 2 = execution/configuration error.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from . import __version__
from .api import check, detect_target, diff, snapshot
from .core.policy import PolicyError, SAMPLE_POLICY_EXCEL, SAMPLE_POLICY_WORD
from .reporters import RENDERERS


def main(argv: list[str] | None = None) -> int:
    # Windows consoles often default to a legacy codepage; reports contain
    # non-ASCII sheet names and cell values.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(
        prog="docgate",
        description=(
            "Policy-as-code acceptance gate for AI-edited Office documents. "
            "Verifies that only intended changes were made to .xlsx/.xlsm/.docx files."
        ),
    )
    parser.add_argument("--version", action="version", version=f"docgate {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    p_check = sub.add_parser("check", help="check a candidate against a baseline using a policy")
    p_check.add_argument("--baseline", required=True)
    p_check.add_argument("--candidate", required=True)
    p_check.add_argument("--policy", required=True)
    p_check.add_argument("--report", choices=sorted(RENDERERS), default="text")
    p_check.add_argument("--output", help="write the report to a file instead of stdout")
    p_check.add_argument("--semantic", choices=["off", "review", "gate"],
                         help="override the policy's semantic mode")

    p_diff = sub.add_parser("diff", help="list all changes without a policy")
    p_diff.add_argument("--baseline", required=True)
    p_diff.add_argument("--candidate", required=True)
    p_diff.add_argument("--json", action="store_true", help="output JSON")

    p_snap = sub.add_parser("snapshot", help="dump a document's structural snapshot as JSON")
    p_snap.add_argument("file")
    p_snap.add_argument("--output")

    p_init = sub.add_parser("init", help="write a sample policy file")
    p_init.add_argument("--target", choices=["excel", "word"], default="excel")
    p_init.add_argument("--output", default="docgate.policy.yaml")

    args = parser.parse_args(argv)
    try:
        return _dispatch(args)
    except PolicyError as exc:
        print(f"docgate: policy error: {exc}", file=sys.stderr)
        return 2
    except (ValueError, OSError) as exc:
        print(f"docgate: error: {exc}", file=sys.stderr)
        return 2


def _dispatch(args: argparse.Namespace) -> int:
    if args.command == "check":
        result = check(args.baseline, args.candidate, args.policy,
                       semantic_mode=args.semantic)
        report = RENDERERS[args.report](result)
        if args.output:
            Path(args.output).write_text(report, encoding="utf-8")
            print(f"report written to {args.output}")
        else:
            print(report)
        return 0 if result.passed else 1

    if args.command == "diff":
        changes = diff(args.baseline, args.candidate)
        if args.json:
            print(json.dumps([c.to_dict() for c in changes],
                             ensure_ascii=False, indent=2, default=str))
        elif not changes:
            print("no changes")
        else:
            for c in changes:
                line = f"{c.location} ({c.attribute}): {c.old!r} -> {c.new!r}"
                if c.detail:
                    line += f"  [{c.detail}]"
                print(line)
        return 0

    if args.command == "snapshot":
        snap = snapshot(args.file)
        text = json.dumps(snap, ensure_ascii=False, indent=2, default=str)
        if args.output:
            Path(args.output).write_text(text, encoding="utf-8")
            print(f"snapshot written to {args.output}")
        else:
            print(text)
        return 0

    if args.command == "init":
        sample = SAMPLE_POLICY_EXCEL if args.target == "excel" else SAMPLE_POLICY_WORD
        path = Path(args.output)
        if path.exists():
            print(f"docgate: {path} already exists, not overwriting", file=sys.stderr)
            return 2
        path.write_text(sample, encoding="utf-8")
        print(f"sample {args.target} policy written to {path}")
        return 0

    raise AssertionError(f"unhandled command {args.command}")


if __name__ == "__main__":
    sys.exit(main())
