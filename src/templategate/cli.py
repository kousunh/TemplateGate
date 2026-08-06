"""TemplateGate command-line interface.

Exit codes: 0 = PASS, 1 = FAIL, 2 = execution/configuration error.

A document that opens as a package but has lost a part it still references is
damage, not an error: `check` reports the missing parts and exits 1, because
default deny means a document nobody could verify does not pass.  `diff` makes
no such judgement, so for it the same document is exit 2 — it could not do the
comparison it was asked for.  Exit 2 also covers files that are not readable
containers at all, and policy and IO problems.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import warnings
from pathlib import Path

from . import __version__
from .api import check, diff_report, snapshot
from .core.policy import PolicyError, SAMPLE_POLICY_EXCEL, SAMPLE_POLICY_WORD
from .core.messages import LANGUAGES, Translator, normalise, say
from .reporters import RENDERERS


def main(argv: list[str] | None = None) -> int:
    # openpyxl warns "... extension is not supported and will be removed" for
    # the extLst blocks it cannot model — sparklines, x14 rules, slicers.
    # Those are exactly the blocks the package layer hashes and compares
    # itself, so the warning would tell a reader those features went
    # unchecked, which is the opposite of what happens.  CLI only: the
    # library API leaves warning policy to the caller.
    warnings.filterwarnings(
        "ignore", message=r".*extension is not supported and will be removed",
        category=UserWarning, module=r"openpyxl\.")
    # Windows consoles often default to a legacy codepage; reports contain
    # non-ASCII sheet names and cell values.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(
        prog="templategate",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description=(
            "Policy-as-code acceptance gate for AI-edited Office documents. "
            "Verifies that only intended changes were made to .xlsx/.xlsm/.docx files."
        ),
        epilog=(
            "Exit codes:\n"
            "  0  the candidate obeys the policy\n"
            "  1  it does not, or the document is damaged\n"
            "  2  the tool could not run the check (bad policy, unreadable file)\n"
            "\n"
            "Start with:  templategate init --target excel"
        ),
    )
    parser.add_argument("--version", action="version", version=f"templategate {__version__}")
    sub = parser.add_subparsers(dest="command", required=True, metavar="COMMAND")

    baseline_help = "the original, untouched document (the edit's starting point)"
    candidate_help = "the edited document to judge"

    p_check = sub.add_parser(
        "check", help="check a candidate against a baseline using a policy",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description="Judge an edited document against the policy you trust.",
        epilog=("Examples:\n"
                "  templategate check --baseline plan.xlsx --candidate plan.edited.xlsx \\\n"
                "      --policy templategate.policy.yaml\n"
                "  templategate check --baseline a.docx --candidate b.docx \\\n"
                "      --policy p.yaml --report markdown --output report.md"))
    p_check.add_argument("--baseline", required=True, metavar="FILE", help=baseline_help)
    p_check.add_argument("--candidate", required=True, metavar="FILE", help=candidate_help)
    p_check.add_argument("--policy", required=True, metavar="FILE",
                         help="the trusted policy that says what may change")
    p_check.add_argument("--report", choices=sorted(RENDERERS), default="text",
                         help="report format (default: text)")
    p_check.add_argument("--output", metavar="FILE",
                         help="write the report to a file instead of stdout")
    p_check.add_argument("--semantic", choices=["off", "review", "gate"],
                         help="override the policy's semantic mode for this run")
    p_check.add_argument("--lang", choices=list(LANGUAGES),
                         help="language for the human-readable report "
                              "(default: en, or TEMPLATEGATE_LANG). "
                              "The JSON report is always English.")

    p_diff = sub.add_parser(
        "diff", help="list every change, with no policy applied",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description="Show what changed between two documents, judging nothing. "
                    "Useful for writing a policy in the first place.",
        epilog=("Examples:\n"
                "  templategate diff --baseline plan.xlsx --candidate plan.edited.xlsx\n"
                "  templategate diff --baseline a.docx --candidate b.docx --json"))
    p_diff.add_argument("--baseline", required=True, metavar="FILE", help=baseline_help)
    p_diff.add_argument("--candidate", required=True, metavar="FILE", help=candidate_help)
    p_diff.add_argument("--json", action="store_true",
                        help="output JSON instead of one line per change")
    p_diff.add_argument("--lang", choices=list(LANGUAGES),
                        help="language for the human-readable output "
                             "(default: en, or TEMPLATEGATE_LANG). "
                             "--json output is always English.")

    p_snap = sub.add_parser(
        "snapshot", help="dump a document's structural snapshot as JSON",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description="Print everything TemplateGate can see in one document.",
        epilog="Examples:\n  templategate snapshot plan.xlsx --output plan.snapshot.json")
    p_snap.add_argument("file", metavar="FILE",
                        help="the .xlsx/.xlsm/.docx file to inspect")
    p_snap.add_argument("--output", metavar="FILE",
                        help="write the JSON to a file instead of stdout")

    p_suggest = sub.add_parser(
        "suggest", help="draft a policy from an edit you have already approved",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description="Turn one reviewed edit into a starting policy. Give it a "
                    "baseline and a candidate you have checked by eye; it "
                    "allows what that edit did, explains the collateral every "
                    "save produces, and refuses to allow anything suspicious.",
        epilog=("Examples:\n"
                "  templategate suggest --baseline plan.xlsx --candidate plan.edited.xlsx\n"
                "  templategate suggest --baseline a.docx --candidate b.docx"
                " --output .templategate/draft.yaml\n"
                "  templategate suggest --baseline a.xlsx --candidate b.xlsx"
                " --policy existing.yaml"))
    p_suggest.add_argument("--baseline", required=True, metavar="FILE",
                           help=baseline_help)
    p_suggest.add_argument("--candidate", required=True, metavar="FILE",
                           help="a candidate whose changes you have reviewed "
                                "and consider correct")
    p_suggest.add_argument("--policy", metavar="FILE",
                           help="an existing policy to propose additions to, "
                                "rather than drafting from scratch")
    p_suggest.add_argument("--output", metavar="FILE",
                           help="write the draft to a file instead of stdout")
    p_suggest.add_argument("--force", action="store_true",
                           help="overwrite the output file if it already exists")

    p_init = sub.add_parser(
        "init", help="write a starter policy file",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description="Write a commented starter policy to edit and then pin in CI.",
        epilog=("Examples:\n"
                "  templategate init --target excel\n"
                "  templategate init --target word --output .templategate/contract.yaml"))
    p_init.add_argument("--target", choices=["excel", "word"], default="excel",
                        help="which format the starter policy is for (default: excel)")
    p_init.add_argument("--output", default="templategate.policy.yaml", metavar="FILE",
                        help="where to write it (default: templategate.policy.yaml)")

    args = parser.parse_args(argv)
    # Resolved before dispatch so a policy that will not parse is still
    # complained about in the reader's language.  The exception types and the
    # exit codes do not move: only the words do.
    t = Translator(_language(args))
    try:
        return _dispatch(args)
    except PolicyError as exc:
        print(t("cli.policy_error", reason=say(exc.args[0] if exc.args else exc, t)),
              file=sys.stderr)
        return 2
    except (ValueError, OSError) as exc:
        print(t("cli.error", reason=say(exc.args[0] if exc.args else exc, t)),
              file=sys.stderr)
        return 2


def _language(args: argparse.Namespace) -> str:
    """The flag if given, else the environment, else English.

    An environment variable is how a Japanese office sets this once for a
    whole CI job; the flag has to win so a single run can still be read by
    somebody else.
    """
    chosen = getattr(args, "lang", None)
    if chosen:
        return normalise(chosen)
    return normalise(os.environ.get("TEMPLATEGATE_LANG"))


def _dispatch(args: argparse.Namespace) -> int:
    if args.command == "check":
        lang = _language(args)
        t = Translator(lang)
        result = check(args.baseline, args.candidate, args.policy,
                       semantic_mode=args.semantic)
        for warning in result.warnings:
            print(t("cli.warning", message=say(warning, t)), file=sys.stderr)
        # The JSON report is a machine contract: it stays English whatever
        # the reader asked for, so agents and CI parse one stable shape.
        report = (RENDERERS[args.report](result) if args.report == "json"
                  else RENDERERS[args.report](result, lang))
        if args.output:
            Path(args.output).write_text(report, encoding="utf-8")
            print(t("cli.report_written", path=args.output))
        else:
            print(report)
        return 0 if result.passed else 1

    if args.command == "diff":
        t = Translator(_language(args))
        changes, degraded = diff_report(args.baseline, args.candidate)
        for role, reason in degraded.items():
            print(t("cli.warning", message=t("damaged.line",
                                             role=t(f"role.{role}"),
                                             reason=reason)),
                  file=sys.stderr)
            print(t("cli.warning", message=t("damaged.parts_only.long")),
                  file=sys.stderr)
        if args.json:
            print(json.dumps([c.to_dict() for c in changes],
                             ensure_ascii=False, indent=2, default=str))
        elif not changes:
            # Never a bare "no changes" for a document we could not read: the
            # whole point of running diff is to find out what moved, and
            # silence would read as "nothing did".
            print(t("cli.no_changes.degraded") if degraded
                  else t("cli.no_changes"))
        else:
            for c in changes:
                line = f"{c.location} ({c.attribute}): {c.old!r} -> {c.new!r}"
                detail = say(c.detail, t)
                if detail:
                    line += f"  [{detail}]"
                print(line)
        return 2 if degraded else 0

    if args.command == "snapshot":
        snap = snapshot(args.file)
        text = json.dumps(snap, ensure_ascii=False, indent=2, default=str)
        if args.output:
            Path(args.output).write_text(text, encoding="utf-8")
            print(Translator(_language(args))("cli.snapshot_written", path=args.output))
        else:
            print(text)
        return 0

    if args.command == "suggest":
        from .suggest import draft

        text = draft(args.baseline, args.candidate,
                     existing=Path(args.policy) if args.policy else None)
        if not args.output:
            print(text, end="")
            return 0
        path = Path(args.output)
        if path.exists() and not args.force:
            print(f"templategate: {path} already exists, not overwriting "
                  "(use --force)", file=sys.stderr)
            return 2
        path.write_text(text, encoding="utf-8")
        print(f"policy draft written to {path}")
        print("templategate: review it before pinning it in CI — a draft only "
              "knows the one edit it was shown", file=sys.stderr)
        return 0

    if args.command == "init":
        sample = SAMPLE_POLICY_EXCEL if args.target == "excel" else SAMPLE_POLICY_WORD
        path = Path(args.output)
        if path.exists():
            print(f"templategate: {path} already exists, not overwriting", file=sys.stderr)
            return 2
        path.write_text(sample, encoding="utf-8")
        print(f"sample {args.target} policy written to {path}")
        return 0

    raise AssertionError(f"unhandled command {args.command}")


if __name__ == "__main__":
    sys.exit(main())
