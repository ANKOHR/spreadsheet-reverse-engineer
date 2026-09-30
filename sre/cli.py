"""Deterministic CLI for read-only spreadsheet analysis."""

import argparse
import json
import sys
from pathlib import Path

from .analysis import BANNER, analyze
from .fixtures import generate


def serialize(value):
    return json.dumps(value, indent=2, sort_keys=True, ensure_ascii=True, allow_nan=False) + "\n"


def markdown(result):
    lines = ["# Spreadsheet Reverse Engineer", "", f"> {BANNER}", "",
             "Analysis documents formula structure without evaluating or certifying results.", "",
             "## Workbook summary", "", "```json", serialize(result["summary"]).rstrip(), "```", "",
             "## Rule catalogue", ""]
    for rule in result["rules"]:
        lines += [f"### {rule['id']}: {rule['kind']}", "", rule["description"], "",
                  "Cells: " + ", ".join(rule["cells"]), ""]
    lines += ["## Smell list", ""]
    if not result["smells"]:
        lines += ["No smells found under the declared structural heuristics. This is not a correctness claim.", ""]
    for smell in result["smells"]:
        lines += [f"### {smell['id']}: {smell['kind']}", "", smell["explanation"], "",
                  "Cells: " + (", ".join(smell["cells"]) or "workbook-level finding"), ""]
        if "path" in smell:
            lines += ["Cycle: " + " -> ".join(smell["path"]), ""]
    lines += ["## Migration plan", "", result["plan"]["disclaimer"], "",
              f"Manual review required: {result['plan']['manual_review_required']}", ""]
    for phase in result["plan"]["phases"]:
        lines += [f"### Phase {phase['phase']}", "", phase["action"], "",
                  "Cells: " + ", ".join(phase["cells"]), "",
                  "Rules: " + ", ".join(phase["rules"]), "",
                  "Smells to address: " + ", ".join(phase["smells_to_address"]), "",
                  "Blocked cells: " + ", ".join(phase["blocked_cells"]), ""]
    lines += ["Unassigned workbook findings: " + ", ".join(result["plan"]["unassigned_smells"]), ""]
    # Escape user-provided em dash characters in reports as well as JSON.
    return "\n".join(lines).replace(chr(0x2014), "[U+2014]")


def main(argv=None):
    parser = argparse.ArgumentParser(description=BANNER)
    commands = parser.add_subparsers(dest="command", required=True)
    fixtures = commands.add_parser("fixtures", help="Generate labelled fictional workbooks")
    fixtures.add_argument("--out", required=True)
    for command in ("parse", "graph", "rules", "smells", "plan", "report"):
        sub = commands.add_parser(command)
        sub.add_argument("file", type=Path)
        sub.add_argument("--out", type=Path)
        if command == "report":
            sub.add_argument("--format", choices=("md", "json"), default="md")
    args = parser.parse_args(argv)
    try:
        if args.command == "fixtures":
            print(serialize({"banner": BANNER, "files": generate(args.out)}), end="")
            return 0
        # Check before analysis and before any output write to protect the input.
        if args.out and args.out.resolve() == args.file.resolve():
            parser.error("Output must not overwrite the input workbook")
        if args.out and args.out.exists() and args.file.exists() and args.out.samefile(args.file):
            parser.error("Output must not overwrite an alias of the input workbook")
        result = analyze(args.file)
        if args.command == "report":
            rendered = markdown(result) if args.format == "md" else serialize(result)
        else:
            key = "inventory" if args.command == "parse" else args.command
            rendered = serialize({"banner": BANNER, key: result[key]})
        if args.out:
            args.out.write_text(rendered, encoding="utf-8", newline="\n")
        else:
            sys.stdout.write(rendered)
        return 0
    except (OSError, ValueError, KeyError) as exc:
        print(f"sre: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
