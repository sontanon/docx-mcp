"""Command-line interface for docx-mcp.

Subcommands:

* ``apply``    — Apply changes JSON to a .docx, producing a redlined output.
* ``convert``  — Extract fragment text from a .docx (pseudo-Markdown).
* ``validate`` — Validate a redlined .docx for structural correctness.

Usage::

    docx-mcp apply input.docx changes.json -o output.docx
    docx-mcp convert input.docx
    docx-mcp validate redlined.docx
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from docx_mcp.converter import document_to_fragments, fragments_to_tagged_text
from docx_mcp.document import DocxDocument
from docx_mcp.models import Change, ChangeType, RedlineConfig
from docx_mcp.redliner import apply_redlines
from docx_mcp.validator import validate_document


def main(argv: list[str] | None = None) -> int:
    """Entry point for the ``docx-mcp`` CLI."""
    parser = argparse.ArgumentParser(
        prog="docx-mcp",
        description="Legal document redlining engine",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # --- apply ---
    apply_parser = subparsers.add_parser(
        "apply",
        help="Apply tracked changes to a .docx file",
    )
    apply_parser.add_argument("input", type=Path, help="Input .docx file")
    apply_parser.add_argument("changes", type=Path, help="Changes JSON file")
    apply_parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Output .docx file (default: <input>_redlined.docx)",
    )
    apply_parser.add_argument(
        "--author",
        default="AI Review",
        help="Author name for tracked changes (default: 'AI Review')",
    )
    apply_parser.add_argument(
        "--validate",
        action="store_true",
        default=True,
        help="Validate output before saving (default: True)",
    )
    apply_parser.add_argument(
        "--no-validate",
        action="store_false",
        dest="validate",
        help="Skip validation",
    )

    # --- convert ---
    convert_parser = subparsers.add_parser(
        "convert",
        help="Extract fragment text from a .docx file",
    )
    convert_parser.add_argument("input", type=Path, help="Input .docx file")
    convert_parser.add_argument(
        "--format",
        choices=["tagged", "json"],
        default="tagged",
        help="Output format: 'tagged' (human-readable) or 'json' (default: tagged)",
    )

    # --- validate ---
    validate_parser = subparsers.add_parser(
        "validate",
        help="Validate a redlined .docx for structural correctness",
    )
    validate_parser.add_argument("input", type=Path, help="Redlined .docx file to validate")

    args = parser.parse_args(argv)

    if args.command == "apply":
        return _cmd_apply(args)
    if args.command == "convert":
        return _cmd_convert(args)
    if args.command == "validate":
        return _cmd_validate(args)

    parser.print_help()
    return 1


def _cmd_apply(args: argparse.Namespace) -> int:
    """Execute the ``apply`` subcommand."""
    input_path: Path = args.input
    changes_path: Path = args.changes

    if not input_path.exists():
        print(f"Error: input file not found: {input_path}", file=sys.stderr)
        return 1
    if not changes_path.exists():
        print(f"Error: changes file not found: {changes_path}", file=sys.stderr)
        return 1

    # Parse changes JSON
    try:
        raw = json.loads(changes_path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as e:
        print(f"Error reading changes file: {e}", file=sys.stderr)
        return 1

    try:
        changes = _parse_changes(raw)
    except (ValueError, KeyError) as e:
        print(f"Error parsing changes: {e}", file=sys.stderr)
        return 1

    # Determine output path
    output_path: Path = args.output or input_path.with_stem(input_path.stem + "_redlined")

    # Build config
    config = RedlineConfig(author=args.author)

    # Apply changes
    try:
        doc = apply_redlines(input_path, changes, config=config)
    except ValueError as e:
        print(f"Error applying changes: {e}", file=sys.stderr)
        return 1

    # Validate if requested
    if args.validate:
        result = validate_document(doc)
        if not result.ok:
            print("Validation errors:", file=sys.stderr)
            for err in result.errors:
                print(f"  ERROR: {err}", file=sys.stderr)
            return 1
        if result.warnings:
            for warn in result.warnings:
                print(f"  WARNING: {warn}", file=sys.stderr)

    # Save
    doc.save(output_path)
    print(f"Redlined document saved to: {output_path}")
    return 0


def _cmd_convert(args: argparse.Namespace) -> int:
    """Execute the ``convert`` subcommand."""
    input_path: Path = args.input

    if not input_path.exists():
        print(f"Error: input file not found: {input_path}", file=sys.stderr)
        return 1

    doc = DocxDocument(path=input_path)
    fragments = document_to_fragments(doc.paragraphs)

    if args.format == "json":
        output = [{"fragment_id": fid, "text": text} for fid, text in fragments]
        print(json.dumps(output, indent=2, ensure_ascii=False))
    else:
        # Tagged format
        print(fragments_to_tagged_text(fragments))

    return 0


def _cmd_validate(args: argparse.Namespace) -> int:
    """Execute the ``validate`` subcommand."""
    input_path: Path = args.input

    if not input_path.exists():
        print(f"Error: input file not found: {input_path}", file=sys.stderr)
        return 1

    doc = DocxDocument(path=input_path)
    result = validate_document(doc)

    if result.errors:
        print("ERRORS:")
        for err in result.errors:
            print(f"  {err}")

    if result.warnings:
        print("WARNINGS:")
        for warn in result.warnings:
            print(f"  {warn}")

    if result.ok:
        print("Validation passed." if not result.warnings else "Validation passed (with warnings).")
        return 0

    print("Validation FAILED.")
    return 1


def _parse_changes(raw: list[dict] | dict) -> list[Change]:
    """Parse a changes JSON structure into Change objects.

    Accepts either:
    - A list of change dicts
    - A dict with a "changes" key containing a list

    Each change dict must have:
    - fragment_id (int)
    - change_type (str: "modify", "delete", "append_after")
    - new_text (str, required for modify/append_after)
    - justification (str)
    """
    if isinstance(raw, dict):
        raw = raw["changes"]

    changes: list[Change] = []
    for item in raw:
        change = Change(
            fragment_id=item["fragment_id"],
            change_type=ChangeType(item["change_type"]),
            new_text=item.get("new_text"),
            justification=item.get("justification", ""),
        )
        changes.append(change)

    return changes
