#!/usr/bin/env python3
"""
Sovereign Doc – CLI
"""

import argparse
import sys
from pathlib import Path

from sovereign_doc_core import convert_any, SovereignDocError


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description="Sovereign Doc – local round-trip converter")
    parser.add_argument(
        "input",
        type=str,
        help="Path to a file (.docx/.txt/.md) OR a folder (batch on .docx only)."
    )
    parser.add_argument(
        "--to",
        dest="dst_format",
        choices=["txt", "md", "html", "docx"],
        required=True,
        help="Destination format."
    )
    parser.add_argument(
        "--no-log",
        action="store_true",
        help="Disable logging this run to sovereign_doc_log.csv"
    )
    return parser.parse_args(argv)


def _convert_single(path: Path, dst_format: str, logging_enabled: bool) -> int:
    try:
        out_path = convert_any(path, dst_format, mode="cli-single", enable_log=logging_enabled)
    except SovereignDocError as e:
        print(f"[SOVEREIGN DOC ERROR] {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"[SOVEREIGN DOC FATAL] {e}", file=sys.stderr)
        return 1
    print(f"Sovereign Doc: {path.name} → {out_path.name}")
    print(f"Output saved to: {out_path}")
    return 0


def _convert_folder(folder: Path, dst_format: str, logging_enabled: bool) -> int:
    docx_files = sorted(folder.glob("*.docx"))
    if not docx_files:
        print(f"[SOVEREIGN DOC] No .docx files found in folder: {folder}")
        return 0
    print(f"[SOVEREIGN DOC] Converting {len(docx_files)} .docx file(s) in folder: {folder} → {dst_format}")
    success_count = 0
    for f in docx_files:
        try:
            out_path = convert_any(f, dst_format, mode="cli-batch", enable_log=logging_enabled)
        except SovereignDocError as e:
            print(f"[SOVEREIGN DOC ERROR] {f.name}: {e}", file=sys.stderr)
            continue
        except Exception as e:
            print(f"[SOVEREIGN DOC FATAL] {f.name}: {e}", file=sys.stderr)
            continue
        else:
            success_count += 1
            print(f"  OK: {f.name} → {out_path.name}")
    print(f"[SOVEREIGN DOC] Completed. Successfully converted {success_count} of {len(docx_files)} file(s).")
    return 0


def main(argv=None) -> int:
    args = parse_args(argv)
    input_path = Path(args.input).expanduser().resolve()
    logging_enabled = not args.no_log
    if input_path.is_dir():
        return _convert_folder(input_path, args.dst_format, logging_enabled)
    else:
        return _convert_single(input_path, args.dst_format, logging_enabled)


if __name__ == "__main__":
    raise SystemExit(main())
