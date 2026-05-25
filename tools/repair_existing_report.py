from __future__ import annotations

import argparse
from pathlib import Path

from app.services.docx_renderer import DocxRenderer


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Repair an existing generated report layout.")
    parser.add_argument("input_path", type=Path, help="Path to the existing .docx report")
    parser.add_argument(
        "--output",
        dest="output_path",
        type=Path,
        default=None,
        help="Optional output path. Defaults to <input>_repaired.docx",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    renderer = DocxRenderer()
    output_path = renderer.repair_existing_report(
        input_path=args.input_path,
        output_path=args.output_path,
    )
    print(output_path)


if __name__ == "__main__":
    main()
