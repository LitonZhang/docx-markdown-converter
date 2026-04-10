from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
import time

from .converter import DocxToMarkdownConverter, RecoverableConversionError


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Convert DOCX file to Markdown")
    parser.add_argument("--input", required=True, help="Input .docx file path")
    parser.add_argument("--output", required=True, help="Output .md file path")
    parser.add_argument("--math", default="latex", choices=["latex"], help="Math output mode")
    parser.add_argument("--extract-images", action="store_true", help="Enable image extraction")
    parser.add_argument("--image-dir", dest="image_dir", help="Directory for extracted images")
    parser.add_argument("--assets-dir", dest="assets_dir", help="Directory for extracted image assets")
    parser.add_argument("--report", help="Optional path to conversion report JSON")
    return parser


def _write_report(path: Path, report: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")


def _empty_stats() -> dict[str, int]:
    return {
        "headings": 0,
        "tables": 0,
        "images": 0,
        "equations": 0,
    }


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    started_at = time.perf_counter()
    report_path = Path(args.report) if args.report else None

    try:
        converter = DocxToMarkdownConverter(math=args.math, extract_images=args.extract_images)
        image_dir = args.image_dir if args.extract_images else None
        result = converter.convert_file(
            args.input,
            args.output,
            image_dir=image_dir,
            assets_dir=args.assets_dir,
        )
        duration_ms = int((time.perf_counter() - started_at) * 1000)
        result.duration_ms = duration_ms

        report = result.to_report()
        if report_path is not None:
            _write_report(report_path, report)
        print(json.dumps(report, ensure_ascii=False))
        return 0

    except RecoverableConversionError as exc:
        duration_ms = int((time.perf_counter() - started_at) * 1000)
        report = {
            "success": False,
            "warnings": [str(exc)],
            "stats": _empty_stats(),
            "duration_ms": duration_ms,
        }
        if report_path is not None:
            _write_report(report_path, report)
        print(str(exc), file=sys.stderr)
        return 1

    except OSError as exc:
        duration_ms = int((time.perf_counter() - started_at) * 1000)
        report = {
            "success": False,
            "warnings": [f"System error: {exc}"],
            "stats": _empty_stats(),
            "duration_ms": duration_ms,
        }
        if report_path is not None:
            _write_report(report_path, report)
        print(f"System error: {exc}", file=sys.stderr)
        return 2

    except Exception as exc:  # noqa: BLE001
        duration_ms = int((time.perf_counter() - started_at) * 1000)
        report = {
            "success": False,
            "warnings": [f"Unexpected error: {exc}"],
            "stats": _empty_stats(),
            "duration_ms": duration_ms,
        }
        if report_path is not None:
            _write_report(report_path, report)
        print(f"Unexpected error: {exc}", file=sys.stderr)
        return 2


def main_entry() -> int:
    return main()


if __name__ == "__main__":
    raise SystemExit(main())
