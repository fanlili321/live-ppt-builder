#!/usr/bin/env python3
"""Direct Keynote automation for live-ppt-builder."""

from __future__ import annotations

import argparse
import shutil
import subprocess
from pathlib import Path

DEFAULT_TEMPLATE = Path(__file__).resolve().parent.parent / "assets" / "指数模版.key"


def q(value: Path | str) -> str:
    return str(value).replace("\\", "\\\\").replace('"', '\\"')


def run_osascript(script: str) -> None:
    subprocess.run(["osascript", "-e", script], check=True)


def create_deck(template: Path, out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template, out_path)


def open_deck(path: Path) -> None:
    run_osascript(f'''
    tell application "Keynote"
        activate
        open POSIX file "{q(path)}"
    end tell
    ''')


def export_pdf(path: Path, out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    run_osascript(f'''
    tell application "Keynote"
        activate
        set theDoc to open POSIX file "{q(path)}"
        export theDoc to POSIX file "{q(out_path)}" as PDF
        save theDoc
    end tell
    ''')


def duplicate_slide(path: Path, source: int) -> None:
    run_osascript(f'''
    tell application "Keynote"
        activate
        set theDoc to open POSIX file "{q(path)}"
        tell theDoc
            duplicate slide {source} to after last slide
            save
        end tell
    end tell
    ''')


def add_text(path: Path, slide: int, text: str, x: float, y: float, w: float, h: float) -> None:
    run_osascript(f'''
    tell application "Keynote"
        activate
        set theDoc to open POSIX file "{q(path)}"
        tell slide {slide} of theDoc
            make new text item with properties {{object text:"{q(text)}", position:{{{x}, {y}}}, width:{w}, height:{h}}}
        end tell
        save theDoc
    end tell
    ''')


def add_image(path: Path, slide: int, image: Path, x: float, y: float, w: float, h: float) -> None:
    run_osascript(f'''
    tell application "Keynote"
        activate
        set theDoc to open POSIX file "{q(path)}"
        tell slide {slide} of theDoc
            set newImage to make new image with properties {{file:POSIX file "{q(image)}", position:{{{x}, {y}}}}}
            set width of newImage to {w}
            set height of newImage to {h}
        end tell
        save theDoc
    end tell
    ''')


def main() -> int:
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="cmd", required=True)

    create = subparsers.add_parser("create")
    create.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    create.add_argument("--out", type=Path, required=True)

    open_cmd = subparsers.add_parser("open")
    open_cmd.add_argument("deck", type=Path)

    export = subparsers.add_parser("export-pdf")
    export.add_argument("deck", type=Path)
    export.add_argument("--out", type=Path, required=True)

    duplicate = subparsers.add_parser("duplicate-slide")
    duplicate.add_argument("deck", type=Path)
    duplicate.add_argument("--source", type=int, required=True)

    text = subparsers.add_parser("add-text")
    text.add_argument("deck", type=Path)
    text.add_argument("--slide", type=int, required=True)
    text.add_argument("--text", required=True)
    text.add_argument("--x", type=float, required=True)
    text.add_argument("--y", type=float, required=True)
    text.add_argument("--w", type=float, required=True)
    text.add_argument("--h", type=float, required=True)

    image = subparsers.add_parser("add-image")
    image.add_argument("deck", type=Path)
    image.add_argument("--slide", type=int, required=True)
    image.add_argument("--image", type=Path, required=True)
    image.add_argument("--x", type=float, required=True)
    image.add_argument("--y", type=float, required=True)
    image.add_argument("--w", type=float, required=True)
    image.add_argument("--h", type=float, required=True)

    args = parser.parse_args()
    if args.cmd == "create":
        create_deck(args.template.expanduser().resolve(), args.out.expanduser().resolve())
    elif args.cmd == "open":
        open_deck(args.deck.expanduser().resolve())
    elif args.cmd == "export-pdf":
        export_pdf(args.deck.expanduser().resolve(), args.out.expanduser().resolve())
    elif args.cmd == "duplicate-slide":
        duplicate_slide(args.deck.expanduser().resolve(), args.source)
    elif args.cmd == "add-text":
        add_text(args.deck.expanduser().resolve(), args.slide, args.text, args.x, args.y, args.w, args.h)
    elif args.cmd == "add-image":
        add_image(args.deck.expanduser().resolve(), args.slide, args.image.expanduser().resolve(), args.x, args.y, args.w, args.h)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
