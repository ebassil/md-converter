#!/usr/bin/env python3
"""Convert a .docx file, .html file, or the clipboard to pure Markdown.

Rendering HTML tables as Markdown tables. With no input file given, reads
the current clipboard contents (preferring rich HTML, falling back to
plain text) and converts that instead.
"""

import argparse
import re
import subprocess
import sys
from pathlib import Path

import mammoth
from markdownify import markdownify


def docx_to_html(path: Path) -> str:
    with open(path, "rb") as f:
        result = mammoth.convert_to_html(f)
    return result.value


def html_to_markdown(html: str) -> str:
    md = markdownify(
        html,
        heading_style="ATX",
        bullets="-",
        strip=["script", "style"],
    )
    # Collapse 3+ blank lines down to 1, and trim trailing whitespace per line.
    md = "\n".join(line.rstrip() for line in md.splitlines())
    md = re.sub(r"\n{3,}", "\n\n", md)
    return md.strip() + "\n"


def get_clipboard_html() -> str | None:
    """Return the clipboard's HTML flavor as a string, or None if unavailable."""
    if sys.platform != "darwin":
        return None
    script = 'the clipboard as «class HTML»'
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True, text=True, timeout=5,
        )
    except (OSError, subprocess.SubprocessError):
        return None
    if result.returncode != 0:
        return None
    match = re.search(r"«data HTML([0-9A-Fa-f]+)»", result.stdout)
    if not match:
        return None
    try:
        return bytes.fromhex(match.group(1)).decode("utf-8", errors="replace")
    except ValueError:
        return None


def set_clipboard_text(text: str) -> None:
    """Write text back to the system clipboard."""
    if sys.platform == "darwin":
        cmd = ["pbcopy"]
    elif sys.platform.startswith("linux"):
        cmd = ["xclip", "-selection", "clipboard"]
    elif sys.platform == "win32":
        cmd = ["clip"]
    else:
        raise RuntimeError(f"Clipboard access not supported on platform: {sys.platform}")

    result = subprocess.run(cmd, input=text, text=True, capture_output=True)
    if result.returncode != 0:
        raise RuntimeError(f"Failed to write to clipboard: {result.stderr.strip()}")


def get_clipboard_text() -> str:
    """Return the clipboard's plain-text contents."""
    if sys.platform == "darwin":
        cmd = ["pbpaste"]
    elif sys.platform.startswith("linux"):
        cmd = ["xclip", "-selection", "clipboard", "-o"]
    elif sys.platform == "win32":
        cmd = ["powershell", "-command", "Get-Clipboard"]
    else:
        raise RuntimeError(f"Clipboard access not supported on platform: {sys.platform}")

    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"Failed to read clipboard: {result.stderr.strip()}")
    return result.stdout


def convert_clipboard() -> str:
    html = get_clipboard_html()
    if html:
        return html_to_markdown(html)
    text = get_clipboard_text()
    if not text.strip():
        raise ValueError("Clipboard is empty")
    return text.strip() + "\n"


def convert_file(input_path: Path) -> str:
    suffix = input_path.suffix.lower()
    if suffix == ".docx":
        html = docx_to_html(input_path)
    elif suffix in (".html", ".htm"):
        html = input_path.read_text(encoding="utf-8")
    else:
        raise ValueError(f"Unsupported file type: {suffix} (expected .docx, .html, or .htm)")
    return html_to_markdown(html)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "input", type=Path, nargs="?", default=None,
        help="Path to a .docx or .html file. Omit to read from the clipboard.",
    )
    parser.add_argument(
        "-o", "--output", type=Path, default=None,
        help="Output .md path (default: same name as input, .md extension; "
             "written back to the clipboard when reading from the clipboard)",
    )
    parser.add_argument(
        "--stdout", action="store_true",
        help="When reading from the clipboard, print the Markdown to stdout instead of writing it back to the clipboard.",
    )
    args = parser.parse_args()

    try:
        if args.input is None:
            markdown = convert_clipboard()
        else:
            if not args.input.exists():
                print(f"Error: input file not found: {args.input}", file=sys.stderr)
                sys.exit(1)
            markdown = convert_file(args.input)
    except (ValueError, RuntimeError) as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    if args.output:
        args.output.write_text(markdown, encoding="utf-8")
        print(f"Wrote {args.output}")
    elif args.input is not None:
        output_path = args.input.with_suffix(".md")
        output_path.write_text(markdown, encoding="utf-8")
        print(f"Wrote {output_path}")
    elif args.stdout:
        sys.stdout.write(markdown)
    else:
        set_clipboard_text(markdown)
        print("Wrote Markdown back to the clipboard")


if __name__ == "__main__":
    main()
