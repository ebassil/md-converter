#!/usr/bin/env python3
"""Convert a .docx file, .html file, or the clipboard to pure Markdown.

HTML tables become Markdown tables, and every other HTML tag is mapped to its
Markdown equivalent (or to plain text / Unicode where Markdown has none), so
the output contains no HTML tags. With no input file given, reads the current
clipboard contents (preferring rich HTML, falling back to plain text) and
converts that instead.
"""

import argparse
import re
import subprocess
import sys
from pathlib import Path

import mammoth
from bs4 import BeautifulSoup, Comment
from markdownify import MarkdownConverter


def docx_to_html(path: Path) -> str:
    with open(path, "rb") as f:
        result = mammoth.convert_to_html(f)
    return result.value


SUPERSCRIPTS = str.maketrans("0123456789+-=()ni", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱ")
SUBSCRIPTS = str.maketrans("0123456789+-=()aeoxhklmnpst", "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₒₓₕₖₗₘₙₚₛₜ")

# Elements whose content must never reach the output.
DROP_TAGS = ["script", "style", "noscript", "template", "head", "title", "meta", "link"]
# Embedded media: Markdown can only link to them.
MEDIA_TAGS = ("video", "audio", "iframe", "embed", "object", "source")


def _script(text: str, table: dict, fallback_prefix: str) -> str:
    """Use Unicode super/subscripts when every character maps, else `^(x)` / `_(x)`."""
    if text.strip() and all(ord(c) in table or c == " " for c in text):
        return text.translate(table)
    return f"{fallback_prefix}({text})"


class PureMarkdownConverter(MarkdownConverter):
    """MarkdownConverter extended so no HTML tag survives into the output."""

    # --- inline formatting that Markdown has no syntax for: keep the text -------
    def convert_u(self, el, text, parent_tags):
        return text

    convert_ins = convert_mark = convert_small = convert_big = convert_span = convert_u
    convert_font = convert_bdi = convert_bdo = convert_wbr = convert_u
    convert_data = convert_time = convert_label = convert_u

    def convert_sub(self, el, text, parent_tags):
        return _script(text, SUBSCRIPTS, "_")

    def convert_sup(self, el, text, parent_tags):
        return _script(text, SUPERSCRIPTS, "^")

    def convert_cite(self, el, text, parent_tags):
        return self.convert_em(el, text, parent_tags)

    convert_var = convert_dfn = convert_cite

    def convert_kbd(self, el, text, parent_tags):
        return self.convert_code(el, text, parent_tags)

    convert_samp = convert_tt = convert_kbd

    def convert_abbr(self, el, text, parent_tags):
        title = (el.get("title") or "").strip()
        return f"{text} ({title})" if title and title != text.strip() else text

    convert_acronym = convert_abbr

    def convert_ruby(self, el, text, parent_tags):
        return text

    def convert_rt(self, el, text, parent_tags):
        return f"({text})" if text.strip() else ""

    def convert_rp(self, el, text, parent_tags):
        return ""

    def convert_br(self, el, text, parent_tags):
        if "_inline" in parent_tags or "td" in parent_tags or "th" in parent_tags:
            return " "
        return "\\\n"  # backslash hard break survives line-trimming

    # --- blocks ---------------------------------------------------------------
    def convert_details(self, el, text, parent_tags):
        return f"\n\n{text.strip()}\n\n"

    def convert_summary(self, el, text, parent_tags):
        return f"\n\n**{text.strip()}**\n\n" if text.strip() else ""

    def convert_figure(self, el, text, parent_tags):
        return f"\n\n{text.strip()}\n\n"

    def convert_figcaption(self, el, text, parent_tags):
        return f"\n\n*{text.strip()}*\n\n" if text.strip() else ""

    def convert_address(self, el, text, parent_tags):
        return self.convert_p(el, text, parent_tags)

    convert_section = convert_article = convert_header = convert_footer = MarkdownConverter.convert_div
    convert_main = convert_aside = convert_nav = convert_fieldset = convert_center = MarkdownConverter.convert_div
    convert_form = convert_legend = MarkdownConverter.convert_div
    convert_menu = convert_dir = MarkdownConverter.convert_list

    def convert_input(self, el, text, parent_tags):
        if (el.get("type") or "").lower() in ("checkbox", "radio"):
            return "[x]" if el.has_attr("checked") else "[ ]"
        return ""

    def convert_button(self, el, text, parent_tags):
        return text

    convert_select = convert_textarea = convert_option = convert_optgroup = convert_button

    def _media(self, el, text, parent_tags):
        src = el.get("src") or el.get("data") or ""
        if not src:
            source = el.find("source", src=True)
            src = source["src"] if source else ""
        if not src:
            return text
        label = (el.get("title") or el.get("alt") or text.strip() or src).strip()
        return f"[{label}]({src})"

    convert_video = convert_audio = convert_iframe = convert_embed = convert_object = _media

    def convert_source(self, el, text, parent_tags):
        return ""

    convert_track = convert_param = convert_map = convert_area = convert_source

    def convert_picture(self, el, text, parent_tags):
        return text

    def convert_svg(self, el, text, parent_tags):
        return ""

    convert_canvas = convert_svg

    # --- tables ---------------------------------------------------------------
    @staticmethod
    def _cell_text(text: str) -> str:
        text = re.sub(r"(?m)^\s*(?:[-*+]|\d+\.)\s+", "• ", text.strip())
        text = re.sub(r"\s*\n\s*", " ", text)
        return text.replace("|", "\\|")

    def convert_td(self, el, text, parent_tags):
        colspan = 1
        if str(el.get("colspan", "")).isdigit():
            colspan = max(1, min(1000, int(el["colspan"])))
        return " " + self._cell_text(text) + " |" * colspan

    convert_th = convert_td

    def convert_caption(self, el, text, parent_tags):
        return f"\n\n**{text.strip()}**\n\n" if text.strip() else ""

    def convert_colgroup(self, el, text, parent_tags):
        return ""

    convert_col = convert_colgroup


def _preprocess(html: str) -> BeautifulSoup:
    soup = BeautifulSoup(html, "html.parser")
    for node in soup.find_all(string=lambda s: isinstance(s, Comment)):
        node.extract()
    for el in soup(DROP_TAGS):
        el.decompose()
    for el in soup.find_all(attrs={"hidden": True}):
        el.decompose()
    # Prefer the <body> when a full document was supplied.
    return soup.body or soup


def _code_language(el) -> str:
    code = el.find("code") or el
    for cls in (code.get("class") or []) + (el.get("class") or []):
        m = re.match(r"(?:language|lang|highlight-source)-(.+)", cls)
        if m:
            return m.group(1)
    return ""


def html_to_markdown(html: str) -> str:
    md = PureMarkdownConverter(
        heading_style="ATX",
        bullets="-",
        code_language_callback=_code_language,
        newline_style="backslash",
    ).convert_soup(_preprocess(html))
    md = md.replace("\xa0", " ")
    # Trim trailing whitespace per line and collapse 3+ blank lines down to 1.
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
