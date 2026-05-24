"""Inline markdown formatting for PowerPoint text runs.

Supports: **bold**, *italic*, ***bold italic***, ~~strikethrough~~,
__underline__, `code` (Courier New font). Handles nested formatting
and backslash escapes.

The API mirrors the approach used in docx_tools/inline_formatting.py but
targets python-pptx paragraph/run objects instead of python-docx ones.
"""

import re
import logging
from typing import Optional

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Inline formatting regex (subset of docx_tools/patterns.py)
# Covers the formats that render well in PowerPoint.
# ---------------------------------------------------------------------------

_INLINE_FORMAT_RE = re.compile(
    r'(\*{3}(?:[^*]|\*(?!\*{2}))+\*{3}'  # ***bold italic***
    r'|\*\*(?:[^*]|\*(?!\*))+\*\*'       # **bold**
    r'|~~.+?~~'                           # ~~strikethrough~~
    r'|__(?!_).+?__'                      # __underline__
    r'|\*(?:[^*]|\*\*[^*]+\*\*)+\*'       # *italic* (allows nested **bold**)
    r'|`[^`]+`)'                          # `code`
)

_ESCAPE_RE = re.compile(r'\\(.)')  # backslash-escaped character


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def has_inline_formatting(text: str) -> bool:
    """Quick check whether text contains any inline markdown markers."""
    return bool(_INLINE_FORMAT_RE.search(text))


def apply_inline_formatting(
    text_frame_or_paragraph,
    text: str,
    font_size: Optional[int] = None,
    bold: bool = False,
    italic: bool = False,
    alignment=None,
) -> None:
    """Parse inline markdown and render into a pptx text frame paragraph.

    If *text_frame_or_paragraph* is a TextFrame, uses its first paragraph.
    Otherwise treats it as a paragraph directly.

    Args:
        text_frame_or_paragraph: pptx TextFrame or paragraph object.
        text: Text potentially containing inline markdown.
        font_size: Optional font size to apply to all runs.
        bold: Inherited bold context.
        italic: Inherited italic context.
        alignment: Optional PP_ALIGN value for the paragraph.
    """
    # Determine if we got a text frame or a paragraph
    if hasattr(text_frame_or_paragraph, 'paragraphs'):
        paragraph = text_frame_or_paragraph.paragraphs[0]
    else:
        paragraph = text_frame_or_paragraph

    if alignment is not None:
        paragraph.alignment = alignment

    # Handle escape sequences
    escape_ctx = {"map": {}, "counter": 0}
    text = _handle_escapes(text, escape_ctx)

    # Parse and render
    _parse_segment(text, paragraph, font_size=font_size, bold=bold, italic=italic, escape_ctx=escape_ctx)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _handle_escapes(text: str, escape_ctx: dict) -> str:
    """Replace backslash-escaped characters with PUA placeholders."""
    def _replace(match):
        placeholder = chr(0xE000 + escape_ctx["counter"])
        escape_ctx["map"][placeholder] = match.group(1)
        escape_ctx["counter"] += 1
        return placeholder
    return _ESCAPE_RE.sub(_replace, text)


def _restore_escapes(text: str, escape_ctx: dict) -> str:
    """Replace PUA placeholders back with their original literal characters."""
    if not escape_ctx or not escape_ctx.get("map"):
        return text
    for placeholder, char in escape_ctx["map"].items():
        text = text.replace(placeholder, char)
    return text


def _add_run(paragraph, text: str, font_size=None, bold=False, italic=False,
             strike=False, underline=False, font_name=None):
    """Add a formatted run to a paragraph."""
    run = paragraph.add_run()
    run.text = text
    if font_size:
        run.font.size = font_size
    if bold:
        run.font.bold = True
    if italic:
        run.font.italic = True
    if strike:
        # python-pptx Font doesn't expose strikethrough — set via XML attribute
        rPr = run._r.get_or_add_rPr()
        rPr.set('strike', 'sngStrike')
    if underline:
        run.font.underline = True
    if font_name:
        run.font.name = font_name
    return run


def _parse_segment(text: str, paragraph, font_size=None, bold=False, italic=False, escape_ctx=None):
    """Parse a text segment for inline markdown and create runs."""
    for part in _INLINE_FORMAT_RE.split(text):
        if not part:
            continue

        if part.startswith('***') and part.endswith('***') and len(part) > 6:
            # Bold italic
            _parse_segment(part[3:-3], paragraph, font_size=font_size,
                           bold=True, italic=True, escape_ctx=escape_ctx)

        elif part.startswith('**') and part.endswith('**') and len(part) > 4:
            # Bold
            _parse_segment(part[2:-2], paragraph, font_size=font_size,
                           bold=True, italic=italic, escape_ctx=escape_ctx)

        elif part.startswith('~~') and part.endswith('~~') and len(part) > 4:
            # Strikethrough
            _add_run(paragraph, _restore_escapes(part[2:-2], escape_ctx),
                     font_size=font_size, bold=bold, italic=italic, strike=True)

        elif part.startswith('__') and part.endswith('__') and len(part) > 4 and not part.startswith('___'):
            # Underline
            _add_run(paragraph, _restore_escapes(part[2:-2], escape_ctx),
                     font_size=font_size, bold=bold, italic=italic, underline=True)

        elif part.startswith('*') and part.endswith('*') and not part.startswith('**') and len(part) > 2:
            # Italic
            _parse_segment(part[1:-1], paragraph, font_size=font_size,
                           bold=bold, italic=True, escape_ctx=escape_ctx)

        elif part.startswith('`') and part.endswith('`') and len(part) > 2:
            # Code (monospace)
            _add_run(paragraph, _restore_escapes(part[1:-1], escape_ctx),
                     font_size=font_size, bold=bold, italic=italic, font_name='Courier New')

        else:
            # Plain text
            _add_run(paragraph, _restore_escapes(part, escape_ctx),
                     font_size=font_size, bold=bold, italic=italic)




