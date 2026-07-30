from .json_reporter import render_json
from .markdown_reporter import render_markdown
from .text_reporter import render_text

RENDERERS = {
    "json": render_json,
    "text": render_text,
    "markdown": render_markdown,
}

__all__ = ["RENDERERS", "render_json", "render_text", "render_markdown"]
