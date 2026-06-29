"""Stack layout — children stacked vertically with equal height distribution."""
from __future__ import annotations
from dataclasses import dataclass, field
from scripts.components.base import BaseLayout


@dataclass
class Stack(BaseLayout):
    children: list  = field(default_factory=list)
    gap_mm:   float = 4.0

    def render_pptx(self, slide, x: int, y: int, w: int, h: int) -> None:
        from pptx.util import Mm
        n = len(self.children)
        if not n:
            return
        gap    = int(Mm(self.gap_mm))
        cell_h = (h - gap * (n - 1)) // n
        for i, child in enumerate(self.children):
            cy = y + i * (cell_h + gap)
            child.render_pptx(slide, x, int(cy), w, int(cell_h))

    def render_html(self) -> str:
        gap_css = f"{self.gap_mm / 10:.2f}rem"
        items   = "".join(
            f'<div class="stack__item">{c.render_html()}</div>'
            for c in self.children
        )
        return (
            f'<div class="stack" '
            f'style="display:flex;flex-direction:column;gap:{gap_css}">'
            f'{items}</div>'
        )
