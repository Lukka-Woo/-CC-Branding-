"""Grid layout — N-column equal-width grid."""
from __future__ import annotations
from dataclasses import dataclass, field
from scripts.components.base import BaseLayout


@dataclass
class Grid(BaseLayout):
    cols:      int   = 3
    children:  list  = field(default_factory=list)
    gap_mm:    float = 5.0
    row_gap_mm: float | None = None

    def render_pptx(self, slide, x: int, y: int, w: int, h: int) -> None:
        from pptx.util import Mm
        if not self.children:
            return
        n       = len(self.children)
        cols    = self.cols
        rows    = (n + cols - 1) // cols
        gap     = int(Mm(self.gap_mm))
        row_gap = int(Mm(self.row_gap_mm if self.row_gap_mm is not None else self.gap_mm))
        cell_w  = (w - gap * (cols - 1)) // cols
        cell_h  = (h - row_gap * (rows - 1)) // rows

        for i, child in enumerate(self.children):
            col = i % cols
            row = i // cols
            cx  = x + col * (cell_w + gap)
            cy  = y + row * (cell_h + row_gap)
            child.render_pptx(slide, int(cx), int(cy), int(cell_w), int(cell_h))

    def render_html(self) -> str:
        gap_css = f"{self.gap_mm / 10:.2f}rem"
        items   = "".join(
            f'<div class="grid__cell">{c.render_html()}</div>'
            for c in self.children
        )
        return (
            f'<div class="grid" style="display:grid;'
            f'grid-template-columns:repeat({self.cols},1fr);gap:{gap_css}">'
            f'{items}</div>'
        )
