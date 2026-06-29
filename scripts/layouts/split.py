"""Split layout — two panes with a configurable left/right ratio."""
from __future__ import annotations
from dataclasses import dataclass
from scripts.components.base import BaseLayout


@dataclass
class Split(BaseLayout):
    left:   object = None
    right:  object = None
    ratio:  float  = 0.5
    gap_mm: float  = 6.0

    def render_pptx(self, slide, x: int, y: int, w: int, h: int) -> None:
        from pptx.util import Mm
        gap     = int(Mm(self.gap_mm))
        left_w  = int(w * self.ratio - gap / 2)
        right_w = w - left_w - gap
        right_x = x + left_w + gap
        if self.left:
            self.left.render_pptx(slide, x, y, left_w, h)
        if self.right:
            self.right.render_pptx(slide, right_x, y, right_w, h)

    def render_html(self) -> str:
        lp       = int(self.ratio * 100)
        rp       = 100 - lp
        gap_css  = f"{self.gap_mm / 10:.2f}rem"
        left_h   = self.left.render_html()  if self.left  else ""
        right_h  = self.right.render_html() if self.right else ""
        return (
            f'<div class="split" '
            f'style="display:grid;grid-template-columns:{lp}fr {rp}fr;gap:{gap_css}">'
            f'<div class="split__left">{left_h}</div>'
            f'<div class="split__right">{right_h}</div>'
            f'</div>'
        )
