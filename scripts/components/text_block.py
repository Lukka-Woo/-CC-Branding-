"""TextBlock component — text content area with optional card background."""
from __future__ import annotations
from dataclasses import dataclass, field
from scripts.components.base import BaseComponent
import scripts.brand_tokens as BT


@dataclass
class TextBlock(BaseComponent):
    content:     str   = ""
    bullets:     list  = field(default_factory=list)
    sz:          int   = 14
    color:       str   = ""           # hex, default NEUTRAL_700
    arrow_style: str   = "primary"    # "primary" | "white"
    bg_color:    str   = ""           # if set, draw a rounded-rect card FIRST, then text inside
    pad_s_mm:    float = 5.0          # inner side padding (only when bg_color is set)
    pad_t_mm:    float = 6.0          # inner top/bottom padding (only when bg_color is set)

    def render_pptx(self, slide, x: int, y: int, w: int, h: int) -> None:
        from scripts.pptx_builder import _render_content_with_arrows, _txb, _card
        from pptx.util import Mm
        color = self.color or BT.NEUTRAL_700_HEX

        # Draw card background FIRST so text appears on top (correct z-order)
        if self.bg_color:
            _card(slide, l=x, t=y, w=w, h=h, bg=self.bg_color)
            ps   = int(Mm(self.pad_s_mm))
            pt   = int(Mm(self.pad_t_mm))
            ix, iy = x + ps, y + pt
            iw, ih = w - 2 * ps, h - 2 * pt
        else:
            ix, iy, iw, ih = x, y, w, h

        if self.content:
            has_arrows = "→" in self.content
            if has_arrows:
                _render_content_with_arrows(
                    slide, self.content, ix, iy, iw, ih,
                    sz=self.sz, color=color, ls_pt=round(self.sz * 1.5),
                    arrow_style=self.arrow_style,
                )
            else:
                _txb(slide, self.content, l=ix, t=iy, w=iw, h=ih,
                     sz=self.sz, color=color, ls_pt=round(self.sz * 1.5))
        elif self.bullets:
            line_h = Mm(self.sz * 0.353 * 1.8)
            y_off  = iy
            for bullet in self.bullets:
                if y_off + line_h > iy + ih:
                    break
                _txb(slide, f"• {bullet}", l=ix, t=int(y_off), w=iw, h=int(line_h),
                     sz=self.sz, color=color)
                y_off += line_h

    def render_html(self) -> str:
        color = self.color or BT.NEUTRAL_700_HEX
        if self.content:
            lines = self.content.split("\n")
            parts = []
            for line in lines:
                if line.strip().startswith("→"):
                    text = line.strip()[1:].lstrip()
                    parts.append(
                        f'<div class="arrow-line" style="display:flex;gap:6px;align-items:baseline">'
                        f'<span style="color:var(--color-primary-500);font-weight:700">→</span>'
                        f'<span style="font-weight:600">{text}</span></div>'
                    )
                elif line:
                    parts.append(f'<p style="margin:2px 0">{line}</p>')
                else:
                    parts.append('<br>')
            return (
                f'<div class="text-block" style="color:{color};font-size:{self.sz}px">'
                + "".join(parts) + "</div>"
            )
        if self.bullets:
            items = "".join(f'<li>{b}</li>' for b in self.bullets)
            return f'<ul class="text-block__bullets" style="color:{color}">{items}</ul>'
        return ""
