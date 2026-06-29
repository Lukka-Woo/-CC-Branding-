"""Callout / Note block component."""
from __future__ import annotations
from dataclasses import dataclass
from scripts.components.base import BaseComponent


@dataclass
class Callout(BaseComponent):
    text:    str   = ""
    label:   str   = ""
    style:   str   = "note"   # note | info | tip | warning | danger
    font_sz: int   = 11

    def render_pptx(self, slide, x: int, y: int, w: int, h: int) -> None:
        from scripts.pptx_builder import _callout
        _callout(slide, self.text, l=x, t=y, w=w,
                 style=self.style, label=self.label or None, font_sz=self.font_sz)

    def render_html(self) -> str:
        style_map = {
            "note":    ("#E8F9F3", "#3EC99E", "注"),
            "info":    ("#E8F9F3", "#3EC99E", "ℹ"),
            "tip":     ("#F8FBE8", "#5CC13C", "TIP"),
            "warning": ("#FFF1DF", "#FFB928", "⚠"),
            "danger":  ("#FFF2F2", "#F12D2D", "!"),
        }
        bg, pill_color, default_label = style_map.get(self.style, style_map["note"])
        label = self.label or default_label
        return (
            f'<div class="callout callout--{self.style}" '
            f'style="background:{bg};border-radius:6px;padding:8px 12px;'
            f'display:flex;align-items:flex-start;gap:8px">'
            f'<span class="callout__pill" '
            f'style="background:{pill_color};color:#fff;border-radius:99px;'
            f'padding:2px 8px;font-size:11px;font-weight:700;white-space:nowrap">'
            f'{label}</span>'
            f'<span class="callout__text" style="font-size:12px">{self.text}</span>'
            f'</div>'
        )
