"""Stat component — a single KPI stat card."""
from __future__ import annotations
from dataclasses import dataclass
from scripts.components.base import BaseComponent, resolve_card_color
import scripts.brand_tokens as BT


@dataclass
class Stat(BaseComponent):
    val:    str    = ""      # numeric/symbol — never translated
    label:  object = ""     # str | {"cn": ..., "en": ...}
    desc:   object = ""     # str | {"cn": ..., "en": ...}
    color:  str    = "primary"
    layout: str    = "horizontal_split"   # "horizontal_split" | "vertical_stack"

    def render_pptx(self, slide, x: int, y: int, w: int, h: int) -> None:
        from scripts.pptx_builder import _card, _render_card_inner
        from scripts.i18n import T
        from pptx.util import Mm

        bg, acc, txt = resolve_card_color(self.color)
        _card(slide, l=x, t=y, w=w, h=h, bg=bg)

        data = {
            "metric":    self.val,
            "metric_sz": 32,
            "title":     T(self.label),
            "title_sz":  12,
            "body":      T(self.desc),
            "body_sz":   11,
        }
        _render_card_inner(slide, "horizontal_split", x, y, w, h, data, acc,
                           txt_color=txt,
                           body_color=BT.NEUTRAL_700_HEX if bg != BT.NEUTRAL_900_HEX
                           else BT.NEUTRAL_100_HEX,
                           pad_s=Mm(5), pad_t=Mm(5))

    def render_html(self) -> str:
        from scripts.i18n import T
        bg, acc, txt = resolve_card_color(self.color)
        return (
            f'<div class="stat-card" style="background:{bg};border-radius:8px;padding:16px">'
            f'<div class="stat-card__metric" style="color:{acc};font-size:36px;font-weight:700">{self.val}</div>'
            f'<div class="stat-card__label" style="color:{txt};font-weight:600;font-size:13px">{T(self.label)}</div>'
            f'<div class="stat-card__desc" style="color:#3D444A;font-size:11px;margin-top:4px">{T(self.desc)}</div>'
            f'</div>'
        )
