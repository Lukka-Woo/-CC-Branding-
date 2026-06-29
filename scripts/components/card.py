"""Card component — renders a single branded card for Grid/Stack layouts."""
from __future__ import annotations
from dataclasses import dataclass, field
from pptx.util import Mm, Pt
from scripts.components.base import BaseComponent, resolve_card_color
import scripts.brand_tokens as BT


@dataclass
class Card(BaseComponent):
    title: str = ""
    body:  str = ""
    tag:   str = ""
    color: str = "primary"   # semantic name → resolve_card_color()
    layout: str = "vertical_stack"
    # horizontal_split
    metric: str = ""
    # icon_left
    icon: str = ""
    # quote_card
    quote: str = ""
    attribution: str = ""
    # override
    dark: bool = False

    def _colors(self):
        if self.dark:
            return BT.NEUTRAL_900_HEX, BT.SECONDARY_500_HEX, BT.WHITE_HEX
        return resolve_card_color(self.color)

    def render_pptx(self, slide, x: int, y: int, w: int, h: int) -> None:
        from scripts.pptx_builder import _card, _render_card_inner, _txb, _rect
        from pptx.enum.text import PP_ALIGN
        from pptx.enum.dml import MSO_THEME_COLOR
        from pptx.dml.color import RGBColor
        try:
            from pptx.util import MSO_AUTO_SIZE
        except ImportError:
            from pptx.enum.text import MSO_AUTO_SIZE

        bg, acc, txt = self._colors()
        border = None if bg != BT.WHITE_HEX else BT.BORDER_DEFAULT_HEX
        _card(slide, l=x, t=y, w=w, h=h, bg=bg, border=border)

        if self.layout != "vertical_stack":
            data = {
                "title": self.title, "body": self.body, "tag": self.tag,
                "metric": self.metric, "icon": self.icon,
                "quote": self.quote, "attribution": self.attribution,
            }
            _render_card_inner(slide, self.layout, x, y, w, h, data, acc,
                               txt_color=txt, body_color=BT.NEUTRAL_700_HEX
                               if bg != BT.NEUTRAL_900_HEX else BT.NEUTRAL_100_HEX,
                               pad_s=Mm(5), pad_t=Mm(5))
            return

        # vertical_stack — adaptive heights to prevent overflow
        body_color = BT.NEUTRAL_700_HEX if bg != BT.NEUTRAL_900_HEX else BT.NEUTRAL_100_HEX
        PAD_S = Mm(5)
        PAD_T = Mm(5)
        PAD_B = Mm(6)
        inner_l = x + PAD_S
        inner_w = w - 2 * PAD_S
        available = h - PAD_T - PAD_B   # total vertical space for content

        # Reserve space proportionally
        has_tag  = bool(self.tag)
        tag_h    = Mm(6)  if has_tag  else 0
        tag_gap  = Mm(3)  if has_tag  else 0
        bar_h    = Mm(1)
        bar_gap  = Mm(3)
        # title gets 30% of remaining, body gets the rest
        after_tag    = available - tag_h - tag_gap
        title_budget = min(Mm(20), max(Mm(8), int(after_tag * 0.30)))
        bar_space    = bar_h + bar_gap if after_tag - title_budget > Mm(12) else 0
        body_h       = after_tag - title_budget - bar_space

        y_off = y + PAD_T

        if has_tag:
            _txb(slide, self.tag, l=inner_l, t=y_off, w=inner_w, h=tag_h,
                 sz=9, bold=True, color=acc)
            y_off += tag_h + tag_gap

        if self.title:
            tb_title = _txb(slide, self.title, l=inner_l, t=y_off, w=inner_w,
                            h=title_budget, sz=15, bold=True, color=txt)
            try:
                tb_title.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            except Exception:
                pass
            y_off += title_budget

        if bar_space:
            _rect(slide, l=inner_l, t=y_off, w=Mm(24), h=bar_h, fill=acc)
            y_off += bar_space

        if self.body and body_h > Mm(4):
            tb_body = _txb(slide, self.body, l=inner_l, t=y_off, w=inner_w,
                           h=body_h, sz=12, color=body_color, ls_pt=17)
            try:
                tb_body.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            except Exception:
                pass

    def render_html(self) -> str:
        bg, acc, txt = self._colors()
        cls = f"card card--{self.color}"
        if self.layout == "horizontal_split":
            return (
                f'<div class="{cls}" style="background:{bg}">'
                f'<div class="card__left">'
                f'<div class="card__tag" style="color:{acc}">{self.tag}</div>'
                f'<div class="card__title" style="color:{txt}">{self.title}</div>'
                f'<div class="card__body">{self.body}</div>'
                f'</div>'
                f'<div class="card__metric" style="color:{acc}">{self.metric}</div>'
                f'</div>'
            )
        if self.layout == "icon_left":
            return (
                f'<div class="{cls}" style="background:{bg}">'
                f'<div class="card__icon-badge" style="background:{acc}">{self.icon[:2].upper()}</div>'
                f'<div class="card__content">'
                f'<div class="card__tag" style="color:{acc}">{self.tag}</div>'
                f'<div class="card__title" style="color:{txt}">{self.title}</div>'
                f'<div class="card__body">{self.body}</div>'
                f'</div></div>'
            )
        if self.layout == "quote_card":
            return (
                f'<div class="{cls} card--quote" style="background:{bg}">'
                f'<div class="card__quote-mark" style="color:{acc}">“</div>'
                f'<div class="card__quote">{self.quote or self.body}</div>'
                f'<div class="card__attribution">— {self.attribution}</div>'
                f'</div>'
            )
        # vertical_stack (default)
        bar = f'<div class="card__bar" style="background:{acc}"></div>' if self.title else ""
        return (
            f'<div class="{cls}" style="background:{bg}">'
            f'<div class="card__tag" style="color:{acc}">{self.tag}</div>'
            f'<div class="card__title" style="color:{txt}">{self.title}</div>'
            f'{bar}'
            f'<div class="card__body">{self.body}</div>'
            f'</div>'
        )
