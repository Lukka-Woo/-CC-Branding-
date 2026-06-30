"""TextBlock component — text content area with optional card background."""
from __future__ import annotations
from dataclasses import dataclass, field
from scripts.components.base import BaseComponent
import scripts.brand_tokens as BT


@dataclass
class TextBlock(BaseComponent):
    content:     object = ""          # str | {"cn": ..., "en": ...}
    bullets:     object = None        # list of str | list of {"cn":..,"en":..}
    sz:          int    = 14
    color:       str    = ""          # hex override for single-lang mode
    arrow_style: str    = "primary"
    bg_color:    str    = ""          # draw card background first if set
    pad_s_mm:    float  = 5.0
    pad_t_mm:    float  = 6.0

    def render_pptx(self, slide, x: int, y: int, w: int, h: int) -> None:
        from scripts.pptx_builder import _render_content_with_arrows, _txb, _card
        from scripts.i18n import T, is_bilingual, get_cn, get_en, \
            CN_COLOR, EN_COLOR, EN_SZ_RATIO
        from pptx.util import Mm

        bilingual = is_bilingual()

        # Draw card background FIRST (correct z-order: bg → text)
        if self.bg_color:
            _card(slide, l=x, t=y, w=w, h=h, bg=self.bg_color)
            ps = int(Mm(self.pad_s_mm))
            pt = int(Mm(self.pad_t_mm))
            ix, iy = x + ps, y + pt
            iw, ih = w - 2 * ps, h - 2 * pt
        else:
            ix, iy, iw, ih = x, y, w, h

        def _render_text(text, ox, oy, ow, oh, sz, color, en_mode=False):
            ls = round(sz * 1.5)
            kw = dict(sz=sz, color=color, ls_pt=ls)
            if en_mode:
                kw["en_font"] = "Inter"
                kw["cn_font"] = "Inter"
            if "→" in text:
                _render_content_with_arrows(
                    slide, text, ox, oy, ow, oh,
                    sz=sz, color=color, ls_pt=ls,
                    arrow_style="white" if en_mode else self.arrow_style,
                    **({k: v for k, v in kw.items()
                        if k not in ("sz", "color", "ls_pt")})
                )
            else:
                _txb(slide, text, l=ox, t=oy, w=ow, h=oh, **kw)

        if self.content is not None and self.content != "":
            cn_text = get_cn(self.content) if bilingual else T(self.content)
            en_text = get_en(self.content) if bilingual else ""

            if bilingual and en_text:
                cn_h = int(ih * 0.58)
                en_h = ih - cn_h - int(Mm(3))
                cn_color = self.color or CN_COLOR
                _render_text(cn_text, ix, iy, iw, cn_h, self.sz, cn_color)
                if en_h > int(Mm(4)):
                    en_sz = max(10, int(self.sz * EN_SZ_RATIO))
                    _render_text(en_text, ix, iy + cn_h + int(Mm(3)), iw, en_h,
                                 en_sz, EN_COLOR, en_mode=True)
            else:
                color = self.color or (BT.NEUTRAL_700_HEX)
                _render_text(cn_text or T(self.content), ix, iy, iw, ih,
                             self.sz, color)

        elif self.bullets:
            bullets = self.bullets or []
            line_h = Mm(self.sz * 0.353 * 1.8)
            y_off  = iy
            for bullet in bullets:
                if y_off + line_h > iy + ih:
                    break
                cn_b = get_cn(bullet) if bilingual else T(bullet)
                en_b = get_en(bullet) if bilingual else ""
                text = cn_b + (f"\n{en_b}" if en_b else "")
                color = self.color or BT.NEUTRAL_700_HEX
                _txb(slide, text, l=ix, t=int(y_off), w=iw, h=int(line_h),
                     sz=self.sz, color=color)
                y_off += line_h

    def render_html(self) -> str:
        from scripts.i18n import T, is_bilingual, get_cn, get_en, EN_COLOR, EN_SZ_RATIO
        color = self.color or BT.NEUTRAL_700_HEX
        bilingual = is_bilingual()

        bg_style = ""
        if self.bg_color:
            bg_style = (f"background:{self.bg_color};border-radius:8px;"
                        f"padding:{self.pad_t_mm/10:.1f}rem {self.pad_s_mm/10:.1f}rem;")

        def _html_content(text, sz, clr, en_mode=False):
            font = "Inter,sans-serif" if en_mode else "inherit"
            lines = text.split("\n")
            parts = []
            for line in lines:
                if line.strip().startswith("→"):
                    t = line.strip()[1:].lstrip()
                    parts.append(
                        f'<div style="display:flex;gap:6px;align-items:baseline">'
                        f'<span style="color:var(--color-primary-500);font-weight:700">→</span>'
                        f'<span style="font-weight:600">{t}</span></div>'
                    )
                elif line:
                    parts.append(f'<p style="margin:2px 0">{line}</p>')
                else:
                    parts.append('<br>')
            return (
                f'<div style="color:{clr};font-size:{sz}px;font-family:{font}">'
                + "".join(parts) + "</div>"
            )

        cn_text = get_cn(self.content) if bilingual else T(self.content)
        en_text = get_en(self.content) if bilingual else ""
        out = _html_content(cn_text, self.sz, color)
        if bilingual and en_text:
            en_sz = max(10, int(self.sz * EN_SZ_RATIO))
            out += _html_content(en_text, en_sz, EN_COLOR, en_mode=True)

        return f'<div class="text-block" style="{bg_style}">{out}</div>'
