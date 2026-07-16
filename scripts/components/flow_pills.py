"""FlowPills component — pill → arrow → pill chain."""
from __future__ import annotations
from dataclasses import dataclass, field
from scripts.components.base import BaseComponent
import scripts.brand_tokens as BT


@dataclass
class FlowPills(BaseComponent):
    items:   list  = field(default_factory=list)  # list of str | {"cn":..,"en":..}
    style:   str   = "primary"
    font_sz: int   = 11

    def render_pptx(self, slide, x: int, y: int, w: int, h: int) -> None:
        from scripts.pptx_builder import _flow_pills, _PILL_STYLES
        from scripts.i18n import T
        bg, fg = _PILL_STYLES.get(self.style, _PILL_STYLES["primary"])
        resolved = [T(item) for item in self.items]
        _flow_pills(slide, resolved, l=x, t=y, w=w, h=h,
                    pill_bg=bg, pill_fg=fg, sz=self.font_sz)

    def render_html(self) -> str:
        from scripts.pptx_builder import _PILL_STYLES
        from scripts.i18n import T
        bg, fg = _PILL_STYLES.get(self.style, _PILL_STYLES["primary"])
        parts = []
        for i, item in enumerate(self.items):
            label = T(item)
            parts.append(
                f'<span style="background:{bg};color:{fg};border-radius:99px;'
                f'padding:3px 10px;font-size:{self.font_sz}px;font-weight:600">{label}</span>'
            )
            if i < len(self.items) - 1:
                parts.append(
                    f'<span style="color:{BT.NEUTRAL_400_HEX};font-size:{self.font_sz}px;padding:0 4px">→</span>'
                )
        return (
            f'<div class="flow-pills" style="display:flex;align-items:center;'
            f'gap:4px;flex-wrap:wrap">{"".join(parts)}</div>'
        )
