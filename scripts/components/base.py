"""Base classes and color resolver for component-first layout system."""
from __future__ import annotations
from abc import ABC, abstractmethod
import scripts.brand_tokens as BT

CARD_COLOR_MAP: dict[str, tuple[str, str, str]] = {
    "primary":   (BT.PRIMARY_100_HEX,   BT.PRIMARY_500_HEX,   BT.NEUTRAL_900_HEX),
    "success":   (BT.NEUTRAL_100_HEX,   BT.SUCCESS_HEX,        BT.NEUTRAL_900_HEX),
    "secondary": (BT.SECONDARY_100_HEX, BT.SECONDARY_500_HEX,  BT.NEUTRAL_900_HEX),
    "warning":   (BT.CARD_ORANGE_BG,    BT.WARNING_HEX,        BT.NEUTRAL_900_HEX),
    "teal":      (BT.CARD_TEAL_BG,      BT.TEAL_HEX,           BT.NEUTRAL_900_HEX),
    "purple":    (BT.CARD_PURPLE_BG,    BT.PURPLE_HEX,         BT.NEUTRAL_900_HEX),
    "dark":      (BT.NEUTRAL_900_HEX,   BT.SECONDARY_500_HEX,  BT.WHITE_HEX),
    "danger":    (BT.CARD_DANGER_BG,    BT.DANGER_HEX,         BT.NEUTRAL_900_HEX),
    "neutral":   (BT.NEUTRAL_100_HEX,   BT.NEUTRAL_400_HEX,    BT.NEUTRAL_900_HEX),
}


def resolve_card_color(name: str) -> tuple[str, str, str]:
    """Return (bg_hex, accent_hex, txt_hex) for a semantic color name."""
    return CARD_COLOR_MAP.get(name or "primary", CARD_COLOR_MAP["primary"])


class BaseComponent(ABC):
    @abstractmethod
    def render_pptx(self, slide, x: int, y: int, w: int, h: int) -> None: ...

    @abstractmethod
    def render_html(self) -> str: ...


class BaseLayout(ABC):
    @abstractmethod
    def render_pptx(self, slide, x: int, y: int, w: int, h: int) -> None: ...

    @abstractmethod
    def render_html(self) -> str: ...
