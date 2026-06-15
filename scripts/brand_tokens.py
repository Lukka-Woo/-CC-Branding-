"""
Brand tokens loaded from tokens.json — single source of truth.
Import this module instead of hard-coding any brand value.

Usage:
    import scripts.brand_tokens as BT
    color = BT.PRIMARY_500_HEX          # "#3EC99E"
    rgb   = BT.PRIMARY_500_RGB          # (49, 158, 124)
    emu   = BT.logo_emu(width_mm=48)    # (1728000, 436320)
"""

import json, os, pathlib
from typing import Tuple

_TOKENS_PATH = pathlib.Path(__file__).parent.parent / "tokens.json"
_ASSETS_PATH = pathlib.Path(__file__).parent.parent / "assets"

with open(_TOKENS_PATH, encoding="utf-8") as _f:
    _T = json.load(_f)

# ── Company identity ──────────────────────────────────────────────────────────
# 修改公司全称时，请同步更新：
# - templates/html/base.html (2处)
# - templates/pdf/brand.css (1处)
# - templates/ppt/gen_brand_master.py (2处)
BRAND_NAME_CN  = "未来方舟"
BRAND_NAME_EN  = "ArktechX"
BRAND_FULL_CN  = "上海未来方舟智能科技有限公司"
BRAND_FULL_EN  = "Shanghai ArktechX Intelligent Technology Co., Ltd."
BRAND_TAGLINE  = _T["brand"]["tagline"]

# 地址信息（用于名片、证明文件等）
COMPANY_ADDRESS_CN = "上海市浦河创业中心（桂平路）302栋"
COMPANY_ADDRESS_EN = "Building 302, Puhe Entrepreneurship Center (Guiping Road), Shanghai"

# ── Colors — hex strings ──────────────────────────────────────────────────────
PRIMARY_500_HEX   = _T["colors"]["primary"]["500"]
PRIMARY_600_HEX   = _T["colors"]["primary"]["600"]
PRIMARY_100_HEX   = _T["colors"]["primary"]["100"]
SECONDARY_500_HEX = _T["colors"]["secondary"]["500"]
SECONDARY_100_HEX = _T["colors"]["secondary"]["100"]
NEUTRAL_900_HEX   = _T["colors"]["neutral"]["900"]
NEUTRAL_700_HEX   = _T["colors"]["neutral"]["700"]
NEUTRAL_400_HEX   = _T["colors"]["neutral"]["400"]
NEUTRAL_200_HEX   = _T["colors"]["neutral"]["200"]
NEUTRAL_100_HEX   = _T["colors"]["neutral"]["100"]
NEUTRAL_50_HEX    = _T["colors"]["neutral"]["50"]
WHITE_HEX         = _T["colors"]["neutral"]["0"]
SUCCESS_HEX       = _T["colors"]["semantic"]["success"]
WARNING_HEX       = _T["colors"]["semantic"]["warning"]
DANGER_HEX        = _T["colors"]["semantic"]["danger"]

# ── Background / Text / Border semantic tokens ────────────────────────────────
BG_PAGE_HEX        = _T["colors"]["background"]["page"]     # #F8FAFC  page-level
BG_NEUTRAL_HEX     = _T["colors"]["background"]["neutral"]  # #FAFAFA  weak section
BG_SURFACE_HEX     = _T["colors"]["background"]["surface"]  # #FFFFFF  card/container
BORDER_DEFAULT_HEX = _T["border"]["color"]["default"]       # #EAECF0
BORDER_SUBTLE_HEX  = _T["border"]["color"]["subtle"]        # #D0D5DD
TEXT_PRIMARY_HEX   = _T["colors"]["text"]["primary"]        # #101828
TEXT_SECONDARY_HEX = _T["colors"]["text"]["secondary"]      # #475467

# Full palette list (for compliance checking)
BRAND_PALETTE_HEX = {
    "primary-500":   PRIMARY_500_HEX,
    "primary-600":   PRIMARY_600_HEX,
    "primary-100":   PRIMARY_100_HEX,
    "secondary-500": SECONDARY_500_HEX,
    "secondary-100": SECONDARY_100_HEX,
    "neutral-900":   NEUTRAL_900_HEX,
    "neutral-700":   NEUTRAL_700_HEX,
    "neutral-400":   NEUTRAL_400_HEX,
    "neutral-200":   NEUTRAL_200_HEX,
    "neutral-100":   NEUTRAL_100_HEX,
    "neutral-50":    NEUTRAL_50_HEX,
    "white":         WHITE_HEX,
    "success":       SUCCESS_HEX,
    "warning":       WARNING_HEX,
    "danger":        DANGER_HEX,
    # Semantic
    "bg-page":       BG_PAGE_HEX,
    "bg-neutral":    BG_NEUTRAL_HEX,
    "bg-surface":    BG_SURFACE_HEX,
    "text-primary":  TEXT_PRIMARY_HEX,
    "text-secondary": TEXT_SECONDARY_HEX,
    "border-default": BORDER_DEFAULT_HEX,
}

# ── Colors — RGB tuples (r, g, b) ─────────────────────────────────────────────
def _hex_to_rgb(h: str) -> Tuple[int, int, int]:
    h = h.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)

PRIMARY_500_RGB   = _hex_to_rgb(PRIMARY_500_HEX)
PRIMARY_600_RGB   = _hex_to_rgb(PRIMARY_600_HEX)
PRIMARY_100_RGB   = _hex_to_rgb(PRIMARY_100_HEX)
SECONDARY_500_RGB = _hex_to_rgb(SECONDARY_500_HEX)
NEUTRAL_900_RGB   = _hex_to_rgb(NEUTRAL_900_HEX)
NEUTRAL_700_RGB   = _hex_to_rgb(NEUTRAL_700_HEX)
NEUTRAL_400_RGB   = _hex_to_rgb(NEUTRAL_400_HEX)
NEUTRAL_200_RGB   = _hex_to_rgb(NEUTRAL_200_HEX)
NEUTRAL_100_RGB   = _hex_to_rgb(NEUTRAL_100_HEX)
WHITE_RGB         = _hex_to_rgb(WHITE_HEX)

# ── python-docx RGBColor (lazy import so module usable without docx) ──────────
def docx_rgb(hex_color: str):
    from docx.shared import RGBColor
    r, g, b = _hex_to_rgb(hex_color)
    return RGBColor(r, g, b)

# ── python-pptx RGBColor ──────────────────────────────────────────────────────
def pptx_rgb(hex_color: str):
    from pptx.util import Pt  # noqa
    from pptx.dml.color import RGBColor
    r, g, b = _hex_to_rgb(hex_color)
    return RGBColor(r, g, b)

# ── Typography ────────────────────────────────────────────────────────────────
FONT_CN     = "PingFang SC"
FONT_CN_FB  = "Microsoft YaHei"
FONT_CN_WEB = "Noto Sans SC"
FONT_EN     = "Inter"
FONT_MONO   = "JetBrains Mono"

# Docx scale sizes (pt)
FONT_H1_PT    = _T["media"]["docx"]["heading1"]["size"]
FONT_H2_PT    = _T["media"]["docx"]["heading2"]["size"]
FONT_H3_PT    = _T["media"]["docx"]["heading3"]["size"]
FONT_BODY_PT  = 11
FONT_SMALL_PT = _T["typography"]["scale"]["body2"]["size"] // 1

FONT_H1_COLOR = _T["media"]["docx"]["heading1"]["color"]
FONT_H2_COLOR = _T["media"]["docx"]["heading2"]["color"]
FONT_H3_COLOR = _T["media"]["docx"]["heading3"]["color"]

# ── Spacing (mm) ──────────────────────────────────────────────────────────────
PAGE_MARGIN_TOP_MM    = _T["media"]["docx"]["margin"]["top"]
PAGE_MARGIN_BOTTOM_MM = _T["media"]["docx"]["margin"]["bottom"]
PAGE_MARGIN_LEFT_MM   = _T["media"]["docx"]["margin"]["left"]
PAGE_MARGIN_RIGHT_MM  = _T["media"]["docx"]["margin"]["right"]

SPACING_XXS_MM = _T["spacing"]["xxs"] * 0.265
SPACING_XS_MM  = _T["spacing"]["xs"]  * 0.265
SPACING_S_MM   = _T["spacing"]["s"]   * 0.265
SPACING_M_MM   = _T["spacing"]["m"]   * 0.265
SPACING_L_MM   = _T["spacing"]["l"]   * 0.265

# ── Table colors ──────────────────────────────────────────────────────────────
TABLE_HEADER_BG  = _T["media"]["docx"]["tableHeaderBg"]
TABLE_HEADER_FG  = _T["media"]["docx"]["tableHeaderColor"]
TABLE_BORDER     = _T["media"]["docx"]["tableBorderColor"]
TABLE_STRIPE_BG  = _T["media"]["docx"]["tableStripeBg"]

# ── Logo paths ────────────────────────────────────────────────────────────────
LOGO_HORIZONTAL_PRIMARY_PNG = str(_ASSETS_PATH / "logo-horizontal-primary.png")
LOGO_HORIZONTAL_PRIMARY_SVG = str(_ASSETS_PATH / "logo-horizontal-primary.svg")
LOGO_HORIZONTAL_REVERSE_PNG = None
LOGO_STACKED_PRIMARY_PNG    = str(_ASSETS_PATH / "logo-stacked-primary.png")
LOGO_STACKED_REVERSE_PNG    = str(_ASSETS_PATH / "logo-stacked-reverse.png")
LOGO_MARK_PNG               = str(_ASSETS_PATH / "logo-mark-primary.png")

# Decorative elements (abstract graphic assets)
DECO_3_PNG = str(_ASSETS_PATH / "装饰性元素" / "3.png")

# SVG viewBox natural ratio: 171.34 : 43.28
LOGO_HORIZONTAL_ASPECT = 171.34 / 43.28   # ≈ 3.96

def logo_emu(width_mm: float = 48.0) -> Tuple[int, int]:
    """Return (width_emu, height_emu) for the horizontal logo at given width."""
    w = int(width_mm * 36000)
    h = int(width_mm / LOGO_HORIZONTAL_ASPECT * 36000)
    return w, h


if __name__ == "__main__":
    print("=== Brand Tokens ===")
    print(f"Primary:   {PRIMARY_500_HEX}  rgb{PRIMARY_500_RGB}")
    print(f"Secondary: {SECONDARY_500_HEX}  rgb{SECONDARY_500_RGB}")
    print(f"Font CN:   {FONT_CN}")
    print(f"Font EN:   {FONT_EN}")
    print(f"Margins:   {PAGE_MARGIN_LEFT_MM}mm / {PAGE_MARGIN_TOP_MM}mm")
    print(f"Logo EMU:  {logo_emu()}")
