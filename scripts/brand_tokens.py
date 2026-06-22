"""
Brand tokens loaded from tokens.json — single source of truth.
Import this module instead of hard-coding any brand value.

Usage:
    import scripts.brand_tokens as BT
    color = BT.PRIMARY_500_HEX          # "#3EC99E"
    rgb   = BT.PRIMARY_500_RGB          # (49, 158, 124)
    emu   = BT.logo_emu(width_mm=48)    # (1728000, 436320)

Visual behavior preferences are read from brand_config.json (sibling of tokens.json).
Builders reference BT.DARK_ACCENT_CARDS_ENABLED / BT.MIN_CARDS_FOR_DARK / BT.MAX_DARK_PER_SLIDE
instead of hard-coding policy — the UI layer only needs to write brand_config.json.
"""

import json, os, pathlib
from typing import Tuple

_TOKENS_PATH = pathlib.Path(__file__).parent.parent / "tokens.json"
_CONFIG_PATH = pathlib.Path(__file__).parent.parent / "brand_config.json"
_ASSETS_PATH = pathlib.Path(__file__).parent.parent / "assets"

with open(_TOKENS_PATH, encoding="utf-8") as _f:
    _T = json.load(_f)

# ── Brand behavior config (brand_config.json) ────────────────────────────────
# Separate from tokens.json: controls WHEN/HOW brand elements appear,
# not what they look like. UI layer writes this file; builders read BT constants.
try:
    with open(_CONFIG_PATH, encoding="utf-8") as _fc:
        _C = json.load(_fc)
except (FileNotFoundError, ValueError):
    _C = {}

_vm  = _C.get("visual_mode", {})
_cr  = _C.get("card_rules",  {})

# Whether dark (#0E1216) accent cards are allowed at all.
# Set to False to keep every slide light and airy regardless of script intent.
DARK_ACCENT_CARDS_ENABLED: bool = _vm.get("dark_accent_cards", {}).get("enabled", True)

# Minimum cards on a slide before a dark accent card is permitted.
# Below this count the slide uses the full light palette only.
MIN_CARDS_FOR_DARK: int = int(_cr.get("min_cards_for_dark", 5))

# Maximum dark accent cards allowed per slide (applied after the min-cards gate).
MAX_DARK_PER_SLIDE: int = int(_cr.get("max_dark_per_slide", 1))

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
# DANGER_HEX — 仅限"危险、警告、风险"场景。
# 使用方式与其他语义卡片一致：CARD_DANGER_BG 作为卡片底色，DANGER_HEX 作为 accent/文字色。
# 严禁在：数据统计卡片、目录、过渡页、普通列举中使用此色。
DANGER_HEX        = _T["colors"]["semantic"]["danger"]   # #F12D2D

# ── Supplement colors (from tokens.json › colors.supplement) ─────────────────
BLUE_HEX   = _T["colors"]["supplement"]["blue"]    # #007BFF
TEAL_HEX   = _T["colors"]["supplement"]["teal"]    # #3CC5CF
PURPLE_HEX = _T["colors"]["supplement"]["purple"]  # #8255E1
ORANGE_HEX = _T["colors"]["supplement"]["orange"]  # #FFB928  (≡ WARNING_HEX)

# ── Card background tints — 语义卡片底色（浅色 tint 规律）────────────────────────
# 所有语义卡片统一遵循：浅色底（"100"级）+ 对应的 accent 色（"500"级）作文字/装饰色。
# 与 tokens.json 的 primary-100/500、secondary-100/500 同一逻辑，
# 补充色（supplement）没有写入 tokens.json，在此命名以保持一致性。
#
# 使用优先级（参见母版 P11）：
#  ★★★ 优先推荐  底色常量          │ accent 常量          │ 语义
#  ────────────────────────────────┼─────────────────────┼──────────────────────
#  ★★★           PRIMARY_100_HEX   │ PRIMARY_500_HEX      │ 标准 / 主要特性
#  ★★★           NEUTRAL_100_HEX   │ SUCCESS_HEX          │ 安全 / 补充 / 次要
#  ★★★           SECONDARY_100_HEX │ SECONDARY_500_HEX    │ 创新 / 机遇
#  ★★★           CARD_ORANGE_BG    │ WARNING_HEX          │ 高风险 / 注意（优先于红色）
#  ★★★           CARD_TEAL_BG      │ TEAL_HEX             │ 扩张 / 生态
#  ★★★ 深色卡    NEUTRAL_900_HEX   │ SECONDARY_500_HEX    │ 突出 / 亮点 / 强调（非危险）
#  ─── 补充 ─────────────────────────────────────────────────────────────────────
#  ★★  补充       CARD_PURPLE_BG    │ PURPLE_HEX           │ 战略 / 特殊（颜色不够时用）
#  ─── 慎用 ─────────────────────────────────────────────────────────────────────
#  ★   慎用       CARD_DANGER_BG    │ DANGER_HEX           │ 危险 / 风险（确实是危险才用）
CARD_ORANGE_BG  = "#FFF1DF"   # light tint of WARNING  #FFB928
CARD_TEAL_BG    = "#E0F7FA"   # light tint of TEAL     #3CC5CF
CARD_PURPLE_BG  = "#F0E8FF"   # light tint of PURPLE   #8255E1
CARD_DANGER_BG  = "#FFF2F2"   # light tint of DANGER   #F12D2D  (同 brand.css .brand-note--danger)

# ── Chart & extended palettes ─────────────────────────────────────────────────
# CHART_PALETTE — 6 slots from tokens.json；用于折线图/柱状图/饼图等数据可视化。
# 注意：最后一槽 #F12D2D（红色）仅用于"负向/危险"系列（如超标、亏损、风险项），
# 中性数据系列请优先用前 5 槽。
CHART_PALETTE: list = _T["charts"]["palette"]
# ['#53AF36', '#007BFF', '#3CC5CF', '#8255E1', '#FFB928', '#F12D2D']

# EXTENDED_PALETTE — 11 slots, 最大感知对比度排列。
# 用于 8+ 张白底卡片（neutral 模式）及多系列图表；8 张卡片无重复色。
# 刻意排除 DANGER_HEX（#F12D2D）：本品牌定位"自然、科技、可持续、安全"，
# 红色与整体视觉调性不符，不作为通用数据色使用。
# Slot:  0          1        2          3         4        5
#        brand-grn  blue     dk-green   purple    orange   teal
#        6          7        8          9         10
#        yel-green  suc-grn  near-blk   dk-gray   mid-gray
EXTENDED_PALETTE: list = [
    PRIMARY_500_HEX,   # #3EC99E  品牌主绿（明亮）
    BLUE_HEX,          # #007BFF  蓝色
    PRIMARY_600_HEX,   # #319E7C  品牌深绿（与主绿同族但明度差异大）
    PURPLE_HEX,        # #8255E1  紫色
    WARNING_HEX,       # #FFB928  橙色（表示"提示/注意"，非红色危险）
    TEAL_HEX,          # #3CC5CF  青色
    SECONDARY_500_HEX, # #C8E13C  黄绿色
    SUCCESS_HEX,       # #5CC13C  成功绿
    NEUTRAL_900_HEX,   # #0E1216  近黑（第9槽起循环用于需要高对比的场景）
    NEUTRAL_700_HEX,   # #3D444A  深灰
    NEUTRAL_400_HEX,   # #8A9199  中灰
]

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
FONT_CN     = "Alibaba PuHuiTi 2.0"   # 系统家族名（无下划线）
FONT_CN_FB  = "PingFang SC"           # 降级备用：Mac 自带
FONT_CN_FB2 = "Microsoft YaHei"       # 降级备用：Windows
FONT_CN_WEB = "Noto Sans SC"          # 降级备用：Web Google Fonts
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

# ── Radius (mm) — PPTX / DOCX 物理尺寸文档专用 ───────────────────────────────
# tokens.json › radiusMm（px 单位的 radiusPx 节仅用于 HTML/CSS）
# 用法：_card(..., radius_mm=BT.RADIUS_SM_MM)
RADIUS_XS_MM   = _T["radiusMm"]["xs"]    # 2mm  细节圆角（角标、小装饰点）
RADIUS_SM_MM   = _T["radiusMm"]["sm"]    # 4mm  ★ 默认：所有普通卡片、按钮
RADIUS_MD_MM   = _T["radiusMm"]["md"]    # 6mm  大卡片、面板（备用）
RADIUS_LG_MM   = _T["radiusMm"]["lg"]    # 8mm  外层大容器 / 图片占位区块
RADIUS_PILL_MM = _T["radiusMm"]["pill"]  # 9999 胶囊标签 / 序号徽章 / 圆形点

# ── 圆角层级（官方推荐搭配）────────────────────────────────────────────────────
# 同一页面建议遵循以下层级，最多两档混用：
#
#   外层容器  RADIUS_LG_MM   (8mm)  → 大面板、图片区块、背景容器
#   内层卡片  RADIUS_SM_MM   (4mm)  → 所有普通卡片（三卡/六卡/统计卡）★ 默认值
#   胶囊元素  RADIUS_PILL_MM  (∞)   → 状态徽章、分类标签、序号圆点
#   微小元素  RADIUS_XS_MM   (2mm)  → 角标、小装饰点（谨慎使用）
#
# 禁止在同一页面混用三档或以上圆角。

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
