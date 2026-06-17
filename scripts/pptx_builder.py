"""
BrandPptx — Brand-consistent PPTX builder for 未来方舟 / ArktechX  v2.0

Usage:
    from scripts.pptx_builder import BrandPptx

    prs = BrandPptx()
    prs.add_cover("AI 智能化解决方案", "ArktechX 2025")
    prs.add_cover_light("轻色封面版本", "Sub-Title")
    prs.add_divider("第一章：产品概述")
    prs.add_body_slide("核心功能", ["要点一", "要点二"])
    prs.add_big_stats("核心指标", [("98%", "准确率", "自动核算"), ...])
    prs.add_three_cards("三大优势", [{"title": "...", "body": "..."}, ...])
    prs.add_timeline("路线图", [("2024 Q1", "里程碑A", "说明"), ...])
    # add_timeline is adaptive: ≤5 → horizontal, 6+ → vertical two-column
    prs.add_quote("引用内容", "作者姓名", "职位")
    prs.add_table_slide("数据对比", ["列1","列2"], [["A","B"], ...])
    prs.add_closing("谢谢", "contact@futureark.ai")
    prs.save("output.pptx")

Gradient title note:
    add_cover_light / add_cover / add_closing / add_divider all use the
    brand title gradient: primary-500 → primary-600 → secondary-500.
    This mirrors the VS template's gradFill text approach.
"""

import os
from typing import List, Optional, Tuple, Dict, Any
import scripts.brand_tokens as BT

from pptx import Presentation
from pptx.util import Inches, Pt, Mm, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from lxml import etree


def _rgb(hex_color: str) -> RGBColor:
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ── Slide dimensions (16:9) ───────────────────────────────────────────────────
SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

ML = Mm(21)              # left / right margin
MR = Mm(21)
CW = SLIDE_W - ML - MR  # usable content width

HEADER_H  = Mm(36)
FOOTER_H  = Mm(13)
CONTENT_Y = HEADER_H
CONTENT_H = SLIDE_H - CONTENT_Y - FOOTER_H

C2_GAP = Mm(6)
C2_W   = (CW - C2_GAP) // 2
C3_GAP = Mm(5)
C3_W   = (CW - 2 * C3_GAP) // 3

# Logo (horizontal) dimensions for footer
_LOGO_H = Mm(7)
_LOGO_W = int(_LOGO_H * BT.LOGO_HORIZONTAL_ASPECT)

# Brand gradient stops (matches VS template palette, mapped to our tokens)
TITLE_GRADIENT = [
    (0,   BT.PRIMARY_500_HEX),    # #3EC99E
    (48,  BT.SUCCESS_HEX),        # #5CC13C  (semantic/success)
    (100, BT.SECONDARY_500_HEX),  # #C8E13C
]


# ── Low-level primitives ──────────────────────────────────────────────────────

def _rect(slide, l, t, w, h, fill=None, line=None, lw_pt=0.75, rounded=False):
    sid = 5 if rounded else 1
    s = slide.shapes.add_shape(sid, int(l), int(t), int(w), int(h))
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = _rgb(fill)
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = _rgb(line)
        s.line.width = Pt(lw_pt)
    else:
        s.line.fill.background()
    return s


def _set_run_fonts(run, en_font=BT.FONT_EN, cn_font=BT.FONT_CN):
    """Set latin + east-asian font on a run."""
    run.font.name = en_font
    rPr = run._r.get_or_add_rPr()
    for tag, face in [("a:latin", en_font), ("a:ea", cn_font)]:
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", face)


def _apply_gradient_to_run(run, stops=None, angle_deg=0):
    """
    Replace the run's color with a gradient fill (DrawingML gradFill).
    stops: list of (pct_0_to_100, '#hexcolor')
    angle_deg: 0 = left→right, 90 = top→bottom
    """
    if stops is None:
        stops = TITLE_GRADIENT
    ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    Q = lambda tag: f"{{{ns}}}{tag}"

    rPr = run._r.get_or_add_rPr()

    # Remove existing fill elements
    for tag in ["solidFill", "gradFill", "pattFill", "noFill", "blipFill"]:
        el = rPr.find(Q(tag))
        if el is not None:
            rPr.remove(el)

    gf = etree.Element(Q("gradFill"))
    gf.set("rotWithShape", "1")

    gsLst = etree.SubElement(gf, Q("gsLst"))
    for pct, hex_color in stops:
        gs = etree.SubElement(gsLst, Q("gs"))
        gs.set("pos", str(int(pct * 1000)))
        srgb = etree.SubElement(gs, Q("srgbClr"))
        srgb.set("val", hex_color.lstrip("#"))
        alpha_el = etree.SubElement(srgb, Q("alpha"))
        alpha_el.set("val", "100000")

    lin = etree.SubElement(gf, Q("lin"))
    lin.set("ang", str(int(angle_deg * 60000)))

    # Insert before any font spec elements so rendering is correct
    rPr.insert(0, gf)


def _txb(slide, text, l, t, w, h,
         sz=16, bold=False, color=None,
         align=PP_ALIGN.LEFT, wrap=True,
         ls_pt=None,
         cn_font=BT.FONT_CN, en_font=BT.FONT_EN):
    """Add a plain text box. color=None keeps default (white)."""
    if color is None:
        color = BT.NEUTRAL_700_HEX
    tb = slide.shapes.add_textbox(int(l), int(t), int(w), int(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    if ls_pt:
        pPr = p._p.get_or_add_pPr()
        lnSpc = etree.SubElement(pPr, qn("a:lnSpc"))
        etree.SubElement(lnSpc, qn("a:spcPts"),
                         attrib={"val": str(int(ls_pt * 100))})
    run = p.add_run()
    run.text = text
    run.font.size = Pt(sz)
    run.font.bold = bold
    run.font.color.rgb = _rgb(color)
    _set_run_fonts(run, en_font=en_font, cn_font=cn_font)
    return tb


def _txb_gradient(slide, text, l, t, w, h,
                  sz=40, bold=True,
                  stops=None, angle_deg=0,
                  align=PP_ALIGN.LEFT, wrap=True,
                  ls_pt=None,
                  cn_font=BT.FONT_CN, en_font=BT.FONT_EN):
    """Add a text box with gradient-filled text (like VS template title)."""
    if stops is None:
        stops = TITLE_GRADIENT
    tb = slide.shapes.add_textbox(int(l), int(t), int(w), int(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    if ls_pt:
        pPr = p._p.get_or_add_pPr()
        lnSpc = etree.SubElement(pPr, qn("a:lnSpc"))
        etree.SubElement(lnSpc, qn("a:spcPts"),
                         attrib={"val": str(int(ls_pt * 100))})
    run = p.add_run()
    run.text = text
    run.font.size = Pt(sz)
    run.font.bold = bold
    # Set fonts first, then replace color with gradient
    _set_run_fonts(run, en_font=en_font, cn_font=cn_font)
    _apply_gradient_to_run(run, stops=stops, angle_deg=angle_deg)
    return tb


def _set_slide_bg(slide, hex_color: str):
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = _rgb(hex_color)


def _add_logo_h(slide, right_mm=9, bottom_mm=3, h_mm=7, reverse=False):
    """Add horizontal logo to bottom-right."""
    if reverse:
        path = BT.LOGO_HORIZONTAL_REVERSE_PNG or BT.LOGO_HORIZONTAL_PRIMARY_PNG
    else:
        path = BT.LOGO_HORIZONTAL_PRIMARY_PNG
    if not (path and os.path.exists(path)):
        return None
    lh = Mm(h_mm)
    lw = int(lh * BT.LOGO_HORIZONTAL_ASPECT)
    return slide.shapes.add_picture(
        path,
        int(SLIDE_W - Mm(right_mm) - lw),
        int(SLIDE_H - Mm(bottom_mm) - lh),
        width=lw, height=lh)


def _add_logo_stacked(slide, reverse=True, l_mm=12, t_mm=10, h_mm=26):
    """Add stacked logo (cover / section pages)."""
    path = BT.LOGO_STACKED_REVERSE_PNG if reverse else BT.LOGO_STACKED_PRIMARY_PNG
    if not (path and os.path.exists(path)):
        return None
    lh = Mm(h_mm)
    from PIL import Image as _PILImage
    img = _PILImage.open(path)
    lw = int(lh * img.width / img.height)
    return slide.shapes.add_picture(path, Mm(l_mm), Mm(t_mm), width=lw, height=lh)


# ── Composite helpers: header / footer / card ─────────────────────────────────

def _header(slide, title, subtitle="", label=""):
    """Standard page header: title + optional subtitle."""
    y_title = Mm(5.5) if not label else Mm(11)
    if label:
        _txb(slide, label, l=ML, t=Mm(4.5), w=Mm(80), h=Mm(6),
             sz=9, color=BT.PRIMARY_500_HEX)
    _txb(slide, title, l=ML, t=y_title, w=CW * 0.82, h=Mm(18),
         sz=26, bold=True, color=BT.NEUTRAL_900_HEX)
    if subtitle:
        _txb(slide, subtitle, l=ML, t=Mm(24 if not label else 27),
             w=CW * 0.80, h=Mm(9), sz=12, color=BT.NEUTRAL_400_HEX)


def _footer(slide, layout_label=""):
    """Standard page footer: logo + optional layout label."""
    _add_logo_h(slide, right_mm=9, bottom_mm=3, h_mm=7)
    if layout_label:
        _txb(slide, layout_label,
             l=SLIDE_W - Mm(72), t=SLIDE_H - FOOTER_H + Mm(1),
             w=Mm(60), h=Mm(10),
             sz=8, color=BT.NEUTRAL_200_HEX, align=PP_ALIGN.RIGHT)


def _card(slide, l, t, w, h, bg=None, border=None, rounded=True):
    """Card-shaped rectangle. Default: light surface + subtle border."""
    if bg is None:
        bg = BT.NEUTRAL_100_HEX
    if border is None:
        border = BT.NEUTRAL_200_HEX
    s = slide.shapes.add_shape(5 if rounded else 1, int(l), int(t), int(w), int(h))
    if rounded:
        adj = s._element.find(qn("p:spPr")).find(qn("a:prstGeom"))
        if adj is not None:
            avLst = adj.find(qn("a:avLst"))
            if avLst is None:
                avLst = etree.SubElement(adj, qn("a:avLst"))
            gd = avLst.find(qn("a:gd"))
            if gd is None:
                gd = etree.SubElement(avLst, qn("a:gd"))
            gd.set("name", "adj")
            gd.set("fmla", "val 20000")
    s.fill.solid()
    s.fill.fore_color.rgb = _rgb(bg)
    s.line.color.rgb = _rgb(border)
    s.line.width = Pt(0.75)
    return s


# ── BrandPptx ─────────────────────────────────────────────────────────────────

class BrandPptx:
    """
    Brand-consistent PPTX builder v2.0.
    Automatically uses templates/ppt/brand_master.pptx as the base template
    if it exists, so that theme colors are properly embedded.
    """

    _TEMPLATE_PATH = os.path.abspath(
        os.path.join(os.path.dirname(__file__), "..", "templates", "ppt", "brand_master.pptx")
    )

    def __init__(self, use_template: bool = True):
        if use_template and os.path.exists(self._TEMPLATE_PATH):
            self._prs = Presentation(self._TEMPLATE_PATH)
            self._clear_slides()
        else:
            self._prs = Presentation()
        self._prs.slide_width  = SLIDE_W
        self._prs.slide_height = SLIDE_H
        self._blank = self._prs.slide_layouts[6]  # blank layout

    def _clear_slides(self):
        """Remove all content slides, keeping master and layouts."""
        r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        sldIdLst = self._prs.slides._sldIdLst
        for sldId in list(sldIdLst):
            rId = sldId.get(f"{{{r_ns}}}id")
            if rId:
                try:
                    self._prs.part.drop_rel(rId)
                except Exception:
                    pass
            sldIdLst.remove(sldId)

    def _new_slide(self):
        return self._prs.slides.add_slide(self._blank)

    # ── Cover Slide — Dark ────────────────────────────────────────────────────

    def add_cover(self, title: str, subtitle: str = "", tagline: str = ""):
        """
        Dark cover (#0E1216 bg) with gradient title text.
        Stacked reverse logo top-left, green accent bar left, bottom bar.
        """
        slide = self._new_slide()
        _set_slide_bg(slide, BT.NEUTRAL_900_HEX)

        # Left accent bar
        _rect(slide, l=0, t=Mm(28), w=Mm(3.5), h=Mm(85),
              fill=BT.PRIMARY_500_HEX)

        # Stacked reverse logo top-left
        _add_logo_stacked(slide, reverse=True, l_mm=13, t_mm=11, h_mm=26)

        # Gradient title
        _txb_gradient(slide, title,
                      l=Mm(22), t=Mm(65), w=Mm(205), h=Mm(38),
                      sz=42, bold=True, align=PP_ALIGN.LEFT,
                      stops=TITLE_GRADIENT)

        # Subtitle
        if subtitle:
            _txb(slide, subtitle,
                 l=Mm(22), t=Mm(107), w=Mm(205), h=Mm(16),
                 sz=16, color=BT.NEUTRAL_400_HEX)

        # Tagline
        if tagline:
            _txb(slide, tagline,
                 l=Mm(22), t=Mm(126), w=Mm(205), h=Mm(10),
                 sz=11, color=BT.PRIMARY_100_HEX)

        # Bottom bar
        _rect(slide, l=0, t=SLIDE_H - Mm(3.5), w=SLIDE_W, h=Mm(3.5),
              fill=BT.PRIMARY_500_HEX)
        return slide

    # ── Cover Slide — Light (VS Template Style) ───────────────────────────────

    def add_cover_light(self, title: str, subtitle: str = "",
                        date_or_meta: str = "", tagline: str = ""):
        """
        Light cover (#FAFAFA bg) with large gradient title — mirrors the VS template.
        Primary logo top-right, thin separator line, gradient 96pt title.
        """
        slide = self._new_slide()
        _set_slide_bg(slide, BT.BG_PAGE_HEX)   # #F8FAFC — page-level bg

        # Thin left accent bar
        _rect(slide, l=0, t=0, w=Mm(3), h=SLIDE_H, fill=BT.PRIMARY_500_HEX)

        # Horizontal separator line
        _rect(slide, l=Mm(16), t=Mm(90), w=SLIDE_W - Mm(20), h=Mm(0.5),
              fill=BT.NEUTRAL_200_HEX)

        # Gradient title (large, 96pt equivalent ≈ 44pt at 1333px canvas)
        _txb_gradient(slide, title,
                      l=Mm(22), t=Mm(94), w=Mm(200), h=Mm(42),
                      sz=44, bold=True, align=PP_ALIGN.LEFT,
                      stops=TITLE_GRADIENT, ls_pt=52)

        # Subtitle
        if subtitle:
            _txb(slide, subtitle,
                 l=Mm(22), t=Mm(139), w=Mm(200), h=Mm(14),
                 sz=15, bold=False, color=BT.NEUTRAL_900_HEX)

        # Meta / date
        if date_or_meta:
            _txb(slide, date_or_meta,
                 l=Mm(22), t=Mm(155), w=Mm(200), h=Mm(10),
                 sz=12, color=BT.NEUTRAL_400_HEX)

        # Tagline
        if tagline:
            _txb(slide, tagline,
                 l=Mm(22), t=Mm(168), w=Mm(200), h=Mm(10),
                 sz=11, color=BT.NEUTRAL_400_HEX)

        # Primary stacked logo — top left
        _add_logo_stacked(slide, reverse=False, l_mm=13, t_mm=14, h_mm=24)

        return slide

    # ── Section Divider ───────────────────────────────────────────────────────

    def add_divider(self, chapter_title: str, chapter_num: str = ""):
        """
        Section divider: green gradient background, white chapter title.
        """
        slide = self._new_slide()

        _rect(slide, l=0, t=0, w=SLIDE_W, h=SLIDE_H, fill=BT.PRIMARY_500_HEX)

        if chapter_num:
            _txb(slide, chapter_num,
                 l=ML, t=Mm(60), w=SLIDE_W - 2*ML, h=Mm(14),
                 sz=13, color=BT.PRIMARY_100_HEX,
                 align=PP_ALIGN.CENTER)

        _txb(slide, chapter_title,
             l=ML, t=Mm(78), w=SLIDE_W - 2*ML, h=Mm(44),
             sz=38, bold=True, color=BT.WHITE_HEX, align=PP_ALIGN.CENTER,
             ls_pt=46)

        # Logo reversed bottom-right
        _add_logo_stacked(slide, reverse=True,
                          l_mm=int((SLIDE_W / Mm(1) - 30)),
                          t_mm=int((SLIDE_H / Mm(1) - 28)),
                          h_mm=18)
        return slide

    # ── Standard Body Slide ───────────────────────────────────────────────────

    def add_body_slide(self,
                       title: str,
                       bullets: Optional[List[str]] = None,
                       body_text: str = "",
                       subtitle: str = "",
                       slide_label: str = ""):
        """Standard white body slide with header, bullets or free text."""
        slide = self._new_slide()
        _set_slide_bg(slide, BT.WHITE_HEX)
        _header(slide, title, subtitle=subtitle)
        _footer(slide, layout_label=slide_label)

        content_top = CONTENT_Y + Mm(4)
        content_h   = CONTENT_H - Mm(4)

        if bullets:
            tb = slide.shapes.add_textbox(
                ML, content_top, CW, content_h)
            tf = tb.text_frame
            tf.word_wrap = True
            for i, bullet in enumerate(bullets):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                pPr = p._p.get_or_add_pPr()
                lnSpc = etree.SubElement(pPr, qn("a:lnSpc"))
                etree.SubElement(lnSpc, qn("a:spcPts"), attrib={"val": "2100"})
                spc_bef = etree.SubElement(pPr, qn("a:spcBef"))
                etree.SubElement(spc_bef, qn("a:spcPts"), attrib={"val": "800"})
                run = p.add_run()
                run.text = f"• {bullet}"
                run.font.size = Pt(17)
                run.font.color.rgb = _rgb(BT.NEUTRAL_700_HEX)
                _set_run_fonts(run)
        elif body_text:
            _txb(slide, body_text,
                 l=ML, t=content_top, w=CW, h=content_h,
                 sz=16, color=BT.NEUTRAL_700_HEX, ls_pt=26)
        return slide

    # ── Two-Column Slide ──────────────────────────────────────────────────────

    def add_two_col_slide(self,
                          title: str,
                          left_content: str,
                          right_content: str,
                          left_title: str = "",
                          right_title: str = "",
                          subtitle: str = ""):
        """Two-column body slide with optional column subtitles."""
        slide = self._new_slide()
        _set_slide_bg(slide, BT.WHITE_HEX)
        _header(slide, title, subtitle=subtitle)
        _footer(slide)

        top    = CONTENT_Y + Mm(5)
        height = CONTENT_H - Mm(5)

        for i, (col_title, content) in enumerate([
            (left_title,  left_content),
            (right_title, right_content),
        ]):
            left_x = ML + i * (C2_W + C2_GAP)
            if col_title:
                _txb(slide, col_title,
                     l=left_x, t=top, w=C2_W, h=Mm(12),
                     sz=13, bold=True, color=BT.PRIMARY_500_HEX)
                body_top = top + Mm(14)
            else:
                body_top = top
            _txb(slide, content,
                 l=left_x, t=body_top, w=C2_W, h=height - Mm(14),
                 sz=14, color=BT.NEUTRAL_700_HEX, ls_pt=22)

        # Vertical divider
        _rect(slide,
              l=ML + C2_W + C2_GAP // 2,
              t=top + Mm(4),
              w=Mm(0.4),
              h=height - Mm(8),
              fill=BT.NEUTRAL_200_HEX)
        return slide

    # ── Three-Cards Slide ─────────────────────────────────────────────────────

    def add_three_cards(self,
                        title: str,
                        cards: List[Dict[str, str]],
                        subtitle: str = ""):
        """
        Three-column feature cards.
        cards: [{"title": "...", "body": "...", "tag": "(optional)"}]
        """
        slide = self._new_slide()
        _set_slide_bg(slide, BT.WHITE_HEX)
        _header(slide, title, subtitle=subtitle)
        _footer(slide)

        card_top = CONTENT_Y + Mm(6)
        card_h   = CONTENT_H - Mm(6)

        card_colors = [BT.PRIMARY_100_HEX, BT.NEUTRAL_100_HEX, "#F8FBE7"]
        border_colors = ["#B8EDD8", BT.NEUTRAL_200_HEX, "#DBE89A"]
        accent_colors = [BT.PRIMARY_500_HEX, BT.SUCCESS_HEX, BT.SECONDARY_500_HEX]

        for i, c in enumerate(cards[:3]):
            x = ML + i * (C3_W + C3_GAP)
            _card(slide, l=x, t=card_top, w=C3_W, h=card_h,
                  bg=card_colors[i % 3], border=border_colors[i % 3])

            # Top accent strip
            _rect(slide, l=x, t=card_top, w=C3_W, h=Mm(3),
                  fill=accent_colors[i % 3])

            inner_l = x + Mm(5)
            inner_w = C3_W - Mm(10)
            y_off   = card_top + Mm(8)

            # Optional tag
            tag = c.get("tag", "")
            if tag:
                _txb(slide, tag, l=inner_l, t=y_off, w=inner_w, h=Mm(8),
                     sz=9, bold=True, color=accent_colors[i % 3])
                y_off += Mm(9)

            # Card title — h=Mm(24) handles 2-line titles at sz=16
            _txb(slide, c.get("title", ""),
                 l=inner_l, t=y_off, w=inner_w, h=Mm(24),
                 sz=16, bold=True, color=BT.NEUTRAL_900_HEX)
            y_off += Mm(26)

            # Card body — auto_size shrinks font to fit, never clips
            body_tb = _txb(slide, c.get("body", ""),
                           l=inner_l, t=y_off, w=inner_w,
                           h=card_h - (y_off - card_top) - Mm(6),
                           sz=13, color=BT.NEUTRAL_700_HEX, ls_pt=18)
            body_tb.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

        return slide

    # ── Six-Cards Slide (2 × 3 grid) ─────────────────────────────────────────

    def add_six_cards(self,
                      title: str,
                      cards: List[Dict[str, str]],
                      subtitle: str = ""):
        """
        2-row × 3-column compact cards grid.
        cards: [{"title": "...", "body": "...", "tag": "(optional)"}] — up to 6
        Column index drives accent colour so each column pair shares a theme.
        """
        slide = self._new_slide()
        _set_slide_bg(slide, BT.WHITE_HEX)
        _header(slide, title, subtitle=subtitle)
        _footer(slide)

        ROW_GAP = Mm(4)
        PAD_TOP = Mm(4)
        card_h  = (CONTENT_H - PAD_TOP - ROW_GAP) // 2

        card_colors   = [BT.PRIMARY_100_HEX, BT.NEUTRAL_100_HEX, "#F8FBE7"]
        border_colors = ["#B8EDD8", BT.NEUTRAL_200_HEX, "#DBE89A"]
        accent_colors = [BT.PRIMARY_500_HEX, BT.SUCCESS_HEX, BT.SECONDARY_500_HEX]

        for i, c in enumerate(cards[:6]):
            col = i % 3
            row = i // 3
            x   = ML + col * (C3_W + C3_GAP)
            y   = CONTENT_Y + PAD_TOP + row * (card_h + ROW_GAP)

            _card(slide, l=x, t=y, w=C3_W, h=card_h,
                  bg=card_colors[col % 3], border=border_colors[col % 3])
            _rect(slide, l=x, t=y, w=C3_W, h=Mm(2.5),
                  fill=accent_colors[col % 3])

            inner_l = x + Mm(4)
            inner_w = C3_W - Mm(8)
            y_off   = y + Mm(6)

            tag = c.get("tag", "")
            if tag:
                _txb(slide, tag, l=inner_l, t=y_off, w=inner_w, h=Mm(7),
                     sz=8, bold=True, color=accent_colors[col % 3])
                y_off += Mm(8)

            _txb(slide, c.get("title", ""),
                 l=inner_l, t=y_off, w=inner_w, h=Mm(16),
                 sz=14, bold=True, color=BT.NEUTRAL_900_HEX)
            y_off += Mm(18)

            body_h = card_h - (y_off - y) - Mm(4)
            if body_h > Mm(8):
                body_tb = _txb(slide, c.get("body", ""),
                               l=inner_l, t=y_off, w=inner_w, h=body_h,
                               sz=11, color=BT.NEUTRAL_700_HEX, ls_pt=16)
                body_tb.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

        return slide

    # ── Big Stats Slide ───────────────────────────────────────────────────────

    def add_big_stats(self,
                      title: str,
                      stats: List[Tuple[str, str, str]],
                      subtitle: str = ""):
        """
        2×2 stats grid.
        stats: [("98%", "准确率", "核算自动化"), ...]  → (value, label, desc)
        Supports 2 or 4 stats. Values get gradient text treatment.
        """
        slide = self._new_slide()
        _set_slide_bg(slide, BT.WHITE_HEX)
        _header(slide, title, subtitle=subtitle)
        _footer(slide)

        stats = stats[:4]
        cols   = 2 if len(stats) <= 2 else 2
        rows   = (len(stats) + cols - 1) // cols

        gap    = Mm(8)
        card_w = (CW - gap) // 2
        card_h = (CONTENT_H - gap - Mm(6)) // rows

        card_bgs    = [BT.PRIMARY_100_HEX, BT.NEUTRAL_100_HEX,
                       BT.NEUTRAL_100_HEX, "#F8FBE7"]
        card_border = ["#B8EDD8", BT.NEUTRAL_200_HEX,
                       BT.NEUTRAL_200_HEX, "#DBE89A"]
        # Solid value colors — same family as card bg, higher saturation for contrast
        val_colors  = [BT.PRIMARY_500_HEX,   # teal-100 bg → teal-500
                       BT.NEUTRAL_900_HEX,   # gray-100 bg → near-black
                       BT.NEUTRAL_900_HEX,   # gray-100 bg → near-black
                       BT.SUCCESS_HEX]       # yellow-green bg → success green

        for i, (val, label, desc) in enumerate(stats):
            col = i % 2
            row = i // 2
            x   = ML + col * (card_w + gap)
            y   = CONTENT_Y + Mm(6) + row * (card_h + gap)

            _card(slide, l=x, t=y, w=card_w, h=card_h,
                  bg=card_bgs[i], border=card_border[i])

            # Value — solid color (gradient reserved for covers/section/closing)
            _txb(slide, val,
                 l=x + Mm(6), t=y + Mm(8),
                 w=card_w - Mm(12), h=Mm(28),
                 sz=44, bold=True, align=PP_ALIGN.LEFT,
                 color=val_colors[i])

            # Label
            _txb(slide, label,
                 l=x + Mm(6), t=y + Mm(37),
                 w=card_w - Mm(12), h=Mm(10),
                 sz=14, bold=True, color=BT.NEUTRAL_900_HEX)

            # Description
            if desc:
                _txb(slide, desc,
                     l=x + Mm(6), t=y + Mm(49),
                     w=card_w - Mm(12), h=card_h - Mm(55),
                     sz=12, color=BT.NEUTRAL_400_HEX, ls_pt=18)

        return slide

    # ── Timeline Slide ────────────────────────────────────────────────────────

    def add_timeline(self,
                     title: str,
                     milestones: List[Tuple[str, str, str]],
                     subtitle: str = ""):
        """
        Adaptive timeline — no hard cap on item count.
        - 1–6  items → single horizontal row (full content height)
        - 7–12 items → two-row wrap: top row = n//2 items, bottom = remainder
          Both rows share the same visual language; split is as balanced as possible.
        milestones: [("period", "title", "description"), ...]
        """
        n = len(milestones)
        slide = self._new_slide()
        _set_slide_bg(slide, BT.WHITE_HEX)
        _header(slide, title, subtitle=subtitle)
        _footer(slide)
        if n == 0:
            return slide

        if n <= 6:
            # Single row — full content area
            self._draw_timeline_row(
                slide, milestones,
                row_top=CONTENT_Y + Mm(2),
                row_h=CONTENT_H - Mm(4),
            )
        else:
            # Two-row wrap
            # Row 1 spine extends to slide right edge;
            # Row 2 spine extends from slide left edge.
            # This creates a visual "snake" showing timeline continuity.
            n1      = n // 2          # top row (floor → fewer or equal)
            ROW_GAP = Mm(7)
            PAD_V   = Mm(2)
            row_h   = (CONTENT_H - ROW_GAP - PAD_V * 2) // 2

            row1_top = CONTENT_Y + PAD_V
            row2_top = row1_top + row_h + ROW_GAP

            self._draw_timeline_row(slide, milestones[:n1],
                                    row1_top, row_h, extend_right=True)
            self._draw_timeline_row(slide, milestones[n1:],
                                    row2_top, row_h, extend_left=True)

        return slide

    def _draw_timeline_row(self,
                           slide,
                           items: List[Tuple[str, str, str]],
                           row_top: int,
                           row_h: int,
                           extend_left: bool = False,
                           extend_right: bool = False):
        """
        Draw one horizontal timeline row inside the given vertical bounds.
        Font sizes and dot radius scale automatically with item density.

        items:        [("period", "title", "description"), ...]
        row_top:      top-y of the row area (EMU)
        row_h:        total height of the row area (EMU)
        extend_left:  stretch spine from slide left edge (x=0) instead of
                      stopping at the first dot — used for the bottom wrap row
        extend_right: stretch spine to slide right edge (x=SLIDE_W) instead of
                      stopping at the last dot — used for the top wrap row
        """
        n = len(items)
        if n == 0:
            return

        shrink   = max(0, n - 3)
        sz_per   = max(9,  12 - shrink)
        sz_title = max(10, 14 - shrink)
        sz_desc  = max(9,  12 - shrink)
        ls_desc  = max(13, 17 - shrink)
        dot_r    = Mm(3.5) if n <= 2 else (Mm(3) if n <= 4 else Mm(2.5))

        item_w  = (CW - Mm(4)) // n
        # Spine sits at ~30 % from the top of the row area so period labels
        # have room above and descriptions have the majority of space below.
        spine_y = row_top + int(row_h * 0.30)
        dot_y   = spine_y - dot_r

        # Dot centre positions
        cx_first = ML + item_w // 2
        cx_last  = ML + item_w // 2 + (n - 1) * item_w

        # Spine: extend to slide edges when wrapping, otherwise dot-to-dot
        spine_l = 0        if extend_left  else cx_first
        spine_r = SLIDE_W  if extend_right else cx_last

        _rect(slide,
              l=spine_l,
              t=spine_y - Mm(0.7),
              w=spine_r - spine_l,
              h=Mm(1.4),
              fill=BT.PRIMARY_500_HEX)

        ACCENT = [BT.PRIMARY_500_HEX, BT.SUCCESS_HEX,
                  BT.SECONDARY_500_HEX, BT.SUCCESS_HEX,
                  BT.PRIMARY_500_HEX, BT.SUCCESS_HEX]

        for i, (period, m_title, desc) in enumerate(items):
            cx = ML + item_w // 2 + i * item_w
            ac = ACCENT[i % len(ACCENT)]

            # White halo so dot pops over spine
            _rect(slide,
                  l=cx - dot_r - Mm(1), t=dot_y - Mm(1),
                  w=(dot_r + Mm(1)) * 2, h=(dot_r + Mm(1)) * 2,
                  fill=BT.WHITE_HEX, rounded=True)
            _rect(slide,
                  l=cx - dot_r, t=dot_y,
                  w=dot_r * 2, h=dot_r * 2,
                  fill=ac, rounded=True)

            # Period label — above spine
            _txb(slide, period,
                 l=cx - item_w // 2, t=row_top + Mm(2),
                 w=item_w, h=spine_y - (row_top + Mm(2)) - dot_r - Mm(1),
                 sz=sz_per, bold=True, color=ac, align=PP_ALIGN.CENTER)

            # Title — below dot
            title_top = spine_y + dot_r + Mm(3)
            title_h   = Mm(14)
            _txb(slide, m_title,
                 l=cx - item_w // 2 + Mm(2), t=title_top,
                 w=item_w - Mm(4), h=title_h,
                 sz=sz_title, bold=True, color=BT.NEUTRAL_900_HEX,
                 align=PP_ALIGN.CENTER)

            # Description — fills remaining row height
            if desc:
                desc_top = title_top + title_h + Mm(2)
                desc_h   = row_top + row_h - desc_top - Mm(2)
                if desc_h > Mm(6):
                    _txb(slide, desc,
                         l=cx - item_w // 2 + Mm(2), t=desc_top,
                         w=item_w - Mm(4), h=desc_h,
                         sz=sz_desc, color=BT.NEUTRAL_400_HEX,
                         align=PP_ALIGN.CENTER, ls_pt=ls_desc)

    # ── Quote Slide ───────────────────────────────────────────────────────────

    def add_quote(self,
                  quote_text: str,
                  author: str = "",
                  role: str = "",
                  subtitle: str = ""):
        """
        Dark quote / testimonial slide.
        Large quote mark, body text white, author in gradient.
        """
        slide = self._new_slide()
        _set_slide_bg(slide, BT.NEUTRAL_900_HEX)

        # Big quotation mark
        _txb(slide, "“",
             l=ML, t=Mm(28), w=Mm(30), h=Mm(26),
             sz=56, bold=True, color=BT.PRIMARY_500_HEX)

        # Quote body
        _txb(slide, quote_text,
             l=ML + Mm(2), t=Mm(54),
             w=CW - Mm(4), h=Mm(52),
             sz=22, bold=False, color=BT.WHITE_HEX,
             align=PP_ALIGN.LEFT, ls_pt=33)

        # Thin accent line
        _rect(slide, l=ML, t=Mm(112), w=Mm(32), h=Mm(1.5),
              fill=BT.SECONDARY_500_HEX)

        # Author — solid primary color (gradient reserved for covers/section/closing)
        if author:
            _txb(slide, author,
                 l=ML, t=Mm(117), w=Mm(180), h=Mm(12),
                 sz=16, bold=True, color=BT.PRIMARY_500_HEX)

        # Role
        if role:
            _txb(slide, role,
                 l=ML, t=Mm(131), w=Mm(180), h=Mm(10),
                 sz=12, color=BT.NEUTRAL_400_HEX)

        # Optional subtitle label (top-right)
        if subtitle:
            _txb(slide, subtitle,
                 l=SLIDE_W - ML - Mm(80), t=Mm(12),
                 w=Mm(80), h=Mm(10),
                 sz=10, color=BT.NEUTRAL_400_HEX, align=PP_ALIGN.RIGHT)

        # Logo reversed
        _add_logo_h(slide, right_mm=10, bottom_mm=4, h_mm=7)

        # Bottom bar
        _rect(slide, l=0, t=SLIDE_H - Mm(3.5), w=SLIDE_W, h=Mm(3.5),
              fill=BT.PRIMARY_500_HEX)
        return slide

    # ── Table Slide ───────────────────────────────────────────────────────────

    def add_table_slide(self,
                        title: str,
                        headers: List[str],
                        rows: List[List[str]],
                        subtitle: str = "",
                        note: str = ""):
        """
        Data table slide.
        Header row: #3EC99E bg + white text.
        Zebra rows: alternate #F2F3F5 / white.
        """
        from pptx.util import Pt
        from pptx.oxml.ns import qn

        slide = self._new_slide()
        _set_slide_bg(slide, BT.WHITE_HEX)
        _header(slide, title, subtitle=subtitle)
        _footer(slide)

        n_cols = len(headers)
        n_rows = len(rows)
        if n_cols == 0:
            return slide

        t_top  = int(CONTENT_Y + Mm(5))
        t_h    = int(CONTENT_H - Mm(8) - (Mm(8) if note else 0))
        row_h  = min(int(t_h // (n_rows + 1)), int(Mm(12)))

        table  = slide.shapes.add_table(
            n_rows + 1, n_cols,
            int(ML), t_top,
            int(CW), row_h * (n_rows + 1)
        ).table

        # Header row
        for j, hdr in enumerate(headers):
            cell = table.cell(0, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(BT.PRIMARY_500_HEX)
            p = cell.text_frame.paragraphs[0]
            run = p.add_run()
            run.text = hdr
            run.font.size = Pt(12)
            run.font.bold = True
            run.font.color.rgb = _rgb(BT.WHITE_HEX)
            _set_run_fonts(run)
            p.alignment = PP_ALIGN.LEFT

        # Data rows
        for i, row in enumerate(rows):
            bg = BT.NEUTRAL_100_HEX if i % 2 == 0 else BT.WHITE_HEX
            for j, val in enumerate(row[:n_cols]):
                cell = table.cell(i + 1, j)
                cell.fill.solid()
                cell.fill.fore_color.rgb = _rgb(bg)
                p = cell.text_frame.paragraphs[0]
                run = p.add_run()
                run.text = str(val)
                run.font.size = Pt(11)
                run.font.color.rgb = _rgb(
                    BT.NEUTRAL_900_HEX if j == 0 else BT.NEUTRAL_700_HEX
                )
                _set_run_fonts(run)
                p.alignment = PP_ALIGN.LEFT

        # Note
        if note:
            _txb(slide, f"注：{note}",
                 l=ML, t=t_top + row_h * (n_rows + 1) + Mm(3),
                 w=CW, h=Mm(8),
                 sz=10, color=BT.NEUTRAL_400_HEX)

        return slide

    # ── Text + Image Layouts ──────────────────────────────────────────────────

    def add_text_image_right(self,
                             title: str,
                             body_text: str,
                             image_path: Optional[str] = None,
                             subtitle: str = ""):
        """Left 60% text, right 40% image placeholder."""
        slide = self._new_slide()
        _set_slide_bg(slide, BT.WHITE_HEX)
        _header(slide, title, subtitle=subtitle)
        _footer(slide)

        text_w  = int(CW * 0.58)
        img_w   = int(CW - text_w - Mm(6))
        img_l   = ML + text_w + Mm(6)

        _txb(slide, body_text,
             l=ML, t=CONTENT_Y + Mm(6), w=text_w,
             h=CONTENT_H - Mm(6), sz=14, ls_pt=22)

        if image_path and os.path.exists(image_path):
            slide.shapes.add_picture(
                image_path, img_l,
                int(CONTENT_Y + Mm(6)), img_w,
                int(CONTENT_H - Mm(8)))
        else:
            _card(slide, img_l, int(CONTENT_Y + Mm(6)),
                  img_w, int(CONTENT_H - Mm(8)),
                  bg=BT.PRIMARY_100_HEX, border="#B8EDD8")
        return slide

    def add_image_left_text(self,
                            title: str,
                            body_text: str,
                            image_path: Optional[str] = None,
                            subtitle: str = ""):
        """Left 40% image, right 60% text."""
        slide = self._new_slide()
        _set_slide_bg(slide, BT.WHITE_HEX)
        _header(slide, title, subtitle=subtitle)
        _footer(slide)

        img_w  = int(CW * 0.40)
        text_w = int(CW - img_w - Mm(6))
        text_l = ML + img_w + Mm(6)

        if image_path and os.path.exists(image_path):
            slide.shapes.add_picture(
                image_path, int(ML),
                int(CONTENT_Y + Mm(6)), img_w,
                int(CONTENT_H - Mm(8)))
        else:
            _card(slide, int(ML), int(CONTENT_Y + Mm(6)),
                  img_w, int(CONTENT_H - Mm(8)),
                  bg=BT.PRIMARY_100_HEX, border="#B8EDD8")

        _txb(slide, body_text,
             l=text_l, t=CONTENT_Y + Mm(6), w=text_w,
             h=CONTENT_H - Mm(6), sz=14, ls_pt=22)
        return slide

    # ── Table of Contents ─────────────────────────────────────────────────────

    def add_toc(self,
                title: str,
                chapters: List[Dict[str, str]],
                label: str = "TABLE OF CONTENTS",
                description: str = ""):
        """
        2×3 chapter card grid TOC.
        chapters: [{"num":"01","title":"关于我们","subtitle":"公司简介·愿景","state":"done|current|upcoming"}]

        state values:
          "done"     — green num, light card bg
          "current"  — dark card (#0E1216), green num, white text
          "upcoming" — dark num, light card, neutral arrow
        """
        slide = self._new_slide()
        _set_slide_bg(slide, BT.WHITE_HEX)
        _footer(slide)

        if label:
            _txb(slide, label, l=ML, t=Mm(17), w=Mm(106), h=Mm(5),
                 sz=7, bold=True, color=BT.PRIMARY_500_HEX)
        _txb(slide, title, l=ML, t=Mm(22.6), w=Mm(133), h=Mm(18),
             sz=36, bold=True, color=BT.NEUTRAL_900_HEX)
        if description:
            _txb(slide, description, l=Mm(208), t=Mm(22.6), w=Mm(109), h=Mm(16),
                 sz=9, color=BT.NEUTRAL_700_HEX, wrap=True)

        _rect(slide, l=ML, t=Mm(43.6), w=SLIDE_W - 2*ML, h=Mm(0.3),
              fill=BT.NEUTRAL_200_HEX)

        CARD_W  = Mm(93.5)
        CARD_H  = Mm(56.4)
        COL_GAP = Mm(7.9)
        ROW_GAP = Mm(5.3)
        GRID_Y  = Mm(55)

        for i, ch in enumerate(chapters[:6]):
            col   = i % 3
            row   = i // 3
            cx    = ML + col * (CARD_W + COL_GAP)
            cy    = GRID_Y + row * (CARD_H + ROW_GAP)
            state = ch.get("state", "done")

            if state == "current":
                card_bg = BT.NEUTRAL_900_HEX
                num_c   = BT.PRIMARY_500_HEX
                ttl_c   = BT.WHITE_HEX
                sub_c   = BT.NEUTRAL_200_HEX
                arr_bg  = BT.PRIMARY_500_HEX
                arr_c   = BT.WHITE_HEX
                line_c  = BT.PRIMARY_500_HEX
            elif state == "upcoming":
                card_bg = BT.BG_PAGE_HEX
                num_c   = BT.NEUTRAL_900_HEX
                ttl_c   = BT.NEUTRAL_900_HEX
                sub_c   = BT.NEUTRAL_700_HEX
                arr_bg  = BT.NEUTRAL_100_HEX
                arr_c   = BT.NEUTRAL_700_HEX
                line_c  = BT.NEUTRAL_400_HEX
            else:  # done
                card_bg = BT.BG_PAGE_HEX
                num_c   = BT.PRIMARY_500_HEX
                ttl_c   = BT.NEUTRAL_900_HEX
                sub_c   = BT.NEUTRAL_700_HEX
                arr_bg  = BT.PRIMARY_100_HEX
                arr_c   = BT.SUCCESS_HEX
                line_c  = BT.PRIMARY_500_HEX

            _rect(slide, l=cx, t=cy, w=CARD_W, h=CARD_H, fill=card_bg)

            _txb(slide, ch.get("num", ""),
                 l=cx + Mm(5.6), t=cy + Mm(5.6), w=Mm(24), h=Mm(18),
                 sz=44, bold=True, color=num_c)

            arr_x = cx + Mm(81.4)
            arr_y = cy + Mm(14.6)
            _rect(slide, l=arr_x, t=arr_y, w=Mm(6.3), h=Mm(6.3),
                  fill=arr_bg, rounded=True)
            _txb(slide, "→", l=arr_x, t=arr_y + Mm(0.5), w=Mm(6.3), h=Mm(5),
                 sz=8, bold=True, color=arr_c, align=PP_ALIGN.CENTER)

            _txb(slide, ch.get("title", ""),
                 l=cx + Mm(5.6), t=cy + Mm(27.5), w=CARD_W - Mm(10), h=Mm(9),
                 sz=15, bold=True, color=ttl_c)

            _rect(slide, l=cx + Mm(5.6), t=cy + Mm(37.5), w=Mm(20), h=Mm(0.5),
                  fill=line_c)

            sub = ch.get("subtitle", "")
            if sub:
                _txb(slide, sub,
                     l=cx + Mm(5.6), t=cy + Mm(40), w=CARD_W - Mm(10), h=Mm(12),
                     sz=8, color=sub_c, wrap=True)

        return slide

    # ── Dark Cover with Background Image (Slide 4 style) ──────────────────────

    def add_cover_image(self,
                        title: str,
                        subtitle: str = "",
                        tagline: str = "",
                        date_label: str = "",
                        meta_pairs: Optional[List[Tuple[str, str]]] = None,
                        bg_image_path: Optional[str] = None):
        """
        Full-bleed dark cover with optional background image overlay.
        Mirrors the photo-cover variant in the master template.

        meta_pairs: up to 3 (value, label) pairs shown as bottom info strip.
                    e.g. [("2026·03","PRESENTATION VERSION"),
                          ("www.example.com","OFFICIAL WEBSITE"),
                          ("Enterprise AI","CORE FOCUS")]
        bg_image_path: optional path to a full-slide background photo.
        """
        slide = self._new_slide()
        _set_slide_bg(slide, BT.NEUTRAL_900_HEX)

        if bg_image_path and os.path.exists(bg_image_path):
            slide.shapes.add_picture(
                bg_image_path, 0, 0, int(SLIDE_W), int(SLIDE_H))

        # Stacked reverse logo — top left
        _add_logo_stacked(slide, reverse=True, l_mm=25, t_mm=13, h_mm=26)

        if tagline:
            _txb(slide, tagline,
                 l=SLIDE_W - Mm(100), t=Mm(13), w=Mm(96), h=Mm(5),
                 sz=7, color=BT.NEUTRAL_400_HEX, align=PP_ALIGN.RIGHT)

        # Section tag
        if subtitle:
            _txb(slide, subtitle,
                 l=Mm(28), t=Mm(56), w=Mm(200), h=Mm(6),
                 sz=8, color=BT.PRIMARY_500_HEX)

        # Main title (white, 48pt, multiline)
        _txb(slide, title,
             l=Mm(28), t=Mm(65), w=Mm(256), h=Mm(50),
             sz=48, bold=True, color=BT.WHITE_HEX, ls_pt=56)

        if date_label:
            _txb(slide, date_label,
                 l=Mm(28), t=Mm(124), w=Mm(200), h=Mm(8),
                 sz=14, color=BT.NEUTRAL_200_HEX)

        # Bottom meta strip
        if meta_pairs:
            strip_y = Mm(164)
            col_w   = (SLIDE_W - Mm(56)) // len(meta_pairs)
            for j, (val, lbl) in enumerate(meta_pairs[:3]):
                bx = Mm(28) + j * col_w
                _txb(slide, val,
                     l=bx, t=strip_y, w=col_w - Mm(4), h=Mm(8),
                     sz=11, color=BT.WHITE_HEX)
                _txb(slide, lbl,
                     l=bx, t=strip_y + Mm(8), w=col_w - Mm(4), h=Mm(6),
                     sz=9, color=BT.NEUTRAL_400_HEX)

        # Bottom bar
        _rect(slide, l=0, t=SLIDE_H - Mm(3.5), w=SLIDE_W, h=Mm(3.5),
              fill=BT.PRIMARY_500_HEX)
        return slide

    # ── Rich Chapter Divider (Slide 6 style) ──────────────────────────────────

    def add_divider_rich(self,
                         chapter_num: str,
                         chapter_title: str,
                         subtitle: str = "",
                         chapter_items: Optional[List[Tuple[str, str]]] = None,
                         current_item: int = 0,
                         bg_image_path: Optional[str] = None):
        """
        Dark divider with oversized decorative chapter number and right sidebar.

        chapter_num:   "01"  (shown at 260pt as teal background accent)
        chapter_title: main title (60pt, white)
        subtitle:      optional subtitle below the title
        chapter_items: [("01","关于我们"), ("02","公司简介"), ...]
                       shown in the right sidebar; current_item is highlighted white
        current_item:  0-based index of the active entry in chapter_items
        bg_image_path: optional full-slide background image overlay
        """
        slide = self._new_slide()
        _set_slide_bg(slide, BT.NEUTRAL_900_HEX)

        if bg_image_path and os.path.exists(bg_image_path):
            slide.shapes.add_picture(
                bg_image_path, 0, 0, int(SLIDE_W), int(SLIDE_H))

        # Oversized decorative number (behind chapter title, serves as texture)
        _txb(slide, chapter_num,
             l=Mm(28), t=Mm(38), w=Mm(160), h=Mm(83),
             sz=260, bold=True, color=BT.PRIMARY_500_HEX)

        # "CHAPTER · XX" small label
        _txb(slide, f"CHAPTER · {chapter_num}",
             l=Mm(28), t=Mm(84.7), w=Mm(106), h=Mm(6),
             sz=9, bold=True, color=BT.PRIMARY_500_HEX)

        # Chapter title
        _txb(slide, chapter_title,
             l=Mm(28), t=Mm(93), w=Mm(212), h=Mm(32),
             sz=60, bold=True, color=BT.WHITE_HEX, ls_pt=68)

        if subtitle:
            _txb(slide, subtitle,
                 l=Mm(28.9), t=Mm(124), w=Mm(212), h=Mm(10),
                 sz=14, color=BT.NEUTRAL_200_HEX)

        # Thin accent line
        _rect(slide, l=Mm(28), t=Mm(140), w=Mm(21), h=Mm(0.7),
              fill=BT.WHITE_HEX)

        # Right sidebar — chapter list
        SIDEBAR_X = Mm(232.8)
        sidebar_w = SLIDE_W - SIDEBAR_X - MR

        if chapter_items:
            _txb(slide, "IN THIS CHAPTER",
                 l=SIDEBAR_X, t=Mm(93), w=sidebar_w, h=Mm(7),
                 sz=7, color=BT.NEUTRAL_400_HEX)
            _rect(slide, l=SIDEBAR_X, t=Mm(101), w=sidebar_w, h=Mm(0.2),
                  fill=BT.WHITE_HEX)
            for j, (num, name) in enumerate(chapter_items):
                item_y  = Mm(103.5) + j * Mm(6.4)
                is_curr = (j == current_item)
                _txb(slide, f"{num}  {name}",
                     l=SIDEBAR_X, t=item_y, w=sidebar_w, h=Mm(6),
                     sz=9,
                     color=BT.WHITE_HEX if is_curr else BT.NEUTRAL_200_HEX)

        _footer(slide)
        return slide

    # ── About / Intro Slide (Slide 13/14 style) ───────────────────────────────

    def add_about_slide(self,
                        title: str,
                        body_text: str,
                        label: str = "",
                        callout_items: Optional[List[Dict[str, str]]] = None,
                        right_panel: Optional[Dict[str, Any]] = None):
        """
        Left text + right dark panel layout (company intro / team page).

        body_text:     main paragraph on the left
        label:         small section label above title (e.g. "01 · ABOUT US")
        callout_items: compact info boxes below body_text.
                       [{"label":"OFFICIAL WEBSITE","value":"https://..."}]
        right_panel:   dark panel on the right with accent items.
                       {"title_label":"OUR VISION","items":[{"label":"...","text":"...","accent":"#3EC99E"}]}
        """
        slide = self._new_slide()
        _set_slide_bg(slide, BT.WHITE_HEX)
        _footer(slide)

        # Thin top divider
        _rect(slide, l=ML, t=Mm(40.9), w=SLIDE_W - 2*ML, h=Mm(0.2),
              fill=BT.NEUTRAL_200_HEX)

        # Section label
        if label:
            _txb(slide, label, l=ML, t=Mm(14.1), w=Mm(100), h=Mm(4),
                 sz=7, bold=True, color=BT.PRIMARY_500_HEX)

        # Title
        _txb(slide, title, l=ML, t=Mm(19.8), w=Mm(257), h=Mm(14),
             sz=28, bold=True, color=BT.NEUTRAL_900_HEX)

        LEFT_W = Mm(138)

        # Body text
        _txb(slide, body_text,
             l=ML, t=Mm(58.6), w=LEFT_W, h=Mm(30),
             sz=11, color=BT.NEUTRAL_700_HEX, ls_pt=17, wrap=True)

        # Callout info boxes
        if callout_items:
            BOX_H = Mm(13.8)
            for k, item in enumerate(callout_items[:2]):
                bx = ML + k * (LEFT_W // 2 + Mm(3))
                by = Mm(91.1)
                _rect(slide, l=bx, t=by, w=LEFT_W // 2 - Mm(3), h=BOX_H,
                      fill=BT.BG_PAGE_HEX)
                _rect(slide, l=bx, t=by, w=Mm(0.5), h=BOX_H,
                      fill=BT.PRIMARY_500_HEX)
                _txb(slide, item.get("label", ""),
                     l=bx + Mm(4), t=by + Mm(2.8), w=Mm(60), h=Mm(4),
                     sz=6, color=BT.NEUTRAL_400_HEX)
                _txb(slide, item.get("value", ""),
                     l=bx + Mm(4), t=by + Mm(7.5), w=Mm(60), h=Mm(5),
                     sz=11, color=BT.NEUTRAL_900_HEX)

        # Right dark panel
        PANEL_X = Mm(172.8)
        PANEL_W = SLIDE_W - PANEL_X - MR
        _rect(slide, l=PANEL_X, t=Mm(49.4), w=PANEL_W, h=Mm(53),
              fill=BT.NEUTRAL_900_HEX)

        if right_panel:
            PANEL_Y = Mm(49.4)
            items   = right_panel.get("items", [])
            item_h  = Mm(15.5)
            for k, it in enumerate(items[:3]):
                iy     = PANEL_Y + k * item_h + Mm(8)
                accent = it.get("accent", BT.PRIMARY_500_HEX)
                # Accent dot
                _rect(slide, l=PANEL_X + Mm(9), t=iy,
                      w=Mm(8.5), h=Mm(8.5), fill=accent, rounded=True)
                # Item label
                _txb(slide, it.get("label", ""),
                     l=PANEL_X + Mm(20), t=iy, w=PANEL_W - Mm(22), h=Mm(5),
                     sz=7, bold=True, color=accent)
                # Item text
                _txb(slide, it.get("text", ""),
                     l=PANEL_X + Mm(20), t=iy + Mm(5.5), w=PANEL_W - Mm(22), h=Mm(8),
                     sz=14, bold=True, color=BT.WHITE_HEX, ls_pt=18)

        # Bottom cards row (brand tone / core focus chips)
        return slide

    # ── Feature Module Grid (Slide 15 style) ──────────────────────────────────

    def add_module_grid(self,
                        title: str,
                        modules: List[Dict],
                        subtitle: str = "",
                        label: str = ""):
        """
        4-per-row feature module grid, up to 8 modules.

        modules: [{
            "num":      "01",          # sequence number shown top-right
            "title":    "智能检查",
            "en":       "SMART INSPECTION",
            "bullets":  ["AI 一键创建检查表单", "..."],
            "accent":   "#4B9E31",     # EN label color  (optional, defaults cycle)
            "icon_bg":  "#EAFAF5",     # icon square bg  (optional, defaults cycle)
            "featured": False,         # True → dark bg, white text (flagship)
        }]
        """
        _ICON_DEFAULTS = [
            "#EAFAF5", "#FFF1DF", "#F2F3F5", "#3EC99E",
            "#F2F3DC", "#F0E8FF", "#E0F7FA", BT.SECONDARY_500_HEX,
        ]
        _ACCENT_DEFAULTS = [
            BT.SUCCESS_HEX, "#FFB928", BT.NEUTRAL_700_HEX, BT.PRIMARY_500_HEX,
            BT.SUCCESS_HEX, "#8255E1", "#3CC5CF",           BT.SECONDARY_500_HEX,
        ]

        slide = self._new_slide()
        _set_slide_bg(slide, BT.WHITE_HEX)
        _header(slide, title, subtitle=subtitle, label=label)
        _footer(slide)

        COLS    = 4
        CELL_W  = Mm(74.1)
        CELL_H  = Mm(60)
        COL_GAP = Mm(3.5)
        ROW_GAP = Mm(3.5)
        GRID_Y  = CONTENT_Y + Mm(3)

        for i, mod in enumerate(modules[:8]):
            col      = i % COLS
            row      = i // COLS
            cx       = ML + col * (CELL_W + COL_GAP)
            cy       = GRID_Y + row * (CELL_H + ROW_GAP)
            featured = mod.get("featured", False)
            accent   = mod.get("accent",  _ACCENT_DEFAULTS[i % len(_ACCENT_DEFAULTS)])
            icon_bg  = mod.get("icon_bg", _ICON_DEFAULTS[i % len(_ICON_DEFAULTS)])
            cell_bg  = BT.NEUTRAL_900_HEX if featured else BT.WHITE_HEX
            ttl_c    = BT.WHITE_HEX if featured else BT.NEUTRAL_900_HEX
            bul_c    = BT.NEUTRAL_200_HEX if featured else BT.NEUTRAL_700_HEX
            num_c    = accent if featured else BT.NEUTRAL_400_HEX

            _rect(slide, l=cx, t=cy, w=CELL_W, h=CELL_H,
                  fill=cell_bg,
                  line=None if featured else BT.NEUTRAL_200_HEX)

            # Icon background square
            _rect(slide, l=cx + Mm(4.9), t=cy + Mm(5), w=Mm(8.5), h=Mm(8.5),
                  fill=icon_bg)

            # Sequence number — top right of cell
            num_str = mod.get("num", f"{i+1:02d}")
            _txb(slide, num_str,
                 l=cx + CELL_W - Mm(9), t=cy + Mm(5), w=Mm(8), h=Mm(5),
                 sz=9, bold=True, color=num_c, align=PP_ALIGN.RIGHT)

            # Module title
            _txb(slide, mod.get("title", ""),
                 l=cx + Mm(4.9), t=cy + Mm(16.6), w=CELL_W - Mm(9), h=Mm(7),
                 sz=11, bold=True, color=ttl_c)

            # EN subtitle
            en = mod.get("en", "")
            if en:
                _txb(slide, en,
                     l=cx + Mm(4.9), t=cy + Mm(24), w=CELL_W - Mm(9), h=Mm(4),
                     sz=6, bold=True, color=accent)

            # Bullet points
            bullets = mod.get("bullets", [])
            if bullets:
                body = "\n".join(f"· {b}" for b in bullets)
                _txb(slide, body,
                     l=cx + Mm(4.9), t=cy + Mm(29.5),
                     w=CELL_W - Mm(9), h=CELL_H - Mm(33),
                     sz=7, color=bul_c, ls_pt=11, wrap=True)

        return slide

    # ── Closing Slide ─────────────────────────────────────────────────────────

    def add_closing(self, message: str = "谢谢", contact: str = ""):
        """Closing slide: dark background, gradient message, logo centered."""
        slide = self._new_slide()
        _set_slide_bg(slide, BT.NEUTRAL_900_HEX)

        # Stacked logo centered top
        _add_logo_stacked(slide, reverse=True,
                          l_mm=int((SLIDE_W / Mm(1) - 24) / 2),
                          t_mm=32, h_mm=24)

        # Gradient message
        _txb_gradient(slide, message,
                      l=ML, t=Mm(76),
                      w=SLIDE_W - 2 * ML, h=Mm(40),
                      sz=50, bold=True, align=PP_ALIGN.CENTER,
                      stops=TITLE_GRADIENT)

        if contact:
            _txb(slide, contact,
                 l=ML, t=Mm(122),
                 w=SLIDE_W - 2 * ML, h=Mm(12),
                 sz=14, color=BT.NEUTRAL_400_HEX, align=PP_ALIGN.CENTER)

        _rect(slide, l=0, t=SLIDE_H - Mm(3.5), w=SLIDE_W, h=Mm(3.5),
              fill=BT.PRIMARY_500_HEX)
        return slide

    # ── Save ─────────────────────────────────────────────────────────────────

    def save(self, output_path: str):
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        self._prs.save(output_path)
        print(f"✓ Saved: {output_path}")

    @property
    def presentation(self):
        return self._prs
