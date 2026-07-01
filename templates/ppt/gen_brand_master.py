#!/usr/bin/env python3
"""
gen_brand_master.py

生成 ArktechX 品牌 PPT 母版文件 → templates/ppt/brand_master.pptx

包含 13 张示范幻灯片（1 目录索引 + 12 版式），注入品牌主题色。
后续项目用 BrandPptx() 时会自动加载本文件作为基础模板，
继承主题色和布局配置，保证视觉一致性。

运行方式：
    cd "brand 3"
    python3 templates/ppt/gen_brand_master.py
"""

import sys, os, zipfile, io
from lxml import etree

_BRAND     = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
_TEMPLATES = os.path.abspath(os.path.dirname(__file__))
sys.path.insert(0, _BRAND)

OUTPUT = os.path.join(_TEMPLATES, "brand_master.pptx")

from scripts.pptx_builder import (
    BrandPptx, _txb, _txb_gradient, _rect, _card, _header, _footer,
    _set_slide_bg, _add_logo_h, _add_logo_stacked,
    SLIDE_W, SLIDE_H, ML, MR, CW, HEADER_H, FOOTER_H, CONTENT_Y, CONTENT_H,
    C2_W, C2_GAP, C3_W, C3_GAP, TITLE_GRADIENT
)
import scripts.brand_tokens as BT
from pptx.util import Mm, Pt
from pptx.enum.text import PP_ALIGN


# ─── Brand Theme XML ──────────────────────────────────────────────────────────
#   Replaces the default Office theme inside the PPTX zip.
#   Ensures brand palette is available as theme accent colors.

BRAND_THEME_XML = '''\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="ArktechX">
  <a:themeElements>
    <a:clrScheme name="ArktechX Colors">
      <a:dk1><a:srgbClr val="0E1216"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="3D444A"/></a:dk2>
      <a:lt2><a:srgbClr val="F2F3F5"/></a:lt2>
      <a:accent1><a:srgbClr val="3EC99E"/></a:accent1>
      <a:accent2><a:srgbClr val="C8E13C"/></a:accent2>
      <a:accent3><a:srgbClr val="319E7C"/></a:accent3>
      <a:accent4><a:srgbClr val="8A9199"/></a:accent4>
      <a:accent5><a:srgbClr val="D0D5DD"/></a:accent5>
      <a:accent6><a:srgbClr val="EAFAF5"/></a:accent6>
      <a:hlink><a:srgbClr val="3EC99E"/></a:hlink>
      <a:folHlink><a:srgbClr val="319E7C"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="ArktechX Fonts">
      <a:majorFont>
        <a:latin typeface="Arial"/>
        <a:ea typeface="PingFang SC"/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Arial"/>
        <a:ea typeface="PingFang SC"/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="ArktechX">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="50000"/>
                <a:satMod val="300000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:tint val="15000"/>
                <a:satMod val="350000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="1"/>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:shade val="51000"/>
                <a:satMod val="130000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="94000"/>
                <a:satMod val="135000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="0"/>
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr">
              <a:shade val="95000"/>
              <a:satMod val="105000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
        <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
        <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="20000" dist="12000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000"><a:alpha val="20000"/></a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill>
          <a:schemeClr val="phClr">
            <a:tint val="95000"/>
            <a:satMod val="170000"/>
          </a:schemeClr>
        </a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="93000"/>
                <a:satMod val="150000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="63000"/>
                <a:satMod val="120000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="0"/>
        </a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>'''


def inject_brand_theme(pptx_path: str):
    """Replace all theme XML files inside the PPTX with the ArktechX brand theme."""
    with zipfile.ZipFile(pptx_path, 'r') as zin:
        names  = zin.namelist()
        files  = {n: zin.read(n) for n in names}

    theme_files = [k for k in files if 'theme' in k.lower() and k.endswith('.xml')]
    for key in theme_files:
        files[key] = BRAND_THEME_XML.encode('utf-8')

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)

    with open(pptx_path, 'wb') as f:
        f.write(buf.getvalue())


# ─── Slide builders ───────────────────────────────────────────────────────────

def slide_00_index(prs: BrandPptx):
    """目录索引页 — 供 AI 快速识别版式编号。"""
    slide = prs._new_slide()
    _set_slide_bg(slide, BT.WHITE_HEX)
    _rect(slide, l=0, t=0, w=SLIDE_W, h=Mm(2.5), fill=BT.PRIMARY_500_HEX)

    _txb(slide, "品牌 PPT 母版版式目录",
         l=ML, t=Mm(8), w=CW, h=Mm(22),
         sz=28, bold=True, align=PP_ALIGN.CENTER,
         color=BT.NEUTRAL_900_HEX)

    layouts = [
        ("01", "封面页 — 深色",       "Cover Dark"),
        ("02", "封面页 — 浅色",       "Cover Light  (VS Style)"),
        ("03", "章节分隔",            "Section Divider"),
        ("04", "标准内容",            "Title + Content"),
        ("05", "两列内容",            "Two Columns"),
        ("06", "三列卡片",            "Three Cards"),
        ("07", "大数据指标",          "Big Stats (2×2)"),
        ("08", "水平时间轴",          "Timeline"),
        ("09", "引用证言",            "Quote"),
        ("10", "表格页",             "Table"),
        ("11", "左文右图",            "Text + Image"),
        ("12", "结尾致谢",            "Closing"),
    ]

    cols = 3
    item_w = CW // cols
    item_h = Mm(18)
    start_y = Mm(38)
    gap_y = Mm(2)

    for i, (num, cn, en) in enumerate(layouts):
        col = i % cols
        row = i // cols
        x   = ML + col * item_w
        y   = start_y + row * (item_h + gap_y)

        # Number badge
        _rect(slide, l=x, t=y, w=Mm(12), h=Mm(10),
              fill=BT.PRIMARY_500_HEX, rounded=True)
        _txb(slide, num,
             l=x + Mm(1), t=y + Mm(1), w=Mm(10), h=Mm(9),
             sz=10, bold=True, color=BT.WHITE_HEX, align=PP_ALIGN.CENTER)

        _txb(slide, cn,
             l=x + Mm(14), t=y, w=item_w - Mm(16), h=Mm(11),
             sz=12, bold=True, color=BT.NEUTRAL_900_HEX)
        _txb(slide, en,
             l=x + Mm(14), t=y + Mm(10), w=item_w - Mm(16), h=Mm(8),
             sz=9, color=BT.NEUTRAL_400_HEX)

    _footer(slide, layout_label="00 · Index")


def slide_01_cover_dark(prs: BrandPptx):
    """01 — 封面页（深色）"""
    prs.add_cover(
        title="能碳智能管理平台",
        subtitle="Industrial Carbon & Energy Management Platform",
        tagline="上海未来方舟智能科技有限公司  ·  ArktechX"
    )


def slide_02_cover_light(prs: BrandPptx):
    """02 — 封面页（浅色 VS 风格）"""
    prs.add_cover_light(
        title="能碳智能管理平台",
        subtitle="Industrial Carbon & Energy Management Platform",
        date_or_meta="上海未来方舟智能科技有限公司  ·  2025",
    )


def slide_03_section(prs: BrandPptx):
    """03 — 章节分隔"""
    prs.add_divider("第一章：产品概述", chapter_num="CHAPTER ONE")


def slide_04_body(prs: BrandPptx):
    """04 — 标准内容页"""
    prs.add_body_slide(
        title="平台核心能力",
        subtitle="Platform Core Capabilities",
        bullets=[
            "Scope 1/2/3 排放源自动识别与核算，符合 ISO 14064 & GHG Protocol",
            "多工厂、多能源介质数据实时接入，支持工业协议 OPC-UA / Modbus",
            "碳资产台账管理：CCER 挂牌、注销、交易全流程追踪",
            "报告中心：一键生成 GHG 声明、核查支撑材料、绿色工厂评价报告",
            "驾驶舱看板：集团/工厂/设备三级穿透，KPI 实时预警",
        ],
        slide_label="04 · Title + Content"
    )


def slide_05_two_cols(prs: BrandPptx):
    """05 — 两列内容页"""
    prs.add_two_col_slide(
        title="行业挑战与平台应对",
        subtitle="Challenge vs. Solution",
        left_title="企业面临的核心痛点",
        right_title="ArkSus 平台解法",
        left_content=(
            "• 数据孤岛：能耗数据分散在 ERP / MES / SCADA，人工汇总误差高\n"
            "• 核算滞后：月度对账周期长，无法实时掌握排放趋势\n"
            "• 合规压力：核查报告格式多变，第三方审计沟通成本高\n"
            "• 范围三盲区：供应链碳数据收集缺乏标准工具"
        ),
        right_content=(
            "• 统一数据底座：IoT + API 双通道接入，自动清洗标准化\n"
            "• 实时核算引擎：基于 GHG Protocol 排放因子库，T+0 出数\n"
            "• 智能报告中心：预置核查机构模板，一键生成可追溯报告\n"
            "• Scope 3 模块：供应商碳数据采集问卷 + LCA 集成"
        )
    )


def slide_06_three_cards(prs: BrandPptx):
    """06 — 三列卡片"""
    prs.add_three_cards(
        title="三大核心价值",
        subtitle="Three Core Value Propositions",
        cards=[
            {
                "tag": "ACCURACY",
                "title": "数据可信",
                "body": "基于国际标准（ISO 14064、GHG Protocol）自动核算，全链路留痕，支持第三方核查。"
            },
            {
                "tag": "EFFICIENCY",
                "title": "流程提效",
                "body": "碳盘查周期从 3 个月缩短至 3 天，报告生成自动化率 ≥ 90%，合规成本大幅下降。"
            },
            {
                "tag": "INSIGHT",
                "title": "决策支持",
                "body": "集团→工厂→设备三级穿透看板，实时识别高排放环节，为减碳投资提供数据支撑。"
            }
        ]
    )


def slide_07_big_stats(prs: BrandPptx):
    """07 — 大数据指标"""
    prs.add_big_stats(
        title="平台关键指标",
        subtitle="Key Performance Metrics",
        stats=[
            ("98%",  "核算准确率",  "基于国际标准自动核算，误差 < 2%"),
            ("3天",  "盘查周期",   "传统方式 3 个月缩短至 3 天"),
            ("30+",  "接入工厂",   "支持多能源介质、多标准并行"),
            ("100%", "报告合规率", "预置核查机构标准模板"),
        ]
    )


def slide_08_timeline(prs: BrandPptx):
    """08 — 水平时间轴"""
    prs.add_timeline(
        title="双碳路线图",
        subtitle="Carbon Neutrality Roadmap",
        milestones=[
            ("2024 Q2", "基础建设",   "完成能耗数据接入\nScope 1+2 核算上线"),
            ("2024 Q4", "合规认证",   "完成首次第三方核查\n绿色工厂评价申报"),
            ("2025 Q2", "范围扩展",   "Scope 3 供应链模块\n上线运行"),
            ("2026",    "碳资产运营", "CCER 项目开发\n碳市场交易"),
            ("2030",    "碳中和目标", "集团净零排放\n对标 SBTi 1.5°C"),
        ]
    )


def slide_09_quote(prs: BrandPptx):
    """09 — 引用证言"""
    prs.add_quote(
        quote_text="ArkSus 平台让我们第一次真正实现了碳数据的实时可见，\n盘查时间从三个月压缩到三天，核查报告自动生成，\n为集团碳中和目标落地提供了坚实的数字底座。",
        author="能源总监  ·  某汽车零部件集团",
        role="Global Energy Director, Automotive Tier-1 Supplier",
        subtitle="客户案例 / Client Testimonial"
    )


def slide_10_table(prs: BrandPptx):
    """10 — 表格页"""
    prs.add_table_slide(
        title="Scope 1/2/3 排放汇总",
        subtitle="GHG Inventory Summary — 2024",
        headers=["排放范围", "类别", "排放量 (tCO₂e)", "占比", "同比变化"],
        rows=[
            ["Scope 1", "直接排放（燃烧）",   "12,450",  "28.3%", "▼ 3.2%"],
            ["Scope 1", "直接排放（工艺）",    "3,200",   "7.3%",  "▼ 1.1%"],
            ["Scope 2", "外购电力（市网）",    "18,900",  "43.0%", "▼ 8.5%"],
            ["Scope 2", "外购热力",           "2,100",   "4.8%",  "▼ 2.0%"],
            ["Scope 3", "上游原材料运输",      "4,360",   "9.9%",  "▲ 0.6%"],
            ["Scope 3", "员工通勤",           "2,980",   "6.8%",  "— 0.0%"],
        ],
        note="数据来源：ArkSus 平台 2024 年度碳盘查报告；Scope 3 仅含重大类别"
    )


def slide_11_text_image(prs: BrandPptx):
    """11 — 左文右图"""
    prs.add_text_image_right(
        title="驾驶舱看板",
        subtitle="Real-Time Carbon & Energy Dashboard",
        body_text=(
            "集团→工厂→设备三级穿透\n\n"
            "• 实时能耗与碳排放趋势图\n"
            "• KPI 红绿灯预警（超标自动推送）\n"
            "• 月度/季度/年度多维对比\n"
            "• 支持大屏展示（1920×1080）\n\n"
            "右侧图片区域放置产品截图"
        )
    )


def slide_12_closing(prs: BrandPptx):
    """12 — 结尾致谢"""
    prs.add_closing(
        message="共建绿色制造未来",
        contact="contact@futureark.ai  ·  www.futureark.ai  ·  +86 021-xxxx-xxxx"
    )


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    print("⚙  Generating brand_master.pptx …")

    prs = BrandPptx(use_template=False)   # fresh start — this IS the template

    # Build all 13 slides
    slide_00_index(prs)
    slide_01_cover_dark(prs)
    slide_02_cover_light(prs)
    slide_03_section(prs)
    slide_04_body(prs)
    slide_05_two_cols(prs)
    slide_06_three_cards(prs)
    slide_07_big_stats(prs)
    slide_08_timeline(prs)
    slide_09_quote(prs)
    slide_10_table(prs)
    slide_11_text_image(prs)
    slide_12_closing(prs)

    # Save
    prs.save(OUTPUT)
    print(f"   Saved to: {OUTPUT}")

    # Inject ArktechX brand theme colors into the PPTX
    print("⚙  Injecting brand theme …")
    inject_brand_theme(OUTPUT)
    print("✓  brand_master.pptx ready.")
    print()
    print("Usage in project scripts:")
    print("    from scripts.pptx_builder import BrandPptx")
    print("    prs = BrandPptx()   # auto-loads brand_master.pptx as base template")


if __name__ == "__main__":
    main()
