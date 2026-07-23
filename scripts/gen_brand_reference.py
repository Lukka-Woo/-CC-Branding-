"""
生成 reference/brand-values.md —— 品牌数值参考表（唯一数据源：scripts/brand_tokens.py → tokens.json）

CLAUDE.md / brand_system.md 中涉及具体色值 / 字体名 / 页面数值的规则，
一律指向本文件而不重复硬编码，避免文档与 tokens.json 实际值出现漂移。

用途说明文字（"标题、强调、Logo、表格 Header" 这类展示层描述）手工维护于本脚本，
与数值本身（来自 BT）解耦——改数值改 tokens.json，改用途说明改本脚本，互不影响。

运行方式（改完 tokens.json 后执行一次）：
    python3 scripts/gen_brand_reference.py
"""
import sys, os

_BRAND = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, _BRAND)

import scripts.brand_tokens as BT

_OUT = os.path.join(_BRAND, 'reference', 'brand-values.md')

_COLOR_USAGE = [
    ('主绿',    BT.PRIMARY_500_HEX,   '标题、强调、Logo、表格 Header'),
    ('主绿深',  BT.PRIMARY_600_HEX,   'Hover / 深色变体'),
    ('主绿浅',  BT.PRIMARY_100_HEX,   '卡片背景、交替行'),
    ('辅色',    BT.SECONDARY_500_HEX, '次级强调、Badge'),
    ('中性900', BT.NEUTRAL_900_HEX,   '正文标题'),
    ('中性700', BT.NEUTRAL_700_HEX,   '正文文字'),
    ('中性400', BT.NEUTRAL_400_HEX,   '说明文字、Footer'),
    ('中性200', BT.NEUTRAL_200_HEX,   '边框、分隔线'),
    ('中性100', BT.NEUTRAL_100_HEX,   '斑马纹、Note 背景'),
    ('白色',    BT.WHITE_HEX,         '页面背景、反白文字'),
]

_CARD_SEMANTICS = [
    (BT.PRIMARY_100_HEX,   'PRIMARY_100',    BT.PRIMARY_500_HEX,   'PRIMARY_500',          '标准/主要特性',             '★★★'),
    (BT.NEUTRAL_100_HEX,   'NEUTRAL_100',    BT.SUCCESS_HEX,       'SUCCESS',              '安全/补充/次要',            '★★★'),
    (BT.SECONDARY_100_HEX, 'SECONDARY_100',  BT.SECONDARY_500_HEX, 'SECONDARY_500',        '创新/机遇',                '★★★'),
    (BT.CARD_ORANGE_BG,    'CARD_ORANGE_BG', BT.WARNING_HEX,       'WARNING',              '高风险/注意',              '★★★'),
    (BT.CARD_TEAL_BG,      'CARD_TEAL_BG',   BT.TEAL_HEX,          'TEAL',                 '扩展/生态',                '★★'),
    (BT.CARD_PURPLE_BG,    'CARD_PURPLE_BG', BT.PURPLE_HEX,        'PURPLE',               '战略/特殊',                '★★'),
    (BT.NEUTRAL_900_HEX,   'NEUTRAL_900',    BT.SECONDARY_500_HEX, 'SECONDARY_500 + 白字',  '突出/亮点（非危险）',        '★★★ 深色卡'),
    (BT.CARD_DANGER_BG,    'CARD_DANGER_BG', BT.DANGER_HEX,        'DANGER',               '危险/风险（确实危险才用）',   '★ 慎用'),
]


def _fmt_color_table():
    lines = ['| 角色 | 色值 | 用途 |', '|---|---|---|']
    for name, hex_val, usage in _COLOR_USAGE:
        lines.append(f'| {name} | `{hex_val}` | {usage} |')
    return '\n'.join(lines)


def _fmt_card_table():
    lines = ['| 底色 | Accent 色 | 语义 | 优先级 |', '|---|---|---|---|']
    for bg, bg_name, ac, ac_name, sem, star in _CARD_SEMANTICS:
        lines.append(f'| `{bg}` {bg_name} | `{ac}` {ac_name} | {sem} | {star} |')
    return '\n'.join(lines)


def _fmt_font_table():
    lines = ['| 用途 | 字体族 |', '|---|---|']
    lines.append(f'| 中文（主 / 降级） | {BT.FONT_CN} / {BT.FONT_CN_FB} / {BT.FONT_CN_FB2} / {BT.FONT_CN_WEB}（Web） |')
    lines.append(f'| 英文 | {BT.FONT_EN} |')
    lines.append(f'| 代码 | {BT.FONT_MONO} |')
    return '\n'.join(lines)


def _fmt_docx_page_table():
    lines = ['| 项 | 值 |', '|---|---|']
    lines.append(f'| 上下边距 | {BT.PAGE_MARGIN_TOP_MM}mm |')
    lines.append(f'| 左右边距 | {BT.PAGE_MARGIN_LEFT_MM}mm |')
    lines.append(f'| H1 | {BT.FONT_H1_PT}pt `{BT.FONT_H1_COLOR}` |')
    lines.append(f'| H2 | {BT.FONT_H2_PT}pt `{BT.FONT_H2_COLOR}` |')
    lines.append(f'| H3 | {BT.FONT_H3_PT}pt `{BT.FONT_H3_COLOR}` |')
    return '\n'.join(lines)


def _fmt_radius_table():
    lines = ['| 常量 | 值 | 用途 |', '|---|---|---|']
    lines.append(f'| RADIUS_XS_MM | {BT.RADIUS_XS_MM}mm | 细节圆角（角标、小装饰点） |')
    lines.append(f'| RADIUS_SM_MM | {BT.RADIUS_SM_MM}mm | ★ 默认：所有普通卡片、按钮 |')
    lines.append(f'| RADIUS_MD_MM | {BT.RADIUS_MD_MM}mm | 大卡片、面板 |')
    lines.append(f'| RADIUS_LG_MM | {BT.RADIUS_LG_MM}mm | 外层大容器 / 图片占位区块 |')
    lines.append(f'| RADIUS_PILL_MM | pill | 胶囊标签 / 序号徽章 / 圆形点 |')
    return '\n'.join(lines)


CONTENT = f"""# 品牌数值参考表（自动生成，请勿手动编辑）

> 本文件由 `scripts/gen_brand_reference.py` 从 `scripts/brand_tokens.py`
> （唯一数据源，追溯至 `tokens.json`）生成。
>
> 修改颜色 / 字体 / 页面数值 → 编辑 `tokens.json` → 重新运行：
> `python3 scripts/gen_brand_reference.py`
>
> `CLAUDE.md` / `brand_system.md` 中涉及具体数值的规则一律指向本文件，
> 不在多处重复硬编码同一个值。

## 颜色

{_fmt_color_table()}

所有颜色常量通过 `import scripts.brand_tokens as BT` 使用，禁止在脚本中内联十六进制。

## 字体

{_fmt_font_table()}

PPTX/DOCX 写字体名固定引用 `BT.FONT_CN` / `BT.FONT_EN`；
HTML/PDF 通过 `@font-face` 加载本地文件（`fonts/alibaba-puhuiti/`）。

## DOCX 页面（A4）

{_fmt_docx_page_table()}

## 卡片颜色语义（全系统统一）

适用于 `add_three_cards`、`add_six_cards`、`add_big_stats`、`add_module_grid`。

{_fmt_card_table()}

## 圆角（PPTX / DOCX）

{_fmt_radius_table()}
"""

if __name__ == '__main__':
    os.makedirs(os.path.dirname(_OUT), exist_ok=True)
    with open(_OUT, 'w', encoding='utf-8') as f:
        f.write(CONTENT)
    print(f'✅ Generated: {_OUT}')
