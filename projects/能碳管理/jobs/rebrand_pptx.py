"""
能碳平台介绍 PPT 品牌化脚本
将原始 Tailwind 蓝/橙配色体系替换为 ArkSus® 品牌色系
"""

import sys, os, zipfile, shutil, re, json

_BRAND   = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..'))
_PROJECT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
_DOCS    = os.path.join(_PROJECT, "docs")
os.makedirs(_DOCS, exist_ok=True)

SRC = os.path.join(_PROJECT, "references", "260522-能碳平台介绍-交付.pptx")
DST = os.path.join(_DOCS, "能碳平台介绍-ArkSus品牌版.pptx")

# ── Color Mapping ─────────────────────────────────────────────
# Keys: lowercase hex (no #), Values: uppercase hex (no #)
# Based on tokens.json + brand_system.md
COLOR_MAP = {
    # Dark backgrounds — 3-level hierarchy
    '0b0f1a': '0E1216',  # deepest slide bg → brand n900
    '111827': '152030',  # panel/card bg    → brand dark-elevated-1
    '1f2937': '1E2D40',  # elevated card bg → brand dark-elevated-2

    # Primary accent: sky blue → brand primary 智慧绿
    '38bdf8': '3EC99E',

    # Secondary accent: orange → brand secondary 荧光柠绿
    'f97316': 'C8E13C',

    # Text grays
    '9aa4b8': '8A9199',  # muted text → brand n400
    'f5f7fb': 'F8FAFC',  # near-white  → brand n50

    # Semantic colors
    'ef4444': 'F12D2D',  # red/danger  → brand semantic-danger
    'a78bfa': '8255E1',  # purple      → brand supplement-purple
    'f59e0b': 'F3B021',  # amber       → brand semantic-warning
}

# Build regex patterns (case-insensitive match on hex val attributes)
PATTERNS = []
for old_lower, new_upper in COLOR_MAP.items():
    # Match: val="XXXXXX"  or  lastClr="XXXXXX"
    pat = re.compile(
        r'(?i)((?:val|lastClr)=")(' + old_lower + r')(")' ,
        re.IGNORECASE
    )
    PATTERNS.append((pat, r'\g<1>' + new_upper + r'\3'))


def replace_colors_in_xml(xml_text: str) -> str:
    for pat, repl in PATTERNS:
        xml_text = pat.sub(repl, xml_text)
    return xml_text


def process_pptx(src_path: str, dst_path: str):
    total_replacements = 0

    with zipfile.ZipFile(src_path, 'r') as zin:
        with zipfile.ZipFile(dst_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                # Only process XML-based files (slides, layouts, masters, theme, chart data)
                is_xml = (
                    item.filename.endswith('.xml') or
                    item.filename.endswith('.rels') or
                    item.filename.endswith('.vml')
                )

                if is_xml:
                    try:
                        text = data.decode('utf-8')
                        new_text = replace_colors_in_xml(text)

                        # Count replacements for this file
                        changed = sum(
                            len(pat.findall(text)) for pat, _ in PATTERNS
                        )
                        if changed:
                            total_replacements += changed
                            print(f"  [{changed:3d} hits] {item.filename}")

                        data = new_text.encode('utf-8')
                    except UnicodeDecodeError:
                        pass  # binary file, skip

                zout.writestr(item, data)

    return total_replacements


# ── Main ──────────────────────────────────────────────────────
print("=" * 60)
print("ArkSus® PPT 品牌化工具")
print("=" * 60)
print(f"\nSource : {SRC}")
print(f"Output : {DST}\n")

print("Color mapping:")
for old, new in COLOR_MAP.items():
    print(f"  #{old.upper()} → #{new}")

print("\nProcessing slides...\n")
n = process_pptx(SRC, DST)

if os.path.exists(DST):
    size_kb = os.path.getsize(DST) / 1024
    print(f"\n✓ Done — {n} color replacements across all XML")
    print(f"✓ Output: {DST} ({size_kb:.0f} KB)")
else:
    print("\n✗ Failed — output file not created")
    sys.exit(1)
