"""
One-time repair script: remove 華康雅宋體(P) from brand_master.pptx.

Changes:
  1. slide5.xml  — replace typeface with "Alibaba PuHuiTi 2.0"
  2. slide14.xml — replace typeface with "Alibaba PuHuiTi 2.0"
  3. presentation.xml — remove <p:embeddedFont> block for 華康雅宋體(P) (rId38)
  4. ppt/_rels/presentation.xml.rels — remove rId38 relationship
  5. ppt/fonts/font4.fntdata — drop the font binary file
  6. docProps/app.xml — remove the font name from the HeadingPairs vector
"""
import os, re, zipfile, shutil, sys

_BRAND = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
TEMPLATE = os.path.join(_BRAND, "templates", "ppt", "brand_master.pptx")
BACKUP   = TEMPLATE + ".bak"
TMP      = TEMPLATE + ".tmp"

FONT_TO_REMOVE  = "華康雅宋體(P)"
FONT_REPLACEMENT = "Alibaba PuHuiTi 2.0"
EMBEDDED_RID    = "rId38"
EMBEDDED_FILE   = "ppt/fonts/font4.fntdata"

if not os.path.exists(TEMPLATE):
    print(f"ERROR: {TEMPLATE} not found")
    sys.exit(1)

# Backup
shutil.copy2(TEMPLATE, BACKUP)
print(f"Backup → {BACKUP}")

def fix_slide_xml(content: str) -> str:
    """Replace 華康雅宋體(P) typeface with brand font in slide XML."""
    return content.replace(
        f'typeface="{FONT_TO_REMOVE}"',
        f'typeface="{FONT_REPLACEMENT}"'
    )

def fix_presentation_xml(content: str) -> str:
    """Remove the <p:embeddedFont> block for 華康雅宋體(P)."""
    # Remove the entire <p:embeddedFont>...</p:embeddedFont> that contains the bad font
    pattern = r'<p:embeddedFont><p:font typeface="華康雅宋體\(P\)"[^<]*/>.*?</p:embeddedFont>'
    cleaned = re.sub(pattern, '', content, flags=re.DOTALL)
    if cleaned == content:
        print("  WARNING: embeddedFont block not found in presentation.xml — skipping")
    return cleaned

def fix_rels(content: str) -> str:
    """Remove the rId38 relationship line."""
    pattern = rf'<Relationship[^>]+Id="{EMBEDDED_RID}"[^>]*/>'
    cleaned = re.sub(pattern, '', content)
    if cleaned == content:
        print(f"  WARNING: {EMBEDDED_RID} not found in rels — skipping")
    return cleaned

def fix_app_xml(content: str) -> str:
    """Remove 華康雅宋體(P) entry from the font list vector and decrement its size."""
    # The font appears as <vt:lpstr>華康雅宋體(P)</vt:lpstr>
    cleaned = content.replace(f'<vt:lpstr>{FONT_TO_REMOVE}</vt:lpstr>', '')
    if cleaned == content:
        print("  WARNING: font name not found in docProps/app.xml — skipping")
        return content
    # Decrement the vector size attribute (size="N") by 1
    def decrement_size(m):
        n = int(m.group(1))
        return m.group(0).replace(f'size="{n}"', f'size="{n - 1}"')
    cleaned = re.sub(r'size="(\d+)"', decrement_size, cleaned, count=1)
    return cleaned

# ─── Read, patch, write ──────────────────────────────────────────────────────
with zipfile.ZipFile(TEMPLATE, 'r') as zin, \
     zipfile.ZipFile(TMP, 'w', zipfile.ZIP_DEFLATED) as zout:

    for item in zin.infolist():
        data = zin.read(item.filename)

        if item.filename == EMBEDDED_FILE:
            print(f"  Dropping: {item.filename}")
            continue  # skip — don't write font binary

        if item.filename in ('ppt/slides/slide5.xml', 'ppt/slides/slide14.xml'):
            text = data.decode('utf-8')
            if FONT_TO_REMOVE in text:
                text = fix_slide_xml(text)
                print(f"  Fixed font in: {item.filename}")
            data = text.encode('utf-8')

        elif item.filename == 'ppt/presentation.xml':
            text = data.decode('utf-8')
            text = fix_presentation_xml(text)
            print(f"  Cleaned embeddedFont from: {item.filename}")
            data = text.encode('utf-8')

        elif item.filename == 'ppt/_rels/presentation.xml.rels':
            text = data.decode('utf-8')
            text = fix_rels(text)
            print(f"  Removed {EMBEDDED_RID} from: {item.filename}")
            data = text.encode('utf-8')

        elif item.filename == 'docProps/app.xml':
            text = data.decode('utf-8')
            text = fix_app_xml(text)
            print(f"  Cleaned font name from: {item.filename}")
            data = text.encode('utf-8')

        zout.writestr(item, data)

# Replace original with patched version
os.replace(TMP, TEMPLATE)
print(f"\n✅ brand_master.pptx patched. Original backup at: {BACKUP}")

# ─── Verify ──────────────────────────────────────────────────────────────────
print("\nVerifying …")
with zipfile.ZipFile(TEMPLATE, 'r') as z:
    remaining = []
    for n in z.namelist():
        if n.endswith('.xml') or n.endswith('.rels'):
            try:
                if FONT_TO_REMOVE in z.read(n).decode('utf-8'):
                    remaining.append(n)
            except Exception:
                pass
    if remaining:
        print(f"  ⚠ Still found font in: {remaining}")
    else:
        print(f"  ✓ No references to '{FONT_TO_REMOVE}' remain")
    if EMBEDDED_FILE in z.namelist():
        print(f"  ⚠ Font binary still present: {EMBEDDED_FILE}")
    else:
        print(f"  ✓ Font binary removed ({EMBEDDED_FILE})")
