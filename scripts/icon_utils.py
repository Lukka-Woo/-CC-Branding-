"""
Auto-fetch Lucide icons that are missing from assets/icons/png/.

Pipeline:  icon name  →  download SVG  →  render.mjs  →  PNG  →  pptx_builder

Usage in gen scripts:
    from scripts.icon_utils import ensure_icons
    ensure_icons(["satellite", "bell", "sparkles"], _BRAND)
"""
import os
import subprocess
import urllib.request
import urllib.error

LUCIDE_RAW = "https://raw.githubusercontent.com/lucide-icons/lucide/main/icons"


def ensure_icons(icon_names: list, brand_root: str) -> None:
    """
    For each name, ensure assets/icons/png/{name}.png exists.
    Downloads missing SVGs from Lucide and converts via render.mjs.

    Raises FileNotFoundError if a name doesn't exist in Lucide.
    Raises RuntimeError if render.mjs fails.
    """
    svg_dir    = os.path.join(brand_root, "assets", "icons", "svg")
    png_dir    = os.path.join(brand_root, "assets", "icons", "png")
    render_mjs = os.path.abspath(os.path.join(brand_root, "assets", "icons", "render.mjs"))

    os.makedirs(svg_dir, exist_ok=True)
    os.makedirs(png_dir, exist_ok=True)

    needs_convert = []

    for name in icon_names:
        if not name:
            continue
        # Skip non-filename values (emoji, absolute paths)
        if len(name) > 64 or os.sep in name or name.startswith("/"):
            continue

        png_path = os.path.join(png_dir, f"{name}.png")
        if os.path.exists(png_path):
            continue

        svg_path = os.path.join(svg_dir, f"{name}.svg")
        if not os.path.exists(svg_path):
            url = f"{LUCIDE_RAW}/{name}.svg"
            try:
                urllib.request.urlretrieve(url, svg_path)
                print(f"  ↓ lucide/{name}.svg")
            except urllib.error.HTTPError as exc:
                os.path.exists(svg_path) and os.remove(svg_path)
                raise FileNotFoundError(
                    f"Icon '{name}' not found in Lucide — "
                    f"check the name at https://lucide.dev/icons/ ({exc})"
                ) from exc

        needs_convert.append(name)

    if not needs_convert:
        return

    print(f"Converting {len(needs_convert)} new SVG(s) → PNG …")
    result = subprocess.run(
        ["node", render_mjs],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"render.mjs failed (exit {result.returncode}):\n{result.stderr}"
        )
    print(f"  ✓ {', '.join(needs_convert)}")
