#!/usr/bin/env python3
"""
gen_font_css.py — 字体配置一键同步到 HTML/PDF/Puppeteer

工作流：
  1. 读取 brand_tokens.FONT_CN（当前激活字体名）
  2. 扫描 fonts/ 子目录，用 fonttools 匹配正确的字体包
  3. 读取每个字体文件的 usWeightClass，生成精确的 @font-face
  4. 更新 templates/html/brand.css（@font-face 块 + --font-cn 变量）
  5. 更新 templates/pdf/brand.css（@top-left font-family）
  6. 写出 templates/pdf/_font_inject.js（Puppeteer 注入脚本）

用法：
  python3 scripts/gen_font_css.py              # 正式写入
  python3 scripts/gen_font_css.py --dry-run    # 预览，不写文件
  python3 scripts/gen_font_css.py --list       # 列出 fonts/ 下所有可识别字体
"""

import os, re, sys, argparse

_BRAND = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, _BRAND)
import scripts.brand_tokens as BT

# ── 路径常量 ──────────────────────────────────────────────────────────────────
FONTS_DIR      = os.path.join(_BRAND, 'fonts')
BRAND_CSS      = os.path.join(_BRAND, 'templates', 'html', 'brand.css')
PDF_CSS        = os.path.join(_BRAND, 'templates', 'pdf', 'brand.css')
PDF_INJECT_JS  = os.path.join(_BRAND, 'templates', 'pdf', '_font_inject.js')

# CSS 中 @font-face 相对路径（从 templates/html/ 出发）
CSS_FONT_REL = '../../fonts'

# 支持的字体扩展名及对应 CSS format()
FONT_EXTS = {'.ttf': 'truetype', '.otf': 'opentype',
             '.woff': 'woff', '.woff2': 'woff2'}

# brand.css 里 @font-face 块的边界注释（精确到前缀，不含尾部细节）
_FACE_OPEN  = '/* ── @font-face'
_FACE_CLOSE = '/* ── CSS Custom Properties'


# ── fonttools 读取 ─────────────────────────────────────────────────────────────

def _font_families(path: str) -> set:
    """返回字体文件所有平台的 family name（nameID 1 + 16 合集）。"""
    try:
        from fontTools.ttLib import TTFont
        tt = TTFont(path, lazy=True)
        families = set()
        for rec in tt['name'].names:
            if rec.nameID in (1, 16):
                try:
                    families.add(rec.toUnicode())
                except Exception:
                    pass
        return families
    except Exception:
        return set()


def _font_weight(path: str) -> int:
    """读取 OS/2.usWeightClass；失败则从文件名关键字推断。"""
    try:
        from fontTools.ttLib import TTFont
        tt = TTFont(path, lazy=True)
        w = tt['OS/2'].usWeightClass
        return max(1, min(1000, w))
    except Exception:
        pass
    return _weight_from_name(os.path.basename(path))


def _weight_from_name(filename: str) -> int:
    """文件名关键字 → CSS font-weight（兜底方案）。"""
    kw_map = [
        ('black',      900), ('heavy',      900), ('extrabold',  800),
        ('ultrabold',  800), ('bold',        700), ('semibold',   600),
        ('demibold',   600), ('medium',      500), ('regular',    400),
        ('normal',     400), ('book',        400), ('light',      300),
        ('extralight', 200), ('ultralight',  200), ('thin',       100),
        ('hairline',   100),
    ]
    name = re.sub(r'[\-_\s]', '', filename.lower())
    for kw, w in sorted(kw_map, key=lambda x: -len(x[0])):
        if kw.replace('-', '') in name:
            return w
    return 400


# ── 字体目录查找 ───────────────────────────────────────────────────────────────

def _normalize(s: str) -> str:
    return re.sub(r'[\s\-_\.\d]+', '', s).lower()


def find_font_dir(font_cn: str):
    """
    在 fonts/ 里找出与 font_cn 匹配的子目录。
    优先用 fonttools 匹配 family name；其次模糊比较目录名。
    返回 (dirpath, subdir_name) 或 (None, None)。
    """
    if not os.path.isdir(FONTS_DIR):
        return None, None

    target_norm = _normalize(font_cn)
    candidates = []

    for subdir in sorted(os.listdir(FONTS_DIR)):
        dirpath = os.path.join(FONTS_DIR, subdir)
        if not os.path.isdir(dirpath):
            continue

        font_files = [f for f in os.listdir(dirpath)
                      if os.path.splitext(f)[1].lower() in FONT_EXTS]
        if not font_files:
            continue

        # 用第一个文件的 family names 做精确匹配
        sample = os.path.join(dirpath, font_files[0])
        families = _font_families(sample)
        if any(_normalize(f) == target_norm for f in families):
            return dirpath, subdir

        # 记录模糊分数（目录名 vs font_cn 的公共前缀长度）
        score = len(os.path.commonprefix([_normalize(subdir), target_norm]))
        candidates.append((score, dirpath, subdir))

    if candidates:
        candidates.sort(key=lambda x: -x[0])
        best_score, dirpath, subdir = candidates[0]
        if best_score >= 4:   # 至少 4 字符公共前缀才算匹配
            return dirpath, subdir

    return None, None


def collect_faces(dirpath: str) -> list:
    """返回 [(weight, filename, format)] 按 weight 升序，同 weight 取第一个文件。"""
    seen_weights: dict = {}
    for fname in sorted(os.listdir(dirpath)):
        ext = os.path.splitext(fname)[1].lower()
        if ext not in FONT_EXTS:
            continue
        fpath = os.path.join(dirpath, fname)
        w   = _font_weight(fpath)
        fmt = FONT_EXTS[ext]
        if w not in seen_weights:
            seen_weights[w] = (fname, fmt)

    return sorted((w, fname, fmt) for w, (fname, fmt) in seen_weights.items())


# ── CSS / JS 生成 ──────────────────────────────────────────────────────────────

_FACE_HEADER_TMPL = (
    '/* ── @font-face — {family}（{n} 字重）'
    '─────────────────────────────────────────────────────── */\n'
    '/* 路径相对于此 CSS 文件位置：templates/html/ → {rel}/  */\n'
)


def build_face_block(font_cn: str, subdir: str, faces: list) -> str:
    header = _FACE_HEADER_TMPL.format(
        family=font_cn, n=len(faces), rel=f'{CSS_FONT_REL}/{subdir}')
    rules = []
    for weight, fname, fmt in faces:
        rel_path = f'{CSS_FONT_REL}/{subdir}/{fname}'
        rules.append(
            f'@font-face {{\n'
            f'  font-family: "{font_cn}";\n'
            f'  font-weight: {weight};\n'
            f'  font-style: normal;\n'
            f'  src: url("{rel_path}") format("{fmt}");\n'
            f'}}'
        )
    return header + '\n'.join(rules)


def build_js_inject(font_cn: str, subdir: str, faces: list) -> str:
    """生成 templates/pdf/_font_inject.js（Puppeteer 用）。"""
    face_lines = []
    for weight, fname, fmt in faces:
        rel = f'{subdir}/{fname}'
        # JS 模板字面量中 ${_f(...)} 在运行时被替换为绝对 file:// 路径
        face_str = (
            f'@font-face{{font-family:"{font_cn}";'
            f'font-weight:{weight};font-style:normal;'
            f"src:url(\"${{_f('{rel}')}}\") format(\"{fmt}\")}}"
        )
        face_lines.append(f'  `{face_str}`,')

    faces_js = '\n'.join(face_lines)
    return f"""\
// templates/pdf/_font_inject.js — auto-generated by scripts/gen_font_css.py
// DO NOT EDIT MANUALLY. Re-run the script after changing fonts or FONT_CN.
//
// Usage in render.mjs:
//   import {{ FONT_CSS }} from '../../templates/pdf/_font_inject.js';
//   await page.addStyleTag({{ content: FONT_CSS }});

import path from 'path';
import {{ fileURLToPath }} from 'url';

const _d = path.dirname(fileURLToPath(import.meta.url));
const _f = (p) => 'file://' + path.resolve(_d, '..', '..', 'fonts', p);

export const FONT_CSS = [
{faces_js}
].join('');
"""


# ── 文件更新 ───────────────────────────────────────────────────────────────────

def update_brand_css(face_block: str, font_cn: str, dry_run: bool = False):
    with open(BRAND_CSS, 'r', encoding='utf-8') as f:
        css = f.read()

    # 替换 @font-face 块（从 _FACE_OPEN 到 _FACE_CLOSE 之前）
    start = css.find(_FACE_OPEN)
    end   = css.find(_FACE_CLOSE)
    if start == -1 or end == -1 or start >= end:
        print('⚠  brand.css 找不到 @font-face 标记区间，跳过 @font-face 更新')
    else:
        css = css[:start] + face_block + '\n\n' + css[end:]

    # 替换 --font-cn 变量值
    fallback = '"PingFang SC", "Microsoft YaHei", "Noto Sans SC", sans-serif'
    css = re.sub(
        r'--font-cn:\s+[^\n]+;',
        f'--font-cn:    "{font_cn}", {fallback};',
        css,
    )

    if dry_run:
        _preview(css, 'brand.css')
        return
    with open(BRAND_CSS, 'w', encoding='utf-8') as f:
        f.write(css)
    print(f'  ✓ templates/html/brand.css   --font-cn → "{font_cn}"，@font-face 已更新')


def update_pdf_css(font_cn: str, dry_run: bool = False):
    with open(PDF_CSS, 'r', encoding='utf-8') as f:
        css = f.read()

    fallback = '"PingFang SC", "Microsoft YaHei", "Noto Sans SC", sans-serif'
    new_val  = f'"{font_cn}", {fallback}'
    # 替换 @top-left 里的 font-family（只改第一处，即中文页眉）
    css = re.sub(
        r'(font-family:\s*)(?:"[^"]+",?\s*)+(?:sans-serif)?;',
        lambda m: m.group(1) + new_val + ';',
        css,
        count=1,
    )

    if dry_run:
        _preview(css, 'pdf/brand.css')
        return
    with open(PDF_CSS, 'w', encoding='utf-8') as f:
        f.write(css)
    print(f'  ✓ templates/pdf/brand.css    font-family → "{font_cn}"')


def write_js_inject(js_content: str, dry_run: bool = False):
    if dry_run:
        _preview(js_content, '_font_inject.js')
        return
    os.makedirs(os.path.dirname(PDF_INJECT_JS), exist_ok=True)
    with open(PDF_INJECT_JS, 'w', encoding='utf-8') as f:
        f.write(js_content)
    print(f'  ✓ templates/pdf/_font_inject.js 已生成（Puppeteer 注入脚本）')


def _preview(content: str, label: str):
    lines = content.splitlines()
    print(f'\n── {label} 预览（前 30 行）──')
    for i, line in enumerate(lines[:30], 1):
        print(f'{i:3}  {line}')
    if len(lines) > 30:
        print(f'    ... （共 {len(lines)} 行）')


# ── --list 命令 ────────────────────────────────────────────────────────────────

def cmd_list():
    if not os.path.isdir(FONTS_DIR):
        print(f'fonts/ 目录不存在：{FONTS_DIR}')
        return

    print(f'fonts/ 目录：{FONTS_DIR}\n')
    found_any = False
    for subdir in sorted(os.listdir(FONTS_DIR)):
        dirpath = os.path.join(FONTS_DIR, subdir)
        if not os.path.isdir(dirpath):
            continue
        files = [f for f in sorted(os.listdir(dirpath))
                 if os.path.splitext(f)[1].lower() in FONT_EXTS]
        if not files:
            continue
        found_any = True

        # 读 family name
        families = _font_families(os.path.join(dirpath, files[0]))
        family_str = ' / '.join(sorted(families)) if families else '（无法读取）'
        print(f'  [{subdir}]')
        print(f'    family names : {family_str}')
        print(f'    字体文件数   : {len(files)}')

        faces = collect_faces(dirpath)
        for w, fname, fmt in faces:
            marker = ' ◀ FONT_CN 匹配' if any(
                _normalize(f) == _normalize(BT.FONT_CN) for f in families
            ) else ''
            print(f'    w={w:4d}  {fname}{marker}')
        print()

    if not found_any:
        print('  （fonts/ 下没有找到字体文件）')

    print(f'当前 brand_tokens.FONT_CN = {BT.FONT_CN!r}')


# ── 主流程 ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description='将 brand_tokens.FONT_CN 同步到 HTML/PDF CSS 和 Puppeteer 脚本')
    parser.add_argument('--dry-run', action='store_true',
                        help='预览输出，不写入文件')
    parser.add_argument('--list', action='store_true',
                        help='列出 fonts/ 目录下所有可识别字体')
    args = parser.parse_args()

    if args.list:
        cmd_list()
        return

    font_cn = BT.FONT_CN
    print(f'FONT_CN = {font_cn!r}')

    # 定位字体目录
    dirpath, subdir = find_font_dir(font_cn)
    if dirpath is None:
        print(
            f'\n✗  在 fonts/ 中找不到匹配 {font_cn!r} 的字体目录。\n'
            f'  请把字体文件放入 fonts/{{目录名}}/，再重新运行。\n'
            f'  运行 --list 查看已有目录。'
        )
        sys.exit(1)

    print(f'字体目录 : fonts/{subdir}/')

    faces = collect_faces(dirpath)
    if not faces:
        print(f'✗  {dirpath} 里没有找到字体文件')
        sys.exit(1)

    print(f'字重     : {[w for w, _, _ in faces]}')
    if args.dry_run:
        print('[dry-run 模式，不写入文件]\n')

    face_block = build_face_block(font_cn, subdir, faces)
    js_content = build_js_inject(font_cn, subdir, faces)

    print()
    update_brand_css(face_block, font_cn, dry_run=args.dry_run)
    update_pdf_css(font_cn,               dry_run=args.dry_run)
    write_js_inject(js_content,           dry_run=args.dry_run)

    if not args.dry_run:
        print('\n✓ 全部完成。新生成的文档将使用字体：', font_cn)


if __name__ == '__main__':
    main()
