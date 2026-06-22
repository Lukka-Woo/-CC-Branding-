#!/usr/bin/env node
// assets/icons/render.mjs — SVG → PNG batch converter
//
// Usage:
//   node assets/icons/render.mjs             # convert all, 64px (default)
//   node assets/icons/render.mjs --size=128  # higher resolution
//
// Reads:  assets/icons/svg/*.svg
// Writes: assets/icons/png/*.png
//
// Dependencies: sharp  (installed once in /tmp/icon-gen)
//   cd /tmp/icon-gen && npm install sharp
//
// Icon library: Lucide (lucide.dev) — MIT license
// Drop any .svg file from Lucide into assets/icons/svg/ and re-run this script.
// For HTML output, use the SVG files directly (<img> or inline).
// For PPTX output, the builder reads from assets/icons/png/.

import { createRequire } from 'module';
import { readdir, mkdir } from 'fs/promises';
import { existsSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SHARP_DIR  = '/tmp/icon-gen';
const SVG_DIR    = path.join(__dirname, 'svg');
const PNG_DIR    = path.join(__dirname, 'png');

// Parse --size=N flag
const sizeArg = process.argv.find(a => a.startsWith('--size='));
const SIZE    = sizeArg ? parseInt(sizeArg.split('=')[1], 10) : 64;

// Load sharp from /tmp/icon-gen
if (!existsSync(path.join(SHARP_DIR, 'node_modules', 'sharp'))) {
  console.error('sharp not found. Install it first:\n  mkdir -p /tmp/icon-gen && cd /tmp/icon-gen && npm install sharp');
  process.exit(1);
}
const req   = createRequire(new URL(`file://${SHARP_DIR}/`));
const sharp = req('sharp');

await mkdir(PNG_DIR, { recursive: true });

const files = (await readdir(SVG_DIR).catch(() => [])).filter(f => f.endsWith('.svg'));
if (!files.length) {
  console.log('No SVGs found in assets/icons/svg/');
  console.log('Download icons from https://lucide.dev and place them there.');
  process.exit(0);
}

// Lucide icons are black-stroke on transparent. Recolor to neutral-700 (#3D444A)
// so they look correct on white card backgrounds.
const ICON_COLOR = '#3D444A';
const recolor = (svg) =>
  svg.replace(/stroke="currentColor"/g, `stroke="${ICON_COLOR}"`)
     .replace(/fill="none"/g, 'fill="none"');

console.log(`Converting ${files.length} SVG → ${SIZE}×${SIZE}px PNG …`);

for (const file of files) {
  const svgPath = path.join(SVG_DIR, file);
  const pngName = file.replace('.svg', '.png');
  const pngPath = path.join(PNG_DIR, pngName);

  const { readFile } = await import('fs/promises');
  const svgSrc = await readFile(svgPath, 'utf8');
  const svgColored = Buffer.from(recolor(svgSrc));

  await sharp(svgColored)
    .resize(SIZE, SIZE, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } })
    .png()
    .toFile(pngPath);

  console.log(`  ✓ ${file} → png/${pngName}`);
}

console.log(`Done. PNGs saved to: ${PNG_DIR}`);
