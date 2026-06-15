/**
 * render_final_v3.mjs
 * Renders Vesuvius_proposal_v3_final.html → Vesuvius_培训平台增强建议书_v3_final.pdf
 *
 * Image paths in the HTML are bare filenames; we resolve them to absolute
 * file:// URLs before handing the content to Puppeteer.
 */

import puppeteer from '/tmp/pdf-gen/node_modules/puppeteer/lib/puppeteer/puppeteer.js';
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const BRAND   = path.resolve(__dirname, '..', '..', '..');
const PROJECT = path.resolve(__dirname, '..');
const DOCS    = path.join(PROJECT, 'docs');
const MEDIA   = path.join(PROJECT, 'media');

const SRC  = path.join(DOCS, 'Vesuvius_proposal_v3_final.html');
const DEST = path.join(DOCS, 'Vesuvius_培训平台增强建议书_v3_final.pdf');

// ── Absolute paths for the two images ────────────────────────────────────────
const VESUVIUS_LOGO = path.join(MEDIA, 'covers', 'VESUVIUS_logo.png');
const OUR_LOGO      = path.join(BRAND, 'assets', 'logo-horizontal-primary.png');

// Convert to file:// URLs safe for HTML src attributes
function toFileUrl(p) {
  return 'file://' + p.replace(/ /g, '%20');
}

// ── Patch HTML ────────────────────────────────────────────────────────────────
let html = readFileSync(SRC, 'utf8');

// Replace bare filenames (possibly with leading ./ too) with absolute URLs
html = html
  .replace(/src="(?:\.\/)?VESUVIUS_logo\.png"/g,
           `src="${toFileUrl(VESUVIUS_LOGO)}"`)
  .replace(/src="(?:\.\/)?logo-horizontal-primary\.png"/g,
           `src="${toFileUrl(OUR_LOGO)}"`)
  // Remove the Cloudflare beacon script so it doesn't slow down rendering
  .replace(/<script[^>]*cloudflareinsights[^>]*>[\s\S]*?<\/script>/gi, '');

// ── Puppeteer ─────────────────────────────────────────────────────────────────
const browser = await puppeteer.launch({
  headless: true,
  executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
  args: ['--no-sandbox', '--disable-web-security', '--allow-file-access-from-files'],
});

const page = await browser.newPage();

// Set content directly so all file:// URLs resolve with no base-URL issues
await page.setContent(html, { waitUntil: 'networkidle0' });

// Load Google Fonts (Chinese + English) — skip gracefully if offline
try {
  await page.addStyleTag({
    url: 'https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@300;400;500;600;700&family=Inter:wght@300;400;500;600;700&display=swap',
  });
} catch (_) { /* use system fonts as fallback */ }

await page.evaluateHandle('document.fonts.ready');

// Give images and layout time to settle
await new Promise(r => setTimeout(r, 2800));

await page.pdf({
  path: DEST,
  width:  '210mm',
  height: '297mm',
  printBackground: true,
  margin: { top: 0, bottom: 0, left: 0, right: 0 },
  preferCSSPageSize: true,
});

await browser.close();
console.log(`✓ PDF saved: ${DEST}`);
