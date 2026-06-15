import puppeteer from '/tmp/pdf-gen/node_modules/puppeteer/lib/puppeteer/puppeteer.js';
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOCS = path.join(__dirname, '..', 'docs');

const SRC  = path.join(DOCS, 'Vesuvius_proposal_v3.html');
const DEST = path.join(DOCS, 'Vesuvius_培训平台增强建议书_v3.0.pdf');

const browser = await puppeteer.launch({
  headless: true,
  executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
  args: ['--no-sandbox', '--disable-web-security', '--allow-file-access-from-files'],
});

const page = await browser.newPage();

// Use file:// URL so relative paths in HTML resolve correctly
await page.goto(`file://${SRC}`, { waitUntil: 'networkidle0', timeout: 30000 });

// Load Google Fonts for Chinese + English rendering
try {
  await page.addStyleTag({
    url: 'https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@300;400;500;600;700&family=Inter:wght@300;400;500;600;700&display=swap',
  });
} catch (_) {
  // Offline fallback — system fonts will be used
}

await page.evaluateHandle('document.fonts.ready');
await new Promise(r => setTimeout(r, 2500));

await page.pdf({
  path: DEST,
  width: '210mm',
  height: '297mm',
  printBackground: true,
  margin: { top: 0, bottom: 0, left: 0, right: 0 },
  preferCSSPageSize: true,
});

await browser.close();
console.log(`✓ PDF saved: ${DEST}`);
