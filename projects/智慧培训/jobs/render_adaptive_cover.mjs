/**
 * Puppeteer PDF 渲染器 - 自适应封面版本
 * 专为新的统一封面系统设计，支持从 HTML 生成高质量 PDF
 */

import puppeteer from 'puppeteer';
import { readFileSync, writeFileSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// 路径常量
const PROJECT_DIR = join(__dirname, '..');
const DOCS_DIR = join(PROJECT_DIR, 'docs');

/**
 * 从 HTML 生成 PDF
 * @param {string} htmlPath - 输入 HTML 文件路径
 * @param {string} outputPath - 输出 PDF 文件路径
 * @param {Object} options - PDF 生成选项
 */
async function renderPDF(htmlPath, outputPath, options = {}) {
    const defaultOptions = {
        format: 'A4',
        printBackground: true,
        margin: { top: 0, bottom: 0, left: 0, right: 0 },
        preferCSSPageSize: true,
        displayHeaderFooter: false,
        ...options
    };

    console.log(`📄 开始渲染: ${htmlPath} → ${outputPath}`);

    const browser = await puppeteer.launch({
        headless: true,
        executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
        args: [
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--disable-dev-shm-usage',
            '--disable-gpu',
            '--no-first-run',
            '--no-default-browser-check'
        ],
    });

    try {
        const page = await browser.newPage();

        // 读取 HTML 内容
        const htmlContent = readFileSync(htmlPath, 'utf8');
        await page.setContent(htmlContent, {
            waitUntil: 'networkidle0',
            timeout: 30000
        });

        // 加载品牌字体
        await page.addStyleTag({
            url: 'https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@300;400;500;600;700&family=Inter:wght@300;400;500;600;700&display=swap'
        });

        // 等待字体加载
        await page.evaluateHandle('document.fonts.ready');

        // 额外等待时间确保渲染稳定
        await new Promise(resolve => setTimeout(resolve, 2000));

        // 生成 PDF
        await page.pdf({
            path: outputPath,
            ...defaultOptions,
        });

        console.log(`✅ PDF 生成完成: ${outputPath}`);

    } catch (error) {
        console.error(`❌ PDF 生成失败:`, error.message);
        throw error;
    } finally {
        await browser.close();
    }
}

/**
 * 批量渲染多个 HTML 文件
 * @param {Array} tasks - 渲染任务列表 [{html: 'input.html', pdf: 'output.pdf'}]
 */
async function batchRender(tasks) {
    console.log(`🚀 开始批量渲染 ${tasks.length} 个文件...`);

    for (const [index, task] of tasks.entries()) {
        try {
            await renderPDF(
                join(DOCS_DIR, task.html),
                join(DOCS_DIR, task.pdf),
                task.options || {}
            );
            console.log(`📊 进度: ${index + 1}/${tasks.length}`);
        } catch (error) {
            console.error(`❌ 任务 ${index + 1} 失败:`, error.message);
        }
    }

    console.log('🎉 批量渲染完成!');
}

/**
 * 生成封面专用的高质量 PDF
 * @param {string} htmlPath - HTML 封面文件
 * @param {string} outputPath - 输出 PDF 路径
 */
async function renderCover(htmlPath, outputPath) {
    await renderPDF(htmlPath, outputPath, {
        format: 'A4',
        printBackground: true,
        margin: { top: 0, bottom: 0, left: 0, right: 0 },
        preferCSSPageSize: true,
        displayHeaderFooter: false,
        // 封面专用设置
        quality: 100,  // 最高质量（仅对 JPEG 有效，PDF 忽略此参数）
    });
}

// 如果直接运行此脚本，执行默认渲染任务
if (import.meta.url === `file://${process.argv[1]}`) {
    (async () => {
        try {
            // 检查是否有测试 HTML 文件
            const testHtmlPath = join(DOCS_DIR, 'test_cover.html');
            const testPdfPath = join(DOCS_DIR, 'test_cover_from_html.pdf');

            try {
                await renderCover(testHtmlPath, testPdfPath);
            } catch (error) {
                console.log('⚠️  没有找到测试封面 HTML，生成一个简单示例...');

                // 创建简单的测试 HTML
                const simpleHtml = `<!DOCTYPE html>
<html lang="zh-Hans">
<head>
    <meta charset="utf-8">
    <title>测试封面</title>
    <style>
        .cover {
            width: 210mm;
            height: 297mm;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            background: linear-gradient(45deg, #3EC99E, #C8E13C);
            color: white;
            font-family: "Microsoft YaHei", sans-serif;
        }
        h1 { font-size: 36pt; margin-bottom: 12mm; }
        p { font-size: 14pt; }
    </style>
</head>
<body>
    <div class="cover">
        <h1>自适应封面测试</h1>
        <p>这是一个测试封面，用于验证 PDF 渲染功能</p>
        <p>生成时间: ${new Date().toLocaleString('zh-CN')}</p>
    </div>
</body>
</html>`;

                writeFileSync(testHtmlPath, simpleHtml, 'utf8');
                console.log(`📝 已创建测试 HTML: ${testHtmlPath}`);

                await renderCover(testHtmlPath, testPdfPath);
            }

        } catch (error) {
            console.error('❌ 渲染失败:', error.message);
            process.exit(1);
        }
    })();
}

export { renderPDF, batchRender, renderCover };