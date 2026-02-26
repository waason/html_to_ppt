const puppeteer = require('puppeteer');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

async function run() {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';

    const rootDir = process.cwd();
    // 修正點：直接指向根目錄下的 inhtml
    const htmlDir = path.join(rootDir, 'inhtml');
    const outDir = path.join(rootDir, 'outppt');

    console.log(`[Debug] 正在掃描目錄: ${htmlDir}`);

    if (!fs.existsSync(htmlDir)) {
        console.error(`❌ 錯誤：找不到目錄 ${htmlDir}`);
        console.log("根目錄內容為：", fs.readdirSync(rootDir));
        process.exit(1);
    }

    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

    const files = fs.readdirSync(htmlDir)
        .filter(f => f.endsWith('.html'))
        .sort((a, b) => parseInt(a) - parseInt(b));

    console.log(`✅ 找到 ${files.length} 個檔案，開始轉換...`);

    const browser = await puppeteer.launch({ 
        args: ['--no-sandbox', '--disable-setuid-sandbox'] 
    });
    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 720 });

    for (const file of files) {
        const filePath = path.join(htmlDir, file);
        await page.goto(`file://${filePath}`, { waitUntil: 'networkidle0' });
        await new Promise(r => setTimeout(r, 1000)); // 等待動畫

        const screenshot = await page.screenshot({ encoding: 'base64' });
        const slide = pptx.addSlide();
        slide.addImage({ data: `image/png;base64,${screenshot}`, x: 0, y: 0, w: '100%', h: '100%' });
        console.log(`- 已完成: ${file}`);
    }

    const outputPath = path.join(outDir, '2025_防詐分析報告.pptx');
    await pptx.writeFile({ fileName: outputPath });
    console.log(`\n🎉 轉換成功！產出檔案：${outputPath}`);
    await browser.close();
}

run().catch(err => { console.error(err); process.exit(1); });
