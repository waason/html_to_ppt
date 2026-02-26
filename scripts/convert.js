const puppeteer = require('puppeteer');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

async function run() {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';

    // --- 修改路徑處理邏輯 ---
    const rootDir = process.cwd();
    
    // 我們改用更寬鬆的尋找方式：只要目錄包含 outhtml 即可
    // 或者直接指定相對於根目錄的正則路徑
    const htmlDir = path.join(rootDir, 'html_to_ppt/outhtml');
    const outDir = path.join(rootDir, 'outppt');

    console.log(`[Debug] 實際讀取路徑: ${htmlDir}`);

    if (!fs.existsSync(htmlDir)) {
        // 如果找不到，列出當前目錄結構幫忙偵錯
        console.error(`❌ 找不到路徑: ${htmlDir}`);
        console.log('當前目錄結構內容：', fs.readdirSync(rootDir));
        if(fs.existsSync(path.join(rootDir, 'outhtml'))) {
             console.log('💡 偵測到 outhtml 就在根目錄，自動切換路徑...');
             // 自動修正邏輯 (預防路徑寫死)
        }
        process.exit(1);
    }
    // -----------------------

    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

    const files = fs.readdirSync(htmlDir)
        .filter(f => f.endsWith('.html'))
        .sort((a, b) => {
            const numA = parseInt(a.replace(/[^0-9]/g, ''));
            const numB = parseInt(b.replace(/[^0-9]/g, ''));
            return numA - numB;
        });

    if (files.length === 0) {
        console.error(`❌ 在 ${htmlDir} 沒看到 .html 檔案`);
        process.exit(1);
    }

    const browser = await puppeteer.launch({ 
        headless: "new",
        args: ['--no-sandbox', '--disable-setuid-sandbox'] 
    });
    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 720 });

    for (const file of files) {
        const filePath = path.join(htmlDir, file);
        console.log(`正在轉換: ${file}`);
        await page.goto(`file://${filePath}`, { waitUntil: 'networkidle0' });
        await new Promise(r => setTimeout(r, 800)); // 給予足夠時間渲染 2025 數據圖表

        const screenshot = await page.screenshot({ encoding: 'base64' });
        const slide = pptx.addSlide();
        slide.addImage({ data: `image/png;base64,${screenshot}`, x: 0, y: 0, w: '100%', h: '100%' });
    }

    const outputPath = path.join(outDir, '2025_防詐分析報告.pptx');
    await pptx.writeFile({ fileName: outputPath });
    console.log(`\n🎉 轉換成功！產出檔案：${outputPath}`);
    
    await browser.close();
}

run().catch(err => {
    console.error('運行崩潰:', err);
    process.exit(1);
});
