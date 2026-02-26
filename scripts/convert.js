const puppeteer = require('puppeteer');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

async function run() {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';

    // 強制使用絕對路徑
    const rootDir = process.cwd();
    const htmlDir = path.join(rootDir, 'html_to_ppt/outhtml');
    const outDir = path.join(rootDir, 'outppt');

    console.log(`[Debug] 目標 HTML 路徑: ${htmlDir}`);
    console.log(`[Debug] 目標輸出路徑: ${outDir}`);

    if (!fs.existsSync(htmlDir)) {
        console.error(`❌ 錯誤：找不到 HTML 資料夾！請檢查路徑是否為 html_to_ppt/outhtml`);
        process.exit(1);
    }

    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

    // 讀取檔案
    const files = fs.readdirSync(htmlDir)
        .filter(f => f.endsWith('.html'))
        .sort((a, b) => parseInt(a) - parseInt(b));

    if (files.length === 0) {
        console.error(`❌ 錯誤：在資料夾內找不到任何 .html 檔案！`);
        process.exit(1);
    }

    console.log(`✅ 找到 ${files.length} 個檔案，開始渲染...`);

    const browser = await puppeteer.launch({ 
        args: ['--no-sandbox', '--disable-setuid-sandbox'] 
    });
    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 720 });

    for (const file of files) {
        const filePath = path.join(htmlDir, file);
        // 使用 file:// 協定開啟本地檔案
        await page.goto(`file://${filePath}`, { waitUntil: 'networkidle0' });
        await new Promise(r => setTimeout(r, 500)); // 等待動畫

        const screenshot = await page.screenshot({ encoding: 'base64' });
        const slide = pptx.addSlide();
        slide.addImage({ data: `image/png;base64,${screenshot}`, x: 0, y: 0, w: '100%', h: '100%' });
        console.log(`- 頁面 ${file} 已加入投影片`);
    }

    const outputFileName = '詐騙手法分析報告.pptx';
    const outputPath = path.join(outDir, outputFileName);
    
    // 儲存檔案
    await pptx.writeFile({ fileName: outputPath });
    console.log(`\n🎉 成功！檔案已產出至: ${outputPath}`);
    
    await browser.close();
}

run().catch(err => {
    console.error('執行過程發生崩潰:', err);
    process.exit(1);
});
