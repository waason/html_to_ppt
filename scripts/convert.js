const puppeteer = require('puppeteer');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

async function run() {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE'; // 16:9 比例

    const rootDir = process.cwd();
    // 修正為您指定的新路徑
    const htmlDir = path.join(rootDir, 'html_to_ppt/inhtml');
    const outDir = path.join(rootDir, 'outppt');

    console.log(`[Debug] 正在掃描目錄: ${htmlDir}`);

    if (!fs.existsSync(htmlDir)) {
        console.error(`❌ 錯誤：找不到目錄 ${htmlDir}`);
        process.exit(1);
    }

    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

    // 讀取檔案並按數字排序 (1, 2, 3... 20)
    const files = fs.readdirSync(htmlDir)
        .filter(f => f.endsWith('.html'))
        .sort((a, b) => {
            const numA = parseInt(a.replace(/[^0-9]/g, ''));
            const numB = parseInt(b.replace(/[^0-9]/g, ''));
            return numA - numB;
        });

    if (files.length === 0) {
        console.error("❌ 錯誤：資料夾內沒有 .html 檔案");
        process.exit(1);
    }

    console.log(`✅ 找到 ${files.length} 個檔案，準備轉換...`);

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
        // 給予足夠時間讓 2025 詐騙財損圖表動畫完成
        await new Promise(r => setTimeout(r, 1000));

        const screenshot = await page.screenshot({ encoding: 'base64' });
        const slide = pptx.addSlide();
        slide.addImage({ 
            data: `image/png;base64,${screenshot}`, 
            x: 0, y: 0, w: '100%', h: '100%' 
        });
    }

    const outputPath = path.join(outDir, '2025_詐騙分析報告.pptx');
    await pptx.writeFile({ fileName: outputPath });
    
    console.log(`\n🎉 轉換成功！檔案已存至: ${outputPath}`);
    await browser.close();
}

run().catch(err => {
    console.error('執行失敗:', err);
    process.exit(1);
});
