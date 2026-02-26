const puppeteer = require('puppeteer');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

async function run() {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE'; // 符合您的 1280x720 設計比例

    // 路徑定義
    const htmlDir = path.join(process.cwd(), 'html_to_ppt/outhtml');
    const outDir = path.join(process.cwd(), 'outppt');

    // 自動建立輸出目錄
    if (!fs.existsSync(outDir)) {
        fs.mkdirSync(outDir, { recursive: true });
    }

    // 讀取並排序 1.html ~ 20.html
    const files = fs.readdirSync(htmlDir)
        .filter(file => file.endsWith('.html'))
        .sort((a, b) => parseInt(a) - parseInt(b));

    const browser = await puppeteer.launch({ args: ['--no-sandbox'] });
    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 720 });

    for (const file of files) {
        console.log(`正在處理第 ${file} 頁...`);
        const filePath = path.join(htmlDir, file);
        await page.goto(`file://${filePath}`, { waitUntil: 'networkidle0' });
        
        // 確保 2025 年詐騙概況等複雜數據與動畫渲染完成
        await new Promise(r => setTimeout(r, 500)); 

        const screenshot = await page.screenshot({ encoding: 'base64' });
        const slide = pptx.addSlide();
        slide.addImage({ 
            data: `image/png;base64,${screenshot}`, 
            x: 0, y: 0, w: '100%', h: '100%' 
        });
    }

    const outputFileName = '2025_詐騙趨勢分析報告.pptx';
    const outputPath = path.join(outDir, outputFileName);
    
    // 使用 await 確保檔案完全寫入磁碟
    await pptx.writeFile({ fileName: outputPath });
    console.log(`✅ 成功產出至: ${outputPath}`);
    
    await browser.close();
}

run().catch(console.error);
