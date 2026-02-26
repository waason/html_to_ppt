const puppeteer = require('puppeteer');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

async function run() {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE'; // 設定為 16:9 寬螢幕

    // 指定您的路徑
    const htmlDir = path.join(__dirname, '../html_to_ppt/outhtml');
    
    // 讀取該資料夾下所有 .html 檔案並排序 (1.html, 2.html...)
    const files = fs.readdirSync(htmlDir)
        .filter(file => file.endsWith('.html'))
        .sort((a, b) => parseInt(a) - parseInt(b));

    const browser = await puppeteer.launch({ 
        args: ['--no-sandbox', '--disable-setuid-sandbox'] 
    });
    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 720 });

    console.log(`找到 ${files.length} 份檔案，開始轉換...`);

    for (const file of files) {
        const filePath = path.join(htmlDir, file);
        console.log(`正在處理: ${file}`);
        
        await page.goto(`file://${filePath}`, { waitUntil: 'networkidle0' });
        
        // 針對 2025 詐騙分析簡報的動畫，等待 500ms 確保渲染完成
        await new Promise(r => setTimeout(r, 500));

        const screenshot = await page.screenshot({ encoding: 'base64' });
        
        const slide = pptx.addSlide();
        slide.addImage({ 
            data: `image/png;base64,${screenshot}`, 
            x: 0, y: 0, w: '100%', h: '100%' 
        });
    }

    const outputFile = '2025_詐騙高風險廣告手法分析.pptx';
    await pptx.writeFile({ fileName: outputFile });
    console.log(`✅ 轉換完成！檔案已儲存為: ${outputFile}`);
    
    await browser.close();
}

run().catch(console.error);
