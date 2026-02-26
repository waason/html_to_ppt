const puppeteer = require('puppeteer');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

async function run() {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE'; // 設定為 16:9
    const browser = await puppeteer.launch({ args: ['--no-sandbox'] });
    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 720 });

    // 假設檔案名為 1.html 到 20.html
    for (let i = 1; i <= 20; i++) {
        const filePath = path.join(__dirname, `../slides/${i}.html`);
        if (fs.existsSync(filePath)) {
            console.log(`Processing slide ${i}...`);
            await page.goto(`file://${filePath}`, { waitUntil: 'networkidle0' });
            
            // 截圖存為 Buffer
            const screenshot = await page.screenshot({ encoding: 'base64' });
            
            // 加入 PPT
            const slide = pptx.addSlide();
            slide.addImage({ 
                data: `image/png;base64,${screenshot}`, 
                x: 0, y: 0, w: '100%', h: '100%' 
            });
        }
    }

    await pptx.writeFile({ fileName: 'Anti-Scam-2025.pptx' });
    console.log('PPTX created successfully!');
    await browser.close();
}

run().catch(console.error);
