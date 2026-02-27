const puppeteer = require('puppeteer');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

// 簡易的多執行緒限制器 (Concurrency Limiter)
const limitConcurrency = async (tasks, limit) => {
    const results = [];
    const executing = [];
    for (const task of tasks) {
        const p = Promise.resolve().then(() => task());
        results.push(p);
        if (limit <= tasks.length) {
            const e = p.then(() => executing.splice(executing.indexOf(e), 1));
            executing.push(e);
            if (executing.length >= limit) {
                await Promise.race(executing);
            }
        }
    }
    return Promise.all(results);
};

// 解析命令列參數
const args = process.argv.slice(2);
let inputDirStr = 'inhtml';
let outputDirStr = 'outppt';

for (let i = 0; i < args.length; i++) {
    if (args[i] === '--inputDir' && args[i + 1]) {
        inputDirStr = args[i + 1];
        i++;
    } else if (args[i] === '--outputDir' && args[i + 1]) {
        outputDirStr = args[i + 1];
        i++;
    }
}

async function run() {
    const rootDir = process.cwd();
    const htmlDir = path.resolve(rootDir, inputDirStr);
    const outDir = path.resolve(rootDir, outputDirStr);
    const outputPath = path.join(outDir, 'output.pptx');

    console.log(`[Info] 輸入目錄: ${htmlDir}`);
    console.log(`[Info] 輸出目錄: ${outDir}`);

    if (!fs.existsSync(htmlDir)) {
        console.error(`❌ 錯誤：找不到輸入目錄 ${htmlDir}`);
        process.exit(1);
    }

    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

    let files = fs.readdirSync(htmlDir)
        .filter(f => f.endsWith('.html'));

    // 檔名排序 (假設檔名是數字)
    files.sort((a, b) => {
        const numA = parseInt(a.match(/\d+/) || [0])[0];
        const numB = parseInt(b.match(/\d+/) || [0])[0];
        return numA - numB;
    });

    console.log(`✅ 找到 ${files.length} 個 HTML 檔案，開啟瀏覽器...`);

    const browser = await puppeteer.launch({
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    // 負責處理單一頁面的任務函數
    const processPage = async (file) => {
        const page = await browser.newPage();
        await page.setViewport({ width: 1280, height: 720 });
        const filePath = path.join(htmlDir, file);

        try {
            await page.goto(`file://${filePath}`, { waitUntil: 'networkidle0', timeout: 30000 });
            await new Promise(r => setTimeout(r, 1000)); // 等待動畫

            // 截圖
            const screenshot = await page.screenshot({ encoding: 'base64' });

            // 提取文字邏輯 (讀取所有的 p, h1, h2, h3, li, div 文字)
            let extractedText = await page.evaluate(() => {
                // 找出所有可能有意義的文字節點
                const elements = document.querySelectorAll('h1, h2, h3, h4, h5, h6, p, li, td, th');
                let texts = Array.from(elements)
                    .map(el => el.innerText.trim())
                    .filter(text => text.length > 0);

                // 如果找不到特定標籤，就回傳整個 body 的文字
                if (texts.length === 0) {
                    texts = [document.body.innerText.trim()];
                }

                // 移除重複並組合
                return [...new Set(texts)].join('\n\n').substring(0, 5000); // 限制文字長度
            });

            await page.close();
            console.log(`- 已截圖並提取文字: ${file}`);
            return {
                file,
                screenshot,
                text: extractedText
            };
        } catch (error) {
            console.error(`❌ 處理 ${file} 時發生錯誤:`, error.message);
            await page.close();
            return { file, screenshot: null, text: '' };
        }
    };

    console.log(`[Info] 開始併發處理網頁...`);
    // 將所有檔案包裝成任務，交由 limiter 控制最大併發數量 (例如 5)
    const MAX_CONCURRENCY = 5;
    const tasks = files.map(file => () => processPage(file));

    const results = await limitConcurrency(tasks, MAX_CONCURRENCY);

    console.log(`[Info] 網頁處理完畢，開始產生 PPTX...`);
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE'; // 16:9

    for (const result of results) {
        if (!result.screenshot) continue; // 跳過錯誤的頁面

        // 第一張投影片放截圖
        const slide = pptx.addSlide();
        slide.addImage({ data: `image/png;base64,${result.screenshot}`, x: 0, y: 0, w: '100%', h: '100%' });

        // 若有提取到文字，新增一頁純文字摘要（方便複製）
        if (result.text) {
            const textSlide = pptx.addSlide();
            textSlide.addText(`文字擷取結果：${result.file}`, { x: 0.5, y: 0.5, w: 9, h: 0.5, fontSize: 18, bold: true });
            textSlide.addText(result.text, { x: 0.5, y: 1.2, w: 12, h: 5.5, fontSize: 12, valign: "top" });
        }
    }

    try {
        await pptx.writeFile({ fileName: outputPath });
        console.log(`\n🎉 轉換成功！產出檔案：${outputPath}`);
    } catch (writeErr) {
        console.error(`❌ 儲存 PPTX 檔案時發生錯誤:`, writeErr);
    }

    await browser.close();
}

run().catch(err => { console.error(err); process.exit(1); });
