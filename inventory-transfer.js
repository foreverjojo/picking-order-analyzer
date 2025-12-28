/**
 * @file inventory-transfer.js
 * @description 庫存轉移工具 - 核心邏輯
 * @purpose 將生產統計表的數據從工作表 2 映射並填入工作表 1，實現 Excel 資料自動轉移
 * @author Ivy House TW Development Team
 */

// ==================== 庫存轉移工具核心邏輯 ====================

// 工作表 2 產品名稱到工作表 1 產品名稱的映射
const productNameMapping = {
    // 塔類 - 豆塔
    'A1蔓越莓': '豆塔-蔓越莓',
    'A2焦糖': '豆塔-焦糖',
    'A3巧克力': '豆塔-巧克力',
    'A4抹茶': '豆塔-抹茶',
    'A5椒麻': '豆塔-椒麻',
    // 塔類 - 堅果塔
    'B1蜂蜜': '堅果塔-蜂蜜',
    'B2焦糖': '堅果塔-焦糖',
    'B3巧克力': '堅果塔-巧克力',
    'B4海苔': '堅果塔-海苔',
    'B5咖哩': '堅果塔-咖哩',
    // 酥餅
    'P1土鳳梨酥': '土鳳梨酥(紅點)',
    'P2鳳凰酥': '鳳凰酥',
    // 雪花餅
    'H1蔓越莓': '雪花餅-蔓越莓',
    'H2巧克力': '雪花餅-巧克力',
    'H3金沙': '雪花餅-金沙',
    'H4抹茶': '雪花餅-抹茶',
    'H5肉鬆': '雪花餅-肉鬆',
    // 瑪德蓮
    'M1蜂蜜': '瑪德蓮-蜂蜜',
    'M2巧克力': '瑪德蓮-巧克力',
    'M3紅茶': '瑪德蓮-紅茶',
    'M4抹茶': '瑪德蓮-抹茶',
    'M5柑橘': '瑪德蓮-柑橘',
    'M6檸檬': '瑪德蓮-檸檬',
    // 堅果類 - 注意：不在此處設置映射，由動態邏輯處理 (來源名 -> 無調味+名)
    // 椰棗類 - 注意：不在此處設置映射，由動態邏輯處理 (來源名 -> ★中東椰棗300g 或 名+150g)
    // 糖果類
    'S1牛奶糖': '牛奶糖',
    'S2核桃糕': '南棗核桃糕',
    '牛奶糖-50g': '牛奶糖-50g',
    '牛奶糖50g': '牛奶糖-50g',
    '牛奶糖': '牛奶糖',
    '南棗核桃糕': '南棗核桃糕',
    // 瓦片類
    '瓦片-綜合': '瓦片-綜合',
    '瓦片-原味': '瓦片-原味',
    '瓦片-紅茶': '瓦片-紅茶',
    '瓦片-巧克力': '瓦片-巧克力',
    '瓦片-海苔': '瓦片-海苔',
    '瓦片-抹茶': '瓦片-抹茶',
    '瓦片-黑糖': '瓦片-黑糖',
    '瓦片-青花椒': '瓦片-青花椒',
    '瓦片-原味45克': '瓦片-原味45克',
    '綜合瓦片': '瓦片-綜合',
    // 注意：曖昧名稱 (原味、紅茶、巧克力、海苔、抹茶、黑糖、青花椒) 不在此處設置映射
    // 由動態邏輯根據區段決定 (瓦片類 -> 瓦片-X, 奶油餅乾 -> 奶油-X)
    '*原味45g': '瓦片-原味45克',
    // 奶油餅乾
    '奶油-焦糖牛奶': '奶油-焦糖牛奶',
    '奶油-法國巧克力': '奶油-法國巧克力',
    '奶油-蜂蜜檸檬': '奶油-蜂蜜檸檬',
    '奶油-伯爵紅茶': '奶油-伯爵紅茶',
    '奶油-抹茶': '奶油-抹茶',
    '焦糖牛奶': '奶油-焦糖牛奶',
    '法國巧克力': '奶油-法國巧克力',
    '蜂蜜檸檬': '奶油-蜂蜜檸檬',
    '伯爵紅茶': '奶油-伯爵紅茶',
    // 西餅餅乾
    '西點-綜合': '西點-綜合',
    '西點-巧克力貝殼': '西點-巧克力貝殼',
    '西點-咖啡小花': '西點-咖啡小花',
    '西點-藍莓小花': '西點-藍莓小花',
    '西點-蔓越莓貝殼': '西點-蔓越莓貝殼',
    '西點-乳酪酥條': '西點-乳酪酥條',
    '綜合西點': '西點-綜合',
    '巧克力貝殼': '西點-巧克力貝殼',
    '咖啡小花': '西點-咖啡小花',
    '藍莓小花': '西點-藍莓小花',
    '蔓越莓貝殼': '西點-蔓越莓貝殼',
    '乳酪酥條': '西點-乳酪酥條',
    // 千層餅乾
    '千層餅乾': '千層餅乾',
    '千層-小酥條': '千層-小酥條'
};

// 類別欄位映射規則
// key: 來源類別, value: { sourceSmall, sourceLarge, targetSmall, targetLarge }
const categoryColumnMapping = {
    '塔類': { sourceCol: 'B', targetCol: 'G' },
    '酥餅': { sourceCol: 'B', targetCol: 'G' },
    '雪花餅': { sourceCol: 'B', targetCol: 'G' },
    '瑪德蓮': { sourceCol: 'B', targetCol: 'G' },
    '堅果類': { sourceSmall: 'B', sourceLarge: 'C', targetSmall: 'E', targetLarge: 'F' },
    '椰棗類': { sourceCol: 'D', targetCol: 'E' },
    '糖果類': { sourceSmall: 'D', sourceLarge: 'E', targetSmall: 'E', targetLarge: 'F' },
    '瓦片類': { sourceSmall: 'E', sourceLarge: 'F', targetSmall: 'E', targetLarge: 'F' },
    '奶油餅乾': { sourceSmall: 'E', sourceLarge: 'F', targetSmall: 'E', targetLarge: 'F' },
    '西餅餅乾': { sourceSmall: 'E', sourceLarge: 'F', targetSmall: 'E', targetLarge: 'F' },
    '千層餅乾': { sourceSmall: 'E', sourceLarge: 'F', targetSmall: 'E', targetLarge: 'F' }
};

// 產品到類別的映射
function getProductCategory(productName) {
    const name = productName.replace(/\s+/g, '');
    if (/^A\d|豆塔/.test(name)) return '塔類';
    if (/^B\d|堅果塔/.test(name)) return '塔類';
    if (/^P\d|鳳梨酥|鳳凰酥/.test(name)) return '酥餅';
    if (/^H\d|雪花餅/.test(name)) return '雪花餅';
    if (/^M\d|瑪德蓮/.test(name)) return '瑪德蓮';
    // 糖果類要在堅果類之前檢測（避免「核桃糕」被當成「核桃」）
    if (/^S\d|牛奶糖|核桃糕|南棗/.test(name)) return '糖果類';
    if (/椰棗/.test(name)) return '椰棗類';
    if (/綜合堅果|夏威夷豆|腰果|杏仁|核桃/.test(name) && !/椰棗|糕/.test(name)) return '堅果類';
    if (/瓦片|原味|紅茶|海苔|抹茶|黑糖|青花椒/.test(name) && !/餅乾|西點|奶油/.test(name)) return '瓦片類';
    if (/奶油餅乾|焦糖牛奶|法國巧克力|蜂蜜檸檬|伯爵紅茶/.test(name) && !/牛奶糖/.test(name)) return '奶油餅乾';
    if (/西點|西餅|貝殼|小花|酥條/.test(name)) return '西餅餅乾';
    if (/千層/.test(name)) return '千層餅乾';
    return null;
}

// 全局變數
let yesterdayWorkbook = null;
let todayWorkbook = null;
let todayFileBuffer = null;
let transferResults = [];

// DOM 元素
document.addEventListener('DOMContentLoaded', () => {
    const yesterdayFileInput = document.getElementById('yesterdayFile');
    const todayFileInput = document.getElementById('todayFile');
    const yesterdayBox = document.getElementById('yesterdayBox');
    const todayBox = document.getElementById('todayBox');
    const transferBtn = document.getElementById('transferBtn');
    const downloadBtn = document.getElementById('downloadBtn');

    // 檔案輸入事件
    yesterdayFileInput.addEventListener('change', handleYesterdayFile);
    todayFileInput.addEventListener('change', handleTodayFile);
    transferBtn.addEventListener('click', performTransfer);
    downloadBtn.addEventListener('click', downloadResult);

    // 點擊上傳區域觸發檔案選擇
    yesterdayBox.addEventListener('click', () => yesterdayFileInput.click());
    todayBox.addEventListener('click', () => todayFileInput.click());

    // 拖放功能 - 本日報表
    setupDragDrop(yesterdayBox, async (file) => {
        console.log('開始讀取本日報表:', file.name);
        document.getElementById('yesterdayFileName').textContent = file.name;
        yesterdayBox.classList.add('has-file');

        try {
            const arrayBuffer = await file.arrayBuffer();
            yesterdayWorkbook = new ExcelJS.Workbook();
            await yesterdayWorkbook.xlsx.load(arrayBuffer);
            console.log('本日報表已載入:', yesterdayWorkbook.worksheets.map(ws => ws.name));
            checkReadyToTransfer();
        } catch (error) {
            console.error('讀取本日報表失敗:', error);
            alert('讀取本日報表失敗，請檢查檔案是否損毀');
        }
    });

    // 拖放功能 - 報表原始檔
    setupDragDrop(todayBox, async (file) => {
        console.log('開始讀取報表原始檔:', file.name);
        document.getElementById('todayFileName').textContent = file.name;
        todayBox.classList.add('has-file');

        try {
            const arrayBuffer = await file.arrayBuffer();
            todayFileBuffer = new Uint8Array(arrayBuffer);
            todayWorkbook = new ExcelJS.Workbook();
            await todayWorkbook.xlsx.load(arrayBuffer.slice(0));
            console.log('報表原始檔已載入:', todayWorkbook.worksheets.map(ws => ws.name));
            checkReadyToTransfer();
        } catch (error) {
            console.error('讀取報表原始檔失敗:', error);
            alert('讀取報表原始檔失敗，請檢查檔案是否損毀');
        }
    });
});

// 設置拖放功能
function setupDragDrop(element, onFile) {
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        element.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
        });
    });

    ['dragenter', 'dragover'].forEach(eventName => {
        element.addEventListener(eventName, () => {
            element.classList.add('dragover');
        });
    });

    ['dragleave', 'drop'].forEach(eventName => {
        element.addEventListener(eventName, () => {
            element.classList.remove('dragover');
        });
    });

    element.addEventListener('drop', (e) => {
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            onFile(files[0]);
        }
    });
}

async function handleYesterdayFile(e) {
    try {
        const file = e.target.files[0];
        if (!file) return;

        console.log('開始讀取昨天報表:', file.name);
        document.getElementById('yesterdayFileName').textContent = file.name;
        document.getElementById('yesterdayBox').classList.add('has-file');

        const arrayBuffer = await file.arrayBuffer();
        yesterdayWorkbook = new ExcelJS.Workbook();
        await yesterdayWorkbook.xlsx.load(arrayBuffer);

        console.log('昨天報表已載入:', yesterdayWorkbook.worksheets.map(ws => ws.name));
        checkReadyToTransfer();
    } catch (error) {
        console.error('讀取昨天報表失敗:', error);
        alert('讀取昨天報表失敗，請檢查檔案是否損毀 (詳細錯誤請看 Console)');
    }
}

async function handleTodayFile(e) {
    try {
        const file = e.target.files[0];
        if (!file) return;

        console.log('開始讀取今天報表:', file.name);
        document.getElementById('todayFileName').textContent = file.name;
        document.getElementById('todayBox').classList.add('has-file');

        // 儲存原始檔案的完整副本（用於 JSZip 輸出）
        const arrayBuffer = await file.arrayBuffer();
        todayFileBuffer = new Uint8Array(arrayBuffer);

        // 讀取給 ExcelJS 使用（需要另一個副本，因為 ExcelJS 可能會修改 buffer）
        todayWorkbook = new ExcelJS.Workbook();
        await todayWorkbook.xlsx.load(arrayBuffer.slice(0));

        console.log('今天報表已載入:', todayWorkbook.worksheets.map(ws => ws.name));
        checkReadyToTransfer();
    } catch (error) {
        console.error('讀取今天報表失敗:', error);
        alert('讀取今天報表失敗，請檢查檔案是否損毀 (詳細錯誤請看 Console)');
    }
}

function checkReadyToTransfer() {
    const transferBtn = document.getElementById('transferBtn');
    const ready = !!(yesterdayWorkbook && todayWorkbook);
    transferBtn.disabled = !ready;
    console.log(`檢查轉移狀態: 昨天=${!!yesterdayWorkbook}, 今天=${!!todayWorkbook} => 按鈕${ready ? '啟用' : '停用'}`);
}

function colLetterToNumber(letter) {
    return letter.charCodeAt(0) - 64;
}

// 處理公式儲存格，提取實際數值
function getCellValue(cell) {
    const value = cell.value;
    if (value === null || value === undefined) return null;

    // 如果是物件（公式儲存格），提取 result
    if (typeof value === 'object') {
        // ExcelJS 公式儲存格格式: { formula: '...', result: 123 }
        if (value.result !== undefined) {
            return value.result;
        }
        // 富文本格式: { richText: [...] }
        if (value.richText) {
            return value.richText.map(r => r.text).join('');
        }
        // 其他物件類型，嘗試轉換
        return null;
    }

    return value;
}

async function performTransfer() {
    transferResults = [];

    // 讀取昨天報表的工作表 2
    const sheet2 = yesterdayWorkbook.worksheets[1];
    if (!sheet2) {
        alert('昨天的報表沒有工作表 2');
        return;
    }

    // 讀取今天報表的工作表 1
    const sheet1 = todayWorkbook.worksheets[0];
    if (!sheet1) {
        alert('今天的報表沒有工作表 1');
        return;
    }

    // 建立今天報表的產品名稱到行號映射
    const todayProductRows = new Map();
    sheet1.eachRow((row, rowNumber) => {
        const cellA = row.getCell(1).value;
        if (cellA) {
            const name = String(cellA).trim();
            todayProductRows.set(name, rowNumber);
        }
    });

    console.log('今天報表產品列表:', Array.from(todayProductRows.keys()));

    // DEBUG: 輸出工作表 2 的前 50 行內容
    console.log('=== 工作表 2 原始內容 ===');
    for (let i = 1; i <= Math.min(50, sheet2.rowCount); i++) {
        const row = sheet2.getRow(i);
        const colA = getCellValue(row.getCell(1));
        const colB = getCellValue(row.getCell(2));
        const colC = getCellValue(row.getCell(3));
        if (colA || colB || colC) {
            console.log(`Row ${i}: A="${colA}", B="${colB}", C="${colC}"`);
        }
    }
    console.log('=========================');

    // 區段追蹤變數 - 用於處理無法單獨判斷類別的產品名稱
    let currentSection = null;

    // 區段標題到類別的映射
    const sectionHeaders = {
        '杏仁瓦片': '瓦片類',
        '瓦片類': '瓦片類',
        '奶油餅乾': '奶油餅乾',
        '西餅餅乾': '西餅餅乾',
        '堅果類': '堅果類',
        '椰棗類': '椰棗類',
        '糖果': '糖果類',      // 注意：console 顯示的是 "糖果" 不是 "糖果類"
        '糖果類': '糖果類'
    };

    // 需要區段上下文才能判斷類別的產品名稱
    const ambiguousNames = new Set([
        '原味', '紅茶', '巧克力', '海苔', '抹茶', '黑糖', '青花椒',
        '伯爵紅茶',  // 這個在奶油餅乾區段
        '綜合堅果', '夏威夷豆', '腰果', '杏仁', '核桃',
        '中東椰棗', '椰棗豆子', '椰棗腰果', '椰棗杏仁', '椰棗核桃',
        '綜合瓦片'  // 瓦片區段的第一個產品
    ]);

    // 遍歷昨天報表工作表 2 的每一行
    sheet2.eachRow((row, rowNumber) => {
        const cellA = getCellValue(row.getCell(1));
        if (!cellA) return;

        const sourceName = String(cellA).trim();

        // 檢查是否為區段標題行
        if (sectionHeaders[sourceName]) {
            currentSection = sectionHeaders[sourceName];
            console.log(`[區段切換] Row ${rowNumber}: 進入「${currentSection}」區段`);
            // 區段標題本身不處理，繼續下一行
            return;
        }

        // 跳過標題行（含有「*庫存」的值）
        const cellB = getCellValue(row.getCell(2));
        if (cellB && String(cellB).includes('*庫存')) {
            console.log(`跳過標題行 ${rowNumber}: ${sourceName}`);
            // 但仍然嘗試識別區段
            if (sectionHeaders[sourceName]) {
                currentSection = sectionHeaders[sourceName];
            }
            return;
        }

        // 判斷類別：對於曖昧名稱，區段追蹤優先於 getProductCategory
        let category;

        // 如果是曖昧名稱且有區段上下文，直接使用區段
        if (ambiguousNames.has(sourceName) && currentSection) {
            category = currentSection;
            console.log(`Row ${rowNumber}: "${sourceName}" -> 使用區段推斷類別: ${category}`);
        } else {
            category = getProductCategory(sourceName);
            console.log(`Row ${rowNumber}: "${sourceName}" -> 類別: ${category || '無'}`);
        }

        if (!category) return;

        const mapping = categoryColumnMapping[category];
        if (!mapping) {
            console.log(`  無欄位映射規則: ${category}`);
            return;
        }

        // 找到今天報表中對應的產品名稱
        // 對於曖昧名稱，需要根據區段來決定正確的目標名稱
        let targetName = productNameMapping[sourceName];

        // 如果沒有直接映射，且是曖昧名稱，根據區段決定目標名稱
        if (!targetName && ambiguousNames.has(sourceName)) {
            if (category === '瓦片類') {
                targetName = '瓦片-' + sourceName;
            } else if (category === '奶油餅乾') {
                targetName = '奶油-' + sourceName;
            } else if (category === '堅果類') {
                targetName = '無調味' + sourceName;
            } else if (category === '椰棗類') {
                // 椰棗類產品映射到今日報表格式
                if (sourceName === '中東椰棗') {
                    targetName = '★中東椰棗300g';
                } else {
                    targetName = sourceName + '150g';
                }
            } else {
                targetName = sourceName;
            }
        } else if (!targetName) {
            targetName = sourceName;
        }

        const targetRow = todayProductRows.get(targetName);

        console.log(`  映射: ${sourceName} -> ${targetName}, 目標行: ${targetRow || '未找到'}`);

        if (!targetRow) {
            console.log(`未找到匹配: ${sourceName} -> ${targetName}`);
            return;
        }

        // 根據類別處理欄位映射
        if (mapping.sourceCol) {
            // 單一欄位映射
            const sourceValue = getCellValue(row.getCell(colLetterToNumber(mapping.sourceCol)));
            if (sourceValue !== null && sourceValue !== undefined) {
                const targetColNum = colLetterToNumber(mapping.targetCol);
                sheet1.getRow(targetRow).getCell(targetColNum).value = sourceValue;

                transferResults.push({
                    sourceName,
                    targetName,
                    targetRow,
                    targetCol: mapping.targetCol,
                    value: sourceValue,
                    matched: true
                });
            }
        } else {
            // 大小包雙欄位映射
            const smallValue = getCellValue(row.getCell(colLetterToNumber(mapping.sourceSmall)));
            const largeValue = getCellValue(row.getCell(colLetterToNumber(mapping.sourceLarge)));

            if (smallValue !== null && smallValue !== undefined) {
                const targetColNum = colLetterToNumber(mapping.targetSmall);
                sheet1.getRow(targetRow).getCell(targetColNum).value = smallValue;

                transferResults.push({
                    sourceName,
                    targetName,
                    targetRow,
                    targetCol: mapping.targetSmall,
                    value: smallValue,
                    type: '小包',
                    matched: true
                });
            }

            if (largeValue !== null && largeValue !== undefined) {
                const targetColNum = colLetterToNumber(mapping.targetLarge);
                sheet1.getRow(targetRow).getCell(targetColNum).value = largeValue;

                transferResults.push({
                    sourceName,
                    targetName,
                    targetRow,
                    targetCol: mapping.targetLarge,
                    value: largeValue,
                    type: '大包',
                    matched: true
                });
            }
        }
    });

    // === 工作表 2 到工作表 2 的直接欄位映射 ===
    // 昨天工作表 2 的 C 欄 -> 今天工作表 2 的 B 欄 (同一行號)
    const todaySheet2 = todayWorkbook.worksheets[1];
    if (todaySheet2) {
        const sheet2DirectMappings = [
            37, 38, 39, 40, 41,  // 椰棗類
            43, 44,              // 糖果類
            48, 49, 50, 51, 52, 53, 54,  // 瓦片類
            57, 58, 59, 60, 61,  // 奶油餅乾
            64, 65, 66, 67, 68   // 西餅餅乾
        ];

        console.log('=== 工作表 2 直接欄位映射 ===');
        sheet2DirectMappings.forEach(rowNum => {
            const sourceValue = getCellValue(sheet2.getRow(rowNum).getCell(3)); // C 欄 = 第 3 欄
            if (sourceValue !== null && sourceValue !== undefined) {
                todaySheet2.getRow(rowNum).getCell(2).value = sourceValue; // B 欄 = 第 2 欄

                transferResults.push({
                    sourceName: `工作表2 C${rowNum}`,
                    targetName: `工作表2 B${rowNum}`,
                    targetRow: rowNum,
                    targetCol: 'B',
                    value: sourceValue,
                    type: '直接映射',
                    matched: true
                });

                console.log(`[工作表2] C${rowNum} -> B${rowNum}: ${sourceValue}`);
            }
        });
        console.log('=========================');
    } else {
        console.log('今天報表沒有工作表 2，跳過直接映射');
    }

    // === 工作表 2 D 欄到工作表 1 N 欄的跨工作表映射 ===
    // 本日報表工作表 2 的 D 欄 -> 報表原始檔工作表 1 的 N 欄
    console.log('=== 工作表 2 D欄 -> 工作表 1 N欄 跨表映射 ===');
    const crossSheetMappings = [
        // 瓦片類
        { sourceRow: 48, targetRow: 12 },  // D48 -> N12
        { sourceRow: 49, targetRow: 13 },  // D49 -> N13
        { sourceRow: 50, targetRow: 14 },  // D50 -> N14
        { sourceRow: 51, targetRow: 22 },  // D51 -> N22
        { sourceRow: 52, targetRow: 23 },  // D52 -> N23
        { sourceRow: 53, targetRow: 24 },  // D53 -> N24
        { sourceRow: 54, targetRow: 25 },  // D54 -> N25
        // 奶油餅乾
        { sourceRow: 57, targetRow: 15 },  // D57 -> N15
        { sourceRow: 58, targetRow: 16 },  // D58 -> N16
        { sourceRow: 59, targetRow: 17 },  // D59 -> N17
        { sourceRow: 60, targetRow: 26 },  // D60 -> N26
        { sourceRow: 61, targetRow: 27 },  // D61 -> N27
        // 西餅餅乾
        { sourceRow: 64, targetRow: 18 },  // D64 -> N18
        { sourceRow: 65, targetRow: 19 },  // D65 -> N19
        { sourceRow: 66, targetRow: 20 },  // D66 -> N20
        { sourceRow: 67, targetRow: 28 },  // D67 -> N28
        { sourceRow: 68, targetRow: 29 },  // D68 -> N29
    ];

    crossSheetMappings.forEach(mapping => {
        const sourceValue = getCellValue(sheet2.getRow(mapping.sourceRow).getCell(4)); // D 欄 = 第 4 欄
        if (sourceValue !== null && sourceValue !== undefined) {
            // 寫入工作表 1 的 N 欄
            sheet1.getRow(mapping.targetRow).getCell(14).value = sourceValue; // N 欄 = 第 14 欄

            transferResults.push({
                sourceName: `工作表2 D${mapping.sourceRow}`,
                targetName: `工作表1 N${mapping.targetRow}`,
                targetRow: mapping.targetRow,
                targetCol: 'N',
                value: sourceValue,
                type: '跨表映射',
                matched: true
            });

            console.log(`[跨表] D${mapping.sourceRow} -> N${mapping.targetRow}: ${sourceValue}`);
        }
    });
    console.log('=========================');

    // 顯示結果
    showResults();
}

function showResults() {
    const resultCard = document.getElementById('resultCard');
    resultCard.style.display = 'block';

    const matched = transferResults.filter(r => r.matched).length;
    document.getElementById('totalMatched').textContent = matched;
    document.getElementById('totalUnmatched').textContent = 0;

    // 建立預覽表格
    let html = '<table class="preview-table">';
    html += '<tr><th>來源產品</th><th>目標產品</th><th>目標儲存格</th><th>庫存值</th></tr>';

    transferResults.forEach(result => {
        const typeLabel = result.type ? ` (${result.type})` : '';
        html += `<tr>
            <td>${result.sourceName}${typeLabel}</td>
            <td class="match-success">${result.targetName}</td>
            <td>${result.targetCol}${result.targetRow}</td>
            <td>${result.value}</td>
        </tr>`;
    });

    html += '</table>';
    document.getElementById('previewContainer').innerHTML = html;

    console.log('轉移完成，共 ' + matched + ' 項');
}

async function downloadResult() {
    try {
        // 使用 JSZip 直接修改 XML，保留 .xlsm 格式和巨集
        const zip = await JSZip.loadAsync(todayFileBuffer);

        // 讀取 sheet1.xml
        // 注意：這裡假設第一個工作表總是 sheet1.xml。這在大多數情況下是正確的。
        const sheetPath = 'xl/worksheets/sheet1.xml';
        if (!zip.file(sheetPath)) {
            alert('錯誤：找不到 sheet1.xml，無法寫入數據。請確認範本格式。');
            return;
        }

        let sheetXml = await zip.file(sheetPath).async('string');

        // 將 transferResults 中的數據寫入 XML (只處理工作表 1 的映射)
        let updatedCount = 0;

        // 過濾出工作表 1 的結果（排除 type='直接映射' 的工作表 2 項目）
        const sheet1Results = transferResults.filter(r => r.type !== '直接映射');

        sheet1Results.forEach(result => {
            const cellRef = result.targetCol + result.targetRow;
            const value = result.value;

            if (value === null || value === undefined) return;

            // --- 手術刀式字串替換策略 ---
            // 直接在字串中定位 <c ... r="CELL_REF" ...>，準確找出標籤範圍並替換

            // 1. 定位 r="CELL_REF"
            const refStr = `r="${cellRef}"`;
            const refIdx = sheetXml.indexOf(refStr);

            if (refIdx === -1) {
                console.log(`[XML] 找不到儲存格 ${cellRef} (可能該行無此欄位)`);
                return;
            }

            // 2. 往回找標籤開頭 <c
            const tagStart = sheetXml.lastIndexOf('<c', refIdx);
            if (tagStart === -1) {
                console.error(`[XML] 異常: 找到 ${refStr} 但找不到開頭 <c`);
                return;
            }

            // 3. 判斷標籤結束位置
            const tagInnerEnd = sheetXml.indexOf('>', tagStart);
            if (tagInnerEnd === -1) return;

            let tagEnd = -1;

            // 檢查是否為 self-closing (/>)
            if (sheetXml[tagInnerEnd - 1] === '/') {
                // <c r="G5" ... />
                tagEnd = tagInnerEnd + 1;
            } else {
                // <c r="G5" ...>...</c>
                const closeTag = '</c>';
                const closeTagIdx = sheetXml.indexOf(closeTag, tagInnerEnd);
                if (closeTagIdx !== -1) {
                    tagEnd = closeTagIdx + closeTag.length;
                } else {
                    console.error(`[XML] 異常: 儲存格 ${cellRef} 是展開的但找不到 </c>`);
                    return;
                }
            }

            const originalTag = sheetXml.substring(tagStart, tagEnd);

            // 4. 解析原有 Style (s="...")
            const styleMatch = originalTag.match(/ s="([0-9]+)"/);
            const styleAttr = styleMatch ? ` s="${styleMatch[1]}"` : '';

            // 5. 構建新標籤
            // 始終展開為標準格式: <c r="..." s="..."> <v>...</v> </c>
            // 這會自動移除 t="s" (如有)，將其變為預設數值型態
            const newTag = `<c r="${cellRef}"${styleAttr}><v>${value}</v></c>`;

            // 6. 執行字串替換
            sheetXml = sheetXml.substring(0, tagStart) + newTag + sheetXml.substring(tagEnd);

            updatedCount++;
            console.log(`[XML] 更新 ${cellRef}`);
        });

        console.log(`XML 更新完成，共修改 ${updatedCount} 個儲存格 (工作表1)`);

        // 寫回 ZIP
        zip.file(sheetPath, sheetXml);

        // === 處理工作表 2 ===
        const sheet2Path = 'xl/worksheets/sheet2.xml';
        if (zip.file(sheet2Path)) {
            let sheet2Xml = await zip.file(sheet2Path).async('string');
            let sheet2UpdatedCount = 0;

            // 篩選出工作表 2 的結果（type 為 '直接映射' 的項目）
            const sheet2Results = transferResults.filter(r => r.type === '直接映射');

            sheet2Results.forEach(result => {
                const cellRef = result.targetCol + result.targetRow;
                const value = result.value;

                if (value === null || value === undefined) return;

                // 使用相同的手術刀式字串替換策略
                const refStr = `r="${cellRef}"`;
                const refIdx = sheet2Xml.indexOf(refStr);

                if (refIdx === -1) {
                    console.log(`[工作表2 XML] 找不到儲存格 ${cellRef}`);
                    return;
                }

                const tagStart = sheet2Xml.lastIndexOf('<c', refIdx);
                if (tagStart === -1) return;

                const tagInnerEnd = sheet2Xml.indexOf('>', tagStart);
                if (tagInnerEnd === -1) return;

                let tagEnd = -1;
                if (sheet2Xml[tagInnerEnd - 1] === '/') {
                    tagEnd = tagInnerEnd + 1;
                } else {
                    const closeTagIdx = sheet2Xml.indexOf('</c>', tagInnerEnd);
                    if (closeTagIdx !== -1) {
                        tagEnd = closeTagIdx + 4;
                    } else {
                        return;
                    }
                }

                const originalTag = sheet2Xml.substring(tagStart, tagEnd);
                const styleMatch = originalTag.match(/ s="([0-9]+)"/);
                const styleAttr = styleMatch ? ` s="${styleMatch[1]}"` : '';
                const newTag = `<c r="${cellRef}"${styleAttr}><v>${value}</v></c>`;

                sheet2Xml = sheet2Xml.substring(0, tagStart) + newTag + sheet2Xml.substring(tagEnd);
                sheet2UpdatedCount++;
                console.log(`[工作表2 XML] 更新 ${cellRef}`);
            });

            console.log(`XML 更新完成，共修改 ${sheet2UpdatedCount} 個儲存格 (工作表2)`);
            zip.file(sheet2Path, sheet2Xml);
        }

        // === 強制 Excel 重新計算公式 ===
        // 1. 刪除 calcChain.xml
        const calcChainPath = 'xl/calcChain.xml';
        if (zip.file(calcChainPath)) {
            zip.remove(calcChainPath);
            console.log('已刪除 calcChain.xml');

            // 2. 從 [Content_Types].xml 中移除對 calcChain 的引用
            const contentTypesPath = '[Content_Types].xml';
            let contentTypesXml = await zip.file(contentTypesPath).async('string');
            contentTypesXml = contentTypesXml.replace(/<Override[^>]*calcChain[^>]*\/>/g, '');
            zip.file(contentTypesPath, contentTypesXml);
            console.log('已從 Content_Types.xml 移除 calcChain 引用');
        }

        // 3. 設定 workbook.xml 強制開啟時重新計算公式
        const workbookPath = 'xl/workbook.xml';
        let workbookXml = await zip.file(workbookPath).async('string');

        // 檢查是否已有 calcPr 元素
        if (workbookXml.includes('<calcPr')) {
            // 如果已有 fullCalcOnLoad，先移除它
            workbookXml = workbookXml.replace(/fullCalcOnLoad="[^"]*"/g, '');
            // 在 calcPr 開頭加入 fullCalcOnLoad="1"
            workbookXml = workbookXml.replace(/<calcPr\s*/g, '<calcPr fullCalcOnLoad="1" ');
        } else {
            // 在 </workbook> 前插入 calcPr
            workbookXml = workbookXml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
        }
        zip.file(workbookPath, workbookXml);
        console.log('已設定公式在開啟時自動重新計算');

        // 生成檔案
        const content = await zip.generateAsync({ type: 'blob' });
        const url = URL.createObjectURL(content);
        const a = document.createElement('a');
        a.href = url;
        const today = new Date().toISOString().split('T')[0].replace(/-/g, '');
        a.download = `生產統計表_庫存已轉移_${today}.xlsm`;
        a.click();
        URL.revokeObjectURL(url);

    } catch (error) {
        console.error('產生檔案時發生錯誤:', error);
        alert('產生檔案時發生錯誤，請查看 Console');
    }
}
