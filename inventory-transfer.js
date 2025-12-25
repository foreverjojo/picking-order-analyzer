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
    // 瑪德蓮
    'M1蜂蜜': '瑪德蓮-蜂蜜',
    'M2巧克力': '瑪德蓮-巧克力',
    'M3柑橘': '瑪德蓮-柑橘',
    'M4檸檬': '瑪德蓮-檸檬',
    'M5綜合': '瑪德蓮-綜合',
    // 堅果類
    '無調味綜合堅果': '無調味綜合堅果',
    '無調味夏威夷豆': '無調味夏威夷豆',
    '無調味腰果': '無調味腰果',
    '無調味杏仁': '無調味杏仁',
    '無調味核桃': '無調味核桃',
    '綜合堅果': '綜合堅果',
    '夏威夷豆': '夏威夷豆',
    '腰果': '腰果',
    '杏仁': '杏仁',
    '核桃': '核桃',
    // 椰棗類
    '★中東椰棗300g': '★中東椰棗300g',
    '椰棗豆子150g': '椰棗豆子150g',
    '椰棗腰果150g': '椰棗腰果150g',
    '椰棗杏仁150g': '椰棗杏仁150g',
    '椰棗核桃150g': '椰棗核桃150g',
    '中東椰棗': '中東椰棗',
    '椰棗豆子': '椰棗豆子',
    '椰棗腰果': '椰棗腰果',
    '椰棗杏仁': '椰棗杏仁',
    '椰棗核桃': '椰棗核桃',
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
    '原味': '瓦片-原味',
    '紅茶': '瓦片-紅茶',
    '巧克力': '瓦片-巧克力',
    '海苔': '瓦片-海苔',
    '抹茶': '瓦片-抹茶',
    '黑糖': '瓦片-黑糖',
    '青花椒': '瓦片-青花椒',
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
    const transferBtn = document.getElementById('transferBtn');
    const downloadBtn = document.getElementById('downloadBtn');

    yesterdayFileInput.addEventListener('change', handleYesterdayFile);
    todayFileInput.addEventListener('change', handleTodayFile);
    transferBtn.addEventListener('click', performTransfer);
    downloadBtn.addEventListener('click', downloadResult);
});

async function handleYesterdayFile(e) {
    const file = e.target.files[0];
    if (!file) return;

    document.getElementById('yesterdayFileName').textContent = file.name;
    document.getElementById('yesterdayBox').classList.add('has-file');

    const arrayBuffer = await file.arrayBuffer();
    yesterdayWorkbook = new ExcelJS.Workbook();
    await yesterdayWorkbook.xlsx.load(arrayBuffer);

    console.log('昨天報表已載入:', yesterdayWorkbook.worksheets.map(ws => ws.name));
    checkReadyToTransfer();
}

async function handleTodayFile(e) {
    const file = e.target.files[0];
    if (!file) return;

    document.getElementById('todayFileName').textContent = file.name;
    document.getElementById('todayBox').classList.add('has-file');

    todayFileBuffer = await file.arrayBuffer();
    todayWorkbook = new ExcelJS.Workbook();
    await todayWorkbook.xlsx.load(todayFileBuffer);

    console.log('今天報表已載入:', todayWorkbook.worksheets.map(ws => ws.name));
    checkReadyToTransfer();
}

function checkReadyToTransfer() {
    const transferBtn = document.getElementById('transferBtn');
    transferBtn.disabled = !(yesterdayWorkbook && todayWorkbook);
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

    // 遍歷昨天報表工作表 2 的每一行
    sheet2.eachRow((row, rowNumber) => {
        const cellA = row.getCell(1).value;
        if (!cellA) return;

        const sourceName = String(cellA).trim();

        // 跳過標題行（含有「*庫存」的值）
        const cellB = getCellValue(row.getCell(2));
        if (cellB && String(cellB).includes('*庫存')) {
            console.log(`跳過標題行 ${rowNumber}: ${sourceName}`);
            return;
        }

        const category = getProductCategory(sourceName);
        console.log(`Row ${rowNumber}: "${sourceName}" -> 類別: ${category || '無'}`);
        if (!category) return;

        const mapping = categoryColumnMapping[category];
        if (!mapping) {
            console.log(`  無欄位映射規則: ${category}`);
            return;
        }

        // 找到今天報表中對應的產品名稱
        const targetName = productNameMapping[sourceName] || sourceName;
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
    const buffer = await todayWorkbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const today = new Date().toISOString().split('T')[0].replace(/-/g, '');
    a.download = `生產統計表_庫存已轉移_${today}.xlsm`;
    a.click();
    URL.revokeObjectURL(url);
}
