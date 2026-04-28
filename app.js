/**
 * @file app.js
 * @description 撿貨單分析系統 - 核心應用程式
 * @purpose 處理多平台訂單 (蝦皮、MOMO、官網等) 的解析、商品映射、統計及報表生成
 * @author Ivy House TW Development Team
 */

// 全域變數
let uploadedFiles = {
    momo: null,
    official: null,
    shopee: null,
    orangepoint: null,  // 橘點子
    template: null
};

let parsedData = {
    momo: [],
    official: [],
    shopee: [],
    orangepoint: []  // 橘點子
};

let allProducts = []; // 所有解析出的商品
let productMapping = {}; // 商品對應關係
let statistics = {}; // 統計結果
let templateWorkbook = null; // 報表範本 (SheetJS)
let excelWorkbook = null; // 報表範本 (ExcelJS - 保留樣式)
let templateBuffer = null; // 報表範本原始二進制數據
let mappedProducts = []; // 映射後的商品（包含報表對應資訊）

// 注意：standardProductNames 和映射函數現在從 mapping-rules.js 載入

// PDF.js 設定
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';


// ==================== 初始化 ====================
document.addEventListener('DOMContentLoaded', () => {
    initializeEventListeners();
});

function initializeEventListeners() {
    const fileInput = document.getElementById('fileInput');
    const uploadBox = document.getElementById('uploadBox');
    const parseBtn = document.getElementById('parseBtn');
    const autoMapBtn = document.getElementById('autoMapBtn');
    const loadMappingBtn = document.getElementById('loadMappingBtn');
    const saveMappingBtn = document.getElementById('saveMappingBtn');
    const confirmMappingBtn = document.getElementById('confirmMappingBtn');
    const generateReportBtn = document.getElementById('generateReportBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    const resetBtn = document.getElementById('resetBtn');
    const addMoreFilesBtn = document.getElementById('addMoreFilesBtn');
    const additionalFileInput = document.getElementById('additionalFileInput');
    const templateInput = document.getElementById('templateInput');

    // 檔案上傳
    uploadBox.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', handleFileSelect);

    // 拖放上傳
    uploadBox.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadBox.classList.add('drag-over');
    });

    uploadBox.addEventListener('dragleave', () => {
        uploadBox.classList.remove('drag-over');
    });

    uploadBox.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadBox.classList.remove('drag-over');
        handleFileSelect({ target: { files: e.dataTransfer.files } });
    });

    // 步驟 2 上傳更多檔案
    if (addMoreFilesBtn) {
        addMoreFilesBtn.addEventListener('click', () => additionalFileInput.click());
        additionalFileInput.addEventListener('change', handleAdditionalFiles);
    }

    // 報表範本上傳
    if (templateInput) {
        templateInput.addEventListener('change', handleTemplateUpload);
    }

    // 按鈕事件
    parseBtn.addEventListener('click', parseAllFiles);
    autoMapBtn.addEventListener('click', autoMapProducts);
    loadMappingBtn.addEventListener('click', loadMappingRules);
    saveMappingBtn.addEventListener('click', saveMappingRules);
    confirmMappingBtn.addEventListener('click', promptTemplateUpload);
    generateReportBtn.addEventListener('click', generateReport);
    downloadBtn.addEventListener('click', downloadReport);
    resetBtn.addEventListener('click', resetApplication);

    // 新按鈕事件
    const skipToStep3Btn = document.getElementById('skipToStep3Btn');
    const downloadPickingListBtn = document.getElementById('downloadPickingListBtn');
    if (skipToStep3Btn) {
        skipToStep3Btn.addEventListener('click', skipToStep3);
    }
    if (downloadPickingListBtn) {
        downloadPickingListBtn.addEventListener('click', downloadPickingList);
    }
}

// ==================== 檔案處理 ====================
function handleFileSelect(event) {
    const files = Array.from(event.target.files || event.dataTransfer.files);

    files.forEach(file => {
        const fileName = file.name.toLowerCase();

        // 橘點子識別
        if (fileName.includes('橘點子') || fileName.includes('jellytree') || (fileName.endsWith('.xls') && !fileName.endsWith('.xlsx'))) {
            uploadedFiles.orangepoint = file;
            addFileToList(file, '橘點子撿貨單', '🍊');
        } else if ((fileName.includes('momo') || fileName.includes('富邦') || fileName.includes('order_export')) && fileName.endsWith('.xlsx')) {
            uploadedFiles.momo = file;
            addFileToList(file, 'MOMO 撿貨單', '📊');
        } else if (fileName.includes('官網') && fileName.endsWith('.xlsx')) {
            uploadedFiles.official = file;
            addFileToList(file, '官網撿貨單', '📊');
        } else if (fileName.endsWith('.pdf')) {
            uploadedFiles.shopee = file;
            addFileToList(file, '蝦皮撿貨單', '📄');
        } else if (fileName.includes('統計') || fileName.endsWith('.xlsm')) {
            uploadedFiles.template = file;
            addFileToList(file, '報表範本', '📋');
        } else if (fileName.endsWith('.xlsx')) {
            // 預設為官網
            uploadedFiles.official = file;
            addFileToList(file, '官網撿貨單', '📊');
        } else {
            alert('未知格式或無法識別平台: ' + file.name);
        }
    });

    updateParseButton();
}

function addFileToList(file, label, icon) {
    const fileList = document.getElementById('fileList');

    // 檢查是否已存在
    const existingItem = document.querySelector(`[data-file-label="${label}"]`);
    if (existingItem) {
        existingItem.remove();
    }

    const fileItem = document.createElement('div');
    fileItem.className = 'file-item';
    fileItem.setAttribute('data-file-label', label);

    fileItem.innerHTML = `
        <div class="file-info">
            <div class="file-icon">${icon}</div>
            <div class="file-details">
                <h4>${label}</h4>
                <span class="file-size">${file.name} (${formatFileSize(file.size)})</span>
            </div>
        </div>
        <button class="file-remove" onclick="removeFile('${label}')">✕ 移除</button>
    `;

    fileList.appendChild(fileItem);
}

function removeFile(label) {
    const fileItem = document.querySelector(`[data-file-label="${label}"]`);
    if (fileItem) {
        fileItem.remove();
    }

    // 清除對應的檔案
    if (label === 'MOMO 撿貨單') uploadedFiles.momo = null;
    else if (label === '官網撿貨單') uploadedFiles.official = null;
    else if (label === '蝦皮撿貨單') uploadedFiles.shopee = null;
    else if (label === '橘點子撿貨單') uploadedFiles.orangepoint = null;

    updateParseButton();
}

function updateParseButton() {
    const parseBtn = document.getElementById('parseBtn');
    const hasAnyPickingOrder = uploadedFiles.momo || uploadedFiles.official || uploadedFiles.shopee || uploadedFiles.orangepoint;
    parseBtn.disabled = !hasAnyPickingOrder;
}

// 步驟 2 上傳更多撿貨單
async function handleAdditionalFiles(event) {
    const files = event.target.files;
    if (!files.length) return;

    showLoading('解析中...');

    for (const file of files) {
        await handleSingleFile(file);
    }

    // 重新整合並顯示
    consolidateProducts();
    buildMappingTable();
    hideLoading();
    showToast('已添加新的撿貨單', 'success');
}

// 處理單個檔案（步驟 2 追加）
async function handleSingleFile(file) {
    const fileName = file.name.toLowerCase();

    if (fileName.includes('momo') || fileName.includes('order_export')) {
        const result = await parseMomoExcel(file);
        if (result && result.length > 0) {
            parsedData.momo = [...parsedData.momo, ...result];
            console.log('追加 MOMO 解析結果:', result.length, '筆');
        }
    } else if (fileName.includes('官網')) {
        const result = await parseOfficialExcel(file);
        if (result && result.length > 0) {
            parsedData.official = [...parsedData.official, ...result];
            console.log('追加官網解析結果:', result.length, '筆');
        }
    } else if (fileName.endsWith('.pdf')) {
        const result = await parseShopeePDF(file);
        if (result && result.length > 0) {
            parsedData.shopee = [...parsedData.shopee, ...result];
            console.log('追加蝦皮解析結果:', result.length, '筆');
        }
    } else if (fileName.includes('橘點子') || fileName.includes('jellytree') ||
        (fileName.endsWith('.xls') && !fileName.endsWith('.xlsx'))) {
        const result = await parseOrangePointExcel(file);
        if (result && result.length > 0) {
            parsedData.orangepoint = [...parsedData.orangepoint, ...result];
            console.log('追加橘點子解析結果:', result.length, '筆');
        }
    } else if (fileName.endsWith('.xlsx')) {
        // 預設為官網
        const result = await parseOfficialExcel(file);
        if (result && result.length > 0) {
            parsedData.official = [...parsedData.official, ...result];
            console.log('追加 Excel 解析結果:', result.length, '筆');
        }
    }
}

// 提示上傳報表範本
function promptTemplateUpload() {
    // 先收集用戶在表格中的調整
    collectMappingFromTable();

    // 觸發報表範本上傳
    document.getElementById('templateInput').click();
}

// 處理報表範本上傳
async function handleTemplateUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    showLoading('載入報表範本...');
    uploadedFiles.template = file;

    try {
        // 讀取檔案為 ArrayBuffer
        const arrayBuffer = await file.arrayBuffer();
        templateBuffer = arrayBuffer; // 保存原始二進制數據

        // 使用 SheetJS 解析（用於讀取資料）
        const data = new Uint8Array(arrayBuffer);
        templateWorkbook = XLSX.read(data, {
            type: 'array',
            cellStyles: true,
            cellFormula: true,
            bookVBA: true
        });
        console.log('SheetJS 報表範本已載入:', templateWorkbook.SheetNames);

        // 使用 ExcelJS 載入（保留完整樣式）
        excelWorkbook = new ExcelJS.Workbook();
        await excelWorkbook.xlsx.load(arrayBuffer);
        console.log('ExcelJS 報表範本已載入:', excelWorkbook.worksheets.map(ws => ws.name));

        // 調試：顯示行高列寬資訊
        const ws = excelWorkbook.worksheets[0];
        console.log('列寬資訊:', ws.columns.slice(0, 5).map(c => ({ col: c.number, width: c.width })));
        console.log('前5行高度:', Array.from({ length: 5 }, (_, i) => {
            const row = ws.getRow(i + 1);
            return { row: i + 1, height: row.height };
        }));

        hideLoading();
        // 繼續執行確認對應流程
        confirmMapping();
    } catch (err) {
        hideLoading();
        console.error('報表範本載入失敗:', err);
        showToast('報表範本載入失敗: ' + err.message, 'error');
    }
}

// 從表格收集使用者調整的對應關係
function collectMappingFromTable() {
    mappedProducts.forEach((product, index) => {
        const nameSelect = document.getElementById(`mapped-name-${index}`);
        const columnSelect = document.getElementById(`mapped-column-${index}`);
        const specSelect = document.getElementById(`mapped-spec-${index}`);

        if (nameSelect) product.templateProduct = nameSelect.value || null;
        if (columnSelect) product.templateColumn = columnSelect.value || null;
        if (specSelect) product.templateSpec = specSelect.value || null;
    });
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

// ==================== 檔案解析 ====================
async function parseAllFiles() {
    showLoading('正在解析撿貨單...');

    try {
        // 解析 MOMO 撿貨單
        if (uploadedFiles.momo) {
            parsedData.momo = await parseMomoExcel(uploadedFiles.momo);
            console.log('MOMO 解析結果:', parsedData.momo);
        }

        // 解析官網撿貨單
        if (uploadedFiles.official) {
            parsedData.official = await parseOfficialExcel(uploadedFiles.official);
            console.log('官網解析結果:', parsedData.official);
        }

        // 解析蝦皮撿貨單
        if (uploadedFiles.shopee) {
            parsedData.shopee = await parseShopeePDF(uploadedFiles.shopee);
            console.log('蝦皮解析結果:', parsedData.shopee);
        }

        // 解析橘點子撿貨單
        if (uploadedFiles.orangepoint) {
            parsedData.orangepoint = await parseOrangePointExcel(uploadedFiles.orangepoint);
            console.log('橘點子解析結果:', parsedData.orangepoint);
        }

        // 解析報表範本
        if (uploadedFiles.template) {
            templateWorkbook = await parseTemplateExcel(uploadedFiles.template);
            console.log('報表範本已載入');
        }

        // 整合所有商品
        consolidateProducts();

        hideLoading();
        showToast('撿貨單解析完成！', 'success');

        // 切換到步驟 2
        document.getElementById('step1').classList.add('hidden');
        document.getElementById('step2').classList.remove('hidden');

        // 建立對應表格
        buildMappingTable();

    } catch (error) {
        hideLoading();
        showToast('解析失敗：' + error.message, 'error');
        console.error('解析錯誤:', error);
    }
}

// 解析 MOMO Excel 撿貨單
async function parseMomoExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(worksheet);

                const products = jsonData
                    .filter(row => {
                        // 過濾掉運費行
                        const productCode = row['商品編碼'] || '';
                        const productName = row['商品名稱'] || '';
                        return !productName.includes('運費') &&
                            !productCode.includes('TP00019370000000');
                    })
                    .map(row => ({
                        name: row['商品名稱'] || '',
                        quantity: parseInt(row['撿貨數量'] || 0),
                        source: 'MOMO',
                        spec: row['單品規格'] || '',
                        rawData: row  // 保留原始資料供調試用
                    }))
                    .filter(p => p.name && p.quantity > 0);

                console.log('MOMO 原始資料行數:', jsonData.length);
                console.log('MOMO 過濾後商品數:', products.length);

                resolve(products);
            } catch (error) {
                console.error('MOMO 解析錯誤:', error);
                reject(error);
            }
        };

        reader.onerror = () => reject(new Error('無法讀取 MOMO 檔案'));
        reader.readAsArrayBuffer(file);
    });
}

// 解析官網 Excel 撿貨單
async function parseOfficialExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                // 官網撿貨單前6行是元數據，第7行是標題，第8行開始是資料
                // 使用 range: 6 跳過前6行，讓第7行成為標題行
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { range: 6 });

                console.log('官網原始資料:', jsonData);
                console.log('官網資料行數:', jsonData.length);
                if (jsonData.length > 0) {
                    console.log('官網欄位名稱:', Object.keys(jsonData[0]));
                }

                const products = jsonData.map(row => {
                    // 嘗試多種可能的欄位名稱
                    const name = row['商品名稱'] || row['品名'] || row['商品'] || row['產品名稱'] || '';
                    const quantity = parseInt(row['數量'] || row['撿貨數量'] || row['訂購數量'] || row['Qty'] || 0);
                    const spec = row['規格'] || row['單品規格'] || row['品規'] || '';

                    return {
                        name: name,
                        quantity: quantity,
                        source: '官網',
                        spec: spec,
                        rawData: row
                    };
                }).filter(p => p.name && p.quantity > 0);

                console.log('官網過濾後商品數:', products.length);
                resolve(products);
            } catch (error) {
                console.error('官網解析錯誤:', error);
                reject(error);
            }
        };

        reader.onerror = () => reject(new Error('無法讀取官網檔案'));
        reader.readAsArrayBuffer(file);
    });
}

// 解析蝦皮 PDF 撿貨單（使用 X 座標精確解析）
async function parseShopeePDF(file) {
    return new Promise(async (resolve, reject) => {
        try {
            const arrayBuffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

            const allProducts = [];
            for (let i = 1; i <= pdf.numPages; i++) {
                const page = await pdf.getPage(i);
                const textContent = await page.getTextContent();
                const pageItems = textContent.items.map(item => ({
                    text: item.str.trim(),
                    x: Math.round(item.transform[4]),
                    y: Math.round(item.transform[5])
                }));

                const products = parseShopeePage(pageItems);
                allProducts.push(...products);
            }

            resolve(allProducts.map(p => ({
                name: p.name,
                quantity: p.quantity,
                spec: p.spec,
                source: '蝦皮'
            })));

        } catch (error) {
            reject(new Error('無法解析蝦皮 PDF：' + error.message));
        }
    });
}

function parseShopeePage(items) {
    const rows = groupByY(items);
    const columnRanges = {
        seq: { min: 32, max: 59 },
        productName: { min: 166, max: 307 },
        spec: { min: 308, max: 396 },
        quantity: { min: 397, max: 520 }
    };

    let headerY = null;
    for (let row of rows) {
        if (row.items.some(i => i.text.includes('商品名稱'))) {
            headerY = row.y;
            break;
        }
    }

    if (!headerY) return [];

    const products = [];
    let currentProduct = null;

    rows.forEach(row => {
        if (row.y >= headerY) return;

        let seqNum = null;
        row.items.forEach(item => {
            if (item.x >= columnRanges.seq.min && item.x <= columnRanges.seq.max) {
                const num = parseInt(item.text.trim());
                if (!isNaN(num) && num > 0) seqNum = num;
            }
        });

        if (seqNum) {
            if (currentProduct) products.push(currentProduct);
            currentProduct = { seq: seqNum, name: [], spec: [], quantity: null };
        }

        if (currentProduct) {
            const rowText = row.items.map(i => i.text).join(' ');
            if (rowText.includes('合計')) return;

            row.items.forEach(item => {
                const x = item.x;
                const t = item.text.trim();

                if (x >= columnRanges.productName.min && x <= columnRanges.productName.max) {
                    if (!t.match(/^\d+$/)) currentProduct.name.push(t);
                } else if (x >= columnRanges.spec.min && x <= columnRanges.spec.max) {
                    if (!t.match(/^\d+$/)) currentProduct.spec.push(t);
                } else if (x >= columnRanges.quantity.min && x <= columnRanges.quantity.max) {
                    const num = parseInt(t);
                    if (!isNaN(num) && num > 0) currentProduct.quantity = num;
                }
            });
        }
    });

    if (currentProduct) products.push(currentProduct);

    return products.filter(p => p.seq && p.seq > 0).map(p => ({
        name: p.name.join(' '),
        spec: p.spec.join(' '),
        quantity: p.quantity || 0
    }));
}

function groupByY(items, tolerance = 5) {
    const sorted = [...items].sort((a, b) => b.y - a.y);
    const groups = [];
    sorted.forEach(item => {
        let found = false;
        for (let g of groups) {
            if (Math.abs(g.y - item.y) <= tolerance) {
                g.items.push(item);
                found = true;
                break;
            }
        }
        if (!found) groups.push({ y: item.y, items: [item] });
    });
    groups.forEach(g => g.items.sort((a, b) => a.x - b.x));
    return groups;
}

// 解析報表範本
async function parseTemplateExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {
                    type: 'array',
                    cellStyles: true,      // 保留儲存格樣式
                    cellFormula: true,     // 保留公式
                    cellHTML: true,        // 保留 HTML
                    cellNF: true,          // 保留數字格式
                    cellDates: true,       // 保留日期
                    bookVBA: true,         // 保留巨集
                    bookImages: true,      // 保留圖片
                    sheetStubs: true       // 保留空白儲存格
                });
                console.log('報表範本載入完成，工作表:', workbook.SheetNames);
                resolve(workbook);
            } catch (error) {
                reject(error);
            }
        };

        reader.onerror = () => reject(new Error('無法讀取報表範本'));
        reader.readAsArrayBuffer(file);
    });
}

// ==================== 商品整合 ====================

// 拆分特殊商品規則
function splitSpecialProducts(products) {
    const result = [];

    products.forEach(product => {
        const fullText = ((product.name || '') + (product.spec || '')).replace(/\s+/g, '');

        // 胖貝殼瑪德蓮-2入 拆分為 蜂蜜 + 巧克力（固定組合）
        if (/胖貝殼瑪德蓮/.test(fullText) && /2入/.test(fullText)) {
            console.log(`拆分商品: ${product.name} (數量 ${product.quantity}) → 蜂蜜 + 巧克力 各 ${product.quantity}`);

            // 拆成兩筆，每筆數量與原始訂單相同
            result.push({
                ...product,
                name: '瑪德蓮-蜂蜜（活動拆分）',
                spec: '單顆',
                quantity: product.quantity,
                originalName: product.name,
                originalSpec: product.spec,
                isSplit: true
            });
            result.push({
                ...product,
                name: '瑪德蓮-巧克力（活動拆分）',
                spec: '單顆',
                quantity: product.quantity,
                originalName: product.name,
                originalSpec: product.spec,
                isSplit: true
            });
        }
        // 夏威夷豆塔組合包：10入蜂蜜蔓越莓x2包+10入焦糖口味x1包 拆分為 蔓越莓 + 焦糖
        else if (/夏威夷豆塔/.test(fullText) && /蜂蜜蔓越莓.*2包/.test(fullText) && /焦糖.*1包/.test(fullText)) {
            console.log(`拆分商品: ${product.name} (數量 ${product.quantity}) → 蔓越莓 ${product.quantity * 2}包 + 焦糖 ${product.quantity}包`);

            // 蔓越莓 x2包
            result.push({
                ...product,
                name: '豆塔-蔓越莓（組合拆分）',
                spec: '10入袋裝',
                quantity: product.quantity * 2,
                originalName: product.name,
                originalSpec: product.spec,
                isSplit: true
            });
            // 焦糖 x1包
            result.push({
                ...product,
                name: '豆塔-焦糖（組合拆分）',
                spec: '10入袋裝',
                quantity: product.quantity * 1,
                originalName: product.name,
                originalSpec: product.spec,
                isSplit: true
            });
        }
        // 夏威夷豆塔組合包：10入焦糖口味x2包+10入蜂蜜蔓越莓x1包 拆分為 焦糖 + 蔓越莓（順序相反）
        else if (/夏威夷豆塔/.test(fullText) && /焦糖.*2包/.test(fullText) && /蜂蜜蔓越莓.*1包/.test(fullText)) {
            console.log(`拆分商品: ${product.name} (數量 ${product.quantity}) → 焦糖 ${product.quantity * 2}包 + 蔓越莓 ${product.quantity}包`);

            // 焦糖 x2包
            result.push({
                ...product,
                name: '豆塔-焦糖（組合拆分）',
                spec: '10入袋裝',
                quantity: product.quantity * 2,
                originalName: product.name,
                originalSpec: product.spec,
                isSplit: true
            });
            // 蔓越莓 x1包
            result.push({
                ...product,
                name: '豆塔-蔓越莓（組合拆分）',
                spec: '10入袋裝',
                quantity: product.quantity * 1,
                originalName: product.name,
                originalSpec: product.spec,
                isSplit: true
            });
        } else {
            result.push(product);
        }
    });

    return result;
}

function consolidateProducts() {
    allProducts = [];
    mappedProducts = [];

    // 合併所有平台的商品（包含橘點子）- 根據名稱、來源和規格合併
    [...parsedData.momo, ...parsedData.official, ...parsedData.shopee, ...parsedData.orangepoint].forEach(product => {
        const existing = allProducts.find(p =>
            p.name === product.name && p.source === product.source && p.spec === product.spec
        );

        if (existing) {
            existing.quantity += product.quantity;
        } else {
            allProducts.push({ ...product });
        }
    });

    // 拆分特殊商品（如：胖貝殼瑪德蓮-2入 → 蜂蜜 + 巧克力）
    allProducts = splitSpecialProducts(allProducts);

    // 應用智慧映射
    allProducts.forEach(product => {
        let platform = 'momo';
        if (product.source === '蝦皮') platform = 'shopee';
        else if (product.source === '橘點子') platform = 'orangepoint';
        else if (product.source === '官網') platform = 'official';
        // MOMO 也使用 momo 平台映射

        const mapping = autoMapProduct(product.name, product.spec || '', product.quantity, platform);

        console.log(`映射: ${product.name} [${product.source}] -> ${mapping.templateProduct || '未識別'} (${Math.round((mapping.confidence || 0) * 100)}%)`);

        mappedProducts.push({
            ...product,
            ...mapping
        });
    });

    console.log('整合後的商品:', allProducts);
    console.log('映射後的商品:', mappedProducts);
}

// ==================== 商品對應 ====================
// 規格選項
const specOptions = ['', '30g', '45g', '50g', '60g', '90g', '120g', '135g', '150g', '200g', '280g', '300g', '8入袋裝', '10入袋裝', '12入袋裝', '15入袋裝', '小包裝', '單顆', '禮盒'];
// 欄位選項
const columnOptions = ['', 'B', 'C', 'D', 'E'];

function buildMappingTable() {
    const tbody = document.getElementById('mappingTableBody');
    tbody.innerHTML = '';

    // 按照商品順序排序 mappedProducts（與步驟 3 相同的排序邏輯）
    const sortedProducts = [...mappedProducts].sort((a, b) => {
        const orderA = getProductCategoryOrder(a.templateProduct || '');
        const orderB = getProductCategoryOrder(b.templateProduct || '');
        if (orderA !== orderB) return orderA - orderB;
        return (a.templateProduct || '').localeCompare(b.templateProduct || '', 'zh-TW');
    });

    // 更新 mappedProducts 的順序（保持同步）
    mappedProducts.length = 0;
    mappedProducts.push(...sortedProducts);

    mappedProducts.forEach((product, index) => {
        const row = document.createElement('tr');
        const confidence = Math.round((product.confidence || 0) * 100);
        const confidenceColor = confidence >= 90 ? '#10b981' : confidence >= 70 ? '#f59e0b' : '#ef4444';

        // 生成報表商品下拉選單選項
        const productOptionsHtml = standardProductNames.map(name =>
            `<option value="${name}" ${product.templateProduct === name ? 'selected' : ''}>${name}</option>`
        ).join('');

        // 生成欄位下拉選單
        const columnOptionsHtml = columnOptions.map(col =>
            `<option value="${col}" ${product.templateColumn === col ? 'selected' : ''}>${col || '-'}</option>`
        ).join('');

        // 生成規格下拉選單
        const specOptionsHtml = specOptions.map(spec =>
            `<option value="${spec}" ${product.templateSpec === spec ? 'selected' : ''}>${spec || '-'}</option>`
        ).join('');

        row.innerHTML = `
            <td title="${product.name}">${product.name.length > 25 ? product.name.substring(0, 25) + '...' : product.name}</td>
            <td><small>${product.spec || '-'}</small></td>
            <td><span style="color: ${getSourceColor(product.source)}">${product.source}</span></td>
            <td><strong>${product.quantity}</strong></td>
            <td style="text-align: center;">→</td>
            <td>
                <select id="mapped-name-${index}" style="background: ${product.templateProduct ? '#e8f5e9' : '#fff3e0'}; color: #1e293b; min-width: 140px;">
                    <option value="">-- 選擇商品 --</option>
                    ${productOptionsHtml}
                </select>
            </td>
            <td>
                <select id="mapped-column-${index}" style="width: 50px; color: #1e293b;">
                    ${columnOptionsHtml}
                </select>
            </td>
            <td>
                <select id="mapped-spec-${index}" style="min-width: 80px; color: #1e293b;">
                    ${specOptionsHtml}
                </select>
            </td>
            <td>
                <strong style="color: #3b82f6;">${product.mappedQuantity || product.quantity}</strong>
            </td>
            <td>
                <span style="color: ${confidenceColor}; font-weight: bold;">${confidence}%</span>
            </td>
        `;
        tbody.appendChild(row);
    });
}

function getSourceColor(source) {
    const colors = {
        'MOMO': '#f59e0b',
        '官網': '#10b981',
        '蝦皮': '#ef4444',
        '橘點子': '#FF6B00'
    };
    return colors[source] || '#cbd5e1';
}

// 自動智慧對應
function autoMapProducts() {
    showLoading('智慧對應中...');

    setTimeout(() => {
        allProducts.forEach((product, index) => {
            const mappedNameInput = document.getElementById(`mapped-name-${index}`);
            const categorySelect = document.getElementById(`mapped-category-${index}`);

            // 簡單的智慧對應邏輯
            const name = product.name.toLowerCase();

            // 根據關鍵字判斷分類
            if (name.includes('戚風') || name.includes('米') || name.includes('基底')) {
                categorySelect.value = '訂購基底類';
            } else if (name.includes('蛋糕') || name.includes('提拉米蘇')) {
                categorySelect.value = '蛋糕類別表單';
            } else if (name.includes('10個') || name.includes('十個')) {
                categorySelect.value = '10個裝';
            } else if (name.includes('15個') || name.includes('十五個')) {
                categorySelect.value = '15個裝';
            } else if (name.includes('10入') || name.includes('十入')) {
                categorySelect.value = '10入裝';
            } else if (name.includes('小包')) {
                categorySelect.value = '小包裝';
            } else if (name.includes('果乾') || name.includes('乾果')) {
                categorySelect.value = '果乾類';
            }
        });

        hideLoading();
        showToast('自動對應完成！請檢查並調整', 'success');
    }, 500);
}

// 載入對應規則
function loadMappingRules() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json';

    input.onchange = (e) => {
        const file = e.target.files[0];
        const reader = new FileReader();

        reader.onload = (event) => {
            try {
                const rules = JSON.parse(event.target.result);
                applyMappingRules(rules);
                showToast('對應規則載入成功！', 'success');
            } catch (error) {
                showToast('載入失敗：無效的 JSON 格式', 'error');
            }
        };

        reader.readAsText(file);
    };

    input.click();
}

function applyMappingRules(rules) {
    allProducts.forEach((product, index) => {
        const rule = rules[product.name];
        if (rule) {
            document.getElementById(`mapped-name-${index}`).value = rule.reportName || product.name;
            document.getElementById(`mapped-category-${index}`).value = rule.category || '';
        }
    });
}

// 儲存對應規則
function saveMappingRules() {
    const rules = {};

    allProducts.forEach((product, index) => {
        const mappedName = document.getElementById(`mapped-name-${index}`).value;
        const category = document.getElementById(`mapped-category-${index}`).value;

        rules[product.name] = {
            reportName: mappedName,
            category: category
        };
    });

    const blob = new Blob([JSON.stringify(rules, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = '商品對應規則.json';
    a.click();
    URL.revokeObjectURL(url);

    showToast('對應規則已儲存！', 'success');
}

// 確認對應
function confirmMapping() {
    showLoading('處理對應關係...');

    setTimeout(() => {
        try {
            // 建立對應關係
            productMapping = {};
            statistics = {};

            // 使用 mappedProducts 而不是 allProducts
            mappedProducts.forEach((product, index) => {
                const nameSelect = document.getElementById(`mapped-name-${index}`);
                const columnSelect = document.getElementById(`mapped-column-${index}`);
                const specSelect = document.getElementById(`mapped-spec-${index}`);

                const mappedName = nameSelect ? nameSelect.value.trim() : product.templateProduct;
                const column = columnSelect ? columnSelect.value : product.templateColumn;
                const spec = specSelect ? specSelect.value : product.templateSpec;

                if (!mappedName) {
                    return;
                }

                // 更新 mappedProducts 中的對應資訊
                product.templateProduct = mappedName;
                product.templateColumn = column;
                product.templateSpec = spec;

                productMapping[product.name] = {
                    reportName: mappedName,
                    column: column,
                    spec: spec
                };

                // 統計數量（按商品名+規格分組）
                const key = `${mappedName}_${spec || 'default'}`;
                if (!statistics[key]) {
                    statistics[key] = {
                        name: mappedName,
                        column: column,
                        spec: spec,
                        flavor: product.spec || '',  // 使用原始商品的規格作為口味
                        quantity: 0
                    };
                }
                statistics[key].quantity += product.mappedQuantity || product.quantity;
            });

            console.log('統計結果:', statistics);
            console.log('商品對應:', productMapping);

            hideLoading();

            // 切換到步驟 3
            document.getElementById('step2').classList.add('hidden');
            document.getElementById('step3').classList.remove('hidden');

            // 啟用兩個按鈕（因為有上傳報表範本）
            const generateReportBtn = document.getElementById('generateReportBtn');
            if (generateReportBtn) {
                generateReportBtn.disabled = false;
                generateReportBtn.style.opacity = '1';
                generateReportBtn.style.cursor = 'pointer';
                generateReportBtn.style.pointerEvents = 'auto';
                generateReportBtn.title = '';
            }

            const downloadPickingListBtn = document.getElementById('downloadPickingListBtn');
            if (downloadPickingListBtn) {
                downloadPickingListBtn.disabled = false;
                downloadPickingListBtn.style.opacity = '1';
                downloadPickingListBtn.style.cursor = 'pointer';
            }

            // 顯示統計結果
            displayStatistics();
        } catch (error) {
            console.error('確認對應錯誤:', error);
            hideLoading();
            showToast('處理對應關係失敗: ' + error.message, 'error');
        }
    }, 500);
}

// ==================== 統計顯示 ====================

// 商品排序配置（明確順序列表）
const productSortOrder = [
    '豆塔-蔓越莓',
    '豆塔-焦糖',
    '豆塔-巧克力',
    '豆塔-抹茶',
    '豆塔-椒麻',
    '堅果塔-蜂蜜',
    '堅果塔-焦糖',
    '堅果塔-巧克力',
    '堅果塔-海苔',
    '堅果塔-咖哩',
    '土鳳梨酥(紅點)',
    '鳳凰酥',
    '豆塔-綜合',
    '堅果塔-綜合',
    '雙塔',
    '雪花餅-綜合',
    '雪花餅-蔓越莓',
    '雪花餅-巧克力',
    '雪花餅-金沙',
    '雪花餅-抹茶',
    '雪花餅-肉鬆',
    '瑪德蓮-綜合',
    '瑪德蓮-蜂蜜',
    '瑪德蓮-巧克力',
    '瑪德蓮-紅茶',
    '瑪德蓮-抹茶',
    '瑪德蓮-柑橘',
    '瑪德蓮-檸檬',
    '無調味綜合堅果',
    '無調味夏威夷豆',
    '無調味腰果',
    '無調味杏仁',
    '無調味核桃',
    '★中東椰棗300g',
    '椰棗豆子150g',
    '椰棗腰果150g',
    '椰棗杏仁150g',
    '椰棗核桃150g',
    '牛奶糖',
    '南棗核桃糕',
    '牛奶糖-50g',
    '瓦片-綜合',
    '瓦片-原味',
    '瓦片-紅茶',
    '瓦片-巧克力',
    '瓦片-海苔',
    '瓦片-抹茶',
    '瓦片-黑糖',
    '瓦片-青花椒',
    '瓦片-原味45克',
    '奶油-焦糖牛奶',
    '奶油-法國巧克力',
    '奶油-蜂蜜檸檬',
    '奶油-伯爵紅茶',
    '奶油-抹茶',
    '西點-綜合',
    '西點-巧克力貝殼',
    '西點-咖啡小花',
    '西點-藍莓小花',
    '西點-蔓越莓貝殼',
    '西點-乳酪酥條',
    '千層-捲捲酥',
    '千層-小酥條',
    '千層-小酥條30g',
    '千層-捲捲酥30g',
    '雙塔禮盒',
    '蔓越莓禮盒',
    '綜豆禮盒',
    '綜堅禮盒',
    '晴空塔餅禮盒',
    '暖暖幸福禮盒',
    '臻愛時光禮盒',
    '濃情滿載禮盒',
    '浪漫詩篇禮盒',
    '戀戀雪花禮盒',
    '午後漫步禮盒',
    '那年花開禮盒',
    '花間逸韻禮盒',
    '輕-香榭漫遊禮盒',
    '輕-晨曦物語禮盒',
    '輕-金緻典藏禮盒',
    '輕-月光序曲禮盒'
];

function getProductCategoryOrder(productName) {
    const name = productName || '';
    const index = productSortOrder.indexOf(name);
    if (index !== -1) {
        return index;
    }
    // 未在列表中的商品排到最後，按名稱排序
    return 9999;
}

function displayStatistics() {
    const container = document.getElementById('statsContainer');
    container.innerHTML = '';

    // 新的統計格式：statistics 是一個物件，key 是 "商品名_規格"，value 是 { name, column, spec, quantity, flavor }
    let statsArray = Object.values(statistics);

    if (statsArray.length === 0) {
        container.innerHTML = '<p style="text-align: center; color: #f59e0b;">沒有統計資料</p>';
        return;
    }

    // 按分類順序排序
    statsArray.sort((a, b) => {
        const orderA = getProductCategoryOrder(a.name);
        const orderB = getProductCategoryOrder(b.name);
        if (orderA !== orderB) return orderA - orderB;
        // 同分類內按名稱排序
        return (a.name || '').localeCompare(b.name || '', 'zh-TW');
    });

    // 建立表格顯示統計結果
    const table = document.createElement('table');
    table.className = 'mapping-table';
    table.innerHTML = `
        <thead>
            <tr>
                <th>報表商品</th>
                <th>口味</th>
                <th>規格</th>
                <th>欄位</th>
                <th>總數量</th>
            </tr>
        </thead>
        <tbody>
            ${statsArray.map(stat => `
                <tr>
                    <td>${stat.name}</td>
                    <td>${stat.flavor || '-'}</td>
                    <td>${stat.spec || '-'}</td>
                    <td>${stat.column || '-'}</td>
                    <td><strong>${stat.quantity}</strong></td>
                </tr>
            `).join('')}
        </tbody>
    `;
    container.appendChild(table);

    // 顯示總計
    const totalQuantity = statsArray.reduce((sum, s) => sum + s.quantity, 0);
    const summary = document.createElement('div');
    summary.style.cssText = 'margin-top: 20px; text-align: center; font-size: 1.2em;';
    summary.innerHTML = `<strong>總計：${statsArray.length} 種商品，${totalQuantity} 件</strong>`;
    container.appendChild(summary);
}

// ==================== 報表產生 ====================
function generateReport() {
    if (!templateWorkbook) {
        showToast('請先上傳報表範本！', 'error');
        return;
    }

    showLoading('正在產生報表...');

    setTimeout(() => {
        try {
            fillTemplate();

            hideLoading();
            showToast('報表產生成功！', 'success');

            // 切換到步驟 4
            document.getElementById('step3').classList.add('hidden');
            document.getElementById('step4').classList.remove('hidden');

        } catch (error) {
            hideLoading();
            showToast('報表產生失敗：' + error.message, 'error');
            console.error('報表產生錯誤:', error);
        }
    }, 1000);
}

function fillTemplate() {
    // 使用 ExcelJS 填入資料（保留完整樣式）
    if (!excelWorkbook) {
        console.error('ExcelJS workbook 未載入');
        return;
    }

    const worksheet = excelWorkbook.worksheets[0];
    console.log('開始填入報表...');
    console.log('統計資料:', statistics);
    console.log('使用工作表:', worksheet.name);

    // 欄位字母轉數字 (A=1, B=2, C=3, D=4, E=5)
    const colToNum = { 'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5 };
    let filledCount = 0;

    // 遍歷工作表的每一行
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        const nameCell = row.getCell(1); // A欄
        if (!nameCell.value) return;

        const cellValue = String(nameCell.value).trim();

        // 遍歷所有統計項目尋找匹配
        Object.values(statistics).forEach(stat => {
            // 使用精確匹配，避免 "瓦片-原味" 誤匹配到 "瓦片-原味45克"
            // 只有當報表中的商品名稱與統計的商品名稱完全相等時才匹配
            if (stat.name && cellValue === stat.name) {
                // 獲取要填入的欄位
                const colLetter = stat.column || 'B';
                const colNum = colToNum[colLetter] || 2;

                // 填入數量（只修改值，保留原有樣式）
                const targetCell = row.getCell(colNum);
                targetCell.value = stat.quantity;
                filledCount++;
                console.log(`填入: ${stat.name} 在 ${colLetter}${rowNumber} = ${stat.quantity}`);
            }
        });
    });

    console.log(`總共填入 ${filledCount} 筆資料`);

    if (filledCount === 0) {
        console.warn('警告：沒有找到匹配的商品名稱填入報表');
        // 顯示工作表中的商品名稱供調試
        const names = [];
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            if (rowNumber <= 15) {
                const val = row.getCell(1).value;
                if (val) names.push(String(val));
            }
        });
        console.log('報表中的商品名稱（A欄前15行）:', names);
    }
}

async function downloadReport() {
    try {
        console.log('開始使用 JSZip 方法下載報表...');

        // 使用 JSZip 載入原始範本
        const zip = await JSZip.loadAsync(templateBuffer);

        // 讀取 sheet1.xml（工作表數據）
        const sheetPath = 'xl/worksheets/sheet1.xml';
        let sheetXml = await zip.file(sheetPath).async('string');

        // 讀取 sharedStrings.xml（共享字串表）
        const sharedStringsPath = 'xl/sharedStrings.xml';
        let sharedStringsXml = '';
        const sharedStringsFile = zip.file(sharedStringsPath);
        if (sharedStringsFile) {
            sharedStringsXml = await sharedStringsFile.async('string');
        }

        // 建立商品名稱到行號的映射（從 ExcelJS 獲取）
        const ws = excelWorkbook.worksheets[0];
        const cellUpdates = [];

        // 欄位字母轉數字
        const colToNum = { 'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5 };

        // 遍歷工作表找到需要更新的儲存格
        ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            const nameCell = row.getCell(1);
            if (!nameCell.value) return;
            const cellValue = String(nameCell.value).trim();

            Object.values(statistics).forEach(stat => {
                // 使用精確匹配，避免重複填入
                if (stat.name && cellValue === stat.name) {
                    const colLetter = stat.column || 'B';
                    const colNum = colToNum[colLetter] || 2;
                    const cellRef = `${colLetter}${rowNumber}`;
                    cellUpdates.push({ cellRef, value: stat.quantity, row: rowNumber, col: colNum });
                    console.log(`準備更新: ${cellRef} = ${stat.quantity}`);
                }
            });
        });

        // 在 XML 中更新儲存格值（保留所有屬性）
        cellUpdates.forEach(update => {
            const cellRef = update.cellRef;
            let matched = false;

            // 1. 嘗試匹配含公式和值的儲存格：<c r="B5" s="1"><f>...</f><v>123</v></c>
            // 移除公式，只保留新值
            if (!matched) {
                const cellWithFormulaAndValuePattern = new RegExp(
                    `(<c r="${cellRef}"[^>]*>)<f[^>]*>[\\s\\S]*?</f><v>[^<]*</v>(</c>)`
                );
                if (cellWithFormulaAndValuePattern.test(sheetXml)) {
                    sheetXml = sheetXml.replace(cellWithFormulaAndValuePattern, `$1<v>${update.value}</v>$2`);
                    console.log(`更新儲存格 ${cellRef}（移除公式）的值為 ${update.value}`);
                    matched = true;
                }
            }

            // 2. 嘗試匹配只有公式的儲存格：<c r="B5" s="1"><f>...</f></c>
            // 移除公式，插入新值
            if (!matched) {
                const cellWithOnlyFormulaPattern = new RegExp(
                    `(<c r="${cellRef}"[^>]*>)<f[^>]*>[\\s\\S]*?</f>(</c>)`
                );
                if (cellWithOnlyFormulaPattern.test(sheetXml)) {
                    sheetXml = sheetXml.replace(cellWithOnlyFormulaPattern, `$1<v>${update.value}</v>$2`);
                    console.log(`更新儲存格 ${cellRef}（移除公式）的值為 ${update.value}`);
                    matched = true;
                }
            }

            // 3. 嘗試匹配已有值的儲存格：<c r="B5" s="1"><v>123</v></c>
            // 只替換 <v>...</v> 部分
            if (!matched) {
                const cellWithValuePattern = new RegExp(
                    `(<c r="${cellRef}"[^>]*>)<v>[^<]*</v>(</c>)`
                );
                if (cellWithValuePattern.test(sheetXml)) {
                    sheetXml = sheetXml.replace(cellWithValuePattern, `$1<v>${update.value}</v>$2`);
                    console.log(`更新儲存格 ${cellRef} 的值為 ${update.value}`);
                    matched = true;
                }
            }

            // 4. 嘗試匹配空儲存格：<c r="B5" s="1"/>
            if (!matched) {
                const emptyCellPattern = new RegExp(
                    `<c r="${cellRef}"([^/>]*)/>`
                );
                if (emptyCellPattern.test(sheetXml)) {
                    sheetXml = sheetXml.replace(emptyCellPattern, `<c r="${cellRef}"$1><v>${update.value}</v></c>`);
                    console.log(`填入空儲存格 ${cellRef} 值為 ${update.value}`);
                    matched = true;
                }
            }

            // 5. 嘗試匹配空內容儲存格：<c r="B5" s="1"></c>
            if (!matched) {
                const emptyContentCellPattern = new RegExp(
                    `(<c r="${cellRef}"[^>]*>)(</c>)`
                );
                if (emptyContentCellPattern.test(sheetXml)) {
                    sheetXml = sheetXml.replace(emptyContentCellPattern, `$1<v>${update.value}</v>$2`);
                    console.log(`填入空內容儲存格 ${cellRef} 值為 ${update.value}`);
                    matched = true;
                }
            }

            if (!matched) {
                console.log(`儲存格 ${cellRef} 在 XML 中未找到匹配模式`);
            }
        });

        // 將修改後的 XML 寫回 ZIP
        zip.file(sheetPath, sheetXml);

        // 刪除 calcChain.xml 強制 Excel 重新計算所有公式
        if (zip.file('xl/calcChain.xml')) {
            zip.remove('xl/calcChain.xml');
            console.log('已刪除 calcChain.xml');

            // 同時從 [Content_Types].xml 中移除對 calcChain 的引用
            const contentTypesPath = '[Content_Types].xml';
            let contentTypesXml = await zip.file(contentTypesPath).async('string');
            contentTypesXml = contentTypesXml.replace(/<Override[^>]*calcChain[^>]*\/>/g, '');
            zip.file(contentTypesPath, contentTypesXml);
            console.log('已從 Content_Types.xml 移除 calcChain 引用');
        }

        // 設定 workbook.xml 強制開啟時重新計算公式
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

        // 注意：不再清除公式緩存值，因為正則表達式無法正確處理複雜公式
        // 刪除 calcChain.xml 已經足夠強制 Excel 重新計算公式

        // 生成最終文件
        const content = await zip.generateAsync({
            type: 'blob',
            mimeType: 'application/vnd.ms-excel.sheet.macroEnabled.12'
        });

        // 下載文件
        const url = URL.createObjectURL(content);
        const a = document.createElement('a');
        a.href = url;
        const today = new Date().toISOString().split('T')[0].replace(/-/g, '');
        // 保持原始副檔名
        a.download = `生產統計表_${today}.xlsm`;
        a.click();
        URL.revokeObjectURL(url);

        console.log(`成功更新 ${cellUpdates.length} 個儲存格`);
        showToast('報表下載完成！', 'success');
    } catch (error) {
        showToast('下載失敗：' + error.message, 'error');
        console.error('下載錯誤:', error);
    }
}

// ==================== 工具函數 ====================
function showLoading(text = '處理中...') {
    document.getElementById('loadingText').textContent = text;
    document.getElementById('loadingOverlay').classList.remove('hidden');
}

function hideLoading() {
    document.getElementById('loadingOverlay').classList.add('hidden');
}

function showToast(message, type = 'info') {
    const toast = document.getElementById('toast');
    toast.textContent = message;
    toast.className = 'toast ' + type;

    setTimeout(() => {
        toast.classList.add('hidden');
    }, 3000);
}

function resetApplication() {
    if (confirm('確定要重新開始嗎？所有資料將會清除。')) {
        location.reload();
    }
}

// ==================== 橘點子解析 ====================
async function parseOrangePointExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

                const products = [];
                const seenProducts = new Set();

                for (let row of jsonData) {
                    for (let cell of row) {
                        if (cell === '品名') continue;
                        const parsedItems = parseOrangePointCell(String(cell));
                        for (const parsed of parsedItems) {
                            const key = `${parsed.name}|${parsed.quantity}`;
                            if (!seenProducts.has(key)) {
                                seenProducts.add(key);
                                products.push({
                                    name: parsed.name,
                                    quantity: parsed.quantity,
                                    spec: '',
                                    source: '橘點子'
                                });
                            }
                        }
                    }
                }

                resolve(products);
            } catch (error) {
                reject(error);
            }
        };

        reader.onerror = () => reject(new Error('檔案讀取失敗'));
        reader.readAsArrayBuffer(file);
    });
}

// 解析橘點子單一欄位
function parseOrangePointCell(cellValue) {
    if (!cellValue || typeof cellValue !== 'string') return [];

    const text = cellValue.trim();
    if (!text) return [];

    const results = [];

    // 過濾咖啡
    if (text.includes('咖啡') && !text.includes('咖啡小花')) {
        return [];
    }

    // 贈品格式
    if (text.startsWith('贈品:')) {
        const giftMatch = text.match(/[#](.+?)\s*[x×]?\s*(\d+)$/i);
        if (giftMatch) {
            const quantity = parseInt(giftMatch[2]);
            if (quantity === 0) return [];
            results.push({ name: giftMatch[1].trim(), quantity: quantity });
        }
        return results;
    }

    // 試吃格式
    if (text.startsWith('試吃:')) {
        const quantityMatch = text.match(/\s+(\d+)$/);
        if (!quantityMatch) return [];
        const totalQuantity = parseInt(quantityMatch[1]);
        if (totalQuantity === 0) return [];

        const productPart = text.replace(/^試吃:/, '').replace(/\s+\d+$/, '').trim();
        const productItems = productPart.split('+');
        for (const item of productItems) {
            const cleanName = item.replace(/\*\d+$/, '').trim();
            if (cleanName) {
                results.push({ name: cleanName, quantity: totalQuantity });
            }
        }
        return results;
    }

    // 普通格式
    const match = text.match(/^(.+?)\s+(\d+)$/);
    if (match) {
        const name = match[1].trim();
        const quantity = parseInt(match[2]);
        if (name.toLowerCase() === 'total' || quantity === 0) return [];
        results.push({ name: name, quantity: quantity });
    }

    return results;
}

// ==================== 跳過上傳報表直接預覽 ====================
function skipToStep3() {
    // 收集對應資料
    collectMappingFromTable();

    // 建立統計資料（與 confirmMapping 相同邏輯）
    productMapping = {};
    statistics = {};

    mappedProducts.forEach((product, index) => {
        const nameSelect = document.getElementById(`mapped-name-${index}`);
        const columnSelect = document.getElementById(`mapped-column-${index}`);
        const specSelect = document.getElementById(`mapped-spec-${index}`);

        const mappedName = nameSelect ? nameSelect.value.trim() : product.templateProduct;
        const column = columnSelect ? columnSelect.value : product.templateColumn;
        const spec = specSelect ? specSelect.value : product.templateSpec;

        if (!mappedName) {
            return;
        }

        // 更新 mappedProducts 中的對應資訊
        product.templateProduct = mappedName;
        product.templateColumn = column;
        product.templateSpec = spec;

        productMapping[product.name] = {
            reportName: mappedName,
            column: column,
            spec: spec
        };

        // 統計數量（按商品名+規格分組）
        const key = `${mappedName}_${spec || 'default'}`;
        if (!statistics[key]) {
            statistics[key] = {
                name: mappedName,
                column: column,
                spec: spec,
                flavor: product.spec || '',  // 使用原始商品的規格作為口味
                quantity: 0
            };
        }
        statistics[key].quantity += product.mappedQuantity || product.quantity;
    });

    console.log('跳過報表 - 統計結果:', statistics);

    // 直接顯示步驟 3（不需要上傳報表範本）
    document.getElementById('step2').classList.add('hidden');
    document.getElementById('step3').classList.remove('hidden');

    // 停用「下載生產統計表」按鈕（因為沒有上傳報表範本）
    const generateReportBtn = document.getElementById('generateReportBtn');
    console.log('generateReportBtn 元素:', generateReportBtn);
    if (generateReportBtn) {
        generateReportBtn.disabled = true;
        generateReportBtn.style.opacity = '0.5';
        generateReportBtn.style.cursor = 'not-allowed';
        generateReportBtn.style.pointerEvents = 'none';
        generateReportBtn.title = '需要先上傳生產統計表範本才能使用此功能';
        console.log('已停用「下載生產統計表」按鈕');
    }

    // 確保「下載撿貨單」按鈕可用
    const downloadPickingListBtn = document.getElementById('downloadPickingListBtn');
    if (downloadPickingListBtn) {
        downloadPickingListBtn.disabled = false;
        downloadPickingListBtn.style.opacity = '1';
        downloadPickingListBtn.style.cursor = 'pointer';
    }

    displayStatistics();
    showToast('已跳過上傳報表，進入預覽模式', 'success');
}

// ==================== 下載撿貨單 Excel ====================
async function downloadPickingList() {
    try {
        showLoading('正在產生撿貨單...');

        // 取得統計資料並排序
        let statsArray = Object.values(statistics);

        if (statsArray.length === 0) {
            hideLoading();
            showToast('沒有統計資料可下載', 'error');
            return;
        }

        // 按分類順序排序
        statsArray.sort((a, b) => {
            const orderA = getProductCategoryOrder(a.name);
            const orderB = getProductCategoryOrder(b.name);
            if (orderA !== orderB) return orderA - orderB;
            return (a.name || '').localeCompare(b.name || '', 'zh-TW');
        });

        // 使用 ExcelJS 建立工作簿
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('撿貨單');

        // 設定欄位標題（不包含「欄位」和「口味」）
        worksheet.columns = [
            { header: '報表商品', key: 'name', width: 30 },
            { header: '規格', key: 'spec', width: 15 },
            { header: '總數量', key: 'quantity', width: 12 }
        ];

        // 設定標題列樣式
        const headerRow = worksheet.getRow(1);
        headerRow.font = { bold: true, size: 12 };
        headerRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF4A5568' }
        };
        headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        headerRow.alignment = { horizontal: 'center', vertical: 'middle' };

        // 添加資料列
        statsArray.forEach(stat => {
            worksheet.addRow({
                name: stat.name || '',
                spec: stat.spec || '',
                quantity: stat.quantity || 0
            });
        });

        // 添加總計列
        const totalRow = worksheet.addRow({
            name: '總計',
            flavor: '',
            spec: `${statsArray.length} 種商品`,
            quantity: statsArray.reduce((sum, s) => sum + s.quantity, 0)
        });
        totalRow.font = { bold: true };
        totalRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFEDF2F7' }
        };

        // 設定邊框
        worksheet.eachRow((row, rowNumber) => {
            row.eachCell((cell) => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });
        });

        // 產生檔案並下載
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const today = new Date().toISOString().split('T')[0].replace(/-/g, '');
        a.download = `撿貨單_${today}.xlsx`;
        a.click();
        URL.revokeObjectURL(url);

        hideLoading();
        showToast('撿貨單下載完成！', 'success');
    } catch (error) {
        hideLoading();
        console.error('下載撿貨單失敗:', error);
        showToast('下載失敗：' + error.message, 'error');
    }
}
