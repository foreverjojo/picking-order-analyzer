// 撿貨單分析系統 - 修正版
// 版本：v1.1 - 修正檔案解析邏輯

// 全域變數
let uploadedFiles = {
    momo: null,
    official: null,
    shopee: null,
    template: null
};

let parsedData = {
    momo: [],
    official: [],
    shopee: []
};

let allProducts = []; // 所有解析出的商品
let productMapping = {}; // 商品對應關係
let statistics = {}; // 統計結果
let templateWorkbook = null; // 報表範本

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

    // 按鈕事件
    parseBtn.addEventListener('click', parseAllFiles);
    autoMapBtn.addEventListener('click', autoMapProducts);
    loadMappingBtn.addEventListener('click', loadMappingRules);
    saveMappingBtn.addEventListener('click', saveMappingRules);
    confirmMappingBtn.addEventListener('click', confirmMapping);
    generateReportBtn.addEventListener('click', generateReport);
    downloadBtn.addEventListener('click', downloadReport);
    resetBtn.addEventListener('click', resetApplication);
}

// ==================== 檔案處理 ====================
function handleFileSelect(event) {
    const files = Array.from(event.target.files);

    files.forEach(file => {
        const fileName = file.name.toLowerCase();

        if (fileName.includes('momo') && fileName.endsWith('.xlsx')) {
            uploadedFiles.momo = file;
            addFileToList(file, 'MOMO 撿貨單', '📊');
        } else if (fileName.includes('官網') && fileName.endsWith('.xlsx')) {
            uploadedFiles.official = file;
            addFileToList(file, '官網撿貨單', '📊');
        } else if (fileName.includes('蝦皮') && fileName.endsWith('.pdf')) {
            uploadedFiles.shopee = file;
            addFileToList(file, '蝦皮撿貨單', '📄');
        } else if (fileName.includes('統計表') && (fileName.endsWith('.xlsm') || fileName.endsWith('.xlsx'))) {
            uploadedFiles.template = file;
            addFileToList(file, '報表範本', '📋');
        } else {
            // 嘗試根據副檔名判斷
            if (fileName.endsWith('.pdf')) {
                uploadedFiles.shopee = file;
                addFileToList(file, '蝦皮撿貨單', '📄');
            } else if (fileName.endsWith('.xlsm') || fileName.includes('統計')) {
                uploadedFiles.template = file;
                addFileToList(file, '報表範本', '📋');
            } else if (fileName.endsWith('.xlsx')) {
                // 讓用戶選擇這是哪個平台的檔案
                if (!uploadedFiles.momo) {
                    uploadedFiles.momo = file;
                    addFileToList(file, 'MOMO 撿貨單', '📊');
                } else if (!uploadedFiles.official) {
                    uploadedFiles.official = file;
                    addFileToList(file, '官網撿貨單', '📊');
                }
            }
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
    else if (label === '報表範本') uploadedFiles.template = null;

    updateParseButton();
}

function updateParseButton() {
    const parseBtn = document.getElementById('parseBtn');
    const hasAnyPickingOrder = uploadedFiles.momo || uploadedFiles.official || uploadedFiles.shopee;
    parseBtn.disabled = !hasAnyPickingOrder;
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

// 解析 MOMO Excel 撿貨單 - 修正版
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

// 解析官網 Excel 撿貨單 - 修正版
async function parseOfficialExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];

                // 官網檔案前7行是元數據（標題、統計資訊等），從第8行開始才是實際商品資料
                // 使用 range 選項從第8行開始讀取（0-indexed，所以是7）
                const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                    range: 7  // 跳過前7行元數據
                });

                console.log('官網原始資料:', jsonData);

                const products = jsonData
                    .filter(row => {
                        // 確保有商品名稱和數量欄位
                        return row['商品名稱'] && row['數量'];
                    })
                    .map(row => ({
                        name: (row['商品名稱'] || '').trim(),
                        quantity: parseInt(row['數量'] || 0),
                        source: '官網',
                        spec: (row['規格'] || '').trim(),
                        rawData: row  // 保留原始資料供調試用
                    }))
                    .filter(p => p.name && p.quantity > 0);

                console.log('官網原始資料行數:', jsonData.length);
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

// 解析蝦皮 PDF 撿貨單 - 改善版
async function parseShopeePDF(file) {
    return new Promise(async (resolve, reject) => {
        try {
            const arrayBuffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

            let allText = '';
            let allItems = [];

            // 提取所有頁面的文字項目
            for (let i = 1; i <= pdf.numPages; i++) {
                const page = await pdf.getPage(i);
                const textContent = await page.getTextContent();

                // 保留每個文字項目的位置資訊
                textContent.items.forEach(item => {
                    allItems.push({
                        text: item.str,
                        x: item.transform[4],
                        y: item.transform[5]
                    });
                });

                const pageText = textContent.items.map(item => item.str).join(' ');
                allText += pageText + '\n';
            }

            console.log('蝦皮 PDF 原始文字:', allText);
            console.log('蝦皮 PDF 文字項目:', allItems);

            // 解析文字內容
            const products = parseShopeeText(allText, allItems);

            if (products.length === 0) {
                console.warn('蝦皮 PDF 解析結果為空，可能需要手動檢查格式');
                showToast('蝦皮 PDF 解析可能不完整，請檢查結果', 'warning');
            }

            resolve(products);

        } catch (error) {
            reject(new Error('無法解析蝦皮 PDF：' + error.message));
        }
    });
}

function parseShopeeText(text, items) {
    const products = [];
    const lines = text.split('\n');

    // 蝦皮 PDF 格式：序 商店 商品名稱 規格 出貨數量 檢貨確認
    // 嘗試多種解析策略

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();

        // 跳過空行和標題行
        if (!line || line.includes('序 商店 商品名稱') || line.includes('檢貨確認')) {
            continue;
        }

        // 策略1：尋找符合「商品名稱 + 數量」模式的行
        // 假設格式：數字(序號) 商店名 【商品名稱】 規格 數量 檢貨標記
        const match = line.match(/【([^】]+)】.*?(\d+)\s*$/);
        if (match) {
            const name = match[1].trim();
            const quantity = parseInt(match[2]);
            if (name && quantity > 0) {
                products.push({
                    name: name,
                    quantity: quantity,
                    source: '蝦皮',
                    spec: ''
                });
                continue;
            }
        }

        // 策略2：尋找任何以數字結尾的行
        const parts = line.split(/\s+/);
        if (parts.length >= 3) {
            const lastPart = parts[parts.length - 1];
            const quantity = parseInt(lastPart);

            if (!isNaN(quantity) && quantity > 0 && quantity < 1000) {
                // 移除最後的數量，剩下的作為商品名稱
                const name = parts.slice(0, -1).join(' ').trim();

                // 過濾掉可能的標題和無效資料
                if (name.length > 2 &&
                    !name.includes('序') &&
                    !name.includes('商店') &&
                    !name.includes('檢貨')) {
                    products.push({
                        name: name,
                        quantity: quantity,
                        source: '蝦皮',
                        spec: ''
                    });
                }
            }
        }
    }

    console.log('蝦皮解析出的商品:', products);
    return products;
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
                    cellStyles: true,
                    bookVBA: true // 保留巨集
                });
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
function consolidateProducts() {
    allProducts = [];

    // 合併所有平台的商品
    [...parsedData.momo, ...parsedData.official, ...parsedData.shopee].forEach(product => {
        const existing = allProducts.find(p =>
            p.name === product.name && p.source === product.source
        );

        if (existing) {
            existing.quantity += product.quantity;
        } else {
            allProducts.push({ ...product });
        }
    });

    console.log('整合後的商品:', allProducts);
}

// ==================== 商品對應 ====================
function buildMappingTable() {
    const tbody = document.getElementById('mappingTableBody');
    tbody.innerHTML = '';

    allProducts.forEach((product, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${product.name}</td>
            <td><span style="color: ${getSourceColor(product.source)}">${product.source}</span></td>
            <td>${product.quantity}</td>
            <td style="text-align: center;">→</td>
            <td>
                <input type="text" 
                       id="mapped-name-${index}" 
                       value="${product.name}" 
                       placeholder="報表中的商品名稱">
            </td>
            <td>
                <select id="mapped-category-${index}">
                    <option value="">選擇分類</option>
                    <option value="訂購基底類">訂購基底類</option>
                    <option value="蛋糕類別表單">蛋糕類別表單</option>
                    <option value="10個裝">10個裝</option>
                    <option value="15個裝">15個裝</option>
                    <option value="10入裝">10入裝</option>
                    <option value="小包裝">小包裝</option>
                    <option value="果乾類">果乾類</option>
                    <option value="其他">其他</option>
                </select>
            </td>
        `;
        tbody.appendChild(row);
    });
}

function getSourceColor(source) {
    const colors = {
        'MOMO': '#f59e0b',
        '官網': '#10b981',
        '蝦皮': '#ef4444'
    };
    return colors[source] || '#cbd5e1';
}

// 自動智慧對應 - 增強版
function autoMapProducts() {
    showLoading('智慧對應中...');

    setTimeout(() => {
        allProducts.forEach((product, index) => {
            const mappedNameInput = document.getElementById(`mapped-name-${index}`);
            const categorySelect = document.getElementById(`mapped-category-${index}`);

            // 智慧對應邏輯
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
            } else if (name.includes('塔') || name.includes('夏威夷豆')) {
                categorySelect.value = '10入裝';
            } else if (name.includes('瓦片') || name.includes('杏仁')) {
                categorySelect.value = '小包裝';
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
        // 建立對應關係
        productMapping = {};
        statistics = {};

        allProducts.forEach((product, index) => {
            const mappedName = document.getElementById(`mapped-name-${index}`).value.trim();
            const category = document.getElementById(`mapped-category-${index}`).value;

            if (!mappedName || !category) {
                return;
            }

            productMapping[product.name] = {
                reportName: mappedName,
                category: category
            };

            // 統計數量
            if (!statistics[category]) {
                statistics[category] = {};
            }

            if (!statistics[category][mappedName]) {
                statistics[category][mappedName] = 0;
            }

            statistics[category][mappedName] += product.quantity;
        });

        console.log('統計結果:', statistics);

        hideLoading();

        // 切換到步驟 3
        document.getElementById('step2').classList.add('hidden');
        document.getElementById('step3').classList.remove('hidden');

        // 顯示統計結果
        displayStatistics();
    }, 500);
}

// ==================== 統計顯示 ====================
function displayStatistics() {
    const container = document.getElementById('statsContainer');
    container.innerHTML = '';

    Object.keys(statistics).forEach(category => {
        const section = document.createElement('div');
        section.className = 'category-section';

        const title = document.createElement('div');
        title.className = 'category-title';
        title.textContent = category;
        section.appendChild(title);

        const grid = document.createElement('div');
        grid.className = 'product-grid';

        Object.keys(statistics[category]).forEach(productName => {
            const quantity = statistics[category][productName];

            const card = document.createElement('div');
            card.className = 'product-card';
            card.innerHTML = `
                <div class="product-name">${productName}</div>
                <div class="product-quantity">數量：<strong>${quantity}</strong></div>
            `;
            grid.appendChild(card);
        });

        section.appendChild(grid);
        container.appendChild(section);
    });
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
    // 這裡實作填入報表的邏輯
    // 由於報表結構複雜，這裡提供基本框架

    const sheetName = templateWorkbook.SheetNames[0];
    const worksheet = templateWorkbook.Sheets[sheetName];

    // 遍歷工作表，尋找商品名稱並填入數量
    const range = XLSX.utils.decode_range(worksheet['!ref']);

    for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            const cell = worksheet[cellAddress];

            if (cell && cell.v) {
                const cellValue = String(cell.v).trim();

                // 在統計結果中尋找匹配的商品
                Object.keys(statistics).forEach(category => {
                    Object.keys(statistics[category]).forEach(productName => {
                        if (cellValue.includes(productName) || productName.includes(cellValue)) {
                            // 找到商品，在右側幾欄填入數量
                            // 這裡假設數量在商品名稱右側的欄位
                            const quantityCell = XLSX.utils.encode_cell({ r: R, c: C + 1 });
                            if (!worksheet[quantityCell]) {
                                worksheet[quantityCell] = {};
                            }
                            worksheet[quantityCell].v = statistics[category][productName];
                            worksheet[quantityCell].t = 'n';
                        }
                    });
                });
            }
        }
    }
}

function downloadReport() {
    try {
        const wbout = XLSX.write(templateWorkbook, {
            bookType: 'xlsm',
            type: 'array',
            bookVBA: true // 保留巨集
        });

        const blob = new Blob([wbout], { type: 'application/vnd.ms-excel.sheet.macroEnabled.12' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;

        const today = new Date().toISOString().split('T')[0].replace(/-/g, '');
        a.download = `生產統計表_${today}.xlsm`;
        a.click();
        URL.revokeObjectURL(url);

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
