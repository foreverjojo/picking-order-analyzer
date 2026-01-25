/**
 * @file app.js
 * @description 系統入口與事件協調
 * @author Ivy House TW Development Team
 */

import { state, resetState, setUploadedFile } from './modules/core/StateManager.js';
import { parseFile, previewExcelForPlatform } from './modules/parsers/ParserFactory.js';
import { consolidateAndMap, calculateStatistics } from './modules/core/LogicEngine.js';
import { downloadReport, downloadPickingList } from './modules/export/ExportManager.js';
import { showLoading, hideLoading, showToast } from './modules/ui/UIUtils.js';
import * as UI from './modules/ui/UIManager.js';

import { initializeData } from './modules/core/DataManager.js';
import { updateStandardNamesCache } from './modules/rules/StandardProducts.js';

// 初始化
document.addEventListener('DOMContentLoaded', init);

async function init() {
    try {
        await initializeData();
        updateStandardNamesCache();
    } catch (e) {
        showToast('資料庫初始化失敗，請檢查網路或伺服器', 'error');
    }

    setupGlobalEvents();
    setupFileEvents();
    setupActionEvents();
}

function setupGlobalEvents() {
    // PDF.js worker
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

    document.getElementById('resetBtn').addEventListener('click', () => {
        if (confirm('確定要重新開始嗎？')) location.reload();
    });
}

function setupFileEvents() {
    const fileInput = document.getElementById('fileInput');
    const uploadBox = document.getElementById('uploadBox');
    const addMoreBtn = document.getElementById('addMoreFilesBtn');
    const additionalInput = document.getElementById('additionalFileInput');
    const templateInput = document.getElementById('templateInput');

    uploadBox.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', (e) => handleFiles(e.target.files));

    uploadBox.addEventListener('dragover', (e) => { e.preventDefault(); uploadBox.classList.add('drag-over'); });
    uploadBox.addEventListener('dragleave', () => uploadBox.classList.remove('drag-over'));
    uploadBox.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadBox.classList.remove('drag-over');
        handleFiles(e.dataTransfer.files);
    });

    if (addMoreBtn) addMoreBtn.addEventListener('click', () => additionalInput.click());
    if (additionalInput) additionalInput.addEventListener('change', (e) => handleAdditionalFiles(e.target.files));
    if (templateInput) templateInput.addEventListener('change', handleTemplateUpload);

    // 委派移除事件
    document.getElementById('fileList').addEventListener('click', (e) => {
        if (e.target.classList.contains('file-remove')) {
            const platform = e.target.getAttribute('data-platform');
            state.uploadedFiles[platform] = null;
            UI.updateFileList();
        }
    });
}

function setupActionEvents() {
    document.getElementById('parseBtn').addEventListener('click', performParsing);
    document.getElementById('confirmMappingBtn').addEventListener('click', () => {
        UI.collectMappingFromUI();
        document.getElementById('templateInput').click();
    });
    document.getElementById('skipToStep3Btn').addEventListener('click', handleSkipToStep3);
    document.getElementById('generateReportBtn').addEventListener('click', async () => {
        showLoading('產生中...');
        try {
            await downloadReport();
            UI.navigateToStep(4);
            showToast('產生成功！', 'success');
        } catch (e) { showToast(e.message, 'error'); }
        hideLoading();
    });
    document.getElementById('downloadBtn').addEventListener('click', downloadReport);
    document.getElementById('downloadPickingListBtn').addEventListener('click', async () => {
        showLoading('產生中...');
        try {
            await downloadPickingList();
            showToast('下載成功！', 'success');
        } catch (e) { showToast(e.message, 'error'); }
        hideLoading();
    });
}

async function handleFiles(files) {
    let hasError = false;
    for (const file of Array.from(files)) {
        try {
            const { platform } = await identifyPlatform(file);
            if (platform) {
                state.uploadedFiles[platform] = file;
            } else {
                showToast(`無法識別檔案類型: ${file.name}`, 'warning');
                hasError = true;
            }
        } catch (e) {
            console.warn(e.message);
            showToast(`檔案處理失敗: ${file.name} - ${e.message}`, 'error');
            hasError = true;
        }
    }
    UI.updateFileList();
    if (!hasError && files.length > 0) {
        showToast('檔案已成功加入列表', 'success');
    }
}

async function identifyPlatform(file) {
    const fileName = file.name.toLowerCase();

    // 1. 優先使用檔名判斷（速度最快）
    if (fileName.includes('橘點子') || fileName.includes('jellytree')) {
        return { platform: 'orangepoint' };
    }
    if ((fileName.includes('momo') || fileName.includes('富邦') || fileName.includes('order_export')) && fileName.endsWith('.xlsx')) {
        return { platform: 'momo' };
    }
    if (fileName.includes('官網') && fileName.endsWith('.xlsx')) {
        return { platform: 'official' };
    }
    if (fileName.endsWith('.pdf')) {
        return { platform: 'shopee' };
    }
    if (fileName.includes('統計') || fileName.endsWith('.xlsm')) {
        return { platform: 'template' };
    }

    // 2. 無法從檔名識別 - 呼叫內容偵測（統一使用 ParserFactory 的邏輯）
    if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
        const detection = await previewExcelForPlatform(file);

        if (detection.platform) {
            console.log(`智慧識別：檔案「${file.name}」判定為 ${detection.platform}`);
            return { platform: detection.platform };
        }

        // 3. 內容也無法判斷 - 使用預設邏輯
        const defaultPlatform = fileName.endsWith('.xls') ? 'orangepoint' : 'official';
        console.log(`預設邏輯：檔案「${file.name}」預設為 ${defaultPlatform}`);
        return { platform: defaultPlatform };
    }

    throw new Error('未知格式');
}

async function performParsing() {
    showLoading('解析中...');
    try {
        const platforms = ['momo', 'official', 'shopee', 'orangepoint'];
        for (const p of platforms) {
            if (state.uploadedFiles[p]) {
                const { data } = await parseFile(state.uploadedFiles[p]);
                state.parsedData[p] = data;
            }
        }
        consolidateAndMap();
        UI.buildMappingTable();
        UI.navigateToStep(2);
        showToast('完成');
    } catch (e) { showToast(e.message, 'error'); }
    hideLoading();
}

async function handleAdditionalFiles(files) {
    showLoading('追加中...');
    for (const file of Array.from(files)) {
        try {
            const { platform, data } = await parseFile(file);
            state.parsedData[platform] = [...state.parsedData[platform], ...data];
        } catch (e) { showToast(e.message, 'error'); }
    }
    consolidateAndMap();
    UI.buildMappingTable();
    hideLoading();
}

async function handleTemplateUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    showLoading('載入模組...');
    try {
        state.templateBuffer = await file.arrayBuffer();
        state.excelWorkbook = new ExcelJS.Workbook();
        await state.excelWorkbook.xlsx.load(state.templateBuffer);

        calculateStatistics();
        UI.displayStatistics();
        UI.navigateToStep(3);

        // 啟用按鈕
        const genBtn = document.getElementById('generateReportBtn');
        genBtn.disabled = false;
        genBtn.style.pointerEvents = 'auto';
        genBtn.style.opacity = '1';
    } catch (e) { showToast(e.message, 'error'); }
    hideLoading();
}

function handleSkipToStep3() {
    UI.collectMappingFromUI();
    calculateStatistics();
    UI.displayStatistics();
    UI.navigateToStep(3);

    const genBtn = document.getElementById('generateReportBtn');
    genBtn.disabled = true;
    genBtn.style.pointerEvents = 'none';
    genBtn.style.opacity = '0.5';
    showToast('已跳過上傳報表範本');
}
