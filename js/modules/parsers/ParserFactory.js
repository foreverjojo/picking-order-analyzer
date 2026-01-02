/**
 * @file ParserFactory.js
 * @description 解析器工廠模組
 * @author Ivy House TW Development Team
 */

import { parseMomoExcel } from './MomoParser.js';
import { parseOfficialExcel } from './OfficialParser.js';
import { parseShopeePDF } from './ShopeeParser.js';
import { parseOrangePointExcel } from './OrangePointParser.js';

/**
 * 預讀 Excel 內容以識別平台
 * @param {File} file 檔案物件
 * @returns {Promise<{isMomo: boolean}>} 平台識別結果
 */
async function previewExcelForPlatform(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(worksheet);

                // 檢查是否有商品編碼欄位且以 TP0001937 開頭 (MOMO 特徵)
                if (jsonData.length > 0 && jsonData[0]['商品編碼']) {
                    const firstProductCode = String(jsonData[0]['商品編碼'] || '');
                    if (firstProductCode.startsWith('TP0001937')) {
                        resolve({ isMomo: true });
                        return;
                    }
                }

                resolve({ isMomo: false });
            } catch (error) {
                console.warn('預讀 Excel 失敗，使用預設邏輯:', error);
                resolve({ isMomo: false });
            }
        };

        reader.onerror = () => resolve({ isMomo: false });
        reader.readAsArrayBuffer(file);
    });
}

/**
 * 根據檔案類型或名稱自動解析
 * @param {File} file 檔案物件
 * @returns {Promise<Array>} 解析後的商品
 */
export async function parseFile(file) {
    const fileName = file.name.toLowerCase();

    // 識別平台 - 優先使用檔名判斷
    if (fileName.includes('橘點子') || fileName.includes('jellytree') || (fileName.endsWith('.xls') && !fileName.endsWith('.xlsx'))) {
        return { platform: 'orangepoint', data: await parseOrangePointExcel(file) };
    } else if ((fileName.includes('momo') || fileName.includes('富邦') || fileName.includes('order_export')) && fileName.endsWith('.xlsx')) {
        return { platform: 'momo', data: await parseMomoExcel(file) };
    } else if (fileName.includes('官網') && fileName.endsWith('.xlsx')) {
        return { platform: 'official', data: await parseOfficialExcel(file) };
    } else if (fileName.endsWith('.pdf')) {
        return { platform: 'shopee', data: await parseShopeePDF(file) };
    } else if (fileName.endsWith('.xlsx')) {
        // 無法從檔名識別 - 嘗試從內容識別
        const preview = await previewExcelForPlatform(file);

        if (preview.isMomo) {
            console.log('根據商品編碼識別為 MOMO 撿貨單');
            return { platform: 'momo', data: await parseMomoExcel(file) };
        }

        // 預設官網
        return { platform: 'official', data: await parseOfficialExcel(file) };
    }

    throw new Error('不支援的檔案格式或無法識別平台: ' + file.name);
}
