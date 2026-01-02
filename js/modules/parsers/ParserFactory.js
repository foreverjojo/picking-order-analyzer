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
 * @returns {Promise<{platform: string|null}>} 平台識別結果
 */
async function previewExcelForPlatform(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }); // 使用陣列格式讀取，解析更快更準確

                if (jsonData.length === 0) {
                    resolve({ platform: null });
                    return;
                }

                // 1. 檢查橘點子特徵 (通常在前幾行文字中出現)
                for (let i = 0; i < Math.min(5, jsonData.length); i++) {
                    const rowText = JSON.stringify(jsonData[i]);
                    if (rowText.includes('橘點子')) {
                        resolve({ platform: 'orangepoint' });
                        return;
                    }
                }

                // 2. 檢查官網特徵 (建表日期)
                for (let i = 0; i < Math.min(10, jsonData.length); i++) {
                    const rowText = JSON.stringify(jsonData[i]);
                    if (rowText.includes('建表日期')) {
                        resolve({ platform: 'official' });
                        return;
                    }
                }

                // 3. 檢查 MOMO 特徵 (商品編碼以 TP0001937 開頭)
                // 轉換回物件格式以檢查特定欄位，或直接在陣列中尋找
                const headerRow = jsonData.find(row => Array.isArray(row) && row.includes('商品編碼'));
                if (headerRow) {
                    const colIdx = headerRow.indexOf('商品編碼');
                    // 檢查前幾個資料列
                    const startIdx = jsonData.indexOf(headerRow) + 1;
                    for (let i = startIdx; i < Math.min(startIdx + 5, jsonData.length); i++) {
                        const code = String(jsonData[i][colIdx] || '');
                        if (code.startsWith('TP0001937')) {
                            resolve({ platform: 'momo' });
                            return;
                        }
                    }
                }

                resolve({ platform: null });
            } catch (error) {
                console.warn('預讀 Excel 失敗，使用預設邏輯:', error);
                resolve({ platform: null });
            }
        };

        reader.onerror = () => resolve({ platform: null });
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

    // 識別平台 - 優先使用檔名判斷 (這部速度最快)
    if (fileName.includes('橘點子') || fileName.includes('jellytree')) {
        return { platform: 'orangepoint', data: await parseOrangePointExcel(file) };
    } else if ((fileName.includes('momo') || fileName.includes('富邦') || fileName.includes('order_export')) && fileName.endsWith('.xlsx')) {
        return { platform: 'momo', data: await parseMomoExcel(file) };
    } else if (fileName.includes('官網') && fileName.endsWith('.xlsx')) {
        return { platform: 'official', data: await parseOfficialExcel(file) };
    } else if (fileName.endsWith('.pdf')) {
        return { platform: 'shopee', data: await parseShopeePDF(file) };
    } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
        // 無法從檔名識別 - 嘗試從內容識別
        const detection = await previewExcelForPlatform(file);

        if (detection.platform === 'momo') {
            console.log('智慧識別：根據內容判定為 MOMO');
            return { platform: 'momo', data: await parseMomoExcel(file) };
        } else if (detection.platform === 'official') {
            console.log('智慧識別：根據內容判定為 官網');
            return { platform: 'official', data: await parseOfficialExcel(file) };
        } else if (detection.platform === 'orangepoint') {
            console.log('智慧識別：根據內容判定為 橘點子');
            return { platform: 'orangepoint', data: await parseOrangePointExcel(file) };
        }

        // 預設邏輯
        if (fileName.endsWith('.xls')) {
            return { platform: 'orangepoint', data: await parseOrangePointExcel(file) };
        }
        return { platform: 'official', data: await parseOfficialExcel(file) };
    }

    throw new Error('不支援的檔案格式或無法識別平台: ' + file.name);
}

// 匯出內容偵測函數供 app.js 使用
export { previewExcelForPlatform };
