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
 * 根據檔案類型或名稱自動解析
 * @param {File} file 檔案物件
 * @returns {Promise<Array>} 解析後的商品
 */
export async function parseFile(file) {
    const fileName = file.name.toLowerCase();

    // 識別平台
    if (fileName.includes('橘點子') || fileName.includes('jellytree') || (fileName.endsWith('.xls') && !fileName.endsWith('.xlsx'))) {
        return { platform: 'orangepoint', data: await parseOrangePointExcel(file) };
    } else if ((fileName.includes('momo') || fileName.includes('富邦') || fileName.includes('order_export')) && fileName.endsWith('.xlsx')) {
        return { platform: 'momo', data: await parseMomoExcel(file) };
    } else if (fileName.includes('官網') && fileName.endsWith('.xlsx')) {
        return { platform: 'official', data: await parseOfficialExcel(file) };
    } else if (fileName.endsWith('.pdf')) {
        return { platform: 'shopee', data: await parseShopeePDF(file) };
    } else if (fileName.endsWith('.xlsx')) {
        // 預設官網
        return { platform: 'official', data: await parseOfficialExcel(file) };
    }

    throw new Error('不支援的檔案格式或無法識別平台: ' + file.name);
}
