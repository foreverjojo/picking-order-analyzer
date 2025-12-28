/**
 * @file ReportFiller.js
 * @description 報表填寫邏輯
 * @author Ivy House TW Development Team
 */

import { state } from '../core/StateManager.js';

/**
 * 使用 ExcelJS 填入資料（保留完整樣式）
 */
export async function fillTemplate() {
    if (!state.excelWorkbook) {
        throw new Error('ExcelJS workbook 未載入');
    }

    const worksheet = state.excelWorkbook.worksheets[0];
    const colToNum = { 'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5 };
    let filledCount = 0;

    // 遍歷工作表的每一行
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        const nameCell = row.getCell(1); // A欄
        if (!nameCell.value) return;

        const cellValue = String(nameCell.value).trim();

        // 遍歷所有統計項目尋找匹配
        Object.values(state.statistics).forEach(stat => {
            if (stat.name && cellValue === stat.name) {
                const colLetter = stat.column || 'B';
                const colNum = colToNum[colLetter] || 2;

                const targetCell = row.getCell(colNum);
                targetCell.value = stat.quantity;
                filledCount++;
            }
        });
    });

    console.log(`總共填入 ${filledCount} 筆資料`);
    return filledCount;
}

/**
 * 使用 JSZip 核心邏輯更新 XML 以確保樣式和宏完全保留
 */
export async function updateXmlViaZip(zip) {
    // 讀取 sheet1.xml
    const sheetPath = 'xl/worksheets/sheet1.xml';
    let sheetXml = await zip.file(sheetPath).async('string');

    // 建立需要更新的儲存格列表
    const ws = state.excelWorkbook.worksheets[0];
    const cellUpdates = [];
    const colToNum = { 'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5 };

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        const nameCell = row.getCell(1);
        if (!nameCell.value) return;
        const cellValue = String(nameCell.value).trim();

        Object.values(state.statistics).forEach(stat => {
            if (stat.name && cellValue === stat.name) {
                const colLetter = stat.column || 'B';
                const cellRef = `${colLetter}${rowNumber}`;
                cellUpdates.push({ cellRef, value: stat.quantity });
            }
        });
    });

    // XML 替換邏輯 (與 app.js 相同)
    cellUpdates.forEach(update => {
        const cellRef = update.cellRef;
        let matched = false;

        // 1. 公式 + 值
        const fvPattern = new RegExp(`(<c r="${cellRef}"[^>]*>)<f[^>]*>[\\s\\S]*?</f><v>[^<]*</v>(</c>)`);
        if (fvPattern.test(sheetXml)) {
            sheetXml = sheetXml.replace(fvPattern, `$1<v>${update.value}</v>$2`);
            matched = true;
        }

        // 2. 只有公式
        if (!matched) {
            const fPattern = new RegExp(`(<c r="${cellRef}"[^>]*>)<f[^>]*>[\\s\\S]*?</f>(</c>)`);
            if (fPattern.test(sheetXml)) {
                sheetXml = sheetXml.replace(fPattern, `$1<v>${update.value}</v>$2`);
                matched = true;
            }
        }

        // 3. 只有值
        if (!matched) {
            const vPattern = new RegExp(`(<c r="${cellRef}"[^>]*>)<v>[^<]*</v>(</c>)`);
            if (vPattern.test(sheetXml)) {
                sheetXml = sheetXml.replace(vPattern, `$1<v>${update.value}</v>$2`);
                matched = true;
            }
        }

        // 4. 空儲存格
        if (!matched) {
            const ePattern = new RegExp(`<c r="${cellRef}"([^/>]*)/>`);
            if (ePattern.test(sheetXml)) {
                sheetXml = sheetXml.replace(ePattern, `<c r="${cellRef}"$1><v>${update.value}</v></c>`);
                matched = true;
            }
        }
    });

    zip.file(sheetPath, sheetXml);

    // 強制重新計算
    if (zip.file('xl/calcChain.xml')) {
        zip.remove('xl/calcChain.xml');
        const contentTypesPath = '[Content_Types].xml';
        let contentTypesXml = await zip.file(contentTypesPath).async('string');
        contentTypesXml = contentTypesXml.replace(/<Override[^>]*calcChain[^>]*\/>/g, '');
        zip.file(contentTypesPath, contentTypesXml);
    }

    const workbookPath = 'xl/workbook.xml';
    let workbookXml = await zip.file(workbookPath).async('string');
    if (workbookXml.includes('<calcPr')) {
        workbookXml = workbookXml.replace(/fullCalcOnLoad="[^"]*"/g, '');
        workbookXml = workbookXml.replace(/<calcPr\s*/g, '<calcPr fullCalcOnLoad="1" ');
    } else {
        workbookXml = workbookXml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
    }
    zip.file(workbookPath, workbookXml);
}
