/**
 * @file analyze-excel.js
 * @description Excel 報表結構分析工具 (Node.js 腳本)
 * @purpose 用於開發階段分析 .xlsm 報表的工作表結構與欄位內容
 * @author Ivy House TW Development Team
 * @note 此為開發輔助工具，非前端功能
 */

// 分析 Excel 報表結構的腳本
const ExcelJS = require('exceljs');

async function analyzeExcel() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('c:/Users/forev/.gemini/antigravity/scratch/picking-order-analyzer/sample-data/20251222-新生產統計表.xlsm');

    console.log('=== 工作表列表 ===');
    workbook.worksheets.forEach((ws, index) => {
        console.log(`${index + 1}. ${ws.name}`);
    });

    // 分析第 1 個工作表
    console.log('\n=== 工作表 1 結構 ===');
    const sheet1 = workbook.worksheets[0];
    console.log(`名稱: ${sheet1.name}`);
    console.log(`行數: ${sheet1.rowCount}`);

    // 顯示前 50 行的 A 欄和 G 欄
    console.log('\n前 50 行 (A欄, G欄, H欄):');
    for (let i = 1; i <= Math.min(50, sheet1.rowCount); i++) {
        const row = sheet1.getRow(i);
        const colA = row.getCell(1).value;
        const colG = row.getCell(7).value;
        const colH = row.getCell(8).value;
        if (colA || colG || colH) {
            console.log(`Row ${i}: A="${colA}", G="${colG}", H="${colH}"`);
        }
    }

    // 分析第 2 個工作表
    console.log('\n=== 工作表 2 結構 ===');
    const sheet2 = workbook.worksheets[1];
    if (sheet2) {
        console.log(`名稱: ${sheet2.name}`);
        console.log(`行數: ${sheet2.rowCount}`);

        // 顯示前 50 行
        console.log('\n前 50 行 (A欄, B欄):');
        for (let i = 1; i <= Math.min(50, sheet2.rowCount); i++) {
            const row = sheet2.getRow(i);
            const colA = row.getCell(1).value;
            const colB = row.getCell(2).value;
            if (colA || colB) {
                console.log(`Row ${i}: A="${colA}", B="${colB}"`);
            }
        }
    } else {
        console.log('沒有第 2 個工作表');
    }
}

analyzeExcel().catch(console.error);
