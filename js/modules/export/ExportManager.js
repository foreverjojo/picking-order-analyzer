/**
 * @file ExportManager.js
 * @description 導出管理模組
 * @author Ivy House TW Development Team
 */

import { state } from '../core/StateManager.js';
import { updateXmlViaZip } from './ReportFiller.js';
import { getProductSortOrder } from '../rules/MappingEngine.js';

/**
 * 下載生產統計表 (Excel)
 */
export async function downloadReport() {
    try {
        const zip = await JSZip.loadAsync(state.templateBuffer);
        await updateXmlViaZip(zip);

        const content = await zip.generateAsync({
            type: 'blob',
            mimeType: 'application/vnd.ms-excel.sheet.macroEnabled.12'
        });

        const url = URL.createObjectURL(content);
        const a = document.createElement('a');
        a.href = url;
        const today = new Date().toISOString().split('T')[0].replace(/-/g, '');
        a.download = `生產統計表_${today}.xlsm`;
        a.click();
        URL.revokeObjectURL(url);

        return true;
    } catch (error) {
        console.error('下載報表失敗:', error);
        throw error;
    }
}

/**
 * 下載撿貨單 (Excel)
 */
export async function downloadPickingList() {
    try {
        const statsArray = Object.values(state.statistics);
        if (statsArray.length === 0) throw new Error('沒有統計資料');

        // 排序
        statsArray.sort((a, b) => {
            const orderA = getProductSortOrder().indexOf(a.name);
            const orderB = getProductSortOrder().indexOf(b.name);
            const rankA = orderA === -1 ? 9999 : orderA;
            const rankB = orderB === -1 ? 9999 : orderB;
            if (rankA !== rankB) return rankA - rankB;
            return a.name.localeCompare(b.name, 'zh-TW');
        });

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('撿貨單');

        worksheet.columns = [
            { header: '報表商品', key: 'name', width: 30 },
            { header: '規格', key: 'spec', width: 15 },
            { header: '總數量', key: 'quantity', width: 12 }
        ];

        // 樣式設定
        const headerRow = worksheet.getRow(1);
        headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4A5568' } };
        headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        headerRow.alignment = { horizontal: 'center' };

        statsArray.forEach(stat => {
            worksheet.addRow({
                name: stat.name,
                spec: stat.spec,
                quantity: stat.quantity
            });
        });

        // 總計
        const totalRow = worksheet.addRow({
            name: '總計',
            spec: `${statsArray.length} 種商品`,
            quantity: statsArray.reduce((sum, s) => sum + s.quantity, 0)
        });
        totalRow.font = { bold: true };
        totalRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDF2F7' } };

        worksheet.eachRow(row => {
            row.eachCell(cell => {
                cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            });
        });

        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const today = new Date().toISOString().split('T')[0].replace(/-/g, '');
        a.download = `撿貨單_${today}.xlsx`;
        a.click();
        URL.revokeObjectURL(url);

        return true;
    } catch (error) {
        console.error('下載撿貨單失敗:', error);
        throw error;
    }
}
