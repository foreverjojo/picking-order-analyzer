/**
 * @file MomoParser.js
 * @description MOMO 撿貨單解析模組
 * @author Ivy House TW Development Team
 */

/**
 * 解析 MOMO Excel 撿貨單
 * @param {File} file 檔案物件
 * @returns {Promise<Array>} 解析後的商品列表
 */
export async function parseMomoExcel(file) {
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
                        rawData: row
                    }))
                    .filter(p => p.name && p.quantity > 0);

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
