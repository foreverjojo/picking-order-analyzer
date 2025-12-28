/**
 * @file OfficialParser.js
 * @description 官網撿貨單解析模組
 * @author Ivy House TW Development Team
 */

/**
 * 解析官網 Excel 撿貨單
 * @param {File} file 檔案物件
 * @returns {Promise<Array>} 解析後的商品列表
 */
export async function parseOfficialExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { range: 6 });

                const products = jsonData.map(row => {
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

                resolve(products);
            } catch (error) {
                reject(error);
            }
        };

        reader.onerror = () => reject(new Error('無法讀取官網檔案'));
        reader.readAsArrayBuffer(file);
    });
}
