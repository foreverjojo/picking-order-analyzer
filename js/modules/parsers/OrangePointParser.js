/**
 * @file OrangePointParser.js
 * @description 橘點子撿貨單解析模組
 * @author Ivy House TW Development Team
 */

/**
 * 解析橘點子 Excel 撿貨單
 */
export async function parseOrangePointExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

                const products = [];
                const seenProducts = new Set();

                for (let row of jsonData) {
                    for (let cell of row) {
                        if (cell === '品名') continue;
                        const parsedItems = parseOrangePointCell(String(cell));
                        for (const parsed of parsedItems) {
                            const key = `${parsed.name}|${parsed.quantity}`;
                            if (!seenProducts.has(key)) {
                                seenProducts.add(key);
                                products.push({
                                    name: parsed.name,
                                    quantity: parsed.quantity,
                                    spec: '',
                                    source: '橘點子'
                                });
                            }
                        }
                    }
                }

                resolve(products);
            } catch (error) {
                reject(error);
            }
        };

        reader.onerror = () => reject(new Error('檔案讀取失敗'));
        reader.readAsArrayBuffer(file);
    });
}

/**
 * 解析橘點子單一欄位
 */
function parseOrangePointCell(cellValue) {
    if (!cellValue || typeof cellValue !== 'string') return [];

    const text = cellValue.trim();
    if (!text) return [];

    const results = [];

    // 過濾咖啡
    if (text.includes('咖啡') && !text.includes('咖啡小花')) {
        return [];
    }

    // 贈品格式
    if (text.startsWith('贈品:')) {
        const giftMatch = text.match(/[#](.+?)\s*[x×]?\s*(\d+)$/i);
        if (giftMatch) {
            const quantity = parseInt(giftMatch[2]);
            if (quantity === 0) return [];
            results.push({ name: giftMatch[1].trim(), quantity: quantity });
        }
        return results;
    }

    // 試吃格式
    if (text.startsWith('試吃:')) {
        const quantityMatch = text.match(/\s+(\d+)$/);
        if (!quantityMatch) return [];
        const totalQuantity = parseInt(quantityMatch[1]);
        if (totalQuantity === 0) return [];

        const productPart = text.replace(/^試吃:/, '').replace(/\s+\d+$/, '').trim();
        const productItems = productPart.split('+');
        for (const item of productItems) {
            const cleanName = item.replace(/\*\d+$/, '').trim();
            if (cleanName) {
                results.push({ name: cleanName, quantity: totalQuantity });
            }
        }
        return results;
    }

    // 普通格式
    const match = text.match(/^(.+?)\s+(\d+)$/);
    if (match) {
        const name = match[1].trim();
        const quantity = parseInt(match[2]);
        if (name.toLowerCase() === 'total' || quantity === 0) return [];
        results.push({ name: name, quantity: quantity });
    }

    return results;
}
