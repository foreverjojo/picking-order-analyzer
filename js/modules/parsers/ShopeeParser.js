/**
 * @file ShopeeParser.js
 * @description 蝦皮 PDF 撿貨單解析模組
 * @author Ivy House TW Development Team
 */

/**
 * 解析蝦皮 PDF 撿貨單
 * @param {File} file 檔案物件
 * @returns {Promise<Array>} 解析後的商品列表
 */
export async function parseShopeePDF(file) {
    return new Promise(async (resolve, reject) => {
        try {
            const arrayBuffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

            const allParsedProducts = [];
            for (let i = 1; i <= pdf.numPages; i++) {
                const page = await pdf.getPage(i);
                const textContent = await page.getTextContent();
                const pageItems = textContent.items.map(item => ({
                    text: item.str.trim(),
                    x: Math.round(item.transform[4]),
                    y: Math.round(item.transform[5])
                }));

                const products = parseShopeePage(pageItems);
                allParsedProducts.push(...products);
            }

            resolve(allParsedProducts.map(p => ({
                name: p.name,
                quantity: p.quantity,
                spec: p.spec,
                source: '蝦皮'
            })));

        } catch (error) {
            reject(new Error('無法解析蝦皮 PDF：' + error.message));
        }
    });
}

/**
 * 內部頁面解析邏輯
 */
function parseShopeePage(items) {
    const rows = groupByY(items);
    const columnRanges = {
        seq: { min: 32, max: 59 },
        productName: { min: 166, max: 307 },
        spec: { min: 308, max: 396 },
        quantity: { min: 397, max: 520 }
    };

    let headerY = null;
    for (let row of rows) {
        if (row.items.some(i => i.text.includes('商品名稱'))) {
            headerY = row.y;
            break;
        }
    }

    if (!headerY) return [];

    const products = [];
    let currentProduct = null;

    rows.forEach(row => {
        if (row.y >= headerY) return;

        let seqNum = null;
        row.items.forEach(item => {
            if (item.x >= columnRanges.seq.min && item.x <= columnRanges.seq.max) {
                const num = parseInt(item.text.trim());
                if (!isNaN(num) && num > 0) seqNum = num;
            }
        });

        if (seqNum) {
            if (currentProduct) products.push(currentProduct);
            currentProduct = { seq: seqNum, name: [], spec: [], quantity: null };
        }

        if (currentProduct) {
            const rowText = row.items.map(i => i.text).join(' ');
            if (rowText.includes('合計')) return;

            row.items.forEach(item => {
                const x = item.x;
                const t = item.text.trim();

                if (x >= columnRanges.productName.min && x <= columnRanges.productName.max) {
                    if (!t.match(/^\d+$/)) currentProduct.name.push(t);
                } else if (x >= columnRanges.spec.min && x <= columnRanges.spec.max) {
                    if (!t.match(/^\d+$/)) currentProduct.spec.push(t);
                } else if (x >= columnRanges.quantity.min && x <= columnRanges.quantity.max) {
                    const num = parseInt(t);
                    if (!isNaN(num) && num > 0) currentProduct.quantity = num;
                }
            });
        }
    });

    if (currentProduct) products.push(currentProduct);

    return products.filter(p => p.seq && p.seq > 0).map(p => ({
        name: p.name.join(' '),
        spec: p.spec.join(' '),
        quantity: p.quantity || 0
    }));
}

/**
 * 按 Y 座標分組
 */
function groupByY(items, tolerance = 5) {
    const sorted = [...items].sort((a, b) => b.y - a.y);
    const groups = [];
    sorted.forEach(item => {
        let found = false;
        for (let g of groups) {
            if (Math.abs(g.y - item.y) <= tolerance) {
                g.items.push(item);
                found = true;
                break;
            }
        }
        if (!found) groups.push({ y: item.y, items: [item] });
    });
    groups.forEach(g => g.items.sort((a, b) => a.x - b.x));
    return groups;
}
