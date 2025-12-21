function parseShopeePDF(items) {
    // 步驟1: 按 Y 軸分組（同一行的項目 Y 軸相近）
    const rows = groupByY(items);

    console.log('📊 按 Y 軸分組後的行數:', rows.length);

    // 步驟2: 找出標題行（包含「商品名稱」、「規格」、「出貨數量」）
    let headerRow = null;
    let headerY = null;
    let columns = null;

    for (let row of rows) {
        const rowText = row.items.map(item => item.text).join(' ');
        if (rowText.includes('商品名稱') && rowText.includes('出貨數量')) {
            headerRow = row;
            headerY = row.y;
            columns = analyzeColumns(headerRow);
            console.log('✓ 找到標題行，Y軸:', headerY);
            console.log('📍 欄位位置:', columns);
            break;
        }
    }

    if (!headerRow || !columns) {
        console.warn('⚠️ 未找到標題行或無法分析欄位位置');
        return { validProducts: [], rows: rows };
    }

    // 步驟3: 找出所有商品行（包含「序」字且在標題行下方）
    const productRows = [];

    rows.forEach((row, index) => {
        // 只處理標題行下方的行（Y軸較小）
        if (row.y >= headerY) {
            return;
        }

        // 檢查是否包含「序」字（商品行的特徵）
        const hasSeq = row.items.some(item => item.text.match(/^序\d+$/));

        if (hasSeq) {
            productRows.push({
                rowIndex: index,
                row: row
            });
        }
    });

    console.log(`✓ 找到 ${productRows.length} 個商品行（包含「序」字）`);

    // 步驟4: 從每個商品行中提取資料
    const validProducts = [];

    productRows.forEach(({ rowIndex, row }) => {
        const product = extractProductFromRow(row, columns);

        if (product) {
            product.rowIndex = rowIndex;
            validProducts.push(product);
        }
    });

    console.log('✓ 成功提取的商品數:', validProducts.length);

    return {
        validProducts,
        rows
    };
}

// 按 Y 軸分組（容差：5 像素）
function groupByY(items, tolerance = 5) {
    const sorted = [...items].sort((a, b) => b.y - a.y); // Y軸從大到小排序
    const groups = [];

    sorted.forEach(item => {
        let found = false;

        for (let group of groups) {
            if (Math.abs(group.y - item.y) <= tolerance) {
                group.items.push(item);
                found = true;
                break;
            }
        }

        if (!found) {
            groups.push({
                y: item.y,
                items: [item]
            });
        }
    });

    // 每組內的項目按 X 軸排序
    groups.forEach(group => {
        group.items.sort((a, b) => a.x - b.x);
    });

    return groups;
}

// 分析標題行的欄位位置
function analyzeColumns(headerRow) {
    const columns = {
        seq: null,
        shop: null,
        productName: null,
        spec: null,
        quantity: null
    };

    headerRow.items.forEach(item => {
        const text = item.text;
        if (text.includes('序')) {
            columns.seq = item.x;
        } else if (text.includes('商店')) {
            columns.shop = item.x;
        } else if (text.includes('商品名稱')) {
            columns.productName = item.x;
        } else if (text.includes('規格')) {
            columns.spec = item.x;
        } else if (text.includes('出貨數量')) {
            columns.quantity = item.x;
        }
    });

    return columns;
}

// 從一行中提取商品資料
function extractProductFromRow(row, columns) {
    // X軸容差範圍
    const tolerance = 80;

    let seqNum = null;
    let nameItems = [];
    let specItems = [];
    let quantityValue = null;

    row.items.forEach(item => {
        const itemText = item.text.trim();

        // 序號
        if (columns.seq && Math.abs(item.x - columns.seq) < 30) {
            const match = itemText.match(/序(\d+)/);
            if (match) {
                seqNum = parseInt(match[1]);
            }
        }
        // 商品名稱（可能有多個文字項目需要合併）
        else if (columns.productName && Math.abs(item.x - columns.productName) < tolerance) {
            // 排除「序」字和純數字
            if (!itemText.match(/^序\d+$/) && !itemText.match(/^\d+$/)) {
                nameItems.push(itemText);
            }
        }
        // 規格
        else if (columns.spec && Math.abs(item.x - columns.spec) < tolerance) {
            if (!itemText.match(/^\d+$/)) {
                specItems.push(itemText);
            }
        }
        // 出貨數量
        else if (columns.quantity && Math.abs(item.x - columns.quantity) < tolerance) {
            const num = parseInt(itemText);
            if (!isNaN(num) && num > 0) {
                quantityValue = num;
            }
        }
    });

    // 組合商品名稱和規格
    const productName = nameItems.join(' ').trim();
    const productSpec = specItems.join(' ').trim();

    // 至少要有商品名稱或數量才算有效
    if (productName || quantityValue) {
        return {
            seq: seqNum,
            name: productName,
            spec: productSpec,
            quantity: quantityValue || 0,
            rawItems: row.items.length
        };
    }

    return null;
}
