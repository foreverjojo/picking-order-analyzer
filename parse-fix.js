// 解析官網 Excel 撿貨單 - 修正版
// 官網檔案前7行是元數據（標題、統計資訊等），實際商品資料從第8行開始
async function parse官網Excel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];

                // 使用 range 選項從第8行開始讀取（跳過前7行元數據）
                const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                    range: 7  // 從第8行開始（0-indexed，所以是7）
                });

                console.log('官網原始資料:', jsonData);

                const products = jsonData
                    .filter(row => {
                        // 確保有商品名稱和數量欄位
                        return row['商品名稱'] && row['數量'];
                    })
                    .map(row => ({
                        name: (row['商品名稱'] || '').trim(),
                        quantity: parseInt(row['數量'] || 0),
                        source: '官網',
                        spec: (row['規格'] || '').trim(),
                        rawData: row
                    }))
                    .filter(p => p.name && p.quantity > 0);

                console.log('官網原始資料行數:', jsonData.length);
                console.log('官網過濾後商品數:', products.length);

                resolve(products);
            } catch (error) {
                console.error('官網解析錯誤:', error);
                reject(error);
            }
        };

        reader.onerror = () => reject(new Error('無法讀取官網檔案'));
        reader.readAsArrayBuffer(file);
    });
}
