// 自動映射規則分析結果

/*
通過分析用戶提供的 75 筆手動對應數據，我發現了以下映射規則：

## 商品名稱映射規則

1. **移除前綴**：【艾薇手工坊】
2. **簡化產品名稱**：
   - "夏威夷豆塔-XX口味" → "豆塔-XX" (移除"夏威夷"，移除"口味")
   - "堅果塔-XX口味" → "堅果塔-XX"
   - "經典杏仁瓦片-XX口味" → "瓦片-XX" (移除"經典杏仁")
   - "奶油曲奇餅乾-XX口味" → "奶油-XX"
   - "西點餅乾-XX" → "西點-XX"
   - "雪花餅-XX口味" → "雪花餅-XX"
   - "胖貝殼瑪德蓮-XX口味" → "瑪德蓮-XX"
   
3. **特殊對應**：
   - "椰棗夏威夷豆" → "椰棗豆子150g"
   - "椰棗XX" → "椰棗XX150g"
   - "無調味堅果-原味XX" → "無調味XX"
   - "艾薇金磚土鳳梨酥" → "土鳳梨酥(紅點)"
   - "夏威夷豆法式牛奶糖" → "牛奶糖"
   - "招牌雙塔組合" → "雙塔"
   
4. **禮盒系列**：直接提取禮盒名稱

## 規格映射規則

從 pickingName 或 pickingSpec 中提取關鍵字：

1. **袋裝數量**：
   - "10入" → 欄B (10入袋裝)
   - "15入" → 欄C (15入袋裝)
   - "8入" → 欄B (8入袋裝)
   - "12入" → 欄C (12入袋裝)

2. **重量**：
   - "90g" → 欄B
   - "135g" → 欄C
   - "120g" → 欄B
   - "280g" → 欄C
   - "200g" → 欄C
   - "50g" → 欄B
   - "45g" → 欄B (活動專用)
   - "150g" → 欄B
   - "300g" → 欄B

3. **特殊規格**：
   - "單顆"/"2入" → 欄D (單顆)
   - "禮盒" → 欄B

## 實現策略

1. **名稱正規化**：
   - 移除「【艾薇手工坊】」
   - 移除多餘空格和特殊符號
   - 套用簡化規則

2. **規格識別**：
   - 合併 pickingName 和 pickingSpec
   - 使用正則提取數量和重量
   - 根據範圍判斷對應欄位

3. **模糊匹配**：
   - 使用字串相似度算法
   - 關鍵詞匹配（塔、瓦片、餅乾等）
*/

// 實際映射函數
function autoMapProduct(pickingName, pickingSpec) {
    // 合併名稱和規格用於分析
    const fullText = `${pickingName} ${pickingSpec}`.toLowerCase();

    // 1. 移除前綴
    let productName = pickingName.replace(/【艾薇手工坊】/g, '').trim();

    // 2. 簡化產品名稱
    const simplifyRules = [
        { pattern: /夏威夷豆塔[|-](.+?)口味/, replacement: '豆塔-$1' },
        { pattern: /夏威夷豆塔[|-](.+?)\s/, replacement: '豆塔-$1' },
        { pattern: /堅果塔[|-](.+?)口味/, replacement: '堅果塔-$1' },
        { pattern: /經典杏仁瓦片[|-](.+?)口味/, replacement: '瓦片-$1' },
        { pattern: /杏仁瓦片[|-](.+?)口味/, replacement: '瓦片-$1' },
        { pattern: /奶油曲奇餅乾[|-](.+?)口味/, replacement: '奶油-$1' },
        { pattern: /奶油餅乾[|-](.+?)口味/, replacement: '奶油-$1' },
        { pattern: /西點餅乾[|-](.+?)\s/, replacement: '西點-$1' },
        { pattern: /雪花餅[|-](.+?)口味/, replacement: '雪花餅-$1' },
        { pattern: /胖貝殼瑪德蓮[|-](.+?)口味/, replacement: '瑪德蓮-$1' },
        { pattern: /招牌雙塔組合/, replacement: '雙塔' },
    ];

    for (let rule of simplifyRules) {
        if (rule.pattern.test(productName)) {
            productName = productName.replace(rule.pattern, rule.replacement);
            break;
        }
    }

    // 3. 特殊對應
    const specialMappings = {
        '椰棗夏威夷豆': '椰棗豆子150g',
        '椰棗核桃': '椰棗核桃150g',
        '椰棗腰果': '椰棗腰果150g',
        '椰棗杏仁': '椰棗杏仁150g',
        '夏威夷豆法式牛奶糖': '牛奶糖',
        '艾薇金磚土鳳梨酥': '土鳳梨酥(紅點)',
    };

    for (let [key, value] of Object.entries(specialMappings)) {
        if (productName.includes(key)) {
            productName = value;
            break;
        }
    }

    // 4. 提取規格並判斷欄位
    let column = null;
    let spec = null;

    // 檢查重量
    if (/90g/.test(fullText)) {
        column = 'B';
        spec = '90g';
    } else if (/135g/.test(fullText)) {
        column = 'C';
        spec = '135g';
    } else if (/45g/.test(fullText)) {
        column = 'B';
        spec = '45g';
    } else if (/120g/.test(fullText)) {
        column = 'B';
        spec = '120g';
    } else if (/280g/.test(fullText)) {
        column = 'C';
        spec = '280g';
    } else if (/200g/.test(fullText)) {
        column = 'C';
        spec = '200g';
    } else if (/150g/.test(fullText)) {
        column = 'B';
        spec = '150g';
    } else if (/50g/.test(fullText)) {
        column = 'B';
        spec = '50g';
    } else if (/300g/.test(fullText)) {
        column = 'B';
        spec = '300g';
    }
    // 檢查數量
    else if (/15入/.test(fullText)) {
        column = 'C';
        spec = '15入袋裝';
    } else if (/12入/.test(fullText)) {
        column = 'C';
        spec = '12入袋裝';
    } else if (/10入/.test(fullText)) {
        column = 'B';
        spec = '10入袋裝';
    } else if (/8入/.test(fullText)) {
        column = 'B';
        spec = '8入袋裝';
    }
    // 檢查單個
    else if (/單顆/.test(fullText) || /2入/.test(fullText) || /單個/.test(fullText)) {
        column = 'D';
        spec = '單顆';
    }
    // 禮盒
    else if (/禮盒/.test(fullText)) {
        column = 'B';
        spec = '禮盒';
    }

    return {
        templateProduct: productName,
        templateColumn: column,
        templateSpec: spec,
        confidence: column ? 0.8 : 0.3  // 信心度
    };
}

// 匯出函數
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { autoMapProduct };
}
