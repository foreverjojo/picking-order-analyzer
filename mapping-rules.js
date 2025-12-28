/**
 * @file mapping-rules.js
 * @description 智慧商品映射規則引擎
 * @purpose 定義撿貨單商品名稱到報表標準名稱的映射規則，支援模糊匹配與欄位/規格判斷
 * @author Ivy House TW Development Team
 */

// 從 test-auto-mapping.html 同步的完整映射規則

// 標準報表商品名稱列表
const standardProductNames = [
    // 塔類
    '豆塔-蔓越莓', '豆塔-焦糖', '豆塔-巧克力', '豆塔-抹茶', '豆塔-椒麻', '豆塔-綜合',
    '堅果塔-蜂蜜', '堅果塔-焦糖', '堅果塔-巧克力', '堅果塔-海苔', '堅果塔-咖哩', '堅果塔-綜合',
    '雙塔',
    // 鳳梨酥
    '土鳳梨酥(紅點)', '鳳凰酥',
    // 雪花餅
    '雪花餅-綜合', '雪花餅-蔓越莓', '雪花餅-巧克力', '雪花餅-金沙', '雪花餅-抹茶', '雪花餅-肉鬆',
    // 瑪德蓮
    '瑪德蓮-綜合', '瑪德蓮-蜂蜜', '瑪德蓮-巧克力', '瑪德蓮-紅茶', '瑪德蓮-抹茶', '瑪德蓮-柑橘', '瑪德蓮-檸檬',
    // 無調味堅果
    '無調味綜合堅果', '無調味夏威夷豆', '無調味腰果', '無調味杏仁', '無調味核桃',
    // 椰棗
    '★中東椰棗300g', '椰棗豆子150g', '椰棗腰果150g', '椰棗杏仁150g', '椰棗核桃150g',
    // 其他
    '牛奶糖', '南棗核桃糕', '牛奶糖-50g',
    // 瓦片
    '瓦片-綜合', '瓦片-原味', '瓦片-紅茶', '瓦片-巧克力', '瓦片-海苔', '瓦片-抹茶', '瓦片-黑糖', '瓦片-青花椒', '瓦片-原味45克',
    // 奶油
    '奶油-焦糖牛奶', '奶油-法國巧克力', '奶油-蜂蜜檸檬', '奶油-伯爵紅茶', '奶油-抹茶',
    // 西點
    '千層-小酥條', '西點-綜合', '西點-巧克力貝殼', '西點-咖啡小花', '西點-藍莓小花', '西點-蔓越莓貝殼', '西點-乳酪酥條',
    // 禮盒
    '雙塔禮盒', '蔓越莓禮盒', '綜豆禮盒', '綜堅禮盒', '晴空塔餅禮盒', '暖暖幸福禮盒', '臻愛時光禮盒', '濃情滿載禮盒',
    '浪漫詩篇禮盒', '戀戀雪花禮盒', '午後漫步禮盒', '那年花開禮盒', '花間逸韻禮盒',
    '輕-香榭漫遊禮盒', '輕-晨曦物語禮盒', '輕-金緻典藏禮盒', '輕-月光序曲禮盒'
];

// 標準化商品名稱
function normalizeProductName(name) {
    if (!name) return null;
    const cleaned = name.replace(/\s+/g, '');

    // 精確匹配
    for (let stdName of standardProductNames) {
        if (cleaned === stdName.replace(/\s+/g, '')) return stdName;
    }

    // 最長匹配
    let bestMatch = null;
    let bestMatchLength = 0;
    for (let stdName of standardProductNames) {
        const stdCleaned = stdName.replace(/\s+/g, '');
        if (cleaned.includes(stdCleaned) && stdCleaned.length > bestMatchLength) {
            bestMatch = stdName;
            bestMatchLength = stdCleaned.length;
        }
    }

    return bestMatch || name;
}

// 根據平台選擇映射函數
function autoMapProduct(pickingName, pickingSpec, quantity, platform = 'shopee') {
    // 優先處理已拆分商品（不管任何平台）
    if (/活動拆分/.test(pickingName)) {
        let flavor = null;
        if (/蜂蜜/.test(pickingName)) flavor = '瑪德蓮-蜂蜜';
        else if (/巧克力/.test(pickingName)) flavor = '瑪德蓮-巧克力';

        if (flavor) {
            console.log(`拆分商品映射: ${pickingName} → ${flavor}, C欄, 單顆, 數量 ${quantity}`);
            return {
                templateProduct: flavor,
                templateColumn: 'C',
                templateSpec: '單顆',
                multiplier: 1,
                mappedQuantity: quantity,
                confidence: 0.95
            };
        }
    }

    // 處理組合拆分商品（豆塔組合包）
    if (/組合拆分/.test(pickingName)) {
        let flavor = null;
        if (/蔓越莓/.test(pickingName)) flavor = '豆塔-蔓越莓';
        else if (/焦糖/.test(pickingName)) flavor = '豆塔-焦糖';

        if (flavor) {
            console.log(`組合拆分映射: ${pickingName} → ${flavor}, B欄, 10入袋裝, 數量 ${quantity}`);
            return {
                templateProduct: flavor,
                templateColumn: 'B',
                templateSpec: '10入袋裝',
                multiplier: 1,
                mappedQuantity: quantity,
                confidence: 0.95
            };
        }
    }

    if (platform === 'momo' || platform === 'official') {
        return autoMapProductMomo(pickingName, pickingSpec, quantity);
    } else if (platform === 'orangepoint') {
        return autoMapProductOrangePoint(pickingName, quantity);
    }
    return autoMapProductShopee(pickingName, pickingSpec, quantity);
}

// MOMO/官網專用映射函數
function autoMapProductMomo(pickingName, pickingSpec, quantity) {
    const fullText = `${pickingName} ${pickingSpec}`;
    const fullTextNoSpace = fullText.replace(/\s+/g, '');
    const nameNoSpace = pickingName.replace(/\s+/g, '');
    const specNoSpace = (pickingSpec || '').replace(/\s+/g, '');

    // 過濾（禮盒不應被過濾）
    const isGiftBox = /禮盒/.test(fullTextNoSpace);
    if (/咖啡/.test(fullTextNoSpace) && !/咖啡小花/.test(fullTextNoSpace) && !isGiftBox) {
        return { templateProduct: null, templateColumn: null, templateSpec: null, multiplier: 1, mappedQuantity: 0, confidence: 0 };
    }
    if (/提袋加購/.test(fullTextNoSpace) || /運費/.test(fullTextNoSpace)) {
        return { templateProduct: null, templateColumn: null, templateSpec: null, multiplier: 1, mappedQuantity: 0, confidence: 0 };
    }

    let productName = null;
    let column = null;
    let spec = null;
    let multiplier = 1;
    let confidence = 0.3;

    // 提取倍數（如 x3包 或 2盒）
    const multiplierMatchPack = fullText.match(/x(\d+)包/);
    const multiplierMatchBox = specNoSpace.match(/^(\d+)盒/);
    if (multiplierMatchPack) {
        multiplier = parseInt(multiplierMatchPack[1]);
    } else if (multiplierMatchBox) {
        multiplier = parseInt(multiplierMatchBox[1]);
    }

    // === 禮盒類型 ===
    if (/禮盒/.test(fullTextNoSpace)) {
        column = 'B'; spec = '禮盒'; confidence = 0.95;
        if (/雙塔禮盒/.test(fullTextNoSpace) || /招牌雙塔/.test(fullTextNoSpace)) {
            productName = '雙塔禮盒';
        } else if (/夏威夷豆塔禮盒/.test(fullTextNoSpace) && /蔓越莓/.test(fullTextNoSpace)) {
            productName = '蔓越莓禮盒';
        } else if (/夏威夷豆塔禮盒/.test(fullTextNoSpace) && /綜合/.test(fullTextNoSpace)) {
            productName = '綜豆禮盒';
        } else if (/堅果塔禮盒/.test(fullTextNoSpace) && /綜合/.test(fullTextNoSpace)) {
            productName = '綜堅禮盒';
        } else if (/戀戀雪花禮盒|戀戀雪花|新戀戀雪花/.test(fullTextNoSpace)) {
            productName = '戀戀雪花禮盒';
        } else if (/浪漫詩篇禮盒|浪漫詩篇|新浪漫詩篇/.test(fullTextNoSpace)) {
            productName = '浪漫詩篇禮盒';
        } else if (/暖暖幸福禮盒|暖暖幸福/.test(fullTextNoSpace)) {
            productName = '暖暖幸福禮盒';
        } else if (/臻愛時光禮盒|臻愛時光/.test(fullTextNoSpace)) {
            productName = '臻愛時光禮盒';
        } else if (/濃情滿載禮盒|濃情滿載/.test(fullTextNoSpace)) {
            productName = '濃情滿載禮盒';
        } else if (/午後漫步禮盒|午後漫步/.test(fullTextNoSpace)) {
            productName = '午後漫步禮盒';
        } else if (/那年花開禮盒|那年花開/.test(fullTextNoSpace)) {
            productName = '那年花開禮盒';
        } else if (/花間逸韻禮盒|花間逸韻/.test(fullTextNoSpace)) {
            productName = '花間逸韻禮盒';
        } else if (/晴空塔餅禮盒|晴空塔餅/.test(fullTextNoSpace)) {
            productName = '晴空塔餅禮盒';
        } else if (/金緻典藏禮盒|金緻典藏/.test(fullTextNoSpace)) {
            productName = '輕-金緻典藏禮盒';
        } else if (/香榭漫遊禮盒|香榭漫遊/.test(fullTextNoSpace)) {
            productName = '輕-香榭漫遊禮盒';
        } else if (/晨曦物語禮盒|晨曦物語/.test(fullTextNoSpace)) {
            productName = '輕-晨曦物語禮盒';
        } else if (/月光序曲禮盒|月光序曲/.test(fullTextNoSpace)) {
            productName = '輕-月光序曲禮盒';
        } else if (/蔓越莓禮盒/.test(fullTextNoSpace)) {
            productName = '蔓越莓禮盒';
        } else if (/綜豆禮盒/.test(fullTextNoSpace)) {
            productName = '綜豆禮盒';
        } else if (/綜堅禮盒/.test(fullTextNoSpace)) {
            productName = '綜堅禮盒';
        }
    }

    // === 豆塔 ===
    if (!productName && /夏威夷豆塔/.test(fullTextNoSpace)) {
        if (/蜂蜜蔓越莓|蔓越莓/.test(fullTextNoSpace)) productName = '豆塔-蔓越莓';
        else if (/焦糖/.test(fullTextNoSpace)) productName = '豆塔-焦糖';
        else if (/巧克力/.test(fullTextNoSpace)) productName = '豆塔-巧克力';
        else if (/抹茶/.test(fullTextNoSpace)) productName = '豆塔-抹茶';
        else if (/椒麻/.test(fullTextNoSpace)) productName = '豆塔-椒麻';
        else if (/綜合/.test(fullTextNoSpace)) productName = '豆塔-綜合';
        confidence = 0.9;
    }

    // === 堅果塔 ===
    if (!productName && /堅果塔/.test(fullTextNoSpace)) {
        if (/蜂蜜/.test(fullTextNoSpace)) productName = '堅果塔-蜂蜜';
        else if (/焦糖/.test(fullTextNoSpace)) productName = '堅果塔-焦糖';
        else if (/巧克力/.test(fullTextNoSpace)) productName = '堅果塔-巧克力';
        else if (/海苔/.test(fullTextNoSpace)) productName = '堅果塔-海苔';
        else if (/咖哩/.test(fullTextNoSpace)) productName = '堅果塔-咖哩';
        else if (/綜合/.test(fullTextNoSpace)) productName = '堅果塔-綜合';
        confidence = 0.9;
    }

    // === 雪花餅 ===
    if (!productName && /雪花餅/.test(fullTextNoSpace)) {
        if (/蔓越莓/.test(fullTextNoSpace)) productName = '雪花餅-蔓越莓';
        else if (/巧克力/.test(fullTextNoSpace)) productName = '雪花餅-巧克力';
        else if (/金沙/.test(fullTextNoSpace)) productName = '雪花餅-金沙';
        else if (/抹茶/.test(fullTextNoSpace)) productName = '雪花餅-抹茶';
        else if (/肉鬆/.test(fullTextNoSpace)) productName = '雪花餅-肉鬆';
        else if (/綜合/.test(fullTextNoSpace)) productName = '雪花餅-綜合';
        confidence = 0.9;
    }

    // === 瑪德蓮 ===
    if (!productName && /瑪德蓮/.test(fullTextNoSpace)) {
        if (/抹茶/.test(fullTextNoSpace)) productName = '瑪德蓮-抹茶';
        else if (/檸檬/.test(fullTextNoSpace)) productName = '瑪德蓮-檸檬';
        else if (/紅茶/.test(fullTextNoSpace)) productName = '瑪德蓮-紅茶';
        else if (/蜂蜜/.test(fullTextNoSpace)) productName = '瑪德蓮-蜂蜜';
        else if (/巧克力/.test(fullTextNoSpace)) productName = '瑪德蓮-巧克力';
        else if (/柑橘/.test(fullTextNoSpace)) productName = '瑪德蓮-柑橘';
        else if (/綜合/.test(fullTextNoSpace)) productName = '瑪德蓮-綜合';
        confidence = 0.9;
    }

    // === 瓦片 ===
    if (!productName && /瓦片|杏仁瓦片/.test(fullTextNoSpace)) {
        if (/原味/.test(fullTextNoSpace) && /45g/.test(fullTextNoSpace)) productName = '瓦片-原味45克';
        else if (/原味/.test(fullTextNoSpace)) productName = '瓦片-原味';
        else if (/抹茶/.test(fullTextNoSpace)) productName = '瓦片-抹茶';
        else if (/紅茶/.test(fullTextNoSpace)) productName = '瓦片-紅茶';
        else if (/巧克力/.test(fullTextNoSpace)) productName = '瓦片-巧克力';
        else if (/海苔/.test(fullTextNoSpace)) productName = '瓦片-海苔';
        else if (/黑糖/.test(fullTextNoSpace)) productName = '瓦片-黑糖';
        else if (/青花椒/.test(fullTextNoSpace)) productName = '瓦片-青花椒';
        else if (/綜合/.test(fullTextNoSpace)) productName = '瓦片-綜合';
        confidence = 0.9;
    }

    // === 奶油 ===
    if (!productName && /奶油曲奇|奶油餅乾/.test(fullTextNoSpace)) {
        if (/抹茶|小山園抹茶/.test(fullTextNoSpace)) productName = '奶油-抹茶';
        else if (/焦糖牛奶/.test(fullTextNoSpace)) productName = '奶油-焦糖牛奶';
        else if (/法國巧克力|巧克力/.test(fullTextNoSpace)) productName = '奶油-法國巧克力';
        else if (/蜂蜜檸檬/.test(fullTextNoSpace)) productName = '奶油-蜂蜜檸檬';
        else if (/伯爵紅茶/.test(fullTextNoSpace)) productName = '奶油-伯爵紅茶';
        confidence = 0.9;
    }

    // === 西點 ===
    if (!productName && /西點餅乾/.test(fullTextNoSpace)) {
        if (/乳酪酥條/.test(fullTextNoSpace)) productName = '西點-乳酪酥條';
        else if (/蔓越莓貝殼/.test(fullTextNoSpace)) productName = '西點-蔓越莓貝殼';
        else if (/藍莓小花/.test(fullTextNoSpace)) productName = '西點-藍莓小花';
        else if (/咖啡小花/.test(fullTextNoSpace)) productName = '西點-咖啡小花';
        else if (/巧克力貝殼/.test(fullTextNoSpace)) productName = '西點-巧克力貝殼';
        else if (/綜合/.test(fullTextNoSpace)) productName = '西點-綜合';
        confidence = 0.9;
    }

    // === 千層小酥條 ===
    if (!productName && /千層.*小酥條|小酥條/.test(fullTextNoSpace)) {
        productName = '千層-小酥條';
        column = 'B';
        spec = '小包裝';
        confidence = 0.95;
    }

    // === 無調味堅果 ===
    if (!productName && /無調味堅果/.test(fullTextNoSpace)) {
        if (/綜合|原味綜合/.test(fullTextNoSpace)) productName = '無調味綜合堅果';
        else if (/核桃/.test(fullTextNoSpace)) productName = '無調味核桃';
        else if (/腰果/.test(fullTextNoSpace)) productName = '無調味腰果';
        else if (/杏仁/.test(fullTextNoSpace)) productName = '無調味杏仁';
        else if (/夏威夷豆/.test(fullTextNoSpace)) productName = '無調味夏威夷豆';
        confidence = 0.9;
    }

    // === 椰棗 ===
    if (!productName && /椰棗/.test(fullTextNoSpace)) {
        if (/中東/.test(fullTextNoSpace)) productName = '★中東椰棗300g';
        else if (/夏威夷豆/.test(fullTextNoSpace)) productName = '椰棗豆子150g';
        else if (/核桃/.test(fullTextNoSpace)) productName = '椰棗核桃150g';
        else if (/腰果/.test(fullTextNoSpace)) productName = '椰棗腰果150g';
        else if (/杏仁/.test(fullTextNoSpace)) productName = '椰棗杏仁150g';
        confidence = 0.9;
    }

    // === 牛奶糖 ===
    if (!productName && /牛奶糖/.test(fullTextNoSpace)) {
        if (/50g/.test(fullTextNoSpace)) productName = '牛奶糖-50g';
        else productName = '牛奶糖';
        confidence = 0.9;
    }

    // === 土鳳梨酥 ===
    if (!productName && /土鳳梨酥|金磚土鳳梨酥/.test(fullTextNoSpace)) {
        productName = '土鳳梨酥(紅點)';
        confidence = 0.9;
    }

    // === 招牌雙塔組合 ===
    if (!productName && /招牌雙塔組合|招牌雙塔/.test(fullTextNoSpace)) {
        productName = '雙塔';
        confidence = 0.9;
    }

    // === 南棗核桃糕 ===
    if (!productName && /南棗核桃糕/.test(fullTextNoSpace)) {
        productName = '南棗核桃糕';
        confidence = 0.9;
    }

    // === 鳳凰酥 ===
    if (!productName && /鳳凰酥/.test(fullTextNoSpace)) {
        productName = '鳳凰酥';
        confidence = 0.9;
    }

    // === 規格提取 ===
    if (!column) {
        if (/300g/.test(fullTextNoSpace)) { column = 'B'; spec = '300g'; confidence = 0.9; }
        else if (/280g/.test(fullTextNoSpace)) { column = 'C'; spec = '280g'; confidence = 0.9; }
        else if (/200g/.test(fullTextNoSpace)) { column = 'C'; spec = '200g'; confidence = 0.9; }
        else if (/150g/.test(fullTextNoSpace)) { column = 'B'; spec = '150g'; confidence = 0.9; }
        else if (/135g/.test(fullTextNoSpace)) { column = 'C'; spec = '135g'; confidence = 0.9; }
        else if (/120g/.test(fullTextNoSpace)) { column = 'B'; spec = '120g'; confidence = 0.9; }
        else if (/90g/.test(fullTextNoSpace)) { column = 'B'; spec = '90g'; confidence = 0.9; }
        else if (/50g/.test(fullTextNoSpace) && !/150g/.test(fullTextNoSpace)) { column = 'B'; spec = '50g'; confidence = 0.9; }
        else if (/45g/.test(fullTextNoSpace)) { column = 'B'; spec = '45g'; confidence = 0.9; }
        // 優先從規格提取入數
        else if (/10入/.test(specNoSpace)) { column = 'B'; spec = '10入袋裝'; confidence = 0.9; }
        else if (/15入/.test(specNoSpace)) { column = 'C'; spec = '15入袋裝'; confidence = 0.9; }
        else if (/12入/.test(specNoSpace)) { column = 'C'; spec = '12入袋裝'; confidence = 0.9; }
        else if (/8入/.test(specNoSpace)) { column = 'B'; spec = '8入袋裝'; confidence = 0.9; }
        // 從商品名稱提取入數（次優先）
        else if (/10入/.test(fullTextNoSpace) && !/15入/.test(fullTextNoSpace)) { column = 'B'; spec = '10入袋裝'; confidence = 0.85; }
        else if (/15入/.test(fullTextNoSpace)) { column = 'C'; spec = '15入袋裝'; confidence = 0.85; }
        else if (/12入/.test(fullTextNoSpace)) { column = 'C'; spec = '12入袋裝'; confidence = 0.85; }
        else if (/8入/.test(fullTextNoSpace)) { column = 'B'; spec = '8入袋裝'; confidence = 0.85; }
        // 單顆/活動專用 - 瑪德蓮用 C 欄，其他用 D 欄
        else if (/單顆/.test(fullTextNoSpace) || /活動專用/.test(fullTextNoSpace) || /活動/.test(fullTextNoSpace)) {
            const isMadeleine = /瑪德蓮/.test(fullTextNoSpace);
            column = isMadeleine ? 'C' : 'D'; spec = '單顆'; confidence = 0.9;
            const activityMatch = specNoSpace.match(/(\d+)入/);
            if (activityMatch) {
                multiplier = parseInt(activityMatch[1]);
            }
        }
    }

    const mappedQuantity = column ? quantity * multiplier : 0;
    const standardizedName = productName ? normalizeProductName(productName) : null;

    return {
        templateProduct: standardizedName,
        templateColumn: column,
        templateSpec: spec,
        multiplier: multiplier,
        mappedQuantity: mappedQuantity,
        confidence: productName ? confidence : 0.3
    };
}

// 蝦皮專用映射函數
function autoMapProductShopee(pickingName, pickingSpec, quantity) {
    const fullText = `${pickingName} ${pickingSpec}`.toLowerCase();
    const fullTextNoSpaceTemp = `${pickingName} ${pickingSpec}`.replace(/\s+/g, '');

    // 過濾（禮盒不應被過濾）
    const isGiftBox = /禮盒/.test(fullTextNoSpaceTemp);
    if ((fullText.includes('咖啡') && !fullText.includes('咖啡小花') && !isGiftBox) ||
        fullText.includes('提袋加購')) {
        return {
            templateProduct: null,
            templateColumn: null,
            templateSpec: null,
            multiplier: 1,
            mappedQuantity: 0,
            confidence: 0
        };
    }

    // 1. 移除前綴
    let productName = pickingName.replace(/【艾薇手工坊】/g, '').trim();

    // 合併商品名稱和規格
    const fullTextNoSpace = `${pickingName} ${pickingSpec}`.replace(/\s+/g, '');
    const specNoSpace = pickingSpec.replace(/\s+/g, '');
    const nameNoSpace = productName.replace(/\s+/g, '');

    // 2. 優先從規格中提取口味（拆分商品已在 autoMapProduct 中處理）
    let extractedFlavor = null;

    // === 禮盒類型判斷 ===
    if (/禮盒/.test(fullTextNoSpace)) {
        if (/雙塔禮盒/.test(specNoSpace) || /雙塔禮盒/.test(fullTextNoSpace)) {
            extractedFlavor = '雙塔禮盒';
        } else if (/招牌雙塔/.test(fullTextNoSpace) && /禮盒/.test(fullTextNoSpace)) {
            extractedFlavor = '雙塔禮盒';
        } else if (/夏威夷豆塔禮盒/.test(fullTextNoSpace) && /蔓越莓/.test(fullTextNoSpace)) {
            extractedFlavor = '蔓越莓禮盒';
        } else if (/夏威夷豆塔禮盒/.test(fullTextNoSpace) && /綜合/.test(fullTextNoSpace)) {
            extractedFlavor = '綜豆禮盒';
        } else if (/堅果塔禮盒/.test(fullTextNoSpace) && /綜合/.test(fullTextNoSpace)) {
            extractedFlavor = '綜堅禮盒';
        } else if (/戀戀雪花禮盒|新戀戀雪花/.test(fullTextNoSpace)) {
            extractedFlavor = '戀戀雪花禮盒';
        } else if (/浪漫詩篇禮盒|新浪漫詩篇/.test(fullTextNoSpace)) {
            extractedFlavor = '浪漫詩篇禮盒';
        } else if (/暖暖幸福禮盒|新暖暖幸福/.test(fullTextNoSpace)) {
            extractedFlavor = '暖暖幸福禮盒';
        } else if (/臻愛時光禮盒|新臻愛時光/.test(fullTextNoSpace)) {
            extractedFlavor = '臻愛時光禮盒';
        } else if (/濃情滿載禮盒|新濃情滿載/.test(fullTextNoSpace)) {
            extractedFlavor = '濃情滿載禮盒';
        } else if (/午後漫步禮盒|新午後漫步/.test(fullTextNoSpace)) {
            extractedFlavor = '午後漫步禮盒';
        } else if (/那年花開禮盒|新那年花開/.test(fullTextNoSpace)) {
            extractedFlavor = '那年花開禮盒';
        } else if (/花間逸韻禮盒|新花間逸韻/.test(fullTextNoSpace)) {
            extractedFlavor = '花間逸韻禮盒';
        } else if (/晴空塔餅禮盒|新晴空塔餅/.test(fullTextNoSpace)) {
            extractedFlavor = '晴空塔餅禮盒';
        } else if (/金緻典藏禮盒/.test(fullTextNoSpace)) {
            extractedFlavor = '輕-金緻典藏禮盒';
        } else if (/香榭漫遊禮盒/.test(fullTextNoSpace)) {
            extractedFlavor = '輕-香榭漫遊禮盒';
        } else if (/晨曦物語禮盒/.test(fullTextNoSpace)) {
            extractedFlavor = '輕-晨曦物語禮盒';
        } else if (/月光序曲禮盒/.test(fullTextNoSpace)) {
            extractedFlavor = '輕-月光序曲禮盒';
        } else if (/蔓越莓禮盒/.test(fullTextNoSpace)) {
            extractedFlavor = '蔓越莓禮盒';
        } else if (/綜豆禮盒/.test(fullTextNoSpace)) {
            extractedFlavor = '綜豆禮盒';
        } else if (/綜堅禮盒/.test(fullTextNoSpace)) {
            extractedFlavor = '綜堅禮盒';
        }
    }

    // === 瑪德蓮口味 ===
    if (!extractedFlavor && /瑪德蓮/.test(specNoSpace)) {
        if (/抹茶/.test(specNoSpace)) extractedFlavor = '瑪德蓮-抹茶';
        else if (/檸檬/.test(specNoSpace)) extractedFlavor = '瑪德蓮-檸檬';
        else if (/紅茶/.test(specNoSpace)) extractedFlavor = '瑪德蓮-紅茶';
        else if (/蜂蜜/.test(specNoSpace)) extractedFlavor = '瑪德蓮-蜂蜜';
        else if (/巧克力/.test(specNoSpace)) extractedFlavor = '瑪德蓮-巧克力';
        else if (/柑橘/.test(specNoSpace)) extractedFlavor = '瑪德蓮-柑橘';
        else if (/綜合|隨機/.test(specNoSpace)) extractedFlavor = '瑪德蓮-綜合';
    }

    // === 杏仁瓦片口味 ===
    // 檢查商品名稱或規格中是否包含「杏仁瓦片」
    if (!extractedFlavor && (/杏仁瓦片/.test(specNoSpace) || /杏仁瓦片/.test(nameNoSpace))) {
        // 優先判斷 45g 原味，直接映射到完整名稱
        if (/原味/.test(fullTextNoSpace) && /45g/.test(fullTextNoSpace)) {
            extractedFlavor = '瓦片-原味45克';
        } else if (/原味/.test(specNoSpace) || /原味/.test(nameNoSpace)) extractedFlavor = '瓦片-原味';
        else if (/抹茶/.test(specNoSpace) || /抹茶/.test(nameNoSpace)) extractedFlavor = '瓦片-抹茶';
        else if (/紅茶/.test(specNoSpace) || /紅茶/.test(nameNoSpace)) extractedFlavor = '瓦片-紅茶';
        else if (/巧克力/.test(specNoSpace) || /巧克力/.test(nameNoSpace)) extractedFlavor = '瓦片-巧克力';
        else if (/海苔/.test(specNoSpace) || /海苔/.test(nameNoSpace)) extractedFlavor = '瓦片-海苔';
        else if (/黑糖/.test(specNoSpace) || /黑糖/.test(nameNoSpace)) extractedFlavor = '瓦片-黑糖';
        else if (/青花椒/.test(specNoSpace) || /青花椒/.test(nameNoSpace)) extractedFlavor = '瓦片-青花椒';
        else if (/綜合/.test(specNoSpace) || /綜合/.test(nameNoSpace)) extractedFlavor = '瓦片-綜合';
    }

    // === 50g牛奶糖 ===
    if (!extractedFlavor && fullTextNoSpace.includes('牛奶糖') && fullTextNoSpace.includes('50g')) {
        extractedFlavor = '牛奶糖-50g';
    } else if (!extractedFlavor && fullTextNoSpace.includes('夏威夷豆法式牛奶糖') && !fullTextNoSpace.includes('50g')) {
        extractedFlavor = '牛奶糖';
    }

    // === 無調味堅果 ===
    if (!extractedFlavor && /無調味堅果/.test(fullTextNoSpace)) {
        if (/原味核桃|核桃/.test(fullTextNoSpace) && !fullTextNoSpace.includes('綜合')) {
            extractedFlavor = '無調味核桃';
        } else if (/原味腰果|腰果/.test(fullTextNoSpace) && !fullTextNoSpace.includes('綜合')) {
            extractedFlavor = '無調味腰果';
        } else if (/原味杏仁|杏仁/.test(fullTextNoSpace) && !fullTextNoSpace.includes('綜合')) {
            extractedFlavor = '無調味杏仁';
        } else if (/夏威夷豆/.test(fullTextNoSpace)) {
            extractedFlavor = '無調味夏威夷豆';
        } else if (/綜合/.test(fullTextNoSpace) || /原味綜合/.test(fullTextNoSpace)) {
            extractedFlavor = '無調味綜合堅果';
        }
    }

    // === 土鳳梨酥 ===
    if (!extractedFlavor && /金磚土鳳梨酥/.test(fullTextNoSpace)) {
        extractedFlavor = '土鳳梨酥(紅點)';
    }

    // === 椰棗系列 ===
    if (!extractedFlavor && /椰棗/.test(fullTextNoSpace)) {
        if (/椰棗夏威夷豆/.test(fullTextNoSpace)) extractedFlavor = '椰棗豆子150g';
        else if (/椰棗核桃/.test(fullTextNoSpace)) extractedFlavor = '椰棗核桃150g';
        else if (/椰棗腰果/.test(fullTextNoSpace)) extractedFlavor = '椰棗腰果150g';
        else if (/椰棗杏仁/.test(fullTextNoSpace)) extractedFlavor = '椰棗杏仁150g';
        else if (/中東.*椰棗|中東椰棗/.test(fullTextNoSpace)) extractedFlavor = '★中東椰棗300g';
    }

    // === 杏仁瓦片-原味45g ===
    if (!extractedFlavor && /杏仁瓦片/.test(fullTextNoSpace) && /原味/.test(fullTextNoSpace) && /45g/.test(fullTextNoSpace)) {
        extractedFlavor = '瓦片-原味45克';
    }

    // === 西點餅乾 ===
    if (!extractedFlavor && (/西點餅乾/.test(specNoSpace) || /西點餅乾/.test(nameNoSpace))) {
        if (/綜合西點|綜合/.test(specNoSpace) || /綜合西點/.test(nameNoSpace)) extractedFlavor = '西點-綜合';
        else if (/乳酪酥條/.test(specNoSpace)) extractedFlavor = '西點-乳酪酥條';
        else if (/蔓越莓貝殼/.test(specNoSpace)) extractedFlavor = '西點-蔓越莓貝殼';
        else if (/藍莓小花/.test(specNoSpace)) extractedFlavor = '西點-藍莓小花';
        else if (/咖啡小花/.test(specNoSpace)) extractedFlavor = '西點-咖啡小花';
        else if (/巧克力貝殼/.test(specNoSpace)) extractedFlavor = '西點-巧克力貝殼';
    }

    // === 夏威夷豆塔 ===
    if (!extractedFlavor && /夏威夷豆塔/.test(nameNoSpace)) {
        if (/蜂蜜蔓越莓|蔓越莓/.test(nameNoSpace)) extractedFlavor = '豆塔-蔓越莓';
        else if (/焦糖/.test(nameNoSpace)) extractedFlavor = '豆塔-焦糖';
        else if (/巧克力/.test(nameNoSpace)) extractedFlavor = '豆塔-巧克力';
        else if (/抹茶/.test(nameNoSpace)) extractedFlavor = '豆塔-抹茶';
        else if (/椒麻/.test(nameNoSpace)) extractedFlavor = '豆塔-椒麻';
        else if (/綜合/.test(nameNoSpace)) extractedFlavor = '豆塔-綜合';
    }

    // === 奶油餅乾 ===
    if (!extractedFlavor && /奶油曲奇餅乾|奶油餅乾/.test(nameNoSpace)) {
        if (/抹茶|小山園抹茶/.test(fullTextNoSpace)) extractedFlavor = '奶油-抹茶';
        else if (/焦糖牛奶/.test(fullTextNoSpace)) extractedFlavor = '奶油-焦糖牛奶';
        else if (/法國巧克力|巧克力/.test(fullTextNoSpace)) extractedFlavor = '奶油-法國巧克力';
        else if (/蜂蜜檸檬/.test(fullTextNoSpace)) extractedFlavor = '奶油-蜂蜜檸檬';
        else if (/伯爵紅茶/.test(fullTextNoSpace)) extractedFlavor = '奶油-伯爵紅茶';
    }

    // === 千層小酥條 ===
    if (!extractedFlavor && /千層.*小酥條|小酥條/.test(fullTextNoSpace)) {
        extractedFlavor = '千層-小酥條';
    }

    // === 從規格提取堅果塔口味 ===
    if (!extractedFlavor) {
        const nutTartMatch = specNoSpace.match(/堅果塔-?(.+?)口味/);
        if (nutTartMatch) extractedFlavor = `堅果塔-${nutTartMatch[1]}`;
    }

    // === 從規格提取夏威夷豆塔口味 ===
    if (!extractedFlavor) {
        const beanTartMatch = specNoSpace.match(/夏威夷豆塔-?(.+?)(\d+入|口味)/);
        if (beanTartMatch) {
            let flavor = beanTartMatch[1].replace(/蜂蜜蔓越莓/, '蔓越莓');
            extractedFlavor = `豆塔-${flavor}`;
        }
    }

    // === 從規格提取奶油曲奇口味 ===
    if (!extractedFlavor) {
        const butterMatch = specNoSpace.match(/奶油曲奇-?(.+?)口味/);
        if (butterMatch) extractedFlavor = `奶油-${butterMatch[1]}`;
    }

    // === 從規格提取瓦片口味 ===
    if (!extractedFlavor) {
        const waferMatch = specNoSpace.match(/杏仁瓦片-?(.+?)口味/) || specNoSpace.match(/^(.+?)口味\d+g/);
        if (waferMatch) extractedFlavor = `瓦片-${waferMatch[1]}`;
    }

    // === 從規格提取雪花餅口味 ===
    if (!extractedFlavor) {
        const snowMatch = specNoSpace.match(/雪花餅-(.+?)口味/);
        if (snowMatch) extractedFlavor = `雪花餅-${snowMatch[1]}`;
    }

    // === 南棗核桃糕 ===
    if (!extractedFlavor && /南棗核桃糕/.test(fullTextNoSpace)) {
        extractedFlavor = '南棗核桃糕';
    }

    // === 招牌雙塔組合 ===
    if (!extractedFlavor && /招牌雙塔組合/.test(fullTextNoSpace)) {
        extractedFlavor = '雙塔';
    }

    // 使用提取到的口味
    if (extractedFlavor) {
        productName = extractedFlavor;
    } else {
        // 簡化產品名稱
        const simplifyRules = [
            { pattern: /夏威夷豆塔[|-](.+?)口味/, replacement: '豆塔-$1' },
            { pattern: /夏威夷豆塔[|-](.+?)\s/, replacement: '豆塔-$1' },
            { pattern: /夏威夷豆塔[|-](.+?)$/, replacement: '豆塔-$1' },
            { pattern: /堅果塔[|-](.+?)口味/, replacement: '堅果塔-$1' },
            { pattern: /經典杏仁瓦片[|-](.+?)口味/, replacement: '瓦片-$1' },
            { pattern: /杏仁瓦片[|-](.+?)口味/, replacement: '瓦片-$1' },
            { pattern: /奶油曲奇餅乾[|-](.+?)口味/, replacement: '奶油-$1' },
            { pattern: /奶油餅乾[|-](.+?)口味/, replacement: '奶油-$1' },
            { pattern: /雪花餅[|-](.+?)口味/, replacement: '雪花餅-$1' },
            { pattern: /胖貝殼瑪德蓮[|-](.+?)口味/, replacement: '瑪德蓮-$1' },
            { pattern: /招牌雙塔組合/, replacement: '雙塔' },
        ];

        for (let rule of simplifyRules) {
            if (rule.pattern.test(productName)) {
                productName = productName.replace(rule.pattern, rule.replacement).trim();
                break;
            }
        }
    }

    // 清理
    if (!extractedFlavor) {
        productName = productName
            .replace(/\s+/g, ' ')
            .replace(/\d+g\/\d+g.*$/i, '')
            .replace(/\d+g.*$/i, '')
            .replace(/\d+入.*$/i, '')
            .replace(/\(袋裝\)$/i, '')
            .replace(/\(盒裝.*?\)$/i, '')
            .replace(/\/袋裝.*$/i, '')
            .trim();
    }

    // 4. 提取倍數
    let multiplier = 1;
    const multiplierMatch = pickingSpec.match(/x\s*(\d+)\s*包/);
    if (multiplierMatch) {
        const num = parseInt(multiplierMatch[1]);
        if (!isNaN(num) && num > 0) multiplier = num;
    }

    // 5. 提取規格並判斷欄位
    let column = null;
    let spec = null;
    let confidence = 0.3;

    // 禮盒
    if (/禮盒/.test(fullText)) {
        column = 'B'; spec = '禮盒'; confidence = 0.95; multiplier = 1;
    }
    // 瓦片-原味45克 是獨立產品，不需要額外設置欄位規格
    else if (productName === '瓦片-原味45克') {
        column = 'B'; spec = null; confidence = 0.95; multiplier = 1;
    }
    // 千層小酥條 - 直接映射到 B 欄/小包裝
    else if (productName === '千層-小酥條') {
        column = 'B'; spec = '小包裝'; confidence = 0.95; multiplier = 1;
    }
    // 從規格提取重量（優先）
    else if (/45g/.test(pickingSpec)) { column = 'B'; spec = '45g'; confidence = 0.95; multiplier = 1; }
    else if (/50g/.test(pickingSpec)) { column = 'B'; spec = '50g'; confidence = 0.95; multiplier = 1; }
    else if (/90g/.test(pickingSpec)) { column = 'B'; spec = '90g'; confidence = 0.95; multiplier = 1; }
    else if (/120g/.test(pickingSpec)) { column = 'B'; spec = '120g'; confidence = 0.95; multiplier = 1; }
    else if (/135g/.test(pickingSpec)) { column = 'C'; spec = '135g'; confidence = 0.95; multiplier = 1; }
    else if (/150g/.test(pickingSpec)) { column = 'B'; spec = '150g'; confidence = 0.95; multiplier = 1; }
    else if (/200g/.test(pickingSpec)) { column = 'C'; spec = '200g'; confidence = 0.95; multiplier = 1; }
    else if (/280g/.test(pickingSpec)) { column = 'C'; spec = '280g'; confidence = 0.95; multiplier = 1; }
    else if (/300g/.test(pickingSpec)) { column = 'B'; spec = '300g'; confidence = 0.95; multiplier = 1; }
    // 從商品名稱提取重量（次優先）
    else if (/300g/.test(pickingName)) { column = 'B'; spec = '300g'; confidence = 0.9; multiplier = 1; }
    else if (/280g/.test(pickingName)) { column = 'C'; spec = '280g'; confidence = 0.9; multiplier = 1; }
    else if (/200g/.test(pickingName)) { column = 'C'; spec = '200g'; confidence = 0.9; multiplier = 1; }
    else if (/150g/.test(pickingName)) { column = 'B'; spec = '150g'; confidence = 0.9; multiplier = 1; }
    else if (/135g/.test(pickingName)) { column = 'C'; spec = '135g'; confidence = 0.9; multiplier = 1; }
    else if (/120g/.test(pickingName)) { column = 'B'; spec = '120g'; confidence = 0.9; multiplier = 1; }
    else if (/90g/.test(pickingName) && !/\//.test(pickingName)) { column = 'B'; spec = '90g'; confidence = 0.9; multiplier = 1; }
    // 入數規格
    else if (/10入/.test(specNoSpace)) { column = 'B'; spec = '10入袋裝'; confidence = 0.9; }
    else if (/15入/.test(specNoSpace)) { column = 'C'; spec = '15入袋裝'; confidence = 0.9; }
    else if (/12入/.test(specNoSpace)) { column = 'C'; spec = '12入袋裝'; confidence = 0.9; }
    else if (/8入/.test(specNoSpace)) { column = 'B'; spec = '8入袋裝'; confidence = 0.9; }
    // 從商品名稱提取入數
    else if (/10入/.test(nameNoSpace) && !/15入/.test(specNoSpace)) { column = 'B'; spec = '10入袋裝'; confidence = 0.85; }
    else if (/15入/.test(nameNoSpace) && !/10入/.test(specNoSpace)) { column = 'C'; spec = '15入袋裝'; confidence = 0.85; }
    else if (/12入/.test(nameNoSpace)) { column = 'C'; spec = '12入袋裝'; confidence = 0.85; }
    else if (/8入/.test(nameNoSpace)) { column = 'B'; spec = '8入袋裝'; confidence = 0.85; }
    // 活動專用 - 瑪德蓮用 C 欄，其他用 D 欄
    else if (/單顆/.test(fullText) || /單個/.test(fullText)) {
        const isMadeleine = /瑪德蓮/.test(fullTextNoSpaceTemp);
        column = isMadeleine ? 'C' : 'D'; spec = '單顆'; confidence = 0.85;
        const singleMatch = specNoSpace.match(/(\d+)入/);
        if (singleMatch) multiplier = parseInt(singleMatch[1]);
    } else if (/\d+入/.test(specNoSpace) && /活動/.test(fullTextNoSpaceTemp)) {
        const isMadeleine = /瑪德蓮/.test(fullTextNoSpaceTemp);
        column = isMadeleine ? 'C' : 'D'; spec = '單顆'; confidence = 0.8;
        const activityMatch = specNoSpace.match(/(\d+)入/);
        if (activityMatch) multiplier = parseInt(activityMatch[1]);
    }

    // 6. 計算對應數量
    const mappedQuantity = column ? quantity * multiplier : 0;

    // 7. 標準化商品名稱
    const standardizedName = normalizeProductName(productName);

    return {
        templateProduct: standardizedName,
        templateColumn: column,
        templateSpec: spec,
        multiplier: multiplier,
        mappedQuantity: mappedQuantity,
        confidence: confidence
    };
}

// 橘點子專用映射函數
function autoMapProductOrangePoint(productName, quantity) {
    const nameNoSpace = productName.replace(/\s+/g, '');

    let mappedName = null;
    let column = null;
    let spec = null;
    let confidence = 0.3;

    // 瓦片
    if (/杏仁瓦片/.test(nameNoSpace)) {
        const flavorMatch = productName.match(/[\[\(](.+?)[\]\)]/);
        let flavor = flavorMatch ? flavorMatch[1].trim() : '';

        if (flavor === '原味' || /原味/.test(nameNoSpace)) mappedName = '瓦片-原味';
        else if (flavor === '巧克力' || /巧克力/.test(nameNoSpace)) mappedName = '瓦片-巧克力';
        else if (flavor === '紅茶' || /紅茶/.test(nameNoSpace)) mappedName = '瓦片-紅茶';
        else if (flavor === '黑糖' || /黑糖/.test(nameNoSpace)) mappedName = '瓦片-黑糖';
        else if (flavor === '抹茶' || /抹茶/.test(nameNoSpace)) mappedName = '瓦片-抹茶';
        else if (flavor === '海苔' || /海苔/.test(nameNoSpace)) mappedName = '瓦片-海苔';
        else if (flavor === '青花椒' || /青花椒/.test(nameNoSpace)) mappedName = '瓦片-青花椒';
        else if (/綜合/.test(flavor) || /綜合/.test(nameNoSpace)) mappedName = '瓦片-綜合';
        confidence = 0.95;

        if (/135g/.test(nameNoSpace)) { column = 'C'; spec = '135g'; }
        else if (/90g/.test(nameNoSpace)) { column = 'B'; spec = '90g'; }
        else if (/45g/.test(nameNoSpace)) { column = 'B'; spec = '45g'; }
    }

    // 豆塔
    if (!mappedName && /夏威夷豆塔/.test(nameNoSpace)) {
        const flavorMatch = productName.match(/\[(.+?)\]/);
        if (flavorMatch) {
            const flavor = flavorMatch[1];
            if (flavor.includes('蔓越莓')) mappedName = '豆塔-蔓越莓';
            else if (flavor.includes('焦糖')) mappedName = '豆塔-焦糖';
            else if (flavor.includes('巧克力')) mappedName = '豆塔-巧克力';
            else if (flavor.includes('抹茶')) mappedName = '豆塔-抹茶';
            else if (flavor.includes('椒麻')) mappedName = '豆塔-椒麻';
            else if (flavor.includes('綜合')) mappedName = '豆塔-綜合';
        }
        confidence = 0.95;
        if (/15入/.test(nameNoSpace)) { column = 'C'; spec = '15入袋裝'; }
        else if (/10入/.test(nameNoSpace)) { column = 'B'; spec = '10入袋裝'; }
    }

    // 堅果塔
    if (!mappedName && /堅果塔/.test(nameNoSpace)) {
        const flavorMatch = productName.match(/\[(.+?)\]/);
        if (flavorMatch) {
            const flavor = flavorMatch[1];
            if (flavor.includes('蜂蜜')) mappedName = '堅果塔-蜂蜜';
            else if (flavor.includes('焦糖')) mappedName = '堅果塔-焦糖';
            else if (flavor.includes('巧克力')) mappedName = '堅果塔-巧克力';
            else if (flavor.includes('海苔')) mappedName = '堅果塔-海苔';
            else if (flavor.includes('咖哩')) mappedName = '堅果塔-咖哩';
            else if (flavor.includes('綜合')) mappedName = '堅果塔-綜合';
        }
        confidence = 0.95;
        if (/15入/.test(nameNoSpace)) { column = 'C'; spec = '15入袋裝'; }
        else if (/10入/.test(nameNoSpace)) { column = 'B'; spec = '10入袋裝'; }
    }

    // 雪花餅
    if (!mappedName && /雪花餅/.test(nameNoSpace)) {
        if (/蔓越莓/.test(nameNoSpace)) mappedName = '雪花餅-蔓越莓';
        else if (/巧克力/.test(nameNoSpace)) mappedName = '雪花餅-巧克力';
        else if (/金沙/.test(nameNoSpace)) mappedName = '雪花餅-金沙';
        else if (/抹茶/.test(nameNoSpace)) mappedName = '雪花餅-抹茶';
        else if (/肉鬆/.test(nameNoSpace)) mappedName = '雪花餅-肉鬆';
        else if (/綜合/.test(nameNoSpace)) mappedName = '雪花餅-綜合';
        confidence = 0.95;
        if (/15入/.test(nameNoSpace)) { column = 'C'; spec = '15入袋裝'; }
        else if (/10入/.test(nameNoSpace)) { column = 'B'; spec = '10入袋裝'; }
    }

    // 禮盒
    if (!mappedName && /禮盒/.test(nameNoSpace)) {
        column = 'B'; spec = '禮盒'; confidence = 0.95;
        if (/雙塔禮盒|招牌雙塔/.test(nameNoSpace)) mappedName = '雙塔禮盒';
        else if (/戀戀雪花/.test(nameNoSpace)) mappedName = '戀戀雪花禮盒';
        else if (/浪漫詩篇/.test(nameNoSpace)) mappedName = '浪漫詩篇禮盒';
        else if (/暖暖幸福/.test(nameNoSpace)) mappedName = '暖暖幸福禮盒';
    }

    const standardizedName = mappedName ? normalizeProductName(mappedName) : null;
    const mappedQuantity = column ? quantity : 0;

    return {
        templateProduct: standardizedName,
        templateColumn: column,
        templateSpec: spec,
        multiplier: 1,
        mappedQuantity: mappedQuantity,
        confidence: mappedName ? confidence : 0.3
    };
}
