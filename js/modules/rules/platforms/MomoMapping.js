/**
 * @file MomoMapping.js
 * @description MOMO 與 官網 平台商品映射邏輯
 * @author Ivy House TW Development Team
 */

import { normalizeProductName } from '../StandardProducts.js';

/**
 * MOMO/官網專用映射函數
 * @param {string} pickingName 撿貨單品名
 * @param {string} pickingSpec 規格
 * @param {number} quantity 數量
 * @returns {object} 映射結果
 */
export function autoMapProductMomo(pickingName, pickingSpec, quantity) {
    const fullText = `${pickingName} ${pickingSpec}`;
    const fullTextNoSpace = fullText.replace(/\s+/g, '');
    const nameNoSpace = pickingName.replace(/\s+/g, '');
    const specNoSpace = (pickingSpec || '').replace(/\s+/g, '');

    // 過濾
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

    // 提取倍數
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
        column = 'B'; spec = '小包裝'; confidence = 0.95;
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
        else if (/10入/.test(specNoSpace)) { column = 'B'; spec = '10入袋裝'; confidence = 0.9; }
        else if (/15入/.test(specNoSpace)) { column = 'C'; spec = '15入袋裝'; confidence = 0.9; }
        else if (/12入/.test(specNoSpace)) { column = 'C'; spec = '12入袋裝'; confidence = 0.9; }
        else if (/8入/.test(specNoSpace)) { column = 'B'; spec = '8入袋裝'; confidence = 0.9; }
        else if (/10入/.test(fullTextNoSpace) && !/15入/.test(fullTextNoSpace)) { column = 'B'; spec = '10入袋裝'; confidence = 0.85; }
        else if (/15入/.test(fullTextNoSpace)) { column = 'C'; spec = '15入袋裝'; confidence = 0.85; }
        else if (/12入/.test(fullTextNoSpace)) { column = 'C'; spec = '12入袋裝'; confidence = 0.85; }
        else if (/8入/.test(fullTextNoSpace)) { column = 'B'; spec = '8入袋裝'; confidence = 0.85; }
        else if (/單顆/.test(fullTextNoSpace) || /活動專用/.test(fullTextNoSpace) || /活動/.test(fullTextNoSpace)) {
            const isMadeleine = /瑪德蓮/.test(fullTextNoSpace);
            column = isMadeleine ? 'C' : 'D'; spec = '單顆'; confidence = 0.9;
            const activityMatch = specNoSpace.match(/(\d+)入/);
            if (activityMatch) multiplier = parseInt(activityMatch[1]);
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
