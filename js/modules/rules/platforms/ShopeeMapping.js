/**
 * @file ShopeeMapping.js
 * @description 蝦皮平台商品映射邏輯
 * @author Ivy House TW Development Team
 */

import { normalizeProductName } from '../StandardProducts.js';

/**
 * 蝦皮專用映射函數
 * @param {string} pickingName 撿貨單品名
 * @param {string} pickingSpec 規格
 * @param {number} quantity 數量
 * @returns {object} 映射結果
 */
export function autoMapProductShopee(pickingName, pickingSpec, quantity) {
    const fullText = `${pickingName} ${pickingSpec}`.toLowerCase();
    const fullTextNoSpaceTemp = `${pickingName} ${pickingSpec}`.replace(/\s+/g, '');

    // 過濾
    const isGiftBox = /禮盒/.test(fullTextNoSpaceTemp);
    if ((fullText.includes('咖啡') && !fullText.includes('咖啡小花') && !isGiftBox) ||
        fullText.includes('提袋加購')) {
        return { templateProduct: null, templateColumn: null, templateSpec: null, multiplier: 1, mappedQuantity: 0, confidence: 0 };
    }

    // 1. 移除前綴
    let productName = pickingName.replace(/【艾薇手工坊】/g, '').trim();
    const specNoSpace = pickingSpec.replace(/\s+/g, '');
    const nameNoSpace = productName.replace(/\s+/g, '');
    const fullTextNoSpace = fullTextNoSpaceTemp;

    // 2. 優先從規格中提取口味
    let extractedFlavor = null;

    // === 禮盒類型判斷 ===
    if (/禮盒/.test(fullTextNoSpace)) {
        if (/雙塔禮盒/.test(specNoSpace) || /雙塔禮盒/.test(fullTextNoSpace)) extractedFlavor = '雙塔禮盒';
        else if (/招牌雙塔/.test(fullTextNoSpace) && /禮盒/.test(fullTextNoSpace)) extractedFlavor = '雙塔禮盒';
        else if (/夏威夷豆塔禮盒/.test(fullTextNoSpace) && /蔓越莓/.test(fullTextNoSpace)) extractedFlavor = '蔓越莓禮盒';
        else if (/夏威夷豆塔禮盒/.test(fullTextNoSpace) && /綜合/.test(fullTextNoSpace)) extractedFlavor = '綜豆禮盒';
        else if (/堅果塔禮盒/.test(fullTextNoSpace) && /綜合/.test(fullTextNoSpace)) extractedFlavor = '綜堅禮盒';
        else if (/戀戀雪花禮盒|新戀戀雪花/.test(fullTextNoSpace)) extractedFlavor = '戀戀雪花禮盒';
        else if (/浪漫詩篇禮盒|新浪漫詩篇/.test(fullTextNoSpace)) extractedFlavor = '浪漫詩篇禮盒';
        else if (/暖暖幸福禮盒|新暖暖幸福/.test(fullTextNoSpace)) extractedFlavor = '暖暖幸福禮盒';
        else if (/臻愛時光禮盒|新臻愛時光/.test(fullTextNoSpace)) extractedFlavor = '臻愛時光禮盒';
        else if (/濃情滿載禮盒|新濃情滿載/.test(fullTextNoSpace)) extractedFlavor = '濃情滿載禮盒';
        else if (/午後漫步禮盒|新午後漫步/.test(fullTextNoSpace)) extractedFlavor = '午後漫步禮盒';
        else if (/那年花開禮盒|新那年花開/.test(fullTextNoSpace)) extractedFlavor = '那年花開禮盒';
        else if (/花間逸韻禮盒|新花間逸韻/.test(fullTextNoSpace)) extractedFlavor = '花間逸韻禮盒';
        else if (/晴空塔餅禮盒|新晴空塔餅/.test(fullTextNoSpace)) extractedFlavor = '晴空塔餅禮盒';
        else if (/金緻典藏禮盒/.test(fullTextNoSpace)) extractedFlavor = '輕-金緻典藏禮盒';
        else if (/香榭漫遊禮盒/.test(fullTextNoSpace)) extractedFlavor = '輕-香榭漫遊禮盒';
        else if (/晨曦物語禮盒/.test(fullTextNoSpace)) extractedFlavor = '輕-晨曦物語禮盒';
        else if (/月光序曲禮盒/.test(fullTextNoSpace)) extractedFlavor = '輕-月光序曲禮盒';
        else if (/蔓越莓禮盒/.test(fullTextNoSpace)) extractedFlavor = '蔓越莓禮盒';
        else if (/綜豆禮盒/.test(fullTextNoSpace)) extractedFlavor = '綜豆禮盒';
        else if (/綜堅禮盒/.test(fullTextNoSpace)) extractedFlavor = '綜堅禮盒';
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
    if (!extractedFlavor && (/杏仁瓦片/.test(specNoSpace) || /杏仁瓦片/.test(nameNoSpace))) {
        if (/原味/.test(fullTextNoSpace) && /45g/.test(fullTextNoSpace)) extractedFlavor = '瓦片-原味45克';
        else if (/原味/.test(specNoSpace) || /原味/.test(nameNoSpace)) extractedFlavor = '瓦片-原味';
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
        if (/原味核桃|核桃/.test(fullTextNoSpace) && !fullTextNoSpace.includes('綜合')) extractedFlavor = '無調味核桃';
        else if (/原味腰果|腰果/.test(fullTextNoSpace) && !fullTextNoSpace.includes('綜合')) extractedFlavor = '無調味腰果';
        else if (/原味杏仁|杏仁/.test(fullTextNoSpace) && !fullTextNoSpace.includes('綜合')) extractedFlavor = '無調味杏仁';
        else if (/夏威夷豆/.test(fullTextNoSpace)) extractedFlavor = '無調味夏威夷豆';
        else if (/綜合/.test(fullTextNoSpace) || /原味綜合/.test(fullTextNoSpace)) extractedFlavor = '無調味綜合堅果';
    }

    // === 土鳳梨酥 ===
    if (!extractedFlavor && /金磚土鳳梨酥/.test(fullTextNoSpace)) extractedFlavor = '土鳳梨酥(紅點)';

    // === 椰棗系列 ===
    if (!extractedFlavor && /椰棗/.test(fullTextNoSpace)) {
        if (/椰棗夏威夷豆/.test(fullTextNoSpace)) extractedFlavor = '椰棗豆子150g';
        else if (/椰棗核桃/.test(fullTextNoSpace)) extractedFlavor = '椰棗核桃150g';
        else if (/椰棗腰果/.test(fullTextNoSpace)) extractedFlavor = '椰棗腰果150g';
        else if (/椰棗杏仁/.test(fullTextNoSpace)) extractedFlavor = '椰棗杏仁150g';
        else if (/中東.*椰棗|中東椰棗/.test(fullTextNoSpace)) extractedFlavor = '★中東椰棗300g';
    }

    // === 西點餅乾 ===
    if (!extractedFlavor && (/西點餅乾/.test(specNoSpace) || /西點餅乾/.test(nameNoSpace) || /藍莓小花/.test(fullTextNoSpace) || /咖啡小花/.test(fullTextNoSpace) || /蔓越莓貝殼/.test(fullTextNoSpace) || /巧克力貝殼/.test(fullTextNoSpace) || /乳酪酥條/.test(fullTextNoSpace))) {
        if (/綜合/.test(specNoSpace) || /綜合西點/.test(nameNoSpace)) extractedFlavor = '西點-綜合';
        else if (/乳酪酥條/.test(fullTextNoSpace)) extractedFlavor = '西點-乳酪酥條';
        else if (/蔓越莓貝殼/.test(fullTextNoSpace)) extractedFlavor = '西點-蔓越莓貝殼';
        else if (/藍莓小花/.test(fullTextNoSpace)) extractedFlavor = '西點-藍莓小花';
        else if (/咖啡小花/.test(fullTextNoSpace)) extractedFlavor = '西點-咖啡小花';
        else if (/巧克力貝殼/.test(fullTextNoSpace)) extractedFlavor = '西點-巧克力貝殼';
    }

    // === 堅果塔 ===
    if (!extractedFlavor && /堅果塔/.test(fullTextNoSpace)) {
        if (/蜂蜜/.test(fullTextNoSpace)) extractedFlavor = '堅果塔-蜂蜜';
        else if (/法式焦糖|焦糖/.test(fullTextNoSpace)) extractedFlavor = '堅果塔-焦糖';
        else if (/巧克力/.test(fullTextNoSpace)) extractedFlavor = '堅果塔-巧克力';
        else if (/海苔/.test(fullTextNoSpace)) extractedFlavor = '堅果塔-海苔';
        else if (/咖哩/.test(fullTextNoSpace)) extractedFlavor = '堅果塔-咖哩';
        else if (/綜合/.test(fullTextNoSpace)) extractedFlavor = '堅果塔-綜合';
    }

    // === 夏威夷豆塔 ===
    if (!extractedFlavor && /夏威夷豆塔/.test(fullTextNoSpace)) {
        if (/蜂蜜蔓越莓|蔓越莓/.test(fullTextNoSpace)) extractedFlavor = '豆塔-蔓越莓';
        else if (/焦糖/.test(fullTextNoSpace)) extractedFlavor = '豆塔-焦糖';
        else if (/巧克力/.test(fullTextNoSpace)) extractedFlavor = '豆塔-巧克力';
        else if (/抹茶/.test(fullTextNoSpace)) extractedFlavor = '豆塔-抹茶';
        else if (/椒麻/.test(fullTextNoSpace)) extractedFlavor = '豆塔-椒麻';
        else if (/綜合/.test(fullTextNoSpace)) extractedFlavor = '豆塔-綜合';
    }

    // === 奶油餅乾 ===
    if (!extractedFlavor && (/奶油曲奇餅乾|奶油餅乾|奶油曲奇|經典熱銷餅乾/.test(fullTextNoSpace))) {
        if (/抹茶|小山園抹茶/.test(fullTextNoSpace)) extractedFlavor = '奶油-抹茶';
        else if (/焦糖牛奶/.test(fullTextNoSpace)) extractedFlavor = '奶油-焦糖牛奶';
        else if (/法國巧克力|巧克力/.test(fullTextNoSpace)) extractedFlavor = '奶油-法國巧克力';
        else if (/蜂蜜檸檬/.test(fullTextNoSpace)) extractedFlavor = '奶油-蜂蜜檸檬';
        else if (/伯爵紅茶/.test(fullTextNoSpace)) extractedFlavor = '奶油-伯爵紅茶';
    }

    // === 千層小酥條 ===
    if (!extractedFlavor && /千層.*小酥條|小酥條/.test(fullTextNoSpace)) extractedFlavor = '千層-小酥條';

    // === 雪花餅 ===
    if (!extractedFlavor && /雪花餅/.test(fullTextNoSpace)) {
        if (/蔓越莓/.test(fullTextNoSpace)) extractedFlavor = '雪花餅-蔓越莓';
        else if (/巧克力/.test(fullTextNoSpace)) extractedFlavor = '雪花餅-巧克力';
        else if (/金沙/.test(fullTextNoSpace)) extractedFlavor = '雪花餅-金沙';
        else if (/抹茶/.test(fullTextNoSpace)) extractedFlavor = '雪花餅-抹茶';
        else if (/肉鬆/.test(fullTextNoSpace)) extractedFlavor = '雪花餅-肉鬆';
        else if (/綜合/.test(fullTextNoSpace)) extractedFlavor = '雪花餅-綜合';
    }

    // === 南棗核桃糕 ===
    if (!extractedFlavor && /南棗核桃糕/.test(fullTextNoSpace)) extractedFlavor = '南棗核桃糕';

    // === 鳳凰酥 ===
    if (!extractedFlavor && /鳳凰酥/.test(fullTextNoSpace)) extractedFlavor = '鳳凰酥';

    // === 雙塔 (非禮盒) ===
    if (!extractedFlavor && /招牌雙塔組合|招牌雙塔/.test(fullTextNoSpace) && !/禮盒/.test(fullTextNoSpace)) extractedFlavor = '雙塔';

    if (extractedFlavor) productName = extractedFlavor;

    // 提取倍數
    let multiplier = 1;
    const multiplierMatch = pickingSpec.match(/x\s*(\d+)\s*包/);
    if (multiplierMatch) multiplier = parseInt(multiplierMatch[1]) || 1;

    let column = null;
    let spec = null;
    let confidence = 0.3;

    // === 規格判斷 (優先順序：pickingSpec > pickingName) ===
    // 正規化全形 g 為半形
    const normalizedSpecNoSpace = specNoSpace.replace(/ｇ/g, 'g');
    const normalizedNameNoSpace = nameNoSpace.replace(/ｇ/g, 'g');

    // 先檢查來源規格是否包含有效的規格資訊 (克重或入數)
    const hasValidSpec = /\d+g|\d+入/.test(normalizedSpecNoSpace);
    const specSource = hasValidSpec ? normalizedSpecNoSpace : normalizedNameNoSpace;

    if (/禮盒/.test(fullText)) { column = 'B'; spec = '禮盒'; confidence = 0.95; multiplier = 1; }
    // 克重規格：優先從 specSource 判斷 (從大到小)
    else if (/300g/.test(specSource)) { column = 'B'; spec = '300g'; confidence = 0.95; multiplier = 1; }
    else if (/280g/.test(specSource)) { column = 'C'; spec = '280g'; confidence = 0.95; multiplier = 1; }
    else if (/200g/.test(specSource)) { column = 'C'; spec = '200g'; confidence = 0.95; multiplier = 1; }
    else if (/150g/.test(specSource)) { column = 'B'; spec = '150g'; confidence = 0.95; multiplier = 1; }
    else if (/135g/.test(specSource)) { column = 'C'; spec = '135g'; confidence = 0.95; multiplier = 1; }
    else if (/120g/.test(specSource)) { column = 'B'; spec = '120g'; confidence = 0.95; multiplier = 1; }
    else if (/90g/.test(specSource)) { column = 'B'; spec = '90g'; confidence = 0.95; multiplier = 1; }
    else if (/60g/.test(specSource)) { column = 'B'; spec = '60g'; confidence = 0.95; multiplier = 1; }
    else if (/50g/.test(specSource)) { column = 'B'; spec = '50g'; confidence = 0.95; multiplier = 1; }
    else if (/45g/.test(specSource)) { column = 'B'; spec = '45g'; confidence = 0.95; multiplier = 1; }
    // 入數規格
    else if (/15入/.test(specSource)) { column = 'C'; spec = '15入袋裝'; confidence = 0.9; }
    else if (/12入/.test(specSource)) { column = 'C'; spec = '12入袋裝'; confidence = 0.9; }
    else if (/10入/.test(specSource)) { column = 'B'; spec = '10入袋裝'; confidence = 0.9; }
    else if (/8入/.test(specSource)) { column = 'B'; spec = '8入袋裝'; confidence = 0.9; }
    // 特殊商品規格
    else if (productName === '瓦片-原味45克') { column = 'B'; spec = null; confidence = 0.95; multiplier = 1; }
    else if (productName === '千層-小酥條') { column = 'B'; spec = '小包裝'; confidence = 0.95; multiplier = 1; }
    else if (/單顆/.test(fullText) || /單個/.test(fullText) || /活動/.test(fullText)) {
        const isMadeleine = /瑪德蓮/.test(fullTextNoSpaceTemp);
        column = isMadeleine ? 'C' : 'D'; spec = '單顆'; confidence = 0.85;
        const countMatch = specNoSpace.match(/(\d+)入/);
        if (countMatch) multiplier = parseInt(countMatch[1]);
    }

    // 標準化商品名稱
    const standardizedName = normalizeProductName(productName);

    // 如果沒有成功映射到商品，清除欄位和規格
    const finalProduct = standardizedName || null;
    if (!finalProduct) {
        column = null;
        spec = null;
    }

    return {
        templateProduct: finalProduct,
        templateColumn: column,
        templateSpec: spec,
        multiplier: multiplier,
        mappedQuantity: finalProduct ? quantity * multiplier : 0,
        confidence: finalProduct ? confidence : 0
    };
}
