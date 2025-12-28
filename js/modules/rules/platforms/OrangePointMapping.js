/**
 * @file OrangePointMapping.js
 * @description 橘點子平台商品映射邏輯
 * @author Ivy House TW Development Team
 */

import { normalizeProductName } from '../StandardProducts.js';

/**
 * 橘點子專用映射函數
 * @param {string} productName 撿貨單品名
 * @param {number} quantity 數量
 * @returns {object} 映射結果
 */
export function autoMapProductOrangePoint(productName, quantity) {
    const nameNoSpace = productName.replace(/\s+/g, '');

    let mappedName = null;
    let column = null;
    let spec = null;
    let confidence = 0.3;

    // 瓦片
    if (/杏仁瓦片/.test(nameNoSpace)) {
        const flavorMatch = productName.match(/[\[\(](.+?)[\]\)]/);
        const flavor = flavorMatch ? flavorMatch[1].trim() : '';
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
