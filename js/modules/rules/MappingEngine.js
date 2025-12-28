/**
 * @file MappingEngine.js
 * @description 智慧商品映射核心引擎
 * @author Ivy House TW Development Team
 */

import { autoMapProductMomo } from './platforms/MomoMapping.js';
import { autoMapProductShopee } from './platforms/ShopeeMapping.js';
import { autoMapProductOrangePoint } from './platforms/OrangePointMapping.js';

import { getProductSortOrder as getSortOrder } from '../core/DataManager.js';

/**
 * 獲取商品排序列表
 */
export function getProductSortOrder() {
    return getSortOrder();
}

/**
 * 根據平台選擇映射函數
 */
export function autoMapProduct(pickingName, pickingSpec, quantity, platform = 'shopee') {
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
