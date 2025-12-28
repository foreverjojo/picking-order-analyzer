/**
 * @file LogicEngine.js
 * @description 核心業務邏輯模組
 * @purpose 處理商品拆分、整合、統計計算等業務邏輯
 * @author Ivy House TW Development Team
 */

import { state } from './StateManager.js';
import { autoMapProduct } from '../rules/MappingEngine.js';

/**
 * 拆分特殊商品規則 (例如：瑪德蓮 2 入拆分為單顆)
 * @param {Array} products 原始商品列表
 * @returns {Array} 拆分後的商品列表
 */
export function splitSpecialProducts(products) {
    const result = [];

    products.forEach(product => {
        const fullText = ((product.name || '') + (product.spec || '')).replace(/\s+/g, '');

        // 胖貝殼瑪德蓮-2入 拆分為 蜂蜜 + 巧克力
        if (/胖貝殼瑪德蓮/.test(fullText) && /2入/.test(fullText)) {
            result.push({
                ...product,
                name: '瑪德蓮-蜂蜜（活動拆分）',
                spec: '單顆',
                quantity: product.quantity,
                originalName: product.name,
                originalSpec: product.spec,
                isSplit: true
            });
            result.push({
                ...product,
                name: '瑪德蓮-巧克力（活動拆分）',
                spec: '單顆',
                quantity: product.quantity,
                originalName: product.name,
                originalSpec: product.spec,
                isSplit: true
            });
        }
        // 夏威夷豆塔組合包 拆分
        else if (/夏威夷豆塔/.test(fullText) && /蜂蜜蔓越莓.*2包/.test(fullText) && /焦糖.*1包/.test(fullText)) {
            result.push({
                ...product,
                name: '豆塔-蔓越莓（組合拆分）',
                spec: '10入袋裝',
                quantity: product.quantity * 2,
                originalName: product.name,
                originalSpec: product.spec,
                isSplit: true
            });
            result.push({
                ...product,
                name: '豆塔-焦糖（組合拆分）',
                spec: '10入袋裝',
                quantity: product.quantity * 1,
                originalName: product.name,
                originalSpec: product.spec,
                isSplit: true
            });
        }
        else if (/夏威夷豆塔/.test(fullText) && /焦糖.*2包/.test(fullText) && /蜂蜜蔓越莓.*1包/.test(fullText)) {
            result.push({
                ...product,
                name: '豆塔-焦糖（組合拆分）',
                spec: '10入袋裝',
                quantity: product.quantity * 2,
                originalName: product.name,
                originalSpec: product.spec,
                isSplit: true
            });
            result.push({
                ...product,
                name: '豆塔-蔓越莓（組合拆分）',
                spec: '10入袋裝',
                quantity: product.quantity * 1,
                originalName: product.name,
                originalSpec: product.spec,
                isSplit: true
            });
        } else {
            result.push(product);
        }
    });

    return result;
}

/**
 * 整合所有平台的商品並執行映射
 */
export function consolidateAndMap() {
    state.allProducts = [];
    state.mappedProducts = [];

    // 扁平化並初步合併 (同名稱、同來源、同規格)
    const rawData = [
        ...state.parsedData.momo,
        ...state.parsedData.official,
        ...state.parsedData.shopee,
        ...state.parsedData.orangepoint
    ];

    rawData.forEach(product => {
        const existing = state.allProducts.find(p =>
            p.name === product.name && p.source === product.source && p.spec === product.spec
        );

        if (existing) {
            existing.quantity += product.quantity;
        } else {
            state.allProducts.push({ ...product });
        }
    });

    // 執行拆分邏輯
    state.allProducts = splitSpecialProducts(state.allProducts);

    // 執行自動映射
    state.allProducts.forEach(product => {
        let platform = 'momo';
        if (product.source === '蝦皮') platform = 'shopee';
        else if (product.source === '橘點子') platform = 'orangepoint';
        else if (product.source === '官網') platform = 'official';

        const mapping = autoMapProduct(product.name, product.spec || '', product.quantity, platform);

        state.mappedProducts.push({
            ...product,
            ...mapping
        });
    });

    return state.mappedProducts;
}

/**
 * 統計報表商品數量
 */
export function calculateStatistics() {
    state.statistics = {};

    state.mappedProducts.forEach(product => {
        const mappedName = product.templateProduct;
        if (!mappedName) return;

        const spec = product.templateSpec || 'default';
        const key = `${mappedName}_${spec}`;

        if (!state.statistics[key]) {
            state.statistics[key] = {
                name: mappedName,
                column: product.templateColumn || '',
                spec: product.templateSpec || '',
                flavor: product.spec || '',
                quantity: 0
            };
        }
        state.statistics[key].quantity += product.mappedQuantity || product.quantity;
    });

    return state.statistics;
}
