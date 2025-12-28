/**
 * @file StandardProducts.js
 * @description 標準商品名稱定義與正規化邏輯
 * @author Ivy House TW Development Team
 */

import { getProductNames } from '../core/DataManager.js';

/**
 * 獲取目前的標準商品名稱列表
 * @returns {string[]}
 */
export function getStandardProductNames() {
    return getProductNames();
}

// 舊的匯出保持相容性（但現在會是空陣列直到初始化完成）
export let standardProductNames = [];

/**
 * 更新標準名稱緩存 (供內部使用)
 */
export function updateStandardNamesCache() {
    standardProductNames = getProductNames();
}

/**
 * 標準化商品名稱
 * @param {string} name 原始名稱
 * @returns {string|null} 標準化後的名稱
 */
export function normalizeProductName(name) {
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
