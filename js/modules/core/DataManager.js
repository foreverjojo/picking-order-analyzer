/**
 * @file DataManager.js
 * @description 數據管理模組，負責載入與提供商品核心資料庫
 * @author Ivy House TW Development Team
 */

let products = [];
let isLoaded = false;

/**
 * 初始化並載入商品資料庫
 */
export async function initializeData() {
    if (isLoaded) return products;

    try {
        const response = await fetch('js/data/products.json');
        if (!response.ok) throw new Error('無法讀取商品資料庫');
        products = await response.json();
        isLoaded = true;
        console.log('✅ 商品資料庫載入成功', products.length, '筆產品');
        return products;
    } catch (error) {
        console.error('❌ 商品資料庫載入失敗:', error);
        // 如果失敗，嘗試從 data 目錄下的 JS 備份載入，或拋出錯誤
        throw error;
    }
}

/**
 * 獲取所有商品對象
 */
export function getAllProducts() {
    return products;
}

/**
 * 獲取所有標準商品名稱
 */
export function getProductNames() {
    return products.map(p => p.name);
}

/**
 * 獲取商品排序清單
 */
export function getProductSortOrder() {
    return products.map(p => p.name);
}

/**
 * 根據名稱獲取商品資訊
 */
export function getProductByName(name) {
    return products.find(p => p.name === name);
}

/**
 * 資料是否已載入
 */
export function isDataReady() {
    return isLoaded;
}
