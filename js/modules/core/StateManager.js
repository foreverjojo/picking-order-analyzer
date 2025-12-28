/**
 * @file StateManager.js
 * @description 全域狀態管理模組
 * @purpose 集中管理應用程式狀態，如上傳檔案、解析數據、統計結果等
 * @author Ivy House TW Development Team
 */

export const state = {
    uploadedFiles: {
        momo: null,
        official: null,
        shopee: null,
        orangepoint: null,
        template: null
    },
    parsedData: {
        momo: [],
        official: [],
        shopee: [],
        orangepoint: []
    },
    allProducts: [],
    mappedProducts: [],
    statistics: {},
    templateWorkbook: null, // SheetJS
    excelWorkbook: null,    // ExcelJS
    templateBuffer: null    //原始二進制數據
};

/**
 * 重設狀態
 */
export function resetState() {
    state.uploadedFiles = { momo: null, official: null, shopee: null, orangepoint: null, template: null };
    state.parsedData = { momo: [], official: [], shopee: [], orangepoint: [] };
    state.allProducts = [];
    state.mappedProducts = [];
    state.statistics = {};
    state.templateWorkbook = null;
    state.excelWorkbook = null;
    state.templateBuffer = null;
}

/**
 * 更新上傳檔案
 * @param {string} platform 平台名稱
 * @param {File} file 檔案物件
 */
export function setUploadedFile(platform, file) {
    if (Object.keys(state.uploadedFiles).includes(platform)) {
        state.uploadedFiles[platform] = file;
    }
}

/**
 * 獲取所有解析後的商品扁平列表
 * @returns {Array} 
 */
export function getRawParsedData() {
    return [
        ...state.parsedData.momo,
        ...state.parsedData.official,
        ...state.parsedData.shopee,
        ...state.parsedData.orangepoint
    ];
}
