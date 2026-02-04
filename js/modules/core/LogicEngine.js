/**
 * @file LogicEngine.js
 * @description 核心業務邏輯模組
 * @purpose 處理商品拆分、整合、統計計算等業務邏輯
 * @author Ivy House TW Development Team
 */

import { state } from './StateManager.js';
import { autoMapProduct } from '../rules/MappingEngine.js';

/**
 * 檢查是否為多口味組合
 * @param {string} spec 規格字串
 * @returns {boolean}
 */
function isMultiFlavorCombo(spec) {
    if (!spec) return false;
    const specNoSpace = spec.replace(/\s+/g, '');
    // 檢查是否有多個不同口味組合 (至少 2 個 xN[袋包] 且用 / 分隔)
    const matches = specNoSpace.match(/x\d+[袋包]/g);
    return matches && matches.length >= 2 && /\//.test(specNoSpace);
}

/**
 * 解析多口味組合
 * @param {string} fullText 完整文字 (商品名稱+規格)
 * @param {string} spec 規格字串
 * @param {string} productName 原始商品名稱
 * @returns {Array} 解析結果陣列
 */
function parseMultiFlavorCombo(fullText, spec, productName) {
    const results = [];

    // 判斷商品類型
    const fullTextNoSpace = fullText.replace(/\s+/g, '');
    let productType = '未知';
    if (/夏威夷豆塔/.test(fullTextNoSpace)) productType = '豆塔';
    else if (/堅果塔/.test(fullTextNoSpace)) productType = '堅果塔';
    else if (/雪花餅/.test(fullTextNoSpace)) productType = '雪花餅';
    else if (/瑪德蓮/.test(fullTextNoSpace)) productType = '瑪德蓮';
    else if (/瓦片/.test(fullTextNoSpace)) productType = '瓦片';

    // MOMO 常見格式：無 / 分隔，口味段落以空白串接，例如：綜合口味10入x2袋 蜂蜜蔓越莓口味10入x1袋
    if (!/\//.test(spec || '')) {
        const comboTextNoSpace = (fullText || '').replace(/\s+/g, '');
        const segmentRegex = /(蜂蜜蔓越莓|蔓越莓|焦糖|巧克力|抹茶|椒麻|綜合|蜂蜜|原味|紅茶|海苔|黑糖|青花椒|金沙|肉鬆|檸檬|柑橘|咖哩)(?:口味)?(?:(\d+)入)?x(\d+)[袋包]/g;

        for (const match of comboTextNoSpace.matchAll(segmentRegex)) {
            let flavor = match[1];
            // 特殊處理: 蜂蜜蔓越莓 → 蔓越莓
            if (flavor === '蜂蜜蔓越莓') flavor = '蔓越莓';

            const unitCount = match[2] ? parseInt(match[2], 10) : 10;
            const multiplier = parseInt(match[3], 10);

            if (Number.isFinite(multiplier) && multiplier > 0) {
                results.push({
                    productType: productType,
                    flavor: flavor,
                    multiplier: multiplier,
                    spec: `${unitCount}入袋裝`
                });
            }
        }

        return results;
    }

    const parts = (spec || '').split('/');

    parts.forEach(part => {
        const partNoSpace = part.replace(/\s+/g, '');

        // 提取口味關鍵字 (擴充支援更多口味)
        const flavorMatch = partNoSpace.match(/(蜂蜜蔓越莓|蔓越莓|焦糖|巧克力|抹茶|椒麻|綜合|蜂蜜|原味|紅茶|海苔|黑糖|青花椒|金沙|肉鬆|檸檬|柑橘|咖哩)/);

        // 提取倍數
        const multiplierMatch = partNoSpace.match(/x(\d+)[袋包]/);

        // 提取規格 (入數)
        const specMatch = partNoSpace.match(/(\d+)入/);

        if (flavorMatch && multiplierMatch) {
            let flavor = flavorMatch[1];
            // 特殊處理: 蜂蜜蔓越莓 → 蔓越莓
            if (flavor === '蜂蜜蔓越莓') flavor = '蔓越莓';

            results.push({
                productType: productType,
                flavor: flavor,
                multiplier: parseInt(multiplierMatch[1]),
                spec: specMatch ? specMatch[1] + '入袋裝' : '10入袋裝'
            });
        }
    });

    return results;
}

/**
 * 拆分特殊商品規則 (例如：瑪德蓮 2 入拆分為單顆)
 * @param {Array} products 原始商品列表
 * @returns {Array} 拆分後的商品列表
 */
export function splitSpecialProducts(products) {
    const result = [];

    products.forEach(product => {
        const fullText = ((product.name || '') + (product.spec || '')).replace(/\s+/g, '');

        // === 新增: MOMO 多口味組合拆分（支援無 / 分隔，僅套用 source=MOMO） ===
        if (product.source === 'MOMO') {
            const specNoSpace = (product.spec || '').replace(/\s+/g, '');
            const multiplierTokens = specNoSpace.match(/x\d+[袋包]/g);
            if (multiplierTokens && multiplierTokens.length >= 2) {
                const flavorCombos = parseMultiFlavorCombo(fullText, product.spec || '', product.name);
                const distinctFlavors = new Set(flavorCombos.map(c => c.flavor));
                if (flavorCombos.length >= 2 && distinctFlavors.size >= 2) {
                    flavorCombos.forEach(combo => {
                        result.push({
                            ...product,
                            name: `${combo.productType}-${combo.flavor}（組合拆分）`,
                            spec: combo.spec,
                            quantity: product.quantity * combo.multiplier,
                            originalName: product.name,
                            originalSpec: product.spec,
                            isSplit: true
                        });
                    });
                    return;
                }
            }
        }

        // === 新增: 通用多口味組合拆分 ===
        if (isMultiFlavorCombo(product.spec)) {
            const flavorCombos = parseMultiFlavorCombo(fullText, product.spec, product.name);

            if (flavorCombos.length >= 2) {
                // 成功解析出多個口味，進行拆分
                flavorCombos.forEach(combo => {
                    result.push({
                        ...product,
                        name: `${combo.productType}-${combo.flavor}（組合拆分）`,
                        spec: combo.spec,
                        quantity: product.quantity * combo.multiplier,
                        originalName: product.name,
                        originalSpec: product.spec,
                        isSplit: true
                    });
                });
                return; // 跳過原始商品，不加入 result
            }
        }

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
