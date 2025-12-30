/**
 * @file UIManager.js
 * @description UI 管理模組
 * @author Ivy House TW Development Team
 */

import { state } from '../core/StateManager.js';
import { standardProductNames } from '../rules/StandardProducts.js';
import { getProductSortOrder } from '../rules/MappingEngine.js';
import { formatFileSize } from './UIUtils.js';

// 規格與欄位選項
const specOptions = ['', '45g', '50g', '60g', '90g', '120g', '135g', '150g', '200g', '280g', '300g', '8入袋裝', '10入袋裝', '12入袋裝', '15入袋裝', '單顆', '禮盒', '小包裝'];
const columnOptions = ['', 'B', 'C', 'D', 'E'];

/**
 * 切換步驟
 */
export function navigateToStep(stepNumber) {
    for (let i = 1; i <= 4; i++) {
        const step = document.getElementById(`step${i}`);
        if (step) {
            if (i === stepNumber) step.classList.remove('hidden');
            else step.classList.add('hidden');
        }
    }
}

/**
 * 更新檔案列表顯示
 */
export function updateFileList() {
    const fileList = document.getElementById('fileList');
    if (!fileList) return;
    fileList.innerHTML = '';

    const labels = {
        momo: { name: 'MOMO 撿貨單', icon: '📊' },
        official: { name: '官網撿貨單', icon: '📊' },
        shopee: { name: '蝦皮撿貨單', icon: '📄' },
        orangepoint: { name: '橘點子撿貨單', icon: '🍊' },
        template: { name: '報表範本', icon: '📋' }
    };

    Object.entries(state.uploadedFiles).forEach(([key, file]) => {
        if (!file || !labels[key]) return;

        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';
        fileItem.innerHTML = `
            <div class="file-info">
                <div class="file-icon">${labels[key].icon}</div>
                <div class="file-details">
                    <h4>${labels[key].name}</h4>
                    <span class="file-size">${file.name} (${formatFileSize(file.size)})</span>
                </div>
            </div>
            <button class="file-remove" data-platform="${key}">✕ 移除</button>
        `;
        fileList.appendChild(fileItem);
    });

    const parseBtn = document.getElementById('parseBtn');
    if (parseBtn) {
        const hasFiles = state.uploadedFiles.momo || state.uploadedFiles.official || state.uploadedFiles.shopee || state.uploadedFiles.orangepoint;
        parseBtn.disabled = !hasFiles;
    }
}

/**
 * 建立映射表格 (步驟 2)
 */
export function buildMappingTable() {
    const tbody = document.getElementById('mappingTableBody');
    if (!tbody) return;
    tbody.innerHTML = '';

    // 排序
    const sortOrder = getProductSortOrder();
    state.mappedProducts.sort((a, b) => {
        const orderA = sortOrder.indexOf(a.templateProduct);
        const orderB = sortOrder.indexOf(b.templateProduct);
        const rankA = orderA === -1 ? 9999 : orderA;
        const rankB = orderB === -1 ? 9999 : orderB;
        if (rankA !== rankB) return rankA - rankB;
        return (a.templateProduct || '').localeCompare(b.templateProduct || '', 'zh-TW');
    });

    state.mappedProducts.forEach((product, index) => {
        const row = document.createElement('tr');
        const confidence = Math.round((product.confidence || 0) * 100);
        const confidenceColor = confidence >= 90 ? '#10b981' : confidence >= 70 ? '#f59e0b' : '#ef4444';

        const productOptions = standardProductNames.map(name =>
            `<option value="${name}" ${product.templateProduct === name ? 'selected' : ''}>${name}</option>`
        ).join('');

        const columnOptionsHtml = columnOptions.map(col =>
            `<option value="${col}" ${product.templateColumn === col ? 'selected' : ''}>${col || '-'}</option>`
        ).join('');

        const specOptionsHtml = specOptions.map(spec =>
            `<option value="${spec}" ${product.templateSpec === spec ? 'selected' : ''}>${spec || '-'}</option>`
        ).join('');

        row.innerHTML = `
            <td title="${product.name}">${product.name.length > 25 ? product.name.substring(0, 25) + '...' : product.name}</td>
            <td><small>${product.spec || '-'}</small></td>
            <td><span style="color: ${getSourceColor(product.source)}">${product.source}</span></td>
            <td><strong>${product.quantity}</strong></td>
            <td style="text-align: center;">→</td>
            <td>
                <select id="mapped-name-${index}" class="mapping-select" style="background: ${product.templateProduct ? '#e8f5e9' : '#fff3e0'};">
                    <option value="">-- 選擇商品 --</option>
                    ${productOptions}
                </select>
            </td>
            <td>
                <select id="mapped-column-${index}" class="mapping-select-small">
                    ${columnOptionsHtml}
                </select>
            </td>
            <td>
                <select id="mapped-spec-${index}" class="mapping-select-med">
                    ${specOptionsHtml}
                </select>
            </td>
            <td><strong style="color: #3b82f6;">${product.mappedQuantity || product.quantity}</strong></td>
            <td><span style="color: ${confidenceColor}; font-weight: bold;">${confidence}%</span></td>
        `;
        tbody.appendChild(row);
    });
}

function getSourceColor(source) {
    const colors = { 'MOMO': '#f59e0b', '官網': '#10b981', '蝦皮': '#ef4444', '橘點子': '#FF6B00' };
    return colors[source] || '#cbd5e1';
}

/**
 * 顯示統計結果 (步驟 3)
 */
export function displayStatistics() {
    const container = document.getElementById('statsContainer');
    if (!container) return;
    container.innerHTML = '';

    const statsArray = Object.values(state.statistics);
    if (statsArray.length === 0) {
        container.innerHTML = '<p style="text-align: center; color: #f59e0b;">沒有統計資料</p>';
        return;
    }

    // 排序
    const sortOrder = getProductSortOrder();
    statsArray.sort((a, b) => {
        const orderA = sortOrder.indexOf(a.name);
        const orderB = sortOrder.indexOf(b.name);
        const rankA = orderA === -1 ? 9999 : orderA;
        const rankB = orderB === -1 ? 9999 : orderB;
        if (rankA !== rankB) return rankA - rankB;
        return a.name.localeCompare(b.name, 'zh-TW');
    });

    const table = document.createElement('table');
    table.className = 'mapping-table';
    table.innerHTML = `
        <thead>
            <tr><th>報表商品</th><th>口味</th><th>規格</th><th>欄位</th><th>總數量</th></tr>
        </thead>
        <tbody>
            ${statsArray.map(stat => `
                <tr>
                    <td>${stat.name}</td>
                    <td>${stat.flavor || '-'}</td>
                    <td>${stat.spec || '-'}</td>
                    <td>${stat.column || '-'}</td>
                    <td><strong>${stat.quantity}</strong></td>
                </tr>
            `).join('')}
        </tbody>
    `;
    container.appendChild(table);

    const totalQuantity = statsArray.reduce((sum, s) => sum + s.quantity, 0);
    const summary = document.createElement('div');
    summary.style.cssText = 'margin-top: 20px; text-align: center; font-size: 1.2em;';
    summary.innerHTML = `<strong>總計：${statsArray.length} 種商品，${totalQuantity} 件</strong>`;
    container.appendChild(summary);
}

/**
 * 從 UI 表格收集人工調整的對應
 */
export function collectMappingFromUI() {
    state.mappedProducts.forEach((product, index) => {
        const nameEl = document.getElementById(`mapped-name-${index}`);
        const colEl = document.getElementById(`mapped-column-${index}`);
        const specEl = document.getElementById(`mapped-spec-${index}`);

        if (nameEl) product.templateProduct = nameEl.value;
        if (colEl) product.templateColumn = colEl.value;
        if (specEl) product.templateSpec = specEl.value;
    });
}
