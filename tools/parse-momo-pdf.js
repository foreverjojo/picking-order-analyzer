#!/usr/bin/env node
/**
 * parse-momo-pdf.js
 * 用 Node + pdfjs-dist 解析 MOMO PDF（或一般格式的撿貨單 PDF）
 * Usage: node tools/parse-momo-pdf.js sample-data/momo-全部撿貨單.pdf
 * 輸出 JSON 到 stdout
 */

const fs = require('fs');
const path = require('path');
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');

function groupByY(items, tolerance = 5) {
  const sorted = [...items].sort((a, b) => b.y - a.y);
  const groups = [];
  sorted.forEach(item => {
    let found = false;
    for (let g of groups) {
      if (Math.abs(g.y - item.y) <= tolerance) {
        g.items.push(item);
        found = true;
        break;
      }
    }
    if (!found) groups.push({ y: item.y, items: [item] });
  });
  groups.forEach(g => g.items.sort((a, b) => a.x - b.x));
  return groups;
}

function parseShopeePage(items) {
  const rows = groupByY(items);
  const columnRanges = {
    seq: { min: 32, max: 59 },
    productName: { min: 166, max: 307 },
    spec: { min: 308, max: 396 },
    quantity: { min: 397, max: 520 }
  };

  let headerY = null;
  for (let row of rows) {
    if (row.items.some(i => i.text.includes('商品名稱') || i.text.includes('商品'))) {
      headerY = row.y;
      break;
    }
  }

  if (!headerY) return { products: [], rows, headerFound: false };

  const products = [];
  let currentProduct = null;

  rows.forEach(row => {
    if (row.y >= headerY) return; // skip header and rows below header (consistent with browser parser)

    let seqNum = null;
    row.items.forEach(item => {
      if (item.x >= columnRanges.seq.min && item.x <= columnRanges.seq.max) {
        const num = parseInt(item.text.trim());
        if (!isNaN(num) && num > 0) seqNum = num;
      }
    });

    if (seqNum) {
      if (currentProduct) products.push(currentProduct);
      currentProduct = { seq: seqNum, name: [], spec: [], quantity: null };
    }

    if (currentProduct) {
      const rowText = row.items.map(i => i.text).join(' ');
      if (rowText.includes('合計')) return;

      row.items.forEach(item => {
        const x = item.x;
        const t = item.text.trim();
        if (x >= columnRanges.productName.min && x <= columnRanges.productName.max) {
          if (!t.match(/^\d+$/)) currentProduct.name.push(t);
        } else if (x >= columnRanges.spec.min && x <= columnRanges.spec.max) {
          if (!t.match(/^\d+$/)) currentProduct.spec.push(t);
        } else if (x >= columnRanges.quantity.min && x <= columnRanges.quantity.max) {
          const num = parseInt(t);
          if (!isNaN(num) && num > 0) currentProduct.quantity = num;
        }
      });
    }
  });

  if (currentProduct) products.push(currentProduct);

  return {
    products: products.filter(p => p.seq && p.seq > 0).map(p => ({
      name: p.name.join(' ').trim(),
      spec: p.spec.join(' ').trim(),
      quantity: p.quantity || 0
    })),
    rows,
    headerFound: true
  };
}

function simpleLineParse(items) {
  // fallback: join items by Y and make heuristic lines, then extract lines ending with a number
  const lines = groupByY(items).map(g => g.items.map(i => i.text).join(' '));

  const products = [];
  lines.forEach(line => {
    // match lines ending with quantity number
    const m = line.match(/^(.+?)\s+(\d+)$/);
    if (m) {
      const name = m[1].trim();
      const qty = parseInt(m[2]);
      if (name && qty > 0) products.push({ name, spec: '', quantity: qty });
    }
  });
  return products;
}

(async function main() {
  try {
    const arg = process.argv[2];
    if (!arg) {
      console.error('Usage: node tools/parse-momo-pdf.js path/to/file.pdf');
      process.exit(2);
    }

    const filePath = path.resolve(arg);
    if (!fs.existsSync(filePath)) {
      console.error('File not found:', filePath);
      process.exit(3);
    }

    const data = fs.readFileSync(filePath);
    const pdf = await pdfjsLib.getDocument({ data }).promise;

    const pages = [];

    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const textContent = await page.getTextContent();
      const items = textContent.items.map(item => ({
        text: String(item.str).trim(),
        x: Math.round(item.transform[4]),
        y: Math.round(item.transform[5])
      }));

      pages.push({ page: i, items });
    }

    // Detect MOMO shop signature
    const fullText = pages.map(p => p.items.map(i => i.text).join(' ')).join('\n');
    const momoDetected = /艾薇手工坊.*MO店|MO店|MO 店|MO店\+|MO店\+/.test(fullText);

    // Try page-by-page parsing with Shopee parser logic
    const parsedProducts = [];
    let headerFoundAny = false;

    for (const p of pages) {
      const res = parseShopeePage(p.items);
      if (res.headerFound) {
        headerFoundAny = true;
      }
      if (res.products && res.products.length) {
        res.products.forEach(prod => parsedProducts.push(prod));
      }
    }

    // If no header found or no products, fallback to simple line-based parse
    if (!headerFoundAny || parsedProducts.length === 0) {
      for (const p of pages) {
        const fallback = simpleLineParse(p.items);
        if (fallback.length) parsedProducts.push(...fallback);
      }
    }

    // De-duplicate and combine quantities for identical names/spec
    const combined = {};
    parsedProducts.forEach(p => {
      const key = (p.name || '').trim() + '||' + (p.spec || '').trim();
      if (!combined[key]) combined[key] = { name: p.name, spec: p.spec, quantity: 0 };
      combined[key].quantity += p.quantity || 0;
    });

    const outputProducts = Object.values(combined);

    const output = {
      file: filePath,
      pages: pages.length,
      momoDetected,
      products: outputProducts,
      rawTextSample: fullText.slice(0, 2000)
    };

    console.log(JSON.stringify(output, null, 2));
  } catch (err) {
    console.error('解析失敗:', err);
    process.exit(1);
  }
})();
