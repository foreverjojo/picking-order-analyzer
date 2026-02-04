#!/usr/bin/env python3
"""
parse_momo_pdf.py
使用 pdfminer.six 提取 PDF 文字與座標，並嘗試解析撿貨單中的商品名稱/規格/數量
用法: python tools/parse_momo_pdf.py sample-data/momo-全部撿貨單.pdf
輸出 JSON
"""
import sys
import json
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTTextLine, LTChar, LAParams
from collections import defaultdict


def extract_items_from_pdf(pdf_path):
    pages = []
    laparams = LAParams()
    for page_layout in extract_pages(pdf_path, laparams=laparams):
        items = []
        for element in page_layout:
            if isinstance(element, LTTextContainer):
                for text_line in element:
                    if isinstance(text_line, LTTextLine):
                        text = text_line.get_text().strip()
                        if not text:
                            continue
                        x0, y0, x1, y1 = text_line.bbox
                        # We'll use x0 and y1 (left and top) to approximate position
                        items.append({
                            'text': text,
                            'x': int(round(x0)),
                            'y': int(round(y1)),
                        })
        pages.append(items)
    return pages


def group_by_y(items, tol=5):
    items_sorted = sorted(items, key=lambda i: (-i['y'], i['x']))
    groups = []
    for it in items_sorted:
        placed = False
        for g in groups:
            if abs(g['y'] - it['y']) <= tol:
                g['items'].append(it)
                placed = True
                break
        if not placed:
            groups.append({'y': it['y'], 'items': [it]})
    # sort items in each group by x
    for g in groups:
        g['items'].sort(key=lambda i: i['x'])
    return groups


def parse_page_items(items):
    groups = group_by_y(items)
    # detect header row
    headerY = None
    for g in groups:
        rowtext = ' '.join([it['text'] for it in g['items']])
        if '商品名稱' in rowtext or '商品' in rowtext and '數量' in rowtext:
            headerY = g['y']
            break
    # Build lines (left-to-right)
    lines = []
    for g in groups:
        # skip header and below header (to mimic original behavior)
        if headerY is not None and g['y'] >= headerY:
            continue
        rowtext = ' '.join([it['text'] for it in g['items']])
        lines.append({'y': g['y'], 'text': rowtext})

    products = []
    for ln in lines:
        t = ln['text']
        if '合計' in t or '總計' in t:
            continue
        # heuristic: last token is quantity
        parts = t.split()
        if not parts:
            continue
        last = parts[-1]
        # sometimes quantity can be like 'x2' or '2' or '\u3000 2'
        q = None
        try:
            q = int(''.join(ch for ch in last if ch.isdigit()))
        except Exception:
            q = None
        if q is None or q == 0:
            # fallback: search for any small integer token in the line (prefer tail)
            for tok in reversed(parts):
                digits = ''.join(ch for ch in tok if ch.isdigit())
                if digits:
                    try:
                        val = int(digits)
                        if 0 < val < 10000:
                            q = val
                            break
                    except:
                        pass
        if q and q > 0:
            # remove the trailing quantity token(s) to get name/spec
            # find first occurrence of the quantity token from the end
            idx = None
            for i in range(len(parts)-1, -1, -1):
                if any(ch.isdigit() for ch in parts[i]):
                    # check that this token contains the digits we matched
                    # assume this is the quantity token
                    idx = i
                    break
            if idx is None:
                name_part = ' '.join(parts[:-1])
            else:
                name_part = ' '.join(parts[:idx])
            name_part = name_part.strip()
            if name_part:
                # try to split name and spec if pattern present (eg: name / spec or name (spec))
                spec = ''
                name = name_part
                # common pattern: name (spec)
                if '(' in name_part and ')' in name_part:
                    # split last parentheses
                    start = name_part.rfind('(')
                    spec = name_part[start+1:name_part.rfind(')')]
                    name = name_part[:start].strip()
                # pattern: name / spec
                elif '/' in name_part:
                    parts2 = [p.strip() for p in name_part.split('/')]
                    if len(parts2) >= 2:
                        name = parts2[0]
                        spec = ' / '.join(parts2[1:])
                products.append({'name': name, 'spec': spec, 'quantity': q, 'line': t})
    return products


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python tools/parse_momo_pdf.py path/to/file.pdf', file=sys.stderr)
        sys.exit(2)
    pdf_path = sys.argv[1]
    pages = extract_items_from_pdf(pdf_path)
    all_text = '\n'.join([' '.join([it['text'] for it in pg]) for pg in pages])
    momo_detected = False
    if '艾薇手工坊' in all_text and ('MO店' in all_text or 'MO 店' in all_text or 'MO店+' in all_text):
        momo_detected = True

    parsed = []
    for pg in pages:
        parsed += parse_page_items(pg)

    if not parsed:
        # fallback: try a simple line-based parse using all text lines
        lines = all_text.split('\n')
        for line in lines:
            line = line.strip()
            if not line or '合計' in line or '總計' in line:
                continue
            parts = line.split()
            if len(parts) < 2:
                continue
            # last token as qty
            last = parts[-1]
            digits = ''.join(ch for ch in last if ch.isdigit())
            if digits:
                q = int(digits)
                if q > 0:
                    name = ' '.join(parts[:-1])
                    parsed.append({'name': name.strip(), 'spec': '', 'quantity': q, 'line': line})

    # combine
    comb = {}
    for p in parsed:
        key = (p['name'].strip(), p.get('spec','').strip())
        if key not in comb:
            comb[key] = {'name': p['name'].strip(), 'spec': p.get('spec','').strip(), 'quantity': 0, 'exampleLine': p.get('line','')}
        comb[key]['quantity'] += p['quantity']

    output = {
        'file': pdf_path,
        'pages': len(pages),
        'momoDetected': momo_detected,
        'products': list(comb.values()),
        'rawTextSample': all_text[:2000]
    }

    print(json.dumps(output, ensure_ascii=False, indent=2))
