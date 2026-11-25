/*
 * TC Excel Tool - PDF parsers (H&M invoice, H&M purchase order, H&M BOM, DO file name).
 * Input for every parser: array of pages, each page = array of text items
 *   { str, x, y, w }   (x/y in PDF points, y grows upwards, as given by pdf.js)
 * Pure functions, no DOM: works in the browser (window.TcParsers) and in Node (module.exports).
 */
(function (root, factory) {
    if (typeof module === 'object' && module.exports) module.exports = factory();
    else root.TcParsers = factory();
})(typeof self !== 'undefined' ? self : this, function () {
    'use strict';

    // ---------------------------------------------------------------- helpers
    const ROW_TOL = 2.5;

    function clean(s) {
        return String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
    }

    // pdf.js sometimes splits a word into single letters ("K","N","I"...). Merge items on the
    // same row whose horizontal gap is tiny, so "KNITTED" becomes one item again.
    function normalizePage(items) {
        const arr = items
            .filter(i => clean(i.str) !== '')
            .map(i => ({ str: i.str, x: i.x, y: i.y, w: i.w || 0 }))
            .sort((a, b) => (Math.abs(a.y - b.y) > ROW_TOL ? b.y - a.y : a.x - b.x));
        const out = [];
        for (const it of arr) {
            const last = out[out.length - 1];
            if (last && Math.abs(last.y - it.y) <= ROW_TOL && last.w > 0 &&
                it.x - (last.x + last.w) < 1.5 && it.x >= last.x) {
                last.str += it.str;
                last.w = it.x + it.w - last.x;
            } else {
                out.push({ ...it });
            }
        }
        out.forEach(i => { i.str = clean(i.str); });
        return out;
    }

    function pageText(items) {
        return items.map(i => i.str).join(' ');
    }

    function find(items, re, pred) {
        return items.find(i => re.test(i.str) && (!pred || pred(i))) || null;
    }

    // Items on the same row as label, right of it, x < xMax
    function inline(items, label, xMax) {
        if (!label) return '';
        const end = label.x + Math.max(label.w, 1) - 1;
        return clean(items
            .filter(i => i !== label && Math.abs(i.y - label.y) <= ROW_TOL && i.x > end && i.x < (xMax || 9999))
            .sort((a, b) => a.x - b.x)
            .map(i => i.str).join(' '));
    }

    // Lines below label, within [xMin, xMax), above stopY (exclusive)
    function blockLines(items, label, xMin, xMax, stopY) {
        if (!label) return [];
        const sel = items.filter(i => i.y < label.y - ROW_TOL && i.y > stopY + ROW_TOL && i.x >= xMin && i.x < xMax);
        const rows = [];
        sel.sort((a, b) => (Math.abs(a.y - b.y) > ROW_TOL ? b.y - a.y : a.x - b.x)).forEach(i => {
            const r = rows[rows.length - 1];
            if (r && Math.abs(r.y - i.y) <= ROW_TOL) r.parts.push(i.str);
            else rows.push({ y: i.y, parts: [i.str] });
        });
        return rows.map(r => clean(r.parts.join(' ')));
    }

    function num(s) {
        if (s == null) return null;
        const t = String(s).replace(/[^\d.,\-]/g, '');
        if (!t) return null;
        // H&M PO uses "5 210" (space thousands) - spaces already removed above; invoice uses "1017.00"
        const v = parseFloat(t.replace(/,/g, ''));
        return isNaN(v) ? null : v;
    }

    function poNum(s) { // "5 210" / "6 268" / "306"
        const t = String(s || '').replace(/\s/g, '');
        return /^\d+(\.\d+)?$/.test(t) ? parseFloat(t) : null;
    }

    const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
    function isoFromText(s) { // "31 Aug, 2026" -> "2026-08-31"; "2026-08-26" stays
        s = clean(s);
        let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (m) return s;
        m = s.match(/^(\d{1,2})\s+([A-Za-z]{3})[a-z]*,?\s+(\d{4})$/);
        if (m && MONTHS[m[2].toLowerCase()]) {
            return `${m[3]}-${String(MONTHS[m[2].toLowerCase()]).padStart(2, '0')}-${m[1].padStart(2, '0')}`;
        }
        return '';
    }

    // "60% ORGANIC COTTON 20% RECYCLE PRE CONSUMER COTTON 20% RECYCLE POST-CONSUMER POLYESTER"
    function parseComposition(text) {
        const t = clean(text).toUpperCase().replace(/(\d)\s*%/g, '$1%');
        const re = /(\d+(?:\.\d+)?)%\s*([A-Z][A-Z\s\-]*?)(?=\s*\d+(?:\.\d+)?%|,|$)/g;
        const out = [];
        let m;
        while ((m = re.exec(t)) !== null) {
            let mat = clean(m[2]).replace(/\bRECYCLE\b/g, 'RECYCLED').replace(/\bPOST\s*-?\s*CONSUMER\b/g, 'POST-CONSUMER')
                .replace(/\bPRE\s*-?\s*CONSUMER\b/g, 'PRE-CONSUMER');
            if (mat) out.push({ pct: parseFloat(m[1]), material: mat });
        }
        return out;
    }

    function programFor(material) {
        if (/ORGANIC/.test(material)) return 'OCS';
        if (/RECYCLED/.test(material)) return 'RCS/GRS';
        return '';
    }

    // ---------------------------------------------------------------- detect
    function detect(fileName, pages) {
        const all = pages.map(pageText).join(' ');
        const name = fileName || '';
        if (all.trim().length < 20) {
            if (parseDoFileName(name)) return 'do';
            return 'scan';
        }
        if (/Invoice No and Date/i.test(all) && /H&M Order No/i.test(all)) return 'invoice';
        if (/Purchase Order\s*(–|-)?\s*Online/i.test(all) || (/Order No:/.test(all) && /Planning Markets/.test(all))) return 'po';
        if (/Supplier Quote BOM|Component \/ Product|Placement/i.test(all) && /Composition/i.test(all)) return 'bom';
        return 'other';
    }

    // ---------------------------------------------------------------- DO file name
    // "000721 - LYG - EXP - VIII - 2026 - LD, OT, OU.pdf" -> {doNo:"000721/LYG-EXP/VIII/2026", codes:[LD,OT,OU]}
    function parseDoFileName(fileName) {
        const base = String(fileName || '').replace(/\.[^.]+$/, '');
        const m = base.match(/^\s*(\d{4,8})\s*-\s*LYG\s*-\s*EXP\s*-\s*([IVXLC]+)\s*-\s*(\d{2,4})\s*(?:-\s*(.*))?$/i);
        if (!m) return null;
        const codes = (m[4] || '').split(/[,\s]+/).map(s => s.trim().toUpperCase()).filter(s => /^[A-Z]{2}$/.test(s));
        const warnings = [];
        if (m[3].length !== 4) warnings.push(`Tahun di nama file DO "${m[3]}" tidak 4 digit (kemungkinan typo)`);
        return {
            doNo: `${m[1]}/LYG-EXP/${m[2].toUpperCase()}/${m[3]}`,
            seq: m[1], roman: m[2].toUpperCase(), year: m[3], codes, warnings
        };
    }

    // ---------------------------------------------------------------- DO scan header (OCR text)
    const MONTH_NAMES = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
    const WEEKDAYS = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];

    function editDistance(a, b) {
        const d = Array.from({ length: a.length + 1 }, (_, i) => [i]);
        for (let j = 1; j <= b.length; j++) d[0][j] = j;
        for (let i = 1; i <= a.length; i++) {
            for (let j = 1; j <= b.length; j++) {
                d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
            }
        }
        return d[a.length][b.length];
    }
    // closest word from a list, tolerating OCR typos ("Seplember", "Soturday")
    function fuzzyIndex(word, list, maxDist) {
        const w = String(word || '').toLowerCase();
        if (!w) return -1;
        let best = -1, bestD = maxDist + 1;
        list.forEach((x, i) => { const dd = editDistance(w, x); if (dd < bestD) { bestD = dd; best = i; } });
        return best;
    }

    /*
     * First line of "WAREHOUSE ADDRESS : PT. SARANA PRIMA LOGISTINDO ( SPL)   DELIVERY MODE ..." (OCR text).
     * Known warehouses are snapped to the spelling used in the H&M summary sheet; others are cleaned only.
     */
    function doWarehouse(lines) {
        const l = lines.find(x => /WAREHOUSE\s*ADDR/i.test(x));
        if (!l) return '';
        let v = l.replace(/^.*?WAREHOUSE\s*ADDRESS/i, '').replace(/\bDELIVERY\b.*$/i, '');
        if (/SARANA\s*PRIMA/i.test(v)) return 'PT SARANA PRIMA LOGISTINDO (SPL)';
        const jict = v.match(/J\s*I\s*C\s*T\W*U\s*T\s*C\W*(\d)?/i);
        if (jict) return 'JICT-UTC' + (jict[1] ? ' ' + jict[1] : '');
        v = v.replace(/^[^A-Za-z0-9(]*(?:\b[A-Za-z0-9]\b[^A-Za-z0-9(]+)?/, '') // OCR'd ":" / "©" / "1" before the value
            .replace(/[{]/g, '(').replace(/[}]/g, ')').replace(/\(\s+/g, '(').replace(/\s+\)/g, ')')
            .replace(/^PT\.\s*/i, 'PT ').replace(/\s+/g, ' ').trim();
        return v;
    }

    /**
     * Read DO number + date from the OCR text of the DO header,
     * e.g. "DELIVERY NO : 000720/LYG-EXP/VIII/2026 ..." and "... DATE : Monday, 31 August 2026".
     * OCR may misread the year (2026 -> 2020): when the weekday is readable, the year is corrected to the
     * nearest year (around yearHint) in which that day/month falls on that weekday.
     * Returns {seq, year, date (yyyy-mm-dd), dateRaw, checked (weekday matches), warnings}.
     */
    function parseDoOcr(text, yearHint) {
        const lines = String(text || '').split(/\r?\n/);
        const out = { seq: '', year: '', date: '', dateRaw: '', checked: false, warnings: [] };

        const noLine = lines.find(l => /DELIVERY\s*N/i.test(l)) || '';
        const mNo = noLine.match(/([0-9ODQ]{6})\s*\/\s*LYG/i);
        if (mNo) out.seq = mNo[1].replace(/[ODQ]/gi, '0');
        const mYr = noLine.match(/\/\s*(20\d{2})\b/);
        if (mYr) out.year = mYr[1];
        out.warehouse = doWarehouse(lines);

        let m = null;
        for (const l of lines) {
            const pos = l.search(/\bDATE\b/i);
            if (pos < 0) continue;
            m = l.slice(pos + 4).match(/([A-Za-z]{5,10})?[^A-Za-z0-9]*(\d{1,2})\s+([A-Za-z]{3,10})\s+(\d{4})/);
            if (m) break;
        }
        if (!m) { out.warnings.push('Tanggal DO tidak terbaca dari scan'); return out; }
        out.dateRaw = m[0].trim();
        const day = Number(m[2]);
        const mon = fuzzyIndex(m[3], MONTH_NAMES, 3);
        const wd = m[1] ? fuzzyIndex(m[1], WEEKDAYS, 3) : -1;
        const ocrYear = Number(m[4]);
        if (mon < 0) { out.warnings.push(`Bulan tidak dikenali di "${out.dateRaw}"`); return out; }

        const valid = y => { const dt = new Date(Date.UTC(y, mon, day)); return dt.getUTCMonth() === mon ? dt : null; };
        // year of the DO number (file name, else "…/2026" on the scan) is more reliable than the date's year:
        // OCR'd year is only trusted when it is within 1 year of it (05 Sep 2020 and 2026 are both Saturdays)
        const hint = (/^\d{4}$/.test(String(yearHint)) ? Number(yearHint) : 0) || Number(out.year) || ocrYear;
        const years = (Math.abs(ocrYear - hint) <= 1 ? [ocrYear] : []).concat([hint, hint - 1, hint + 1])
            .filter((y, i, a) => y > 2000 && a.indexOf(y) === i);
        let year = null;
        if (wd >= 0) {
            year = years.find(y => { const dt = valid(y); return dt && dt.getUTCDay() === wd; }) || null;
            out.checked = year != null;
        }
        if (year == null) year = valid(ocrYear) && Math.abs(ocrYear - hint) <= 1 ? ocrYear : (valid(hint) ? hint : null);
        if (year == null) { out.warnings.push(`Tanggal tidak valid di "${out.dateRaw}"`); return out; }
        // a year corrected by a matching weekday is reliable, no warning needed
        if (year !== ocrYear && !out.checked) out.warnings.push(`Tahun hasil OCR ${ocrYear} dikoreksi ke ${year}`);
        if (!out.checked) out.warnings.push(`Nama hari di tanggal DO tidak cocok / tidak terbaca ("${out.dateRaw}"): cek tanggal dengan scan`);
        out.date = `${year}-${String(mon + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        return out;
    }

    // ---------------------------------------------------------------- invoice
    function parseInvoice(pagesRaw) {
        const items = normalizePage(pagesRaw[0] || []);
        const w = [];
        const L = re => find(items, re);

        const lblNumber = find(items, /^Number:/);
        let invoiceNo = lblNumber ? clean(lblNumber.str.replace(/^Number:/, '')) || inline(items, lblNumber, 420) : '';
        const lblOrder = L(/^H&M Order No:/);
        const lblDate = find(items, /^Date:/, i => i.x >= 290 && i.x < 320 && (!lblOrder || i.y > lblOrder.y));
        let invoiceDate = lblDate ? clean(lblDate.str.replace(/^Date:/, '')) || inline(items, lblDate, 420) : '';
        const orderNo = lblOrder ? clean(lblOrder.str.replace(/^H&M Order No:/, '')) || inline(items, lblOrder, 600) : '';
        const lblProd = L(/^Product Id:/);
        const productId = lblProd ? clean(lblProd.str.replace(/^Product Id:/, '')) || inline(items, lblProd, 600) : '';

        // left column blocks
        const leftLabels = ['Exporter:', 'Supplier:', 'Actual Manufacturer:', 'Buyer:', 'Consignee:', 'Ultimate Consignee:', 'Mode of Transport:']
            .map(t => ({ t, it: find(items, new RegExp('^' + t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '$'), i => i.x < 120) }));
        function leftLines(t) {
            const idx = leftLabels.findIndex(l => l.t === t);
            const lab = leftLabels[idx] && leftLabels[idx].it;
            if (!lab) return [];
            const next = leftLabels.slice(idx + 1).map(l => l.it).filter(Boolean).filter(i => i.y < lab.y).sort((a, b) => b.y - a.y)[0];
            return blockLines(items, lab, 55, 295, next ? next.y : 0);
        }
        // printed lines joined with a space (same as the H&M TC summary sheet: line breaks are not commas)
        const leftBlock = t => clean(leftLines(t).join(' '));
        // right column
        const lblCoo = L(/^Country of Origin of Goods:/);
        const lblRemarks = L(/^Remarks:/);
        const countryOfOrigin = lblCoo ? blockLines(items, lblCoo, 295, 600, lblRemarks ? lblRemarks.y : lblCoo.y - 30).join(' ') : '';
        const lblMode = L(/^Mode of Transport:/);
        const remarks = lblRemarks ? blockLines(items, lblRemarks, 295, 600, lblMode ? lblMode.y : 0) : [];
        const remarkText = remarks.join(' | ');
        const mEf = remarkText.match(/E-filing Product Identifier Number:\s*\|?\s*([\w\-]+)/i);
        const mPn = remarkText.match(/Product Name:\s*([^|]+)/i);

        const lblPol = L(/^Port of Loading:/);
        const lblPod = L(/^Port of Discharge:/);
        const lblFd = L(/^Final Destination:/);
        const lblTod = L(/^Terms of Delivery:/);
        const lblTop = L(/^Terms of Payment:/);
        const lblWh = L(/^Warehouse ID$/);
        const mode = lblMode ? blockLines(items, lblMode, 55, lblFd ? lblFd.x - 2 : 180, lblPol ? lblPol.y : 0).join(' ') : '';
        const fdLines = lblFd ? blockLines(items, lblFd, lblFd.x - 2, 295, lblPod ? lblPod.y : 0) : [];
        let destCode = '';
        if (fdLines.length && /^[A-Z]{2}$/.test(fdLines[fdLines.length - 1])) destCode = fdLines.pop();
        const finalDestination = fdLines.join(' ');
        const termsDelivery = lblTod ? clean(lblTod.str.replace(/^Terms of Delivery:/, '')) || inline(items, lblTod, 600) : '';
        const portLoading = lblPol ? blockLines(items, lblPol, 55, lblPod ? lblPod.x - 2 : 180, lblWh ? lblWh.y : 0).join(' ') : '';
        const portDischarge = lblPod ? blockLines(items, lblPod, lblPod.x - 2, 295, lblWh ? lblWh.y : 0).join(' ') : '';
        const termsPayment = lblTop ? blockLines(items, lblTop, 295, 600, lblWh ? lblWh.y : 0).join(' ') : '';

        // goods table
        const lblCont = L(/^Container No:/);
        const lines = [];
        let description = '';
        let unit = '';
        if (lblWh && lblCont) {
            const top = lblWh.y, bot = lblCont.y;
            let region = items.filter(i => i.y < top - ROW_TOL && i.y > bot + ROW_TOL);
            const totalLbl = region.find(i => i.str === 'Total' && i.x > 330);
            if (totalLbl) region = region.filter(i => !(Math.abs(i.y - totalLbl.y) <= ROW_TOL && i.x >= totalLbl.x - 5));
            const descParts = region.filter(i => i.x >= 235 && i.x < 365)
                .sort((a, b) => (Math.abs(a.y - b.y) > ROW_TOL ? b.y - a.y : a.x - b.x)).map(i => i.str);
            const unitItem = region.find(i => i.x >= 365 && i.x < 425 && /^[A-Za-z]+$/.test(i.str));
            description = clean(descParts.join(' '));
            // unit word ("Pieces") is printed at the end of the first description line or in the qty column
            const mU = description.match(/\b(Pieces|Sets|Packs)\b/i);
            if (unitItem) unit = unitItem.str;
            else if (mU) unit = mU[1];
            description = clean(description.replace(/\b(Pieces|Sets|Packs)\b/gi, ' ').replace(/-\s+/g, '-'))
                .replace(/^\d+\s*CTNS?\s*=\s*\d+\s*PCS\s*/i, ''); // packing note "1 CTN = 14 PCS" is not part of the goods
            const qtyItems = region.filter(i => i.x >= 365 && i.x < 425 && poNum(i.str) !== null).sort((a, b) => b.y - a.y);
            const whItems = region.filter(i => i.x < 140).sort((a, b) => b.y - a.y);
            const pkgItems = region.filter(i => i.x >= 140 && i.x < 235 && /^\d+\s+\w+/.test(i.str)).sort((a, b) => b.y - a.y);
            qtyItems.forEach(q => {
                const sameRow = region.filter(i => Math.abs(i.y - q.y) <= ROW_TOL);
                const price = sameRow.find(i => i.x >= 425 && i.x < 475 && num(i.str) !== null);
                const amount = sameRow.find(i => i.x >= 475 && num(i.str) !== null);
                const wh = whItems.filter(i => i.y >= q.y - ROW_TOL)[whItems.filter(i => i.y >= q.y - ROW_TOL).length - 1] || whItems[0];
                const pk = pkgItems.filter(i => i.y >= q.y - ROW_TOL)[pkgItems.filter(i => i.y >= q.y - ROW_TOL).length - 1] || pkgItems[0];
                const mPk = pk ? pk.str.match(/^(\d+)\s+(\w+)/) : null;
                lines.push({
                    warehouseId: wh ? wh.str : '',
                    packages: mPk ? parseInt(mPk[1], 10) : null,
                    packageUnit: mPk ? mPk[2] : '',
                    qty: poNum(q.str),
                    price: price ? num(price.str) : null,
                    amount: amount ? num(amount.str) : null
                });
            });
            if (!lines.length) w.push('Baris barang (qty) tidak terbaca');
        } else {
            w.push('Tabel barang tidak ditemukan');
        }
        const currency = (items.find(i => i.str === 'USD' || i.str === 'IDR' || i.str === 'EUR') || {}).str || '';

        const lblHs = find(items, /^HS Code:/, i => i.x > 200);
        const hsCode = lblHs ? clean(lblHs.str.replace(/^HS Code:/, '')) || inline(items, lblHs, 360) : '';
        const lblCom = find(items, /^Country of$/);
        const countryOfManufacture = lblCom ? inline(items, lblCom, 300) : '';
        const lblNw = L(/^Net Weight:/);
        const lblGw = L(/^Gross Weight:/);
        const netWeight = lblNw ? num(clean(lblNw.str.replace(/^Net Weight:/, '')) || inline(items, lblNw, 300)) : null;
        const grossWeight = lblGw ? num(clean(lblGw.str.replace(/^Gross Weight:/, '')) || inline(items, lblGw, 300)) : null;
        const lblTad = L(/^Total after Discount$/);
        const totalAfterDiscount = lblTad ? num(inline(items, lblTad, 600)) : null;

        // footer "Invoice No: X, HM Order No: 711225, Invoice Date: 2026-08-26, ..."
        const footer = clean(items.filter(i => i.y < 100).sort((a, b) => b.y - a.y || a.x - b.x).map(i => i.str).join(' '));
        const mF = footer.match(/Invoice No:\s*([\w\-]+),\s*HM Order No:\s*(\d+),\s*Invoice Date:\s*(\d{4}-\d{2}-\d{2})/i);

        const qty = lines.reduce((s, l) => s + (l.qty || 0), 0);
        const packages = lines.reduce((s, l) => s + (l.packages || 0), 0);
        const amount = lines.reduce((s, l) => s + (l.amount || 0), 0);
        const composition = parseComposition(description);

        if (!invoiceNo) w.push('Nomor invoice tidak terbaca');
        if (!/^\d{4}-\d{2}-\d{2}$/.test(invoiceDate)) w.push(`Tanggal invoice tidak terbaca ("${invoiceDate}")`);
        if (netWeight == null) w.push('Net weight tidak terbaca');
        if (grossWeight == null) w.push('Gross weight tidak terbaca');
        if (netWeight != null && grossWeight != null && grossWeight < netWeight) w.push('Gross weight < net weight');
        if (mF) {
            if (mF[1] !== invoiceNo) w.push(`Footer invoice no (${mF[1]}) beda dengan header (${invoiceNo})`);
            if (mF[3] !== invoiceDate) w.push(`Footer invoice date (${mF[3]}) beda dengan header (${invoiceDate})`);
            if (orderNo && !orderNo.startsWith(mF[2])) w.push(`Footer order no (${mF[2]}) beda dengan header (${orderNo})`);
        }
        lines.forEach(l => {
            if (l.qty != null && l.price != null && l.amount != null && Math.abs(l.qty * l.price - l.amount) > 0.011)
                w.push(`Qty x harga (${l.qty} x ${l.price}) != amount ${l.amount}`);
        });
        if (totalAfterDiscount != null && Math.abs(totalAfterDiscount - amount) > 0.011)
            w.push(`Total after discount ${totalAfterDiscount} != jumlah baris ${amount.toFixed(2)}`);
        if (!composition.length) w.push('Komposisi tidak terbaca dari deskripsi barang');
        else {
            const tot = composition.reduce((s, c) => s + c.pct, 0);
            if (Math.abs(tot - 100) > 0.01) w.push(`Total komposisi ${tot}% (bukan 100%)`);
        }

        return {
            invoiceNo, invoiceDate, orderNo, productId,
            productName: mPn ? clean(mPn[1]) : '',
            eFilingId: mEf ? mEf[1] : '',
            countryOfOrigin,
            exporter: leftBlock('Exporter:'),
            // company name only (first printed line), e.g. "PT LEE YIN GAPURA GARMENT INDONESIA"
            // (some invoices print the street on the same line: cut at the first comma or "Jl." / "Jalan")
            productionUnit: clean(((leftLines('Actual Manufacturer:')[0]) || (leftLines('Exporter:')[0]) || '')
                .split(/,|\s(?=(?:Jl\.?|Jalan)\s)/i)[0].replace(/^PT\.\s*/i, 'PT ')),
            supplier: leftBlock('Supplier:'),
            actualManufacturer: leftBlock('Actual Manufacturer:'),
            buyer: leftBlock('Buyer:'),
            consignee: leftBlock('Consignee:'),
            ultimateConsignee: leftBlock('Ultimate Consignee:'),
            modeOfTransport: mode, finalDestination, destCode, termsDelivery, portLoading, portDischarge, termsPayment,
            lines, warehouseId: lines.map(l => l.warehouseId).filter(Boolean).join(', '),
            packages, packageUnit: (lines[0] || {}).packageUnit || '',
            qty, unit: unit || '', price: lines.length === 1 ? lines[0].price : (qty ? amount / qty : null),
            amount, currency, totalAfterDiscount,
            description, composition, hsCode, countryOfManufacture, netWeight, grossWeight,
            warnings: w
        };
    }

    // ---------------------------------------------------------------- purchase order
    function parsePo(pagesRaw) {
        const pages = pagesRaw.map(normalizePage);
        const p1 = pages[0] || [];
        const w = [];
        function field(items, label, xMin, xMax) {
            const lab = find(items, new RegExp('^' + label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '$'));
            if (!lab) return '';
            return clean(items.filter(i => Math.abs(i.y - lab.y) <= ROW_TOL && i.x >= xMin && i.x < xMax).map(i => i.str).join(' '));
        }
        const po = {
            orderNo: field(p1, 'Order No:', 100, 255),
            ptProdNo: field(p1, 'PT Prod No:', 100, 255),
            orderDate: isoFromText(field(p1, 'Date of Order:', 100, 255)),
            supplierCode: field(p1, 'Supplier Code:', 100, 255),
            supplierName: field(p1, 'Supplier Name:', 100, 255),
            optionNo: field(p1, 'Option No:', 100, 255),
            developmentNo: field(p1, 'Development No:', 100, 255),
            productNo: field(p1, 'Product No:', 360, 620),
            productName: field(p1, 'Product Name:', 360, 620),
            productDescription: field(p1, 'Product Description:', 360, 620),
            season: field(p1, 'Season:', 360, 620),
            customerGroup: field(p1, 'Customs Customer Group:', 360, 620),
            construction: field(p1, 'Type of Construction:', 360, 620),
            countryOfProduction: field(p1, 'Country of Production:', 100, 255),
            countryOfDelivery: field(p1, 'Country of Delivery:', 100, 255),
            countryOfOrigin: field(p1, 'Country of Origin:', 100, 255),
            termsPayment: field(p1, 'Terms of Payment:', 100, 255),
            noOfPieces: field(p1, 'No of Pieces:', 360, 620),
            salesMode: field(p1, 'Sales Mode:', 360, 620),
            termsDeliveryText: '', incoterm: '',
            todGroups: [], prices: [], articles: [], totalQty: null, details: []
        };
        const lblTd = find(p1, /^Terms of Delivery$/);
        if (lblTd) {
            const tl = blockLines(p1, lblTd, 0, 620, lblTd.y - 40);
            po.termsDeliveryText = tl.join(' | ');
            const mI = po.termsDeliveryText.match(/\b(FCA|FOB|CIF|CFR|EXW|DAP|DDP|CPT|CIP)\b/);
            po.incoterm = mI ? mI[1] : '';
        }
        // Time of Delivery table
        const lblTime = find(p1, /^Time of Delivery$/);
        const lblTotal = find(p1, /^Total:$/);
        if (lblTime) {
            const stop = lblTotal ? lblTotal.y : (find(p1, /^Quantity per Article$/) || { y: 0 }).y;
            const region = p1.filter(i => i.y < lblTime.y - ROW_TOL && i.y > stop + ROW_TOL);
            const rows = [];
            region.sort((a, b) => (Math.abs(a.y - b.y) > ROW_TOL ? b.y - a.y : a.x - b.x)).forEach(i => {
                const r = rows[rows.length - 1];
                if (r && Math.abs(r.y - i.y) <= ROW_TOL) r.items.push(i); else rows.push({ y: i.y, items: [i] });
            });
            rows.forEach(r => {
                const date = isoFromText(r.items.filter(i => i.x < 120).map(i => i.str).join(' '));
                const mk = clean(r.items.filter(i => i.x >= 120 && i.x < 440).map(i => i.str).join(' ')).replace(/\s*,\s*/g, ', ');
                const qtyTxt = r.items.filter(i => i.x >= 440 && i.x < 530).map(i => i.str).join(' ');
                if (!date) return;
                const markets = [];
                const re = /([A-Z]{2})\s*\(([^)]+)\)/g; let m;
                while ((m = re.exec(mk)) !== null) markets.push({ code: m[1], market: m[2] });
                po.todGroups.push({ tod: date, marketsText: mk, markets, qty: poNum(qtyTxt) });
            });
        }
        // Article / total qty / prices (page 1)
        const lblArt = find(p1, /^Article No$/);
        const lblTq = find(p1, /^Total Quantity:$/);
        if (lblArt && lblTq) {
            const region = p1.filter(i => i.y < lblArt.y - ROW_TOL && i.y > lblTq.y + ROW_TOL);
            const ys = [...new Set(region.map(i => Math.round(i.y)))];
            ys.forEach(y => {
                const r = region.filter(i => Math.abs(i.y - y) <= ROW_TOL);
                const g = (a, b) => clean(r.filter(i => i.x >= a && i.x < b).map(i => i.str).join(' '));
                po.articles.push({ articleNo: g(0, 95), colourCode: g(95, 180), ptArticle: g(180, 265), colour: g(265, 345), optionNo: g(345, 445), cost: g(445, 505), qty: poNum(g(505, 620)) });
            });
            po.totalQty = poNum(inline(p1, lblTq, 620));
        }
        const lblIap = find(p1, /^Invoice Average Price$/);
        const lblAcc = find(p1, /^By accepting/);
        if (lblIap) {
            const region = p1.filter(i => i.y < lblIap.y - ROW_TOL && i.y > (lblAcc ? lblAcc.y : 0) + ROW_TOL);
            const ys = [...new Set(region.map(i => Math.round(i.y)))].sort((a, b) => b - a);
            ys.forEach(y => {
                const r = region.filter(i => Math.abs(i.y - y) <= ROW_TOL);
                const pr = clean(r.filter(i => i.x < 130).map(i => i.str).join(' '));
                const cs = clean(r.filter(i => i.x >= 130).map(i => i.str).join(' '));
                const mP = pr.match(/^([\d.,]+)\s*([A-Z]{3})$/);
                if (!mP || !cs) return;
                cs.split(',').map(s => s.trim()).filter(Boolean).forEach(code => {
                    po.prices.push({ code, price: parseFloat(mP[1].replace(/,/g, '')), currency: mP[2] });
                });
            });
        }
        // Detail pages: one market each
        pages.slice(1).forEach((pg, idx) => {
            const td = find(pg, /^Terms of Delivery$/);
            if (!td) return;
            const code = (blockLines(pg, td, 0, 620, td.y - 16)[0] || '').trim();
            const tq = find(pg, /^Total Quantity:$/);
            const buyerLbl = find(pg, /^BUYER:$/);
            const buyer = buyerLbl ? blockLines(pg, { ...buyerLbl, y: buyerLbl.y + 3 }, 80, 620, buyerLbl.y - 60).join(', ') : '';
            po.details.push({ page: idx + 2, code, qty: tq ? poNum(inline(pg, tq, 620)) : null, buyer });
        });
        const sumTod = po.todGroups.reduce((s, g) => s + (g.qty || 0), 0);
        if (po.totalQty != null && sumTod && Math.abs(sumTod - po.totalQty) > 0.001) w.push(`Total qty PO ${po.totalQty} != jumlah tabel TOD ${sumTod}`);
        if (!po.orderNo) w.push('Order No PO tidak terbaca');
        if (!po.todGroups.length) w.push('Tabel Time of Delivery tidak terbaca');
        po.warnings = w;
        return po;
    }

    // ---------------------------------------------------------------- BOM (materials only)
    function parseBom(pagesRaw) {
        const pages = pagesRaw.map(normalizePage);
        const out = { materials: [], warnings: [] };
        for (let p = 0; p < pages.length; p += 2) {
            const left = pages[p], right = pages[p + 1] || [];
            if (!find(left, /^Position$/)) continue;
            const anchors = left.filter(i => i.x < 50 && /^(Shell|Lining|Trim|Pocketing|Interlining|Contrast|Main)$/i.test(i.str))
                .sort((a, b) => b.y - a.y);
            const sectionHdr = left.filter(i => /^(Material|Trims|Placement Graphic|Product Treatments|Labelling|Packaging)\s*\(\d+\)$/.test(i.str));
            anchors.forEach((a, k) => {
                const nextY = Math.max(
                    (anchors[k + 1] || { y: 0 }).y,
                    ...sectionHdr.filter(h => h.y < a.y - ROW_TOL).map(h => h.y),
                    find(left, /^Displaying/) ? find(left, /^Displaying/).y : 0
                );
                const band = left.filter(i => i.y <= a.y + 4 && i.y > nextY + ROW_TOL);
                const col = (x1, x2) => clean(band.filter(i => i.x >= x1 && i.x < x2).sort((m, n) => n.y - m.y || m.x - n.x).map(i => i.str).join(' '));
                const comp = col(250, 300);
                const rband = right.filter(i => i.y <= a.y + 4 && i.y > nextY + ROW_TOL);
                const wIdx = rband.findIndex(i => /^g\/m2?$/.test(i.str));
                let weight = '';
                if (wIdx >= 0) {
                    const g = rband[wIdx];
                    const n = rband.filter(i => Math.abs(i.y - g.y) <= ROW_TOL && i.x < g.x && i.x > g.x - 45 && poNum(i.str) !== null).pop();
                    weight = n ? `${n.str} ${g.str}` : g.str;
                }
                const supplier = clean(rband.filter(i => i.x >= 500 && poNum(i.str) === null).sort((m, n) => n.y - m.y || m.x - n.x).map(i => i.str).join(' '));
                const section = (sectionHdr.filter(h => h.y > a.y).sort((m, n) => m.y - n.y)[0] || {}).str || '';
                out.materials.push({
                    section: section.replace(/\s*\(\d+\)$/, ''),
                    position: a.str, placement: col(50, 95), masterComponent: col(125, 165), component: col(165, 205),
                    supplierArticle: col(205, 250), composition: comp, weight, supplier,
                    compositionParsed: parseComposition(comp)
                });
            });
        }
        if (!out.materials.length) out.warnings.push('Baris material BOM tidak terbaca');
        return out;
    }

    return {
        normalizePage, detect, parseInvoice, parsePo, parseBom, parseDoFileName, parseDoOcr, parseComposition, programFor, isoFromText
    };
});
