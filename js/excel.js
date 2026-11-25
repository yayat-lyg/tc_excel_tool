/*
 * TC Excel Tool - workbook builder (ExcelJS).
 * buildWorkbook(ExcelJS, model) -> ExcelJS.Workbook
 * Sheets: Ringkasan, Shipment TC, Berat Serat, PO vs Invoice, BOM, Temuan, Dokumen.
 * Values read from PDF / typed by the user = blue font; formulas = black.
 */
(function (root, factory) {
    if (typeof module === 'object' && module.exports) module.exports = factory();
    else root.TcExcel = factory();
})(typeof self !== 'undefined' ? self : this, function () {
    'use strict';

    const FONT = 'Arial';
    const C = {
        head: 'FF1F4E78', sub: 'FFDDEBF7', input: 'FF0000FF', yellow: 'FFFFFF00', red: 'FFF8CBAD',
        amber: 'FFFFEB9C', green: 'FFC6EFCE', border: 'FFBFBFBF', grey: 'FF595959'
    };
    const DATE = 'dd-mmm-yyyy';
    const INT = '#,##0;(#,##0);-';
    const DEC2 = '#,##0.00;(#,##0.00);-';
    const DEC3 = '#,##0.000;(#,##0.000);-';

    function d(iso) { // ISO date -> JS Date at UTC midnight (ExcelJS writes it as a date serial)
        if (!/^\d{4}-\d{2}-\d{2}$/.test(iso || '')) return null;
        return new Date(iso + 'T00:00:00Z');
    }
    function col(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
    function q(name) { return `'${name.replace(/'/g, "''")}'`; }
    function esc(s) { return String(s).replace(/"/g, '""'); }

    function styleCell(c, o) {
        o = o || {};
        c.font = { name: FONT, size: o.size || 10, bold: !!o.bold, italic: !!o.italic, color: { argb: o.color || 'FF000000' } };
        if (o.fill) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: o.fill } };
        if (o.fmt) c.numFmt = o.fmt;
        if (o.border !== false) c.border = { top: { style: 'thin', color: { argb: C.border } }, left: { style: 'thin', color: { argb: C.border } }, bottom: { style: 'thin', color: { argb: C.border } }, right: { style: 'thin', color: { argb: C.border } } };
        c.alignment = { vertical: 'top', wrapText: !!o.wrap, horizontal: o.align };
    }
    function header(ws, row, titles, widths) {
        titles.forEach((t, i) => {
            const c = ws.getCell(row, i + 1);
            c.value = t;
            styleCell(c, { bold: true, color: 'FFFFFFFF', fill: C.head, wrap: true, align: 'center' });
        });
        ws.getRow(row).height = 30;
        if (widths) widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    }
    function title(ws, text, sub) {
        ws.getCell('A1').value = text;
        ws.getCell('A1').font = { name: FONT, size: 14, bold: true };
        if (sub) {
            ws.getCell('A2').value = sub;
            ws.getCell('A2').font = { name: FONT, size: 9, italic: true, color: { argb: C.grey } };
        }
    }
    function put(ws, r, c, value, o) {
        const cell = ws.getCell(r, c);
        cell.value = (value === undefined || value === '' ) ? null : value;
        styleCell(cell, o);
        return cell;
    }
    function f(ws, r, c, formula, o) {
        const cell = ws.getCell(r, c);
        cell.value = { formula };
        styleCell(cell, o);
        return cell;
    }
    function statusRules(ws, ref) {
        ws.addConditionalFormatting({
            ref, rules: [
                { type: 'containsText', operator: 'containsText', text: 'OK', priority: 2, style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.green } } } },
                { type: 'containsText', operator: 'containsText', text: 'BEDA', priority: 1, style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.amber } } } },
                { type: 'containsText', operator: 'containsText', text: 'Lengkapi', priority: 1, style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.red } } } },
                { type: 'containsText', operator: 'containsText', text: 'Cek', priority: 1, style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.amber } } } }
            ]
        });
    }

    function buildWorkbook(ExcelJS, model) {
        const wb = new ExcelJS.Workbook();
        wb.creator = 'TC Excel Tool';
        wb.created = new Date();
        wb.calcProperties.fullCalcOnLoad = true;

        const inv = model.invoices;
        const po = model.po;
        const first = inv[0] || {};
        const orderNo = (po && po.orderNo) || first.orderNo || '';
        const S_SUM = 'Ringkasan', S_SHP = 'Shipment TC', S_FIB = 'Berat Serat', S_PO = 'PO vs Invoice', S_BOM = 'BOM', S_ISS = 'Temuan', S_DOC = 'Dokumen';
        const wsSum = wb.addWorksheet(S_SUM);
        const wsShp = wb.addWorksheet(S_SHP);
        const wsFib = wb.addWorksheet(S_FIB);
        const wsPo = wb.addWorksheet(S_PO);
        const wsBom = wb.addWorksheet(S_BOM);
        const wsIss = wb.addWorksheet(S_ISS);
        const wsDoc = wb.addWorksheet(S_DOC);

        // ------------------------------------------------ Shipment TC
        // Columns A–AD follow sheet "Summary Details" of the H&M consequence file (header row 1, data from row 2,
        // Calibri 11, no title/total rows) so the rows can be pasted straight into that file.
        // Tool-only helper columns start at AF.
        const SUMMARY = [
            ['Production Unit', 14.73], ['Style Name', 17.91], ['Order No', 14], ['Invoice No', 16], ['Invoice Date', 12],
            ['Net Weight', 12.54], ['Gross Weight', 13.54], ['Product ID', 11], ['Buyer', 13.63], ['Consignee', 14.36],
            ['Final Destination', 18], ['Final Destination Code', 10], ['Terms of Delivery', 12], ['Description of Goods', 24.73],
            ['Pieces', 11.27], ['Transport No', 14.27], ['Transport Date', 12], ['Address warehouse', 18.63], ['TOD Week', 10],
            ['TC NUMBER', 11], ['SUBMIT DATE TO cu', 12.63], ['REMARK', 16.18], ['Consequence Date', 16],
            // X–AD have no header in the summary file; titles added on request (X / AD meaning = best guess, to confirm)
            ['Konsumsi Kain / pcs (kg)', 13], ['Efficiency', 11], ['Waste %', 9], ['Kain Net (kg)', 12], ['Kain Gross (kg)', 12],
            ['GSM', 10], ['Input TC No (kain)', 13]
        ];
        const EXTRA = [['Warehouse ID', 11], ['Karton', 8], ['Production Date', 13], ['Amount', 11], ['Currency', 8],
            ['File Invoice', 24], ['File DO', 30], ['Sumber Tgl DO', 10], ['Catatan (tool)', 24], ['Status', 12]];
        const X0 = SUMMARY.length + 2; // AF
        const XC = name => col(X0 + EXTRA.findIndex(x => x[0] === name));
        const CAL = { name: 'Calibri', size: 11 };
        const sumHead = (c, text, argb) => { // same look as the summary file header
            c.value = text;
            c.font = { ...CAL, bold: true, color: { argb: 'FFFFFFFF' } };
            c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
            c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        };
        SUMMARY.forEach(([t, w], k) => {
            const c = wsShp.getCell(1, k + 1);
            if (t) sumHead(c, t, t === 'REMARK' ? 'FF002060' : (t === 'TC NUMBER' || t === 'SUBMIT DATE TO cu') ? 'FF203864' : 'FF1F3864');
            wsShp.getColumn(k + 1).width = w;
        });
        EXTRA.forEach(([t, w], k) => {
            const c = wsShp.getCell(1, X0 + k);
            c.value = t;
            styleCell(c, { bold: true, color: 'FFFFFFFF', fill: C.grey, wrap: true, align: 'center' });
            wsShp.getColumn(X0 + k).width = w;
        });
        wsShp.getCell('X1').note = 'Input per order: konsumsi kain per pcs dalam kg (contoh 0.12), termasuk sisa potong.';
        wsShp.getCell('Y1').note = 'Input per order: efficiency potong (0.885 = 88.5%).';
        wsShp.getCell('Z1').note = 'Rumus: 100 - Efficiency x 100.';
        wsShp.getCell('AA1').note = 'Rumus: Pieces x Konsumsi x Efficiency = kain yang jadi garmen.';
        wsShp.getCell('AB1').note = 'Rumus: Kain Net / Efficiency (= Pieces x Konsumsi) = kain terpakai termasuk waste.';
        wsShp.getCell('AC1').note = 'GSM kain shell, dari BOM.';
        wsShp.getCell('AD1').note = 'Input per order: nomor TC kain dari supplier (input TC) – mohon konfirmasi.';
        wsShp.getColumn(SUMMARY.length + 1).width = 3;
        wsShp.getRow(1).height = 43.5;

        const MON = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEPT', 'OCT', 'NOV', 'DEC'];
        const todWeek = code => { // TOD of the PO group holding this country code, e.g. "SEPT 7"
            const g = po && po.todGroups.find(x => x.markets.some(mk => mk.code === code));
            const dt = g && d(g.tod);
            return dt ? `${MON[dt.getUTCMonth()]} ${dt.getUTCDate()}` : '';
        };
        const shell = model.bom && model.bom.materials.find(mt => /shell/i.test(mt.position || mt.placement || '') && /\d/.test(mt.weight || ''));
        const gsm = shell ? shell.weight.replace(/\s*g\/m2\s*$/i, '') + ' GSM' : '';
        const numOr = v => (v === '' || v == null) ? null : (isNaN(Number(v)) ? v : Number(v));
        const S = model.settings || {};
        // Y = efficiency as a fraction (0.885); "88.5" typed as percent is converted
        let efficiency = numOr(S.colY);
        if (typeof efficiency === 'number' && efficiency > 1) efficiency = efficiency / 100;

        const r1 = 2;
        inv.forEach((i, k) => {
            const r = r1 + k;
            const m = i.manual || {};
            const dof = (model.doFiles || []).find(x => x.file === m.doFile) || {};
            const plain = (c, v, fmt) => { const cell = wsShp.getCell(r, c); cell.value = (v === undefined || v === '') ? null : v; cell.font = CAL; if (fmt) cell.numFmt = fmt; return cell; };
            const fx = (c, formula, fmt) => { const cell = wsShp.getCell(r, c); cell.value = { formula }; cell.font = CAL; if (fmt) cell.numFmt = fmt; return cell; };
            plain(1, i.productionUnit);
            plain(2, (po && po.productName) || i.productName);
            plain(3, i.orderNo);
            plain(4, i.invoiceNo);
            plain(5, d(i.invoiceDate), 'd-mmm-yy');
            plain(6, i.netWeight, '#,##0.00');
            plain(7, i.grossWeight, '#,##0.00');
            plain(8, numOr(i.productId));
            plain(9, i.buyer);
            plain(10, i.consignee);
            plain(11, i.finalDestination);
            plain(12, i.destCode);
            plain(13, i.termsDelivery);
            plain(14, i.description);
            plain(15, i.qty, '#,##0');
            plain(16, (m.doNo || '').replace(/\//g, '-'));
            plain(17, d(m.doDate), 'd-mmm-yy');
            plain(18, m.warehouse || dof.warehouse || '');
            plain(19, todWeek(i.destCode));
            plain(20, null); plain(21, null, '[$-409]d\\-mmm\\-yy;@'); plain(22, null); // TC number / submit date / remark: after CU
            plain(23, d(S.consequenceDate), 'dd\\-mmm\\-yy');
            plain(24, numOr(S.colX));
            plain(25, efficiency, '0.00%');
            fx(26, `IF(Y${r}="","",100-Y${r}*100)`);
            fx(27, `IF(OR(X${r}="",Y${r}=""),"",O${r}*X${r}*Y${r})`, '#,##0.00');
            fx(28, `IF(OR(AA${r}="",Y${r}="",Y${r}=0),"",AA${r}/Y${r})`, '#,##0.00');
            plain(29, gsm);
            plain(30, numOr(S.colAD));
            // helper columns (not part of the summary layout)
            const B = { color: C.input };
            put(wsShp, r, X0, i.warehouseId, B);
            put(wsShp, r, X0 + 1, i.packages, { ...B, fmt: INT });
            if (m.prodDate) put(wsShp, r, X0 + 2, d(m.prodDate), { ...B, fmt: DATE });
            else f(wsShp, r, X0 + 2, `IF(E${r}="","",E${r}-${q(S_SUM)}!$C$5)`, { fmt: DATE });
            put(wsShp, r, X0 + 3, i.amount, { ...B, fmt: DEC2 });
            put(wsShp, r, X0 + 4, i.currency, B);
            put(wsShp, r, X0 + 5, i.file, B);
            put(wsShp, r, X0 + 6, m.doFile, B);
            put(wsShp, r, X0 + 7, m.doDate ? (m.doDateSrc === 'ocr' ? 'OCR' : 'manual') : '', B);
            put(wsShp, r, X0 + 8, m.remark, { ...B, wrap: true });
            f(wsShp, r, X0 + 9, `IF(OR(P${r}="",Q${r}=""),"Lengkapi DO",IF(Q${r}<E${r},"Cek tgl DO","OK"))`);
        });
        const rN = r1 + Math.max(inv.length, 1) - 1;
        wsShp.views = [{ state: 'frozen', xSplit: 4, ySplit: 1 }];
        wsShp.autoFilter = { from: { row: 1, column: 1 }, to: { row: rN, column: 23 } };
        statusRules(wsShp, `${XC('Status')}${r1}:${XC('Status')}${rN}`);
        const shp = (c) => `${q(S_SHP)}!$${c}$${r1}:$${c}$${rN}`;

        // ------------------------------------------------ Ringkasan
        title(wsSum, `Ringkasan TC – Order ${orderNo}`, `Dibuat ${new Date().toISOString().slice(0, 10)} dengan TC Excel Tool. Sumber: PDF invoice/PO/BOM + OCR tanggal DO dari scan + input manual.`);
        wsSum.getColumn(1).width = 3; wsSum.getColumn(2).width = 46; wsSum.getColumn(3).width = 26; wsSum.getColumn(4).width = 50;
        put(wsSum, 4, 2, 'Input', { bold: true, border: false });
        put(wsSum, 5, 2, 'Production Date = Invoice Date dikurangi (hari)');
        const pd = put(wsSum, 5, 3, model.settings.prodDays, { color: C.input, fill: C.yellow, fmt: '0' });
        pd.note = 'Ubah angka ini untuk menghitung ulang kolom Production Date (kecuali yang diisi manual). Default 14 hari = aturan "tanggal invoice - 2 minggu" dari scan beranotasi.';
        put(wsSum, 5, 4, 'Ubah sel kuning untuk hitung ulang', { italic: true, border: false, color: C.grey });

        put(wsSum, 7, 2, 'Info Order', { bold: true, border: false });
        const infos = [
            ['Order No', orderNo], ['Product No', (po && po.productNo) || first.productId], ['Product Name', (po && po.productName) || first.productName],
            ['Product Description', po && po.productDescription], ['Season', po && po.season], ['Supplier (PO)', po && po.supplierName],
            ['Option No / Development No', po ? `${po.optionNo} / ${po.developmentNo}` : ''], ['Country of Production (PO)', po && po.countryOfProduction],
            ['Country of Origin (PO)', po && po.countryOfOrigin], ['Incoterm (PO)', po && po.incoterm], ['Tanggal PO', po ? d(po.orderDate) : null]
        ];
        infos.forEach((x, k) => {
            put(wsSum, 8 + k, 2, x[0]);
            put(wsSum, 8 + k, 3, x[1] || null, { color: C.input, fmt: x[1] instanceof Date ? DATE : undefined });
        });
        let r = 8 + infos.length + 1;
        put(wsSum, r, 2, 'Angka Kunci (formula)', { bold: true, border: false }); r++;
        const kpi = [
            ['Jumlah invoice', `COUNTA(${shp('D')})`, INT],
            ['Total qty (pcs)', `SUM(${shp('O')})`, INT],
            ['Total qty PO', po && po.totalQty != null ? String(po.totalQty) : '""', INT],
            ['Selisih qty invoice - PO', po && po.totalQty != null ? `C${r + 1}-C${r + 2}` : '""', INT],
            ['Total karton', `SUM(${shp(XC('Karton'))})`, INT],
            ['Total net weight (kg)', `SUM(${shp('F')})`, DEC2],
            ['Total gross weight (kg)', `SUM(${shp('G')})`, DEC2],
            ['Invoice dengan data DO lengkap', `COUNTIF(${shp(XC('Status'))},"OK")`, INT],
            ['Invoice yang masih perlu dilengkapi', `COUNTA(${shp('D')})-C${r + 7}`, INT],
            ['Temuan tingkat Tinggi', `COUNTIF(${q(S_ISS)}!$B:$B,"Tinggi")`, INT],
            ['Temuan tingkat Sedang', `COUNTIF(${q(S_ISS)}!$B:$B,"Sedang")`, INT]
        ];
        kpi.forEach((x, k) => {
            put(wsSum, r + k, 2, x[0]);
            const c = f(wsSum, r + k, 3, x[1], { bold: true, fmt: x[2] });
            if (x[0] === 'Total qty PO' && po && po.totalQty != null) { c.value = po.totalQty; styleCell(c, { bold: true, fmt: INT, color: C.input }); }
        });
        const rK = r;
        wsSum.addConditionalFormatting({ ref: `C${rK + 8}:C${rK + 9}`, rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: ['0'], priority: 1, style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.red } } } }] });
        r = rK + kpi.length + 1;
        put(wsSum, r, 2, 'Legenda: teks biru = dari PDF / input manual; hitam = formula; kuning = wajib diisi / boleh diubah.', { italic: true, border: false, color: C.grey });

        // ------------------------------------------------ Berat Serat
        title(wsFib, 'Estimasi Berat Serat Bersertifikat per Invoice (ESTIMASI)',
            'Rumus: Net Weight invoice x % komposisi di deskripsi invoice. Verifikasi dengan konsumsi kain aktual sebelum dipakai di CU. % (biru) boleh diubah.');
        const mats = [];
        inv.forEach(i => i.composition.forEach(c => { if (!mats.includes(c.material)) mats.push(c.material); }));
        const fh = ['Invoice No', 'Kode', 'Net Weight (kg)'];
        mats.forEach(mm => { fh.push(`% ${mm}`); fh.push(`kg ${mm}`); });
        header(wsFib, 4, fh, [15, 7, 11].concat(mats.flatMap(() => [9, 13])));
        put(wsFib, 5, 1, 'Program (perkiraan)', { italic: true, color: C.grey });
        mats.forEach((mm, k) => put(wsFib, 5, 4 + k * 2, programForLocal(mm, model.settings.recycledProgram), { italic: true, color: C.grey, align: 'center' }));
        inv.forEach((i, k) => {
            const rr = 6 + k;
            f(wsFib, rr, 1, `${q(S_SHP)}!D${r1 + k}`);
            f(wsFib, rr, 2, `${q(S_SHP)}!L${r1 + k}`);
            f(wsFib, rr, 3, `${q(S_SHP)}!F${r1 + k}`, { fmt: DEC2 });
            mats.forEach((mm, j) => {
                const c = i.composition.find(x => x.material === mm);
                const pc = 4 + j * 2;
                put(wsFib, rr, pc, c ? c.pct / 100 : 0, { color: C.input, fmt: '0%' });
                f(wsFib, rr, pc + 1, `$C${rr}*${col(pc)}${rr}`, { fmt: DEC3 });
            });
        });
        const fN = 6 + Math.max(inv.length, 1) - 1;
        put(wsFib, fN + 1, 1, 'Total', { bold: true });
        f(wsFib, fN + 1, 3, `SUM(C6:C${fN})`, { bold: true, fmt: DEC2 });
        mats.forEach((mm, j) => { const kc = 5 + j * 2; f(wsFib, fN + 1, kc, `SUM(${col(kc)}6:${col(kc)}${fN})`, { bold: true, fmt: DEC3 }); });
        wsFib.views = [{ state: 'frozen', xSplit: 1, ySplit: 4 }];

        // ------------------------------------------------ PO vs Invoice
        title(wsPo, 'PO vs Invoice', po ? `PO ${po.orderNo} tanggal ${po.orderDate}. Qty PO dari tabel Time of Delivery (per kelompok TOD).` : 'PO belum di-upload.');
        let pr = 4;
        if (po) {
            header(wsPo, pr, ['TOD (PO)', 'Planning Markets', 'Kode', 'Qty PO', 'Qty Invoice', 'Selisih', 'Status'], [13, 46, 18, 10, 11, 9, 10]);
            const g0 = pr + 1;
            po.todGroups.forEach((g, k) => {
                const rr = g0 + k;
                const codes = g.markets.map(m => m.code);
                put(wsPo, rr, 1, d(g.tod), { color: C.input, fmt: DATE });
                put(wsPo, rr, 2, g.marketsText, { color: C.input, wrap: true });
                put(wsPo, rr, 3, codes.join(', '), { color: C.input });
                put(wsPo, rr, 4, g.qty, { color: C.input, fmt: INT });
                f(wsPo, rr, 5, codes.length ? codes.map(cd => `SUMIF(${shp('L')},"${esc(cd)}",${shp('O')})`).join('+') : '0', { fmt: INT });
                f(wsPo, rr, 6, `E${rr}-D${rr}`, { fmt: INT });
                f(wsPo, rr, 7, `IF(F${rr}=0,"OK","BEDA")`);
            });
            const gN = g0 + Math.max(po.todGroups.length, 1) - 1;
            put(wsPo, gN + 1, 2, 'Total', { bold: true });
            [4, 5, 6].forEach(c => f(wsPo, gN + 1, c, `SUM(${col(c)}${g0}:${col(c)}${gN})`, { bold: true, fmt: INT }));
            statusRules(wsPo, `G${g0}:G${gN}`);

            pr = gN + 4;
            put(wsPo, pr - 1, 1, 'Harga per kode negara (Invoice Average Price PO vs invoice)', { bold: true, border: false });
            header(wsPo, pr, ['Kode', 'Harga PO', 'Currency PO', 'Currency Invoice', 'Harga Rata-rata Invoice', 'Qty Invoice', 'Status']);
            const p0 = pr + 1;
            po.prices.forEach((p, k) => {
                const rr = p0 + k;
                const curInv = [...new Set(inv.filter(i => i.destCode === p.code).map(i => i.currency))].join('/');
                put(wsPo, rr, 1, p.code, { color: C.input });
                put(wsPo, rr, 2, p.price, { color: C.input, fmt: DEC2 });
                put(wsPo, rr, 3, p.currency, { color: C.input });
                put(wsPo, rr, 4, curInv, { color: C.input });
                f(wsPo, rr, 5, `IFERROR(SUMIF(${shp('L')},A${rr},${shp(XC('Amount'))})/SUMIF(${shp('L')},A${rr},${shp('O')}),0)`, { fmt: DEC2 });
                f(wsPo, rr, 6, `SUMIF(${shp('L')},A${rr},${shp('O')})`, { fmt: INT });
                f(wsPo, rr, 7, `IF(F${rr}=0,"-",IF(C${rr}<>D${rr},"BEDA (mata uang)",IF(ABS(B${rr}-E${rr})<0.005,"OK","BEDA")))`);
            });
            statusRules(wsPo, `G${p0}:G${p0 + Math.max(po.prices.length, 1) - 1}`);
            pr = p0 + po.prices.length + 2;
            if (po.details.length) {
                put(wsPo, pr, 1, 'Halaman detail PO per market', { bold: true, border: false });
                header(wsPo, pr + 1, ['Kode', 'Qty PO (detail)', 'Qty Invoice', 'Selisih', 'Buyer (PO detail)']);
                po.details.forEach((dt, k) => {
                    const rr = pr + 2 + k;
                    put(wsPo, rr, 1, dt.code, { color: C.input });
                    put(wsPo, rr, 2, dt.qty, { color: C.input, fmt: INT });
                    f(wsPo, rr, 3, `SUMIF(${shp('L')},A${rr},${shp('O')})`, { fmt: INT });
                    f(wsPo, rr, 4, `C${rr}-B${rr}`, { fmt: INT });
                    put(wsPo, rr, 5, dt.buyer, { color: C.input, wrap: true });
                });
            }
        }

        // ------------------------------------------------ BOM
        title(wsBom, 'Material dari BOM', model.bom ? 'Dibaca dari PDF BOM (Supplier Quote BOM). Hanya bagian Material & Trims.' : 'BOM belum di-upload.');
        header(wsBom, 4, ['Section', 'Position', 'Placement', 'Master Component', 'Component', 'Supplier Article', 'Composition', 'Weight / GSM', 'Supplier', 'Program (perkiraan)'], [10, 9, 13, 14, 12, 13, 40, 11, 36, 14]);
        ((model.bom && model.bom.materials) || []).forEach((m, k) => {
            const rr = 5 + k;
            [m.section, m.position, m.placement, m.masterComponent, m.component, m.supplierArticle, m.composition, m.weight, m.supplier]
                .forEach((v, j) => put(wsBom, rr, j + 1, v, { color: C.input, wrap: j === 6 || j === 8 }));
            const progs = [...new Set(m.compositionParsed.map(c => programForLocal(c.material, model.settings.recycledProgram)).filter(Boolean))];
            put(wsBom, rr, 10, progs.join(', ') || '-', { color: C.grey });
        });

        // ------------------------------------------------ Temuan
        title(wsIss, 'Temuan / Validasi', 'Hasil cek otomatis antar dokumen saat Excel dibuat. Perbaiki di tool lalu generate ulang.');
        header(wsIss, 4, ['No', 'Tingkat', 'Area', 'Referensi', 'Temuan', 'Saran'], [5, 9, 14, 30, 70, 40]);
        model.issues.forEach((x, k) => {
            const rr = 5 + k;
            put(wsIss, rr, 1, k + 1);
            put(wsIss, rr, 2, x.level);
            put(wsIss, rr, 3, x.area);
            put(wsIss, rr, 4, x.ref, { wrap: true });
            put(wsIss, rr, 5, x.message, { wrap: true });
            put(wsIss, rr, 6, x.action, { wrap: true });
        });
        if (!model.issues.length) put(wsIss, 5, 5, 'Tidak ada temuan.');
        const iN = 5 + Math.max(model.issues.length, 1) - 1;
        wsIss.addConditionalFormatting({
            ref: `B5:B${iN}`, rules: [
                { type: 'cellIs', operator: 'equal', formulae: ['"Tinggi"'], priority: 1, style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.red } } } },
                { type: 'cellIs', operator: 'equal', formulae: ['"Sedang"'], priority: 2, style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.amber } } } }
            ]
        });
        wsIss.views = [{ state: 'frozen', ySplit: 4 }];
        wsIss.autoFilter = { from: { row: 4, column: 1 }, to: { row: iN, column: 6 } };

        // ------------------------------------------------ Dokumen
        title(wsDoc, 'Daftar Dokumen yang Di-upload', 'Jenis dideteksi otomatis dari isi PDF (atau nama file untuk DO hasil scan).');
        header(wsDoc, 4, ['No', 'File', 'Jenis', 'Halaman', 'Status', 'Keterangan'], [5, 60, 14, 9, 12, 60]);
        model.files.forEach((x, k) => {
            const rr = 5 + k;
            put(wsDoc, rr, 1, k + 1);
            put(wsDoc, rr, 2, x.name, { wrap: true });
            put(wsDoc, rr, 3, x.kindLabel);
            put(wsDoc, rr, 4, x.pages, { fmt: INT });
            put(wsDoc, rr, 5, x.error ? 'Gagal' : 'OK');
            put(wsDoc, rr, 6, x.error || x.info || '', { wrap: true });
        });

        [wsSum, wsShp, wsFib, wsPo, wsBom, wsIss, wsDoc].forEach(ws => { ws.properties.defaultRowHeight = 15; });
        return wb;
    }

    function programForLocal(material, recycledProgram) {
        if (/ORGANIC/.test(material)) return 'OCS';
        if (/RECYCLED/.test(material)) return recycledProgram || 'RCS';
        return '';
    }

    return { buildWorkbook };
});
