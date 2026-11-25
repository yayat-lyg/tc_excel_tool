/*
 * TC Excel Tool - cross-document checks and DO auto-assignment.
 * Works in the browser (window.TcValidate) and in Node (module.exports).
 */
(function (root, factory) {
    if (typeof module === 'object' && module.exports) module.exports = factory();
    else root.TcValidate = factory();
})(typeof self !== 'undefined' ? self : this, function () {
    'use strict';

    // DO candidates for an invoice = DO files whose file name lists the invoice country code.
    function doCandidates(invoice, doFiles) {
        return doFiles.filter(d => d.codes.includes(invoice.destCode));
    }

    // OCR look-alikes folded to one character, so "LYGM1S127v26" compares equal to "LYGM15127Y26"
    function ocrNorm(s) {
        return String(s || '').toUpperCase().replace(/¥/g, 'Y').replace(/[^A-Z0-9]/g, '')
            .replace(/[ODQ]/g, '0').replace(/[IL]/g, '1').replace(/Z/g, '2').replace(/S/g, '5')
            .replace(/B/g, '8').replace(/G/g, '6').replace(/V/g, 'Y');
    }
    function editDistance(a, b) {
        const d = Array.from({ length: a.length + 1 }, (_, i) => [i]);
        for (let j = 1; j <= b.length; j++) d[0][j] = j;
        for (let i = 1; i <= a.length; i++) {
            for (let j = 1; j <= b.length; j++) d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
        }
        return d[a.length][b.length];
    }
    // smallest edit distance between `needle` and any same-length window of `hay`
    function nearest(hay, needle) {
        if (!needle || hay.length < needle.length) return Infinity;
        let best = Infinity;
        for (let i = 0; i + needle.length <= hay.length && best > 0; i++) best = Math.min(best, editDistance(hay.substr(i, needle.length), needle));
        return best;
    }

    /*
     * Does a DO's highlighted rows (OCR text, see app.js highlightedRows) mention this invoice?
     * 'invoice' = invoice no found (1 OCR error allowed), 'warehouse' = H&M warehouse id found (exact), '' = no.
     * (The country code is not used: one DO can hold rows of several invoices of the same country.)
     */
    function doRowsMention(doFile, invoice) {
        const rows = (doFile.ocrRows || []).map(ocrNorm);
        const inv = ocrNorm(invoice.invoiceNo), wh = ocrNorm(invoice.warehouseId);
        if (inv.length >= 8 && rows.some(n => nearest(n, inv) <= 1)) return 'invoice';
        if (wh.length >= 5 && rows.some(n => n.includes(wh))) return 'warehouse';
        return '';
    }

    // Auto-assign when exactly one DO file matches the country code, or, when several (or none) do,
    // exactly one of them has a highlighted row with the invoice no (else: with the warehouse id).
    function autoAssignDo(invoices, doFiles) {
        invoices.forEach(inv => {
            const m = inv.manual;
            if (m.doFile) return; // user already chose
            let c = doCandidates(inv, doFiles);
            if (c.length !== 1) {
                const pool = c.length ? c : doFiles;
                const byInv = pool.filter(d => doRowsMention(d, inv) === 'invoice');
                const byAny = pool.filter(d => doRowsMention(d, inv));
                if (byInv.length === 1) c = byInv;
                else if (!byInv.length && byAny.length === 1) c = byAny;
            }
            if (c.length === 1) {
                m.doFile = c[0].file;
                if (!m.doNo) m.doNo = c[0].doNo;
            }
        });
    }

    function addDays(iso, days) {
        if (!/^\d{4}-\d{2}-\d{2}$/.test(iso || '')) return '';
        const d = new Date(iso + 'T00:00:00Z');
        d.setUTCDate(d.getUTCDate() + days);
        return d.toISOString().slice(0, 10);
    }

    function buildIssues(model) {
        const out = [];
        const add = (level, area, ref, message, action) => out.push({ level, area, ref: ref || '', message, action: action || '' });
        const inv = model.invoices;
        const po = model.po;

        model.files.filter(f => f.error).forEach(f => add('Tinggi', 'Dokumen', f.name, 'Gagal dibaca: ' + f.error, 'Cek file PDF'));
        if (!inv.length) add('Tinggi', 'Dokumen', '', 'Belum ada invoice yang terbaca', 'Upload PDF invoice H&M');
        if (!po) add('Sedang', 'Dokumen', '', 'PO H&M belum di-upload: cek qty & harga per negara tidak bisa dilakukan', 'Upload PDF Purchase Order');
        if (!model.bom) add('Info', 'Dokumen', '', 'BOM belum di-upload: GSM & supplier kain tidak terisi', 'Upload PDF BOM (opsional)');

        // invoice-level parser warnings
        inv.forEach(i => (i.warnings || []).forEach(w => add(/kosong/.test(w) ? 'Sedang' : 'Tinggi', 'Invoice', i.invoiceNo || i.file, w)));

        // duplicate invoice numbers
        const seen = {};
        inv.forEach(i => { seen[i.invoiceNo] = (seen[i.invoiceNo] || 0) + 1; });
        Object.keys(seen).filter(k => k && seen[k] > 1).forEach(k => add('Tinggi', 'Invoice', k, `Invoice ${k} di-upload ${seen[k]}x`, 'Hapus file duplikat'));

        // order number consistency
        const orders = [...new Set(inv.map(i => i.orderNo).filter(Boolean))];
        if (orders.length > 1) add('Tinggi', 'Invoice', '', 'Invoice berisi lebih dari 1 order: ' + orders.join(', '), 'Satu Excel = satu order; pisahkan file');
        if (po && orders.length && !orders.includes(po.orderNo)) add('Tinggi', 'PO', po.orderNo, `Order PO (${po.orderNo}) beda dengan order di invoice (${orders.join(', ')})`);
        const prods = [...new Set(inv.map(i => i.productId).filter(Boolean))];
        if (po && prods.length && po.productNo && !prods.includes(po.productNo)) add('Tinggi', 'PO', po.productNo, `Product No PO (${po.productNo}) beda dengan invoice (${prods.join(', ')})`);

        // composition consistency between invoices / BOM shell
        const comps = [...new Set(inv.map(i => i.composition.map(c => c.pct + '% ' + c.material).join(', ')))];
        if (comps.length > 1) add('Sedang', 'Invoice', '', 'Komposisi berbeda antar invoice: ' + comps.join(' || '));

        // country of origin vs manufacture
        if (inv.some(i => i.countryOfOrigin && i.countryOfManufacture && i.countryOfOrigin.toLowerCase() !== i.countryOfManufacture.toLowerCase())) {
            const s = inv.find(i => i.countryOfOrigin && i.countryOfManufacture);
            add('Info', 'Invoice', '', `Country of Origin of Goods = "${s.countryOfOrigin}" sedangkan Country of Manufacture = "${s.countryOfManufacture}"`,
                'Pastikan field negara di form CU memakai negara produksi');
        }

        if (po) {
            // incoterm vs PO
            inv.forEach(i => {
                if (po.incoterm && i.termsDelivery && !i.termsDelivery.toUpperCase().includes(po.incoterm))
                    add('Sedang', 'Invoice vs PO', i.invoiceNo, `Terms of Delivery invoice "${i.termsDelivery}" beda dengan PO (${po.incoterm})`, 'Konfirmasi ke shipping');
            });
            // country codes not in PO
            const poCodes = new Set(po.todGroups.flatMap(g => g.markets.map(m => m.code)));
            inv.forEach(i => { if (i.destCode && poCodes.size && !poCodes.has(i.destCode)) add('Tinggi', 'Invoice vs PO', i.invoiceNo, `Kode negara ${i.destCode} tidak ada di PO`); });
            // qty per TOD group
            po.todGroups.forEach(g => {
                const codes = g.markets.map(m => m.code);
                const q = inv.filter(i => codes.includes(i.destCode)).reduce((s, i) => s + (i.qty || 0), 0);
                if (g.qty != null && q !== g.qty) add('Sedang', 'Invoice vs PO', codes.join(','), `TOD ${g.tod} (${codes.join(', ')}): qty PO ${g.qty} vs invoice ${q}`, 'Cek invoice yang belum di-upload / short-ship');
            });
            // price per code
            po.prices.forEach(p => {
                const rows = inv.filter(i => i.destCode === p.code);
                if (!rows.length) return;
                const cur = [...new Set(rows.map(i => i.currency))];
                const q = rows.reduce((s, i) => s + (i.qty || 0), 0);
                const a = rows.reduce((s, i) => s + (i.amount || 0), 0);
                if (!cur.includes(p.currency)) add('Info', 'Invoice vs PO', p.code, `Harga PO ${p.code} dalam ${p.currency} ${p.price}, invoice dalam ${cur.join('/')} ${q ? (a / q).toFixed(2) : ''}`);
                else if (q && Math.abs(a / q - p.price) > 0.005) add('Info', 'Invoice vs PO', p.code, `Harga PO ${p.code} ${p.price} vs rata-rata invoice ${(a / q).toFixed(2)}`);
            });
        }

        // DO
        model.doFiles.forEach(d => (d.warnings || []).forEach(w => add('Sedang', 'DO', d.file, w, 'Cek nama file / scan DO')));
        model.doFiles.forEach(d => {
            const used = inv.some(i => i.manual.doFile === d.file);
            if (!used) add('Info', 'DO', d.file, 'File DO tidak dipakai di invoice mana pun');
        });
        inv.forEach(i => {
            const m = i.manual;
            const c = doCandidates(i, model.doFiles);
            if (!m.doFile && c.length > 1) add('Tinggi', 'DO', i.invoiceNo, `Ada ${c.length} file DO untuk kode ${i.destCode} (${c.map(x => x.doNo).join(', ')}): pilih manual`);
            if (!m.doNo) add('Tinggi', 'DO', i.invoiceNo, 'Nomor DO belum diisi', 'Isi di tabel invoice');
            if (!m.doDate) add('Tinggi', 'DO', i.invoiceNo, 'Tanggal DO belum diisi (pilih file DO, atau tanggal tidak terbaca dari scan)', 'Isi dari scan DO');
            if (m.doDate && i.invoiceDate && m.doDate < i.invoiceDate) add('Sedang', 'DO', i.invoiceNo, `Tanggal DO ${m.doDate} lebih awal dari tanggal invoice ${i.invoiceDate}`);
            if (m.doFile && i.destCode) {
                const d = model.doFiles.find(x => x.file === m.doFile);
                if (d && d.codes.length && !d.codes.includes(i.destCode)) add('Sedang', 'DO', i.invoiceNo, `File DO ${d.file} tidak mencantumkan kode ${i.destCode}`);
            }
        });
        return out;
    }

    return { doCandidates, doRowsMention, autoAssignDo, buildIssues, addDays };
});
