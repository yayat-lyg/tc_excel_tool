/*
 * TC Excel Tool - UI.
 * Flow: upload PDFs -> pdf.js text extraction -> TcParsers -> manual DO input -> TcValidate -> TcExcel -> download.
 * pdf.worker.min.js is loaded as a normal <script>, so pdf.js runs in the main thread (works from file://).
 */
(function () {
    'use strict';

    const KIND_LABEL = { invoice: 'Invoice', po: 'Purchase Order', bom: 'BOM', do: 'DO (scan)', scan: 'Lampiran scan', other: 'Lainnya' };
    const STORE_KEY = 'tcExcelTool.v1';

    const state = {
        files: [],            // {id, name, pages(items[][]), pageCount, kind, error, parsed, info}
        manual: {},           // invoiceNo -> {doFile, doNo, doDate, prodDate, remark}
        // consequenceDate / colX / colY / colAD: order-level values for Shipment TC columns W, X, Y, AD
        settings: { prodDays: 14, recycledProgram: 'RCS', consequenceDate: '', colX: '', colY: '', colAD: '' }
    };
    let seq = 0;

    const $ = id => document.getElementById(id);
    const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    const fmt = (v, dec) => (v == null || v === '' ? '' : Number(v).toLocaleString('en-US', { minimumFractionDigits: dec || 0, maximumFractionDigits: dec || 0 }));

    // ------------------------------------------------------------- storage (convenience only)
    function loadStore() {
        try {
            const s = JSON.parse(localStorage.getItem(STORE_KEY) || '{}');
            if (s.manual) state.manual = s.manual;
            if (s.settings) Object.assign(state.settings, s.settings);
        } catch (e) { /* storage unavailable */ }
    }
    function saveStore() {
        try { localStorage.setItem(STORE_KEY, JSON.stringify({ manual: state.manual, settings: state.settings })); } catch (e) { /* ignore */ }
    }

    // ------------------------------------------------------------- PDF reading
    if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdfjs/pdf.worker.min.js';

    async function readPdf(bytes) {
        const doc = await pdfjsLib.getDocument({ data: bytes.slice() }).promise; // pdf.js detaches the buffer it gets
        const pages = [];
        for (let p = 1; p <= doc.numPages; p++) {
            const tc = await (await doc.getPage(p)).getTextContent();
            pages.push(tc.items.map(i => ({ str: i.str, x: i.transform[4], y: i.transform[5], w: i.width })));
        }
        return pages;
    }

    // ------------------------------------------------------------- OCR of scanned DO (date in the header)
    const DO_HEADER = [{ left: 0.03, top: 0.06, width: 0.97, height: 0.18 }]; // page fraction: DELIVERY NO ... DATE block

    async function ocrDo(f) {
        const doc = await pdfjsLib.getDocument({ data: f.bytes.slice() }).promise;
        const page = await doc.getPage(1);
        const yearHint = f.parsed && f.parsed.year;
        let res = null;
        for (const scale of [2.5, 3.5]) { // small/grey print (e.g. FCL DO) needs the higher resolution
            const vp = page.getViewport({ scale });
            const c = document.createElement('canvas');
            c.width = Math.round(vp.width); c.height = Math.round(vp.height);
            await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
            const [txt] = await TcOcr.recognize(c, DO_HEADER);
            res = TcParsers.parseDoOcr(txt, yearHint);
            if (res.date && res.checked && res.warehouse) break;
        }
        try { res.rows = await highlightedRows(page); } catch (e) { res.rows = []; }
        return res;
    }

    /*
     * Staff mark the rows of this order on the DO with a yellow highlighter. Find those bands, whiten the
     * highlighter, and OCR each band as one text line. Used to pick the DO when 2 DO files have the same country code.
     */
    async function highlightedRows(page) {
        const SC = 3;
        const vp = page.getViewport({ scale: SC });
        const c = document.createElement('canvas');
        c.width = Math.round(vp.width); c.height = Math.round(vp.height);
        const ctx = c.getContext('2d', { willReadFrequently: true });
        await page.render({ canvasContext: ctx, viewport: vp }).promise;
        const W = c.width, H = c.height;
        const im = ctx.getImageData(0, 0, W, H), d = im.data;
        const bands = [];
        let start = -1;
        for (let y = 0; y <= H; y++) {
            let n = 0;
            if (y < H) for (let x = 0; x < W; x += 2) { const i = (y * W + x) * 4; if (d[i] > 150 && Math.min(d[i], d[i + 1]) - d[i + 2] > 40) n++; }
            const on = n > W / 2 * 0.25; // a quarter of the row is yellow
            if (on && start < 0) start = y;
            if (!on && start >= 0) {
                if (y - start >= 3 * SC) { // two adjacent highlighted rows form one band: split by the usual row height
                    const k = Math.max(1, Math.round((y - start) / (0.0135 * H)));
                    for (let j = 0; j < k; j++) bands.push([start + (y - start) * j / k, start + (y - start) * (j + 1) / k]);
                }
                start = -1;
            }
        }
        if (!bands.length || bands.length > 40) return [];
        for (let i = 0; i < d.length; i += 4) { // binarise: highlighter (light yellow) -> white
            const v = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2] < 165 ? 0 : 255;
            d[i] = d[i + 1] = d[i + 2] = v;
        }
        ctx.putImageData(im, 0, 0);
        // two windows per band (tight / a bit taller): OCR quality differs per scan, matching uses both texts
        const rects = [];
        bands.forEach(([a, b]) => {
            rects.push({ left: 0.02, top: Math.max(0, a - 2) / H, width: 0.96, height: (b - a + 4) / H });
            rects.push({ left: 0.02, top: Math.max(0, (a + b) / 2 - 0.009 * H) / H, width: 0.96, height: 0.018 });
        });
        return (await TcOcr.recognize(c, rects, 7)).map(t => t.trim()).filter(Boolean);
    }

    function mergeOcr(f) {
        const p = f.parsed, o = f.ocr;
        if (!p || !o) return;
        p.warnings = p.warnings.slice();
        if (o.error) { p.warnings.push('OCR gagal: ' + o.error); return; }
        p.ocrDate = o.date;
        p.ocrChecked = o.checked;
        p.ocrRows = o.rows || [];
        p.warehouse = o.warehouse || '';
        o.warnings.forEach(w => p.warnings.push(w));
        if (o.seq && o.seq !== p.seq.padStart(6, '0')) p.warnings.push(`No DO di scan (${o.seq}) beda dengan nama file (${p.seq})`);
        if (p.year.length !== 4 && o.year) {
            p.doNoFile = p.doNo;
            p.doNo = `${p.seq}/LYG-EXP/${p.roman}/${o.year}`;
            p.warnings = p.warnings.map(w => /^Tahun di nama file/.test(w) ? `Tahun di nama file DO "${p.year}" salah ketik; dipakai tahun dari scan: ${o.year}` : w);
        }
        f.info = `${p.doNo} · kode ${p.codes.join(', ') || '-'} · tgl DO ${o.date || '?'} · gudang ${o.warehouse || '?'} (OCR)`;
    }

    let ocrRunning = false;
    async function runOcr() {
        if (ocrRunning) return;
        ocrRunning = true;
        const prog = $('progress');
        try {
            let todo;
            while ((todo = state.files.filter(f => f.kind === 'do' && f.parsed && f.bytes && !f.ocr)).length) {
                const f = todo[0];
                prog.textContent = `OCR tanggal DO: ${f.name} (sisa ${todo.length})`;
                await new Promise(r => setTimeout(r, 30)); // let the progress text paint (OCR blocks the page)
                try { f.ocr = await ocrDo(f); } catch (e) { f.ocr = { error: e.message }; }
                mergeOcr(f);
                refresh();
            }
        } finally {
            prog.textContent = '';
            ocrRunning = false;
        }
    }

    function parseFile(f) {
        f.error = ''; f.parsed = null; f.info = '';
        try {
            if (f.kind === 'invoice') {
                f.parsed = TcParsers.parseInvoice(f.pages);
                f.info = `${f.parsed.invoiceNo} · ${f.parsed.destCode} · ${fmt(f.parsed.qty)} pcs`;
            } else if (f.kind === 'po') {
                f.parsed = TcParsers.parsePo(f.pages);
                f.info = `${f.parsed.orderNo} · ${f.parsed.productName} · ${fmt(f.parsed.totalQty)} pcs`;
            } else if (f.kind === 'bom') {
                f.parsed = TcParsers.parseBom(f.pages);
                f.info = `${f.parsed.materials.length} baris material/trim`;
            } else if (f.kind === 'do') {
                f.parsed = TcParsers.parseDoFileName(f.name);
                if (!f.parsed) f.error = 'Nama file tidak sesuai format "000721 - LYG - EXP - VIII - 2026 - LD, OT, OU.pdf"';
                else {
                    f.info = `${f.parsed.doNo} · kode ${f.parsed.codes.join(', ') || '-'}` + (f.ocr ? '' : ' · OCR tanggal DO menunggu...');
                    mergeOcr(f);
                }
            } else if (f.kind === 'scan') {
                f.info = 'Hasil scan, tidak dibaca (hanya dicatat sebagai lampiran)';
            }
        } catch (e) {
            f.error = 'Gagal parse: ' + e.message;
        }
    }

    async function addFiles(fileList) {
        const list = [...fileList];
        const prog = $('progress');
        for (let k = 0; k < list.length; k++) {
            const file = list[k];
            prog.textContent = `Membaca ${k + 1}/${list.length}: ${file.name}`;
            if (!/\.pdf$/i.test(file.name)) {
                state.files.push({ id: ++seq, name: file.name, pages: [], pageCount: 0, kind: 'other', error: 'Bukan file PDF' });
                continue;
            }
            if (state.files.some(f => f.name === file.name && f.size === file.size)) continue; // same file twice
            const f = { id: ++seq, name: file.name, size: file.size, pages: [], pageCount: 0, kind: 'other' };
            try {
                f.bytes = new Uint8Array(await file.arrayBuffer());
                f.pages = await readPdf(f.bytes);
                f.pageCount = f.pages.length;
                f.kind = TcParsers.detect(file.name, f.pages);
                parseFile(f);
            } catch (e) {
                f.error = 'PDF tidak bisa dibuka: ' + e.message;
            }
            state.files.push(f);
        }
        prog.textContent = '';
        refresh();
        runOcr();
    }

    // ------------------------------------------------------------- model
    function buildModel() {
        const files = state.files;
        const doFiles = files.filter(f => f.kind === 'do' && f.parsed).map(f => ({ ...f.parsed, file: f.name }));
        const invoices = files.filter(f => f.kind === 'invoice' && f.parsed && f.parsed.invoiceNo).map(f => {
            const key = f.parsed.invoiceNo;
            if (!state.manual[key]) state.manual[key] = {};
            return { ...f.parsed, file: f.name, manual: state.manual[key] };
        }).sort((a, b) => (a.invoiceDate + a.invoiceNo).localeCompare(b.invoiceDate + b.invoiceNo));
        // drop a remembered DO file that is no longer uploaded
        invoices.forEach(i => { if (i.manual.doFile && !doFiles.some(d => d.file === i.manual.doFile)) i.manual.doFile = ''; });
        TcValidate.autoAssignDo(invoices, doFiles);
        // DO date from OCR of the chosen DO file, unless typed by the user (doDateSrc 'manual')
        invoices.forEach(i => {
            const m = i.manual;
            const d = doFiles.find(x => x.file === m.doFile);
            if (d && d.doNoFile && m.doNo === d.doNoFile) m.doNo = d.doNo; // year typo in file name fixed from scan
            if (m.doDateSrc === 'ocr' && (!d || d.ocrDate !== m.doDate)) { m.doDate = ''; m.doDateSrc = ''; }
            if (!m.doDate && d && d.ocrDate) { m.doDate = d.ocrDate; m.doDateSrc = 'ocr'; }
        });
        const poF = files.find(f => f.kind === 'po' && f.parsed);
        const bomF = files.find(f => f.kind === 'bom' && f.parsed);
        const model = {
            invoices, doFiles,
            po: poF ? poF.parsed : null,
            bom: bomF ? bomF.parsed : null,
            files: files.map(f => ({ name: f.name, kind: f.kind, kindLabel: KIND_LABEL[f.kind] || f.kind, pages: f.pageCount, error: f.error, info: f.info })),
            settings: { ...state.settings }
        };
        if (files.filter(f => f.kind === 'po').length > 1) model.files.push({ name: '(catatan)', kindLabel: '-', error: '', info: 'Lebih dari 1 PO di-upload; yang dipakai file PO pertama' });
        model.issues = TcValidate.buildIssues(model);
        return model;
    }

    // ------------------------------------------------------------- render
    function renderFiles() {
        const rows = state.files.map(f => {
            const opts = Object.keys(KIND_LABEL).map(k => `<option value="${k}"${k === f.kind ? ' selected' : ''}>${KIND_LABEL[k]}</option>`).join('');
            const warn = (f.parsed && f.parsed.warnings && f.parsed.warnings.length)
                ? `<ul class="warn-list">${f.parsed.warnings.map(w => `<li>${esc(w)}</li>`).join('')}</ul>` : '';
            const status = f.error ? `<span class="badge b-err">Gagal</span>` : (warn ? `<span class="badge b-warn">Cek</span>` : `<span class="badge b-ok">OK</span>`);
            return `<tr>
                <td class="wrap">${esc(f.name)}</td>
                <td><select data-kind="${f.id}">${opts}</select></td>
                <td class="num">${f.pageCount || ''}</td>
                <td>${status}</td>
                <td class="wrap">${esc(f.error || f.info || '')}${warn}</td>
                <td><button class="btn ghost" data-del="${f.id}" type="button">Hapus</button></td>
            </tr>`;
        }).join('');
        $('tblFiles').innerHTML = `<thead><tr><th>File</th><th>Jenis</th><th class="num">Hal.</th><th>Status</th><th>Keterangan</th><th></th></tr></thead><tbody>${rows}</tbody>`;
    }

    function dateSrc(m, model) {
        if (m.doDateSrc === 'manual') return '<div class="muted">manual</div>';
        if (m.doDateSrc === 'ocr') {
            const d = model.doFiles.find(x => x.file === m.doFile);
            return d && d.ocrChecked ? '<div class="muted">OCR</div>' : '<div class="ocr-check">OCR – cek scan</div>';
        }
        const f = state.files.find(x => x.name === m.doFile);
        if (f && !f.ocr) return '<div class="muted">OCR berjalan...</div>';
        return f ? '<div class="ocr-check">tidak terbaca, isi manual</div>' : '';
    }

    function renderInvoices(model) {
        const opts = (inv) => {
            const cands = TcValidate.doCandidates(inv, model.doFiles).map(d => d.file);
            const all = model.doFiles.map(d => d.file);
            const ordered = cands.concat(all.filter(x => !cands.includes(x)));
            return `<option value="">– pilih –</option>` + ordered.map(fn => {
                const d = model.doFiles.find(x => x.file === fn);
                const hit = TcValidate.doRowsMention(d, inv);
                const tag = (cands.includes(fn) ? '' : ' (kode lain)') + (hit === 'invoice' ? ' ✓ invoice ada di DO' : hit === 'warehouse' ? ' ✓ warehouse ada di DO' : '');
                return `<option value="${esc(fn)}"${inv.manual.doFile === fn ? ' selected' : ''}>${esc(d.doNo + ' [' + d.codes.join(',') + ']' + tag)}</option>`;
            }).join('');
        };
        const rows = model.invoices.map(i => {
            const m = i.manual;
            const prodDefault = TcValidate.addDays(i.invoiceDate, -Number(state.settings.prodDays || 0));
            const k = esc(i.invoiceNo);
            const whAuto = (model.doFiles.find(x => x.file === m.doFile) || {}).warehouse || '';
            return `<tr>
                <td>${k}</td>
                <td>${esc(i.invoiceDate)}</td>
                <td>${esc(i.destCode)}</td>
                <td>${esc(i.warehouseId)}</td>
                <td class="short" title="${esc(i.consignee)}">${esc(i.consignee)}</td>
                <td class="num">${fmt(i.packages)}</td>
                <td class="num">${fmt(i.qty)}</td>
                <td class="num">${fmt(i.netWeight, 2)}</td>
                <td class="num">${fmt(i.grossWeight, 2)}</td>
                <td>${esc(i.termsDelivery)}</td>
                <td><select data-inv="${k}" data-f="doFile">${opts(i)}</select></td>
                <td><input type="text" data-inv="${k}" data-f="doNo" value="${esc(m.doNo || '')}" class="${m.doNo ? '' : 'need'}" size="22"></td>
                <td><input type="text" data-inv="${k}" data-f="warehouse" data-auto="${esc(whAuto)}" value="${esc(m.warehouse || whAuto)}" size="20" title="Address warehouse (kolom R). Default dari header scan DO."><div class="muted">${m.warehouse ? 'manual' : (whAuto ? 'dari DO' : '')}</div></td>
                <td><input type="date" data-inv="${k}" data-f="doDate" value="${esc(m.doDate || '')}" class="${m.doDate ? '' : 'need'}">${dateSrc(m, model)}</td>
                <td><input type="date" data-inv="${k}" data-f="prodDate" data-auto="${esc(prodDefault)}" value="${esc(m.prodDate || prodDefault)}" title="Otomatis = Invoice Date − ${esc(state.settings.prodDays)} hari (${esc(prodDefault)})"><div class="muted">${m.prodDate ? 'manual' : 'otomatis (−' + esc(state.settings.prodDays) + ' hari)'}</div></td>
                <td><input type="text" data-inv="${k}" data-f="remark" value="${esc(m.remark || '')}" size="18"></td>
            </tr>`;
        }).join('');
        $('tblInv').innerHTML = `<thead><tr>
            <th>Invoice No</th><th>Tgl Invoice</th><th>Kode</th><th>Warehouse</th><th>Consignee</th>
            <th class="num">Ktn</th><th class="num">Qty</th><th class="num">NW kg</th><th class="num">GW kg</th><th>ToD</th>
            <th>File DO</th><th>No DO</th><th>Gudang</th><th>Tgl DO</th><th>Production Date</th><th>Catatan</th></tr></thead><tbody>${rows}</tbody>`;
        const tq = model.invoices.reduce((s, i) => s + (i.qty || 0), 0);
        const tn = model.invoices.reduce((s, i) => s + (i.netWeight || 0), 0);
        const tg = model.invoices.reduce((s, i) => s + (i.grossWeight || 0), 0);
        const tc = model.invoices.reduce((s, i) => s + (i.packages || 0), 0);
        const poq = model.po && model.po.totalQty != null ? ` · Qty PO ${fmt(model.po.totalQty)} (selisih ${fmt(tq - model.po.totalQty)})` : '';
        $('invTotals').textContent = `${model.invoices.length} invoice · ${fmt(tc)} karton · ${fmt(tq)} pcs · NW ${fmt(tn, 2)} kg · GW ${fmt(tg, 2)} kg${poq}`;
    }

    function renderIssues(model) {
        const lv = { Tinggi: 'b-err', Sedang: 'b-warn', Info: 'b-invoice' };
        const order = { Tinggi: 0, Sedang: 1, Info: 2 };
        const list = model.issues.slice().sort((a, b) => order[a.level] - order[b.level]);
        $('issues').innerHTML = list.length ? list.map(x => `<div class="issue">
            <div class="lvl"><span class="badge ${lv[x.level] || ''}">${esc(x.level)}</span></div>
            <div><div>${esc(x.message)}</div><div class="ref">${esc(x.area)}${x.ref ? ' · ' + esc(x.ref) : ''}${x.action ? ' · Saran: ' + esc(x.action) : ''}</div></div>
        </div>`).join('') : '<p class="muted">Tidak ada temuan.</p>';
    }

    function refresh() {
        const model = buildModel();
        const has = state.files.length > 0;
        $('secFiles').hidden = !has;
        $('secInv').hidden = !model.invoices.length;
        $('secIssues').hidden = !has;
        $('secGen').hidden = !model.invoices.length;
        renderFiles();
        renderInvoices(model);
        renderIssues(model);
        const high = model.issues.filter(x => x.level === 'Tinggi').length;
        $('genMsg').textContent = high ? `Masih ada ${high} temuan Tinggi (Excel tetap bisa dibuat, temuan ikut dicatat di sheet Temuan).` : '';
        saveStore();
    }

    // ------------------------------------------------------------- events
    function onManualChange(e) {
        const el = e.target;
        const key = el.dataset.inv, field = el.dataset.f;
        if (!key || !field) return;
        const m = state.manual[key] || (state.manual[key] = {});
        if (field === 'doFile') {
            const prevFile = m.doFile;
            const prevDo = (state.files.find(f => f.name === prevFile && f.parsed) || {}).parsed;
            m.doFile = el.value;
            const d = (state.files.find(f => f.name === el.value && f.parsed) || {}).parsed;
            if (d && (!m.doNo || (prevDo && m.doNo === prevDo.doNo))) m.doNo = d.doNo;
        } else {
            m[field] = el.value.trim();
            // production date equal to (or cleared back to) the automatic value stays automatic (Excel formula)
            if (field === 'prodDate' && m.prodDate === el.dataset.auto) m.prodDate = '';
            if (field === 'warehouse' && m.warehouse === el.dataset.auto) m.warehouse = ''; // same as DO scan = follow the DO
            if (field === 'doDate') {
                m.doDateSrc = m.doDate ? 'manual' : ''; // cleared -> OCR date (if any) comes back
                // same DO number -> same DO date (only rows still empty or filled by OCR)
                if (m.doNo && m.doDate) {
                    Object.values(state.manual).forEach(o => {
                        if (o !== m && o.doNo === m.doNo && (!o.doDate || o.doDateSrc === 'ocr')) { o.doDate = m.doDate; o.doDateSrc = 'manual'; }
                    });
                }
            }
        }
        refresh();
    }

    async function generate() {
        const btn = $('btnGenerate');
        btn.disabled = true;
        $('genMsg').textContent = 'Membuat Excel...';
        try {
            const model = buildModel();
            const wb = TcExcel.buildWorkbook(ExcelJS, model);
            const buf = await wb.xlsx.writeBuffer();
            const order = ((model.po && model.po.orderNo) || (model.invoices[0] || {}).orderNo || 'order').replace(/[^\w\-]/g, '_');
            const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
            const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = `TC_${order}_${today}.xlsx`;
            document.body.appendChild(a);
            a.click();
            setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
            $('genMsg').textContent = `Excel dibuat: ${a.download}`;
        } catch (e) {
            $('genMsg').textContent = 'Gagal membuat Excel: ' + e.message;
        } finally {
            btn.disabled = false;
        }
    }

    function init() {
        loadStore();
        $('setProdDays').value = state.settings.prodDays;
        $('setRecycled').value = state.settings.recycledProgram;

        const dz = $('dropZone');
        $('fileInput').addEventListener('change', e => { addFiles(e.target.files); e.target.value = ''; });
        ['dragenter', 'dragover'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add('over'); }));
        ['dragleave', 'drop'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove('over'); }));
        dz.addEventListener('drop', e => { if (e.dataTransfer && e.dataTransfer.files.length) addFiles(e.dataTransfer.files); });

        $('tblFiles').addEventListener('change', e => {
            const id = Number(e.target.dataset.kind);
            if (!id) return;
            const f = state.files.find(x => x.id === id);
            f.kind = e.target.value;
            parseFile(f);
            refresh();
            runOcr();
        });
        $('tblFiles').addEventListener('click', e => {
            const id = Number(e.target.dataset.del);
            if (!id) return;
            state.files = state.files.filter(x => x.id !== id);
            refresh();
        });
        $('tblInv').addEventListener('change', onManualChange);
        $('setProdDays').addEventListener('change', e => { state.settings.prodDays = Math.max(0, parseInt(e.target.value, 10) || 0); refresh(); });
        $('setRecycled').addEventListener('change', e => { state.settings.recycledProgram = e.target.value; refresh(); });
        [['setConsequence', 'consequenceDate'], ['setColX', 'colX'], ['setColY', 'colY'], ['setColAD', 'colAD']].forEach(([id, key]) => {
            $(id).value = state.settings[key] || '';
            $(id).addEventListener('change', e => { state.settings[key] = e.target.value.trim(); saveStore(); });
        });
        $('btnReset').addEventListener('click', () => {
            if (!confirm('Hapus semua file yang sudah di-upload? (Isian No/tanggal DO tetap diingat di browser ini)')) return;
            state.files = [];
            refresh();
        });
        $('btnGenerate').addEventListener('click', generate);
    }

    document.addEventListener('DOMContentLoaded', init);
})();
