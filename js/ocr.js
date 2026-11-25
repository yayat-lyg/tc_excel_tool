/*
 * OCR for scanned DO (Surat Jalan) PDFs, fully offline.
 * tesseract.js-core (wasm inlined) runs in a Web Worker so the page does not freeze.
 * From file:// a worker cannot be loaded from a file, so it is built from a blob:
 * core source (lib/tesseract/tesseract-core-src.js, a string) + the handler below.
 * If the browser refuses the worker, OCR falls back to the main thread (page pauses, but still works).
 * Libraries are loaded only when the first DO is processed.
 */
(function (global) {
    'use strict';

    let enginePromise = null;

    function loadScript(src) {
        return new Promise((resolve, reject) => {
            const s = document.createElement('script');
            s.src = src;
            s.onload = resolve;
            s.onerror = () => reject(new Error('Gagal memuat ' + src));
            document.head.appendChild(s);
        });
    }

    async function gunzipBase64(b64) {
        const bin = atob(b64);
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip'));
        return new Uint8Array(await new Response(stream).arrayBuffer());
    }

    // Runs inside the worker (or on the main thread as fallback). Must stay self-contained: it is stringified.
    function createOcrHandler(TesseractCoreFn) {
        let M = null, api = null;
        return async function (msg) {
            if (msg.type === 'init') {
                M = await TesseractCoreFn();
                M.FS.writeFile('./eng.traineddata', msg.data);
                api = new M.TessBaseAPI();
                if (api.Init(null, 'eng', 1) === -1) throw new Error('OCR gagal inisialisasi');
                return true;
            }
            if (msg.type === 'recognize') {
                M.FS.writeFile('/input', msg.png);
                if (api.SetImageFile(1, 0) === 1) throw new Error('OCR gagal membaca gambar');
                api.SetVariable('tessedit_pageseg_mode', String(msg.psm || 6)); // 6 = block of text (header), 7 = single line (table row)
                const out = [];
                for (const r of msg.rects) {
                    api.SetRectangle(r.left, r.top, r.width, r.height);
                    api.Recognize(null);
                    out.push(api.GetUTF8Text() || '');
                }
                return out;
            }
            throw new Error('Perintah OCR tidak dikenal: ' + msg.type);
        };
    }

    function workerBackend(coreSrc) {
        const glue = `
const __h = (${createOcrHandler.toString()})(TesseractCore);
self.onmessage = async (e) => {
    try { self.postMessage({ id: e.data.id, ok: true, result: await __h(e.data) }); }
    catch (err) { self.postMessage({ id: e.data.id, ok: false, error: String((err && err.message) || err) }); }
};`;
        const url = URL.createObjectURL(new Blob([coreSrc, '\n', glue], { type: 'text/javascript' }));
        const worker = new Worker(url);
        const pending = new Map();
        let seq = 0;
        const failAll = msg => { pending.forEach(p => p.reject(new Error(msg))); pending.clear(); };
        worker.onmessage = e => {
            const p = pending.get(e.data.id);
            if (!p) return;
            pending.delete(e.data.id);
            e.data.ok ? p.resolve(e.data.result) : p.reject(new Error(e.data.error));
        };
        worker.onerror = e => { e.preventDefault(); failAll('Worker OCR error: ' + (e.message || 'tidak diketahui')); };
        return {
            mode: 'worker',
            call(msg, transfer) {
                return new Promise((resolve, reject) => {
                    const id = ++seq;
                    pending.set(id, { resolve, reject });
                    worker.postMessage({ ...msg, id }, transfer || []);
                });
            },
            dispose() { worker.terminate(); URL.revokeObjectURL(url); }
        };
    }

    function mainBackend(coreSrc) {
        (0, eval)(coreSrc); // defines global TesseractCore
        const h = createOcrHandler(global.TesseractCore);
        return { mode: 'main', call: msg => h(msg), dispose() {} };
    }

    function engine() {
        if (!enginePromise) {
            enginePromise = (async () => {
                await loadScript('lib/tesseract/tesseract-core-src.js');
                await loadScript('lib/tesseract/eng-data.js');
                const coreSrc = global.TC_TESS_CORE_SRC;
                const data = await gunzipBase64(global.TC_ENG_TRAINEDDATA_GZ_B64);
                delete global.TC_TESS_CORE_SRC;
                delete global.TC_ENG_TRAINEDDATA_GZ_B64;
                try {
                    const w = workerBackend(coreSrc);
                    await w.call({ type: 'init', data: data.slice() }); // copy: keep data for the fallback
                    return w;
                } catch (e) {
                    console.warn('OCR worker tidak tersedia, pakai main thread:', e.message);
                    const m = mainBackend(coreSrc);
                    await m.call({ type: 'init', data });
                    return m;
                }
            })();
            enginePromise.catch(() => { enginePromise = null; });
        }
        return enginePromise;
    }

    function canvasPng(canvas) {
        return new Promise(resolve => canvas.toBlob(b => b.arrayBuffer().then(a => resolve(new Uint8Array(a))), 'image/png'));
    }

    /**
     * Recognise text in a canvas.
     * rects: optional [{left, top, width, height}] in fractions of the canvas size; one text per rect.
     * psm: tesseract page segmentation mode, default 6 (block); 7 = single text line.
     */
    async function recognize(canvas, rects, psm) {
        const eng = await engine();
        const png = await canvasPng(canvas);
        const px = (rects || [{ left: 0, top: 0, width: 1, height: 1 }]).map(r => ({
            left: Math.round(r.left * canvas.width), top: Math.round(r.top * canvas.height),
            width: Math.round(r.width * canvas.width), height: Math.round(r.height * canvas.height)
        }));
        return eng.call({ type: 'recognize', png, rects: px, psm }, [png.buffer]);
    }

    async function mode() { return (await engine()).mode; }

    global.TcOcr = { recognize, engine, mode };
})(window);
