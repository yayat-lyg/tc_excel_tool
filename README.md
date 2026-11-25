# TC Excel Tool

Halaman HTML offline untuk menyiapkan data pengajuan TC Control Union dari dokumen shipment H&M.
Sumber data: **PDF (+ OCR scan DO) dan input manual** (tanpa Sage, tanpa server). File tidak dikirim ke mana pun.

## Cara pakai
1. Buka `index.html` (double-click, Chrome/Edge).
2. Upload semua PDF satu order (extract ZIP dulu):
   - Invoice H&M (wajib, PDF yang ada teksnya)
   - Purchase Order H&M (untuk cek qty & harga per negara)
   - BOM / Supplier Quote BOM (opsional: GSM & supplier kain)
   - Scan DO dengan nama file `000721 - LYG - EXP - VIII - 2026 - LD, OT, OU.pdf` (No DO + kode negara dari nama file, **tanggal DO dibaca OCR** dari header scan)
   - Lampiran lain (letter of statement, marker) hanya dicatat
3. Tunggu OCR tanggal DO selesai (~5 detik per DO, lihat teks progres). Cek kolom Jenis, lalu di tabel invoice: File DO dipilih otomatis (kode negara di nama file; kalau ada 2 DO dengan kode sama, dari baris stabilo kuning di scan). Pilih manual kalau masih kosong, cek **Tgl DO** yang bertanda "OCR – cek scan" atau kosong.
4. Isi sekali per order: Consequence Date (kolom W) dan nilai kolom X, Y, AD (kosongkan kalau tidak dipakai). Kolom Gudang (R) bisa diganti per invoice.
5. Lihat Temuan, lalu klik **Download Excel**.

## Isi Excel
Ringkasan · Shipment TC (1 baris per invoice; kolom A–AD sama dengan sheet "Summary Details" file Consequence, header baris 1, bisa langsung di-copy-paste; kolom bantu tool mulai AF) · Berat Serat (estimasi NW x %) · PO vs Invoice · BOM · Temuan · Dokumen.
Teks biru = dari PDF / input manual; hitam = formula; kuning = wajib diisi / boleh diubah.

## Batasan
- Parser mengikuti template invoice/PO/BOM H&M saat ini (dites dengan order 711225-7668, 14 invoice). Template lain perlu parser baru.
- Pemilihan DO dari scan hanya membaca baris yang di-stabilo kuning (nomor invoice, boleh salah 1 karakter, atau kode warehouse persis). DO tanpa stabilo = pilih manual kalau ada 2 DO dengan kode sama.
- OCR tanggal hanya membaca baris DATE dan DELIVERY NO di header DO. Tahun dikoreksi dengan nama hari + tahun No DO. Kalau tidak terbaca, isi manual. Tanggal OCR diberi catatan di sel Excel.
- OCR jalan di Web Worker (dibuat dari blob agar tetap jalan dari file://), halaman tetap bisa dipakai selama OCR. Kalau browser menolak worker, otomatis pindah ke halaman utama (halaman diam sebentar per DO).
- Isian manual diingat per nomor invoice di browser yang sama (localStorage), bukan backup.

## Struktur
- `js/parsers.js` baca PDF invoice / PO / BOM / nama file DO
- `js/ocr.js` OCR offline (tesseract) untuk header scan DO
- `js/validate.js` cek antar dokumen, auto-pilih DO
- `js/excel.js` susun workbook (ExcelJS)
- `js/app.js` UI
- `lib/` pdf.js 3.11.174 (Apache-2.0), ExcelJS 4.4.0 (MIT), tesseract.js-core 5.1.1 (sebagai string di tesseract-core-src.js) + eng.traineddata 4.0.0_best_int (Apache-2.0), lokal agar jalan offline
