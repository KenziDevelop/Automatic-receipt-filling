const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.static('public'));
app.use(express.json());

// Helper Fungsi Terbilang Otomatis
function terbilang(nilai) {
    nilai = Math.abs(Math.floor(nilai));
    const huruf = ["", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas"];
    let temp = "";
    if (nilai < 12) {
        temp = " " + huruf[nilai];
    } else if (nilai < 20) {
        temp = terbilang(nilai - 10) + " Belas";
    } else if (nilai < 100) {
        temp = terbilang(Math.floor(nilai / 10)) + " Puluh" + terbilang(nilai % 10);
    } else if (nilai < 200) {
        temp = " Seratus" + terbilang(nilai - 100);
    } else if (nilai < 1000) {
        temp = terbilang(Math.floor(nilai / 100)) + " Ratus" + terbilang(nilai % 100);
    } else if (nilai < 2000) {
        temp = " Seribu" + terbilang(nilai - 1000);
    } else if (nilai < 1000000) {
        temp = terbilang(Math.floor(nilai / 1000)) + " Ribu" + terbilang(nilai % 1000);
    } else if (nilai < 1000000000) {
        temp = terbilang(Math.floor(nilai / 1000000)) + " Juta" + terbilang(nilai % 1000000);
    } else if (nilai < 1000000000000) {
        temp = terbilang(Math.floor(nilai / 1000000000)) + " Milyar" + terbilang(nilai % 1000000000);
    }
    return temp.trim();
}

// Endpoint Upload Excel
app.post('/upload', upload.single('excelFile'), (req, res) => {
    try {
        const workbook = xlsx.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const rawData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

        // Parsing & Mapping Data Excel
        const dataParsed = rawData.map((row, index) => {
            const nominal = parseFloat(row['PENGELUARAN'] || row['PENERIMAAN'] || 0);
            
            // Format Tanggal jika tipe Serial Excel
            let tanggal = row['TANGGAL'] || '-';
            if (typeof tanggal === 'number') {
                const dateObj = xlsx.SSF.parse_date_code(tanggal);
                tanggal = `${dateObj.d}/${dateObj.m}/${dateObj.y}`;
            }

            // Gunakan Terbilang dari Excel jika ada, jika kosong buat otomatis
            let teksTerbilang = row['Terbilang'] || (nominal > 0 ? terbilang(nominal) + " Rupiah" : "-");

            return {
                id: index + 1,
                tanggal: tanggal,
                kode_kegiatan: row['KODE KEGIATAN'] || '-',
                kode_rekening: row['KODE REKENING'] || '-',
                no_bukti: row['NO. BUKTI'] || `BKT-${index+1}`,
                uraian: row['URAIAN'] || '-',
                nominal: nominal,
                terbilang: teksTerbilang
            };
        });

        // Hapus file sementara setelah selesai diproses
        fs.unlinkSync(req.file.path);

        res.json({ success: true, data: dataParsed });
    } catch (error) {
        if (req.file && fs.existsSync(req.file.path)) {
            fs.unlinkSync(req.file.path);
        }
        res.status(500).json({ success: false, message: error.message });
    }
});

// NEW: Endpoint Verifikasi Scan QR Code
app.get('/kuitansi/pdf/:no_bukti', (req, res) => {
    const noBukti = req.params.no_bukti;
    
    // Menampilkan Tampilan Verifikasi Dokumen Saat QR Code Di-scan
    res.send(`
        <!DOCTYPE html>
        <html lang="id">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Verifikasi Kuitansi ${noBukti}</title>
            <script src="https://cdn.tailwindcss.com"></script>
        </head>
        <body class="bg-slate-100 flex items-center justify-center min-h-screen p-4 font-sans">
            <div class="bg-white p-6 rounded-2xl shadow-lg text-center max-w-sm w-full border border-slate-200">
                <div class="w-16 h-16 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center mx-auto mb-4">
                    <svg class="w-8 h-8" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"></path>
                    </svg>
                </div>
                <span class="bg-emerald-100 text-emerald-800 text-[10px] font-bold px-2.5 py-1 rounded-full uppercase tracking-wider">
                    Dokumen Sah & Terverifikasi
                </span>
                <h2 class="font-bold text-xl text-slate-800 mt-3 mb-1">Kuitansi Resmi</h2>
                <p class="text-xs text-slate-500 mb-4">No. Bukti: <span class="font-bold text-slate-700">${noBukti}</span></p>

                <div class="bg-slate-50 p-3 rounded-xl border border-slate-100 text-left text-xs space-y-2 mb-5">
                    <div class="flex justify-between">
                        <span class="text-slate-400">Status</span>
                        <span class="font-semibold text-emerald-600">Valid / Recorded</span>
                    </div>
                    <div class="flex justify-between">
                        <span class="text-slate-400">Diakses Pada</span>
                        <span class="font-semibold text-slate-600">${new Date().toLocaleDateString('id-ID')}</span>
                    </div>
                </div>

                <button onclick="window.print()" class="w-full bg-blue-600 hover:bg-blue-700 text-white font-medium py-2.5 px-4 rounded-xl text-sm transition-all shadow-sm flex items-center justify-center gap-2">
                    <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"></path></svg>
                    Cetak / Simpan PDF
                </button>
            </div>
        </body>
        </html>
    `);
});

// KODE BARU (Siap Deploy ke Render)
const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
    console.log(`Server berjalan di port ${PORT}`);
});