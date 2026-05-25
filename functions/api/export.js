import ExcelJS from 'exceljs';

export async function onRequestGet({ env }) {
    try {
        // Ambil semua data dari database
        const { results } = await env.DB.prepare(
            `SELECT id, sheet_name, kolom_data, created_at FROM data_mentah ORDER BY sheet_name, id`
        ).all();

        // Kelompokkan data per sheet
        const dataPerSheet = {};
        for (const row of results) {
            if (!dataPerSheet[row.sheet_name]) {
                dataPerSheet[row.sheet_name] = [];
            }
            dataPerSheet[row.sheet_name].push(JSON.parse(row.kolom_data));
        }

        // Buat workbook Excel
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "Kodim 1609 Buleleng";
        workbook.created = new Date();

        // Definisi kolom per tipe (sama dengan frontend)
        const kolomDefinitions = {
            tipeA: [
                "no", "satuan", "bekal", "jenis", "pemilik", "alamat_desa",
                "jarak_instansi_mil", "jarak_jalur_ekonomi", "luas_ha", 
                "produksi_ton", "waktu_pakai_guna", "no_registrasi", "keterangan"
            ],
            tipeB: [
                "no", "satuan", "bekal", "jenis", "pemilik", "alamat",
                "jarak_instansi_mil", "jarak_jalur_ekonomi", "kapasitas",
                "waktu_pakai_guna", "no_registrasi", "keterangan"
            ],
            tipeC: [
                "no", "satuan", "bekal", "jenis", "pemilik", "alamat",
                "jarak_instansi_mil", "jarak_jalur_ekonomi", "kapasitas",
                "waktu_pakai_guna", "no_registrasi", "keterangan"
            ],
            tipeD: [
                "no", "satuan", "bekal", "jenis", "pemilik", "alamat",
                "jarak_instansi_mil", "jarak_jalur_ekonomi", "kapasitas",
                "waktu_pakai_guna", "no_registrasi", "keterangan"
            ],
            tipeE: [
                "no", "satuan", "bekal", "jenis", "pemilik", "alamat",
                "jarak_instansi_mil", "jarak_jalur_ekonomi", "kapasitas",
                "waktu_pakai_guna", "no_registrasi", "keterangan"
            ],
            tipeF: [
                "no", "satuan", "wilayah", "nama", "nik", "pangkat_jabatan",
                "kategori", "usia", "keterangan"
            ],
            tipeG: [
                "no", "satuan", "nama_pemilik", "nama_perusahaan", "alamat",
                "koordinat_x", "koordinat_y", "jumlah", "no_telepon", "keterangan"
            ]
        };

        // Mapping sheet ke tipe
        const sheetToTipe = {
            "Padi": "tipeA", "Jagung": "tipeA", "Kedelai": "tipeA", "Kacang Tanah": "tipeA",
            "Kacang Hijau": "tipeA", "Ubi Kayu": "tipeA", "Ubi Jalar": "tipeA",
            "Jeruk": "tipeA", "Pisang": "tipeA", "Jahe": "tipeA", "Kentang": "tipeA",
            "Kopi": "tipeA", "Teh": "tipeA", "Kelapa Sawit": "tipeA", "Kakao": "tipeA",
            "Tembakau": "tipeA", "Karet": "tipeA", "Cengkeh": "tipeA", "Tebu": "tipeA",
            "Kapas": "tipeA", "Lada": "tipeA", "Sapi": "tipeA", "Kerbau": "tipeA",
            "Kambing": "tipeA", "Domba": "tipeA", "Ayam Ras": "tipeA",
            "Pakaian": "tipeB", "Sepatu": "tipeB", "Helm": "tipeB", "Ransel": "tipeB", "Alat Makan": "tipeB",
            "Pertamax": "tipeC", "Premium": "tipeC", "Avtur": "tipeC", "Avgas": "tipeC", "Solar": "tipeC",
            "Batu Bata": "tipeD", "Kayu": "tipeD", "Semen": "tipeD", "Kaca": "tipeD", "Pipa": "tipeD",
            "Blusting": "tipeE", "Catridge": "tipeE", "Propelan": "tipeE", "Fuse": "tipeE", "Pyrotechnix": "tipeE",
            "Purnawirawan": "tipeF", "Menwa": "tipeF", "Pol PP": "tipeF", "Polsus": "tipeF", 
            "Satpam": "tipeF", "Linmas": "tipeF", "Ormas": "tipeF",
            "Sarana Transportasi": "tipeG", "Prasarana Transportasi": "tipeG", 
            "Bengkel": "tipeG", "Sarana Siber": "tipeG", "Prasarana Siber": "tipeG"
        };

        // Header label untuk tampilan Excel (sesuai format asli)
        const headerLabels = {
            tipeA: ["No", "Satuan", "Bekal", "Jenis", "Pemilik", "Alamat (Desa)", 
                    "Jarak Dari Instansi Mil (Km)", "Jarak Dari Jalur Ekonomi", 
                    "Luas (Ha)", "Produksi (Ton)", "Waktu Pakai/Guna", "No.Registrasi", "Ket"],
            tipeB: ["No", "Satuan", "Bekal", "Jenis", "Pemilik", "Alamat", 
                    "Jarak Dari Instansi Mil (Km)", "Jarak Dari Jalur Ekonomi", 
                    "Kapasitas", "Waktu Pakai/Guna", "No Registrasi", "Ket"],
            tipeC: ["No", "Satuan", "Bekal", "Jenis", "Pemilik", "Alamat", 
                    "Jarak Dari Instansi Mil (Km)", "Jarak Dari Jalur Ekonomi", 
                    "Kapasitas (Ton)", "Waktu Pakai/Guna", "No Registrasi", "Ket"],
            tipeD: ["No", "Satuan", "Bekal", "Jenis", "Pemilik", "Alamat", 
                    "Jarak Dari Instansi Mil (Km)", "Jarak Dari Jalur Ekonomi", 
                    "Kapasitas", "Waktu Pakai/Guna", "No Registrasi", "Ket"],
            tipeE: ["No", "Satuan", "Bekal", "Jenis", "Pemilik", "Alamat", 
                    "Jarak Dari Instansi Mil (Km)", "Jarak Dari Jalur Ekonomi", 
                    "Kapasitas (Buah)", "Waktu Pakai/Guna", "No Registrasi", "Ket"],
            tipeF: ["No", "Satuan", "Wilayah", "Nama", "NIK", "Pangkat/Jabatan", 
                    "Kategori", "Usia", "Ket"],
            tipeG: ["No", "Satuan", "Nama Pemilik/Pengelola", "Nama Perusahaan", "Alamat", 
                    "Koordinat X", "Koordinat Y", "Jumlah", "No Telepon", "Ket"]
        };

        // Daftar semua sheet yang mungkin (62 sheet)
        const allSheets = [
            "Padi", "Jagung", "Kedelai", "Kacang Tanah", "Kacang Hijau", "Ubi Kayu", "Ubi Jalar",
            "Jeruk", "Pisang", "Jahe", "Kentang",
            "Kopi", "Teh", "Kelapa Sawit", "Kakao", "Tembakau", "Karet", "Cengkeh", "Tebu", "Kapas", "Lada",
            "Sapi", "Kerbau", "Kambing", "Domba", "Ayam Ras",
            "Pakaian", "Sepatu", "Helm", "Ransel", "Alat Makan",
            "Pertamax", "Premium", "Avtur", "Avgas", "Solar",
            "Batu Bata", "Kayu", "Semen", "Kaca", "Pipa",
            "Blusting", "Catridge", "Propelan", "Fuse", "Pyrotechnix",
            "Purnawirawan", "Menwa", "Pol PP", "Polsus", "Satpam", "Linmas", "Ormas",
            "Sarana Transportasi", "Prasarana Transportasi", "Bengkel", "Sarana Siber", "Prasarana Siber"
        ];

        // Buat sheet untuk setiap komoditas
        for (const sheetName of allSheets) {
            const tipe = sheetToTipe[sheetName];
            if (!tipe) continue;

            const worksheet = workbook.addWorksheet(sheetName, { properties: { tabColor: { argb: "FF2C7A2C" } } });
            const kolomKeys = kolomDefinitions[tipe];
            const headerNames = headerLabels[tipe];

            // Set kolom
            worksheet.columns = kolomKeys.map((key, idx) => ({
                header: headerNames[idx],
                key: key,
                width: 20
            }));

            // Style header
            worksheet.getRow(1).font = { bold: true };
            worksheet.getRow(1).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: "FFD3E6D3" }
            };

            // Data dari database
            const rowsData = dataPerSheet[sheetName] || [];
            
            // Urutkan berdasarkan no
            rowsData.sort((a, b) => (a.no || 0) - (b.no || 0));
            
            for (const row of rowsData) {
                const rowData = {};
                for (const key of kolomKeys) {
                    rowData[key] = row[key] || "";
                }
                worksheet.addRow(rowData);
            }

            // Jika tidak ada data, tambahkan baris kosong dengan catatan
            if (rowsData.length === 0) {
                worksheet.addRow({});
                const emptyRow = worksheet.getRow(2);
                emptyRow.getCell(1).value = "Belum ada data";
            }
        }

        // Buat buffer Excel
        const buffer = await workbook.xlsx.writeBuffer();

        return new Response(buffer, {
            headers: {
                "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Content-Disposition": `attachment; filename="Pendataan_Kodim_1609_${new Date().toISOString().slice(0,19)}.xlsx"`
            }
        });
    } catch (error) {
        return new Response(JSON.stringify({ error: error.message }), {
            status: 500,
            headers: { "Content-Type": "application/json" }
        });
    }
}
