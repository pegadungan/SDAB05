import { stringify } from 'csv-stringify/sync';

export async function onRequestGet({ env }) {
    try {
        // Ambil semua data dari database
        const { results } = await env.DB.prepare(
            `SELECT id, sheet_name, kolom_data, created_at FROM data_mentah ORDER BY sheet_name, id`
        ).all();

        // Flatten kolom_data (JSON) + sheet info ke array
        const dataRows = [];
        for (const row of results) {
            const kolomObj = JSON.parse(row.kolom_data);
            kolomObj.sheet_name = row.sheet_name;
            kolomObj.id = row.id;
            kolomObj.created_at = row.created_at;
            dataRows.push(kolomObj);
        }

        // Buat header otomatis dari semua keys
        const allKeysSet = new Set();
        for (const r of dataRows) for (const k of Object.keys(r)) allKeysSet.add(k);
        const columns = Array.from(allKeysSet); // urutan random, bisa disusun

        // Convert ke CSV
        const csv = stringify(dataRows, { header: true, columns });

        return new Response(csv, {
            headers: {
                "Content-Type": "text/csv",
                "Content-Disposition": `attachment; filename=\"Pendataan_Kodim_1609_${new Date().toISOString().slice(0,19)}.csv\"`
            }
        });
    } catch (error) {
        return new Response(JSON.stringify({ error: error.message }), {
            status: 500,
            headers: { "Content-Type": "application/json" }
        });
    }
}
