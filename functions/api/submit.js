export async function onRequestPost({ request, env }) {
    try {
        const { sheet_name, data } = await request.json();
        
        if (!sheet_name || !data) {
            return new Response("Sheet name dan data diperlukan", { status: 400 });
        }

        // Tambahkan timestamp dan nomor urut otomatis
        const now = new Date().toISOString();
        
        // Jika tidak ada no, cari no terakhir + 1
        let no = data.no;
        if (!no || no === "") {
            const lastRecord = await env.DB.prepare(
                `SELECT json_extract(kolom_data, '$.no') as no FROM data_mentah 
                 WHERE sheet_name = ? ORDER BY id DESC LIMIT 1`
            ).bind(sheet_name).first();
            
            if (lastRecord && lastRecord.no) {
                no = parseInt(lastRecord.no) + 1;
            } else {
                no = 1;
            }
            data.no = no;
        }

        // Simpan ke database
        await env.DB.prepare(
            `INSERT INTO data_mentah (sheet_name, kolom_data, created_at, updated_at) 
             VALUES (?, ?, ?, ?)`
        ).bind(sheet_name, JSON.stringify(data), now, now).run();

        return new Response(JSON.stringify({ success: true, message: "Data tersimpan" }), {
            headers: { "Content-Type": "application/json" }
        });
    } catch (error) {
        return new Response(JSON.stringify({ error: error.message }), {
            status: 500,
            headers: { "Content-Type": "application/json" }
        });
    }
}
