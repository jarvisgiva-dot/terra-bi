const { createClient } = require('@supabase/supabase-js');
const XLSX = require('xlsx');

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();
  
  // Parse query string manualmente para garantir
  const url = new URL(req.url, 'https://terra-bi-app.vercel.app');
  const secret = url.searchParams.get('secret');
  
  if (secret !== process.env.SYNC_SECRET) {
    return res.status(401).json({ error: 'Unauthorized', got: secret ? 'present' : 'missing' });
  }

  try {
    const supabase = createClient(
      process.env.SUPABASE_URL,
      process.env.SUPABASE_SERVICE_KEY
    );

    const fileId = process.env.DRIVE_FILE_ID;
    const url2 = `https://drive.google.com/uc?export=download&id=${fileId}&confirm=t`;
    const response = await fetch(url2);
    if (!response.ok) throw new Error(`Drive download failed: ${response.status}`);
    const buffer = await response.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array', cellDates: true });

    // RESUMO
    const wsR = wb.Sheets['RESUMO'];
    const resumoData = XLSX.utils.sheet_to_json(wsR);
    const r = resumoData[0];
    await supabase.from('resumo').upsert([{
      id: 1, safra: '2024/2025',
      total_colhido: r['TOTAL COLHIDO'], area_total: r['ÁREA TOTAL'],
      area_colhida: r['ÁREA COLHIDA'], pct_colhido: r['PERCENTUAL COLHIDO'],
      area_nao_colhida: r['ÁREA NÃO COLHIDA'], media_geral: r['MÉDIA GERAL'],
      media_umidade: r['MÉDIA UMIDADE'], media_impureza: r['MÉDIA IMPUREZA'],
      total_desconto: r['TOTAL DESCONTO'], desconto_sc_ha: r['DESCONTO SC/HÁ'],
      updated_at: new Date().toISOString()
    }]);

    // ARMAZEM
    const wsA = wb.Sheets['ARMAZEM'];
    const armData = XLSX.utils.sheet_to_json(wsA);
    const armRows = armData.filter(a => a['ARMAZEM'] && a['TOTAL'])
      .map(a => ({ safra: '2024/2025', nome: a['ARMAZEM'], total_sc: a['TOTAL'] }));
    await supabase.from('armazem').delete().eq('safra','2024/2025');
    if (armRows.length) await supabase.from('armazem').insert(armRows);

    // COLHEITA DIÁRIA
    const wsD = wb.Sheets['DATA DE COLHEITA'];
    const diaData = XLSX.utils.sheet_to_json(wsD, { cellDates: true });
    const diaRows = diaData.filter(d => d['DATA'] && d['TOTAL COLHIDO']).map(d => ({
      safra: '2024/2025',
      data: d['DATA'] instanceof Date ? d['DATA'].toISOString().split('T')[0] : d['DATA'],
      total_colhido: d['TOTAL COLHIDO'], area_colhida: d['ÁREA COLHIDA'], acumulado: d['TOTAL']
    }));
    await supabase.from('colheita_diaria').delete().eq('safra','2024/2025');
    if (diaRows.length) await supabase.from('colheita_diaria').insert(diaRows);

    // TALHÕES
    const wsT = wb.Sheets['PRODUTIVIDADE'];
    const talData = XLSX.utils.sheet_to_json(wsT);
    const talRows = talData.filter(t => t['TALHÃO'] && t['TALHÃO'] !== 'MÉDIA GERAL' && t['ÁREA TOTAL']).map(t => ({
      safra: '2024/2025', talhao: t['TALHÃO'], area: t['ÁREA TOTAL'],
      total_sc: t['TOTAL SC'], produtividade: t['PRODUTIVIDADE'],
      ha_colhido: t['HECTARES'], pct_colhido: t['PERCENTUAL COLHIDO'], status: t['STATUS']
    }));
    await supabase.from('talhoes').delete().eq('safra','2024/2025');
    if (talRows.length) await supabase.from('talhoes').insert(talRows);

    // HISTÓRICO
    const wsH = wb.Sheets['PRODUTIVIDADE POR TALHÃO'];
    const histData = XLSX.utils.sheet_to_json(wsH);
    const histRows = histData.filter(h => h['TALHÃO'] && h['SAFRA'] && h['PRODUTIVIDADE'] > 0).map(h => ({
      safra: h['SAFRA'], talhao: h['TALHÃO'],
      total_sc: h['TOTAL COLHIDO'], produtividade: h['PRODUTIVIDADE'], area: h['HECTARES']
    }));
    await supabase.from('historico').delete();
    if (histRows.length) await supabase.from('historico').insert(histRows);

    res.json({ ok: true, updated_at: new Date().toISOString(),
      counts: { armazem: armRows.length, diario: diaRows.length, talhoes: talRows.length, historico: histRows.length }
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
};
