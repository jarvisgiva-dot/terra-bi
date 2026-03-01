import { createClient } from '@supabase/supabase-js';
import * as XLSX from 'xlsx';

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_KEY
);

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();
  if (req.query.secret !== process.env.SYNC_SECRET)
    return res.status(401).json({ error: 'Unauthorized' });

  try {
    // Baixar xlsx do Drive (link direto de download)
    const fileId = process.env.DRIVE_FILE_ID;
    const url = `https://drive.google.com/uc?export=download&id=${fileId}`;
    const response = await fetch(url);
    const buffer = await response.arrayBuffer();

    const wb = XLSX.read(buffer, { type: 'array', cellDates: true });

    // --- RESUMO ---
    const wsResumo = wb.Sheets['RESUMO'];
    const resumoData = XLSX.utils.sheet_to_json(wsResumo);
    const r = resumoData[0];
    await supabase.from('resumo').upsert([{
      id: 1,
      safra: '2024/2025',
      total_colhido: r['TOTAL COLHIDO'],
      area_total: r['ÁREA TOTAL'],
      area_colhida: r['ÁREA COLHIDA'],
      pct_colhido: r['PERCENTUAL COLHIDO'],
      area_nao_colhida: r['ÁREA NÃO COLHIDA'],
      media_geral: r['MÉDIA GERAL'],
      media_umidade: r['MÉDIA UMIDADE'],
      media_impureza: r['MÉDIA IMPUREZA'],
      total_desconto: r['TOTAL DESCONTO'],
      desconto_sc_ha: r['DESCONTO SC/HÁ'],
      updated_at: new Date().toISOString()
    }]);

    // --- ARMAZEM ---
    const wsArm = wb.Sheets['ARMAZEM'];
    const armData = XLSX.utils.sheet_to_json(wsArm);
    const armRows = armData.filter(r => r['ARMAZEM'] && r['TOTAL']).map(r => ({
      nome: r['ARMAZEM'], total_sc: r['TOTAL'], safra: '2024/2025'
    }));
    await supabase.from('armazem').delete().eq('safra','2024/2025');
    await supabase.from('armazem').insert(armRows);

    // --- COLHEITA DIÁRIA ---
    const wsData = wb.Sheets['DATA DE COLHEITA'];
    const dataRows = XLSX.utils.sheet_to_json(wsData, { cellDates: true });
    const diarioRows = dataRows.filter(r => r['DATA'] && r['TOTAL COLHIDO']).map(r => ({
      data: r['DATA'] instanceof Date ? r['DATA'].toISOString().split('T')[0] : r['DATA'],
      total_colhido: r['TOTAL COLHIDO'],
      area_colhida: r['ÁREA COLHIDA'],
      acumulado: r['TOTAL'],
      safra: '2024/2025'
    }));
    await supabase.from('colheita_diaria').delete().eq('safra','2024/2025');
    await supabase.from('colheita_diaria').insert(diarioRows);

    // --- TALHÕES ---
    const wsTal = wb.Sheets['PRODUTIVIDADE'];
    const talRows = XLSX.utils.sheet_to_json(wsTal);
    const talhoes = talRows.filter(r => r['TALHÃO'] && r['TALHÃO'] !== 'MÉDIA GERAL' && r['ÁREA TOTAL']).map(r => ({
      talhao: r['TALHÃO'], safra: r['SAFRA'] || '2024/2025',
      area: r['ÁREA TOTAL'], total_sc: r['TOTAL SC'],
      produtividade: r['PRODUTIVIDADE'], ha_colhido: r['HECTARES'],
      pct_colhido: r['PERCENTUAL COLHIDO'], status: r['STATUS']
    }));
    await supabase.from('talhoes').delete().eq('safra','2024/2025');
    await supabase.from('talhoes').insert(talhoes);

    // --- HISTÓRICO ---
    const wsHist = wb.Sheets['PRODUTIVIDADE POR TALHÃO'];
    const histRows = XLSX.utils.sheet_to_json(wsHist);
    const historico = histRows.filter(r => r['TALHÃO'] && r['SAFRA'] && r['PRODUTIVIDADE'] > 0).map(r => ({
      talhao: r['TALHÃO'], safra: r['SAFRA'],
      total_sc: r['TOTAL COLHIDO'], produtividade: r['PRODUTIVIDADE'], area: r['HECTARES']
    }));
    await supabase.from('historico').delete();
    await supabase.from('historico').insert(historico);

    res.json({ ok: true, synced: { talhoes: talhoes.length, diario: diarioRows.length, historico: historico.length }});
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
}
