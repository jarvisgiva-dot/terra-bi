const { createClient } = require('@supabase/supabase-js');

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const url = new URL(req.url, 'https://terra-bi-app.vercel.app');
  const secret = url.searchParams.get('secret');
  if (secret !== process.env.SYNC_SECRET)
    return res.status(401).json({ error: 'Unauthorized' });

  try {
    const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_KEY);
    const data = req.body;

    if (data.resumo) {
      await supabase.from('resumo').upsert([{ id: 1, safra: '2024/2025', ...data.resumo, updated_at: new Date().toISOString() }]);
    }
    if (data.armazem?.length) {
      await supabase.from('armazem').delete().eq('safra','2024/2025');
      await supabase.from('armazem').insert(data.armazem.map(a => ({ ...a, safra: '2024/2025' })));
    }
    if (data.diario?.length) {
      await supabase.from('colheita_diaria').delete().eq('safra','2024/2025');
      await supabase.from('colheita_diaria').insert(data.diario.map(d => ({ ...d, safra: '2024/2025' })));
    }
    if (data.talhoes?.length) {
      await supabase.from('talhoes').delete().eq('safra','2024/2025');
      await supabase.from('talhoes').insert(data.talhoes.map(t => ({ ...t, safra: '2024/2025' })));
    }

    res.json({ ok: true, updated_at: new Date().toISOString() });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
};
