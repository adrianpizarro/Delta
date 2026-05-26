const express = require('express');
const cors    = require('cors');
const XLSX    = require('xlsx');
const fs      = require('fs');
const path    = require('path');

const app        = express();
const PORT       = process.env.PORT || 8080;
const EXCEL_FILE = path.join(__dirname, 'verificacoes_delta.xlsx');
const API_KEY    = process.env.ANTHROPIC_API_KEY;

app.use(cors());
app.use(express.json({ limit: '15mb' }));

/* ── Sirve el HTML principal ── */
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'verificacao-carga.html'));
});

/* ── GET /api/status ── */
app.get('/api/status', (req, res) => {
  const exists = fs.existsSync(EXCEL_FILE);
  let sessions = 0;
  if (exists) {
    try {
      const wb = XLSX.readFile(EXCEL_FILE);
      sessions = wb.SheetNames.filter(n => n !== 'RESUMO').length;
    } catch {}
  }
  res.json({ ok: true, sessions, hasApiKey: !!API_KEY });
});

/* ── POST /api/ocr ── PROXY a Claude API ── */
app.post('/api/ocr', async (req, res) => {
  try {
    if (!API_KEY) {
      return res.status(500).json({ ok: false, error: 'Chave API nao configurada no servidor' });
    }
    const { imageBase64, imageMime } = req.body;
    if (!imageBase64) {
      return res.status(400).json({ ok: false, error: 'Sem imagem' });
    }

    const resp = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': API_KEY,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-6',
        max_tokens: 2000,
        messages: [{
          role: 'user',
          content: [
            { type: 'image', source: { type: 'base64', media_type: imageMime || 'image/jpeg', data: imageBase64 } },
            { type: 'text', text: `Analisa esta imagem de uma lista de picking de armazem logistico Delta Cafes.
A tabela pode estar rodada 90 graus — interpreta-a corretamente.

Extrai TODOS os artigos/linhas. As colunas tipicas sao:
- Material (codigo numerico)
- Texto breve material / Descricao
- Posicao no deposito (ex: EX.TRANS)
- Caixa Palete
- Un Caixa
- Total Caixas
- Paletes de Mercadoria (CRITICA - geralmente destacada a amarelo)
- Pedido Paletes
- Paletes Veiculo

Devolve APENAS JSON valido sem markdown, sem texto extra:
{"items":[{"material":"40166","descricao":"DELTA SOLUVEL CREME 160 G","posicao":"EX.TRANS","caixasPalete":126,"unCaixa":6,"totalCaixas":504,"paleteMercadoria":4,"paletePedido":4,"paleteVeiculo":4}]}

IMPORTANTE:
- Inclui TODOS os artigos visiveis
- Mantem os nomes em maiusculas exatamente como estao
- Se um numero nao for legivel usa 0
- A coluna "Paletes de Mercadoria" e a mais importante - extrai-a com cuidado`}
          ]
        }]
      })
    });

    if (!resp.ok) {
      const e = await resp.json().catch(() => ({}));
      return res.status(500).json({ ok: false, error: e.error?.message || 'Erro Claude API: '+resp.status });
    }

    const data = await resp.json();
    const txt = (data.content||[]).find(b=>b.type==='text')?.text || '{}';
    const clean = txt.replace(/```json|```/g,'').trim();
    const parsed = JSON.parse(clean);
    res.json({ ok: true, items: parsed.items || [] });

  } catch (err) {
    console.error('OCR error:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

/* ── POST /api/guardar ── */
app.post('/api/guardar', (req, res) => {
  try {
    const { company, truck, rows, counters, validade } = req.body;
    let wb;
    if (fs.existsSync(EXCEL_FILE)) {
      wb = XLSX.readFile(EXCEL_FILE);
    } else {
      wb = XLSX.utils.book_new();
      const wsR = XLSX.utils.aoa_to_sheet([
        ['DELTA CAFES — REGISTO MESTRE DE VERIFICACOES'],
        [],
        ['#','Data/Hora','Empresa','Verificadas','Total','Pal.Merc Verificadas','Pal.Merc Total','Estado'],
      ]);
      wsR['!cols'] = [4,18,14,12,8,16,14,12].map(w=>({wch:w}));
      XLSX.utils.book_append_sheet(wb, wsR, 'RESUMO');
    }

    const now = new Date();
    const ts  = now.toLocaleString('pt-PT');
    const okRefs = rows.filter((r,i) => (counters[i]||0) === r.paleteMercadoria && r.paleteMercadoria > 0).length;
    const palMercTotal = rows.reduce((s,r) => s + (r.paleteMercadoria||0), 0);
    const palMercVerif = rows.reduce((s,r,i) => s + (counters[i]||0), 0);

    const rWs   = wb.Sheets['RESUMO'];
    const rData = XLSX.utils.sheet_to_json(rWs, { header: 1 });
    const num   = rData.filter((r,i) => i >= 3 && r.length > 0).length + 1;
    rData.push([num, ts, company, okRefs, rows.length, palMercVerif, palMercTotal,
                okRefs === rows.length ? 'COMPLETO' : 'INCOMPLETO']);
    const newRWs = XLSX.utils.aoa_to_sheet(rData);
    newRWs['!cols'] = [4,18,14,12,8,16,14,12].map(w=>({wch:w}));
    wb.Sheets['RESUMO'] = newRWs;

    const base = (company.slice(0,2).toUpperCase()
      + '_' + (truck.data||'').replace(/-/g,'').slice(4)
      + '_' + (truck.viatura||'').replace(/[^A-Z0-9]/gi,'')
    ).slice(0, 26);
    let shName = base; let s = 1;
    while (wb.SheetNames.includes(shName)) shName = base.slice(0,24)+'_'+s++;

    const aoa = [
      [company.toUpperCase()+' / DELTA CAFES — VERIFICACAO DE CARGA'],
      ['Guardado em: '+ts],
      [],
      ['Material','Descricao','Posicao','Cx/Pal','Un/Cx','Total Cx',
       'Pal.Mercadoria','Pal.Verificadas','Pal.Pedido','Pal.Veiculo','Lote','Validade 1','Validade 2','Anotacoes'],
      ...rows.map((r,i) => [
        r.material, r.descricao, r.posicao,
        r.caixasPalete, r.unCaixa, r.totalCaixas,
        r.paleteMercadoria,
        `${counters[i]||0}/${r.paleteMercadoria} ${(counters[i]||0)===r.paleteMercadoria?'OK':'—'}`,
        r.paletePedido, r.paleteVeiculo,
        validade[i]?.lote||'', validade[i]?.v1||'', validade[i]?.v2||'',
        validade[i]?.anotacao||'',
      ]),
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [10,38,10,8,8,10,15,16,12,12,18,13,13,28].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb, ws, shName.slice(0,31));
    XLSX.writeFile(wb, EXCEL_FILE);

    console.log(`[${ts}] ${company} — ${okRefs}/${rows.length} refs, ${palMercVerif}/${palMercTotal} pal.merc`);
    res.json({ ok: true, session: shName, total: num, okRefs, palMercVerif, palMercTotal });
  } catch (err) {
    console.error('Error:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

/* ── GET /api/download ── */
app.get('/api/download', (req, res) => {
  if (!fs.existsSync(EXCEL_FILE)) {
    return res.status(404).json({ ok: false, error: 'Sem ficheiro ainda' });
  }
  res.download(EXCEL_FILE, 'verificacoes_delta.xlsx');
});

app.listen(PORT, () => {
  console.log('Delta Cafes — Servidor na porta ' + PORT);
  console.log('API Key configurada: ' + (API_KEY ? 'SIM' : 'NAO'));
});
