const express = require('express');
const cors    = require('cors');
const XLSX    = require('xlsx');
const fs      = require('fs');
const path    = require('path');

const app       = express();
const PORT      = 3001;
const EXCEL_FILE = path.join(__dirname, 'verificacoes_delta.xlsx');

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static(__dirname));      // sirve el HTML desde la misma carpeta

/* ── GET /api/status ─────────────────────────────────────── */
app.get('/api/status', (req, res) => {
  const exists = fs.existsSync(EXCEL_FILE);
  let sessions = 0;
  if (exists) {
    try {
      const wb = XLSX.readFile(EXCEL_FILE);
      sessions = wb.SheetNames.filter(n => n !== 'RESUMO').length;
    } catch {}
  }
  res.json({ ok: true, file: EXCEL_FILE, sessions });
});

/* ── POST /api/guardar ───────────────────────────────────── */
app.post('/api/guardar', (req, res) => {
  try {
    const { company, truck, rows, counters, validade } = req.body;

    // Cargar o crear el libro Excel
    let wb;
    if (fs.existsSync(EXCEL_FILE)) {
      wb = XLSX.readFile(EXCEL_FILE);
    } else {
      wb = XLSX.utils.book_new();
      const wsR = XLSX.utils.aoa_to_sheet([
        ['DELTA CAFES — REGISTO MESTRE DE VERIFICACOES'],
        [],
        ['#','Data/Hora','Empresa','Viatura','Motorista','Destino','Verificadas','Total','Estado'],
      ]);
      wsR['!cols'] = [4,18,14,14,20,20,12,8,12].map(w=>({wch:w}));
      XLSX.utils.book_append_sheet(wb, wsR, 'RESUMO');
    }

    const now = new Date();
    const ts  = now.toLocaleString('pt-PT');
    const ok  = rows.filter((r,i) => (counters[i]||0) === r.paletePedido && r.paletePedido > 0).length;

    // Actualizar hoja RESUMO
    const rWs   = wb.Sheets['RESUMO'];
    const rData = XLSX.utils.sheet_to_json(rWs, { header: 1 });
    const num   = rData.filter((r,i) => i >= 3 && r.length > 0).length + 1;
    rData.push([num, ts, company, truck.viatura||'', truck.motorista||'',
                truck.destino||'', ok, rows.length,
                ok === rows.length ? 'COMPLETO' : 'INCOMPLETO']);
    const newRWs = XLSX.utils.aoa_to_sheet(rData);
    newRWs['!cols'] = [4,18,14,14,20,20,12,8,12].map(w=>({wch:w}));
    wb.Sheets['RESUMO']      = newRWs;
    wb.SheetNames[wb.SheetNames.indexOf('RESUMO')] = 'RESUMO';

    // Nombre de la hoja de sesion (max 31 chars Excel)
    const base = (company.slice(0,2).toUpperCase()
      + '_' + (truck.data||'').replace(/-/g,'').slice(4)
      + '_' + (truck.viatura||'').replace(/[^A-Z0-9]/gi,'')
    ).slice(0, 26);
    let shName = base; let s = 1;
    while (wb.SheetNames.includes(shName)) shName = base.slice(0,24)+'_'+s++;

    // Construir hoja de sesion
    const aoa = [
      [company.toUpperCase()+' / DELTA CAFES — VERIFICACAO DE CARGA'],
      ['Guardado em: '+ts],
      [],
      ['Viatura:', truck.viatura||'', '', 'Data:', truck.data||''],
      ['Motorista:', truck.motorista||'', '', 'Destino:', truck.destino||''],
      ['Observacoes:', truck.obs||''],
      [],
      ['Material','Descricao','Posicao','Cx/Pal','Un/Cx','Total Cx',
       'Pal.Mercadoria','Pal.Pedido','Pal.Veiculo','Verificado','Lote','Validade 1','Validade 2'],
      ...rows.map((r,i) => [
        r.material, r.descricao, r.posicao,
        r.caixasPalete, r.unCaixa, r.totalCaixas,
        r.paleteMercadoria, r.paletePedido, r.paleteVeiculo,
        `${counters[i]||0}/${r.paletePedido} ${(counters[i]||0)===r.paletePedido?'OK':'—'}`,
        validade[i]?.lote||'', validade[i]?.v1||'', validade[i]?.v2||'',
      ]),
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [10,38,10,8,8,10,15,12,12,14,18,13,13].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb, ws, shName.slice(0,31));

    XLSX.writeFile(wb, EXCEL_FILE);
    console.log(`[${ts}] Guardado: ${company} ${truck.viatura||''} ${truck.data||''} — ${ok}/${rows.length} refs`);

    res.json({ ok: true, session: shName, total: num });
  } catch (err) {
    console.error('Error al guardar:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

/* ── GET /api/download ───────────────────────────────────── */
app.get('/api/download', (req, res) => {
  if (!fs.existsSync(EXCEL_FILE)) {
    return res.status(404).json({ ok: false, error: 'Sem ficheiro ainda' });
  }
  res.download(EXCEL_FILE, 'verificacoes_delta.xlsx');
});

app.listen(PORT, () => {
  console.log('');
  console.log('  Delta Cafes — Servidor de Verificacao');
  console.log('  Correndo em: http://localhost:' + PORT);
  console.log('  Excel em:    ' + EXCEL_FILE);
  console.log('');
});
