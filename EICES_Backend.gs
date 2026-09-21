// ═══════════════════════════════════════════════════════════════════
// EICES — Backend Google Apps Script v4
// Login por obra (senha única) + controle (senha única)
// Senhas e lista de obras ficam na própria planilha (abas Config e Obras)
// App Controle lê e grava nas abas Patrimônio, Movimentação, Manutenção,
// Locado e Rádio.
// ═══════════════════════════════════════════════════════════════════

const SHEET_ID = '1l6o-pJBzgz0llin0dD1GqkYJ9SifYhwTZ-x0cV1AQ-Y';

const ABA = {
  SOL: 'Solicitacoes', CFG: 'Config', OBRAS: 'Obras',
  PAT: 'Patrimônio', MOV: 'Movimentação', MAN: 'Manutenção',
  RAD: 'Rádio', LOC: 'Locado'
};

const COLS_SOL = [
  'protocolo','tipo','obra','solicitante','urgencia','data_necessidade',
  'data_devolucao','observacao','ferramentas','subtipo','patrimonio',
  'ferramenta','equipamento','descricao','em_uso','data_solicitacao',
  'status','resposta','atualizado_em'
];

const LINHA_INI = 7;      // primeira linha de dados nas abas da planilha de controle
const LINHA_MAX = 1006;   // limite usado pelas fórmulas da planilha

// ── utilidades ────────────────────────────────────────────────────
function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function ss() { return SpreadsheetApp.openById(SHEET_ID); }
function norm(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .trim().toUpperCase();
}
function aba(nome) {
  const planilha = ss();
  const exata = planilha.getSheetByName(nome);
  if (exata) return exata;
  const alvo = norm(nome);
  return planilha.getSheets().find(s => norm(s.getName()).indexOf(alvo) === 0) || null;
}
function fmtData(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}
function paraData(t) {
  if (!t) return new Date();
  const p = String(t).split('-');
  if (p.length === 3) return new Date(+p[0], +p[1] - 1, +p[2]);
  return new Date(t);
}
function txt(v) { return v === null || v === undefined ? '' : String(v).trim(); }

// ── configuração e senhas ─────────────────────────────────────────
function config() {
  const ws = aba(ABA.CFG);
  const cfg = {};
  if (!ws) return cfg;
  ws.getDataRange().getValues().forEach(r => { if (r[0]) cfg[txt(r[0])] = txt(r[1]); });
  return cfg;
}
function senhaObraOk(s)     { return txt(s) !== '' && txt(s) === config().SENHA_OBRA; }
function senhaControleOk(s) { return txt(s) !== '' && txt(s) === config().SENHA_CONTROLE; }

// ── entradas ──────────────────────────────────────────────────────
function doGet(e) {
  const acao = e && e.parameter ? e.parameter.acao : '';
  if (acao === 'ping') return jsonResp({ status: 'ok', hora: new Date().toISOString() });
  return jsonResp({ status: 'ok', msg: 'EICES Backend v4 ativo' });
}

function doPost(e) {
  try {
    if (!e || !e.postData) return jsonResp({ status: 'erro', msg: 'Sem dados' });
    const d = JSON.parse(e.postData.contents);
    const acao = d.acao || '';

    // públicas
    if (acao === 'ping')  return jsonResp({ status: 'ok' });
    if (acao === 'obras') return jsonResp({ status: 'ok', obras: listarObras(false) });

    if (acao === 'login') {
      const ok = d.perfil === 'controle' ? senhaControleOk(d.senha) : senhaObraOk(d.senha);
      return jsonResp(ok ? { status: 'ok' } : { status: 'erro', msg: 'Senha incorreta' });
    }

    // obra
    if (['nova_solicitacao', 'nova_manutencao', 'listar_obra'].indexOf(acao) >= 0) {
      if (!senhaObraOk(d.senha)) return jsonResp({ status: 'negado', msg: 'A senha das obras mudou. Entre novamente.' });
      if (acao === 'listar_obra') return jsonResp({ status: 'ok', registros: listarSolicitacoes(d.obra) });
      gravarSolicitacao(d.dados || {});
      return jsonResp({ status: 'ok' });
    }

    // controle
    if (!senhaControleOk(d.senha)) return jsonResp({ status: 'negado', msg: 'A senha de controle mudou. Entre novamente.' });

    if (acao === 'listar_todas')          return jsonResp({ status: 'ok', registros: listarSolicitacoes('') });
    if (acao === 'dados_controle')        return jsonResp({ status: 'ok', dados: dadosControle() });
    if (acao === 'atualizar_solicitacao') { atualizarSolicitacao(d.protocolo, d.campos || {}); return jsonResp({ status: 'ok' }); }
    if (acao === 'movimentar')            { const n = movimentar(d.itens || []); return jsonResp({ status: 'ok', gravados: n }); }
    if (acao === 'manutencao_enviar')     { enviarConserto(d); return jsonResp({ status: 'ok' }); }
    if (acao === 'manutencao_retorno')    { retornoConserto(d); return jsonResp({ status: 'ok' }); }
    if (acao === 'locado_encerrar')       { encerrarLocado(d.linha, d.data); return jsonResp({ status: 'ok' }); }

    return jsonResp({ status: 'erro', msg: 'Ação desconhecida: ' + acao });
  } catch (err) {
    return jsonResp({ status: 'erro', msg: err.toString() });
  }
}

// ── obras ─────────────────────────────────────────────────────────
function listarObras(incluirLocais) {
  const ws = aba(ABA.OBRAS);
  if (!ws) return [];
  return ws.getDataRange().getValues().slice(1)
    .filter(r => txt(r[0]) && norm(r[2]) !== 'NAO')
    .filter(r => incluirLocais || norm(r[1]) !== 'LOCAL')
    .map(r => ({ nome: txt(r[0]), tipo: norm(r[1]) === 'LOCAL' ? 'LOCAL' : 'OBRA' }));
}

// ── solicitações ──────────────────────────────────────────────────
function abaSolicitacoes() {
  let ws = aba(ABA.SOL);
  if (!ws) {
    ws = ss().insertSheet(ABA.SOL);
    ws.getRange(1, 1, 1, COLS_SOL.length).setValues([COLS_SOL])
      .setBackground('#2F4D65').setFontColor('#FFFFFF').setFontWeight('bold');
    ws.setFrozenRows(1);
  }
  return ws;
}

function gravarSolicitacao(dados) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ws = abaSolicitacoes();
    const linha = COLS_SOL.map(c => {
      if (c === 'ferramentas') return JSON.stringify(dados.ferramentas || []);
      if (c === 'resposta')    return '';
      if (c === 'status')      return dados.status || 'Pendente';
      if (c === 'data_solicitacao') return dados.data_solicitacao || new Date().toISOString();
      if (c === 'atualizado_em')    return new Date().toISOString();
      return txt(dados[c]);
    });
    ws.appendRow(linha);
  } finally { lock.releaseLock(); }
}

function listarSolicitacoes(obraFiltro) {
  const ws = aba(ABA.SOL);
  if (!ws) return [];
  const data = ws.getDataRange().getValues();
  if (data.length <= 1) return [];
  const hdr = data[0].map(txt);
  const alvo = norm(obraFiltro);
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const o = {};
    hdr.forEach((h, j) => {
      const v = data[i][j];
      const ehData = Object.prototype.toString.call(v) === '[object Date]';
      if (h === 'data_necessidade' || h === 'data_devolucao') o[h] = fmtData(v);
      else o[h] = ehData ? v.toISOString() : txt(v);
    });
    if (!o.protocolo) continue;
    if (alvo && norm(o.obra) !== alvo) continue;
    try { o.ferramentas = JSON.parse(o.ferramentas || '[]'); } catch (e) { o.ferramentas = []; }
    try { o.resposta = o.resposta ? JSON.parse(o.resposta) : null; } catch (e) { o.resposta = null; }
    out.push(o);
  }
  return out.reverse();
}

function atualizarSolicitacao(protocolo, campos) {
  const ws = aba(ABA.SOL);
  if (!ws) return;
  const data = ws.getDataRange().getValues();
  const hdr = data[0].map(txt);
  const colProto = hdr.indexOf('protocolo');
  for (let i = 1; i < data.length; i++) {
    if (txt(data[i][colProto]) !== txt(protocolo)) continue;
    Object.keys(campos).forEach(k => {
      const j = hdr.indexOf(k);
      if (j < 0) return;
      let v = campos[k];
      if (typeof v === 'object' && v !== null) v = JSON.stringify(v);
      ws.getRange(i + 1, j + 1).setValue(v);
    });
    const jAt = hdr.indexOf('atualizado_em');
    if (jAt >= 0) ws.getRange(i + 1, jAt + 1).setValue(new Date().toISOString());
    return;
  }
}

// ── leitura da planilha de controle (sem valores em R$) ───────────
function linhasDados(ws, ncol) {
  const ult = Math.min(ws.getLastRow(), LINHA_MAX);
  if (ult < LINHA_INI) return [];
  return ws.getRange(LINHA_INI, 1, ult - LINHA_INI + 1, ncol).getValues();
}

function dadosControle() {
  const out = { patrimonio: [], movimentacoes: [], manutencoes: [], locados: [], radios: [], obras: listarObras(true) };

  const pat = aba(ABA.PAT);
  if (pat) linhasDados(pat, 13).forEach(r => {
    if (!txt(r[0])) return;
    out.patrimonio.push({
      codigo: txt(r[0]), categoria: txt(r[1]), item: txt(r[2]), marca: txt(r[3]),
      modelo: txt(r[4]), serie: txt(r[5]), nf: txt(r[8]), obra: txt(r[9]),
      responsavel: txt(r[10]), status: txt(r[11]), obs: txt(r[12])
    });
  });

  const mov = aba(ABA.MOV);
  if (mov) linhasDados(mov, 8).forEach(r => {
    if (!txt(r[1])) return;
    out.movimentacoes.push({
      data: fmtData(r[0]), codigo: txt(r[1]), item: txt(r[2]), origem: txt(r[3]),
      destino: txt(r[4]), tipo: txt(r[5]), responsavel: txt(r[6]), obs: txt(r[7])
    });
  });

  const man = aba(ABA.MAN);
  if (man) linhasDados(man, 11).forEach(r => {
    if (!txt(r[1])) return;
    out.manutencoes.push({
      entrada: fmtData(r[0]), codigo: txt(r[1]), item: txt(r[2]), categoria: txt(r[3]),
      obra: txt(r[4]), retorno: fmtData(r[6]), status: txt(r[7]),
      obra_retorno: txt(r[8]), obs: txt(r[10])
    });
  });

  const loc = aba(ABA.LOC);
  if (loc) linhasDados(loc, 12).forEach((r, i) => {
    if (!txt(r[0])) return;
    out.locados.push({
      linha: LINHA_INI + i, fornecedor: txt(r[0]), categoria: txt(r[1]), item: txt(r[2]),
      obra: txt(r[3]), qtd: txt(r[4]), cod_locadora: txt(r[7]), inicio: fmtData(r[8]),
      devolucao: fmtData(r[9]), status: txt(r[10]), obs: txt(r[11])
    });
  });

  const rad = aba(ABA.RAD);
  if (rad) {
    const ult = Math.min(rad.getLastRow(), 37);
    if (ult >= 8) rad.getRange(8, 1, ult - 7, 9).getValues().forEach(r => {
      if (!txt(r[0]) || norm(r[3]).indexOf('TOTAL') === 0) return;
      out.radios.push({
        obra: txt(r[0]), locadora: txt(r[1]), qtd: txt(r[2]), inicio: fmtData(r[5]),
        devolucao: fmtData(r[6]), status: txt(r[7]), obs: txt(r[8])
      });
    });
  }
  return out;
}

// ── gravação na Movimentação ──────────────────────────────────────
function obraAtualDe(codigo) {
  const pat = aba(ABA.PAT);
  if (!pat) return '';
  const r = linhasDados(pat, 10).find(l => norm(l[0]) === norm(codigo));
  return r ? txt(r[9]) : '';
}

function proximaLinhaLivre(ws, col) {
  const vals = ws.getRange(LINHA_INI, col, LINHA_MAX - LINHA_INI + 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (!txt(vals[i][0])) return LINHA_INI + i;
  return Math.max(ws.getLastRow() + 1, LINHA_MAX + 1);
}

function movimentar(itens) {
  const ws = aba(ABA.MOV);
  if (!ws) throw new Error('Aba Movimentação não encontrada');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let n = 0;
  try {
    itens.forEach(it => {
      if (!txt(it.codigo)) return;
      const linha = proximaLinhaLivre(ws, 2);
      const origem = txt(it.origem) || obraAtualDe(it.codigo);
      ws.getRange(linha, 1, 1, 2).setValues([[paraData(it.data), txt(it.codigo).toUpperCase()]]);
      ws.getRange(linha, 4, 1, 5).setValues([[origem, txt(it.destino), txt(it.tipo), txt(it.responsavel) || 'EICES', txt(it.obs)]]);
      n++;
    });
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  return n;
}

// ── manutenção (ficha de conserto) ────────────────────────────────
function enviarConserto(d) {
  const ws = aba(ABA.MAN);
  if (!ws) throw new Error('Aba Manutenção não encontrada');
  const obra = txt(d.obra) || obraAtualDe(d.codigo);
  const linha = proximaLinhaLivre(ws, 2);
  ws.getRange(linha, 1, 1, 2).setValues([[paraData(d.data), txt(d.codigo).toUpperCase()]]);
  ws.getRange(linha, 5).setValue(obra);
  ws.getRange(linha, 8).setValue('EM CONSERTO');
  ws.getRange(linha, 11).setValue(txt(d.obs));
  movimentar([{ codigo: d.codigo, origem: obra, destino: 'MANUTENÇÃO', tipo: 'ENVIO P/ MANUTENÇÃO', data: d.data, obs: d.obs }]);
}

function retornoConserto(d) {
  const ws = aba(ABA.MAN);
  if (!ws) throw new Error('Aba Manutenção não encontrada');
  const linhas = linhasDados(ws, 8);
  let alvo = -1;
  linhas.forEach((r, i) => { if (norm(r[1]) === norm(d.codigo) && norm(r[7]) === 'EM CONSERTO') alvo = LINHA_INI + i; });
  const resultado = d.resultado === 'BAIXA' ? 'BAIXA' : 'VOLTOU';
  if (alvo > 0) {
    ws.getRange(alvo, 7).setValue(paraData(d.data));
    ws.getRange(alvo, 8).setValue(resultado);
    if (resultado === 'VOLTOU') ws.getRange(alvo, 9).setValue(txt(d.destino));
  }
  if (resultado === 'BAIXA') {
    movimentar([{ codigo: d.codigo, origem: 'MANUTENÇÃO', destino: '', tipo: 'BAIXA', data: d.data, obs: d.obs || 'Sem conserto' }]);
  } else {
    movimentar([{ codigo: d.codigo, origem: 'MANUTENÇÃO', destino: d.destino, tipo: 'RETORNO MANUTENÇÃO', data: d.data, obs: d.obs }]);
  }
}

// ── locado ────────────────────────────────────────────────────────
function encerrarLocado(linha, data) {
  const ws = aba(ABA.LOC);
  linha = parseInt(linha, 10);
  if (!ws || !(linha >= LINHA_INI) || !txt(ws.getRange(linha, 1).getValue())) throw new Error('Linha de locação inválida');
  ws.getRange(linha, 10).setValue(paraData(data));
}

// ═══════════════════════════════════════════════════════════════════
// SETUP — rode uma vez depois de colar o script
// Cria as abas Solicitacoes, Config e Obras sem mexer nas outras
// ═══════════════════════════════════════════════════════════════════
function setup() {
  const planilha = ss();
  abaSolicitacoes();

  if (!aba(ABA.CFG)) {
    const c = planilha.insertSheet(ABA.CFG);
    c.getRange(2, 2, 2, 1).setNumberFormat('@');
    c.getRange(1, 1, 3, 3).setValues([
      ['CHAVE', 'VALOR', 'O QUE É'],
      ['SENHA_OBRA', 'EICES1', 'Senha única para entrar no App Supervisor (todas as obras)'],
      ['SENHA_CONTROLE', '@EICES', 'Senha única para entrar no App Controle']
    ]);
    c.getRange(1, 1, 1, 3).setBackground('#2F4D65').setFontColor('#FFFFFF').setFontWeight('bold');
    c.setColumnWidth(1, 170); c.setColumnWidth(2, 140); c.setColumnWidth(3, 420);
  }

  if (!aba(ABA.OBRAS)) {
    const o = planilha.insertSheet(ABA.OBRAS);
    const obras = ['ALPHA ONE','ALTIER','ANGELICA','ARENGA (TRISUL)','ARTHUR PRADO','BALTAZAR',
      'BIOMA','CUNHA GAGO','ENXOVIA','INFRA','JERONIMO','JOA','MAURO','ORIÇANGA','PJM',
      'SAMPA CONSOLAÇÃO','TANABI','VIRGILIO'];
    const locais = ['ESCRITÓRIO','ASSISTENCIA TECNICA','MANUTENÇÃO'];
    const linhas = [['OBRA', 'TIPO', 'ATIVA']]
      .concat(obras.map(n => [n, 'OBRA', 'SIM']))
      .concat(locais.map(n => [n, 'LOCAL', 'SIM']));
    o.getRange(1, 1, linhas.length, 3).setValues(linhas);
    o.getRange(1, 1, 1, 3).setBackground('#2F4D65').setFontColor('#FFFFFF').setFontWeight('bold');
    o.setFrozenRows(1);
    o.setColumnWidth(1, 220);
    o.getRange(1, 5).setValue('TIPO: OBRA aparece no login das obras. LOCAL (escritório, assistência) aparece só no Controle. ATIVA = NÃO esconde a obra.');
  }
  Logger.log('✅ Setup concluído: Solicitacoes, Config e Obras prontas.');
}

function zerarSolicitacoes() {
  const ws = aba(ABA.SOL);
  if (ws && ws.getLastRow() > 1) ws.deleteRows(2, ws.getLastRow() - 1);
  Logger.log('✅ Solicitações zeradas');
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('🔧 EICES')
      .addItem('Configurar abas do app', 'setup')
      .addItem('Zerar solicitações dos apps', 'zerarSolicitacoes')
      .addToUi();
  } catch (e) {}
}
