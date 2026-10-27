/************************************************************
 * 🦇 AGENTE A.I.F.R.E.A.D - E-MAILS "AGENDAS PSA" (Gmail → Controle)
 *
 * O planejamento manda e-mails como:
 *   Assunto: "Agendas PSA D000086721001 - <projeto> - WF 5370347"
 *   Agendas criadas:
 *     - CIBKPRTH-15864 - Criação de Agenda - Projeto : D000086721001 - Recurso: NOME - De 02/11 à 22/12 - Exceto 02/11 e 22/12
 *   Observações: Projeto liberado para apontamentos · Margem Projeto · Vlr. Projeto ·
 *                Custo Total/Horas/Despesas Planejadas · Saldo a Planejar
 *   Respostas "Errata" na mesma conversa corrigem a linha da agenda.
 *
 * Lê SÓ e-mails com esse assunto (na caixa de quem roda), guarda o histórico na
 * aba 'PSA_Emails' e resume na coluna 'PSA (e-mail)' do Controle.
 ************************************************************/

const CONFIG_PSA = {
  BUSCA: 'subject:"Agendas PSA" newer_than:180d',
  MAX_CONVERSAS: 60,
  ABA: 'PSA_Emails'
};
const COLS_PSA = ['Conversa (ID)', 'Assunto', 'Código PSA', 'Projeto', 'WF', 'Último e-mail em', 'Remetente',
  'Agendas', 'Liberado p/ apontamento', 'Margem', 'Vlr. Projeto (R$)', 'Custo Total Planejado (R$)',
  'Custo Horas Planejadas (R$)', 'Custo Despesas Planejadas (R$)', 'Saldo a Planejar (R$)', 'Lido em'];

function abaPsaEmails_() {
  const ss = planilha_();
  let aba = ss.getSheetByName(CONFIG_PSA.ABA);
  if (aba) return aba;
  aba = ss.insertSheet(CONFIG_PSA.ABA);
  aba.getRange(1, 1, 1, COLS_PSA.length).setValues([COLS_PSA])
    .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff').setWrap(true);
  aba.setFrozenRows(1);
  [140, 380, 130, 300, 90, 130, 200, 460, 110, 90, 130, 130, 130, 130, 130, 130].forEach((w, i) => aba.setColumnWidth(i + 1, w));
  return aba;
}

/** Só o texto novo de cada e-mail (corta a parte citada da resposta). */
function textoNovoEmail_(txt) {
  const s = String(txt || '').replace(/\r/g, '');
  const corte = s.search(/^\s*(Em .{5,120}escreveu:|On .{5,120}wrote:|-{3,}\s*(Mensagem original|Original Message)|De:\s.+\n\s*(Enviad|Data|Sent))/im);
  return (corte > 0 ? s.substring(0, corte) : s).split('\n').filter(l => !/^\s*>/.test(l)).join('\n');
}

const RX_AGENDA_PSA = /([A-Z][A-Z0-9_]+-\d+)\s*-\s*Cria[çc][ãa]o de Agenda\s*-\s*Projeto\s*:?\s*(D\d{12})\s*-\s*Recurso\s*:?\s*(.+?)\s*-\s*De\s+(\d{2}\/\d{2}(?:\/\d{2,4})?)\s*(?:[àa]|até|a)\s*(\d{2}\/\d{2}(?:\/\d{2,4})?)(?:\s*-\s*Exceto\s+(.+))?/i;

/** Lê um e-mail: { agendas:[...], obs:{...} } */
function lerEmailPsa_(texto) {
  texto = String(texto || '').replace(/\r/g, '').replace(/\*/g, '').replace(/\u00a0/g, ' ');
  const r = { agendas: [], obs: {}, errata: /^\s*errata\b/im.test(texto) };
  // Linhas quebradas pelo e-mail ("Recurso:\n   NOME - De…") voltam a ser uma só
  const unidas = texto.replace(/\n(?![ \t]*\n)(?![ \t]*(?:[-•–·]|\d+[.)])\s)(?![ \t]*[A-ZÀ-Ú][\wÀ-ú ]{2,40}:\s*$)/g, ' ');
  unidas.split('\n').forEach(l => {
    const linha = l.replace(/^[\s*•\-–·]+/, '').trim();
    const m = linha.match(RX_AGENDA_PSA);
    if (m) r.agendas.push({ issue: m[1].toUpperCase(), codigo: m[2], recurso: m[3].trim(), de: m[4], ate: m[5],
      exceto: ((m[6] || '').match(/^[\d\/\s,;e]+/i) || [''])[0].replace(/[\s,;e]+$/i, '').trim() });
  });
  const valor = rx => { const m = texto.match(rx); return m ? m[1].trim() : ''; };
  if (/liberad[oa] para apontamentos?/i.test(texto)) r.obs.liberado = 'Sim';
  r.obs.margem = valor(/Margem(?:\s+(?:do\s+)?Projeto)?\s*:\s*(-?[\d.,]+\s*%)/i);
  r.obs.valorProjeto = valor(/Vl?r\.?\s*(?:do\s+)?Projeto\s*:\s*R?\$?\s*([\d.,]+)/i);
  r.obs.custoTotal = valor(/Custo\s+Total\s+Planejado\s*:\s*R?\$?\s*([\d.,]+)/i);
  r.obs.custoHoras = valor(/Custo\s+(?:de\s+)?Horas\s+Planejadas\s*:\s*R?\$?\s*([\d.,]+)/i);
  r.obs.custoDespesas = valor(/Custo\s+(?:de\s+)?Despesas\s+Planejadas\s*:\s*R?\$?\s*([\d.,]+)/i);
  r.obs.saldo = valor(/Saldo\s+a\s+Planejar\s*:\s*R?\$?\s*(-?[\d.,]+)/i);
  return r;
}

/** Junta uma conversa (e-mails do mais antigo ao mais novo; a errata corrige a agenda). */
function lerConversaPsa_(th) {
  const assunto = th.getFirstMessageSubject() || '';
  const mA = assunto.match(/Agendas\s+PSA\s+(D\d{12})\s*-\s*(.*?)(?:\s*-\s*WF\s*(\d+))?\s*$/i);
  const c = { id: th.getId(), assunto: assunto, codigo: mA ? mA[1] : '', projeto: mA ? mA[2].trim() : '', wf: mA && mA[3] ? mA[3] : '',
    agendas: {}, obs: {}, ultimo: null, remetente: '' };
  th.getMessages().forEach(msg => {
    const novo = textoNovoEmail_(msg.getPlainBody());
    const e = lerEmailPsa_(novo);
    const quando = msg.getDate();
    e.agendas.forEach(a => {
      const ant = c.agendas[a.issue];
      const mudou = ant && (ant.de !== a.de || ant.ate !== a.ate || ant.exceto !== a.exceto);
      c.agendas[a.issue] = Object.assign({}, a, {
        criadaEm: ant ? ant.criadaEm : quando, corrigidaEm: ant && (mudou || e.errata) ? quando : (ant ? ant.corrigidaEm : null) });
      if (!c.codigo) c.codigo = a.codigo;
    });
    Object.keys(e.obs).forEach(k => { if (e.obs[k]) c.obs[k] = e.obs[k]; });
    c.ultimo = quando; c.remetente = msg.getFrom();
  });
  return c;
}

/** Procura os e-mails, grava o histórico e devolve as conversas lidas. */
function buscarEmailsPsa_() {
  const conversas = GmailApp.search(CONFIG_PSA.BUSCA, 0, CONFIG_PSA.MAX_CONVERSAS).map(lerConversaPsa_)
    .filter(c => c.codigo || Object.keys(c.agendas).length);
  if (!conversas.length) return [];
  const aba = abaPsaEmails_();
  const ids = aba.getLastRow() > 1 ? aba.getRange(2, 1, aba.getLastRow() - 1, 1).getDisplayValues().map(r => r[0]) : [];
  const tz = Session.getScriptTimeZone(), fmt = d => d ? Utilities.formatDate(d, tz, 'dd/MM/yyyy HH:mm') : '';
  conversas.forEach(c => {
    const agendas = Object.keys(c.agendas).map(k => { const a = c.agendas[k];
      return `${k} · ${a.recurso} · ${a.de} a ${a.ate}${a.exceto ? ' · exceto ' + a.exceto : ''}${a.corrigidaEm ? ' (errata ' + fmt(a.corrigidaEm).substring(0, 5) + ')' : ''}`; }).join('\n');
    const o = c.obs;
    const linha = [c.id, c.assunto, c.codigo, c.projeto, c.wf, c.ultimo, c.remetente, agendas, o.liberado || '', o.margem || '',
      o.valorProjeto || '', o.custoTotal || '', o.custoHoras || '', o.custoDespesas || '', o.saldo || '', new Date()];
    let i = ids.indexOf(c.id);
    const n = i === -1 ? aba.getLastRow() + 1 : i + 2;
    if (i === -1) ids.push(c.id);
    aba.getRange(n, 1, 1, linha.length).setValues([linha]);
    aba.getRange(n, 6).setNumberFormat('dd/MM/yyyy HH:mm'); aba.getRange(n, 16).setNumberFormat('dd/MM/yyyy HH:mm');
  });
  return conversas;
}

/** Leva o que veio dos e-mails para o Controle (coluna 'PSA (e-mail)' + status). */
function aplicarEmailsPsaNoControle_(aba, conversas) {
  const res = { lidos: conversas.length, agendas: 0, projetos: 0, divergencias: [] };
  if (!conversas.length || aba.getLastRow() < 2) return res;
  const n = aba.getLastRow() - 1;
  const v = aba.getRange(2, 1, n, COLS_CT.length).getDisplayValues();
  const g = (i, c) => String(v[i][C_CT[c] - 1] || '').trim();
  const tz = Session.getScriptTimeZone(), dm = d => d ? Utilities.formatDate(d, tz, 'dd/MM') : '';
  const ddmm = s => String(s || '').substring(0, 5);

  // Agendas: pela chave da issue
  const porIssue = {}, porCodigo = {};
  conversas.forEach(c => {
    Object.keys(c.agendas).forEach(k => porIssue[k] = c.agendas[k]);
    if (c.codigo && (!porCodigo[c.codigo] || c.ultimo > porCodigo[c.codigo].ultimo)) porCodigo[c.codigo] = c;
  });

  v.forEach((row, i) => {
    const linha = i + 2, tipo = g(i, 'Tipo'), key = g(i, 'Issue'), st = g(i, 'Status Controle');
    let resumo = '';
    if (tipo === 'Agenda' && porIssue[key]) {
      const a = porIssue[key];
      const difere = (g(i, 'Período De') && ddmm(g(i, 'Período De')) !== ddmm(a.de)) || (g(i, 'Período Até') && ddmm(g(i, 'Período Até')) !== ddmm(a.ate));
      resumo = `✅ Criada no PSA ${dm(a.criadaEm)} · ${a.de} a ${a.ate}${a.exceto ? ' · exceto ' + a.exceto : ''}` +
        (a.corrigidaEm ? ` · errata ${dm(a.corrigidaEm)}` : '') +
        (difere ? ` · ⚠️ período diferente do Controle (${ddmm(g(i, 'Período De'))} a ${ddmm(g(i, 'Período Até'))})` : '');
      if (difere) res.divergencias.push(key);
      res.agendas++;
    } else if (tipo === 'Projeto Interno' && porCodigo[g(i, 'Código Projeto (PSA)')]) {
      const o = porCodigo[g(i, 'Código Projeto (PSA)')].obs, c = porCodigo[g(i, 'Código Projeto (PSA)')];
      resumo = [o.liberado ? '✅ Liberado para apontamentos' : '', o.margem ? 'Margem ' + o.margem : '',
        o.valorProjeto ? 'Vlr. projeto R$ ' + o.valorProjeto : '', o.custoTotal ? 'Custo planejado R$ ' + o.custoTotal : '',
        o.saldo ? 'Saldo a planejar R$ ' + o.saldo : '', c.wf ? 'WF ' + c.wf : ''].filter(Boolean).join(' · ');
      if (resumo) res.projetos++;
    }
    if (!resumo) return;
    if (g(i, 'PSA (e-mail)') !== resumo) aba.getRange(linha, C_CT['PSA (e-mail)']).setValue(resumo);
    // Agenda criada / projeto liberado no PSA → Aprovado (se ainda estava no começo)
    if (CONFIG_CT.AVANCA_COM_PSA.indexOf(st) !== -1) aba.getRange(linha, C_CT['Status Controle']).setValue('Aprovado');
  });
  return res;
}

/** Usado pelo Atualizar (painel e gatilho). Sem permissão do Gmail, só avisa. */
function sincronizarEmailsPsa_(aba) {
  try {
    return aplicarEmailsPsaNoControle_(aba, buscarEmailsPsa_());
  } catch (e) {
    return { erro: /permiss|autoriz|authoriz|scope/i.test(e.message)
      ? 'Gmail ainda não autorizado: rode no menu 🦇 → 📬 Ler e-mails do PSA (uma vez).' : e.message };
  }
}

/** Menu: lê agora e autoriza o Gmail na primeira vez. */
function lerEmailsPsaMenu() {
  const ui = SpreadsheetApp.getUi();
  const r = sincronizarEmailsPsa_(abaControle_());
  ui.alert('📬 E-mails "Agendas PSA"', r.erro ? '❌ ' + r.erro :
    `${r.lidos} conversa(s) lida(s) · ${r.agendas} agenda(s) e ${r.projetos} projeto(s) atualizados no Controle` +
    (r.divergencias && r.divergencias.length ? `\n⚠️ Período diferente do Controle: ${r.divergencias.join(', ')}` : '') +
    '\n\nHistórico completo na aba PSA_Emails.', ui.ButtonSet.OK);
}
