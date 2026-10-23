/************************************************************
 * AGENTE JIRA IA - CONTROLE DE SOLICITAÇÕES (v1.0)
 * Aba 'Controle_Solicitacoes' com Projetos Internos e Agendas:
 *  - linha criada automaticamente ao abrir a issue pelos formulários
 *  - "Atualizar Controle" busca no Jira o status e o nº da
 *    solicitação no comentário ("Solicitação 5370347 iniciada...")
 *  - issues já criadas são importadas das abas de rascunho
 ************************************************************/

const CONFIG_CT = {
  ABA: 'Controle_Solicitacoes',
  STATUS_CONTROLE: ['Aguardando solicitação', 'Solicitação aberta', 'Aprovado', 'Em andamento',
    'Encerrado', 'Cancelado'],
  // Padrão do comentário automático com o número da solicitação
  REGEX_SOLICITACAO: /Solicita[çc][ãa]o\s*\[?\s*#?\s*(\d{4,})/i
};

const COLS_CT = [
  'Tipo', 'Issue', 'Aberta em', 'Solicitante', 'Código Projeto (PSA)', 'Projeto', 'Analista / Recurso',
  'Squad', 'Período De', 'Período Até', 'Valor Hora (R$)', 'Dias', 'Horas', 'Valor Total (R$)',
  'Status Jira', 'Nº Solicitação', 'Link Solicitação', 'Status Controle', 'Observações', 'Última atualização'
];
const C_CT = {}; COLS_CT.forEach((n, i) => C_CT[n] = i + 1);

/************************************************************
 * ABA
 ************************************************************/
function abaControle_() {
  const ss = planilha_();
  let aba = ss.getSheetByName(CONFIG_CT.ABA);
  if (aba) return aba;

  aba = ss.insertSheet(CONFIG_CT.ABA, 0);
  aba.getRange(1, 1, 1, COLS_CT.length).setValues([COLS_CT])
    .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff').setWrap(true)
    .setVerticalAlignment('middle');
  aba.setFrozenRows(1);
  aba.setFrozenColumns(2);
  aba.setRowHeight(1, 36);

  // Colunas preenchidas manualmente: destaque amarelo no cabeçalho
  [C_CT['Status Controle'], C_CT['Observações']].forEach(c =>
    aba.getRange(1, c).setBackground('#ffab00').setFontColor('#172b4d'));
  // Colunas atualizadas pelo Jira: destaque verde
  [C_CT['Status Jira'], C_CT['Nº Solicitação'], C_CT['Link Solicitação']].forEach(c =>
    aba.getRange(1, c).setBackground('#00875a'));

  const larguras = { 'Tipo': 110, 'Issue': 120, 'Projeto': 280, 'Analista / Recurso': 200, 'Squad': 150,
    'Link Solicitação': 180, 'Status Controle': 160, 'Observações': 240 };
  COLS_CT.forEach((n, i) => aba.setColumnWidth(i + 1, larguras[n] || 105));
  formatarColunasControle_(aba);
  return aba;
}

function formatarColunasControle_(aba) {
  const n = Math.max(aba.getMaxRows() - 1, 1);
  aba.getRange(2, C_CT['Aberta em'], n, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  aba.getRange(2, C_CT['Período De'], n, 2).setNumberFormat('dd/MM/yyyy');
  aba.getRange(2, C_CT['Valor Hora (R$)'], n, 1).setNumberFormat('R$ #,##0.00');
  aba.getRange(2, C_CT['Dias'], n, 2).setNumberFormat('#,##0');
  aba.getRange(2, C_CT['Valor Total (R$)'], n, 1).setNumberFormat('R$ #,##0.00');
  aba.getRange(2, C_CT['Última atualização'], n, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  aba.getRange(2, C_CT['Status Controle'], n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(CONFIG_CT.STATUS_CONTROLE, true)
      .setAllowInvalid(true).build());

  // Cores por status de controle
  const faixa = aba.getRange(2, 1, n, COLS_CT.length);
  const col = colunaLetra_(C_CT['Status Controle']);
  const regra = (txt, cor) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${col}2="${txt}"`).setBackground(cor).setRanges([faixa]).build();
  aba.setConditionalFormatRules([
    regra('Cancelado', '#ffebe6'), regra('Encerrado', '#f4f5f7'),
    regra('Em andamento', '#e3fcef'), regra('Aprovado', '#deebff')
  ]);
}

function colunaLetra_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function abrirControle() {
  const aba = abaControle_();
  planilha_().setActiveSheet(aba);
}

/************************************************************
 * REGISTRO (chamado pelos formulários ao criar a issue)
 ************************************************************/
function registrarNoControle_(tipo, key, url, dados, solicitante, criadaEm) {
  return registrarLinhaControle_(tipo, key, url, resumoParaControle_(tipo, dados), solicitante, criadaEm);
}

/** Grava a linha a partir do resumo já calculado. Devolve false se a issue já estava no controle. */
function registrarLinhaControle_(tipo, key, url, r, solicitante, criadaEm) {
  const aba = abaControle_();
  if (linhaControlePorIssue_(aba, key)) return false;
  const linha = Array(COLS_CT.length).fill('');
  linha[C_CT['Tipo'] - 1] = tipo;
  linha[C_CT['Issue'] - 1] = `=HYPERLINK("${url}"; "${key}")`;
  linha[C_CT['Aberta em'] - 1] = criadaEm || new Date();
  linha[C_CT['Solicitante'] - 1] = String(solicitante || Session.getActiveUser().getEmail() || '').split('@')[0];
  linha[C_CT['Código Projeto (PSA)'] - 1] = r.codigoProjeto;
  linha[C_CT['Projeto'] - 1] = r.projeto;
  linha[C_CT['Analista / Recurso'] - 1] = r.analista;
  linha[C_CT['Squad'] - 1] = r.squad;
  linha[C_CT['Período De'] - 1] = r.de;
  linha[C_CT['Período Até'] - 1] = r.ate;
  linha[C_CT['Valor Hora (R$)'] - 1] = r.valorHora;
  linha[C_CT['Dias'] - 1] = r.dias;
  linha[C_CT['Horas'] - 1] = r.horas;
  linha[C_CT['Valor Total (R$)'] - 1] = r.valorTotal;
  linha[C_CT['Status Controle'] - 1] = 'Aguardando solicitação';
  linha[C_CT['Última atualização'] - 1] = new Date();

  aba.getRange(aba.getLastRow() + 1, 1, 1, COLS_CT.length).setValues([linha]);
  return true;
}

/** Extrai do formulário os números de controle (mesmas regras das Tabelas 02). */
function resumoParaControle_(tipo, d) {
  const data = s => { const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : ''; };

  if (tipo === 'Agenda') {
    let dias = 0, horas = 0;
    (d.meses || []).forEach(m => { dias += num_(m.dias); horas += num_(m.horas); });
    const vh = num_(d.valorHora);
    return {
      codigoProjeto: String(d.codigoProjeto || '').toUpperCase(),
      projeto: d.nomeProjeto || '',
      analista: [d.nomeRecurso, d.consultoria ? `(${d.consultoria})` : ''].filter(Boolean).join(' '),
      squad: d.squad || '', de: data(d.dataDe), ate: data(d.dataAte),
      valorHora: vh || '', dias: dias, horas: horas, valorTotal: Math.round(horas * vh * 100) / 100
    };
  }

  // Projeto Interno: Tabela 02 (A×B×C = horas; F+G = valor)
  let dias = 0, horas = 0, valor = 0, valorHoras = 0;
  (d.meses || []).forEach(m => {
    const a = num_(m.a), b = num_(m.b), c = num_(m.c), e = num_(m.e), g = num_(m.g);
    const D = a * b * c;
    dias += b; horas += D; valorHoras += D * e; valor += D * e + g;
  });
  return {
    codigoProjeto: '', projeto: d.nomeProjeto || '',
    analista: String(d.analistas || '').trim() || (d.teamLeads ? 'TL: ' + d.teamLeads : ''),
    squad: d.squad || '', de: data(d.dataInicio), ate: data(d.dataFim),
    valorHora: horas ? Math.round(valorHoras / horas * 100) / 100 : '',   // média ponderada, sem despesas
    dias: dias, horas: horas || num_(d.horasTotal), valorTotal: Math.round(valor * 100) / 100
  };
}

function linhaControlePorIssue_(aba, key) {
  if (aba.getLastRow() < 2) return 0;
  const keys = aba.getRange(2, C_CT['Issue'], aba.getLastRow() - 1, 1).getDisplayValues();
  for (let i = 0; i < keys.length; i++) if (keys[i][0] === key) return i + 2;
  return 0;
}

/************************************************************
 * IMPORTA ISSUES JÁ CRIADAS (a partir das abas de rascunho)
 ************************************************************/
function importarDosRascunhos_() {
  const ss = planilha_();
  const fontes = [
    { aba: 'Rascunhos_PI', tipo: 'Projeto Interno' },
    { aba: 'Rascunhos_Agenda', tipo: 'Agenda' }
  ];
  let novos = 0;
  const controle = abaControle_();
  fontes.forEach(f => {
    const aba = ss.getSheetByName(f.aba);
    if (!aba || aba.getLastRow() < 2) return;
    const n = aba.getLastRow() - 1;
    const txt = aba.getRange(2, 1, n, 9).getDisplayValues();
    const datas = aba.getRange(2, 2, n, 1).getValues();
    txt.forEach((r, i) => {
      const key = r[5];
      if (r[4] !== 'Criada' || !/^[A-Z][A-Z0-9_]+-\d+$/.test(key)) return;
      if (linhaControlePorIssue_(controle, key)) return;
      let dados = {};
      try { dados = JSON.parse(r[8] || '{}'); } catch (e) {}
      registrarNoControle_(f.tipo, key, `${CONFIG.JIRA_BASE_URL}/browse/${key}`, dados, r[2], datas[i][0] || '');
      novos++;
    });
  });
  return novos;
}

/************************************************************
 * ATUALIZAÇÃO: STATUS JIRA + Nº DA SOLICITAÇÃO (comentários)
 ************************************************************/
function atualizarControle() {
  const ui = SpreadsheetApp.getUi();
  const importados = importarDosRascunhos_();
  const r = atualizarControleSilencioso_();
  planilha_().setActiveSheet(abaControle_());
  ui.alert('🔄 Controle atualizado',
    `Issues verificadas: ${r.verificadas}\n` +
    `Nº de solicitação encontrados: ${r.comSolicitacao}\n` +
    (r.novasSolicitacoes ? `🆕 Novas solicitações nesta atualização: ${r.novasSolicitacoes}\n` : '') +
    (importados ? `Importadas dos rascunhos: ${importados}\n` : '') +
    (r.analistasVinculados ? `👤 Projetos com analistas trazidos das Agendas: ${r.analistasVinculados}\n` : '') +
    (r.canceladas ? `🚫 Marcadas como Cancelado (canceladas no Jira): ${r.canceladas}\n` : '') +
    (r.erros.length ? `\n⚠️ Não encontradas no Jira: ${r.erros.join(', ')}` : ''),
    ui.ButtonSet.OK);
}

/** Versão sem janelas — usada também pelo agendamento automático. */
function atualizarControleSilencioso_() {
  const aba = abaControle_();
  const res = { verificadas: 0, comSolicitacao: 0, novasSolicitacoes: 0, erros: [] };
  if (aba.getLastRow() < 2) return res;

  const n = aba.getLastRow() - 1;
  const keys = aba.getRange(2, C_CT['Issue'], n, 1).getDisplayValues().map(r => r[0].trim());
  const validas = keys.filter(k => /^[A-Z][A-Z0-9_]+-\d+$/.test(k));
  const info = {};

  for (let i = 0; i < validas.length; i += 40) {
    const lote = validas.slice(i, i + 40);
    const resp = chamarJiraApiCriacao_('/rest/api/2/search', 'post', {
      jql: `key in (${lote.join(',')})`, maxResults: lote.length,
      fields: ['status', 'comment'], validateQuery: false
    });
    (resp.issues || []).forEach(is => { info[is.key] = extrairInfoIssue_(is); });
  }

  // Reforço: se a busca não trouxe o nº, lê os comentários direto da issue
  Object.keys(info).forEach(key => {
    if (info[key].numero) return;
    try {
      const comentarios = buscarComentariosIssue_(key);
      const it = extrairInfoIssue_({ fields: { status: { name: info[key].status }, comment: { comments: comentarios } } });
      if (it.numero) info[key] = it;
    } catch (e) {}
  });

  const agora = new Date();
  const colSt = C_CT['Status Jira'], colNum = C_CT['Nº Solicitação'], colLink = C_CT['Link Solicitação'];
  const colCtl = C_CT['Status Controle'], colUpd = C_CT['Última atualização'];
  const atuais = aba.getRange(2, 1, n, COLS_CT.length).getValues();

  keys.forEach((key, idx) => {
    if (!key) return;
    const linha = idx + 2;
    const it = info[key];
    if (!it) { res.erros.push(key); return; }
    res.verificadas++;
    aba.getRange(linha, colSt).setValue(it.status);
    if (it.numero) {
      res.comSolicitacao++;
      if (String(atuais[idx][colNum - 1]) !== String(it.numero)) res.novasSolicitacoes++;
      aba.getRange(linha, colNum).setValue(it.numero);
      aba.getRange(linha, colLink).setFormula(it.link
        ? `=HYPERLINK("${it.link}"; "Acessar #${it.numero}")` : '');
      if (atuais[idx][colCtl - 1] === 'Aguardando solicitação') aba.getRange(linha, colCtl).setValue('Solicitação aberta');
    }
    // Cancelada no Jira → Cancelado no controle (sai das horas e do valor)
    if (/cancel/i.test(it.status) && atuais[idx][colCtl - 1] !== 'Cancelado') {
      aba.getRange(linha, colCtl).setValue('Cancelado');
      res.canceladas = (res.canceladas || 0) + 1;
    }
    aba.getRange(linha, colUpd).setValue(agora);
  });

  try { res.analistasVinculados = vincularAnalistasDasAgendas_(aba); } catch (e) {}
  return res;
}

/************************************************************
 * ANALISTAS: leva para a linha do Projeto Interno os recursos
 * das Agendas abertas para o mesmo projeto.
 * Casa pelo Código PSA (se preenchido nas duas linhas) ou pelo
 * nome do projeto. Nunca apaga nomes digitados à mão.
 ************************************************************/
function vincularAnalistasDasAgendas_(aba) {
  if (aba.getLastRow() < 2) return 0;
  const n = aba.getLastRow() - 1;
  const cTipo = C_CT['Tipo'] - 1, cCod = C_CT['Código Projeto (PSA)'] - 1,
    cProj = C_CT['Projeto'] - 1, cAna = C_CT['Analista / Recurso'] - 1, cCtl = C_CT['Status Controle'] - 1;
  const dados = aba.getRange(2, 1, n, COLS_CT.length).getDisplayValues();
  const norm = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9]+/g, ' ').trim();
  const semConsultoria = s => String(s || '').replace(/\s*\(.*\)\s*$/, '').trim();

  // Agendas (não canceladas) por código PSA e por nome de projeto
  const porCod = {}, porNome = {};
  dados.forEach(r => {
    if (r[cTipo] !== 'Agenda' || r[cCtl] === 'Cancelado') return;
    const nome = semConsultoria(r[cAna]);
    if (!nome) return;
    const add = (mapa, k) => { if (!k) return; (mapa[k] = mapa[k] || []).indexOf(nome) === -1 && mapa[k].push(nome); };
    add(porCod, norm(r[cCod]));
    add(porNome, norm(r[cProj]));
  });

  let alterados = 0;
  dados.forEach((r, i) => {
    if (r[cTipo] !== 'Projeto Interno') return;
    const achados = [].concat(porCod[norm(r[cCod])] || [], porNome[norm(r[cProj])] || [])
      .filter((v, j, a) => a.indexOf(v) === j);
    if (!achados.length) return;

    const atual = String(r[cAna] || '').trim();
    const ehTL = !atual || /^TL:/i.test(atual);
    const existentes = ehTL ? [] : atual.split(/\s*[,;]\s*/).filter(Boolean);
    const faltam = achados.filter(a => existentes.map(norm).indexOf(norm(a)) === -1);
    if (!faltam.length) return;

    aba.getRange(i + 2, cAna + 1).setValue(existentes.concat(faltam).join(', '));
    alterados++;
  });
  return alterados;
}

/** Status atual e nº/link da solicitação a partir dos comentários (o mais recente vence). */
function extrairInfoIssue_(issue) {
  const f = issue.fields || {};
  const comentarios = ((f.comment && f.comment.comments) || []).slice().reverse();
  let numero = '', link = '';
  for (const c of comentarios) {
    const corpo = String(c.body || '');
    const m = corpo.match(CONFIG_CT.REGEX_SOLICITACAO);
    if (!m) continue;
    numero = m[1];
    const urls = corpo.match(/https?:\/\/[^\s\]|"<>]+/g) || [];
    link = urls.find(u => u.indexOf(numero) !== -1) || urls[0] || '';
    break;
  }
  return { status: (f.status && f.status.name) || '', numero: numero, link: link };
}

/************************************************************
 * IMPORTAR ISSUES ABERTAS POR OUTRAS PESSOAS
 * Aceita chaves (CIBKPRTH-7250, CIBKPRTH-7301) ou uma JQL.
 * Os dados vêm da descrição (layout dos templates 7190/7191).
 ************************************************************/
function importarIssuesExistentes() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('📥 Importar issue(s) para o Controle',
    'Informe uma ou mais chaves separadas por vírgula ou espaço:\n' +
    '   ex.: CIBKPRTH-7250, CIBKPRTH-7301\n\n' +
    'ou uma busca JQL:\n' +
    '   ex.: project = CIBKPRTH AND summary ~ "Criação de Agenda" AND created >= -30d',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const entrada = resp.getResponseText().trim();
  if (!entrada) return;

  const r = importarIssuesParaControle_(entrada);
  if (r.importadas.length) atualizarControleSilencioso_();
  planilha_().setActiveSheet(abaControle_());

  ui.alert('📥 Importação concluída',
    `✅ Importadas: ${r.importadas.length}${r.importadas.length ? ' (' + r.importadas.join(', ') + ')' : ''}\n` +
    (r.jaExistiam.length ? `↪️ Já estavam no controle: ${r.jaExistiam.join(', ')}\n` : '') +
    (r.semLayout.length ? `⚠️ Importadas sem valores (descrição fora do template): ${r.semLayout.join(', ')}\n` : '') +
    (r.erros.length ? `❌ Não encontradas: ${r.erros.join(', ')}` : ''),
    ui.ButtonSet.OK);
}

function importarIssuesParaControle_(entrada) {
  const res = { importadas: [], jaExistiam: [], semLayout: [], erros: [] };
  const chaves = (entrada.toUpperCase().match(/\b[A-Z][A-Z0-9_]+-\d+\b/g) || []);
  const pareceJql = /[=~]|\border\s+by\b|\bin\s*\(/i.test(entrada);
  const jql = pareceJql ? entrada : `key in (${[...new Set(chaves)].join(',')})`;
  if (!pareceJql && !chaves.length) throw new Error('Nenhuma chave de issue reconhecida.');

  let issues = [];
  try {
    const r = chamarJiraApiCriacao_('/rest/api/2/search', 'post', {
      jql: jql, maxResults: 200, validateQuery: false,
      fields: ['summary', 'description', 'created', 'reporter', 'status']
    });
    issues = r.issues || [];
  } catch (e) {
    throw new Error('Erro na busca do Jira: ' + e.message);
  }
  if (!pareceJql) {
    const achadas = issues.map(i => i.key);
    chaves.filter(k => achadas.indexOf(k) === -1).forEach(k => res.erros.push(k));
  }

  const aba = abaControle_();
  issues.forEach(is => {
    if (linhaControlePorIssue_(aba, is.key)) { res.jaExistiam.push(is.key); return; }
    const f = is.fields || {};
    const tipo = /cria[çc][ãa]o de agenda/i.test(f.summary || '') ? 'Agenda'
      : /projeto interno/i.test(f.summary || '') ? 'Projeto Interno' : 'Outro';
    const resumo = extrairResumoDaDescricao_(tipo, f.description || '', f.summary || '');
    const solicitante = (f.reporter && (f.reporter.name || f.reporter.displayName)) || '';
    const criada = f.created ? new Date(f.created) : '';
    registrarLinhaControle_(tipo, is.key, `${CONFIG.JIRA_BASE_URL}/browse/${is.key}`, resumo, solicitante, criada);
    res.importadas.push(is.key);
    if (!resumo.encontrouLayout) res.semLayout.push(is.key);
  });
  return res;
}

/** Lê as tabelas wiki da descrição e devolve o mesmo resumo usado no controle. */
function extrairResumoDaDescricao_(tipo, descricao, titulo) {
  const linhas = linhasTabelaWiki_(descricao);
  const campo = {};          // rótulo normalizado -> valor (Tabela 01)
  const t02 = [];            // linhas com várias células (Tabela 02)
  linhas.forEach(cel => {
    if (cel.length === 2) campo[normRotulo_(cel[0])] = cel[1];
    if (cel.length >= 9) t02.push(cel);
  });
  const pega = (...rotulos) => {
    for (const r of rotulos) { const k = Object.keys(campo).find(c => c.indexOf(r) === 0); if (k && campo[k]) return campo[k]; }
    return '';
  };
  const valor = s => ehPlaceholder_(s) ? '' : s;
  const data = s => { const m = String(s || '').match(/(\d{2})\/(\d{2})\/(\d{4})/); return m && m[3] !== '9999' ? new Date(+m[3], +m[2] - 1, +m[1]) : ''; };
  const numero = s => {
    if (ehPlaceholder_(s)) return 0;
    let t = String(s || '').replace(/[^\d.,-]/g, '');
    if (t.indexOf(',') === -1 && /^\d{1,3}(\.\d{3})+$/.test(t)) t = t.replace(/\./g, '');   // "1.440" = mil e quarenta
    return num_(t);
  };

  const r = { codigoProjeto: '', projeto: '', analista: '', squad: '', de: '', ate: '',
    valorHora: '', dias: 0, horas: 0, valorTotal: 0, encontrouLayout: linhas.length > 3 };

  if (tipo === 'Agenda') {
    r.codigoProjeto = valor(pega('codigo do projeto')).toUpperCase();
    r.projeto = valor(pega('nome do projeto'));
    const consult = valor(pega('nome da consultoria'));
    r.analista = [valor(pega('nome completo do recurso')), consult ? `(${consult})` : ''].filter(Boolean).join(' ');
    r.squad = valor(pega('nome do squad'));
    r.de = data(pega('alocacao (periodo - de)', 'alocacao (periodo-de)'));
    r.ate = data(pega('alocacao (periodo - ate)', 'alocacao (periodo-ate)'));
    r.valorHora = numero(pega('valor hora negociado')) || '';
    const linhaT = rot => t02.find(c => normRotulo_(c[0]).indexOf(rot) === 0);
    const total = c => c ? numero(c[c.length - 1]) : 0;
    r.dias = total(linhaT('dias de alocacao'));
    r.horas = total(linhaT('horas de alocacao'));
    r.valorTotal = total(linhaT('valor total da alocacao')) || Math.round(r.horas * (r.valorHora || 0) * 100) / 100;
  } else {
    r.projeto = valor(pega('nome do projeto')) ||
      String(titulo).replace(/^.*nome do projeto:\s*/i, '').trim();
    const tl = valor(pega('team lead'));
    r.analista = tl ? 'TL: ' + tl : '';
    r.squad = valor(pega('squad responsavel'));
    r.de = data(pega('data inicio'));
    r.ate = data(pega('data fim'));
    let horas = 0, valorHoras = 0, total = 0, dias = 0;
    t02.filter(c => CONFIG_PI.MESES.indexOf(normRotulo_(c[0]).toUpperCase()) !== -1).forEach(c => {
      dias += numero(c[2]); horas += numero(c[4]); valorHoras += numero(c[6]); total += numero(c[8]);
    });
    r.dias = dias;
    r.horas = horas || numero(pega('quantidade de horas total'));
    r.valorTotal = total;
    r.valorHora = horas ? Math.round(valorHoras / horas * 100) / 100 : '';
  }
  return r;
}

/** Converte as tabelas do wiki markup em listas de células (texto limpo). */
function linhasTabelaWiki_(desc) {
  const brutas = [];
  String(desc || '').replace(/\r/g, '').split('\n').forEach(l => {
    const t = l.trim();
    if (t.indexOf('|') === 0) brutas.push(t);
    else if (t && brutas.length && !/\|\s*$/.test(brutas[brutas.length - 1]) && !/^h\d\./.test(t)) {
      brutas[brutas.length - 1] += ' \\\\ ' + t;   // célula que continua na linha seguinte
    }
  });
  return brutas.map(l => {
    const semLinks = l.replace(/\[([^\]|]*)\|[^\]]*\]/g, '$1');
    const protegido = semLinks.replace(/\\\|/g, '\u0001');
    return protegido.replace(/^\|\|?/, '').replace(/\|\|?\s*$/, '').split(/\|\|?/)
      .map(c => limparCelulaWiki_(c.replace(/\u0001/g, '|')));
  });
}

function limparCelulaWiki_(c) {
  return String(c)
    .replace(/\{color(:[^}]*)?\}/g, '')
    .replace(/\{\{|\}\}/g, '')
    .replace(/\\\\/g, ' ')
    .replace(/\\(.)/g, '$1')
    .replace(/[*+_]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function normRotulo_(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/:\s*$/, '').replace(/\s+/g, ' ').trim();
}

/** Valores de exemplo deixados no template (XXXX, 999, 99/99/9999). */
function ehPlaceholder_(s) {
  const t = String(s || '').trim();
  return !t || /x{4,}/i.test(t) || /x\s*-{3,}/i.test(t) || /^[R$\s]*9{2,}[9.,\s]*$/.test(t) || /99\/99\/9999/.test(t);
}

function buscarComentariosIssue_(key) {
  const r = chamarJiraApiCriacao_(`/rest/api/2/issue/${encodeURIComponent(key)}/comment?maxResults=200`, 'get');
  return r.comments || [];
}

/************************************************************
 * DIAGNÓSTICO: mostra o que o Jira devolve para a linha selecionada
 ************************************************************/
function diagnosticarLinhaControle() {
  const ui = SpreadsheetApp.getUi();
  const ss = planilha_();
  const aba = ss.getActiveSheet();
  if (aba.getName() !== CONFIG_CT.ABA) {
    ui.alert('Abra a aba Controle_Solicitacoes e selecione a linha da issue.'); return;
  }
  const linha = aba.getActiveRange().getRow();
  if (linha < 2) { ui.alert('Selecione uma linha de issue (a partir da linha 2).'); return; }
  const key = aba.getRange(linha, C_CT['Issue']).getDisplayValue().trim();
  if (!key) { ui.alert('A linha selecionada não tem issue.'); return; }

  let texto = `Issue: ${key}\n`;
  try {
    const is = chamarJiraApiCriacao_(`/rest/api/2/issue/${encodeURIComponent(key)}?fields=status`, 'get');
    texto += `Status no Jira: ${is.fields.status.name}\n`;
    const coms = buscarComentariosIssue_(key);
    texto += `Comentários: ${coms.length}\n\n`;
    coms.slice(-5).forEach((c, i) => {
      const autor = (c.author && c.author.displayName) || '?';
      const corpo = String(c.body || '').replace(/\s+/g, ' ').substring(0, 220);
      const achou = CONFIG_CT.REGEX_SOLICITACAO.test(c.body || '') ? '✅ ' : '— ';
      texto += `${achou}${autor}: ${corpo}\n\n`;
    });
    const it = extrairInfoIssue_({ fields: { status: is.fields.status, comment: { comments: coms } } });
    texto += it.numero ? `➡️ Nº encontrado: ${it.numero}${it.link ? '\nLink: ' + it.link : ''}`
      : '➡️ Nenhum comentário no formato "Solicitação NNNNN iniciada…" ainda.';
  } catch (e) {
    texto += '❌ Erro ao consultar o Jira: ' + e.message;
  }
  ui.alert('🔍 Diagnóstico da linha ' + linha, texto, ui.ButtonSet.OK);
}

/************************************************************
 * ATUALIZAÇÃO AUTOMÁTICA (opcional): a cada 1 hora
 ************************************************************/
function ativarAtualizacaoAutomaticaControle() {
  desativarAtualizacaoAutomaticaControle(true);
  ScriptApp.newTrigger('gatilhoAtualizarControle').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('⏰ Atualização automática ativada',
    'O Controle será atualizado a cada 1 hora (status e nº de solicitação).', SpreadsheetApp.getUi().ButtonSet.OK);
}

function desativarAtualizacaoAutomaticaControle(silencioso) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'gatilhoAtualizarControle')
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (silencioso !== true) {
    SpreadsheetApp.getUi().alert('Atualização automática desativada.');
  }
}

function gatilhoAtualizarControle() {
  try { importarDosRascunhos_(); atualizarControleSilencioso_(); }
  catch (e) { Logger.log('Falha na atualização automática do Controle: ' + e.message); }
}
