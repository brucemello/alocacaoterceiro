/************************************************************
 * AGENTE JIRA IA - SOLICITAÇÃO DE AGENDA (v1.0)
 * Cria a issue no layout do template CIBKPRTH-7191.
 * Título: Criação de Agenda - Projeto : <código> - Recurso: <nome>
 *
 * Reaproveita de ProjetoInterno.gs: formatação (esc_, cor_, lin_,
 * num_, fmtNum_, dataBR_), tipos/campos por projeto, criação
 * adaptável às telas e vínculo de épico.
 ************************************************************/

const CONFIG_AG = {
  PROJETO_JIRA_PADRAO: 'CIBKPRTH',
  ISSUE_TEMPLATE: 'CIBKPRTH-7191',
  // Épico: lido da própria template (Agendas de Terceiros/Totvers...). Preencha aqui só se quiser fixar.
  EPICO_FIXO: '',
  LABELS: [],                           // a template não usa rótulos
  TITULO: (d) => `Criação de Agenda - Projeto : ${String(d.codigoProjeto || '').trim().toUpperCase()} - Recurso: ${String(d.nomeRecurso || '').trim()}`,
  ABA_RASCUNHOS: 'Rascunhos_Agenda',
  INCLUIR_ORIENTACOES: true
};

/* ---------- Menu (já incluído no Principal.gs atualizado) ----------
    .addItem('📅 Solicitar Agenda', 'abrirTelaAgenda')
-------------------------------------------------------------------- */

function abrirTelaAgenda() {
  const html = HtmlService.createTemplateFromFile('TelaAgenda').evaluate()
    .setWidth(1100)
    .setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, '📅 Solicitação de Agenda');
}

/** Lê tipo, prioridade e épico da template 7191 (com cache de 6h). */
function obterMetaTemplateAG_() {
  const cache = CacheService.getScriptCache();
  const salvo = cache.get('AG_META_TEMPLATE');
  if (salvo) return JSON.parse(salvo);

  let meta = { issuetypeId: '', issuetypeName: '', priorityId: '', priorityName: '', epicKey: '' };
  try {
    let campoEpic = '';
    try { campoEpic = obterCampoEpicLink_(); } catch (e) {}
    const r = chamarJiraApiCriacao_(`/rest/api/2/issue/${CONFIG_AG.ISSUE_TEMPLATE}` +
      `?fields=issuetype,priority${campoEpic ? ',' + campoEpic : ''}`, 'get');
    const f = r.fields || {};
    meta = {
      issuetypeId: (f.issuetype && f.issuetype.id) || '',
      issuetypeName: (f.issuetype && f.issuetype.name) || '',
      priorityId: (f.priority && f.priority.id) || '',
      priorityName: (f.priority && f.priority.name) || '',
      epicKey: campoEpic && f[campoEpic] ? String(f[campoEpic]) : ''
    };
    cache.put('AG_META_TEMPLATE', JSON.stringify(meta), 21600);
  } catch (e) {
    Logger.log('Não consegui ler a template de agenda: ' + e.message);
  }
  return meta;
}

/** Configuração inicial do formulário: rápida, sem depender do Jira (usa o cache da template). */
function obterConfigAgenda() {
  let meta = null;
  try { const c = CacheService.getScriptCache().get('AG_META_TEMPLATE'); meta = c ? JSON.parse(c) : null; } catch (e) {}
  return {
    projetoPadrao: CONFIG_AG.PROJETO_JIRA_PADRAO,
    epicoPadrao: CONFIG_AG.EPICO_FIXO || (meta && meta.epicKey) || '',
    prioridade: (meta && meta.priorityName) || 'Média',
    template: CONFIG_AG.ISSUE_TEMPLATE,
    jiraUrl: CONFIG.JIRA_BASE_URL,
    meses: CONFIG_PI.MESES,
    metaPendente: !meta
  };
}

/** Lê a template no Jira (chamado em segundo plano pelo formulário). */
function obterMetaAgendaPainel() {
  const meta = obterMetaTemplateAG_();
  return { epico: CONFIG_AG.EPICO_FIXO || meta.epicKey || '', prioridade: meta.priorityName || 'Média', ok: !!meta.issuetypeId };
}

/** Tipos do projeto, com "Tarefa" (tipo da template de agenda) como padrão. */
function listarTiposProjetoAG(projeto) {
  projeto = String(projeto || '').trim().toUpperCase();
  if (!projeto) return { tipos: [], padrao: '' };
  const r = chamarJiraApiCriacao_(
    `/rest/api/2/issue/createmeta/${encodeURIComponent(projeto)}/issuetypes?maxResults=100`, 'get');
  const tipos = (r.values || r.issueTypes || []).filter(t => !t.subtask)
    .map(t => ({ id: String(t.id), nome: t.name }));
  let padrao = '';
  try { padrao = resolverTipoIssue_(projeto, obterMetaTemplateAG_()); } catch (e) {}
  return { tipos: tipos, padrao: String(padrao || '') };
}

/************************************************************
 * PRÉVIA E CRIAÇÃO
 ************************************************************/
function gerarPreviaAgenda(dados) {
  validarDadosAgenda_(dados);
  return { titulo: CONFIG_AG.TITULO(dados), descricao: montarDescricaoAgenda_(dados) };
}

function criarIssueAgenda(dados) {
  validarDadosAgenda_(dados);
  const meta = obterMetaTemplateAG_();
  const projeto = String(dados.projetoJira || CONFIG_AG.PROJETO_JIRA_PADRAO).trim().toUpperCase();
  const epico = String(dados.epico === undefined ? (CONFIG_AG.EPICO_FIXO || meta.epicKey) : dados.epico)
    .trim().toUpperCase();

  const fields = {
    project: { key: projeto },
    issuetype: { id: dados.tipoIssueId || resolverTipoIssue_(projeto, meta) },
    summary: CONFIG_AG.TITULO(dados).substring(0, 255),
    description: montarDescricaoAgenda_(dados)
  };
  if (meta.priorityId) fields.priority = { id: meta.priorityId };
  if (CONFIG_AG.LABELS.length) fields.labels = CONFIG_AG.LABELS;

  aplicarCamposObrigatoriosProjeto_(fields, projeto, fields.issuetype.id, dados.camposProjeto);

  const pendentes = {};
  const permitidos = obterCamposPermitidos_(projeto, fields.issuetype.id);
  if (permitidos) {
    Object.keys(fields).forEach(k => {
      if (['project', 'issuetype', 'summary'].indexOf(k) === -1 && !permitidos[k]) {
        pendentes[k] = fields[k]; delete fields[k];
      }
    });
  }

  const resp = criarComRetentativa_(fields, pendentes);
  if (!resp.key) throw new Error('O Jira não retornou a chave da issue criada.');
  const avisos = aplicarCamposPendentes_(resp.key, pendentes);
  const status = (epico ? vincularEpico_(resp.key, epico) : 'Sem épico informado') +
    (avisos.length ? '<br>⚠️ Não aplicado (fora das telas do projeto): ' + avisos.join(', ') : '');

  const url = `${CONFIG.JIRA_BASE_URL}/browse/${resp.key}`;
  try {
    const aba = planilha_().getSheetByName(CONFIG.ABA_LOG)
      || planilha_().insertSheet(CONFIG.ABA_LOG);
    aba.appendRow([new Date(), Session.getActiveUser().getEmail() || 'desconhecido',
      'CRIAR_AGENDA', resp.key, projeto, CONFIG_AG.TITULO(dados), status.replace(/<br>/g, ' ')]);
  } catch (e) {}

  let rascunhoId = dados.rascunhoId || '';
  try {
    rascunhoId = salvarRascunhoAG(dados, rascunhoId).id;
    marcarRascunhoGen_(CONFIG_AG.ABA_RASCUNHOS, rascunhoId, resp.key, url);
  } catch (e) {}
  try { registrarNoControle_('Agenda', resp.key, url, dados); } catch (e) {}

  return { key: resp.key, url: url, epico: status, rascunhoId: rascunhoId };
}

/************************************************************
 * VALIDAÇÃO
 ************************************************************/
function validarDadosAgenda_(d) {
  const obrig = {
    codigoRecurso: 'Código Recurso (PSA)', nomeRecurso: 'Nome Completo do Recurso',
    codigoProjeto: 'Código do Projeto (PSA)', nomeProjeto: 'Nome do Projeto (PSA)',
    squad: 'Nome do SQUAD', papel: 'Papel/Categoria do Recurso', dataDe: 'Alocação (Período - De)',
    dataAte: 'Alocação (Período - Até)', email: 'E-mail do Recurso'
  };
  const faltando = Object.keys(obrig).filter(k => !String(d[k] || '').trim()).map(k => obrig[k]);
  if (String(d.consultoria || '').trim() && !(num_(d.valorHora) > 0)) {
    faltando.push('Valor hora negociado (obrigatório para Consultores Externos)');
  }
  if (!d.naoPreencherTabela02 && !(d.meses || []).some(m => num_(m.horas) > 0)) faltando.push('Tabela 02 (ao menos um mês com horas)');
  if (d.propostaExcecao) {
    if (!(num_(d.valorExcecao) > 0 || num_(d.valorHora) > 0)) faltando.push('Valor da Proposta de Exceção');
    if (!String(d.justificativaExcecao || '').trim()) faltando.push('Justificativa da Proposta de Exceção');
  }
  if (faltando.length) throw new Error('Preencha os campos obrigatórios: ' + faltando.join(', '));

  if (String(d.dataAte) < String(d.dataDe)) throw new Error('A data "Até" é anterior à data "De".');
  if (d.email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(d.email).trim())) {
    throw new Error('E-mail do Recurso inválido.');
  }
}

/** Linhas da Observação: Proposta de Exceção de Valor (quando houver) + observação livre. */
function textoObservacaoAgenda_(d) {
  const linhas = [];
  if (d.propostaExcecao) {
    const valor = num_(d.valorExcecao) || num_(d.valorHora);
    let t = `É necessária a montagem da Proposta de Exceção de Valor de Alocação no valor de R$ ${fmtNum_(valor, 2)}/hora`;
    if (num_(d.valorPapel) > 0) t += `, acima do valor de R$ ${fmtNum_(d.valorPapel, 2)}/hora cadastrado para o papel ${String(d.papel || '').trim()}`.replace(/ $/, '');
    linhas.push(t + '.');
    linhas.push('Justificativa da exceção: ' + String(d.justificativaExcecao || '').trim());
  }
  if (String(d.observacao || '').trim()) linhas.push(String(d.observacao).trim());
  return linhas;
}

/************************************************************
 * DESCRIÇÃO (wiki markup) NO LAYOUT DO CIBKPRTH-7191
 ************************************************************/
function montarDescricaoAgenda_(d) {
  const C = CONFIG_PI.CORES;
  const az = t => cor_(C.AZUL, t);
  const azB = t => String(t || '').trim() ? '*' + cor_(C.AZUL, t) + '*' : ' ';
  const vd = t => cor_(C.VERDE, t);
  const lj = t => cor_(C.LARANJA, t);
  const L = [];

  if (CONFIG_AG.INCLUIR_ORIENTACOES) {
    L.push('*Orientações Gerais para preenchimento do Template*:');
    L.push('* Campos em *AZUL* - Preenchimento obrigatório');
    L.push('* Campos em *VERDE* - Se não houver informações pertinentes a eles, deixar em branco');
    L.push('* Campos em *LARANJA* - Preenchimento obrigatórios para os *Projetos SFA (TOTVS CRM)*');
    L.push('* Campos em *VERMELHO* - Não alterá-los.');
    L.push('');
    L.push('É obrigatório preencher os dados das *Tabelas 01* e *02*, conforme orientações das cores acima');
    L.push('');
  }

  // ---------- Tabela 01 ----------
  L.push('h3. +Tabela 01 - Dados da Alocação+');
  L.push('|*Descrição:*|*Solicito a criação da alocação conforme dados do Projeto e Recurso abaixo*|');
  L.push(lin_('Código Recurso (PSA):', azB(esc_(d.codigoRecurso))));
  L.push(lin_('Nome Completo do Recurso:', azB(esc_(d.nomeRecurso))));
  L.push(lin_('Código do Projeto (PSA):', azB(esc_(String(d.codigoProjeto || '').toUpperCase()))));
  L.push(lin_('Nome do Projeto (PSA):', azB(esc_(d.nomeProjeto))));
  L.push(lin_('Nome da Consultoria:', String(d.consultoria || '').trim() ? vd(esc_(d.consultoria)) : ' '));
  L.push(lin_('Nome do SQUAD:', azB(esc_(d.squad))));
  L.push(lin_('Valor hora negociado da alocação:',
    num_(d.valorHora) > 0 ? vd('{{R$ ' + fmtNum_(d.valorHora, 2) + '}}') : ' '));
  L.push(lin_('Papel/Categoria do Recurso (PSA):', azB(esc_(d.papel))));
  L.push(lin_('Alocação (Período - De)', az(dataBR_(d.dataDe))));
  L.push(lin_('Alocação (Período - Até)', az(dataBR_(d.dataAte))));

  L.push(`|${lj('*Para TOTVS CRM*')}| |`);
  L.push(`|      ${lj('*Issue (Épico)*')}|${String(d.crmEpico || '').trim() ? lj(esc_(d.crmEpico)) : ' '}|`);
  L.push(`|      ${lj('*Issue (Subtarefa) / Data / Horas*')}|${String(d.crmSubtarefa || '').trim() ? lj(esc_(d.crmSubtarefa)) : ' '}|`);

  // Dias específicos: só os meses informados; em branco quando a alocação é sequencial (pedido do gestor)
  const esp = d.semDiasEspecificos ? [] : (d.diasEspecificos || []);
  const linhasMes = CONFIG_PI.MESES.map((m, i) => {
    const dias = String(esp[i] || '').trim();
    return dias ? '*' + vd(m) + '* ' + vd('- ' + esc_(dias)) : null;
  }).filter(Boolean);
  L.push(lin_('Dias específicos a alocar:', linhasMes.length ? linhasMes.join(' \\\\ ') : ' '));

  L.push(lin_('Não alocar nas seguintes datas:',
    String(d.naoAlocar || '').trim() ? vd(esc_(d.naoAlocar)) : ' '));
  L.push(lin_('E-mail do Recurso:', az(esc_(String(d.email || '').trim()))));
  const obs = textoObservacaoAgenda_(d);
  L.push(lin_('Observação:', obs.length ? obs.map(t => vd(esc_(t))).join(' \\\\ ') : ' '));
  L.push('');

  // ---------- Tabela 02 ----------
  L.push('h3. +Tabela 02 - Dados Financeiros da Alocação+');
  const vh = num_(d.valorHora);
  let tDias = 0, tHoras = 0, tValor = 0;
  const cDias = [], cHoras = [], cValor = [];
  CONFIG_PI.MESES.forEach((m, i) => {
    const x = (d.meses || [])[i] || {};
    const dias = num_(x.dias), horas = num_(x.horas), valor = horas * vh;
    tDias += dias; tHoras += horas; tValor += valor;
    const vazio = d.naoPreencherTabela02 || (d.mesesVaziosEmBranco !== false && !dias && !horas);   // em branco na issue
    cDias.push(vazio ? ' ' : vd(fmtNum_(dias, 0)));
    cHoras.push(vazio ? ' ' : vd(fmtNum_(horas, 0)));
    cValor.push(vazio ? ' ' : '*' + vd(fmtNum_(valor, 2)) + '*');
  });
  L.push('|*Meses da alocação:*|' + CONFIG_PI.MESES.map(m => `*${m}*`).join('|') + '|*Total*|');
  const tot = (txt, neg) => d.naoPreencherTabela02 ? ' ' : (neg ? '*' + az(txt) + '*' : az(txt));
  L.push('|*Dias de alocação:*|' + cDias.join('|') + '|' + tot(fmtNum_(tDias, 0)) + '|');
  L.push('|*Horas de alocação:*|' + cHoras.join('|') + '|' + tot(fmtNum_(tHoras, 0)) + '|');
  L.push('|*Valor Total da alocação:*|' + cValor.join('|') + '|' + tot(fmtNum_(tValor, 2), true) + '|');

  return L.join('\n');
}

/************************************************************
 * RASCUNHOS (aba 'Rascunhos_Agenda')
 ************************************************************/
const CAB_RASCUNHOS_GEN = ['ID', 'Atualizado em', 'Usuário', 'Identificação', 'Status', 'Issue',
  'Título (prévia)', 'Descrição (prévia wiki)', 'Dados (JSON)'];

function abaRascunhosGen_(nome) {
  const ss = planilha_();
  let aba = ss.getSheetByName(nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
    aba.getRange(1, 1, 1, CAB_RASCUNHOS_GEN.length).setValues([CAB_RASCUNHOS_GEN])
      .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
    aba.setFrozenRows(1);
    aba.setColumnWidths(7, 3, 300);
  }
  return aba;
}

function salvarRascunhoAG(dados, rascunhoId) {
  let titulo = '', descricao = '';
  try { titulo = CONFIG_AG.TITULO(dados); descricao = montarDescricaoAgenda_(dados); } catch (e) {}
  const ident = [dados.nomeRecurso, dados.codigoProjeto].filter(s => String(s || '').trim()).join(' · ');
  return salvarRascunhoGen_(CONFIG_AG.ABA_RASCUNHOS, dados, rascunhoId, ident || '(sem identificação)', titulo, descricao);
}

function listarRascunhosAG() { return listarRascunhosGen_(CONFIG_AG.ABA_RASCUNHOS); }
function carregarRascunhoAG(id) { return carregarRascunhoGen_(CONFIG_AG.ABA_RASCUNHOS, id); }

function salvarRascunhoGen_(nomeAba, dados, rascunhoId, ident, titulo, descricao) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const aba = abaRascunhosGen_(nomeAba);
    const id = rascunhoId || Utilities.getUuid().substring(0, 8);
    const agora = new Date();
    let linha = linhaRascunhoPI_(aba, id);
    const status = linha ? aba.getRange(linha, 5).getValue() : '';
    const issue = linha ? (aba.getRange(linha, 6).getFormula() || aba.getRange(linha, 6).getValue()) : '';
    if (!linha) linha = aba.getLastRow() + 1;
    aba.getRange(linha, 1, 1, CAB_RASCUNHOS_GEN.length).setValues([[
      id, agora, Session.getActiveUser().getEmail() || 'desconhecido', ident,
      status || 'Rascunho', issue, titulo, descricao, JSON.stringify(dados)]]);
    aba.getRange(linha, 2).setNumberFormat('dd/MM/yyyy HH:mm');
    return { id: id, quando: Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM HH:mm') };
  } finally {
    lock.releaseLock();
  }
}

function listarRascunhosGen_(nomeAba) {
  const esc = escopoTime_('');
  return listarRascunhosTodos_(nomeAba).filter(r => esc.filtro(timeDoUsuario_(r.email))).slice(0, 100);
}

function listarRascunhosTodos_(nomeAba) {
  const aba = planilha_().getSheetByName(nomeAba);
  if (!aba || aba.getLastRow() < 2) return [];
  const n = aba.getLastRow() - 1;
  const txt = aba.getRange(2, 1, n, 6).getDisplayValues();
  const datas = aba.getRange(2, 2, n, 1).getValues();
  return txt.map((r, i) => ({ r: r, data: datas[i][0] }))
    .filter(x => x.r[0])
    .map(x => ({
      id: String(x.r[0]),
      quando: x.data instanceof Date ? x.data.getTime() : 0,
      quandoTxt: x.r[1],
      usuario: String(x.r[2] || '').split('@')[0],
      email: String(x.r[2] || ''),
      nome: String(x.r[3] || ''),
      status: String(x.r[4] || ''),
      issue: String(x.r[5] || '')
    }))
    .sort((a, b) => b.quando - a.quando)
    .slice(0, 300);
}

function carregarRascunhoGen_(nomeAba, id) {
  const aba = planilha_().getSheetByName(nomeAba);
  const linha = aba ? linhaRascunhoPI_(aba, id) : 0;
  if (!linha) throw new Error('Rascunho não encontrado.');
  return JSON.parse(aba.getRange(linha, 9).getValue() || '{}');
}

function marcarRascunhoGen_(nomeAba, id, key, url) {
  try {
    const aba = planilha_().getSheetByName(nomeAba);
    const linha = aba ? linhaRascunhoPI_(aba, id) : 0;
    if (!linha) return;
    aba.getRange(linha, 5).setValue('Criada');
    aba.getRange(linha, 6).setFormula(`=HYPERLINK("${url}"; "${key}")`);
  } catch (e) {}
}

/************************************************************
 * COMBOS DO FORMULÁRIO: projetos e squads a partir do Controle
 ************************************************************/
function listarProjetosParaAgenda() {
  const res = { projetos: [], squads: [] };
  const ss = planilha_();
  if (!ss.getSheetByName(CONFIG_CT.ABA)) return res;
  const aba = abaControle_();
  if (aba.getLastRow() < 2) return res;

  const esc = escopoTime_('');
  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, COLS_CT.length).getDisplayValues()
    .filter(l => esc.filtro(String(l[C_CT['Time'] - 1] || '').trim()));
  const g = (l, c) => String(l[C_CT[c] - 1] || '').trim();
  const cancelada = l => g(l, 'Status Controle') === 'Cancelado' || /cancel/i.test(g(l, 'Status Jira'));
  const porCodigo = {}, pendentes = [], squads = {};

  dados.forEach(l => {
    if (!g(l, 'Issue') || cancelada(l)) return;
    const squad = g(l, 'Squad');
    if (squad) squads[squad.toLowerCase()] = squads[squad.toLowerCase()] || squad;
    const codigo = g(l, 'Código Projeto (PSA)').toUpperCase();
    const item = { codigo: codigo, nome: g(l, 'Projeto'), squad: squad, issue: g(l, 'Issue'),
      tipo: g(l, 'Tipo'), status: g(l, 'Status Controle') };
    if (codigo) {
      // Projeto Interno tem prioridade sobre Agendas do mesmo código (nome e squad "oficiais")
      const atual = porCodigo[codigo];
      if (!atual || (atual.tipo !== 'Projeto Interno' && item.tipo === 'Projeto Interno')) porCodigo[codigo] = item;
    } else if (item.tipo === 'Projeto Interno') {
      pendentes.push(item);
    }
  });

  const ordenar = (a, b) => a.nome.localeCompare(b.nome, 'pt-BR');
  res.projetos = Object.keys(porCodigo).map(k => porCodigo[k]).sort(ordenar).concat(pendentes.sort(ordenar));
  res.squads = Object.keys(squads).map(k => squads[k]).sort((a, b) => a.localeCompare(b, 'pt-BR'));
  return res;
}

/************************************************************
 * EDIÇÃO DE UMA AGENDA JÁ CRIADA
 ************************************************************/
/** Dados do formulário para editar a issue: do rascunho, ou lidos da descrição no Jira. */
function carregarAgendaParaEdicao(key) {
  key = String(key || '').trim().toUpperCase();
  if (!/^[A-Z][A-Z0-9_]+-\d+$/.test(key)) throw new Error('Chave de issue inválida: ' + key);

  const rasc = planilha_().getSheetByName(CONFIG_AG.ABA_RASCUNHOS);
  if (rasc && rasc.getLastRow() > 1) {
    const linhas = rasc.getRange(2, 1, rasc.getLastRow() - 1, 9).getDisplayValues();
    for (let i = linhas.length - 1; i >= 0; i--) {
      if (String(linhas[i][5]).trim().toUpperCase() === key) {
        let dados = {};
        try { dados = JSON.parse(rasc.getRange(i + 2, 9).getValue() || '{}'); } catch (e) {}
        return { dados: dados, rascunhoId: linhas[i][0], origem: 'rascunho' };
      }
    }
  }

  const is = chamarJiraApiCriacao_(`/rest/api/2/issue/${encodeURIComponent(key)}?fields=summary,description,project,issuetype`, 'get');
  const f = is.fields || {};
  const dados = lerAgendaDaDescricao_(f.description || '');
  dados.projetoJira = (f.project && f.project.key) || CONFIG_AG.PROJETO_JIRA_PADRAO;
  dados.tipoIssueId = (f.issuetype && f.issuetype.id) || '';
  return { dados: dados, rascunhoId: '', origem: 'jira' };
}

/** Lê a descrição (layout 7191) de volta para os campos do formulário. */
function lerAgendaDaDescricao_(desc) {
  const linhas = linhasTabelaWiki_(desc);
  const campo = {};
  const t02 = [];
  linhas.forEach(c => {
    if (c.length === 2) campo[normRotulo_(c[0])] = c[1];
    if (c.length >= 13) t02.push(c);
  });
  const pega = (...rot) => { for (const r of rot) { const k = Object.keys(campo).find(x => x.indexOf(r) === 0); if (k) return ehPlaceholder_(campo[k]) ? '' : campo[k]; } return ''; };
  const iso = s => { const m = String(s || '').match(/(\d{2})\/(\d{2})\/(\d{4})/); return m && m[3] !== '9999' ? `${m[3]}-${m[2]}-${m[1]}` : ''; };
  const numero = s => { if (ehPlaceholder_(s)) return ''; const n = num_(String(s || '').replace(/[^\d.,]/g, '')); return n ? String(n).replace('.', ',') : ''; };

  const d = {
    codigoRecurso: pega('codigo recurso'), nomeRecurso: pega('nome completo do recurso'),
    codigoProjeto: pega('codigo do projeto').toUpperCase(), nomeProjeto: pega('nome do projeto'),
    consultoria: pega('nome da consultoria').replace(/\(para consultores externos\)/i, '').trim(),
    squad: pega('nome do squad'), valorHora: numero(pega('valor hora negociado')),
    papel: pega('papel/categoria', 'papel'), dataDe: iso(pega('alocacao (periodo - de)', 'alocacao (periodo-de)')),
    dataAte: iso(pega('alocacao (periodo - ate)', 'alocacao (periodo-ate)')),
    crmEpico: pega('issue (epico)'), crmSubtarefa: pega('issue (subtarefa)'),
    naoAlocar: pega('nao alocar'), email: pega('e-mail do recurso', 'email do recurso'),
    observacao: pega('observacao'), horasDia: '8', autoCalc: false, descontarFeriados: true,
    trazerFeriados: false, diasEspecificos: Array(12).fill(''), meses: []
  };
  if (/^dd\/mm/i.test(d.naoAlocar)) d.naoAlocar = '';

  // Proposta de exceção (texto gerado pelo próprio agente) volta para os campos próprios
  const mEx = d.observacao.match(/Proposta de Exce[çc][ãa]o de Valor de Aloca[çc][ãa]o no valor de R\$\s*([\d.,]+)/i);
  if (mEx) {
    d.propostaExcecao = true;
    d.valorExcecao = mEx[1];
    const mPapel = d.observacao.match(/acima do valor de R\$\s*([\d.,]+)\/hora/i);
    if (mPapel) d.valorPapel = mPapel[1];
    const mJus = d.observacao.match(/Justificativa da exce[çc][ãa]o:\s*(.*?)(?=\s{2,}|$)/i);
    d.justificativaExcecao = mJus ? mJus[1].trim() : '';
    d.observacao = d.observacao.replace(/É necessária a montagem da Proposta.*?(Justificativa da exce[çc][ãa]o:.*?)(?=\s{2,}|$)/i, '').trim();
  }

  // Dias específicos ("JAN - 03, 05")
  const esp = pega('dias especificos');
  CONFIG_PI.MESES.forEach((m, i) => {
    const r = new RegExp('\\b' + m + '\\b\\s*-\\s*([\\d\\s,]+)', 'i');
    const x = esp.match(r);
    if (x && /\d/.test(x[1])) d.diasEspecificos[i] = x[1].trim().replace(/,\s*$/, '');
  });

  d.semDiasEspecificos = !d.diasEspecificos.some(Boolean);
  d.mesesVaziosEmBranco = true;
  d.__lerTabela02 = true;

  // Tabela 02: dias e horas por mês
  const linhaT = rot => t02.find(c => normRotulo_(c[0]).indexOf(rot) === 0);
  const ld = linhaT('dias de alocacao'), lh = linhaT('horas de alocacao');
  d.meses = CONFIG_PI.MESES.map((m, i) => ({
    mes: m,
    dias: ld && !ehPlaceholder_(ld[i + 1]) && num_(ld[i + 1]) ? String(num_(ld[i + 1])) : '',
    horas: lh && !ehPlaceholder_(lh[i + 1]) && num_(lh[i + 1]) ? String(num_(String(lh[i + 1]).replace(/\./g, ''))) : ''
  }));
  if (d.__lerTabela02) { delete d.__lerTabela02; d.naoPreencherTabela02 = !d.meses.some(m => num_(m.horas) > 0); if (d.naoPreencherTabela02) d.autoCalc = true; }
  return d;
}

/** Regrava título e descrição da issue, comenta o que mudou e atualiza rascunho e Controle. */
function atualizarIssueAgenda(dados, key, motivo) {
  validarDadosAgenda_(dados);
  key = String(key || '').trim().toUpperCase();
  if (!/^[A-Z][A-Z0-9_]+-\d+$/.test(key)) throw new Error('Issue inválida para atualização.');

  let anterior = {};
  try { anterior = carregarAgendaParaEdicao(key).dados || {}; } catch (e) {}

  chamarJiraApiCriacao_(`/rest/api/2/issue/${encodeURIComponent(key)}`, 'put', {
    fields: { summary: CONFIG_AG.TITULO(dados).substring(0, 255), description: montarDescricaoAgenda_(dados) }
  });

  const mudancas = diferencasAgenda_(anterior, dados);
  let comentario = false, avisoComentario = '';
  try {
    const autor = (Session.getActiveUser().getEmail() || '').split('@')[0];
    const corpo = ['*Agenda atualizada* pelo Agente A.I.F.R.E.A.D' + (autor ? ' (' + autor + ')' : '') + '.',
      String(motivo || '').trim() ? '*Motivo:* ' + String(motivo).trim() : '',
      mudancas.length ? '*Alterações:*\n' + mudancas.map(m => `* ${m.campo}: ${m.de || '—'} → ${m.para || '—'}`).join('\n')
        : 'Descrição regravada no layout da template (sem mudança nos campos principais).'
    ].filter(Boolean).join('\n');
    chamarJiraApiCriacao_(`/rest/api/2/issue/${encodeURIComponent(key)}/comment`, 'post', { body: corpo });
    comentario = true;
  } catch (e) { avisoComentario = 'A issue foi atualizada, mas não consegui adicionar o comentário: ' + e.message; }

  const url = `${CONFIG.JIRA_BASE_URL}/browse/${key}`;
  let rascunhoId = dados.rascunhoId || '';
  try {
    rascunhoId = salvarRascunhoAG(dados, rascunhoId).id;
    marcarRascunhoGen_(CONFIG_AG.ABA_RASCUNHOS, rascunhoId, key, url);
  } catch (e) {}
  try { atualizarLinhaControle_('Agenda', key, url, dados); } catch (e) {}
  try {
    const aba = planilha_().getSheetByName(CONFIG.ABA_LOG) || planilha_().insertSheet(CONFIG.ABA_LOG);
    aba.appendRow([new Date(), Session.getActiveUser().getEmail() || 'desconhecido', 'ATUALIZAR_AGENDA', key, '',
      CONFIG_AG.TITULO(dados), mudancas.map(m => m.campo).join(', ')]);
  } catch (e) {}

  return { key: key, url: url, mudancas: mudancas, comentario: comentario, avisoComentario: avisoComentario, rascunhoId: rascunhoId };
}

function diferencasAgenda_(a, b) {
  const totais = d => (d.meses || []).reduce((s, m) => ({ dias: s.dias + num_(m.dias), horas: s.horas + num_(m.horas) }), { dias: 0, horas: 0 });
  const ta = totais(a), tb = totais(b);
  const valor = (d, t) => t.horas * num_(d.valorHora);
  const br = s => { const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? `${m[3]}/${m[2]}/${m[1]}` : String(s || ''); };
  const campos = [
    ['Código Recurso (PSA)', d => d.codigoRecurso], ['Recurso', d => d.nomeRecurso], ['E-mail', d => d.email],
    ['Código do Projeto (PSA)', d => String(d.codigoProjeto || '').toUpperCase()], ['Projeto', d => d.nomeProjeto],
    ['Consultoria', d => d.consultoria], ['Squad', d => d.squad], ['Papel', d => d.papel],
    ['Valor hora', d => num_(d.valorHora) ? 'R$ ' + fmtNum_(d.valorHora, 2) : ''],
    ['Período - De', d => br(d.dataDe)], ['Período - Até', d => br(d.dataAte)],
    ['Dias específicos', d => (d.diasEspecificos || []).map((x, i) => String(x || '').trim() ? CONFIG_PI.MESES[i] + ' ' + x : '').filter(Boolean).join('; ')],
    ['Não alocar', d => d.naoAlocar], ['Observação', d => textoObservacaoAgenda_(d).join(' ')],
    ['Total de dias', (d, t) => t.dias ? fmtNum_(t.dias, 0) : ''], ['Total de horas', (d, t) => t.horas ? fmtNum_(t.horas, 0) : ''],
    ['Valor total', (d, t) => valor(d, t) ? 'R$ ' + fmtNum_(valor(d, t), 2) : '']
  ];
  const lista = [];
  campos.forEach(([nome, fn]) => {
    const de = String(fn(a, ta) || '').trim(), para = String(fn(b, tb) || '').trim();
    if (de !== para) lista.push({ campo: nome, de: de.substring(0, 160), para: para.substring(0, 160) });
  });
  return lista;
}
