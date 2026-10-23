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

function obterConfigAgenda() {
  const meta = obterMetaTemplateAG_();
  return {
    projetoPadrao: CONFIG_AG.PROJETO_JIRA_PADRAO,
    epicoPadrao: CONFIG_AG.EPICO_FIXO || meta.epicKey || '',
    prioridade: meta.priorityName || 'Média',
    template: CONFIG_AG.ISSUE_TEMPLATE,
    meses: CONFIG_PI.MESES
  };
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
    const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ABA_LOG)
      || SpreadsheetApp.getActiveSpreadsheet().insertSheet(CONFIG.ABA_LOG);
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
  if (!(d.meses || []).some(m => num_(m.horas) > 0)) faltando.push('Tabela 02 (ao menos um mês com horas)');
  if (faltando.length) throw new Error('Preencha os campos obrigatórios: ' + faltando.join(', '));

  if (String(d.dataAte) < String(d.dataDe)) throw new Error('A data "Até" é anterior à data "De".');
  if (d.email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(d.email).trim())) {
    throw new Error('E-mail do Recurso inválido.');
  }
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

  const instr = [
    vd('Preencha SEMPRE os campos de  ALOCAÇÃO (PERÍODO - DE) e ALOCAÇÃO (PERÍODO - ATÉ).'),
    vd('Informar este quadro de DIAS ESPECÍFICOS apenas se precisar determinar DIAS DE ALOCAÇÕES ALTERNADOS dentro de cada mês dentro do período DE-ATÉ.'),
    vd('Se as alocações forem sequenciais, basta informar os quadros de ALOCAÇÃO (PERÍODO - DE) e ALOCAÇÃO (PERÍODO - ATÉ), não sendo necessário o preenchimento deste quadro de DIAS ESPECÍFICOS A ALOCAR.')
  ];
  const esp = d.diasEspecificos || [];
  const linhasMes = CONFIG_PI.MESES.map((m, i) => {
    const dias = String(esp[i] || '').trim();
    return '*' + vd(m) + '*' + (dias ? ' ' + vd('- ' + esc_(dias)) : '');
  });
  L.push(lin_('Dias específicos a alocar:', instr.concat(linhasMes).join(' \\\\ ')));

  L.push(lin_('Não alocar nas seguintes datas:',
    String(d.naoAlocar || '').trim() ? vd(esc_(d.naoAlocar)) : ' '));
  L.push(lin_('E-mail do Recurso:', az(esc_(String(d.email || '').trim()))));
  L.push(lin_('Observação:', String(d.observacao || '').trim() ? vd(esc_(d.observacao)) : ' '));
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
    cDias.push(vd(fmtNum_(dias, 0)));
    cHoras.push(vd(fmtNum_(horas, 0)));
    cValor.push('*' + vd(fmtNum_(valor, 2)) + '*');
  });
  L.push('|*Meses da alocação:*|' + CONFIG_PI.MESES.map(m => `*${m}*`).join('|') + '|*Total*|');
  L.push('|*Dias de alocação:*|' + cDias.join('|') + '|' + az(fmtNum_(tDias, 0)) + '|');
  L.push('|*Horas de alocação:*|' + cHoras.join('|') + '|' + az(fmtNum_(tHoras, 0)) + '|');
  L.push('|*Valor Total da alocação:*|' + cValor.join('|') + '|*' + az(fmtNum_(tValor, 2)) + '*|');

  return L.join('\n');
}

/************************************************************
 * RASCUNHOS (aba 'Rascunhos_Agenda')
 ************************************************************/
const CAB_RASCUNHOS_GEN = ['ID', 'Atualizado em', 'Usuário', 'Identificação', 'Status', 'Issue',
  'Título (prévia)', 'Descrição (prévia wiki)', 'Dados (JSON)'];

function abaRascunhosGen_(nome) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
  const lock = LockService.getDocumentLock();
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
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAba);
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
      nome: String(x.r[3] || ''),
      status: String(x.r[4] || ''),
      issue: String(x.r[5] || '')
    }))
    .sort((a, b) => b.quando - a.quando)
    .slice(0, 100);
}

function carregarRascunhoGen_(nomeAba, id) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAba);
  const linha = aba ? linhaRascunhoPI_(aba, id) : 0;
  if (!linha) throw new Error('Rascunho não encontrado.');
  return JSON.parse(aba.getRange(linha, 9).getValue() || '{}');
}

function marcarRascunhoGen_(nomeAba, id, key, url) {
  try {
    const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAba);
    const linha = aba ? linhaRascunhoPI_(aba, id) : 0;
    if (!linha) return;
    aba.getRange(linha, 5).setValue('Criada');
    aba.getRange(linha, 6).setFormula(`=HYPERLINK("${url}"; "${key}")`);
  } catch (e) {}
}
