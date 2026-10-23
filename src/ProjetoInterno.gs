/************************************************************
 * AGENTE JIRA IA - SOLICITAÇÃO DE PROJETO INTERNO (v1.0)
 * Cria a issue no layout do template CIBKPRTH-7190.
 * O usuário informa só os campos AZUIS e VERDES; os campos
 * VERMELHOS e LARANJAS saem fixos, como no template.
 *
 * Depende de: CriarIssue.gs (chamarJiraApiCriacao_)
 *             CONFIG do código principal (JIRA_BASE_URL, ABA_LOG)
 ************************************************************/

const CONFIG_PI = {
  PROJETO_JIRA_PADRAO: 'CIBKPRTH',
  EPICO_PADRAO: 'CIBKPRTH-7181',
  ISSUE_TEMPLATE: 'CIBKPRTH-7190',
  TIPO_ISSUE: 'Tarefa',
  PRIORIDADE: 'Baixa',
  LABELS: ['Template'],

  // Título: o nome do projeto entra no lugar de "X--- NOME DO PROJETO ---X"
  TITULO_PREFIXO: 'Criação de [PROJETO INTERNO] - Projeto: D0000XXXXXXXX - Nome do Projeto: ',

  INCLUIR_ORIENTACOES: true,

  CORES: {
    AZUL: '#4c9aff',
    VERDE: '#00875a',
    VERMELHO: '#de350b',
    LARANJA: '#ff991f'
  },

  TIPOS_PROJETO: [
    'Arquitetura', 'Desenvolvimento de Produto', 'Empréstimo', 'Entrega SAAStização',
    'Interno - atividades internas', 'Interno - Criação de oferta', 'Interno - Reunião interna',
    'Investimento em cliente - Campanha Promocional', 'Investimento em cliente para retenção',
    'Investimento em cliente para venda', 'Investimento em produto', 'Perda', 'Treinamento'
  ],

  // Campos VERMELHOS / LARANJAS (fixos do template — "Não alterá-los")
  FIXOS: {
    BACKOFFICE: {
      titulo: 'Para Unid.DESENVOLVIMENTO BACKOFFICE PRO',
      Produto: 'Protheus',
      Cliente: '99887 00 - TOTVS - Desenvolvimento Backoffice Protheus',
      Segmento: '000999 - TOTVS (Segmento)',
      Atendimento: 'Atendimento interno',
      Unidade: '184 - DESENVOLVIMENTO BACKOFFICE PRO',
      GPP: 'RODRIGO DE ALMEIDA SARTORIO'
    },
    GESPLAN: {
      titulo: 'Para Unid.GESPLAN',
      Produto: 'Outros',
      Cliente: 'TFDGGH 00 - GESPLAN SA',
      Segmento: '000010 - SERVICOS ( Segmento )',
      Atendimento: 'Avenida Santos Dumont Nº 831',
      Unidade: '[ ] 249 - GESPLAN          [ ] 269 - GESPLAN PARTICIPATIVOS',
      GPP: 'GEOVANI BROCARDO LOPES'
    },
    CP_GP: 'ALEXANDRE DA COSTA',
    RESPONSAVEL_RESERVA: '<NÃO INFORMAR>',
    MODELO_CALENDARIO: 'Modelo de Trabalho Padrão'
  },

  MESES: ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']
};

/* ---------- Menu: adicione no onOpen() do código principal ----------
    .addItem('10. 🏗️ Solicitar Projeto Interno', 'abrirTelaProjetoInterno')
    .addItem('11. 🔎 Exportar template bruto (CIBKPRTH-7190)', 'exportarTemplateBrutoJira')
--------------------------------------------------------------------- */

function abrirTelaProjetoInterno() {
  const html = HtmlService.createTemplateFromFile('TelaProjetoInterno').evaluate()
    .setWidth(1050)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, '🏗️ Solicitação de Projeto Interno');
}

function obterConfigProjetoInterno() {
  return {
    projetoPadrao: CONFIG_PI.PROJETO_JIRA_PADRAO,
    epicoPadrao: CONFIG_PI.EPICO_PADRAO,
    tiposProjeto: CONFIG_PI.TIPOS_PROJETO,
    meses: CONFIG_PI.MESES,
    tituloPrefixo: CONFIG_PI.TITULO_PREFIXO,
    tipo: CONFIG_PI.TIPO_ISSUE,
    prioridade: CONFIG_PI.PRIORIDADE,
    labels: CONFIG_PI.LABELS.join(', ')
  };
}

function listarProjetosJira() {
  const lista = chamarJiraApiCriacao_('/rest/api/2/project', 'get') || [];
  return lista
    .map(p => ({ key: p.key, nome: p.name }))
    .sort((a, b) => a.key.localeCompare(b.key));
}

/************************************************************
 * PRÉVIA E CRIAÇÃO
 ************************************************************/
function gerarPreviaProjetoInterno(dados) {
  validarDadosProjetoInterno_(dados);
  return {
    titulo: montarTituloProjetoInterno_(dados),
    descricao: montarDescricaoProjetoInterno_(dados)
  };
}

function criarIssueProjetoInterno(dados) {
  validarDadosProjetoInterno_(dados);

  const projeto = String(dados.projetoJira || CONFIG_PI.PROJETO_JIRA_PADRAO).trim().toUpperCase();
  // Épico em branco no formulário = não vincular (útil para testes)
  const epico = String(dados.epico === undefined ? CONFIG_PI.EPICO_PADRAO : dados.epico).trim().toUpperCase();

  const meta = obterMetaTemplate_();
  const fields = {
    project: { key: projeto },
    issuetype: { id: dados.tipoIssueId || resolverTipoIssue_(projeto, meta) },
    priority: meta.priorityId ? { id: meta.priorityId } : { name: CONFIG_PI.PRIORIDADE },
    labels: CONFIG_PI.LABELS,
    summary: montarTituloProjetoInterno_(dados),
    description: montarDescricaoProjetoInterno_(dados)
  };

  // 0) Campos obrigatórios próprios do projeto (ex.: Componente(s))
  aplicarCamposObrigatoriosProjeto_(fields, projeto, fields.issuetype.id, dados.camposProjeto);

  // 1) Separa os campos que não estão na tela de criação deste projeto/tipo
  const pendentes = {};
  const permitidos = obterCamposPermitidos_(projeto, fields.issuetype.id);
  if (permitidos) {
    Object.keys(fields).forEach(k => {
      if (['project', 'issuetype', 'summary'].indexOf(k) === -1 && !permitidos[k]) {
        pendentes[k] = fields[k];
        delete fields[k];
      }
    });
  }

  // 2) Cria (Epic Link nunca vai aqui); se o Jira recusar algum campo, tira e tenta de novo
  const resp = criarComRetentativa_(fields, pendentes);
  if (!resp.key) throw new Error('O Jira não retornou a chave da issue criada.');

  // 3) Aplica os campos pendentes editando a issue já criada
  const avisos = aplicarCamposPendentes_(resp.key, pendentes);

  // 4) Vincula ao épico depois de criada
  const statusEpico = (epico ? vincularEpico_(resp.key, epico) : 'Sem épico informado') +
    (avisos.length ? '<br>⚠️ Não aplicado (fora das telas do projeto): ' + avisos.join(', ') : '');

  try {
    const aba = planilha_().getSheetByName(CONFIG.ABA_LOG)
      || planilha_().insertSheet(CONFIG.ABA_LOG);
    aba.appendRow([new Date(), Session.getActiveUser().getEmail() || 'desconhecido',
      'CRIAR_PROJETO_INTERNO', resp.key, projeto, fields.summary, statusEpico]);
  } catch (e) {}

  const url = `${CONFIG.JIRA_BASE_URL}/browse/${resp.key}`;
  let rascunhoId = dados.rascunhoId || '';
  try {
    rascunhoId = salvarRascunhoPI(dados, rascunhoId).id;
    marcarRascunhoCriadoPI_(rascunhoId, resp.key, url);
  } catch (e) {}
  try { registrarNoControle_('Projeto Interno', resp.key, url, dados); } catch (e) {}
  return { key: resp.key, url: url, epico: statusEpico, rascunhoId: rascunhoId };
}

/************************************************************
 * RASCUNHOS (aba 'Rascunhos_PI') — salvar, listar e reabrir
 ************************************************************/
const ABA_RASCUNHOS_PI = 'Rascunhos_PI';
const CAB_RASCUNHOS_PI = ['ID', 'Atualizado em', 'Usuário', 'Nome do Projeto', 'Status', 'Issue',
  'Título (prévia)', 'Descrição (prévia wiki)', 'Dados (JSON)'];

function abaRascunhosPI_() {
  const ss = planilha_();
  let aba = ss.getSheetByName(ABA_RASCUNHOS_PI);
  if (!aba) {
    aba = ss.insertSheet(ABA_RASCUNHOS_PI);
    aba.getRange(1, 1, 1, CAB_RASCUNHOS_PI.length).setValues([CAB_RASCUNHOS_PI])
      .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
    aba.setFrozenRows(1);
    aba.setColumnWidths(7, 3, 300);
  }
  return aba;
}

function linhaRascunhoPI_(aba, id) {
  if (!id || aba.getLastRow() < 2) return 0;
  const ids = aba.getRange(2, 1, aba.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

/** Salva (ou atualiza) o rascunho. Também grava a prévia, se os dados permitirem gerá-la. */
function salvarRascunhoPI(dados, rascunhoId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const aba = abaRascunhosPI_();
    const id = rascunhoId || Utilities.getUuid().substring(0, 8);
    let titulo = '', descricao = '';
    try {
      titulo = montarTituloProjetoInterno_(dados);
      descricao = montarDescricaoProjetoInterno_(dados);
    } catch (e) {}

    const agora = new Date();
    let linha = linhaRascunhoPI_(aba, id);
    const statusAtual = linha ? aba.getRange(linha, 5).getValue() : '';
    const issueAtual = linha ? aba.getRange(linha, 6).getFormula() || aba.getRange(linha, 6).getValue() : '';
    const valores = [[id, agora, Session.getActiveUser().getEmail() || 'desconhecido',
      dados.nomeProjeto || '(sem nome)', statusAtual || 'Rascunho', issueAtual,
      titulo, descricao, JSON.stringify(dados)]];

    if (!linha) linha = aba.getLastRow() + 1;
    aba.getRange(linha, 1, 1, CAB_RASCUNHOS_PI.length).setValues(valores);
    aba.getRange(linha, 2).setNumberFormat('dd/MM/yyyy HH:mm');
    return { id: id, quando: Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM HH:mm') };
  } finally {
    lock.releaseLock();
  }
}

function listarRascunhosPI() {
  const aba = planilha_().getSheetByName(ABA_RASCUNHOS_PI);
  if (!aba || aba.getLastRow() < 2) return [];
  return aba.getRange(2, 1, aba.getLastRow() - 1, 6).getDisplayValues()
    .map((r, i) => ({ r: r, data: aba.getRange(i + 2, 2).getValue() }))
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

function carregarRascunhoPI(id) {
  const aba = planilha_().getSheetByName(ABA_RASCUNHOS_PI);
  const linha = aba ? linhaRascunhoPI_(aba, id) : 0;
  if (!linha) throw new Error('Rascunho não encontrado.');
  return JSON.parse(aba.getRange(linha, 9).getValue() || '{}');
}

function marcarRascunhoCriadoPI_(id, key, url) {
  try {
    const aba = planilha_().getSheetByName(ABA_RASCUNHOS_PI);
    const linha = aba ? linhaRascunhoPI_(aba, id) : 0;
    if (!linha) return;
    aba.getRange(linha, 5).setValue('Criada');
    aba.getRange(linha, 6).setFormula(`=HYPERLINK("${url}"; "${key}")`);
  } catch (e) {}
}

/**
 * Vincula a issue ao épico. Tenta, em ordem:
 *  a) API Agile:  POST /rest/agile/1.0/epic/{epico}/issue
 *  b) Edição:     PUT  /rest/api/2/issue/{key} com o campo Epic Link
 * Nunca lança erro — a issue já existe; devolve o status para exibir ao usuário.
 */
function vincularEpico_(issueKey, epico) {
  const erros = [];
  try {
    chamarJiraApiCriacao_(`/rest/agile/1.0/epic/${encodeURIComponent(epico)}/issue`, 'post',
      { issues: [issueKey] });
    return `✅ Vinculada ao épico ${epico}`;
  } catch (e) { erros.push('Agile: ' + e.message); }

  try {
    const campo = obterCampoEpicLink_();
    chamarJiraApiCriacao_(`/rest/api/2/issue/${issueKey}`, 'put', { fields: { [campo]: epico } });
    return `✅ Vinculada ao épico ${epico}`;
  } catch (e) { erros.push('Edição: ' + e.message); }

  return `⚠️ Issue criada, mas NÃO foi vinculada ao épico ${epico}. Vincule manualmente. (${erros.join(' | ')})`;
}

/**
 * Lê o tipo e a prioridade da issue-template (CIBKPRTH-7190) e guarda os IDs em cache.
 * Usar IDs evita o problema de nomes traduzidos ("Tarefa" x "Task", "Baixa" x "Low").
 */
function obterMetaTemplate_() {
  const cache = CacheService.getScriptCache();
  const salvo = cache.get('PI_META_TEMPLATE');
  if (salvo) return JSON.parse(salvo);

  let meta = { issuetypeId: '', issuetypeName: '', priorityId: '' };
  try {
    const r = chamarJiraApiCriacao_(
      `/rest/api/2/issue/${CONFIG_PI.ISSUE_TEMPLATE}?fields=issuetype,priority`, 'get');
    meta = {
      issuetypeId: (r.fields.issuetype && r.fields.issuetype.id) || '',
      issuetypeName: (r.fields.issuetype && r.fields.issuetype.name) || '',
      priorityId: (r.fields.priority && r.fields.priority.id) || ''
    };
    cache.put('PI_META_TEMPLATE', JSON.stringify(meta), 21600);
  } catch (e) {
    Logger.log('Não consegui ler a issue-template: ' + e.message);
  }
  return meta;
}

/************************************************************
 * CRIAÇÃO ADAPTÁVEL ÀS TELAS DE CADA PROJETO
 ************************************************************/
const NOMES_CAMPOS_PI = { priority: 'Prioridade', labels: 'Rótulo', description: 'Descrição' };

/** Campos aceitos na tela de criação (projeto + tipo). null = não foi possível consultar. */
function obterCamposPermitidos_(projeto, tipoId) {
  try {
    const r = chamarJiraApiCriacao_(`/rest/api/2/issue/createmeta/${encodeURIComponent(projeto)}` +
      `/issuetypes/${encodeURIComponent(tipoId)}?maxResults=500`, 'get');
    const lista = r.values || r.fields || [];
    if (!lista.length) return null;
    const mapa = {};
    lista.forEach(c => { mapa[c.fieldId || c.key] = true; });
    return mapa;
  } catch (e) {
    return null;
  }
}

/** Cria a issue; se o Jira disser "cannot be set" para algum campo, move-o para pendentes e repete. */
function criarComRetentativa_(fields, pendentes) {
  for (let tentativa = 0; tentativa < 4; tentativa++) {
    try {
      return chamarJiraApiCriacao_('/rest/api/2/issue', 'post', { fields: fields });
    } catch (e) {
      const recusados = Object.keys(fields).filter(k =>
        ['project', 'issuetype', 'summary'].indexOf(k) === -1 &&
        new RegExp('\\b' + k + ':[^|]*(cannot be set|not on the appropriate screen)', 'i').test(e.message));
      if (!recusados.length) throw e;
      recusados.forEach(k => { pendentes[k] = fields[k]; delete fields[k]; });
    }
  }
  throw new Error('O Jira continuou recusando campos na criação.');
}

/** Edita a issue com os campos pendentes; devolve os nomes dos que não puderam ser aplicados. */
function aplicarCamposPendentes_(key, pendentes) {
  const chaves = Object.keys(pendentes);
  if (!chaves.length) return [];
  try {
    chamarJiraApiCriacao_(`/rest/api/2/issue/${key}`, 'put', { fields: pendentes });
    return [];
  } catch (e) {}
  const falharam = [];
  chaves.forEach(k => {
    try {
      chamarJiraApiCriacao_(`/rest/api/2/issue/${key}`, 'put', { fields: { [k]: pendentes[k] } });
    } catch (e) {
      falharam.push(NOMES_CAMPOS_PI[k] || k);
    }
  });
  return falharam;
}

/************************************************************
 * CAMPOS OBRIGATÓRIOS ESPECÍFICOS DE CADA PROJETO/TIPO
 * (ex.: Componente(s) no DTRIPRO)
 ************************************************************/
// Campos que o formulário/script já preenche por conta própria
const CAMPOS_JA_TRATADOS_PI = ['project', 'issuetype', 'summary', 'description', 'reporter',
  'priority', 'labels', 'attachment', 'issuelinks'];

function listarCamposObrigatoriosPI(projeto, tipoId) {
  projeto = String(projeto || '').trim().toUpperCase();
  if (!projeto || !tipoId) return [];
  const r = chamarJiraApiCriacao_(`/rest/api/2/issue/createmeta/${encodeURIComponent(projeto)}` +
    `/issuetypes/${encodeURIComponent(tipoId)}?maxResults=500`, 'get');
  let epicLink = '';
  try { epicLink = obterCampoEpicLink_(); } catch (e) {}

  return (r.values || r.fields || [])
    .filter(c => c.required && !c.hasDefaultValue &&
      CAMPOS_JA_TRATADOS_PI.indexOf(c.fieldId) === -1 && c.fieldId !== epicLink)
    .map(c => {
      const schema = c.schema || {};
      return {
        id: c.fieldId,
        nome: c.name,
        tipo: schema.type || 'string',
        itens: schema.items || '',
        multiplo: schema.type === 'array',
        opcoes: (c.allowedValues || []).map(v => ({
          id: String(v.id || v.key || ''),
          nome: v.name || v.value || v.key || String(v.id)
        }))
      };
    });
}

/** Converte o valor digitado/escolhido no formato que a API do Jira espera. */
function formatarValorCampoPI_(def, valor) {
  const vazio = v => v === undefined || v === null || String(v).trim() === '' ||
    (Array.isArray(v) && !v.length);
  if (vazio(valor)) return null;
  const lista = Array.isArray(valor) ? valor : String(valor).split(',').map(s => s.trim()).filter(Boolean);

  if (def.opcoes && def.opcoes.length) {
    return def.multiplo ? lista.map(id => ({ id: String(id) })) : { id: String(lista[0]) };
  }
  if (def.tipo === 'number') return num_(valor);
  if (def.tipo === 'user') return { name: String(valor).trim() };
  if (def.tipo === 'array') {
    return def.itens === 'user' ? lista.map(n => ({ name: n })) :
      def.itens === 'string' ? lista : lista.map(n => ({ name: n }));
  }
  return String(valor).trim();
}

/** Adiciona aos fields os campos obrigatórios do projeto preenchidos no formulário. */
function aplicarCamposObrigatoriosProjeto_(fields, projeto, tipoId, valores) {
  let defs = [];
  try { defs = listarCamposObrigatoriosPI(projeto, tipoId); } catch (e) { return; }
  const faltando = [];
  defs.forEach(def => {
    const v = formatarValorCampoPI_(def, (valores || {})[def.id]);
    if (v === null) faltando.push(def.nome);
    else fields[def.id] = v;
  });
  if (faltando.length) {
    throw new Error(`Preencha os campos obrigatórios do projeto ${projeto}: ${faltando.join(', ')}`);
  }
}

/** Lista os tipos do projeto para o formulário, indicando o padrão ("Tarefa") se existir. */
function listarTiposProjetoPI(projeto) {
  projeto = String(projeto || '').trim().toUpperCase();
  if (!projeto) return { tipos: [], padrao: '' };
  const r = chamarJiraApiCriacao_(
    `/rest/api/2/issue/createmeta/${encodeURIComponent(projeto)}/issuetypes?maxResults=100`, 'get');
  const tipos = (r.values || r.issueTypes || [])
    .filter(t => !t.subtask)
    .map(t => ({ id: String(t.id), nome: t.name }));
  let padrao = '';
  try { padrao = resolverTipoIssue_(projeto, obterMetaTemplate_()); } catch (e) {}
  return { tipos: tipos, padrao: String(padrao || '') };
}

/** Descobre o ID do tipo "Tarefa" válido no projeto escolhido. */
function resolverTipoIssue_(projeto, meta) {
  let tipos = [];
  try {
    const r = chamarJiraApiCriacao_(
      `/rest/api/2/issue/createmeta/${encodeURIComponent(projeto)}/issuetypes?maxResults=100`, 'get');
    tipos = (r.values || r.issueTypes || []).filter(t => !t.subtask);
  } catch (e) {
    if (meta.issuetypeId) return meta.issuetypeId;   // sem createmeta: usa o tipo da template
    throw new Error(`Não consegui ler os tipos de issue do projeto ${projeto}: ${e.message}`);
  }

  const porId = tipos.find(t => String(t.id) === String(meta.issuetypeId));
  if (porId) return porId.id;

  const nomes = [CONFIG_PI.TIPO_ISSUE, meta.issuetypeName, 'Tarefa', 'Task']
    .filter(Boolean).map(n => n.toLowerCase());
  const porNome = tipos.find(t => nomes.indexOf(String(t.name).toLowerCase()) !== -1);
  if (porNome) return porNome.id;

  throw new Error(`O projeto ${projeto} não tem o tipo "Tarefa". Tipos disponíveis: ` +
    tipos.map(t => t.name).join(', '));
}

/** Descobre o ID do campo "Epic Link" (varia por instalação) e guarda em cache. */
function obterCampoEpicLink_() {
  const props = PropertiesService.getScriptProperties();
  const cache = props.getProperty('JIRA_CAMPO_EPIC_LINK');
  if (cache) return cache;

  const campos = chamarJiraApiCriacao_('/rest/api/2/field', 'get') || [];
  const epic = campos.find(f => f.schema && f.schema.custom === 'com.pyxis.greenhopper.jira:gh-epic-link')
    || campos.find(f => /^epic link$/i.test(f.name || ''));
  if (!epic) throw new Error('Não encontrei o campo "Epic Link" no Jira.');

  props.setProperty('JIRA_CAMPO_EPIC_LINK', epic.id);
  return epic.id;
}

/** Salva a descrição bruta (wiki) da issue-template numa aba, para conferência do layout. */
function exportarTemplateBrutoJira() {
  const resp = chamarJiraApiCriacao_(
    `/rest/api/2/issue/${CONFIG_PI.ISSUE_TEMPLATE}?fields=summary,description`, 'get');
  const ss = planilha_();
  let aba = ss.getSheetByName('Template_Bruto');
  if (!aba) aba = ss.insertSheet('Template_Bruto');
  aba.clear();
  aba.getRange('A1:B1').setValues([['Título', 'Descrição (wiki markup)']]).setFontWeight('bold');
  aba.getRange('A2:B2').setValues([[resp.fields.summary || '', resp.fields.description || '']]);
  aba.setColumnWidth(2, 900);
  aba.getRange('B2').setWrap(true);
  SpreadsheetApp.getUi().alert('Template exportado',
    `A descrição bruta do ${CONFIG_PI.ISSUE_TEMPLATE} foi gravada na aba 'Template_Bruto'.`,
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/************************************************************
 * VALIDAÇÃO (campos AZUIS obrigatórios)
 ************************************************************/
function validarDadosProjetoInterno_(d) {
  const faltando = [];
  const obrig = {
    squad: 'Squad Responsável', gestor: 'Gestor Responsável (Tribe)', teamLeads: 'Team Lead(ers)',
    nomeProjeto: 'Nome do Projeto', dataInicio: 'Data Início', dataFim: 'Data Fim',
    centroCusto: 'Centro de Custo', classeValor: 'Classe de Valor', itemContabil: 'Item Contábil',
    objetivo: 'Objetivo do Serviço', resultadoEsperado: 'Resultado Final Esperado',
    cenarioAtual: 'Cenário Atual', horasTotal: 'Quantidade de horas total',
    recursosTotal: 'Quantidade total de Recursos'
  };
  Object.keys(obrig).forEach(k => { if (!String(d[k] || '').trim()) faltando.push(obrig[k]); });
  if (!d.tiposProjeto || !d.tiposProjeto.length) faltando.push('Tipo de Projeto');
  if (!d.custoForecast) faltando.push('Custo não considerado no FORECAST');
  if (!(d.valoresHora || []).some(v => num_(v.recursos) > 0 && num_(v.valorHora) > 0)) {
    faltando.push('Valor(es) hora por Recurso');
  }
  if (!(d.meses || []).some(m => num_(m.a) > 0 && num_(m.b) > 0)) {
    faltando.push('Tabela 02 (ao menos um mês)');
  }
  if (faltando.length) throw new Error('Preencha os campos obrigatórios: ' + faltando.join(', '));

  if (d.dataInicio && d.dataFim && String(d.dataFim) < String(d.dataInicio)) {
    throw new Error('A Data Fim é anterior à Data Início.');
  }
}

/************************************************************
 * MONTAGEM DO TÍTULO E DA DESCRIÇÃO (wiki markup do Jira)
 ************************************************************/
function montarTituloProjetoInterno_(d) {
  return (CONFIG_PI.TITULO_PREFIXO + String(d.nomeProjeto || '').trim()).substring(0, 255);
}

function montarDescricaoProjetoInterno_(d) {
  const C = CONFIG_PI.CORES;
  const F = CONFIG_PI.FIXOS;
  const az = t => cor_(C.AZUL, t);
  const vd = t => cor_(C.VERDE, t);
  const vm = t => cor_(C.VERMELHO, t);
  const lj = t => cor_(C.LARANJA, t);
  const L = [];

  // ---------- Orientações ----------
  if (CONFIG_PI.INCLUIR_ORIENTACOES) {
    L.push('*Orientações Gerais para preenchimento do Template*:');
    L.push('* Campos em *AZUL* - Preenchimento obrigatório');
    L.push('* Campos em *VERDE* - Se não houver informações pertinentes a eles, deixar em branco');
    L.push('* Campos em *LARANJA* - Preenchimento obrigatório para solicitar projetos para as Unidades GESPLAN');
    L.push('* Campos em *VERMELHO* - Não alterá-los.');
    L.push('');
    L.push('É obrigatório preencher os dados das *Tabelas 01* e *02*, conforme orientações das cores acima');
    L.push('');
    L.push('Caso seu Projeto não tenha a comprovação de um retorno financeiro para a Totvs, então deixe os dados da *Tabela 03* em branco, ou a exclua da sua _Issue_,');
    L.push('');
  }

  // ---------- TABELA 01 ----------
  L.push('h3. +TABELA 01 - Definições do Projeto+');
  L.push('||Descrição||Solicito a criação do Projeto Interno conforme abaixo||');
  L.push(lin_('Squad Responsável pelo Projeto', az(esc_(d.squad))));
  L.push(lin_('Gestor Responsável pelo Projeto (Tribe):', az(esc_(d.gestor))));
  L.push(lin_('Team Lead(ers) que irá(ão) solicitar alocações no Projeto:', az(esc_(d.teamLeads))));

  const marcados = d.tiposProjeto || [];
  const tipos = CONFIG_PI.TIPOS_PROJETO
    .map(t => az((marcados.indexOf(t) !== -1 ? '\\[ X \\] ' : '\\[ \\] ') + t))
    .join(' \\\\ ');
  L.push(lin_('Tipo de Projeto:', tipos));

  L.push(lin_('Nome do Projeto:', az(esc_(d.nomeProjeto))));
  L.push(lin_('Data Início:', az(dataBR_(d.dataInicio))));
  L.push(lin_('Data fim:', az(dataBR_(d.dataFim))));

  [F.BACKOFFICE, F.GESPLAN].forEach(bloco => {
    const ehGesplan = bloco === F.GESPLAN;
    L.push(lin_(bloco.titulo, ' '));
    ['Produto', 'Cliente', 'Segmento', 'Atendimento', 'Unidade', 'GPP'].forEach(campo => {
      let valor = esc_(bloco[campo]);
      valor = (ehGesplan && campo === 'Unidade') ? lj(valor) : vm(valor);
      L.push(`|      *${campo}:*|${valor}|`);
    });
  });

  L.push(lin_('CP/GP:', vm(esc_(F.CP_GP))));
  L.push(lin_('Responsável pela Reserva:', vm(esc_(F.RESPONSAVEL_RESERVA))));
  L.push(lin_('Modelo de Calendário:', vm(esc_(F.MODELO_CALENDARIO))));

  const sim = d.custoForecast === 'Sim', nao = d.custoForecast === 'Não';
  L.push(lin_('Custo não considerado no FORECAST:',
    az(`\\[${sim ? ' X ' : ' '}\\] Sim          \\[${nao ? ' X ' : ' '}\\] Não`)));

  L.push(lin_('Código do contrato faturável:', String(d.contratoFaturavel || '').trim() ? vd(esc_(d.contratoFaturavel)) : ' '));
  L.push(lin_('Nr Oportunidade (pré-venda):', String(d.oportunidade || '').trim() ? vd(esc_(d.oportunidade)) : ' '));

  L.push(lin_('Centro de Custo:', az(esc_(d.centroCusto))));
  L.push(lin_('Classe de Valor:', az(esc_(d.classeValor))));
  L.push(lin_('Item Contábil:', az(esc_(d.itemContabil))));
  L.push(lin_('Objetivo do Serviço:', az(esc_(d.objetivo))));
  L.push(lin_('Resultado Final Esperado:', az(esc_(d.resultadoEsperado))));
  L.push(lin_('Cenário Atual do Projeto:', az(esc_(d.cenarioAtual))));
  L.push(lin_('Quantidade de horas total do Projeto:', az(fmtNum_(d.horasTotal, 0) + ' Horas')));
  L.push(lin_('Quantidade total de Recursos a serem alocados:', az(fmtNum_(d.recursosTotal, 0) + ' Recursos')));

  let totalValores = 0;
  const linhasValor = (d.valoresHora || [])
    .filter(v => num_(v.recursos) > 0 && num_(v.valorHora) > 0)
    .map(v => {
      totalValores += num_(v.valorHora) * num_(v.horas);
      return az(`${fmtNum_(v.recursos, 0)} Recursos - R$ ${fmtNum_(v.valorHora, 2)} Vlr/hora - ${fmtNum_(v.horas, 0)} horas de alocação`);
    });
  linhasValor.push(az(`Total R$ ${fmtNum_(totalValores, 2)}`));
  L.push(lin_('Valor(es) hora considerado(s) para alocação (por Recurso):', linhasValor.join(' \\\\ ')));
  L.push('');

  // ---------- TABELA 02 ----------
  L.push('h3. +TABELA 02 - Investimento do Projeto+');
  L.push('||Mês||Qtde. Recursos||Qtde. Dias Alocação||Horas de Alocação/Dia||Total Recursos x Dias x Horas por mês||Vlr. (R$) / Hora (Base de Cálculo)||Valor (R$)||Vlr. Despesas Extras (R$)||TOTAIS (R$)||');
  L.push('|*Fórmulas*|*(A)*|*(B)*|*(C)*|*(D) = (A x B x C)*|*(E)*|*(F) = (D x E)*|*(G)*|*H = (F + G)*|');

  const tot = { d: 0, f: 0, g: 0, h: 0 };
  CONFIG_PI.MESES.forEach((mes, i) => {
    const m = (d.meses || [])[i] || {};
    const a = num_(m.a), b = num_(m.b), c = num_(m.c), e = num_(m.e), g = num_(m.g);
    const D = a * b * c, Fv = D * e, H = Fv + g;
    tot.d += D; tot.f += Fv; tot.g += g; tot.h += H;
    L.push(`|*${mes}*|${vd(fmtNum_(a, 0))}|${vd(fmtNum_(b, 0))}|${vd(fmtNum_(c, 0))}|${vd(fmtNum_(D, 0))}|` +
      `${vd(fmtNum_(e, 2))}|${vd(fmtNum_(Fv, 2))}|${vd(fmtNum_(g, 2))}|*${vd(fmtNum_(H, 2))}*|`);
  });
  L.push(`|*N/A*|*N/A*|*N/A*|*TOTAIS*|*${az(fmtNum_(tot.d, 0))}*|*N/A*|*${az(fmtNum_(tot.f, 2))}*|` +
    `*${az(fmtNum_(tot.g, 2))}*|*${az(fmtNum_(tot.h, 2))}*|`);
  L.push('');

  // ---------- TABELA 03 (opcional) ----------
  const b = d.beneficio || {};
  const temBeneficio = !d.semRetorno && [b.detalhamento, b.valor, b.dataInicio, b.dataFim]
    .some(v => String(v || '').trim());
  if (temBeneficio) {
    L.push('h3. +TABELA 03 - Benefícios e retornos financeiros esperados do Projeto+');
    L.push(lin_('Detalhamento:', vd(esc_(b.detalhamento || ' '))));
    L.push(lin_('Valor do Benefício:', '*' + vd(fmtNum_(b.valor, 2)) + '*'));
    L.push(lin_('Data Início:', vd(dataBR_(b.dataInicio))));
    L.push(lin_('Data Fim:', vd(dataBR_(b.dataFim))));
  }

  return L.join('\n');
}

/************************************************************
 * AUXILIARES DE FORMATAÇÃO
 ************************************************************/
function cor_(cor, texto) {
  const t = String(texto == null ? '' : texto);
  return t.trim() ? `{color:${cor}}${t}{color}` : ' ';
}

function lin_(rotulo, valor) {
  return `|*${esc_(rotulo)}*|${valor || ' '}|`;
}

/** Escapa caracteres especiais do wiki markup e converte quebras de linha. */
function esc_(v) {
  return String(v == null ? '' : v)
    .replace(/\\/g, '/')
    .replace(/([|{}\[\]*_+^~!#])/g, '\\$1')
    .replace(/\r?\n/g, ' \\\\ ')
    .trim();
}

function num_(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const s = String(v).trim();
  const n = s.indexOf(',') !== -1 ? parseFloat(s.replace(/\./g, '').replace(',', '.')) : parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function fmtNum_(v, casas) {
  const n = num_(v);
  const fixo = n.toFixed(casas);
  const partes = fixo.split('.');
  partes[0] = partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return partes.join(',');
}

function dataBR_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const iso = String(v).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return iso ? `${iso[3]}/${iso[2]}/${iso[1]}` : String(v);
}
