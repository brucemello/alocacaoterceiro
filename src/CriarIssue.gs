/************************************************************
 * AGENTE JIRA IA - MÓDULO DE CRIAÇÃO DE ISSUES (v1.0)
 * Reaproveita CONFIG e o token salvo em 'JIRA_TOKEN'
 * (menu "5. Configurar Token Geral Jira").
 *
 * Recursos:
 *  - Criar uma issue pelo painel (TelaCriarIssue.html)
 *  - Criar issues em lote a partir da aba 'Criar_Issues'
 *  - Consultar tipos de issue e campos obrigatórios do projeto
 *  - Vincular a issue criada a outra existente (ex.: CIBKPRTH-7190)
 *
 * ATENÇÃO: a issue é criada em nome do dono do token.
 ************************************************************/

const CONFIG_CRIACAO = {
  ABA_LOTE: 'Criar_Issues',
  TIPO_VINCULO_PADRAO: 'Relates',
  CABECALHO_LOTE: [
    'Projeto', 'Tipo', 'Resumo', 'Descrição', 'Responsável (login)',
    'Prioridade', 'Labels (vírgula)', 'Componentes (vírgula)', 'Data Limite (aaaa-mm-dd)',
    'Vincular a (issue)', 'Campos extras (JSON)', 'Issue Criada', 'Status / Erro'
  ]
};

/* ---------- Menu: adicione estas linhas no onOpen() do código principal ----------
    .addSeparator()
    .addItem('7. ➕ Criar Issue no Jira', 'abrirTelaCriarIssue')
    .addItem('8. 📋 Preparar aba de criação em lote', 'prepararAbaCriacaoLote')
    .addItem('9. 🚀 Criar issues da aba em lote', 'criarIssuesEmLote')
------------------------------------------------------------------------------------ */

function abrirTelaCriarIssue() {
  const html = HtmlService.createTemplateFromFile('TelaCriarIssue').evaluate()
    .setWidth(720)
    .setHeight(780);
  SpreadsheetApp.getUi().showModalDialog(html, '➕ Criar Issue no Jira');
}

/************************************************************
 * CHAMADA À API TOLERANTE A RESPOSTA VAZIA (ex.: issueLink → 201 sem corpo)
 ************************************************************/
function chamarJiraApiCriacao_(endpoint, method, payload) {
  const token = PropertiesService.getScriptProperties().getProperty('JIRA_TOKEN');
  if (!token) throw new Error('Token global Jira não configurado (menu 5).');

  const options = {
    method: method.toLowerCase(),
    muteHttpExceptions: true,
    headers: {
      'Authorization': 'Bearer ' + token.trim(),
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    }
  };
  if (payload) options.payload = JSON.stringify(payload);

  const resp = UrlFetchApp.fetch(CONFIG.JIRA_BASE_URL + endpoint, options);
  const code = resp.getResponseCode();
  const texto = resp.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error(traduzirErroJira_(code, texto));
  }
  return texto ? JSON.parse(texto) : {};
}

/** Transforma o JSON de erro do Jira em mensagem legível. */
function traduzirErroJira_(code, texto) {
  try {
    const j = JSON.parse(texto);
    const msgs = [].concat(j.errorMessages || []);
    Object.keys(j.errors || {}).forEach(campo => msgs.push(`${campo}: ${j.errors[campo]}`));
    if (msgs.length) return `Jira HTTP ${code} → ${msgs.join(' | ')}`;
  } catch (e) {}
  return `Jira HTTP ${code}: ${String(texto).substring(0, 500)}`;
}

/************************************************************
 * METADADOS: TIPOS DE ISSUE E CAMPOS OBRIGATÓRIOS DO PROJETO
 * (Jira Server/Data Center 8.4+ e 9.x)
 ************************************************************/
function listarTiposIssueProjeto(projetoKey) {
  if (!projetoKey) throw new Error('Informe a chave do projeto.');
  const key = encodeURIComponent(projetoKey.trim().toUpperCase());
  const resp = chamarJiraApiCriacao_(`/rest/api/2/issue/createmeta/${key}/issuetypes?maxResults=100`, 'get');
  const lista = resp.values || resp.issueTypes || [];
  return lista
    .filter(t => !t.subtask)
    .map(t => ({ id: t.id, nome: t.name }));
}

function listarCamposObrigatorios(projetoKey, tipoId) {
  if (!projetoKey || !tipoId) return [];
  const key = encodeURIComponent(projetoKey.trim().toUpperCase());
  const resp = chamarJiraApiCriacao_(
    `/rest/api/2/issue/createmeta/${key}/issuetypes/${encodeURIComponent(tipoId)}?maxResults=200`, 'get');
  const campos = resp.values || resp.fields || [];

  const ignorados = ['project', 'issuetype', 'summary', 'reporter'];
  return campos
    .filter(c => c.required && ignorados.indexOf(c.fieldId) === -1)
    .map(c => ({
      id: c.fieldId,
      nome: c.name,
      tipo: (c.schema && (c.schema.custom || c.schema.type)) || '',
      opcoes: (c.allowedValues || []).slice(0, 50).map(v => v.value || v.name || v.key || v.id)
    }));
}

/************************************************************
 * MONTAGEM DO PAYLOAD E CRIAÇÃO
 ************************************************************/
function montarPayloadIssue_(d) {
  if (!d.projeto) throw new Error('Projeto não informado.');
  if (!d.tipo) throw new Error('Tipo de issue não informado.');
  if (!d.resumo) throw new Error('Resumo não informado.');

  const fields = {
    project: { key: String(d.projeto).trim().toUpperCase() },
    issuetype: /^\d+$/.test(String(d.tipo)) ? { id: String(d.tipo) } : { name: String(d.tipo).trim() },
    summary: String(d.resumo).trim()
  };

  if (d.descricao) fields.description = String(d.descricao);
  if (d.responsavel) fields.assignee = { name: String(d.responsavel).trim() };
  if (d.prioridade) fields.priority = { name: String(d.prioridade).trim() };
  if (d.dataLimite) fields.duedate = formatarDataIso_(d.dataLimite);

  const labels = listaDeTexto_(d.labels).map(l => l.replace(/\s+/g, '_'));
  if (labels.length) fields.labels = labels;

  const comps = listaDeTexto_(d.componentes);
  if (comps.length) fields.components = comps.map(c => ({ name: c }));

  // Campos customizados: objeto { "customfield_xxx": valor } ou texto JSON
  let extras = d.camposExtras;
  if (typeof extras === 'string' && extras.trim()) {
    try { extras = JSON.parse(extras); }
    catch (e) { throw new Error('Campos extras não são um JSON válido: ' + e.message); }
  }
  if (extras && typeof extras === 'object') Object.assign(fields, extras);

  return { fields: fields };
}

/**
 * Gera a prévia (sem criar nada). Usada pelo painel antes da confirmação.
 */
function previsualizarIssue(dados) {
  return JSON.stringify(montarPayloadIssue_(dados), null, 2);
}

/**
 * Cria a issue no Jira e, opcionalmente, vincula a outra issue.
 * @return {{key:string, url:string, vinculo:string}}
 */
function criarIssueJira(dados) {
  const payload = montarPayloadIssue_(dados);
  const resp = chamarJiraApiCriacao_('/rest/api/2/issue', 'post', payload);
  const key = resp.key;
  if (!key) throw new Error('O Jira não retornou a chave da issue criada.');

  let vinculo = '';
  if (dados.vincularA) {
    try {
      vincularIssues_(key, String(dados.vincularA).trim().toUpperCase(), dados.tipoVinculo);
      vinculo = `Vinculada a ${String(dados.vincularA).trim().toUpperCase()}`;
    } catch (e) {
      vinculo = `Criada, mas o vínculo falhou: ${e.message}`;
    }
  }

  registrarLogCriacao_(key, payload, vinculo);
  return { key: key, url: `${CONFIG.JIRA_BASE_URL}/browse/${key}`, vinculo: vinculo };
}

function vincularIssues_(novaKey, existenteKey, tipoVinculo) {
  chamarJiraApiCriacao_('/rest/api/2/issueLink', 'post', {
    type: { name: tipoVinculo || CONFIG_CRIACAO.TIPO_VINCULO_PADRAO },
    inwardIssue: { key: novaKey },
    outwardIssue: { key: existenteKey }
  });
}

/************************************************************
 * CRIAÇÃO EM LOTE A PARTIR DA ABA 'Criar_Issues'
 ************************************************************/
function prepararAbaCriacaoLote() {
  const ss = planilha_();
  let aba = ss.getSheetByName(CONFIG_CRIACAO.ABA_LOTE);
  if (!aba) aba = ss.insertSheet(CONFIG_CRIACAO.ABA_LOTE);

  const cab = CONFIG_CRIACAO.CABECALHO_LOTE;
  aba.getRange(1, 1, 1, cab.length).setValues([cab])
    .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
  aba.setFrozenRows(1);
  aba.autoResizeColumns(1, cab.length);

  if (CONFIG.ABAS_PROTEGIDAS.indexOf(CONFIG_CRIACAO.ABA_LOTE) === -1) {
    CONFIG.ABAS_PROTEGIDAS.push(CONFIG_CRIACAO.ABA_LOTE);
  }
  SpreadsheetApp.getUi().alert('Aba pronta',
    `Preencha as linhas da aba '${CONFIG_CRIACAO.ABA_LOTE}' e use o menu "9. Criar issues da aba em lote".\n` +
    'Linhas que já têm valor em "Issue Criada" são ignoradas (evita duplicidade).',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function criarIssuesEmLote() {
  const ui = SpreadsheetApp.getUi();
  const aba = planilha_().getSheetByName(CONFIG_CRIACAO.ABA_LOTE);
  if (!aba || aba.getLastRow() < 2) {
    ui.alert('Nada para criar', `Prepare e preencha a aba '${CONFIG_CRIACAO.ABA_LOTE}' primeiro.`, ui.ButtonSet.OK);
    return;
  }

  const nCols = CONFIG_CRIACAO.CABECALHO_LOTE.length;
  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, nCols).getValues();
  const pendentes = dados
    .map((r, i) => ({ r: r, linha: i + 2 }))
    .filter(x => x.r[2] && !x.r[11]);

  if (!pendentes.length) {
    ui.alert('Nada pendente', 'Todas as linhas com resumo já têm uma issue criada.', ui.ButtonSet.OK);
    return;
  }

  const resumoProjetos = [...new Set(pendentes.map(x => String(x.r[0]).toUpperCase()))].join(', ');
  const conf = ui.alert('Confirmar criação',
    `Serão criadas ${pendentes.length} issue(s) no(s) projeto(s): ${resumoProjetos}.\nDeseja continuar?`,
    ui.ButtonSet.YES_NO);
  if (conf !== ui.Button.YES) return;

  let ok = 0, erro = 0;
  pendentes.forEach(x => {
    const r = x.r;
    try {
      const res = criarIssueJira({
        projeto: r[0], tipo: r[1], resumo: r[2], descricao: r[3], responsavel: r[4],
        prioridade: r[5], labels: r[6], componentes: r[7], dataLimite: r[8],
        vincularA: r[9], camposExtras: r[10]
      });
      aba.getRange(x.linha, 12).setFormula(`=HYPERLINK("${res.url}"; "${res.key}")`);
      aba.getRange(x.linha, 13).setValue('✅ Criada' + (res.vinculo ? ' — ' + res.vinculo : ''));
      ok++;
    } catch (e) {
      aba.getRange(x.linha, 13).setValue('❌ ' + e.message);
      erro++;
    }
    SpreadsheetApp.flush();
  });

  ui.alert('Lote concluído', `✅ Criadas: ${ok}\n❌ Com erro: ${erro}\nVeja a coluna "Status / Erro".`, ui.ButtonSet.OK);
}

/************************************************************
 * AUXILIARES
 ************************************************************/
function listaDeTexto_(v) {
  if (!v) return [];
  if (Array.isArray(v)) return v.map(String).map(s => s.trim()).filter(Boolean);
  return String(v).split(',').map(s => s.trim()).filter(Boolean);
}

function formatarDataIso_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = String(v).trim();
  const br = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  return br ? `${br[3]}-${br[2]}-${br[1]}` : s;
}

function registrarLogCriacao_(key, payload, vinculo) {
  try {
    const ss = planilha_();
    let aba = ss.getSheetByName(CONFIG.ABA_LOG);
    if (!aba) aba = ss.insertSheet(CONFIG.ABA_LOG);
    aba.appendRow([
      new Date(), Session.getActiveUser().getEmail() || 'desconhecido', 'CRIAR_ISSUE', key,
      payload.fields.project.key, payload.fields.summary, vinculo || ''
    ]);
  } catch (e) {}
}
