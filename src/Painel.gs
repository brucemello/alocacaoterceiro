/************************************************************
 * AGENTE JIRA IA - PAINEL CENTRAL (v1.0)
 * Uma janela com abas: Controle · Projeto Interno · Agenda · Configurações
 * Os formulários existentes são carregados dentro do painel
 * (TelaProjetoInterno, TelaAgenda) sem alteração.
 ************************************************************/

function abrirPainel() {
  const html = HtmlService.createHtmlOutputFromFile('TelaPainel')
    .setWidth(1240)
    .setHeight(860);
  SpreadsheetApp.getUi().showModalDialog(html, '🦇 Agente A.I.F.R.E.A.D - Alocação de Recurso Terceiro');
}

/** Devolve o HTML de uma tela para ser carregada dentro do painel. */
function obterHtmlTelaPainel(nome) {
  const permitidas = ['TelaControle', 'TelaProjetoInterno', 'TelaAgenda', 'TelaCriarIssue'];
  if (permitidas.indexOf(nome) === -1) throw new Error('Tela não permitida: ' + nome);
  return HtmlService.createTemplateFromFile(nome).evaluate().getContent();
}

/************************************************************
 * CONTROLE
 ************************************************************/
function obterDadosControle() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_CT.ABA);
  const base = {
    linhas: [], statusOpcoes: CONFIG_CT.STATUS_CONTROLE,
    autoAtivo: autoAtualizacaoAtiva_(), jiraUrl: CONFIG.JIRA_BASE_URL
  };
  if (!aba || aba.getLastRow() < 2) return base;

  const n = aba.getLastRow() - 1;
  const txt = aba.getRange(2, 1, n, COLS_CT.length).getDisplayValues();
  const raw = aba.getRange(2, 1, n, COLS_CT.length).getValues();
  const forms = aba.getRange(2, C_CT['Link Solicitação'], n, 1).getFormulas();

  base.linhas = txt.map((t, i) => {
    const g = nome => t[C_CT[nome] - 1];
    const v = nome => raw[i][C_CT[nome] - 1];
    const aberta = v('Aberta em');
    return {
      tipo: g('Tipo'), issue: g('Issue').trim(), abertaEm: g('Aberta em'),
      abertaTs: aberta instanceof Date ? aberta.getTime() : 0,
      solicitante: g('Solicitante'), codigoPsa: g('Código Projeto (PSA)'), projeto: g('Projeto'),
      analista: g('Analista / Recurso'), squad: g('Squad'), de: g('Período De'), ate: g('Período Até'),
      valorHora: Number(v('Valor Hora (R$)')) || 0, dias: Number(v('Dias')) || 0,
      horas: Number(v('Horas')) || 0, valorTotal: Number(v('Valor Total (R$)')) || 0,
      statusJira: g('Status Jira'), numSolic: g('Nº Solicitação'),
      linkSolic: (String(forms[i][0] || '').match(/HYPERLINK\("([^"]+)"/) || [])[1] || '',
      statusControle: g('Status Controle'), obs: g('Observações'), atualizado: g('Última atualização')
    };
  }).filter(l => l.issue);
  return base;
}

/** Edita um campo do controle a partir do painel (só os campos manuais). */
function alterarCampoControlePainel(issue, campo, valor) {
  const editaveis = ['Status Controle', 'Observações', 'Analista / Recurso', 'Código Projeto (PSA)'];
  if (editaveis.indexOf(campo) === -1) throw new Error('Campo não editável: ' + campo);
  const aba = abaControle_();
  const linha = linhaControlePorIssue_(aba, issue);
  if (!linha) throw new Error('Issue não encontrada no controle: ' + issue);
  aba.getRange(linha, C_CT[campo]).setValue(valor);
  return true;
}

function atualizarControlePainel() {
  const importados = importarDosRascunhos_();
  const r = atualizarControleSilencioso_();
  r.importados = importados;
  return r;
}

function importarIssuesPainel(entrada) {
  if (!String(entrada || '').trim()) throw new Error('Informe as chaves das issues ou uma JQL.');
  const r = importarIssuesParaControle_(String(entrada).trim());
  if (r.importadas.length) atualizarControleSilencioso_();
  return r;
}

function irParaAbaControle() {
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(abaControle_());
  return true;
}

/************************************************************
 * CONFIGURAÇÕES
 ************************************************************/
function autoAtualizacaoAtiva_() {
  try {
    return ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'gatilhoAtualizarControle');
  } catch (e) { return false; }
}

function statusConfigPainel() {
  return {
    temToken: !!PropertiesService.getScriptProperties().getProperty('JIRA_TOKEN'),
    autoAtivo: autoAtualizacaoAtiva_(),
    usuario: Session.getActiveUser().getEmail() || '',
    jiraUrl: CONFIG.JIRA_BASE_URL
  };
}

function salvarTokenPainel(token) {
  token = String(token || '').trim();
  if (token.length < 10) throw new Error('Token muito curto. Cole o Personal Access Token completo.');
  PropertiesService.getScriptProperties().setProperty('JIRA_TOKEN', token);
  return true;
}

function testarConexaoPainel() {
  const d = chamarJiraApiCriacao_('/rest/api/2/myself', 'get');
  return { nome: d.displayName || d.name || '', email: d.emailAddress || '', login: d.name || '' };
}

function definirAutoAtualizacaoPainel(ativar) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'gatilhoAtualizarControle')
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (ativar) ScriptApp.newTrigger('gatilhoAtualizarControle').timeBased().everyHours(1).create();
  return autoAtualizacaoAtiva_();
}

/** Limpa o cache das templates (útil se alguém mudar tipo/prioridade/épico no Jira). */
function limparCacheTemplatesPainel() {
  CacheService.getScriptCache().removeAll(['PI_META_TEMPLATE', 'AG_META_TEMPLATE']);
  PropertiesService.getScriptProperties().deleteProperty('JIRA_CAMPO_EPIC_LINK');
  return true;
}
