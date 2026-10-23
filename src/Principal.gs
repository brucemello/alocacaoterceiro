/************************************************************
 * PROJETO TERCEIRO - ARQUIVO PRINCIPAL (menu, config e token)
 * Use este arquivo numa planilha NOVA, junto com:
 *   CriarIssue.gs, TelaCriarIssue.html,
 *   ProjetoInterno.gs, TelaProjetoInterno.html
 ************************************************************/

const CONFIG = {
  JIRA_BASE_URL: 'https://jiraproducao.totvs.com.br',
  ABA_LOG: 'Logs',
  ABAS_PROTEGIDAS: ['Logs', 'Criar_Issues', 'Template_Bruto', 'Controle_Solicitacoes',
    'Rascunhos_PI', 'Rascunhos_Agenda']
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🦇 Agente A.I.F.R.E.A.D')
    .addItem('🧭 Abrir Painel', 'abrirPainel')
    .addSeparator()
    .addItem('🏗️ Solicitar Projeto Interno', 'abrirTelaProjetoInterno')
    .addItem('📅 Solicitar Agenda', 'abrirTelaAgenda')
    .addSeparator()
    .addItem('📊 Abrir Controle de Solicitações', 'abrirControle')
    .addItem('🔄 Atualizar Controle (status e nº solicitação)', 'atualizarControle')
    .addItem('📥 Importar issue(s) existente(s)', 'importarIssuesExistentes')
    .addItem('🔍 Diagnosticar linha selecionada', 'diagnosticarLinhaControle')
    .addItem('⏰ Ativar atualização automática (1h)', 'ativarAtualizacaoAutomaticaControle')
    .addItem('⏹️ Desativar atualização automática', 'desativarAtualizacaoAutomaticaControle')
    .addSeparator()
    .addItem('➕ Criar Issue no Jira (genérica)', 'abrirTelaCriarIssue')
    .addItem('📋 Preparar aba de criação em lote', 'prepararAbaCriacaoLote')
    .addItem('🚀 Criar issues da aba em lote', 'criarIssuesEmLote')
    .addSeparator()
    .addItem('🔎 Exportar template bruto (CIBKPRTH-7190)', 'exportarTemplateBrutoJira')
    .addItem('🧪 Testar Conexão Jira', 'testarConexaoJira')
    .addItem('🔑 Configurar Token Jira', 'configurarTokenJira')
    .addToUi();
}

function configurarTokenJira() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Configuração de Credenciais',
    'Insira seu Personal Access Token (PAT) do Jira:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() === ui.Button.OK && resp.getResponseText().trim()) {
    PropertiesService.getScriptProperties().setProperty('JIRA_TOKEN', resp.getResponseText().trim());
    ui.alert('Sucesso', 'Token salvo. Agora use "Testar Conexão Jira".', ui.ButtonSet.OK);
  }
}

function testarConexaoJira() {
  const ui = SpreadsheetApp.getUi();
  const token = PropertiesService.getScriptProperties().getProperty('JIRA_TOKEN');
  if (!token) {
    ui.alert('⚠️ Token ausente', 'Configure o token no menu "Configurar Token Jira".', ui.ButtonSet.OK);
    return;
  }
  const r = UrlFetchApp.fetch(CONFIG.JIRA_BASE_URL + '/rest/api/2/myself', {
    method: 'get', muteHttpExceptions: true,
    headers: { 'Authorization': 'Bearer ' + token.trim(), 'Accept': 'application/json' }
  });
  if (r.getResponseCode() === 200) {
    const d = JSON.parse(r.getContentText());
    ui.alert('✅ Conectado', `Jira Produção como:\n👤 ${d.displayName || d.name}\n✉️ ${d.emailAddress || 'N/A'}`, ui.ButtonSet.OK);
  } else {
    ui.alert('❌ Erro HTTP ' + r.getResponseCode(), 'Verifique se o token (PAT) está ativo.', ui.ButtonSet.OK);
  }
}
