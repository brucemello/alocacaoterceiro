/************************************************************
 * 🦇 AGENTE A.I.F.R.E.A.D - ARQUIVO PRINCIPAL
 * Menu, configuração, token do Jira e acesso à planilha
 * (funciona tanto dentro da planilha quanto pelo link do App da Web).
 ************************************************************/

const CONFIG = {
  JIRA_BASE_URL: 'https://jiraproducao.totvs.com.br',
  // Link do App da Web (painel direto no navegador) usado nos e-mails e em Configurações
  URL_APP_WEB: 'https://script.google.com/a/macros/totvs.com.br/s/AKfycbxSmYpwtHQvrXVfDlalr0-WqtEhcaAbbmQ95Z0IKD9vXIf02vN-6cDv1oHyCR1G9wwS/exec',
  // Manual do TL (passo a passo) — aparece no tour guiado (botão ❓ Tour)
  URL_MANUAL: 'https://claude.ai/artifact/DRQeqkWsPJVWfJg7xVLmeq',
  ABA_LOG: 'Logs',
  ABAS_PROTEGIDAS: ['Logs', 'Criar_Issues', 'Template_Bruto', 'Controle_Solicitacoes',
    'Rascunhos_PI', 'Rascunhos_Agenda', 'Acessos_Terceiros', 'Times', 'Perfis_Acesso', 'Credenciais_Terceiros', 'Consultorias', 'PSA_Emails']
};

/**
 * Planilha do agente. Dentro da planilha usa a ativa (e guarda o ID);
 * pelo link do App da Web (sem planilha ativa) abre pelo ID guardado.
 */
function planilha_() {
  const props = PropertiesService.getScriptProperties();
  const ativa = SpreadsheetApp.getActiveSpreadsheet();
  if (ativa) {
    if (props.getProperty('PLANILHA_ID') !== ativa.getId()) props.setProperty('PLANILHA_ID', ativa.getId());
    return ativa;
  }
  const id = props.getProperty('PLANILHA_ID');
  if (!id) throw new Error('Abra a planilha "Projeto Terceiro" uma vez pelo menu do agente para registrar o link.');
  return SpreadsheetApp.openById(id);
}

function onOpen() {
  try { planilha_(); } catch (e) {}
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
    .addItem('📬 Ler e-mails do PSA (Agendas PSA)', 'lerEmailsPsaMenu')
    .addItem('⏰ Ativar atualização automática (15 min + e-mail)', 'ativarAtualizacaoAutomaticaControle')
    .addItem('📧 Enviar e-mail de teste', 'enviarEmailTesteControle')
    .addItem('📊 Enviar relatório semanal agora', 'enviarRelatorioSemanalMenu')
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
    PropertiesService.getUserProperties().setProperty('JIRA_TOKEN', resp.getResponseText().trim());
    ui.alert('Sucesso', 'Seu token pessoal foi salvo (só você tem acesso a ele). Agora use "Testar Conexão Jira".', ui.ButtonSet.OK);
  }
}

function testarConexaoJira() {
  const ui = SpreadsheetApp.getUi();
  const token = tokenJira_();
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
