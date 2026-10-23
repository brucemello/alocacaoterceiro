/************************************************************
 * 🦇 AGENTE A.I.F.R.E.A.D - SOLICITAÇÃO DE ACESSOS DO TERCEIRO
 * Aba 'Acessos_Terceiros': prepara o texto do chamado, abre o
 * formulário do portal de suporte interno e acompanha o chamado.
 * (O portal não é acessível pelo script: o chamado é aberto pela
 *  pessoa; aqui fica o registro e o acompanhamento.)
 ************************************************************/

const CONFIG_AC = {
  ABA: 'Acessos_Terceiros',
  URL_PORTAL: 'https://suporte.interno.totvs.io/hc/pt-br/requests/new?ticket_form_id=43090319176595' +
    '&tf_43653456041107=peso_2' +
    '&tf_43093105543571=TI_Acessos_Acesso+em+Sistemas+TOTVS_Gest%C3%A3o+de+Identidades+-+Horacius_D%C3%BAvida+sobre+Horacius' +
    '&isLeaf=level5&tf_43093056156819=01K5HHB8ZCGVC4ZAE79Y9ARHGN&tf_undefined=portal_ti_brasil&pais=Brasil' +
    '&tf_49875679614355=pais_brasil&level1=Acessos&level2=Acesso+em+Sistemas+TOTVS' +
    '&level3=Gest%C3%A3o+de+Identidades+-+Horacius&level4=D%C3%BAvida+sobre+Horacius',
  URL_TICKET: 'https://suporte.interno.totvs.io/hc/pt-br/requests/',
  STATUS: ['Aguardando abertura', 'Chamado aberto', 'Em atendimento', 'Concluído', 'Negado', 'Revogado', 'Cancelado'],
  SISTEMAS: ['E-mail / Google Workspace', 'Rede / VPN', 'Jira', 'Confluence', 'Fluig', 'Protheus (ambiente de desenvolvimento)',
    'TFS / Azure DevOps', 'GitHub / Bitbucket', 'Teams / Slack'],
  DIAS_AVISO_FIM: 15
};

const COLS_AC = ['ID', 'Criado em', 'Solicitante', 'Recurso', 'Código Recurso (PSA)', 'E-mail do Recurso', 'Consultoria',
  'Código Projeto (PSA)', 'Projeto', 'Squad', 'Período De', 'Período Até', 'Acessos', 'Justificativa',
  'Nº Chamado', 'Status', 'Observações', 'Issue Agenda', 'Última atualização'];
const C_AC = {}; COLS_AC.forEach((n, i) => C_AC[n] = i + 1);

function abaAcessos_() {
  const ss = planilha_();
  let aba = ss.getSheetByName(CONFIG_AC.ABA);
  if (aba) return aba;
  aba = ss.insertSheet(CONFIG_AC.ABA);
  aba.getRange(1, 1, 1, COLS_AC.length).setValues([COLS_AC])
    .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff').setWrap(true).setVerticalAlignment('middle');
  [C_AC['Status'], C_AC['Observações'], C_AC['Nº Chamado']].forEach(c =>
    aba.getRange(1, c).setBackground('#ffab00').setFontColor('#172b4d'));
  aba.setFrozenRows(1);
  aba.setFrozenColumns(4);
  aba.setRowHeight(1, 36);
  const larg = { 'Recurso': 200, 'Projeto': 260, 'Acessos': 240, 'Justificativa': 260, 'Observações': 220 };
  COLS_AC.forEach((n, i) => aba.setColumnWidth(i + 1, larg[n] || 110));
  const nLin = aba.getMaxRows() - 1;
  aba.getRange(2, C_AC['Criado em'], nLin, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  aba.getRange(2, C_AC['Período De'], nLin, 2).setNumberFormat('dd/MM/yyyy');
  aba.getRange(2, C_AC['Última atualização'], nLin, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  aba.getRange(2, C_AC['Status'], nLin, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(CONFIG_AC.STATUS, true).setAllowInvalid(true).build());
  return aba;
}

/************************************************************
 * RECURSOS (para o combo): Agendas do Controle + dados dos rascunhos
 ************************************************************/
function listarRecursosParaAcesso() {
  const ss = planilha_();
  const porIssue = {}, lista = [];

  // 1) Rascunhos de agenda: têm e-mail, código do recurso e consultoria
  const rasc = ss.getSheetByName(CONFIG_AG.ABA_RASCUNHOS);
  if (rasc && rasc.getLastRow() > 1) {
    rasc.getRange(2, 1, rasc.getLastRow() - 1, 9).getDisplayValues().forEach(r => {
      let d = {};
      try { d = JSON.parse(r[8] || '{}'); } catch (e) {}
      if (!d.nomeRecurso) return;
      const item = {
        recurso: d.nomeRecurso, codigoRecurso: d.codigoRecurso || '', email: d.email || '',
        consultoria: d.consultoria || '', codigoProjeto: String(d.codigoProjeto || '').toUpperCase(),
        projeto: d.nomeProjeto || '', squad: d.squad || '', de: d.dataDe || '', ate: d.dataAte || '',
        issue: /^[A-Z][A-Z0-9_]+-\d+$/.test(r[5]) ? r[5] : '', origem: r[4] === 'Criada' ? 'Agenda criada' : 'Rascunho'
      };
      if (item.issue) porIssue[item.issue] = item;
      lista.push(item);
    });
  }

  // 2) Agendas do Controle (inclui as importadas de outras pessoas)
  if (ss.getSheetByName(CONFIG_CT.ABA)) {
    const aba = abaControle_();
    if (aba.getLastRow() > 1) {
      const dados = aba.getRange(2, 1, aba.getLastRow() - 1, COLS_CT.length).getValues();
      const txt = aba.getRange(2, 1, aba.getLastRow() - 1, COLS_CT.length).getDisplayValues();
      dados.forEach((l, i) => {
        const g = c => txt[i][C_CT[c] - 1];
        if (g('Tipo') !== 'Agenda' || g('Status Controle') === 'Cancelado') return;
        const issue = g('Issue').trim();
        const iso = v => v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
        const base = {
          recurso: g('Analista / Recurso').replace(/\s*\(.*\)\s*$/, ''), codigoRecurso: '', email: '',
          consultoria: (g('Analista / Recurso').match(/\(([^)]*)\)\s*$/) || [])[1] || '',
          codigoProjeto: g('Código Projeto (PSA)').toUpperCase(), projeto: g('Projeto'), squad: g('Squad'),
          de: iso(l[C_CT['Período De'] - 1]), ate: iso(l[C_CT['Período Até'] - 1]), issue: issue, origem: 'Controle'
        };
        if (porIssue[issue]) {   // completa o rascunho com o que o Controle tiver de mais novo
          const r = porIssue[issue];
          ['codigoProjeto', 'projeto', 'squad', 'de', 'ate'].forEach(k => { if (!r[k] && base[k]) r[k] = base[k]; });
          r.origem = 'Agenda criada';
        } else if (base.recurso) {
          lista.push(base);
        }
      });
    }
  }

  // Um item por recurso + projeto (o mais completo)
  const unicos = {};
  lista.forEach(it => {
    const k = (it.recurso + '|' + it.codigoProjeto).toLowerCase();
    const pontos = x => ['email', 'codigoRecurso', 'issue', 'de'].filter(c => x[c]).length;
    if (!unicos[k] || pontos(it) > pontos(unicos[k])) unicos[k] = it;
  });
  return Object.keys(unicos).map(k => unicos[k]).sort((a, b) => a.recurso.localeCompare(b.recurso, 'pt-BR'));
}

/************************************************************
 * DADOS DA ABA, SALVAR E EDITAR
 ************************************************************/
function obterDadosAcessos() {
  const ss = planilha_();
  const base = {
    linhas: [], statusOpcoes: CONFIG_AC.STATUS, sistemas: CONFIG_AC.SISTEMAS, diasAviso: CONFIG_AC.DIAS_AVISO_FIM,
    urlPortal: PropertiesService.getScriptProperties().getProperty('URL_PORTAL_ACESSOS') || CONFIG_AC.URL_PORTAL,
    urlTicket: CONFIG_AC.URL_TICKET, jiraUrl: CONFIG.JIRA_BASE_URL, solicitante: ''
  };
  try { base.solicitante = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  const aba = ss.getSheetByName(CONFIG_AC.ABA);
  if (!aba || aba.getLastRow() < 2) return base;

  const n = aba.getLastRow() - 1;
  const txt = aba.getRange(2, 1, n, COLS_AC.length).getDisplayValues();
  const raw = aba.getRange(2, 1, n, COLS_AC.length).getValues();
  base.linhas = txt.map((t, i) => {
    const g = c => t[C_AC[c] - 1], v = c => raw[i][C_AC[c] - 1];
    const ate = v('Período Até');
    return {
      id: g('ID'), criadoEm: g('Criado em'), criadoTs: v('Criado em') instanceof Date ? v('Criado em').getTime() : 0,
      solicitante: g('Solicitante'), recurso: g('Recurso'), codigoRecurso: g('Código Recurso (PSA)'), email: g('E-mail do Recurso'),
      consultoria: g('Consultoria'), codigoProjeto: g('Código Projeto (PSA)'), projeto: g('Projeto'), squad: g('Squad'),
      de: g('Período De'), ate: g('Período Até'), ateTs: ate instanceof Date ? ate.getTime() : 0,
      acessos: g('Acessos'), justificativa: g('Justificativa'), chamado: g('Nº Chamado'), status: g('Status'),
      obs: g('Observações'), issue: g('Issue Agenda'), atualizado: g('Última atualização')
    };
  }).filter(l => l.id);
  return base;
}

function salvarSolicitacaoAcesso(d) {
  if (!String(d.recurso || '').trim()) throw new Error('Informe o nome do recurso.');
  if (!String(d.email || '').trim()) throw new Error('Informe o e-mail do analista.');
  if (!d.de) throw new Error('Informe a data de início.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const aba = abaAcessos_();
    const id = d.id || Utilities.getUuid().substring(0, 8);
    const data = s => { const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : ''; };
    const chamado = numeroChamado_(d.chamado);
    let linha = linhaAcesso_(aba, id);
    const criadoEm = linha ? aba.getRange(linha, C_AC['Criado em']).getValue() : new Date();
    const solicitante = linha ? aba.getRange(linha, C_AC['Solicitante']).getValue() :
      (Session.getActiveUser().getEmail() || '').split('@')[0];
    const statusAtual = linha ? aba.getRange(linha, C_AC['Status']).getValue() : '';
    const status = d.status || (chamado ? (statusAtual && statusAtual !== 'Aguardando abertura' ? statusAtual : 'Chamado aberto')
      : (statusAtual || 'Aguardando abertura'));

    const valores = Array(COLS_AC.length).fill('');
    const set = (c, v) => valores[C_AC[c] - 1] = v;
    set('ID', id); set('Criado em', criadoEm); set('Solicitante', solicitante);
    set('Recurso', String(d.recurso).trim()); set('Código Recurso (PSA)', d.codigoRecurso || '');
    set('E-mail do Recurso', d.email || ''); set('Consultoria', d.consultoria || '');
    set('Código Projeto (PSA)', String(d.codigoProjeto || '').toUpperCase()); set('Projeto', d.projeto || '');
    set('Squad', d.squad || ''); set('Período De', data(d.de)); set('Período Até', data(d.ate));
    set('Acessos', (d.acessos || []).join(', ')); set('Justificativa', d.justificativa || '');
    set('Nº Chamado', chamado ? `=HYPERLINK("${CONFIG_AC.URL_TICKET}${chamado}"; "${chamado}")` : '');
    set('Status', status);
    set('Observações', linha ? aba.getRange(linha, C_AC['Observações']).getValue() : (d.obs || ''));
    set('Issue Agenda', d.issue ? `=HYPERLINK("${CONFIG.JIRA_BASE_URL}/browse/${d.issue}"; "${d.issue}")` : '');
    set('Última atualização', new Date());

    if (!linha) linha = aba.getLastRow() + 1;
    aba.getRange(linha, 1, 1, COLS_AC.length).setValues([valores]);
    return { id: id, status: status };
  } finally {
    lock.releaseLock();
  }
}

/** Edição rápida pela lista (status, nº do chamado, observações). */
function alterarCampoAcesso(id, campo, valor) {
  const editaveis = ['Status', 'Nº Chamado', 'Observações'];
  if (editaveis.indexOf(campo) === -1) throw new Error('Campo não editável: ' + campo);
  const aba = abaAcessos_();
  const linha = linhaAcesso_(aba, id);
  if (!linha) throw new Error('Solicitação não encontrada.');
  if (campo === 'Nº Chamado') {
    const n = numeroChamado_(valor);
    aba.getRange(linha, C_AC[campo]).setFormula(n ? `=HYPERLINK("${CONFIG_AC.URL_TICKET}${n}"; "${n}")` : '');
    const st = aba.getRange(linha, C_AC['Status']);
    if (n && st.getValue() === 'Aguardando abertura') st.setValue('Chamado aberto');
  } else {
    aba.getRange(linha, C_AC[campo]).setValue(valor);
  }
  aba.getRange(linha, C_AC['Última atualização']).setValue(new Date());
  return true;
}

/** Nº do chamado a partir do número ou do link colado (último grupo de 5+ dígitos). */
function numeroChamado_(v) {
  return (String(v || '').match(/\d{5,}/g) || []).pop() || '';
}

function linhaAcesso_(aba, id) {
  if (!id || aba.getLastRow() < 2) return 0;
  const ids = aba.getRange(2, 1, aba.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function irParaAbaAcessos() {
  const ss = planilha_(), aba = abaAcessos_();
  const dentro = !!SpreadsheetApp.getActiveSpreadsheet();
  if (dentro) ss.setActiveSheet(aba);
  return { url: ss.getUrl() + '#gid=' + aba.getSheetId(), web: !dentro };
}
