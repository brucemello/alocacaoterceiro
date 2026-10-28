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
  URL_HORACIUS: 'https://gestaoid.totvs.com.br/portal/',
  STATUS: ['Aguardando abertura', 'Chamado aberto', 'Em atendimento', 'Liberar no Horacius', 'Horacius liberado',
    'Concluído', 'Negado', 'Revogado', 'Cancelado'],
  SISTEMAS: ['E-mail / Google Workspace', 'Rede / VPN', 'Jira', 'Confluence', 'Fluig', 'Protheus (ambiente de desenvolvimento)',
    'TFS / Azure DevOps', 'GitHub / Bitbucket', 'Teams / Slack'],
  DIAS_AVISO_FIM: 15,
  // Checklist do líder no Horacius: perfis que todo terceiro precisa. Editável na aba 'Perfis_Acesso'.
  ABA_PERFIS: 'Perfis_Acesso',
  PERFIS_PADRAO: [
    ['Protheus', 'RDO - Perfil de pagamento - Terceiros'],
    ['Active Directory SP01', 'Conta de Rede EXT TOTVS (SP01)'],
    ['Active Directory SP01', 'ColaboradoresIdentityGsuitesEXTSP'],
    ['Active Directory SP01', 'SP.Genetec.EXT'],
    ['Active Directory SP01', 'mfa.identity.mandatory.sp.ext'],
    ['Active Directory SP01', 'Proxy-Geral'],
    ['Active Directory SP01', 'VPN-ACESSO-PADRAO'],
    ['Active Directory SP01', 'TOTVS-JIRA'],
    ['Active Directory SP01', 'SQUAD.Fiscal'],
    ['Active Directory SP01', 'TOTVS.TDN.MP.EDITORES'],
    ['Active Directory SP01', 'TOTVS.TDN.MP.EDITORES_DOC'],
    ['Active Directory SP01', 'TOTVS.TDN.TAF.TAF.SQA.CRIA'],
    ['Active Directory SP01', 'TOTVS.TDN.TAF.TAF.INOV.CRIA'],
    ['Active Directory SP01', 'TOTVS.TDN.TAF.TAF.MANUT.CRIA'],
    ['Active Directory totvspartners', 'TerceirosIdentityGsuitesTP'],
    ['Active Directory totvspartners', 'JIRA.TOTVSPARTNERS.BASICO'],
    ['PSA Dynamics', 'Contato'],
    ['PSA Dynamics', 'Contato (Recurso Reservável)'],
    ['PSA Dynamics', 'Unidade Organizacional RMO TERCEIROS - CONTATOS'],
    ['Universidade - Automático', 'TERCEIROS']
  ]
};

const COLS_PERFIS = ['Ativo', 'Perfil', 'Usar no checklist (Sim/Não)', 'Time (vazio = todos)'];

/** Aba com a lista de perfis do terceiro; criada com a lista padrão na primeira vez. */
function abaPerfisAcesso_() {
  const ss = planilha_();
  let aba = ss.getSheetByName(CONFIG_AC.ABA_PERFIS);
  if (aba) return aba;
  aba = ss.insertSheet(CONFIG_AC.ABA_PERFIS);
  aba.getRange(1, 1, 1, COLS_PERFIS.length).setValues([COLS_PERFIS])
    .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
  const linhas = CONFIG_AC.PERFIS_PADRAO.map(p => [p[0], p[1], 'Sim', '']);
  aba.getRange(2, 1, linhas.length, COLS_PERFIS.length).setValues(linhas);
  aba.getRange(2, 3, Math.max(aba.getMaxRows() - 1, 1), 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Sim', 'Não'], true).build());
  aba.setFrozenRows(1);
  [230, 330, 130, 170].forEach((w, i) => aba.setColumnWidth(i + 1, w));
  return aba;
}

/** Perfis do checklist do Horacius para um time: [{ativo, perfil, chave:'Ativo: Perfil'}]. */
function perfisAcesso_(time) {
  const aba = abaPerfisAcesso_();
  if (aba.getLastRow() < 2) return [];
  const norm = v => String(v || '').trim().toLowerCase();
  return aba.getRange(2, 1, aba.getLastRow() - 1, COLS_PERFIS.length).getDisplayValues()
    .filter(r => String(r[0]).trim() && String(r[1]).trim())
    .filter(r => !String(r[3]).trim() || !time || time === TODOS_TIMES ||
      String(r[3]).split(/\s*[,;]\s*/).some(t => norm(t) === norm(time)))
    .filter(r => !/^n/i.test(String(r[2]).trim()))
    .map(r => ({ ativo: String(r[0]).trim(), perfil: String(r[1]).trim(),
      chave: String(r[0]).trim() + ': ' + String(r[1]).trim() }));
}

const COLS_AC = ['ID', 'Criado em', 'Solicitante', 'Recurso', 'Código Recurso (PSA)', 'E-mail do Recurso', 'Consultoria',
  'Código Projeto (PSA)', 'Projeto', 'Squad', 'Período De', 'Período Até', 'Acessos', 'Justificativa',
  'Nº Chamado', 'Status', 'Observações', 'Issue Agenda', 'Última atualização', 'Time', 'Perfis Horacius'];
const C_AC = {}; COLS_AC.forEach((n, i) => C_AC[n] = i + 1);

function abaAcessos_() {
  const ss = planilha_();
  let aba = ss.getSheetByName(CONFIG_AC.ABA);
  if (aba) {                                   // planilhas antigas ganham as colunas novas no fim
    const atual = aba.getLastColumn();
    if (atual < COLS_AC.length) {
      if (aba.getMaxColumns() < COLS_AC.length) aba.insertColumnsAfter(aba.getMaxColumns(), COLS_AC.length - aba.getMaxColumns());
      aba.getRange(1, atual + 1, 1, COLS_AC.length - atual).setValues([COLS_AC.slice(atual)])
        .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
    }
    return aba;
  }
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
  const esc = escopoTime_('');
  const porIssue = {}, lista = [];

  // 1) Rascunhos de agenda: têm e-mail, código do recurso e consultoria
  const rasc = ss.getSheetByName(CONFIG_AG.ABA_RASCUNHOS);
  if (rasc && rasc.getLastRow() > 1) {
    rasc.getRange(2, 1, rasc.getLastRow() - 1, 9).getDisplayValues().forEach(r => {
      let d = {};
      try { d = JSON.parse(r[8] || '{}'); } catch (e) {}
      if (!d.nomeRecurso || !esc.filtro(timeDoUsuario_(r[2]))) return;   // rascunhos do meu time
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
        if (!esc.filtro(g('Time') || timeDoUsuario_(g('Solicitante')))) return;
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
function obterDadosAcessos(timeSel) {
  const ss = planilha_();
  const esc = escopoTime_(timeSel);
  const base = {
    time: esc.time, podeVerTodos: esc.podeVerTodos, times: listaTimes_(), meuTime: esc.perfil.time,
    linhas: [], statusOpcoes: CONFIG_AC.STATUS, sistemas: CONFIG_AC.SISTEMAS, perfis: [], diasAviso: CONFIG_AC.DIAS_AVISO_FIM,
    urlPortal: PropertiesService.getScriptProperties().getProperty('URL_PORTAL_ACESSOS') || CONFIG_AC.URL_PORTAL,
    urlTicket: CONFIG_AC.URL_TICKET, urlHoracius: CONFIG_AC.URL_HORACIUS, jiraUrl: CONFIG.JIRA_BASE_URL, solicitante: ''
  };
  try { base.solicitante = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  try { base.perfis = perfisAcesso_(esc.perfil.time || esc.time); } catch (e) { base.erroPerfis = e.message; }
  const aba = ss.getSheetByName(CONFIG_AC.ABA) ? abaAcessos_() : null;
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
      obs: g('Observações'), issue: g('Issue Agenda'), atualizado: g('Última atualização'),
      time: g('Time') || timeDoUsuario_(g('Solicitante')), perfisHoracius: g('Perfis Horacius')
    };
  }).filter(l => l.id && esc.filtro(l.time));
  try { const n = contarCredenciais_(); base.linhas.forEach(l => l.qtdCred = n[l.id] || 0); } catch (e) {}
  try { const ch = resumoChamadosAcesso_(); base.linhas.forEach(l => l.chamadosExtras = ch[l.id] || { total: 0, abertos: 0 }); } catch (e) {}
  base.statusChamados = STATUS_CH;
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
    set('Perfis Horacius', (d.perfisHoracius || []).join(', '));
    set('Time', linha ? (aba.getRange(linha, C_AC['Time']).getValue() || perfil_().time) : perfil_().time);

    if (!linha) linha = aba.getLastRow() + 1;
    aba.getRange(linha, 1, 1, COLS_AC.length).setValues([valores]);
    return { id: id, status: status };
  } finally {
    lock.releaseLock();
  }
}

/** Edição rápida pela lista (status, nº do chamado, observações). */
function alterarCampoAcesso(id, campo, valor) {
  const editaveis = ['Status', 'Nº Chamado', 'Observações', 'Perfis Horacius'];
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

/************************************************************
 * CREDENCIAIS GERADAS NO HORACIUS
 * - Aba 'Credenciais_Terceiros': ativo, acesso, usuário, quem guardou, entrega.
 *   A SENHA NÃO VAI PARA A PLANILHA (todos os editores leriam).
 * - A senha fica nas UserProperties de quem cadastrou: só essa pessoa
 *   consegue ver/copiar, e pode apagar depois de entregar ao analista.
 ************************************************************/
const ABA_CRED = 'Credenciais_Terceiros';
const COLS_CRED = ['ID', 'ID Solicitação', 'Recurso', 'Ativo', 'Acesso', 'Identificador do Usuário',
  'Registrado em', 'Registrado por', 'Senha', 'Entregue ao analista em', 'Observação'];
const C_CRED = {}; COLS_CRED.forEach((n, i) => C_CRED[n] = i + 1);
const PREFIXO_SENHA = 'CRED_SENHA_';

function abaCredenciais_() {
  const ss = planilha_();
  let aba = ss.getSheetByName(ABA_CRED);
  if (aba) return aba;
  aba = ss.insertSheet(ABA_CRED);
  aba.getRange(1, 1, 1, COLS_CRED.length).setValues([COLS_CRED])
    .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
  aba.setFrozenRows(1);
  [80, 90, 200, 210, 240, 160, 120, 150, 230, 140, 200].forEach((w, i) => aba.setColumnWidth(i + 1, w));
  aba.getRange(2, C_CRED['Registrado em'], aba.getMaxRows() - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  aba.getRange(2, C_CRED['Entregue ao analista em'], aba.getMaxRows() - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  return aba;
}

function credenciaisDaAba_() {
  const aba = planilha_().getSheetByName(ABA_CRED);
  if (!aba || aba.getLastRow() < 2) return [];
  return aba.getRange(2, 1, aba.getLastRow() - 1, COLS_CRED.length).getDisplayValues().map((r, i) => {
    const g = c => String(r[C_CRED[c] - 1] || '').trim();
    return { linha: i + 2, id: g('ID'), idSolic: g('ID Solicitação'), recurso: g('Recurso'), ativo: g('Ativo'), acesso: g('Acesso'),
      usuario: g('Identificador do Usuário'), registradoEm: g('Registrado em'), registradoPor: g('Registrado por'),
      senha: g('Senha'), entregue: g('Entregue ao analista em'), obs: g('Observação') };
  }).filter(c => c.id);
}

function senhaGuardada_(id) {
  try { return PropertiesService.getUserProperties().getProperty(PREFIXO_SENHA + id) || ''; } catch (e) { return ''; }
}

/** Credenciais de uma solicitação (sem senha; diz se a senha está com você). */
function listarCredenciaisAcesso(idSolic) {
  const eu = emailAtual_();
  return credenciaisDaAba_().filter(c => c.idSolic === String(idSolic)).map(c => {
    const minha = !!senhaGuardada_(c.id);
    return { id: c.id, ativo: c.ativo, acesso: c.acesso, usuario: c.usuario, registradoEm: c.registradoEm,
      registradoPor: c.registradoPor, entregue: c.entregue, obs: c.obs, senhaComigo: minha,
      senhaStatus: minha ? 'guardada com você' : (c.senha || 'não guardada'), souEu: c.registradoPor === eu };
  });
}

/** Contagem por solicitação, para a lista. */
function contarCredenciais_() {
  const n = {};
  credenciaisDaAba_().forEach(c => n[c.idSolic] = (n[c.idSolic] || 0) + 1);
  return n;
}

function salvarCredencialAcesso(d) {
  const idSolic = String(d.idSolic || '').trim();
  if (!idSolic) throw new Error('Salve a solicitação de acesso antes de registrar as credenciais.');
  ['ativo', 'usuario'].forEach(k => { if (!String(d[k] || '').trim()) throw new Error('Informe ' + (k === 'ativo' ? 'o ativo.' : 'o identificador do usuário.')); });
  const abaAc = abaAcessos_();
  const linhaSolic = linhaAcesso_(abaAc, idSolic);
  if (!linhaSolic) throw new Error('Solicitação não encontrada.');
  const recurso = abaAc.getRange(linhaSolic, C_AC['Recurso']).getValue();
  const eu = emailAtual_();

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const aba = abaCredenciais_();
    const atual = d.id ? credenciaisDaAba_().find(c => c.id === d.id) : null;
    const id = atual ? atual.id : Utilities.getUuid().substring(0, 8);
    const senha = String(d.senha || '');
    if (senha) PropertiesService.getUserProperties().setProperty(PREFIXO_SENHA + id, senha);
    const comigo = !!senhaGuardada_(id);
    const linha = atual ? atual.linha : aba.getLastRow() + 1;
    aba.getRange(linha, 1, 1, COLS_CRED.length).setValues([[
      id, idSolic, recurso, String(d.ativo).trim(), String(d.acesso || '').trim(), String(d.usuario).trim(),
      atual ? aba.getRange(linha, C_CRED['Registrado em']).getValue() : new Date(),
      atual ? atual.registradoPor : eu,
      comigo ? 'guardada com ' + eu + ' (fora da planilha)' : (atual ? atual.senha : 'não guardada'),
      atual ? aba.getRange(linha, C_CRED['Entregue ao analista em']).getValue() : '',
      String(d.obs || (atual ? atual.obs : '')).trim()
    ]]);
  } finally { lock.releaseLock(); }
  return listarCredenciaisAcesso(idSolic);
}

/** Só quem guardou consegue ler (UserProperties é por usuário). */
function revelarSenhaCredencial(id) {
  const s = senhaGuardada_(id);
  if (!s) throw new Error('A senha não está guardada com você (só quem cadastrou consegue ver).');
  return s;
}

function apagarSenhaCredencial(id) {
  PropertiesService.getUserProperties().deleteProperty(PREFIXO_SENHA + id);
  const c = credenciaisDaAba_().find(x => x.id === id);
  if (c) abaCredenciais_().getRange(c.linha, C_CRED['Senha']).setValue('apagada em ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + ' por ' + emailAtual_());
  return c ? listarCredenciaisAcesso(c.idSolic) : [];
}

function marcarCredencialEntregue(id, entregue) {
  const c = credenciaisDaAba_().find(x => x.id === id);
  if (!c) throw new Error('Credencial não encontrada.');
  abaCredenciais_().getRange(c.linha, C_CRED['Entregue ao analista em']).setValue(entregue ? new Date() : '');
  return listarCredenciaisAcesso(c.idSolic);
}

function removerCredencialAcesso(id) {
  const c = credenciaisDaAba_().find(x => x.id === id);
  if (!c) return [];
  PropertiesService.getUserProperties().deleteProperty(PREFIXO_SENHA + id);
  abaCredenciais_().deleteRow(c.linha);
  return listarCredenciaisAcesso(c.idSolic);
}

/** Texto para mandar ao analista (inclui as senhas que estão com você). */
function textoCredenciaisParaAnalista(idSolic) {
  const lista = credenciaisDaAba_().filter(c => c.idSolic === String(idSolic));
  if (!lista.length) throw new Error('Nenhuma credencial registrada.');
  const blocos = lista.map(c => [`Ativo: ${c.ativo}`, c.acesso ? `Acesso: ${c.acesso}` : null,
    `Usuário: ${c.usuario}`, `Senha: ${senhaGuardada_(c.id) || '(com ' + (c.registradoPor || 'quem cadastrou') + ')'}`]
    .filter(Boolean).join('\n'));
  return `Olá ${lista[0].recurso || ''},\n\nSeguem os seus acessos:\n\n${blocos.join('\n\n')}\n\n` +
    'Troque a senha no primeiro acesso e não compartilhe com ninguém.';
}

/************************************************************
 * OUTROS CHAMADOS DE ACESSO (Azure, VPN, sistemas…)
 * Cada solicitação de acesso pode ter vários chamados extras, cada um com
 * seu nº, status e datas, na aba 'Chamados_Acesso'.
 ************************************************************/
const ABA_CH = 'Chamados_Acesso';
const COLS_CH = ['ID', 'ID Solicitação', 'Recurso', 'Acesso / sistema', 'Nº Chamado', 'Status', 'Aberto em',
  'Concluído em', 'Observação', 'Registrado por', 'Atualizado em'];
const C_CH = {}; COLS_CH.forEach((n, i) => C_CH[n] = i + 1);
const STATUS_CH = ['A abrir', 'Aberto', 'Em atendimento', 'Aguardando terceiro', 'Concluído', 'Negado', 'Cancelado'];
const STATUS_CH_FECHADOS = ['Concluído', 'Negado', 'Cancelado'];

function abaChamadosAcesso_() {
  const ss = planilha_();
  let aba = ss.getSheetByName(ABA_CH);
  if (aba) return aba;
  aba = ss.insertSheet(ABA_CH);
  aba.getRange(1, 1, 1, COLS_CH.length).setValues([COLS_CH])
    .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
  aba.setFrozenRows(1);
  [80, 90, 200, 240, 110, 130, 110, 110, 280, 200, 130].forEach((w, i) => aba.setColumnWidth(i + 1, w));
  aba.getRange(2, C_CH['Status'], aba.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUS_CH, true).setAllowInvalid(true).build());
  return aba;
}

function chamadosDaAba_() {
  const aba = planilha_().getSheetByName(ABA_CH);
  if (!aba || aba.getLastRow() < 2) return [];
  return aba.getRange(2, 1, aba.getLastRow() - 1, COLS_CH.length).getDisplayValues().map((r, i) => {
    const g = c => String(r[C_CH[c] - 1] || '').trim();
    return { linha: i + 2, id: g('ID'), idSolic: g('ID Solicitação'), recurso: g('Recurso'), acesso: g('Acesso / sistema'),
      chamado: g('Nº Chamado'), status: g('Status'), abertoEm: g('Aberto em'), concluidoEm: g('Concluído em'),
      obs: g('Observação'), por: g('Registrado por'), atualizado: g('Atualizado em') };
  }).filter(c => c.id);
}

function listarChamadosAcesso(idSolic) {
  return chamadosDaAba_().filter(c => c.idSolic === String(idSolic)).map(c => { delete c.linha; return c; });
}

/** { aberto: n, total: n } por solicitação, para a lista. */
function resumoChamadosAcesso_() {
  const r = {};
  chamadosDaAba_().forEach(c => {
    const x = r[c.idSolic] = r[c.idSolic] || { total: 0, abertos: 0 };
    x.total++; if (STATUS_CH_FECHADOS.indexOf(c.status) === -1) x.abertos++;
  });
  return r;
}

function salvarChamadoAcesso(d) {
  const idSolic = String(d.idSolic || '').trim();
  if (!idSolic) throw new Error('Salve a solicitação de acesso antes de registrar outros chamados.');
  if (!String(d.acesso || '').trim()) throw new Error('Informe qual acesso / sistema o chamado pede.');
  const abaAc = abaAcessos_();
  const linhaSolic = linhaAcesso_(abaAc, idSolic);
  if (!linhaSolic) throw new Error('Solicitação não encontrada.');
  const recurso = abaAc.getRange(linhaSolic, C_AC['Recurso']).getValue();
  const nCh = numeroChamado_(d.chamado);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const aba = abaChamadosAcesso_();
    const atual = d.id ? chamadosDaAba_().find(c => c.id === d.id) : null;
    const id = atual ? atual.id : Utilities.getUuid().substring(0, 8);
    let status = String(d.status || '').trim() || (atual ? atual.status : '');
    if (!status) status = nCh ? 'Aberto' : 'A abrir';
    if (nCh && status === 'A abrir') status = 'Aberto';
    const linha = atual ? atual.linha : aba.getLastRow() + 1;
    const val = c => aba.getRange(linha, C_CH[c]).getValue();
    const abertoEm = atual ? (val('Aberto em') || (nCh ? new Date() : '')) : (nCh ? new Date() : '');
    const concluidoEm = STATUS_CH_FECHADOS.indexOf(status) !== -1 ? (atual && val('Concluído em') ? val('Concluído em') : new Date()) : '';
    aba.getRange(linha, 1, 1, COLS_CH.length).setValues([[id, idSolic, recurso, String(d.acesso).trim(),
      nCh ? `=HYPERLINK("${CONFIG_AC.URL_TICKET}${nCh}"; "${nCh}")` : '', status, abertoEm, concluidoEm,
      String(d.obs || '').trim(), atual ? atual.por : emailAtual_(), new Date()]]);
    [C_CH['Aberto em'], C_CH['Concluído em']].forEach(c => aba.getRange(linha, c).setNumberFormat('dd/MM/yyyy'));
    aba.getRange(linha, C_CH['Atualizado em']).setNumberFormat('dd/MM/yyyy HH:mm');
  } finally { lock.releaseLock(); }
  return listarChamadosAcesso(idSolic);
}

function alterarStatusChamadoAcesso(id, status) {
  const c = chamadosDaAba_().find(x => x.id === id);
  if (!c) throw new Error('Chamado não encontrado.');
  return salvarChamadoAcesso({ id: id, idSolic: c.idSolic, acesso: c.acesso, chamado: c.chamado, status: status, obs: c.obs });
}

function removerChamadoAcesso(id) {
  const c = chamadosDaAba_().find(x => x.id === id);
  if (!c) return [];
  abaChamadosAcesso_().deleteRow(c.linha);
  return listarChamadosAcesso(c.idSolic);
}

/** Vários acessos de uma vez ("Ryver, Kanoah") → um chamado "A abrir" para cada (ignora os que já existem). */
function adicionarChamadosAcesso(idSolic, nomes) {
  const ja = listarChamadosAcesso(idSolic).map(c => c.acesso.toLowerCase());
  const lista = (Array.isArray(nomes) ? nomes : String(nomes || '').split(/[,;\n]+/))
    .map(n => String(n).trim()).filter(Boolean).filter((n, i, a) => a.indexOf(n) === i && ja.indexOf(n.toLowerCase()) === -1);
  if (!lista.length) throw new Error('Informe os acessos (separe por vírgula). Os que já estão na lista não são repetidos.');
  let r = [];
  lista.forEach(n => { r = salvarChamadoAcesso({ idSolic: idSolic, acesso: n, status: 'A abrir' }); });
  return r;
}

/** Edição rápida na lista: nº, status ou observação de um chamado. */
function alterarCampoChamadoAcesso(id, campo, valor) {
  const c = chamadosDaAba_().find(x => x.id === id);
  if (!c) throw new Error('Chamado não encontrado.');
  const d = { id: id, idSolic: c.idSolic, acesso: c.acesso, chamado: c.chamado, status: c.status, obs: c.obs };
  if (campo === 'chamado') { d.chamado = valor; if (numeroChamado_(valor) && c.status === 'A abrir') d.status = 'Aberto'; }
  else if (campo === 'status') d.status = valor;
  else if (campo === 'obs') d.obs = valor;
  else throw new Error('Campo inválido.');
  return salvarChamadoAcesso(d);
}
