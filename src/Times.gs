/************************************************************
 * 🦇 AGENTE A.I.F.R.E.A.D - TIMES, PERFIS E TOKEN POR USUÁRIO
 *
 * - Aba 'Times': quem é quem (e-mail, nome, time, papel, destinatários)
 * - Token do Jira por usuário (UserProperties: só a própria pessoa lê);
 *   o token "do sistema" (ScriptProperties) fica como reserva e para as
 *   rotinas automáticas.
 * - Papéis: TL (vê o próprio time) · Gestor e Admin (veem todos).
 *   Admin também cadastra membros. O primeiro a se cadastrar vira Admin.
 * - A separação é de VISUALIZAÇÃO: quem tem acesso à planilha consegue
 *   abrir as abas e ver tudo.
 ************************************************************/

const ABA_TIMES = 'Times';
const COLS_TIMES = ['E-mail', 'Nome', 'Time', 'Papel', 'Destinatários do time (e-mails extras)', 'Atualizado em'];
const PAPEIS = ['TL', 'Gestor', 'Admin'];
const TODOS_TIMES = '*';

/************************************************************
 * TOKEN
 ************************************************************/
/** Token do usuário atual; se ele não tiver, usa o token do sistema. */
function tokenJira_() {
  let t = '';
  try { t = PropertiesService.getUserProperties().getProperty('JIRA_TOKEN') || ''; } catch (e) {}
  return t || PropertiesService.getScriptProperties().getProperty('JIRA_TOKEN') || '';
}

function origemToken_() {
  let pessoal = false;
  try { pessoal = !!PropertiesService.getUserProperties().getProperty('JIRA_TOKEN'); } catch (e) {}
  return pessoal ? 'pessoal' : (PropertiesService.getScriptProperties().getProperty('JIRA_TOKEN') ? 'sistema' : '');
}

/************************************************************
 * ABA TIMES
 ************************************************************/
function abaTimes_() {
  const ss = planilha_();
  let aba = ss.getSheetByName(ABA_TIMES);
  if (aba) return aba;
  aba = ss.insertSheet(ABA_TIMES);
  aba.getRange(1, 1, 1, COLS_TIMES.length).setValues([COLS_TIMES])
    .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
  aba.setFrozenRows(1);
  [260, 220, 180, 90, 320, 130].forEach((w, i) => aba.setColumnWidth(i + 1, w));
  aba.getRange(2, 4, aba.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(PAPEIS, true).build());
  return aba;
}

function membros_() {
  const aba = planilha_().getSheetByName(ABA_TIMES);
  if (!aba || aba.getLastRow() < 2) return [];
  return aba.getRange(2, 1, aba.getLastRow() - 1, COLS_TIMES.length).getDisplayValues()
    .map((r, i) => ({
      linha: i + 2, email: String(r[0]).trim().toLowerCase(), nome: String(r[1]).trim(), time: String(r[2]).trim(),
      papel: PAPEIS.indexOf(String(r[3]).trim()) !== -1 ? String(r[3]).trim() : 'TL',
      destinatarios: String(r[4]).split(/[,;\s]+/).filter(e => /@/.test(e))
    }))
    .filter(m => m.email);
}

function emailAtual_() {
  let e = '';
  try { e = Session.getActiveUser().getEmail(); } catch (x) {}
  if (!e) { try { e = Session.getEffectiveUser().getEmail(); } catch (x) {} }
  return String(e || '').toLowerCase();
}

/** Perfil de quem está usando (ou de um e-mail/login informado). */
function perfil_(emailOuLogin) {
  const alvo = String(emailOuLogin || emailAtual_()).trim().toLowerCase();
  const login = alvo.split('@')[0];
  const m = membros_().find(x => x.email === alvo || x.email.split('@')[0] === login);
  if (m) return Object.assign({ cadastrado: true }, m);
  return { cadastrado: false, email: alvo.indexOf('@') !== -1 ? alvo : '', nome: '', time: '', papel: 'TL', destinatarios: [] };
}

function timeDoUsuario_(emailOuLogin) { return emailOuLogin ? perfil_(emailOuLogin).time : ''; }

function veTodos_(p) { return p.papel === 'Gestor' || p.papel === 'Admin'; }

function listaTimes_() {
  const s = {};
  membros_().forEach(m => { if (m.time) s[m.time] = true; });
  return Object.keys(s).sort((a, b) => a.localeCompare(b, 'pt-BR'));
}

/**
 * Decide qual time mostrar. timeSel: '' = o meu · '*' = todos (só Gestor/Admin) · nome do time.
 * Devolve { time, todos, filtro(timeDaLinha) }.
 */
function escopoTime_(timeSel) {
  const p = perfil_();
  const pode = veTodos_(p) || !p.time;              // sem time cadastrado ainda: vê tudo
  let time = String(timeSel || '').trim();
  if (time === TODOS_TIMES && !pode) time = p.time;
  if (time && time !== TODOS_TIMES && time !== p.time && !pode) time = p.time;
  if (!time) time = p.time || TODOS_TIMES;
  const norm = s => String(s || '').trim().toLowerCase();
  return {
    perfil: p, time: time, todos: time === TODOS_TIMES, podeVerTodos: pode,
    filtro: t => time === TODOS_TIMES || norm(t) === norm(time) || (!t && pode)
  };
}

/** Destinatários de e-mail de um time: membros do time + extras informados. */
function destinatariosDoTime_(time) {
  const lista = [];
  membros_().forEach(m => {
    if (m.time && time && m.time.toLowerCase() === String(time).toLowerCase()) {
      lista.push(m.email); m.destinatarios.forEach(d => lista.push(d.toLowerCase()));
    }
  });
  return lista.filter((v, i, a) => a.indexOf(v) === i);
}

function gestores_() {
  return membros_().filter(veTodos_).map(m => m.email);
}

/************************************************************
 * PAINEL: PERFIL E CADASTRO
 ************************************************************/
function obterPerfilPainel() {
  const p = perfil_();
  return {
    email: p.email || emailAtual_(), nome: p.nome, time: p.time, papel: p.papel, cadastrado: p.cadastrado,
    destinatarios: (p.destinatarios || []).join(', '), veTodos: veTodos_(p) || !p.time,
    admin: p.papel === 'Admin', times: listaTimes_(), token: origemToken_(),
    semAdmin: !membros_().some(m => m.papel === 'Admin'),
    tourVisto: tourVisto_(), urlManual: urlManual_()
  };
}

/************************************************************
 * TOUR GUIADO E MANUAL
 ************************************************************/
function tourVisto_() {
  try { return PropertiesService.getUserProperties().getProperty('TOUR_VISTO') === '1'; } catch (e) { return true; }
}

function marcarTourVistoPainel() {
  PropertiesService.getUserProperties().setProperty('TOUR_VISTO', '1');
  return true;
}

/** Link do manual do TL (propriedade URL_MANUAL ou CONFIG.URL_MANUAL). */
function urlManual_() {
  return PropertiesService.getScriptProperties().getProperty('URL_MANUAL') || CONFIG.URL_MANUAL || '';
}

/** O próprio usuário cadastra/atualiza nome, time e destinatários extras. */
function salvarPerfilPainel(dados) {
  const email = emailAtual_();
  if (!email) throw new Error('Não consegui identificar o seu e-mail Google.');
  const nome = String(dados.nome || '').trim(), time = String(dados.time || '').trim();
  if (!nome || !time) throw new Error('Informe nome e time.');
  const extras = String(dados.destinatarios || '').split(/[,;\s]+/).filter(Boolean);
  const invalidos = extras.filter(e => !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e));
  if (invalidos.length) throw new Error('E-mail inválido: ' + invalidos.join(', '));

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const aba = abaTimes_();
    const lista = membros_();
    const atual = lista.find(m => m.email === email);
    const papel = atual ? atual.papel : (lista.some(m => m.papel === 'Admin') ? 'TL' : 'Admin');  // primeiro = Admin
    const linha = atual ? atual.linha : aba.getLastRow() + 1;
    aba.getRange(linha, 1, 1, COLS_TIMES.length).setValues([[email, nome, time, papel, extras.join(', '), new Date()]]);
    aba.getRange(linha, 6).setNumberFormat('dd/MM/yyyy HH:mm');
  } finally { lock.releaseLock(); }
  return obterPerfilPainel();
}

/** Token pessoal (só a própria pessoa lê). */
function salvarTokenPessoalPainel(token) {
  token = String(token || '').trim();
  if (token.length < 10) throw new Error('Token muito curto. Cole o Personal Access Token completo.');
  PropertiesService.getUserProperties().setProperty('JIRA_TOKEN', token);
  return origemToken_();
}

function removerTokenPessoalPainel() {
  PropertiesService.getUserProperties().deleteProperty('JIRA_TOKEN');
  return origemToken_();
}

/** Admin: token do sistema (usado pelas rotinas automáticas e por quem não tiver token próprio). */
function salvarTokenSistemaPainel(token) {
  exigirAdmin_();
  token = String(token || '').trim();
  if (token.length < 10) throw new Error('Token muito curto.');
  PropertiesService.getScriptProperties().setProperty('JIRA_TOKEN', token);
  return true;
}

function exigirAdmin_() {
  const p = perfil_();
  const semAdmin = !membros_().some(m => m.papel === 'Admin');
  if (p.papel !== 'Admin' && !semAdmin) throw new Error('Somente o Admin pode fazer isso.');
}

function listarMembrosPainel() {
  exigirAdmin_();
  return membros_().map(m => ({ email: m.email, nome: m.nome, time: m.time, papel: m.papel, destinatarios: m.destinatarios.join(', ') }));
}

function salvarMembroPainel(m) {
  exigirAdmin_();
  const email = String(m.email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('E-mail inválido.');
  if (!String(m.time || '').trim()) throw new Error('Informe o time.');
  const papel = PAPEIS.indexOf(m.papel) !== -1 ? m.papel : 'TL';
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const aba = abaTimes_();
    const atual = membros_().find(x => x.email === email);
    const linha = atual ? atual.linha : aba.getLastRow() + 1;
    aba.getRange(linha, 1, 1, COLS_TIMES.length).setValues([[email, String(m.nome || '').trim(), String(m.time).trim(), papel,
      String(m.destinatarios || '').trim(), new Date()]]);
    aba.getRange(linha, 6).setNumberFormat('dd/MM/yyyy HH:mm');
  } finally { lock.releaseLock(); }
  return listarMembrosPainel();
}

function removerMembroPainel(email) {
  exigirAdmin_();
  email = String(email || '').trim().toLowerCase();
  if (email === emailAtual_()) throw new Error('Você não pode remover a si mesmo.');
  const m = membros_().find(x => x.email === email);
  if (m) abaTimes_().deleteRow(m.linha);
  return listarMembrosPainel();
}

function statusTokenSistemaPainel() {
  exigirAdmin_();
  return { temSistema: !!PropertiesService.getScriptProperties().getProperty('JIRA_TOKEN') };
}
