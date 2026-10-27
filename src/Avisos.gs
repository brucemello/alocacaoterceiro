/************************************************************
 * 🦇 AGENTE A.I.F.R.E.A.D - CONSULTORIAS E AVISO DE INÍCIO DA AGENDA
 *
 * - Aba 'Consultorias': nome + e-mails que recebem os avisos.
 * - "📧 Avisar início" (Controle, linha de Agenda): e-mail para o analista
 *   e a consultoria com os dados da agenda e o link de apontamento no Fluig.
 *   Depois do envio: Status Controle = 'Início comunicado' e coluna
 *   'Aviso de início' com data e destinatários.
 ************************************************************/

const CONFIG_AV = {
  ABA_CONSULTORIAS: 'Consultorias',
  URL_APONTAMENTO: 'https://fluig.totvs.com/portal/p/10097/apontamentoterceiros',
  STATUS_APOS_AVISO: 'Início comunicado'
};
const COLS_CONS = ['Consultoria', 'E-mails para avisos', 'Contato', 'Observação', 'Atualizado em', 'Atualizado por'];

/************************************************************
 * CADASTRO DE CONSULTORIAS
 ************************************************************/
function abaConsultorias_() {
  const ss = planilha_();
  let aba = ss.getSheetByName(CONFIG_AV.ABA_CONSULTORIAS);
  if (aba) return aba;
  aba = ss.insertSheet(CONFIG_AV.ABA_CONSULTORIAS);
  aba.getRange(1, 1, 1, COLS_CONS.length).setValues([COLS_CONS])
    .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
  aba.setFrozenRows(1);
  [300, 340, 200, 240, 130, 200].forEach((w, i) => aba.setColumnWidth(i + 1, w));
  return aba;
}

function consultorias_() {
  const aba = planilha_().getSheetByName(CONFIG_AV.ABA_CONSULTORIAS);
  if (!aba || aba.getLastRow() < 2) return [];
  return aba.getRange(2, 1, aba.getLastRow() - 1, COLS_CONS.length).getDisplayValues()
    .map((r, i) => ({ linha: i + 2, nome: String(r[0]).trim(), emails: listaEmails_(r[1]), contato: String(r[2]).trim(),
      obs: String(r[3]).trim(), atualizado: String(r[4]).trim() }))
    .filter(c => c.nome);
}

function listaEmails_(txt) {
  return String(txt || '').split(/[,;\s]+/).map(e => e.trim().toLowerCase()).filter(e => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e))
    .filter((v, i, a) => a.indexOf(v) === i);
}

/** "NEXPERTI CONSULTORIA EM INFORMATICA LTDA" ≈ "Nexperti" (ignora LTDA, ME, CONSULTORIA…). */
function chaveConsultoria_(nome) {
  return String(nome || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\b(ltda?s?|me|epp|eireli|s\.?\/?a|sa|consultoria|em|informatica|servicos|tecnologia|de|da|do|e)\b/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ').trim();
}

function acharConsultoria_(nome) {
  const k = chaveConsultoria_(nome);
  if (!k) return null;
  const lista = consultorias_();
  return lista.find(c => chaveConsultoria_(c.nome) === k) ||
    lista.find(c => { const ck = chaveConsultoria_(c.nome); return ck && (k.indexOf(ck) === 0 || ck.indexOf(k) === 0); }) || null;
}

function listarConsultoriasPainel() {
  return consultorias_().map(c => ({ nome: c.nome, emails: c.emails.join(', '), contato: c.contato, obs: c.obs, atualizado: c.atualizado }));
}

function salvarConsultoriaPainel(d) {
  const nome = String(d.nome || '').trim();
  if (!nome) throw new Error('Informe o nome da consultoria.');
  const brutos = String(d.emails || '').split(/[,;\s]+/).filter(Boolean);
  const invalidos = brutos.filter(e => !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e));
  if (invalidos.length) throw new Error('E-mail inválido: ' + invalidos.join(', '));
  if (!brutos.length) throw new Error('Informe ao menos um e-mail da consultoria.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const aba = abaConsultorias_();
    const atual = consultorias_().find(c => chaveConsultoria_(c.nome) === chaveConsultoria_(nome));
    const linha = atual ? atual.linha : aba.getLastRow() + 1;
    aba.getRange(linha, 1, 1, COLS_CONS.length).setValues([[nome, listaEmails_(brutos.join(',')).join(', '),
      String(d.contato || '').trim(), String(d.obs || '').trim(), new Date(), emailAtual_()]]);
    aba.getRange(linha, 5).setNumberFormat('dd/MM/yyyy HH:mm');
  } finally { lock.releaseLock(); }
  return listarConsultoriasPainel();
}

function removerConsultoriaPainel(nome) {
  const c = consultorias_().find(x => x.nome === nome);
  if (c) abaConsultorias_().deleteRow(c.linha);
  return listarConsultoriasPainel();
}

/************************************************************
 * AVISO DE INÍCIO DA AGENDA
 ************************************************************/
/** Dados da agenda (Controle + descrição/rascunho) e destinatários sugeridos. */
function dadosAvisoInicio_(key) {
  key = String(key || '').trim().toUpperCase();
  const linha = obterLinhasControle_().find(l => l.issue === key);
  if (!linha) throw new Error('Agenda não encontrada no Controle: ' + key);
  if (linha.tipo !== 'Agenda') throw new Error('O aviso de início é só para Agendas.');
  let ag = {};
  try { ag = carregarAgendaParaEdicao(key).dados || {}; } catch (e) { ag = {}; }
  const p = perfil_();
  const nomeCons = linha.consultoria || ag.consultoria || '';     // associação no Controle vale mais que a descrição
  const cons = acharConsultoria_(nomeCons);
  return {
    issue: key, url: CONFIG.JIRA_BASE_URL + '/browse/' + key,
    recurso: ag.nomeRecurso || linha.analista, email: String(ag.email || '').trim().toLowerCase(),
    codigoRecurso: ag.codigoRecurso || '', papel: ag.papel || '',
    consultoria: cons ? cons.nome : nomeCons, consultoriaAgenda: ag.consultoria || '', consultoriaCadastrada: cons ? cons.nome : '',
    consultorias: consultorias_().map(c => ({ nome: c.nome, emails: c.emails })),
    emailsConsultoria: cons ? cons.emails : [],
    codigoProjeto: ag.codigoProjeto || linha.codigoPsa, projeto: ag.nomeProjeto || linha.projeto,
    squad: ag.squad || linha.squad, de: linha.de, ate: linha.ate,
    horasDia: ag.horasDia || '', dias: linha.dias, horas: linha.horas, valorHora: linha.valorHora,
    naoAlocar: ag.naoAlocar || '', avisoAnterior: linha.avisoInicio || '',
    tlNome: p.nome || nomeDoEmail_(emailAtual_()), tlEmail: emailAtual_(), time: linha.time || p.time
  };
}

function nomeDoEmail_(e) {
  return String(e || '').split('@')[0].split(/[._-]+/).filter(Boolean)
    .map(x => x.charAt(0).toUpperCase() + x.slice(1)).join(' ');
}

/** Prévia do e-mail: { para, cc, assunto, html, dados }. */
function prepararAvisoInicioPainel(key, opcoes) {
  const d = dadosAvisoInicio_(key);
  const o = opcoes || {};
  if (o.consultoria !== undefined) {             // escolhida no comunicado
    const c = consultorias_().find(x => x.nome === o.consultoria);
    d.consultoria = o.consultoria; d.consultoriaCadastrada = c ? c.nome : ''; d.emailsConsultoria = c ? c.emails : [];
  }
  const primeiroNome = String(d.recurso || '').split(/\s+/)[0] || '';
  return {
    dados: d,
    para: d.email ? [d.email] : [],
    cc: d.emailsConsultoria.concat(d.tlEmail ? [d.tlEmail] : []),
    assunto: `Agenda aberta - ${d.recurso} - início ${d.de}`,
    html: htmlAvisoInicio_(d, o, primeiroNome)
  };
}

function htmlAvisoInicio_(d, o, primeiroNome) {
  const e = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const fmtN = n => Number(n || 0).toLocaleString('pt-BR');
  const moeda = n => 'R$ ' + Number(n || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const linha = (k, v) => v ? `<tr><td style="padding:6px 10px;background:#f4f5f7;color:#42526e;font-weight:bold;width:210px;border-bottom:1px solid #fff">${e(k)}</td>` +
    `<td style="padding:6px 10px;border-bottom:1px solid #ebecf0">${v}</td></tr>` : '';
  const extra = String(o.mensagem || '').trim();
  if (!o.detalhado) return htmlAvisoSimples_(d, e, linha, extra, primeiroNome);
  return `<div style="font-family:Arial,sans-serif;font-size:14px;color:#172b4d;max-width:680px">
  <p>Olá${primeiroNome ? ' ' + e(primeiroNome) : ''}${d.consultoria ? ' e equipe ' + e(d.consultoria) : ''},</p>
  <p>A <b>agenda de alocação foi criada</b>. Seguem os dados para o início das atividades:</p>
  <table style="border-collapse:collapse;font-size:13px;width:100%">
    ${linha('Recurso', '<b>' + e(d.recurso) + '</b>')}
    ${linha('Código do recurso (PSA)', e(d.codigoRecurso))}
    ${linha('Consultoria', e(d.consultoria))}
    ${linha('Projeto (PSA)', '<b>' + e(d.codigoProjeto) + '</b>' + (d.projeto ? ' - ' + e(d.projeto) : ''))}
    ${linha('Squad', e(d.squad))}
    ${linha('Período', '<b>' + e(d.de) + (d.ate ? ' a ' + e(d.ate) : '') + '</b>')}
    ${linha('Dias / horas alocados', d.dias ? fmtN(d.dias) + ' dias · ' + fmtN(d.horas) + ' horas' : '')}
    ${o.valorHora && d.valorHora ? linha('Valor hora negociado', moeda(d.valorHora)) : ''}
    ${linha('Datas sem alocação', e(d.naoAlocar))}
    ${linha('Líder responsável', e(d.tlNome) + (d.tlEmail ? ' (' + e(d.tlEmail) + ')' : ''))}
    ${linha('Agenda (Jira)', `<a href="${e(d.url)}">${e(d.issue)}</a>`)}
  </table>
  <p style="margin-top:18px"><b>Apontamento de horas:</b> registre as horas trabalhadas no Fluig, pelo link abaixo, dentro do período da agenda.</p>
  <p><a href="${e(CONFIG_AV.URL_APONTAMENTO)}" style="display:inline-block;background:#0052cc;color:#fff;padding:10px 18px;border-radius:4px;text-decoration:none;font-weight:bold">Apontar horas no Fluig</a></p>
  <p style="font-size:12px;color:#6b778c">${e(CONFIG_AV.URL_APONTAMENTO)}</p>
  ${extra ? `<p>${e(extra).replace(/\n/g, '<br>')}</p>` : ''}
  <p>Qualquer dúvida, fale com o líder responsável.</p>
  <p>Obrigado,<br>${e(d.tlNome)}${d.time ? '<br><span style="color:#6b778c">' + e(d.time) + '</span>' : ''}</p>
</div>`;
}

/** Aviso curto (padrão): agenda aberta, data de início, período e apontamento. Sem valores, dias ou horas. */
function htmlAvisoSimples_(d, e, linha, extra, primeiroNome) {
  return `<div style="font-family:Arial,sans-serif;font-size:14px;color:#172b4d;max-width:640px">
  <p>Olá${primeiroNome ? ' ' + e(primeiroNome) : ''}${d.consultoria ? ' e equipe ' + e(d.consultoria) : ''},</p>
  <p>Informamos que a <b>agenda de alocação de ${e(d.recurso)} está aberta</b>.
     O início das atividades será em <b>${e(d.de)}</b>.</p>
  <table style="border-collapse:collapse;font-size:13px;width:100%">
    ${linha('Data de início', '<b>' + e(d.de) + '</b>')}
    ${linha('Período da agenda', '<b>' + e(d.de) + (d.ate ? ' a ' + e(d.ate) : '') + '</b>')}
    ${linha('Projeto', e(d.projeto || d.codigoProjeto))}
    ${linha('Líder responsável', e(d.tlNome) + (d.tlEmail ? ' (' + e(d.tlEmail) + ')' : ''))}
  </table>
  <p style="margin-top:16px">As horas devem ser apontadas no Fluig: <a href="${e(CONFIG_AV.URL_APONTAMENTO)}">${e(CONFIG_AV.URL_APONTAMENTO)}</a></p>
  ${extra ? `<p>${e(extra).replace(/\n/g, '<br>')}</p>` : ''}
  <p>Obrigado,<br>${e(d.tlNome)}${d.time ? '<br><span style="color:#6b778c">' + e(d.time) + '</span>' : ''}</p>
</div>`;
}

/** Envia o aviso e atualiza o Controle. dados: { issue, para, cc, assunto, mensagem, valorHora, salvarConsultoria } */
function enviarAvisoInicioPainel(dados) {
  const para = listaEmails_((dados.para || []).join(','));
  const cc = listaEmails_((dados.cc || []).join(',')).filter(e => para.indexOf(e) === -1);
  if (!para.length && !cc.length) throw new Error('Informe ao menos um destinatário.');
  const op = { mensagem: dados.mensagem, valorHora: dados.valorHora, detalhado: !!dados.detalhado };
  if (dados.consultoria !== undefined) op.consultoria = dados.consultoria;
  const prev = prepararAvisoInicioPainel(dados.issue, op);
  const d = prev.dados;

  // Consultoria sem cadastro: guarda os e-mails informados (sem o analista e sem o próprio TL)
  if (dados.salvarConsultoria && d.consultoria) {
    const daCons = cc.filter(e => e !== d.email && e !== d.tlEmail);
    if (daCons.length) salvarConsultoriaPainel({ nome: d.consultoria, emails: daCons.join(', ') });
  }

  MailApp.sendEmail({
    to: (para.length ? para : cc).join(','),
    cc: para.length ? cc.join(',') : '',
    subject: String(dados.assunto || prev.assunto).trim(),
    htmlBody: prev.html,
    name: d.tlNome ? d.tlNome + ' (A.I.F.R.E.A.D)' : '🦇 A.I.F.R.E.A.D',
    replyTo: d.tlEmail || undefined
  });

  const aba = abaControle_();
  const linha = linhaControlePorIssue_(aba, d.issue);
  const quando = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const registro = `${quando} · ${para.concat(cc).join(', ')}`;
  if (linha) {
    aba.getRange(linha, C_CT['Aviso de início']).setValue(registro);
    if (d.consultoria) aba.getRange(linha, C_CT['Consultoria']).setValue(d.consultoria);
    const st = String(aba.getRange(linha, C_CT['Status Controle']).getValue());
    if (['Encerrado', 'Cancelado', 'Em andamento'].indexOf(st) === -1) {
      aba.getRange(linha, C_CT['Status Controle']).setValue(CONFIG_AV.STATUS_APOS_AVISO);
    }
    aba.getRange(linha, C_CT['Última atualização']).setValue(new Date());
  }
  return { registro: registro, status: CONFIG_AV.STATUS_APOS_AVISO };
}

/** Associa (ou troca) a consultoria de uma agenda no Controle. */
function associarConsultoriaAgenda(key, nome) {
  const aba = abaControle_();
  const linha = linhaControlePorIssue_(aba, String(key || '').trim().toUpperCase());
  if (!linha) throw new Error('Agenda não encontrada no Controle.');
  aba.getRange(linha, C_CT['Consultoria']).setValue(String(nome || '').trim());
  return true;
}
