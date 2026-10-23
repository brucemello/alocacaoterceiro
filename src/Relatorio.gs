/************************************************************
 * 🦇 AGENTE A.I.F.R.E.A.D - RELATÓRIO SEMANAL POR E-MAIL
 * Toda segunda-feira às 8h (fuso America/Sao_Paulo):
 *  - valores: Projetos (orçado) · Agendas (alocado) · Saldo
 *  - projetos com alocado/saldo, agendas, pendências
 *  - acessos de terceiros terminando
 * Destinatários: os mesmos do e-mail de mudanças (Configurações).
 ************************************************************/

const RELATORIO = { DIA: 'MONDAY', HORA: 8, DIAS_AGENDA_FIM: 30 };

/************************************************************
 * AGENDAMENTO
 ************************************************************/
function relatorioSemanalAtivo_() {
  try { return ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'gatilhoRelatorioSemanal'); }
  catch (e) { return false; }
}

function definirRelatorioSemanalPainel(ativar) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'gatilhoRelatorioSemanal')
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (ativar) {
    ScriptApp.newTrigger('gatilhoRelatorioSemanal').timeBased()
      .onWeekDay(ScriptApp.WeekDay[RELATORIO.DIA]).atHour(RELATORIO.HORA).everyWeeks(1).create();
  }
  return relatorioSemanalAtivo_();
}

function gatilhoRelatorioSemanal() {
  try { enviarRelatorioSemanal_(false); }
  catch (e) { Logger.log('Falha no relatório semanal: ' + e.message); }
}

/** Menu: envia agora (também autoriza o envio na primeira vez). */
function enviarRelatorioSemanalMenu() {
  const para = enviarRelatorioSemanal_(true);
  SpreadsheetApp.getUi().alert('📊 Relatório enviado', 'Enviado para: ' + para.join(', '), SpreadsheetApp.getUi().ButtonSet.OK);
}

function enviarRelatorioAgoraPainel() {
  try { return enviarRelatorioSemanal_(true); }
  catch (e) {
    if (/permiss|autoriz|authoriz/i.test(e.message)) {
      throw new Error('Falta autorizar o envio de e-mails: use o menu 🦇 Agente A.I.F.R.E.A.D → 📊 Enviar relatório semanal agora (uma vez).');
    }
    throw e;
  }
}

/************************************************************
 * DADOS
 ************************************************************/
function dadosRelatorio_() {
  try { atualizarControleSilencioso_(); } catch (e) { Logger.log('Relatório: não atualizou o Jira: ' + e.message); }

  const linhas = obterDadosControle().linhas;
  const cancelada = l => l.statusControle === 'Cancelado' || /cancel/i.test(l.statusJira || '');
  const norm = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/g, ' ').trim();
  const ativos = linhas.filter(l => !cancelada(l));
  const projetos = ativos.filter(l => l.tipo === 'Projeto Interno');
  const agendas = ativos.filter(l => l.tipo === 'Agenda');

  const projetoDaAgenda = a => {
    const cod = String(a.codigoPsa || '').toUpperCase(), nome = norm(a.projeto);
    return projetos.find(p => cod && String(p.codigoPsa || '').toUpperCase() === cod) ||
           projetos.find(p => nome && norm(p.projeto) === nome) || null;
  };
  agendas.forEach(a => { const p = projetoDaAgenda(a); a.projetoIssue = p ? p.issue : ''; });
  projetos.forEach(p => {
    const ags = agendas.filter(a => a.projetoIssue === p.issue);
    p.agendas = ags.length;
    p.alocado = ags.reduce((s, a) => s + a.valorTotal, 0);
    p.horasAlocadas = ags.reduce((s, a) => s + a.horas, 0);
    p.saldo = p.valorTotal - p.alocado;
    p.pct = p.valorTotal ? Math.round(p.alocado / p.valorTotal * 100) : 0;
  });

  const soma = (arr, c) => arr.reduce((s, l) => s + (l[c] || 0), 0);
  const vincs = agendas.filter(a => a.projetoIssue);
  const tot = {
    projetos: projetos.length, agendas: agendas.length, cancelados: linhas.length - ativos.length,
    orcado: soma(projetos, 'valorTotal'), horasOrcadas: soma(projetos, 'horas'),
    alocado: soma(agendas, 'valorTotal'), horasAlocadas: soma(agendas, 'horas'),
    alocadoVinc: soma(vincs, 'valorTotal'), horasAlocadasVinc: soma(vincs, 'horas')
  };
  tot.saldo = tot.orcado - tot.alocadoVinc;
  tot.horasSaldo = tot.horasOrcadas - tot.horasAlocadasVinc;
  tot.pct = tot.orcado ? Math.round(tot.alocadoVinc / tot.orcado * 100) : 0;

  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  const brParaData = s => { const m = String(s || '').match(/(\d{2})\/(\d{2})\/(\d{4})/); return m ? new Date(+m[3], +m[2] - 1, +m[1]) : null; };
  const dias = d => d ? Math.round((d - hoje) / 86400000) : null;
  const agendasTerminando = agendas.map(a => Object.assign(a, { diasFim: dias(brParaData(a.ate)) }))
    .filter(a => a.diasFim !== null && a.diasFim >= 0 && a.diasFim <= RELATORIO.DIAS_AGENDA_FIM)
    .sort((x, y) => x.diasFim - y.diasFim);
  const pendentes = ativos.filter(l => ['Aguardando solicitação', 'Solicitação aberta'].indexOf(l.statusControle) !== -1);

  let acessos = [];
  try {
    const ac = obterDadosAcessos();
    acessos = ac.linhas.filter(l => {
      const d = l.ateTs ? Math.ceil((l.ateTs - Date.now()) / 86400000) : null;
      l.diasFim = d;
      return ['Aguardando abertura', 'Chamado aberto', 'Em atendimento', 'Liberar no Horacius'].indexOf(l.status) !== -1 ||
        ((l.status === 'Concluído' || l.status === 'Horacius liberado') && d !== null && d <= ac.diasAviso);
    });
  } catch (e) {}

  return { tot, projetos, agendas, pendentes, agendasTerminando, acessos };
}

/************************************************************
 * E-MAIL
 ************************************************************/
function enviarRelatorioSemanal_(manual) {
  const cfg = configNotificacao_();
  if (!cfg.destinatarios.length) throw new Error('Nenhum destinatário configurado.');
  const d = dadosRelatorio_();
  const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  MailApp.sendEmail({
    to: cfg.destinatarios.join(','),
    subject: `🦇 A.I.F.R.E.A.D · Relatório semanal ${hoje} · Saldo ${moeda_(d.tot.saldo)} · ${d.tot.pct}% alocado`,
    htmlBody: montarRelatorioHtml_(d, hoje, manual),
    body: montarRelatorioTexto_(d, hoje),
    name: 'Agente A.I.F.R.E.A.D'
  });
  return cfg.destinatarios;
}

function moeda_(n) {
  const v = Number(n || 0), neg = v < 0;
  const [i, c] = Math.abs(v).toFixed(2).split('.');
  return (neg ? '-' : '') + 'R$ ' + i.replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',' + c;
}
function inteiro_(n) { return String(Math.round(Number(n || 0))).replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }

function montarRelatorioHtml_(d, hoje, manual) {
  const esc = v => String(v == null ? '' : v).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const J = CONFIG.JIRA_BASE_URL;
  const issue = k => k ? `<a href="${J}/browse/${esc(k)}" style="color:#0052cc;text-decoration:none;font-weight:bold">${esc(k)}</a>` : '';
  const th = t => `<th style="text-align:left;padding:7px 8px;background:#f4f5f7;color:#42526e;font-size:11px;border-bottom:1px solid #dfe1e6;white-space:nowrap">${t}</th>`;
  const td = (t, extra) => `<td style="padding:7px 8px;border-bottom:1px solid #ebecf0;font-size:12px;vertical-align:top;${extra || ''}">${t}</td>`;
  const num = 'text-align:right;white-space:nowrap';
  const secao = (titulo, conteudo) => `<h3 style="font-size:14px;color:#1f4e78;margin:22px 0 8px">${titulo}</h3>${conteudo}`;
  const tabela = (cab, linhas, vazio) => linhas.length
    ? `<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #dfe1e6">${'<tr>' + cab.map(th).join('') + '</tr>'}${linhas.join('')}</table>`
    : `<div style="color:#6b778c;font-size:12px">${vazio}</div>`;
  const barra = pct => `<div style="background:#ebecf0;height:6px;border-radius:3px;width:120px;margin-top:4px"><div style="background:${pct > 100 ? '#de350b' : '#6554c0'};height:6px;border-radius:3px;width:${Math.min(pct, 100) * 1.2}px"></div></div>`;
  const card = (titulo, valor, sub, cor) => `<td style="width:33%;padding:0 5px;vertical-align:top">
      <div style="border:1px solid #dfe1e6;border-left:4px solid ${cor};border-radius:6px;padding:10px 12px">
        <div style="font-size:10px;font-weight:bold;color:#6b778c;text-transform:uppercase">${titulo}</div>
        <div style="font-size:19px;font-weight:bold;color:${cor};margin-top:3px">${valor}</div>
        <div style="font-size:11px;color:#6b778c;margin-top:2px">${sub}</div></div></td>`;
  const t = d.tot;

  const projetos = d.projetos.sort((a, b) => b.valorTotal - a.valorTotal).map(p => '<tr>' +
    td(issue(p.issue) + `<div style="color:#6b778c;font-size:11px">${esc(p.codigoPsa || 'sem código PSA')}</div>`) +
    td(esc(p.projeto) + (p.analista ? `<div style="color:#6b778c;font-size:11px">${esc(p.analista)}</div>` : '')) +
    td(esc(p.statusControle) + (p.ticket ? `<div style="color:#6b778c;font-size:11px">🎫 ${esc(p.ticket)} · ${esc(p.statusTicket)}</div>` : '')) +
    td(moeda_(p.valorTotal) + `<div style="color:#6b778c;font-size:11px">${inteiro_(p.horas)} h</div>`, num) +
    td(moeda_(p.alocado) + `<div style="color:#6b778c;font-size:11px">${p.agendas} agenda(s) · ${p.pct}%</div>` + barra(p.pct), num) +
    td(`<b style="color:${p.saldo < 0 ? '#bf2600' : '#006644'}">${moeda_(p.saldo)}</b>`, num) + '</tr>');

  const agendas = d.agendas.sort((a, b) => String(a.analista).localeCompare(String(b.analista))).map(a => '<tr>' +
    td(issue(a.issue)) +
    td(`<b>${esc(a.analista)}</b><div style="color:#6b778c;font-size:11px">${esc(a.codigoPsa)} · ${esc(a.projeto)}</div>`) +
    td(`${esc(a.de)} → ${esc(a.ate)}`, 'white-space:nowrap') +
    td(`${inteiro_(a.horas)} h<div style="color:#6b778c;font-size:11px">${a.valorHora ? moeda_(a.valorHora) + '/h' : ''}</div>`, num) +
    td(moeda_(a.valorTotal), num) +
    td(esc(a.statusControle) + (a.projetoIssue ? `<div style="color:#6b778c;font-size:11px">consome ${esc(a.projetoIssue)}</div>` :
      '<div style="color:#b65c00;font-size:11px">projeto fora do Controle</div>')) + '</tr>');

  const pend = d.pendentes.map(l => '<tr>' + td(issue(l.issue)) + td(esc(l.tipo)) + td(esc(l.projeto)) +
    td(`<b>${esc(l.statusControle)}</b>`) + td(esc(l.numSolic ? '#' + l.numSolic : '') + (l.ticket ? ' · 🎫 ' + esc(l.ticket) : '')) + '</tr>');

  const fim = d.agendasTerminando.map(a => '<tr>' + td(issue(a.issue)) + td(esc(a.analista)) + td(esc(a.projeto)) +
    td(`<b style="color:#b65c00">${a.diasFim === 0 ? 'termina hoje' : 'em ' + a.diasFim + ' dia(s)'}</b> (${esc(a.ate)})`) + '</tr>');

  const acessos = d.acessos.map(l => '<tr>' + td(`<b>${esc(l.recurso)}</b><div style="color:#6b778c;font-size:11px">${esc(l.email)}</div>`) +
    td(esc(l.projeto)) + td(l.chamado ? `<a href="${CONFIG_AC.URL_TICKET}${esc(l.chamado)}" style="color:#0052cc">#${esc(l.chamado)}</a>` : '—') +
    td(esc(l.status) + ((l.status === 'Concluído' || l.status === 'Horacius liberado') && l.diasFim !== null ? `<div style="color:${l.diasFim < 0 ? '#bf2600' : '#b65c00'};font-size:11px">${l.diasFim < 0 ? '⛔ período encerrado — revogar' : '⏳ termina em ' + l.diasFim + ' dia(s)'}</div>` : '')) + '</tr>');

  const web = urlAppWeb_();
  return `<div style="font-family:Arial,sans-serif;color:#172b4d;max-width:860px">
    <div style="background:#1f4e78;color:#fff;padding:14px 16px;border-radius:6px">
      <div style="font-size:17px;font-weight:bold">🦇 Agente A.I.F.R.E.A.D · Relatório semanal</div>
      <div style="font-size:12px;color:#cfe0f3">Alocação de Recurso Terceiro · ${hoje}${manual ? ' · envio manual' : ''}</div>
    </div>
    <table cellpadding="0" cellspacing="0" style="width:100%;margin-top:12px;border-collapse:separate"><tr>
      ${card('💼 Projetos (orçado)', moeda_(t.orcado), `${inteiro_(t.horasOrcadas)} h · ${t.projetos} projeto(s)`, '#0747a6')}
      ${card('📅 Agendas (alocado)', moeda_(t.alocado), `${inteiro_(t.horasAlocadas)} h · ${t.agendas} agenda(s)`, '#403294')}
      ${card('⚖️ Saldo dos projetos', moeda_(t.saldo), `${inteiro_(t.horasSaldo)} h livres · ${t.pct}% alocado`, t.saldo < 0 ? '#bf2600' : '#006644')}
    </tr></table>
    <div style="font-size:11px;color:#6b778c;margin-top:6px">Agenda consome o orçamento do projeto: os valores não são somados. ${t.cancelados ? t.cancelados + ' solicitação(ões) cancelada(s) fora das contas.' : ''}</div>
    ${secao('💼 Projetos Internos', tabela(['Issue', 'Projeto', 'Status', 'Orçado', 'Alocado em agendas', 'Saldo'], projetos, 'Nenhum projeto ativo.'))}
    ${secao('📅 Agendas', tabela(['Issue', 'Recurso / Projeto', 'Período', 'Horas', 'Valor', 'Status'], agendas, 'Nenhuma agenda ativa.'))}
    ${d.pendentes.length ? secao('⏳ Pendências de solicitação', tabela(['Issue', 'Tipo', 'Projeto', 'Status', 'Solicitação / Ticket'], pend, '')) : ''}
    ${d.agendasTerminando.length ? secao(`📆 Agendas terminando nos próximos ${RELATORIO.DIAS_AGENDA_FIM} dias`, tabela(['Issue', 'Recurso', 'Projeto', 'Fim'], fim, '')) : ''}
    ${d.acessos.length ? secao('🔑 Acessos de terceiros em andamento / terminando', tabela(['Recurso', 'Projeto', 'Chamado', 'Status'], acessos, '')) : ''}
    <div style="margin-top:18px">
      ${web ? `<a href="${web}" style="background:#0052cc;color:#fff;padding:8px 14px;border-radius:4px;text-decoration:none;font-weight:bold;font-size:12px">Abrir painel</a>` : ''}
    </div>
    <div style="font-size:11px;color:#6b778c;margin-top:14px">Enviado toda segunda-feira às ${RELATORIO.HORA}h. Para mudar destinatários ou desativar: Painel → ⚙️ Configurações.</div>
  </div>`;
}

function montarRelatorioTexto_(d, hoje) {
  const t = d.tot;
  return [`Relatório semanal A.I.F.R.E.A.D - ${hoje}`, '',
    `Projetos (orçado): ${moeda_(t.orcado)} - ${inteiro_(t.horasOrcadas)} h - ${t.projetos} projeto(s)`,
    `Agendas (alocado): ${moeda_(t.alocado)} - ${inteiro_(t.horasAlocadas)} h - ${t.agendas} agenda(s)`,
    `Saldo dos projetos: ${moeda_(t.saldo)} - ${t.pct}% alocado`, '',
    'PROJETOS', ...d.projetos.map(p => `- ${p.issue} ${p.codigoPsa || ''} ${p.projeto}: orçado ${moeda_(p.valorTotal)}, alocado ${moeda_(p.alocado)}, saldo ${moeda_(p.saldo)}`),
    '', 'AGENDAS', ...d.agendas.map(a => `- ${a.issue} ${a.analista} (${a.de} a ${a.ate}): ${inteiro_(a.horas)} h, ${moeda_(a.valorTotal)}`)
  ].join('\n');
}
