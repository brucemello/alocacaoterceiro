/************************************************************
 * 🦇 AGENTE A.I.F.R.E.A.D - EXCLUSÃO DE AGENDA DE TERCEIRO
 *
 * Template CIBKPRTH-7192 ("Exclusão de Agenda"), épico CIBKPRTH-4294
 * (Agendas de Terceiros/Totvers). Os dados vêm da agenda original
 * (rascunho ou descrição da issue); o TL informa só as datas a excluir
 * e o motivo. A issue nasce vinculada à agenda original e entra no
 * Controle como "Exclusão de Agenda" (não soma em valores).
 ************************************************************/

const CONFIG_EX = {
  ISSUE_TEMPLATE: 'CIBKPRTH-7192',
  EPICO: 'CIBKPRTH-4294',
  PROJETO_JIRA_PADRAO: 'CIBKPRTH',
  TIPO_CONTROLE: 'Exclusão de Agenda',
  TITULO: d => `Exclusão de Agenda - Projeto: ${String(d.codigoProjeto || '').trim().toUpperCase()} - Recurso: ` +
    `${String(d.nomeRecurso || '').trim()} - ${String(d.nomeProjeto || '').trim()}`
};

/** Tipo, prioridade e épico da template 7192 (cache de 6h). */
function obterMetaTemplateEX_() {
  const cache = CacheService.getScriptCache();
  const salvo = cache.get('EX_META_TEMPLATE');
  if (salvo) return JSON.parse(salvo);
  let meta = { issuetypeId: '', issuetypeName: '', priorityId: '', priorityName: '', epicKey: '' };
  try {
    let campoEpic = '';
    try { campoEpic = obterCampoEpicLink_(); } catch (e) {}
    const r = chamarJiraApiCriacao_(`/rest/api/2/issue/${CONFIG_EX.ISSUE_TEMPLATE}?fields=issuetype,priority${campoEpic ? ',' + campoEpic : ''}`, 'get');
    const f = r.fields || {};
    meta = { issuetypeId: (f.issuetype && f.issuetype.id) || '', issuetypeName: (f.issuetype && f.issuetype.name) || '',
      priorityId: (f.priority && f.priority.id) || '', priorityName: (f.priority && f.priority.name) || '',
      epicKey: campoEpic && f[campoEpic] ? String(f[campoEpic]) : '' };
    cache.put('EX_META_TEMPLATE', JSON.stringify(meta), 21600);
  } catch (e) { Logger.log('Não consegui ler a template de exclusão: ' + e.message); }
  return meta;
}

/** Dados da agenda original para pré-preencher a exclusão. */
function carregarAgendaParaExclusao(key) {
  key = String(key || '').trim().toUpperCase();
  if (!/^[A-Z][A-Z0-9_]+-\d+$/.test(key)) throw new Error('Informe a chave da agenda (ex.: CIBKPRTH-15864).');
  const ag = carregarAgendaParaEdicao(key).dados || {};
  const linha = obterLinhasControle_().find(l => l.issue === key) || {};
  const exclusoes = obterLinhasControle_().filter(l => l.tipo === CONFIG_EX.TIPO_CONTROLE && l.agendaOrigem === key)
    .map(l => ({ issue: l.issue, status: l.statusJira, datas: l.obs }));
  return {
    agenda: key,
    codigoRecurso: ag.codigoRecurso || '', nomeRecurso: ag.nomeRecurso || String(linha.analista || '').replace(/\s*\(.*\)\s*$/, ''),
    codigoProjeto: String(ag.codigoProjeto || linha.codigoPsa || '').toUpperCase(), nomeProjeto: ag.nomeProjeto || linha.projeto || '',
    consultoria: ag.consultoria || linha.consultoria || '', squad: ag.squad || linha.squad || '',
    valorHora: ag.valorHora || (linha.valorHora ? String(linha.valorHora).replace('.', ',') : ''),
    papel: ag.papel || '', projetoJira: ag.projetoJira || CONFIG_EX.PROJETO_JIRA_PADRAO,
    periodoDe: ag.dataDe || isoDeBR_(linha.de), periodoAte: ag.dataAte || isoDeBR_(linha.ate),
    horasDia: ag.horasDia || '8', exclusoesAnteriores: exclusoes,
    template: CONFIG_EX.ISSUE_TEMPLATE, epico: CONFIG_EX.EPICO
  };
}

function isoDeBR_(s) {
  const m = String(s || '').match(/(\d{2})\/(\d{2})\/(\d{4})/);
  return m ? `${m[3]}-${m[2]}-${m[1]}` : '';
}

/** "05/11/2026, 2026-11-06 07/11/2026" → ['05/11/2026', '06/11/2026', '07/11/2026'] (ordenadas, sem repetir). */
function normalizarDatasExclusao_(txt) {
  const out = [];
  String(txt || '').split(/[,;\s]+/).filter(Boolean).forEach(p => {
    let m = p.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/), d = null;
    if (m) d = new Date(+m[3], +m[2] - 1, +m[1]);
    else if ((m = p.match(/^(\d{4})-(\d{2})-(\d{2})$/))) d = new Date(+m[1], +m[2] - 1, +m[3]);
    if (!d || isNaN(d.getTime())) throw new Error('Data inválida: ' + p + ' (use DD/MM/AAAA).');
    out.push(d);
  });
  return out.sort((a, b) => a - b).map(d => dataBR_(d)).filter((v, i, a) => a.indexOf(v) === i);
}

function validarDadosExclusao_(d) {
  const obrig = { codigoRecurso: 'Código Recurso (PSA)', nomeRecurso: 'Nome Completo do Recurso',
    codigoProjeto: 'Código do Projeto (PSA)', nomeProjeto: 'Nome do Projeto (PSA)', papel: 'Papel/Categoria do Recurso (PSA)',
    motivo: 'Motivo(s) da(s) exclusão(ões)' };
  const faltam = Object.keys(obrig).filter(k => !String(d[k] || '').trim()).map(k => obrig[k]);
  if (!(num_(d.valorHora) > 0)) faltam.push('Valor hora negociado da alocação');
  if (!d.datas || !d.datas.length) faltam.push('Excluir as alocações das seguintes datas');
  if (faltam.length) throw new Error('Preencha: ' + faltam.join(', '));
  if (!/^D\d{12}$/i.test(String(d.codigoProjeto).trim())) throw new Error('Código do Projeto (PSA) deve ser D + 12 dígitos.');
}

function montarDescricaoExclusao_(d) {
  const C = CONFIG_PI.CORES;
  const az = t => cor_(C.AZUL, t);
  const azB = t => String(t || '').trim() ? '*' + cor_(C.AZUL, t) + '*' : ' ';
  const vd = t => cor_(C.VERDE, t);
  const L = [];
  L.push('*Orientações Gerais para preenchimento do Template*:');
  L.push('* Campos em *AZUL* - Preenchimento obrigatório');
  L.push('* Campos em *VERDE* - Se não houver informações pertinentes a eles, deixar em branco');
  L.push('* Campos em *VERMELHO* - Não alterá-los.');
  L.push('');
  L.push('É obrigatório preencher os dados da *Tabelas 01*, conforme orientações das cores acima');
  L.push('');
  L.push('h3. +Tabela 01 - Dados da Alocação+');
  L.push('|*Descrição:*|*Solicito a exclusão da alocação conforme dados do Projeto e Recurso abaixo*|');
  L.push(lin_('Código Recurso (PSA):', azB(esc_(d.codigoRecurso))));
  L.push(lin_('Nome Completo do Recurso:', az(esc_(d.nomeRecurso))));
  L.push(lin_('Código do Projeto (PSA):', azB(esc_(String(d.codigoProjeto || '').toUpperCase()))));
  L.push(lin_('Nome do Projeto (PSA):', az(esc_(d.nomeProjeto))));
  L.push(lin_('Nome da Consultoria:', String(d.consultoria || '').trim() ? vd(esc_(d.consultoria)) : ' '));
  L.push(lin_('Nome do SQUAD:', String(d.squad || '').trim() ? vd(esc_(d.squad)) : ' '));
  L.push(lin_('Valor hora negociado da alocação:', azB('R$ ' + fmtNum_(num_(d.valorHora), 2))));
  L.push(lin_('Papel/Categoria do Recurso (PSA):', az(esc_(d.papel))));
  L.push(lin_('Excluir as alocações das seguintes datas:', azB(esc_((d.datas || []).join(', ')))));
  L.push(lin_('Motivo(s) da(s) exclusão(ões) da(s) agenda(s):', azB(esc_(d.motivo))));
  return L.join('\n');
}

function gerarPreviaExclusao(dados) {
  dados.datas = normalizarDatasExclusao_(dados.datasTexto);
  validarDadosExclusao_(dados);
  return { titulo: CONFIG_EX.TITULO(dados), descricao: montarDescricaoExclusao_(dados), datas: dados.datas };
}

function criarIssueExclusao(dados) {
  dados.datas = normalizarDatasExclusao_(dados.datasTexto);
  validarDadosExclusao_(dados);
  const meta = obterMetaTemplateEX_();
  const projeto = String(dados.projetoJira || CONFIG_EX.PROJETO_JIRA_PADRAO).trim().toUpperCase();
  const fields = {
    project: { key: projeto },
    issuetype: { id: resolverTipoIssue_(projeto, meta) },
    summary: CONFIG_EX.TITULO(dados).substring(0, 255),
    description: montarDescricaoExclusao_(dados)
  };
  if (meta.priorityId) fields.priority = { id: meta.priorityId };
  aplicarCamposObrigatoriosProjeto_(fields, projeto, fields.issuetype.id, dados.camposProjeto);
  const pendentes = {};
  const permitidos = obterCamposPermitidos_(projeto, fields.issuetype.id);
  if (permitidos) Object.keys(fields).forEach(k => {
    if (['project', 'issuetype', 'summary'].indexOf(k) === -1 && !permitidos[k]) { pendentes[k] = fields[k]; delete fields[k]; }
  });
  const resp = criarComRetentativa_(fields, pendentes);
  if (!resp.key) throw new Error('O Jira não retornou a chave da issue criada.');
  const avisos = aplicarCamposPendentes_(resp.key, pendentes);
  let status = vincularEpico_(resp.key, CONFIG_EX.EPICO || meta.epicKey) +
    (avisos.length ? '<br>⚠️ Não aplicado (fora das telas do projeto): ' + avisos.join(', ') : '');

  // Liga a exclusão à agenda original ("relates to")
  if (dados.agenda) {
    try {
      chamarJiraApiCriacao_('/rest/api/2/issueLink', 'post',
        { type: { name: 'Relates' }, inwardIssue: { key: resp.key }, outwardIssue: { key: String(dados.agenda).toUpperCase() } });
      status += '<br>🔗 Vinculada à agenda ' + dados.agenda;
    } catch (e) { status += '<br>⚠️ Não consegui vincular à agenda ' + dados.agenda + ': ' + e.message; }
  }

  const url = `${CONFIG.JIRA_BASE_URL}/browse/${resp.key}`;
  try {
    const log = planilha_().getSheetByName(CONFIG.ABA_LOG) || planilha_().insertSheet(CONFIG.ABA_LOG);
    log.appendRow([new Date(), Session.getActiveUser().getEmail() || 'desconhecido', 'EXCLUIR_AGENDA', resp.key, projeto,
      CONFIG_EX.TITULO(dados), status.replace(/<br>/g, ' ')]);
  } catch (e) {}
  try { registrarExclusaoNoControle_(resp.key, url, dados); } catch (e) {}
  return { key: resp.key, url: url, status: status, datas: dados.datas };
}

/** Linha "Exclusão de Agenda" no Controle: não soma valores; guarda datas e agenda de origem. */
function registrarExclusaoNoControle_(key, url, d) {
  const horasDia = num_(d.horasDia) || 8;
  const datas = d.datas || [];
  const toDate = s => { const m = String(s).match(/(\d{2})\/(\d{2})\/(\d{4})/); return m ? new Date(+m[3], +m[2] - 1, +m[1]) : ''; };
  const r = { codigoProjeto: String(d.codigoProjeto || '').toUpperCase(), projeto: d.nomeProjeto || '',
    analista: [d.nomeRecurso, d.consultoria ? `(${d.consultoria})` : ''].filter(Boolean).join(' '), squad: d.squad || '',
    de: toDate(datas[0]), ate: toDate(datas[datas.length - 1]), valorHora: num_(d.valorHora) || '',
    dias: datas.length, horas: datas.length * horasDia, valorTotal: '' };
  registrarLinhaControle_(CONFIG_EX.TIPO_CONTROLE, key, url, r);
  const aba = abaControle_();
  const linha = linhaControlePorIssue_(aba, key);
  if (linha) {
    aba.getRange(linha, C_CT['Observações']).setValue('Excluir: ' + datas.join(', ') + (d.motivo ? ' · Motivo: ' + d.motivo : ''));
    aba.getRange(linha, C_CT['Agenda de origem']).setValue(String(d.agenda || '').toUpperCase());
    if (d.consultoria) aba.getRange(linha, C_CT['Consultoria']).setValue(d.consultoria);
  }
}
