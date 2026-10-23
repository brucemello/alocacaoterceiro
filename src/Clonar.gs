/************************************************************
 * 🦇 AGENTE A.I.F.R.E.A.D - CLONAR PROJETO INTERNO / AGENDA
 * Carrega os dados de uma issue existente para virar uma NOVA
 * solicitação (renovação): o formulário limpa período, feriados,
 * Tabela 02 e totais; o resto é aproveitado.
 ************************************************************/

/** Dados da issue de origem para o clone: { tipo, dados, origem }. */
function carregarParaClonar(key) {
  key = String(key || '').trim().toUpperCase();
  if (!/^[A-Z][A-Z0-9_]+-\d+$/.test(key)) throw new Error('Chave de issue inválida: ' + key);

  let tipo = '';
  try {
    const aba = abaControle_();
    const linha = linhaControlePorIssue_(aba, key);
    if (linha) tipo = aba.getRange(linha, C_CT['Tipo']).getValue();
  } catch (e) {}

  if (tipo === 'Agenda') {
    const r = carregarAgendaParaEdicao(key);
    return { tipo: 'Agenda', dados: r.dados, origem: r.origem };
  }
  if (tipo === 'Projeto Interno') return carregarProjetoInternoParaClone_(key);

  // Fora do Controle: decide pelo título no Jira
  const is = chamarJiraApiCriacao_(`/rest/api/2/issue/${encodeURIComponent(key)}?fields=summary`, 'get');
  const titulo = (is.fields && is.fields.summary) || '';
  if (/cria[çc][ãa]o de agenda/i.test(titulo)) {
    const r = carregarAgendaParaEdicao(key);
    return { tipo: 'Agenda', dados: r.dados, origem: r.origem };
  }
  return carregarProjetoInternoParaClone_(key);
}

function carregarProjetoInternoParaClone_(key) {
  const rasc = planilha_().getSheetByName(ABA_RASCUNHOS_PI);
  if (rasc && rasc.getLastRow() > 1) {
    const linhas = rasc.getRange(2, 1, rasc.getLastRow() - 1, 6).getDisplayValues();
    for (let i = linhas.length - 1; i >= 0; i--) {
      if (String(linhas[i][5]).trim().toUpperCase() === key) {
        let dados = {};
        try { dados = JSON.parse(rasc.getRange(i + 2, 9).getValue() || '{}'); } catch (e) {}
        return { tipo: 'Projeto Interno', dados: dados, origem: 'rascunho' };
      }
    }
  }
  const is = chamarJiraApiCriacao_(`/rest/api/2/issue/${encodeURIComponent(key)}?fields=summary,description,project,issuetype`, 'get');
  const f = is.fields || {};
  const dados = lerProjetoInternoDaDescricao_(f.description || '', f.summary || '');
  dados.projetoJira = (f.project && f.project.key) || CONFIG_PI.PROJETO_JIRA_PADRAO;
  dados.tipoIssueId = (f.issuetype && f.issuetype.id) || '';
  return { tipo: 'Projeto Interno', dados: dados, origem: 'jira' };
}

/** Lê a descrição (layout 7190) de volta para os campos fixos do formulário de Projeto Interno. */
function lerProjetoInternoDaDescricao_(desc, titulo) {
  const campo = {};
  linhasTabelaWiki_(desc).forEach(c => { if (c.length === 2) campo[normRotulo_(c[0])] = c[1]; });
  const pega = (...rot) => { for (const r of rot) { const k = Object.keys(campo).find(x => x.indexOf(r) === 0); if (k) return ehPlaceholder_(campo[k]) ? '' : campo[k]; } return ''; };
  const marcado = (texto, opcao) => new RegExp('\\[\\s*X\\s*\\]\\s*' + opcao.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i').test(texto);

  const tiposTxt = campo[Object.keys(campo).find(k => k.indexOf('tipo de projeto') === 0)] || '';
  const forecast = pega('custo nao considerado');
  return {
    nomeProjeto: pega('nome do projeto') || String(titulo).replace(/^.*nome do projeto:\s*/i, '').trim(),
    squad: pega('squad responsavel'), gestor: pega('gestor responsavel'), teamLeads: pega('team lead'),
    tiposProjeto: CONFIG_PI.TIPOS_PROJETO.filter(t => marcado(tiposTxt, t)),
    custoForecast: marcado(forecast, 'Sim') ? 'Sim' : marcado(forecast, 'Não') ? 'Não' : '',
    contratoFaturavel: pega('codigo do contrato'), oportunidade: pega('nr oportunidade'),
    centroCusto: pega('centro de custo'), classeValor: pega('classe de valor'), itemContabil: pega('item contabil'),
    objetivo: pega('objetivo do servico'), resultadoEsperado: pega('resultado final'), cenarioAtual: pega('cenario atual'),
    autoCalc: true
  };
}
