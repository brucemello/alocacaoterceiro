# 🦇 Agente A.I.F.R.E.A.D — Alocação de Recurso Terceiro

Automação em Google Apps Script (planilha **Projeto Terceiro**) integrada ao Jira de produção
(`jiraproducao.totvs.com.br`) para abrir e acompanhar as solicitações de **Projeto Interno** e de
**Agenda** de recursos terceiros.

## Funcionalidades

| Módulo | O que faz |
|---|---|
| 🧭 **Painel** | Janela central com abas: Controle · Projeto Interno · Agenda · Configurações |
| 🏗️ **Projeto Interno** | Formulário no layout do template `CIBKPRTH-7190`; Tabela 02 com cálculo automático; cria a Tarefa e vincula ao épico `CIBKPRTH-7181` |
| 📅 **Agenda** | Formulário no layout do template `CIBKPRTH-7191`; dias úteis por mês (feriados nacionais e datas a não alocar); horas × valor hora |
| 📊 **Controle** | Aba `Controle_Solicitacoes`: projeto, analista, valor hora, dias, horas, valor total, status do Jira, nº da solicitação (lido do comentário "Solicitação NNNNNNN iniciada com sucesso"), status de controle e observações |
| 📥 **Importar** | Traz para o controle issues abertas por outras pessoas (por chave ou JQL), lendo os valores da descrição |
| 💾 **Rascunhos** | Abas `Rascunhos_PI` e `Rascunhos_Agenda`, com salvamento automático |

A criação se adapta a cada projeto Jira: tipo de issue e prioridade são resolvidos por ID a partir
das templates, campos obrigatórios do projeto (ex.: *Componente(s)*) aparecem no formulário e
campos fora da tela de criação são aplicados depois por edição.

## Estrutura

```
src/
  appsscript.json          manifesto do projeto
  Principal.gs             menu, CONFIG, token e teste de conexão
  CriarIssue.gs            chamada à API do Jira e criação genérica/lote
  ProjetoInterno.gs        Projeto Interno (descrição, validação, criação, rascunhos)
  Agenda.gs                Agenda (descrição, validação, criação, rascunhos)
  Controle.gs              aba de controle, atualização, importação, analistas
  Painel.gs                painel central e configurações
  TelaPainel.html          painel com abas (carrega as telas abaixo)
  TelaControle.html        aba Controle
  TelaProjetoInterno.html  formulário de Projeto Interno
  TelaAgenda.html          formulário de Agenda
  TelaCriarIssue.html      criação genérica de issue
scripts/verificar-sintaxe.js
.github/workflows/
  publicar.yml             push em main (src/) → clasp push + atualiza implantação
  baixar-fontes.yml        manual → clasp pull + commit (quando editarem no editor online)
```

## Publicação (GitHub Actions)

**Secrets** (Settings → Secrets and variables → Actions):

| Secret | Conteúdo |
|---|---|
| `CLASPRC_JSON` | conteúdo completo do `~/.clasprc.json` gerado por `clasp login` |
| `SCRIPT_ID` | ID do projeto Apps Script (Editor → Configurações do projeto → ID do script) |
| `DEPLOYMENT_ID` | (opcional) ID da implantação a ser atualizada a cada publicação |

**Fluxo:**

1. Altere os arquivos em `src/` e faça commit na `main`.
2. A action **Publicar no Apps Script** verifica a sintaxe e envia tudo com `clasp push --force`.
3. Recarregue a planilha (F5) para ver as mudanças.

> ⚠️ `clasp push --force` **substitui** o projeto Apps Script pelo conteúdo de `src/`.
> Arquivos que existirem só no editor online são apagados. Se alguém editar direto no editor,
> rode antes a action **Baixar fontes do Apps Script** para trazer as mudanças para o Git.

> ℹ️ Cada atualização de implantação cria uma nova versão do script (limite do Google: 200 versões).
> Se não usar a implantação para nada, pode remover o secret `DEPLOYMENT_ID`.

## Link direto (App da Web)

O manifesto publica o painel também como **App da Web** (`doGet`), acessível por quem é do domínio,
executando com a conta de quem acessa:

- `.../exec` → abre o painel na aba **Controle**
- `.../exec?aba=agenda` · `?aba=projeto` · `?aba=cfg` → abre em outra aba

O link aparece em ⚙️ Configurações → *Abrir o Controle direto*. Para existir, é preciso uma
implantação do tipo **App da Web** (Implantar → Nova implantação → App da Web); coloque o ID dela
no secret `DEPLOYMENT_ID` para que cada publicação atualize o link automaticamente.
Pelo link, o agente usa a planilha registrada na última vez que ela foi aberta pelo menu.

Em ⚙️ Configurações também dá para ativar **abrir o painel automaticamente ao abrir a planilha**
(gatilho por usuário).

## Atualização automática e e-mail

Com a atualização automática ativa (⚙️ Configurações ou menu), o agente consulta o Jira **a cada 15 minutos**
e envia **um e-mail** quando algo muda: status do Jira, nº da solicitação, ticket, status do ticket,
resposta do suporte, código PSA, status de controle ou analista. Destinatários em ⚙️ Configurações
(vazio = quem ativou). Na primeira vez, rode **📧 Enviar e-mail de teste** pelo menu para autorizar o envio.

## Uso na planilha

Menu **🦇 Agente A.I.F.R.E.A.D → 🧭 Abrir Painel**.
Na primeira vez: aba ⚙️ Configurações → salvar o Personal Access Token do Jira → 🧪 Testar conexão.
