# Pipeline de dados do Painel Performance Dashboard

Objetivo: parar de colar números à mão no HTML (`painel_parques_rio.html`) e passar a gerar
um `data.json` automaticamente, todo dia, direto das 4 planilhas oficiais, via GitHub Actions.

Isso elimina a classe inteira de erro que caçamos manualmente nesta sessão (Museu de Cera
zerado em Março/Maio, Captação CV em dobro, Share E-commerce desatualizado) — porque o dado
passa a ser recalculado do zero a cada execução, e qualquer coisa estranha aparece no `git diff`
do arquivo `data.json` a cada commit, em vez de ficar escondida dentro de um número copiado.

## Visão geral da arquitetura

```
[4 planilhas no Google Drive]
        │  (Sheets API, leitura)
        ▼
[scripts/extract_data.py]  ← roda dentro de um GitHub Action agendado (cron)
        │  gera
        ▼
[data.json]  ← commitado automaticamente no repositório
        │  lido via fetch() pelo navegador
        ▼
[painel_parques_rio.html]  ← hospedado no GitHub Pages (ou onde você já hospeda hoje)
```

Nenhuma peça nova de infraestrutura: só o repositório que você já tem acesso agora, mais uma
service account do Google (gratuita) para autenticar a leitura das planilhas.

---

## Passo 1 — Criar a Service Account no Google Cloud

Uma "service account" é uma conta-robô que só tem permissão de leitura nas 4 planilhas — não é
sua conta pessoal, então não usa sua senha nem expõe seu acesso total ao Drive.

1. Acesse https://console.cloud.google.com/ com a conta Google da empresa (ou uma pessoal, tanto
   faz — a service account não precisa estar ligada à conta dona das planilhas).
2. Crie um projeto novo (ex.: `cataratas-dashboard`).
3. No menu **APIs e serviços → Biblioteca**, ative a **Google Sheets API**.
4. Vá em **APIs e serviços → Credenciais → Criar credenciais → Conta de serviço**.
   - Nome: `cataratas-dashboard-reader`.
   - Não precisa dar nenhum papel (role) de projeto — ela só vai ler planilhas.
5. Depois de criada, abra a service account, vá na aba **Chaves (Keys) → Adicionar chave →
   Criar nova chave → JSON**. Isso baixa um arquivo `.json` com credenciais.
   **Esse arquivo é uma senha — nunca commitar ele no repositório.**
6. Anote o e-mail da service account (algo como
   `cataratas-dashboard-reader@cataratas-dashboard.iam.gserviceaccount.com`).

## Passo 2 — Dar acesso de leitura às 4 planilhas

A service account só enxerga o que for compartilhado com ela, planilha por planilha:

1. Abra cada uma das 4 planilhas (Visitação Parques 2026, Mix OBZ e visitação, Share
   E-commerce, Investimento Marketing).
2. Clique em **Compartilhar** e cole o e-mail da service account (passo 1.6).
3. Permissão: **Leitor (Viewer)** — o script nunca precisa escrever nelas.
4. Guarde o **ID de cada planilha** (a parte da URL entre `/d/` e `/edit`):
   `https://docs.google.com/spreadsheets/d/AQUI_ESTA_O_ID/edit`

## Passo 3 — Guardar o segredo no GitHub

1. No repositório, vá em **Settings → Secrets and variables → Actions → New repository
   secret**.
2. Nome: `GCP_SERVICE_ACCOUNT_JSON`.
3. Valor: cole o **conteúdo inteiro** do arquivo `.json` baixado no passo 1.5.
4. Crie também 4 secrets (ou variables, não são sigilosos) com os IDs das planilhas — ou,
   mais simples, deixe os IDs direto no `config.json` do repositório (não são segredo, só o
   acesso é controlado pelo compartilhamento do passo 2).

## Passo 4 — Estrutura de pastas no repositório

```
seu-repo/
├── .github/
│   └── workflows/
│       └── update-data.yml
├── scripts/
│   ├── extract_data.py
│   ├── requirements.txt
│   └── config.json
├── data.json                 ← gerado automaticamente, não editar à mão
└── painel_parques_rio.html
```

## Passo 5 — Testar localmente antes de automatizar

Antes de confiar no Action, rode na sua máquina:

```bash
pip install -r scripts/requirements.txt
export GOOGLE_APPLICATION_CREDENTIALS=/caminho/para/a-chave-baixada.json
python scripts/extract_data.py --config scripts/config.json --out data.json
```

Abra o `data.json` gerado e compare alguns números com a planilha aberta ao lado — o mesmo
tipo de double-check que fizemos manualmente nesta conversa, só que reprodutível.

## Passo 6 — Ativar o GitHub Action

O workflow em `.github/workflows/update-data.yml` já vem configurado para:
- Rodar todo dia às 7h (horário de Brasília) — ajuste o cron se quiser outro horário.
- Também poder ser disparado manualmente (`workflow_dispatch`), para quando você quiser forçar
  uma atualização na hora.
- Rodar o script, e se o `data.json` mudou, commitar e dar push automaticamente.

## Passo 7 — O painel passa a ler o `data.json`

Hoje o HTML tem os dados colados dentro de `<script>` (`const VISITACAO = {...}`, `const SHARE
= {...}` etc). O próximo passo (depois que o pipeline estiver rodando e você validar um tempo em
paralelo) é trocar esses `const` fixos por um `fetch('data.json')` no carregamento da página —
veja `dashboard-integration.md` neste mesmo pacote para o trecho de código.

**Recomendação:** não troque os dois de uma vez. Rode o pipeline gerando `data.json` por 1–2
semanas em paralelo ao HTML atual, comparando os números manualmente algumas vezes, e só depois
troque o HTML para consumir o JSON. Isso evita que um bug no script derrube o painel de vez.

## O que este script cobre (testado contra as 4 planilhas reais, número por número)

- Resumo mensal e acumulado por parque (Realizado, OBZ, %, 2025, %) — aba "Visitação Parques
  2026.xlsx", uma sheet por mês.
- Visitação diária de GEX / Mar de Espelhos / Museu de Cera, já tratando o caso de linha
  duplicada vazia que causou o bug do Museu de Cera em Março/Maio (o script sempre pega a
  **última** ocorrência da linha com dado, não a primeira).
- Captação CV (Três Pescadores) mensal e anual, somando só os dias reais do mês (o bug do "valor
  em dobro" que achamos veio de somar a linha de total junto com os dias — o script evita isso
  filtrando por data, não por posição de linha).
- Snapshot geral de Share E-commerce (aba "Dash Share GC").
- `investimentoMidia.meses` (visitação/e-commerce/share/investimento em mídia, mês a mês, para
  os 5 parques com e-commerce) — fonte: aba "Share_Ecommerce_2026", que tem uma série mensal
  histórica por parque desde 2023. Ao mapear essa aba encontrei e já corrigi no
  `painel_parques_rio.html` uma inconsistência: o mês de Julho estava comparando com um "2025"
  errado (visitação e e-commerce de referência não batiam com o mês cheio de julho/2025) —
  agora usa o mesmo critério dos outros meses.
- Resumo de investimento em marketing (disponível/realizado/saldo por parque e mês) — aba
  "acompanhamento mkt".
- `INVEST_MKT_DETAIL` — lista linha a linha de cada campanha/fornecedor por mês, das 7 abas
  mensais da planilha de Investimento Marketing. O cabeçalho dessas abas muda de posição mês a
  mês (ex.: "SETOR" vira "CUSTO" em Junho/Julho) — o script procura as colunas pelo nome, não
  pela posição, então não quebra se isso mudar de novo. Conferi: a soma dessa lista por parque
  bate exatamente com o "Realizado" do resumo acima, em todos os parques de Julho.

## Gap conhecido (não inventei número para isso)

Os parques **Três Pescadores** e **Vila Velha** têm um valor de investimento em mídia mensal no
painel atual (ex.: R$ 5.001,39 e R$ 5.314,67 em Julho) que **não encontrei em nenhuma das 4
planilhas anexadas** — não está na aba de e-commerce (esses dois parques não têm e-commerce
rastreado ali) nem na lista de campanhas da planilha de Investimento Marketing (eles simplesmente
não aparecem nela). Por isso, no `data.json` gerado, esses dois parques ficam sem
`investimento2025`/`investimento2026` por enquanto, em vez de eu inventar um número parecido.
Se você souber em qual planilha/aba esse valor é controlado, me diga e eu mapeio — é rápido,
só precisava da fonte certa.
