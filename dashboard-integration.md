# Como trocar os `const` fixos pelo `data.json`

Hoje o `painel_parques_rio.html` tem, dentro da tag `<script>`, blocos assim:

```js
const VISITACAO = { "JANEIRO": {...}, "FEVEREIRO": {...}, ... };
const SHARE = { "dashShareGC": {...}, "investimentoMidia": {...}, "evolucaoMensal": {...} };
const INVEST_MKT_RESUMO = { "AquaRio": {...}, ... };
const CAPTACAO_CV_3P_BY_MONTH = {...};
const CAPTACAO_CV_3P_ANUAL = {...};
```

O `extract_data.py` já gera um `data.json` com **todas** essas chaves preenchidas:
`VISITACAO`, `CAPTACAO_CV_3P_BY_MONTH`, `CAPTACAO_CV_3P_ANUAL`, `SEMMORADOR_RATIO`,
`SHARE_META_MESES`, `SHARE_META_GRUPO_CATARATAS`, `SHARE.dashShareGC`,
`SHARE.investimentoMidia`, `SHARE.evolucaoMensal`, `INVEST_MKT_RESUMO`, `INVEST_MKT_DETAIL`.
A única coisa que fica de fora, por enquanto, é o investimento de mídia de Três Pescadores e
Vila Velha — ver a seção "Gap conhecido" no README.

## Passo a passo da troca

**1.** Apague as linhas de todos os `const` fixos que o `data.json` agora cobre:
`VISITACAO`, `CAPTACAO_CV_3P_BY_MONTH`, `CAPTACAO_CV_3P_ANUAL`, `SEMMORADOR_RATIO`,
`SHARE_META_MESES`, `SHARE_META_GRUPO_CATARATAS`, `SHARE` (o objeto inteiro), `INVEST_MKT_RESUMO`,
`INVEST_MKT_DETAIL`.

**2.** No lugar, declare como `let`/`const` vazios e preencha via `fetch` antes de chamar as
funções de render que já existem hoje:

```js
let VISITACAO = {};
let CAPTACAO_CV_3P_BY_MONTH = {};
let CAPTACAO_CV_3P_ANUAL = {};
let SEMMORADOR_RATIO = {};
let SHARE_META_MESES = [];
let SHARE_META_GRUPO_CATARATAS = {};
let SHARE = {};
let INVEST_MKT_RESUMO = {};
let INVEST_MKT_DETAIL = {};

async function carregarDados() {
  const resp = await fetch('data.json', { cache: 'no-store' });
  const data = await resp.json();

  VISITACAO = data.VISITACAO;
  CAPTACAO_CV_3P_BY_MONTH = data.CAPTACAO_CV_3P_BY_MONTH;
  CAPTACAO_CV_3P_ANUAL = data.CAPTACAO_CV_3P_ANUAL;
  SEMMORADOR_RATIO = data.SEMMORADOR_RATIO;
  SHARE_META_MESES = data.SHARE_META_MESES;
  SHARE_META_GRUPO_CATARATAS = data.SHARE_META_GRUPO_CATARATAS;
  SHARE = data.SHARE;
  INVEST_MKT_RESUMO = data.INVEST_MKT_RESUMO;
  INVEST_MKT_DETAIL = data.INVEST_MKT_DETAIL;

  // a partir daqui, chame o que hoje já roda direto no carregamento da pagina:
  refreshVisitacao();
  renderInvestRegua();
  renderShareMetaChart();
  // ...(as demais chamadas de render que já existem no fim do script)
}

carregarDados();
```

**3.** Remova a chamada direta a essas funções de render que hoje roda solta no fim do
script (ex.: `refreshVisitacao();` fora de qualquer função) — elas passam a rodar só depois
que o `fetch` terminar, dentro de `carregarDados()`.

**Atenção:** como `VISITACAO`, `SHARE` etc. viram `let` em vez de `const`, procure no HTML se
alguma função declara `const VISITACAO` de novo mais abaixo (duplicar `let`/`const` do mesmo
nome quebra o carregamento). Se aparecer, é só apagar a segunda declaração.

**4.** Sirva o HTML e o `data.json` do mesmo lugar (mesma pasta do repositório / mesmo site do
GitHub Pages), porque `fetch('data.json')` é um caminho relativo — se um dia o HTML for
hospedado em outro domínio, troque para a URL completa do `data.json` (ex.: o "raw" do
GitHub, ou o Pages).

## Por que fazer em duas etapas (pipeline primeiro, HTML depois)

Se trocar tudo de uma vez e o `data.json` vier com algum problema (chave faltando, service
account sem acesso a uma das planilhas, etc.), o painel para de funcionar por inteiro. Rodando
o pipeline em paralelo por um tempo, gerando o `data.json` mas o HTML ainda usando os `const`
fixos, você consegue comparar os dois lado a lado sem risco — só troca de fato quando confiar.
