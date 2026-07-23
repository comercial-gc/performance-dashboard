#!/usr/bin/env python3
"""
Extrai os dados das 4 planilhas do Grupo Cataratas e gera um data.json consolidado
para o painel_parques_rio.html.

A logica de leitura de cada aba foi validada manualmente (celula a celula) durante a
auditoria de dados desta sessao -- inclusive os dois bugs encontrados (Museu de Cera
zerado em Marco/Maio por causa de uma linha duplicada vazia, e Captacao CV em dobro por
somar a linha de total junto com os dias). O codigo abaixo evita os dois de proposito;
os comentarios marcados com "# BUG EVITADO:" explicam onde.

Uso:
    export GOOGLE_APPLICATION_CREDENTIALS=/caminho/para/service-account.json
    python extract_data.py --config config.json --out ../data.json

Requer: google-api-python-client, google-auth (ver requirements.txt)
"""
import argparse
import calendar
import datetime
import json
import sys

from google.oauth2 import service_account
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

MONTH_NUMBER = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARÇO": 3, "ABRIL": 4, "MAIO": 5, "JUNHO": 6,
    "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9, "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12,
}

PARKS = ["AquaRio", "BioParque", "Paineiras", "PNI", "M3F", "AquaFoz", "Três Pescadores", "Vila Velha"]
ATRATIVOS = ["GEX", "MDE", "MDC"]

SHEETS_EPOCH = datetime.date(1899, 12, 30)


def serial_to_date(serial):
    """Converte o numero serial de data do Google Sheets para datetime.date."""
    if serial is None:
        return None
    try:
        return SHEETS_EPOCH + datetime.timedelta(days=float(serial))
    except (TypeError, ValueError):
        return None


def get_sheets_service():
    creds = service_account.Credentials.from_service_account_file(
        _sa_key_path(), scopes=SCOPES
    ) if _sa_key_path() else service_account.Credentials.from_service_account_info(
        json.loads(_sa_key_env()), scopes=SCOPES
    )
    return build("sheets", "v4", credentials=creds)


def _sa_key_path():
    import os
    return os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")


def _sa_key_env():
    import os
    # alternativa: colar o JSON inteiro direto numa env var (usado no workflow do GitHub)
    return os.environ.get("GCP_SERVICE_ACCOUNT_JSON", "{}")


def get_values(service, spreadsheet_id, sheet_name, a1_range="A1:CA2000"):
    """Busca uma aba inteira como lista de listas, com numeros de data como serial."""
    rng = f"'{sheet_name}'!{a1_range}"
    resp = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=rng,
        valueRenderOption="UNFORMATTED_VALUE",
        dateTimeRenderOption="SERIAL_NUMBER",
    ).execute()
    rows = resp.get("values", [])
    # normaliza todas as linhas para o mesmo comprimento (Sheets API omite celulas vazias no fim)
    width = max((len(r) for r in rows), default=0)
    return [r + [None] * (width - len(r)) for r in rows]


def cell(row, idx):
    return row[idx] if idx < len(row) else None


# ---------------------------------------------------------------------------
# Visitação Parques 2026.xlsx (uma aba por mês: JANEIRO..JULHO)
# ---------------------------------------------------------------------------

def parse_month_summary(rows):
    """Linhas 3-10 (index 2-9): resumo por parque. Colunas B-F = mes, H-L = acumulado."""
    summary = {}
    for i, park in enumerate(PARKS):
        r = rows[2 + i]
        summary[park] = {
            "realizado": cell(r, 1),
            "obz": cell(r, 2),
            "pctObz": cell(r, 3),
            "y2025": cell(r, 4),
            "pct2025": cell(r, 5),
            "acumRealizado": cell(r, 7),
            "acumObzParcial": cell(r, 8),
            "acumPctObz": cell(r, 9),
            "acum2025": cell(r, 10),
            "acumPct2025": cell(r, 11),
        }
    return summary


def find_daily_blocks(rows):
    """Acha, para cada parque, a linha 'Realizado 2026' e a linha 'OBZ 2026' logo abaixo.

    O titulo do bloco (nome do parque) fica 1 ou 2 linhas acima -- às vezes tem uma
    linha 'Ações' no meio, e o PNI aparece como 'URBIA + CATARATAS (PNI)' em vez de 'PNI'.
    """
    blocks = {}
    for i, row in enumerate(rows):
        if cell(row, 0) == "Realizado 2026":
            j = i - 1
            while j >= 0 and (not cell(rows[j], 0) or cell(rows[j], 0) == "Ações"):
                j -= 1
            title = str(cell(rows[j], 0) or "")
            norm = None
            for p in PARKS:
                if p.lower().replace("ê", "e") in title.lower().replace("ê", "e"):
                    norm = p
                    break
            if norm and norm not in blocks:
                blocks[norm] = i
    return blocks


def parse_month_daily(rows, n_days):
    daily = {}
    blocks = find_daily_blocks(rows)
    for park, ridx in blocks.items():
        realizado_row = rows[ridx][1:1 + n_days]
        obz_row = rows[ridx + 1][1:1 + n_days] if cell(rows[ridx + 1], 0) == "OBZ 2026" else [None] * n_days
        real2025_row = rows[ridx + 3][1:1 + n_days] if cell(rows[ridx + 3], 0) == "Realizado 2025" else [None] * n_days
        daily[park] = {
            "Realizado 2026": realizado_row,
            "OBZ 2026": obz_row,
            "Realizado 2025": real2025_row,
        }
    return daily


def parse_atrativos_daily(rows, n_days):
    """GEX/MDE/MDC diario.

    # BUG EVITADO: em Marco/Maio existem DUAS linhas 'MDC' na planilha (uma vazia, uma com
    # os dados de verdade, mais abaixo). Por isso sempre ficamos com a ULTIMA linha que
    # tiver algum valor não nulo -- nunca a primeira ocorrência do rótulo.
    """
    daily = {}
    for row in rows:
        label = cell(row, 0)
        if label in ATRATIVOS:
            vals = row[1:1 + n_days]
            if any(v is not None for v in vals):
                daily[label] = vals  # sobrescreve a anterior -> fica a ultima com dado
    return daily


def parse_atrativos_accum(rows):
    """Linhas 3-5 (index 2-4), colunas N-Q (index 13-17): nome, realizado2026, pctAq2026,
    realizado2025, pctAq2025."""
    accum = {}
    for i, a in enumerate(ATRATIVOS):
        r = rows[2 + i]
        name = cell(r, 13)
        if name != a:
            continue
        accum[a] = {
            "realizado2026": cell(r, 14),
            "pctAq2026": cell(r, 15),
            "realizado2025": cell(r, 16) if cell(r, 16) != "-" else None,
            "pctAq2025": cell(r, 17) if cell(r, 17) != "-" else None,
        }
    return accum


def build_visitacao(service, spreadsheet_id, meses_com_dados):
    visitacao = {}
    for mes in meses_com_dados:
        rows = get_values(service, spreadsheet_id, mes)
        month_number = MONTH_NUMBER[mes]
        n_days = calendar.monthrange(2026, month_number)[1]
        visitacao[mes] = {
            "monthNumber": month_number,
            "nDays": n_days,
            "summary": parse_month_summary(rows),
            "daily": parse_month_daily(rows, n_days),
            "atrativos": {
                "daily": parse_atrativos_daily(rows, n_days),
                "accum": parse_atrativos_accum(rows),
            },
        }
    return visitacao


# ---------------------------------------------------------------------------
# CAPTAÇÃO CV - 3P (mesma planilha "Visitação Parques 2026")
# ---------------------------------------------------------------------------

# grupos de colunas por mes: (indice da coluna de data, visitacao, cv), 0-indexed
CAPTACAO_CV_3P_COLS = {
    "JANEIRO": (1, 2, 3), "FEVEREIRO": (6, 7, 8), "MARÇO": (11, 12, 13),
    "ABRIL": (16, 17, 18), "MAIO": (21, 22, 23), "JUNHO": (26, 27, 28), "JULHO": (31, 32, 33),
}


def build_captacao_cv_3p(service, spreadsheet_id, sheet_name):
    rows = get_values(service, spreadsheet_id, sheet_name)
    by_month = {}
    total_vis, total_cv = 0.0, 0.0
    for mes, (dcol, vcol, ccol) in CAPTACAO_CV_3P_COLS.items():
        month_number = MONTH_NUMBER[mes]
        vis_sum, cv_sum = 0.0, 0.0
        for r in rows[2:]:
            d = serial_to_date(cell(r, dcol))
            # BUG EVITADO: filtramos por DATA real do mes, nao por posicao de linha --
            # a planilha tem uma linha de "total do mes" logo apos os dias, que nao tem
            # data preenchida. Se somarmos por posicao de linha (ex.: linhas 3 a 34) esse
            # total entra junto e o resultado sai em dobro.
            if d and d.year == 2026 and d.month == month_number:
                v = cell(r, vcol)
                c = cell(r, ccol)
                if isinstance(v, (int, float)):
                    vis_sum += v
                if isinstance(c, (int, float)):
                    cv_sum += c
        by_month[mes] = {"visitacao": int(vis_sum), "cv": int(cv_sum)}
        total_vis += vis_sum
        total_cv += cv_sum
    return by_month, {"visitacao": int(total_vis), "cv": int(total_cv)}


# ---------------------------------------------------------------------------
# Share E-commerce_Parques Rio - _2026.xlsx -> aba "Share_Ecommerce_2026"
# Serie mensal historica (Jan/2023 em diante) por parque: Visitação total, Ecommerce,
# Share, R$ em mídia. Uma coluna por mes/ano. E' a fonte de "investimentoMidia.meses".
# ---------------------------------------------------------------------------

# linha (0-indexed) onde comeca o bloco de cada parque nesta aba
SHARE_ECOMMERCE_BLOCKS = {
    "BioParque": 0, "AquaRio": 9, "Paineiras": 17, "M3F": 25, "AquaFoz": 33,
}
MESES_PT = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho",
            "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]


def build_investimento_midia(service, spreadsheet_id, sheet_name, meses_com_dados):
    """Monta SHARE.investimentoMidia.meses para os 5 parques que tem e-commerce
    (AquaRio, BioParque, Paineiras, M3F, AquaFoz).

    Para meses fechados (todo mes exceto o corrente), tanto o lado 2025 quanto o 2026 sao
    o total do mes inteiro. Para o mes corrente (ainda em andamento), o lado 2026 e'
    parcial (só os dias já lançados na planilha) e o 2025 e' o mes inteiro do ano passado
    -- usado como referencia de "para onde estamos indo", nao como comparação dia-a-dia.

    NOTA / GAP CONHECIDO: os parques "3P" (Três Pescadores) e "Vila Velha" nao tem
    e-commerce/share rastreados nesta planilha (a coluna "R$ em mídia" deles fica vazia
    nos meses de 2026). Nao encontrei, em nenhuma das 4 planilhas, uma fonte confiavel
    para o investimento de midia desses dois parques especificamente -- por isso eles
    ficam de fora deste dicionario por enquanto. Se voce souber onde esse numero e'
    controlado (pode ser uma planilha/aba que nao foi anexada), me avise que eu mapeio.
    """
    rows = get_values(service, spreadsheet_id, sheet_name)
    meses = {}
    for i, mes_en in enumerate(meses_com_dados):
        mes_idx = MONTH_NUMBER[mes_en] - 1  # 0 = Janeiro
        mes_pt = MESES_PT[mes_idx]
        idx_2026 = 37 + mes_idx  # coluna do mes/ano em 2026 (Jan/2026 comeca no indice 37)
        idx_2025 = 25 + mes_idx  # coluna do mes/ano em 2025 (Jan/2025 comeca no indice 25)
        meses[mes_pt] = {}
        for park, r0 in SHARE_ECOMMERCE_BLOCKS.items():
            vis_row, ecom_row, share_row, inv_row = rows[r0 + 1], rows[r0 + 2], rows[r0 + 3], rows[r0 + 4]

            def num(row, idx):
                v = cell(row, idx)
                return v if isinstance(v, (int, float)) else None

            vis26, ecom26 = num(vis_row, idx_2026), num(ecom_row, idx_2026)
            vis25, ecom25 = num(vis_row, idx_2025), num(ecom_row, idx_2025)
            meses[mes_pt][park] = {
                "visitacao2026": vis26,
                "ecommerce2026": ecom26,
                "share2026": (ecom26 / vis26) if (vis26 and ecom26 is not None) else num(share_row, idx_2026),
                "visitacao2025": vis25,
                "ecommerce2025": ecom25,
                "share2025": (ecom25 / vis25) if vis25 else 0,
                "investimento2026": num(inv_row, idx_2026),
                "investimento2025": num(inv_row, idx_2025),
            }
    return meses


def build_evolucao_mensal(service, spreadsheet_id, sheet_name, ano_inicio=2025, mes_inicio=1, ano_fim=2026, mes_fim=7):
    """Serie historica mes a mes (investimento, share, visitacaoTotal) de Jan/2025 ate o
    mes/ano atual, mesma aba "Share_Ecommerce_2026" — e' o mesmo dado de
    build_investimento_midia, só que olhando pra tras (nao comparando 2026 vs 2025 lado a
    lado, e sim uma linha do tempo unica)."""
    rows = get_values(service, spreadsheet_id, sheet_name)
    labels = []
    idx_por_label = []
    y, m = ano_inicio, mes_inicio
    while (y, m) <= (ano_fim, mes_fim):
        labels.append(f"{m:02d}/{y % 100:02d}")
        idx_por_label.append(25 + (y - 2025) * 12 + (m - 1))
        m += 1
        if m > 12:
            m = 1
            y += 1

    parques = {}
    for park, r0 in SHARE_ECOMMERCE_BLOCKS.items():
        vis_row, share_row, inv_row = rows[r0 + 1], rows[r0 + 3], rows[r0 + 4]

        def num(row, idx):
            v = cell(row, idx)
            return v if isinstance(v, (int, float)) else None

        parques[park] = {
            "investimento": [num(inv_row, i) for i in idx_por_label],
            "share": [num(share_row, i) for i in idx_por_label],
            "visitacaoTotal": [num(vis_row, i) for i in idx_por_label],
        }
    return {"labels": labels, "parques": parques}


def build_share_meta_grupo_cataratas(investimento_midia_meses, meses_com_dados):
    """SHARE_META_GRUPO_CATARATAS: nao precisa de nenhuma planilha nova -- e' a soma dos 5
    parques com e-commerce (AquaRio, BioParque, Paineiras, M3F, AquaFoz) que ja lemos em
    build_investimento_midia. Antes esse bloco vivia hard-coded no HTML; agora e' calculado.

    OBSERVAÇÃO: comparando com os números fixos que estavam no HTML, Julho bate exatamente,
    mas alguns meses mais antigos (ex.: Junho) batem diferente -- a AquaFoz aparenta ter
    passado a ser rastreada no e-commerce só a partir de um certo mês, e o valor antigo
    hard-coded parece ter sido somado sem a AquaFoz nesses meses. Como não dá pra saber com
    certeza, a partir de agora, a soma é sempre com os 5 parques (mais transparente e
    consistente pra frente) -- isso pode mudar levemente os meses fechados mais antigos do
    gráfico de meta, mas o mês mais recente (o que importa pra acompanhar o dia a dia)
    sempre bate."""
    parks = ["AquaRio", "BioParque", "Paineiras", "M3F", "AquaFoz"]
    visitacao, ecommerce, share = [], [], []
    for mes_en in meses_com_dados:
        mes_pt = MESES_PT[MONTH_NUMBER[mes_en] - 1]
        d = investimento_midia_meses.get(mes_pt, {})
        vis_total = sum((d.get(p, {}).get("visitacao2026") or 0) for p in parks)
        ecom_total = sum((d.get(p, {}).get("ecommerce2026") or 0) for p in parks)
        visitacao.append(vis_total)
        ecommerce.append(ecom_total)
        share.append(ecom_total / vis_total if vis_total else None)
    return {"visitacao": visitacao, "ecommerce": ecommerce, "share": share}


# ---------------------------------------------------------------------------
# [2026] Mix OBZ e visitação.xlsx -> aba "AQF E M3F | SMorador"
# Proporcao "Sem morador" / "Com morador" mais recente, usada para estimar a Captação
# PNI "sem morador" (ver comentario original no HTML sobre SEMMORADOR_RATIO).
# ---------------------------------------------------------------------------

def build_semmorador_ratio(service, spreadsheet_id, sheet_name):
    rows = get_values(service, spreadsheet_id, sheet_name)
    # linha 6 (index5) = % C/Morador do mes corrente; linha 7 (index6) = % S/Morador do mes
    # corrente. M3F nas colunas 0-1, AquaFoz nas colunas 4-5 (ver aba "MÊS <mes atual>").
    com_m3f, sem_m3f = cell(rows[5], 0), cell(rows[6], 1)
    com_aqf, sem_aqf = cell(rows[5], 4), cell(rows[6], 5)
    return {
        "M3F": (sem_m3f / com_m3f) if com_m3f else None,
        "AquaFoz": (sem_aqf / com_aqf) if com_aqf else None,
    }


# ---------------------------------------------------------------------------
# Share E-commerce_Parques Rio - _2026.xlsx -> aba "Dash Share GC"
# ---------------------------------------------------------------------------

def build_dash_share_gc(service, spreadsheet_id, sheet_name):
    rows = get_values(service, spreadsheet_id, sheet_name)
    periodo = cell(rows[0], 1)
    nomes = {"Aquario": 4, "Bioparque": 5, "Paineiras": 6, "Marco": 7, "Aquafoz": 8, "Grupo Cataratas": 10}
    parques = []
    for nome, ridx in nomes.items():
        r = rows[ridx]
        parques.append({
            "parque": nome,
            "metaAnual": cell(r, 1),
            "ecommerce2026": cell(r, 2),
            "visitacao2026": cell(r, 3),
            "share2026": cell(r, 4),
            "ecommerce2025": cell(r, 5),
            "visitacao2025": cell(r, 6),
            "share2025": cell(r, 7),
            "deltaPP": cell(r, 8),
            "gapMeta": cell(r, 9),
        })
    return {"periodo": periodo, "parques": parques}


# ---------------------------------------------------------------------------
# INVESTIMENTO MARKETING _ 2026.xlsx -> aba "acompanhamento mkt"
# ---------------------------------------------------------------------------

BLOCK_NAME_MAP = {
    "AQUARIO": "AquaRio", "BIOPARQUE": "BioParque", "PAINEIRAS": "Paineiras",
    "AQUAFOZ": "AquaFoz", "M3F": "M3F", "VILA VELHA": "Vila Velha",
    "3 PESCADORES": "Três Pescadores",
}


def build_invest_mkt_resumo(service, spreadsheet_id, sheet_name):
    rows = get_values(service, spreadsheet_id, sheet_name)
    resumo = {}
    for i, row in enumerate(rows):
        for c, val in enumerate(row):
            if val in BLOCK_NAME_MAP:
                park = BLOCK_NAME_MAP[val]
                resumo[park] = {}
                r = i + 2
                while r < len(rows):
                    mes = cell(rows[r], c)
                    if not mes:
                        break
                    disp = cell(rows[r], c + 1)
                    real = cell(rows[r], c + 2)
                    saldo = cell(rows[r], c + 3)
                    resumo[park][str(mes).strip().upper()] = {
                        "disponivel": disp if disp not in ("", "#N/A") else None,
                        "realizado": real if real not in ("", "#N/A") else None,
                        "saldo": saldo if saldo not in ("", "#N/A") else None,
                    }
                    r += 1
    return resumo


# ---------------------------------------------------------------------------
# INVESTIMENTO MARKETING _ 2026.xlsx -> uma aba por mes (JANEIRO..JULHO), lista
# de campanhas/linhas de gasto. O cabecalho MUDA de posicao de mes para mes
# (ex.: "SETOR" vira "CUSTO" em Junho/Julho, "RUNRUN IT" some em alguns meses),
# entao procuramos as colunas pelo NOME do cabecalho em vez de por indice fixo.
# ---------------------------------------------------------------------------

# nomes possiveis de cabecalho -> campo de saida (primeiro que bater, na ordem da linha)
DETAIL_HEADER_ALIASES = {
    "parque": ["PARQUE"],
    "setor": ["SETOR", "CUSTO"],
    "fornecedor": ["FORNECEDOR"],
    "descricao": ["DESCRIÇAO DO SERVIÇO", "DESCRIÇÃO DO SERVIÇO"],
    "runrun": ["RUNRUN IT"],
    "valor": ["VALOR"],
    "mesCompetencia": ["MêS DE COMPETÊNCIA", "MÊS DE COMPETÊNCIA"],
    "observacao": ["OBSERVAÇÃO"],
}


def _find_detail_columns(header_row):
    cols = {}
    for c, raw in enumerate(header_row):
        if not raw:
            continue
        label = str(raw).strip().upper()
        for field, aliases in DETAIL_HEADER_ALIASES.items():
            if field in cols:
                continue
            if label in [a.upper() for a in aliases]:
                cols[field] = c
    return cols


def _parse_valor(v):
    """VALOR vem ora como numero, ora como texto 'R$ 1.780,00' -- normaliza pros dois casos."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("R$", "").strip()
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def build_invest_mkt_detail(service, spreadsheet_id, meses_com_dados):
    """Le a lista de campanhas de cada aba mensal. Retorna {mes: [ {parque, setor,
    fornecedor, descricao, valor, observacao}, ... ]}.
    """
    detail = {}
    for mes in meses_com_dados:
        rows = get_values(service, spreadsheet_id, mes)
        if not rows:
            detail[mes] = []
            continue
        cols = _find_detail_columns(rows[0])
        if "parque" not in cols or "valor" not in cols:
            detail[mes] = []
            continue
        items = []
        for row in rows[1:]:
            parque = cell(row, cols.get("parque"))
            if not parque:
                continue
            items.append({
                "parque": str(parque).strip(),
                "setor": cell(row, cols.get("setor")) if "setor" in cols else None,
                "fornecedor": cell(row, cols.get("fornecedor")) if "fornecedor" in cols else None,
                "descricao": cell(row, cols.get("descricao")) if "descricao" in cols else None,
                "valor": _parse_valor(cell(row, cols.get("valor"))),
                "observacao": cell(row, cols.get("observacao")) if "observacao" in cols else None,
            })
        detail[mes] = items
    return detail


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    with open(args.config, encoding="utf-8") as f:
        cfg = json.load(f)

    service = get_sheets_service()

    print("Lendo Visitação Parques 2026...", file=sys.stderr)
    visitacao = build_visitacao(service, cfg["visitacao_parques_id"], cfg["meses_com_dados"])

    print("Lendo Captação CV - 3P...", file=sys.stderr)
    cv3p_by_month, cv3p_anual = build_captacao_cv_3p(
        service, cfg["visitacao_parques_id"], cfg["sheet_names"]["captacao_cv_3p"]
    )

    print("Lendo Dash Share GC...", file=sys.stderr)
    dash_share_gc = build_dash_share_gc(
        service, cfg["share_ecommerce_id"], cfg["sheet_names"]["dash_share_gc"]
    )

    print("Lendo Share_Ecommerce_2026 (investimentoMidia)...", file=sys.stderr)
    investimento_midia = build_investimento_midia(
        service, cfg["share_ecommerce_id"], cfg["sheet_names"]["share_ecommerce_2026"],
        cfg["meses_com_dados"]
    )

    print("Lendo acompanhamento mkt...", file=sys.stderr)
    invest_mkt_resumo = build_invest_mkt_resumo(
        service, cfg["investimento_marketing_id"], cfg["sheet_names"]["acompanhamento_mkt"]
    )

    print("Lendo detalhe de campanhas (Investimento Marketing, mes a mes)...", file=sys.stderr)
    invest_mkt_detail = build_invest_mkt_detail(
        service, cfg["investimento_marketing_id"], cfg["meses_com_dados"]
    )

    print("Lendo evolução mensal (série histórica)...", file=sys.stderr)
    evolucao_mensal = build_evolucao_mensal(
        service, cfg["share_ecommerce_id"], cfg["sheet_names"]["share_ecommerce_2026"]
    )

    print("Calculando Share Meta Grupo Cataratas...", file=sys.stderr)
    share_meta_grupo_cataratas = build_share_meta_grupo_cataratas(
        investimento_midia, cfg["meses_com_dados"]
    )

    print("Lendo Captação PNI Sem Morador...", file=sys.stderr)
    semmorador_ratio = build_semmorador_ratio(
        service, cfg["mix_obz_visitacao_id"], cfg["sheet_names"]["smorador"]
    )

    output = {
        "geradoEm": datetime.datetime.utcnow().isoformat() + "Z",
        "VISITACAO": visitacao,
        "CAPTACAO_CV_3P_BY_MONTH": cv3p_by_month,
        "CAPTACAO_CV_3P_ANUAL": cv3p_anual,
        "SEMMORADOR_RATIO": semmorador_ratio,
        "SHARE_META_MESES": [MESES_PT[MONTH_NUMBER[m] - 1] for m in cfg["meses_com_dados"]],
        "SHARE_META_GRUPO_CATARATAS": share_meta_grupo_cataratas,
        "SHARE": {
            "dashShareGC": dash_share_gc,
            "investimentoMidia": {"meses": investimento_midia},
            "evolucaoMensal": evolucao_mensal,
        },
        "INVEST_MKT_RESUMO": invest_mkt_resumo,
        "INVEST_MKT_DETAIL": invest_mkt_detail,
    }

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"OK: {args.out} gerado.", file=sys.stderr)


if __name__ == "__main__":
    main()
