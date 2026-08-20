# -*- coding: utf-8 -*-
"""
Extensão do extract_data.py: Cross AquaFoz (vendas cruzadas AquaFoz / PNI / Marco das 3
Fronteiras)
============================================================================================

CONTEXTO
--------
Pedido do usuário (20/08/2026): trazer pro Performance Dashboard, numa aba nova ("Cross
AquaFoz"), o relatório que hoje ele monta manualmente (a partir da planilha "Vendas Cross
para AquaFoz 2026") pra enviar pra Camila -- duas visões:

  - Diária: divisão por canal de venda (Marco das 3 Fronteiras: Bilheteria/Trade balcão/
    Cross; Aeroporto; PNI: Cross/Totem/Combo), Total cross, Share de visitação AQF/PNI,
    Visitação AquaFoz/PNI, Captação AQF x PNI -- tudo do MÊS VIGENTE, dia a dia.
  - Mensal: série mês a mês (Jan-Dez) com Vendas total, Vendas Totem PNI, Vendas Combo PNI,
    Visitação AQF, Share visitação AQF, Visitação PNI, Captação PNI, Captação AQF x PNI.

FONTE: planilha Google Sheets "Vendas Cross para AquaFoz 2026"
(id em cfg["cross_aquafoz_id"], ver COMO INTEGRAR no fim do arquivo).

  - Aba "Venda mes a mes": UMA linha por métrica, UMA coluna por mês (Jan..Dez, só até o mês
    mais recente com dado) -- fonte da visão Mensal. Achado pelo RÓTULO de cada linha
    (coluna A) e pelo NOME do mês no cabeçalho, não por número de linha/coluna fixo, pra não
    quebrar se alguém inserir uma linha/coluna em cima (mesmo padrão defensivo do resto do
    extract_data.py -- ver build_mix_origem_diario etc.).

  - Aba com o NOME DO MÊS VIGENTE (ex.: "Agosto", "Julho") -- uma linha por DIA do mês,
    colunas: Bilheteria MF3, Trade MF3, Aeroporto, Cross MF3, Cross PNI, Totem PNI, Combo
    PNI, Total cross, Visitação PNI, Capt. Diária PNI, Visitação AQF, Capt Total AQF x PNI
    -- fonte da visão Diária.

    IMPORTANTE -- conferido manualmente em Agosto/2026: essa aba tem também um "painel
    resumo" solto nas colunas à direita (a partir da coluna P) com números DIGITADOS À MÃO
    (ex.: "01/08 a 18/08" seguido do total daquele intervalo) -- não é fórmula viva, fica
    desatualizado a cada dia que passa sem alguém atualizar na mão. Esse módulo IGNORA esse
    painel de propósito e sempre recalcula os totais a partir das colunas diárias (A-O, que
    são alimentadas dia a dia e são a fonte confiável) -- o front-end (index.html) que soma
    o período que o usuário selecionar, do mesmo jeito que as outras abas do painel já
    fazem (ver getMixPrimaryRange/getMixTop5 como referência de padrão).

  - Nomes de aba por mês: Title Case em português (ex. "Agosto", não "AGOSTO" nem
    "agosto") -- mas o CABEÇALHO da aba "Venda mes a mes" vem com capitalização
    inconsistente (“Janeiro”, mas depois “fevereiro”, “março” minúsculo etc.) -- por isso
    _canonical_mes() normaliza tudo (case + acento) antes de comparar.

  - Esquema da aba diária pode mudar mês a mês (conferido: Jan-Junho usava só "Vendas
    Bilhetria"/"Vendas Comercial", sem a quebra por Bilheteria MF3/Trade MF3/Aeroporto/Cross
    MF3/Cross PNI/Totem PNI/Combo PNI que passou a existir a partir de Julho/Agosto) -- por
    isso as colunas são lidas pelo NOME do cabeçalho (DIARIO_COLUNAS), não por posição fixa;
    uma coluna que não existir naquele mês simplesmente não entra no resultado daquele dia
    (fica de fora do dict, front-end trata como ausente/0 conforme o caso).
"""

import datetime
import unicodedata


def _norm(s):
    """Normaliza rótulo de célula (cabeçalho de coluna, nome de mês, rótulo de linha) pra
    comparação: remove acento, baixa a caixa, colapsa espaço. Não usado para os VALORES
    numéricos -- ver _to_num."""
    if s is None:
        return ""
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return " ".join(s.lower().split())


def _to_num(v):
    """Converte valor de célula pra número, tratando 'NA'/'N/A'/'#DIV/0!'/vazio como None
    (em vez de virar 0 escondido ou propagar erro) -- diferença importante pro front-end
    saber quando o dado simplesmente não existe ainda (ex.: Combo PNI não existia nos
    primeiros meses do ano; dias futuros do mês vigente ainda sem visitação lançada)."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    if not s or s.upper() in ("NA", "N/A", "#DIV/0!", "#N/A", "#REF!", "#VALUE!"):
        return None
    try:
        return float(s)
    except ValueError:
        try:
            return float(s.replace(".", "").replace(",", "."))
        except ValueError:
            return None


MES_ALIASES = {
    "janeiro": "Janeiro", "fevereiro": "Fevereiro", "marco": "Março", "abril": "Abril",
    "maio": "Maio", "junho": "Junho", "julho": "Julho", "agosto": "Agosto",
    "setembro": "Setembro", "outubro": "Outubro", "novembro": "Novembro", "dezembro": "Dezembro",
}


def _canonical_mes(label):
    return MES_ALIASES.get(_norm(label))


# Rótulo da linha (coluna A) na aba "Venda mes a mes" -> chave no dict de saída.
MENSAL_METRICAS = {
    "vendas todos canais": "vendasTotal",
    "vendas totem pni": "vendasTotemPni",
    "vendas combo pni": "vendasComboPni",
    "visitacao aqf": "visitacaoAqf",
    "share visitacao aqf": "shareVisitacaoAqf",
    "visitacao pni": "visitacaoPni",
    "captacao pni": "captacaoPni",
    "captacao aqf x pni": "captacaoAqfPni",
}

# Nome da coluna (cabeçalho, linha 1) na aba do mês vigente -> chave no dict de saída.
DIARIO_COLUNAS = {
    "bilheteria mf3": "bilheteriaMf3",
    "trade mf3": "tradeMf3",
    "aeroporto": "aeroporto",
    "cross mf3": "crossMf3",
    "cross pni": "crossPni",
    "totem pni": "totemPni",
    "combo pni": "comboPni",
    "total cross": "totalCross",
    "visitacao pni": "visitacaoPni",
    "capt. diaria pni": "captDiariaPni",
    "visitacao aqf": "visitacaoAqf",
    "capt total aqf x pni": "captTotalAqfPni",
}


def build_cross_aquafoz_mensal(rows):
    """Lê a aba 'Venda mes a mes' (lista de listas, uma por linha da planilha) e devolve
    {mes: {vendasTotal, vendasTotemPni, vendasComboPni, visitacaoAqf, shareVisitacaoAqf,
    visitacaoPni, captacaoPni, captacaoAqfPni}} -- só com os meses que aparecerem no
    cabeçalho (tipicamente Jan até o mês vigente)."""
    header_row = None
    for row in rows:
        if row and _norm(row[0]) == "metrica":
            header_row = row
            break
    if header_row is None:
        raise ValueError("linha de cabeçalho ('Métrica') não encontrada na aba 'Venda mes a mes'")

    col_mes = {}
    for i, v in enumerate(header_row):
        mes = _canonical_mes(v)
        if mes:
            col_mes[i] = mes

    result = {mes: {} for mes in col_mes.values()}
    for row in rows:
        if not row:
            continue
        key = MENSAL_METRICAS.get(_norm(row[0]))
        if not key:
            continue
        for i, mes in col_mes.items():
            result[mes][key] = _to_num(row[i]) if i < len(row) else None
    return result


def build_cross_aquafoz_diario(rows):
    """Lê a aba do mês vigente (ex. 'Agosto') -- uma linha por dia, colunas A-O. O painel
    resumo solto em P+ (ver docstring do módulo) é ignorado de propósito. Devolve
    {'YYYY-MM-DD': {bilheteriaMf3, tradeMf3, aeroporto, crossMf3, crossPni, totemPni,
    comboPni, totalCross, visitacaoPni, captDiariaPni, visitacaoAqf, captTotalAqfPni}} --
    só as chaves cuja coluna existir de fato naquele mês (esquema já variou mês a mês, ver
    docstring do módulo)."""
    if not rows:
        return {}
    header_row = rows[0]
    col_map = {}
    col_data = None
    for i, v in enumerate(header_row):
        nv = _norm(v)
        if nv == "data":
            col_data = i
        key = DIARIO_COLUNAS.get(nv)
        if key:
            col_map[i] = key
    if col_data is None or not col_map:
        raise ValueError("colunas esperadas (Data + métricas) não encontradas na aba do mês vigente")

    result = {}
    for row in rows[1:]:
        if not row or col_data >= len(row):
            continue
        dt_raw = row[col_data]
        if isinstance(dt_raw, datetime.datetime):
            dt = dt_raw.date()
        elif isinstance(dt_raw, datetime.date):
            dt = dt_raw
        else:
            continue  # célula vazia/texto -- fora do range de dias reais do mês
        entry = {}
        for i, key in col_map.items():
            entry[key] = _to_num(row[i]) if i < len(row) else None
        result[dt.strftime("%Y-%m-%d")] = entry
    return result


def _download_workbook(drive_service, spreadsheet_id):
    """Baixa a planilha (Drive API) e abre com openpyxl -- mesmo padrão de
    top5_origem_extension.py (self-contained, não depende do cache/import de
    extract_data.py, pra evitar import circular)."""
    import io
    import openpyxl
    from googleapiclient.http import MediaIoBaseDownload

    meta = drive_service.files().get(
        fileId=spreadsheet_id, fields="mimeType", supportsAllDrives=True
    ).execute()
    is_native_sheet = meta["mimeType"] == "application/vnd.google-apps.spreadsheet"

    buf = io.BytesIO()
    if is_native_sheet:
        request = drive_service.files().export_media(
            fileId=spreadsheet_id,
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        request = drive_service.files().get_media(fileId=spreadsheet_id, supportsAllDrives=True)
    downloader = MediaIoBaseDownload(buf, request, chunksize=10 * 1024 * 1024)
    done = False
    while not done:
        _, done = downloader.next_chunk(num_retries=5)
    buf.seek(0)
    return openpyxl.load_workbook(buf, data_only=True, read_only=True)


def _get_rows(wb, sheet_name):
    ws = wb[sheet_name]
    return [list(r) for r in ws.iter_rows(values_only=True)]


def build_cross_aquafoz(drive_service, spreadsheet_id, mes_atual_nome):
    """Função principal -- chamar de dentro de main() do extract_data.py e salvar o
    resultado em data['CROSS_AQUAFOZ'].

    mes_atual_nome: nome do mês vigente em PT, Title Case (ex. 'Agosto') -- mesmo valor que
    já está em cfg["meses_com_dados"][-1] (MESES_PT), é o nome exato da aba a ler pra visão
    Diária. Se essa aba ainda não existir na planilha (mês ainda não criado manualmente lá,
    ou nome digitado diferente), a visão Diária sai vazia mas a Mensal continua normal.

    Devolve {'mensal': {...}, 'diaria': {...}, 'mesAtual': mes_atual_nome}.
    """
    wb = _download_workbook(drive_service, spreadsheet_id)
    mensal = build_cross_aquafoz_mensal(_get_rows(wb, "Venda mes a mes"))
    diaria = {}
    if mes_atual_nome in wb.sheetnames:
        diaria = build_cross_aquafoz_diario(_get_rows(wb, mes_atual_nome))
    return {"mensal": mensal, "diaria": diaria, "mesAtual": mes_atual_nome}


# =========================================================================================
# COMO INTEGRAR no extract_data.py real
# =========================================================================================
# 1. Colocar este arquivo em scripts/cross_aquafoz_extension.py (mesma pasta do
#    extract_data.py de verdade) e adicionar, junto dos outros imports no topo do arquivo:
#      from cross_aquafoz_extension import build_cross_aquafoz
#
# 2. Adicionar ao config.json a chave (mesmo nível de visitacao_parques_id etc.):
#      "cross_aquafoz_id": "10y-MjciXhxYU7yBoUgymLmphuELd3SbyR6HbRejCGOo"
#
# 3. Em main(), em qualquer ponto depois que cfg["meses_com_dados"] já estiver calculado:
#      print("Lendo Cross AquaFoz...", file=sys.stderr)
#      try:
#          cross_aquafoz_id = cfg.get("cross_aquafoz_id", "")
#          if cross_aquafoz_id:
#              mes_atual = cfg["meses_com_dados"][-1] if cfg["meses_com_dados"] else None
#              mes_atual_nome = MESES_PT[MONTH_NUMBER[mes_atual] - 1] if mes_atual else None
#              cross_aquafoz = build_cross_aquafoz(service, cross_aquafoz_id, mes_atual_nome) \
#                  if mes_atual_nome else {"mensal": {}, "diaria": {}, "mesAtual": None}
#          else:
#              print("AVISO: 'cross_aquafoz_id' ainda não configurado no config.json -- aba Cross AquaFoz fica vazia.", file=sys.stderr)
#              cross_aquafoz = {"mensal": {}, "diaria": {}, "mesAtual": None}
#      except Exception as e:
#          print(f"AVISO: falha ao ler Cross AquaFoz ({e}) -- aba fica vazia neste ciclo.", file=sys.stderr)
#          cross_aquafoz = {"mensal": {}, "diaria": {}, "mesAtual": None}
#
#    e adicionar ao dict `output` final: "CROSS_AQUAFOZ": cross_aquafoz,
#
# 4. A planilha "Vendas Cross para AquaFoz 2026" precisa estar compartilhada com a service
#    account do pipeline (o mesmo e-mail que já tem acesso às outras planilhas do
#    config.json) -- senão a leitura cai no except acima e a aba do painel fica vazia.
# =========================================================================================
