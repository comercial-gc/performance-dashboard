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

Requer: google-api-python-client, google-auth, openpyxl (ver requirements.txt)

NOTA IMPORTANTE: as planilhas do Grupo Cataratas estao guardadas no Drive como arquivos
.xlsx (Excel de verdade), nao como Google Sheets nativos. A API do Google Sheets nao le
arquivos .xlsx diretamente ("This operation is not supported for this document. The
document must not be an Office file."). Por isso este script usa a API do **Google Drive**
para baixar o arquivo (funciona tanto pra .xlsx quanto pra Google Sheets nativo) e le o
conteudo com openpyxl -- a mesma biblioteca usada pra validar os dados nesta conversa.
"""
import argparse
import calendar
import datetime
import io
import json
import socket
import sys

import openpyxl
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# O runner do GitHub Actions as vezes tem uma conexao mais lenta/instavel com a API do
# Google, e o timeout padrao do socket (ilimitado, mas o cliente HTTP interno usa um valor
# curto) estourava no meio do download de planilhas maiores. Aumenta a margem e deixa o
# retry automático (num_retries em next_chunk, abaixo) lidar com falhas passageiras.
socket.setdefaulttimeout(300)

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

MONTH_NUMBER = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARÇO": 3, "ABRIL": 4, "MAIO": 5, "JUNHO": 6,
    "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9, "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12,
}

PARKS = ["AquaRio", "BioParque", "Paineiras", "PNI", "M3F", "AquaFoz", "Três Pescadores", "Vila Velha"]
ATRATIVOS = ["GEX", "MDE", "MDC"]

SHEETS_EPOCH = datetime.date(1899, 12, 30)


def serial_to_date(serial):
    """Converte o numero serial de data (mesmo formato usado pela API do Sheets) para
    datetime.date. Mantido por compatibilidade -- get_values() ja devolve as datas nesse
    formato numerico, igual antes."""
    if serial is None:
        return None
    try:
        return SHEETS_EPOCH + datetime.timedelta(days=float(serial))
    except (TypeError, ValueError):
        return None


def get_drive_service():
    creds = service_account.Credentials.from_service_account_file(
        _sa_key_path(), scopes=SCOPES
    ) if _sa_key_path() else service_account.Credentials.from_service_account_info(
        json.loads(_sa_key_env()), scopes=SCOPES
    )
    return build("drive", "v3", credentials=creds)


def _sa_key_path():
    import os
    return os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")


def _sa_key_env():
    import os
    # alternativa: colar o JSON inteiro direto numa env var (usado no workflow do GitHub)
    return os.environ.get("GCP_SERVICE_ACCOUNT_JSON", "{}")


_WORKBOOK_CACHE = {}


def _download_workbook(drive_service, spreadsheet_id):
    """Baixa o arquivo (Drive API) e abre com openpyxl. Cacheia por spreadsheet_id pra nao
    baixar o mesmo arquivo de novo a cada aba lida (ex.: 7 abas de mes na mesma planilha)."""
    if spreadsheet_id in _WORKBOOK_CACHE:
        return _WORKBOOK_CACHE[spreadsheet_id]

    # supportsAllDrives=True e' necessario quando o arquivo vive dentro de um Drive
    # Compartilhado (Shared Drive) da empresa -- sem isso a API responde "File not found"
    # mesmo quando o arquivo esta compartilhado corretamente com a service account.
    meta = drive_service.files().get(
        fileId=spreadsheet_id, fields="mimeType, name", supportsAllDrives=True
    ).execute()
    is_native_sheet = meta["mimeType"] == "application/vnd.google-apps.spreadsheet"

    buf = io.BytesIO()
    if is_native_sheet:
        # Google Sheets nativo: exporta como .xlsx antes de baixar
        request = drive_service.files().export_media(
            fileId=spreadsheet_id,
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        # ja e' um arquivo .xlsx (ou similar): baixa direto
        request = drive_service.files().get_media(fileId=spreadsheet_id, supportsAllDrives=True)

    # num_retries>0 faz o proprio googleapiclient re-tentar automaticamente (com backoff)
    # em caso de erro de rede passageiro (timeout, conexao resetada etc.) -- foi isso que
    # derrubou uma execucao no GitHub Actions ("TimeoutError: The read operation timed out").
    downloader = MediaIoBaseDownload(buf, request, chunksize=10 * 1024 * 1024)
    done = False
    while not done:
        _, done = downloader.next_chunk(num_retries=5)
    buf.seek(0)

    wb = openpyxl.load_workbook(buf, data_only=True, read_only=True)
    _WORKBOOK_CACHE[spreadsheet_id] = wb
    return wb


def get_values(service, spreadsheet_id, sheet_name, a1_range=None):
    """Busca uma aba inteira como lista de listas. Datas viram numero serial (mesmo
    formato que a API do Sheets usava), pra nao precisar mudar o resto do codigo que
    espera esse formato."""
    wb = _download_workbook(service, spreadsheet_id)
    ws = wb[sheet_name]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    width = max((len(r) for r in rows), default=0)
    out = []
    for r in rows:
        r = r + [None] * (width - len(r))
        r2 = []
        for v in r:
            if isinstance(v, datetime.datetime):
                r2.append(float((v.date() - SHEETS_EPOCH).days))
            else:
                r2.append(v)
        out.append(r2)
    return out


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
# Visitação Diária Ponderada - 2026 - Final.xlsx -> aba "Visitação diária GC V2"
# Tabela dinamica "Soma de Nova2026" (projecao diaria ponderada por parque, baseada no
# padrao de dias da semana de anos anteriores) -- usada como "Meta" nos meses que ainda
# nao tem resultado real (Agosto em diante). Por pedido do usuario, a Meta e' colocada no
# MESMO campo "OBZ 2026" que a regua diaria ja mostra (nao criamos um campo novo) -- assim
# a celula do dia mantem exatamente o layout de sempre (numero grande + "vs 25" + "OBZ"),
# só que com o numero grande e "vs 25" vazios ("-") e o "OBZ" preenchido com a meta.
# ---------------------------------------------------------------------------

VISITACAO_META_PARK_NAMES = {
    "Marco": "M3F", "3P": "Três Pescadores", "VV": "Vila Velha",
    "AquaFoz": "AquaFoz", "AquaRio": "AquaRio", "BioParque": "BioParque",
    "Paineiras": "Paineiras", "PNI": "PNI",
}

# BUG EVITADO: os "Rótulos de Coluna" dessa tabela dinamica NAO sao celulas de data de
# verdade (isinstance datetime) -- sao TEXTO no formato "01/mar", "02/mar" etc. (o Excel
# converte assim quando agrupa datas numa tabela dinamica). Por isso o parser abaixo le
# esse texto em vez de checar isinstance(..., datetime.datetime).
MES_ABREV_PT = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
}


def _parse_dd_mon(label, ano=2026):
    if not isinstance(label, str) or "/" not in label:
        return None
    dia_str, _, mes_str = label.strip().partition("/")
    mes_num = MES_ABREV_PT.get(mes_str.strip().lower())
    if mes_num is None or not dia_str.strip().isdigit():
        return None
    try:
        return datetime.date(ano, mes_num, int(dia_str))
    except ValueError:
        return None


# BUG EVITADO / TEMPORARIO: por pedido do usuario, os dados da planilha "Visitação
# Diária Ponderada - 2026" ainda NAO sao lidos ao vivo do Drive (a planilha ainda nao foi
# compartilhada com a service account) -- por enquanto, ficam "amarrados" aqui como dado
# ESTATICO, extraido uma unica vez da copia que o usuario enviou (Agosto a Dezembro/2026,
# projecao "Nova2026" por parque/dia). Quando a planilha for conectada ao Drive, trocar o
# uso de VISITACAO_META_ESTATICO em main() pela chamada a build_visitacao_meta() (a funcao
# ja esta pronta logo abaixo, so' nao esta sendo chamada por enquanto).
VISITACAO_META_ESTATICO = json.loads('''{"AGOSTO": {"Três Pescadores": [1019, 1019, 357, 0, 357, 357, 357, 1019, 1019, 357, 0, 357, 357, 357, 1019, 1019, 357, 0, 357, 357, 357, 1019, 1019, 357, 0, 357, 357, 357, 1019, 1019, 357], "M3F": [2462, 2462, 0, 862, 862, 862, 862, 2462, 2462, 0, 862, 862, 862, 862, 2462, 2462, 0, 862, 862, 862, 862, 2462, 2462, 0, 862, 862, 862, 862, 2462, 2462, 0], "AquaFoz": [2315, 2315, 810, 810, 810, 810, 810, 2315, 2315, 810, 810, 810, 810, 810, 2315, 2315, 810, 810, 810, 810, 810, 2315, 2315, 810, 810, 810, 810, 810, 2315, 2315, 810], "BioParque": [3816, 3816, 0, 0, 1335, 1335, 1335, 3816, 3816, 0, 0, 1335, 1335, 1335, 3816, 3816, 0, 0, 1335, 1335, 1335, 3816, 3816, 0, 0, 1335, 1335, 1335, 3816, 3816, 0], "AquaRio": [4678, 3236, 1951, 2355, 1970, 1884, 2474, 4678, 3236, 1951, 2355, 1970, 1884, 2474, 4678, 3236, 1951, 2355, 1970, 1884, 2474, 4678, 3236, 1951, 2355, 1970, 1884, 2474, 4678, 3236, 1951], "Vila Velha": [336, 336, 118, 0, 118, 118, 118, 336, 336, 118, 0, 118, 118, 118, 336, 336, 118, 0, 118, 118, 118, 336, 336, 118, 0, 118, 118, 118, 336, 336, 118], "PNI": [6816, 6544, 3308, 4438, 4544, 4625, 5721, 6816, 6544, 3308, 4438, 4544, 4625, 5721, 6816, 6544, 3308, 4438, 4544, 4625, 5721, 6816, 6544, 3308, 4438, 4544, 4625, 5721, 6816, 6544, 3308], "Paineiras": [3751, 4112, 3561, 2665, 3717, 3660, 3828, 3751, 4112, 3561, 2665, 3717, 3660, 3828, 3751, 4112, 3561, 2665, 3717, 3660, 3828, 3751, 4112, 3561, 2665, 3717, 3660, 3828, 3751, 4112, 3561]}, "SETEMBRO": {"Três Pescadores": [0, 376, 376, 376, 1074, 1074, 1074, 0, 376, 376, 376, 1074, 1074, 376, 0, 376, 376, 376, 1074, 1074, 376, 0, 376, 376, 376, 1074, 1074, 376, 0, 376], "M3F": [1102, 1102, 1102, 1102, 2561, 2561, 0, 1268, 1102, 1102, 1102, 2561, 2561, 0, 1102, 1102, 1102, 1102, 2561, 2561, 0, 1102, 1102, 1102, 1102, 2561, 2561, 0, 1102, 1102], "AquaFoz": [1021, 1021, 1021, 1021, 2219, 2219, 2423, 1021, 1021, 1021, 1021, 2219, 2219, 1021, 1021, 1021, 1021, 1021, 2219, 2219, 1021, 1021, 1021, 1021, 1021, 2219, 2219, 1021, 1021, 1021], "BioParque": [0, 1249, 1249, 1249, 3568, 3568, 0, 0, 1249, 1249, 1249, 3568, 3568, 0, 0, 1249, 1249, 1249, 3568, 3568, 0, 0, 1249, 1249, 1249, 3568, 3568, 0, 0, 1249], "AquaRio": [2101, 1993, 2422, 2473, 4021, 3574, 4127, 2101, 1993, 2422, 2473, 4021, 3574, 2127, 2101, 1993, 2422, 2473, 4021, 3574, 2127, 2101, 1993, 2422, 2473, 3021, 2574, 2127, 2101, 1993], "Vila Velha": [0, 147, 147, 147, 420, 420, 420, 0, 147, 147, 147, 420, 420, 147, 0, 147, 147, 147, 420, 420, 147, 0, 147, 147, 147, 420, 420, 147, 0, 147], "PNI": [4509, 5233, 5135, 5159, 7572, 7692, 6880, 4509, 5233, 5135, 5159, 7572, 7692, 3880, 4509, 5233, 5135, 5159, 7572, 7692, 3880, 4509, 4233, 4135, 5159, 7572, 7692, 3880, 4509, 4233], "Paineiras": [2892, 3301, 3120, 4542, 5000, 3360, 3998, 2892, 3301, 3120, 4542, 5000, 3360, 3998, 2892, 3301, 3120, 4542, 5000, 3360, 3998, 2892, 3301, 3120, 4542, 5000, 3360, 3998, 2892, 3301]}, "OUTUBRO": {"Três Pescadores": [322, 322, 920, 920, 322, 0, 322, 322, 322, 920, 920, 920, 0, 322, 322, 322, 920, 920, 322, 0, 322, 322, 322, 920, 920, 322, 0, 322, 322, 322, 920], "M3F": [1104, 1169, 2199, 2199, 0, 1104, 1104, 1104, 1104, 2199, 2199, 2199, 1104, 1104, 1104, 1104, 2199, 2199, 0, 1104, 1104, 1104, 1104, 2199, 2199, 0, 1104, 1104, 1104, 1104, 2199], "AquaFoz": [1115, 1115, 2493, 2493, 1115, 1115, 1115, 1115, 1115, 2493, 2493, 2493, 1115, 1115, 1115, 1115, 2493, 2493, 1115, 1115, 1115, 1115, 1115, 2493, 2493, 1115, 1115, 1115, 1115, 1115, 2484], "BioParque": [1501, 1501, 4729, 4729, 0, 0, 1501, 1501, 1501, 4729, 4729, 8000, 0, 1501, 1501, 1501, 4729, 4729, 0, 0, 1501, 1501, 1501, 4729, 4729, 0, 0, 1501, 1501, 1501, 4729], "AquaRio": [2431, 2742, 4725, 5569, 2503, 2600, 3349, 2431, 2742, 4725, 5569, 4503, 2600, 3349, 2431, 2742, 4725, 5569, 2503, 2600, 3349, 2431, 2742, 3725, 4569, 2503, 2600, 3349, 2431, 2742, 4725], "Vila Velha": [124, 124, 355, 355, 124, 0, 124, 124, 124, 355, 355, 355, 0, 124, 124, 124, 355, 355, 124, 0, 124, 124, 124, 355, 355, 124, 0, 124, 124, 124, 355], "PNI": [5108, 5838, 8611, 7465, 5486, 6140, 4799, 5108, 5838, 8611, 7465, 5486, 6140, 4799, 5108, 5838, 8611, 7465, 5486, 6140, 4799, 5108, 5838, 8611, 7465, 5486, 6140, 4799, 5108, 5838, 8611], "Paineiras": [3121, 3623, 4304, 4389, 4577, 3131, 2398, 3121, 3623, 4304, 4389, 4577, 3131, 2398, 3121, 3623, 4304, 4389, 4577, 3131, 2398, 3121, 3623, 4304, 4389, 4577, 3131, 2398, 3121, 3623, 4304]}, "NOVEMBRO": {"Três Pescadores": [1069, 1069, 0, 374, 374, 374, 1069, 1069, 374, 0, 374, 374, 374, 1069, 1069, 374, 0, 374, 1230, 1069, 1069, 1069, 374, 0, 374, 374, 374, 1069, 1069, 374], "M3F": [2444, 2443, 1254, 1254, 1254, 1254, 2443, 2443, 0, 1254, 1254, 1254, 1254, 2443, 2443, 0, 1254, 1254, 1254, 1254, 2443, 2443, 0, 1254, 1254, 1254, 1254, 2443, 2443, 0], "AquaFoz": [2665, 2665, 1101, 1101, 1101, 1101, 2665, 2665, 1101, 1101, 1101, 1101, 1101, 2665, 2665, 1101, 1101, 1101, 2271, 2665, 2665, 2665, 1101, 1101, 1101, 1101, 1101, 2665, 2665, 999], "BioParque": [2627, 2627, 0, 920, 920, 920, 2627, 2627, 0, 0, 920, 920, 920, 2627, 2627, 0, 0, 920, 3021, 2627, 2627, 2627, 0, 0, 920, 920, 920, 2627, 2627, 0], "AquaRio": [3402, 3548, 1998, 1755, 2046, 2732, 3126, 3002, 2648, 1998, 1755, 2046, 2732, 3126, 3002, 2648, 1998, 1755, 2046, 4232, 3126, 3002, 2648, 1998, 1755, 2046, 2732, 3126, 3002, 2648], "Vila Velha": [456, 456, 0, 160, 160, 160, 456, 456, 160, 0, 160, 160, 160, 456, 456, 160, 0, 160, 456, 524, 456, 456, 160, 0, 160, 160, 160, 456, 456, 160], "PNI": [9084, 5112, 5934, 5232, 6384, 7437, 8572, 9084, 5112, 5934, 5232, 6384, 7437, 8572, 9084, 5112, 5934, 5232, 6384, 7437, 8572, 9084, 5112, 5934, 5232, 6384, 7437, 8572, 9084, 5112], "Paineiras": [3391, 3695, 3327, 4177, 4074, 5648, 5323, 3391, 3695, 3327, 4177, 4074, 5648, 5323, 3391, 3695, 3327, 4177, 4074, 5648, 5323, 3391, 3695, 3327, 4177, 4074, 5648, 5323, 3391, 3695]}, "DEZEMBRO": {"Três Pescadores": [0, 393, 393, 393, 852, 852, 393, 0, 393, 393, 393, 852, 852, 393, 0, 393, 393, 393, 1124, 1124, 420, 420, 420, 1293, 1124, 1124, 1124, 470, 470, 435, 420], "M3F": [1121, 1121, 1121, 1121, 1950, 1950, 0, 1121, 1121, 1121, 1121, 1950, 1950, 0, 1121, 1121, 1121, 1121, 2010, 2010, 2010, 2010, 2010, 0, 2745, 2745, 2745, 2010, 2069, 2500, 0], "AquaFoz": [1042, 1042, 1042, 1042, 2521, 2521, 1042, 1042, 1042, 1042, 1042, 2521, 2521, 1042, 1042, 1042, 1042, 1042, 2977, 2977, 1042, 1042, 1042, 1042, 2000, 2977, 2977, 2400, 2400, 2400, 2151], "BioParque": [0, 1101, 1101, 1101, 2625, 2625, 0, 0, 1101, 1101, 1101, 2625, 2510, 0, 0, 1101, 1101, 1101, 2525, 2525, 1300, 1300, 1300, 0, 1900, 2625, 2625, 1650, 1650, 1650, 0], "AquaRio": [1852, 1862, 1905, 1943, 2621, 2281, 1852, 1862, 1905, 1943, 1943, 3140, 2281, 1943, 1852, 1862, 1905, 1943, 3251, 3096, 1852, 1862, 1905, 1351, 2251, 3685, 3950, 3482, 3485, 3489, 2051], "Vila Velha": [0, 148, 148, 148, 423, 423, 148, 0, 148, 148, 148, 423, 423, 148, 0, 148, 148, 148, 423, 423, 148, 0, 148, 148, 246, 423, 423, 240, 240, 240, 240], "PNI": [5661, 6053, 5842, 5776, 7084, 7008, 5601, 5661, 6053, 5842, 5776, 7984, 7908, 5601, 5661, 6053, 5842, 5776, 8984, 8908, 5661, 6053, 5842, 7893, 8280, 8280, 8084, 7680, 7540, 6852, 7152], "Paineiras": [3004, 3483, 3471, 4076, 3625, 3690, 3203, 3004, 3983, 4471, 4476, 3625, 3690, 3203, 3504, 3983, 4471, 4476, 3625, 3690, 3803, 3404, 3983, 2671, 4476, 3425, 3490, 4203, 4504, 4983, 4471]}}''')


def build_visitacao_meta(service, spreadsheet_id, sheet_name, meses_meta):
    """Le a tabela dinamica 'Soma de Nova2026' (uma coluna por dia do ano, uma linha por
    parque) e devolve {mes: {parque: [valor_dia1, valor_dia2, ...]}} só para os meses
    informados (tipicamente os que ainda nao tem resultado real)."""
    if not meses_meta:
        return {}
    rows = get_values(service, spreadsheet_id, sheet_name)

    # Acha a tabela dinamica de verdade: a aba tem MAIS DE UMA celula solta chamada
    # "Soma de Nova2026" (ex.: uma mini-tabela de referencia perto do topo, com "Rótulos
    # de Linha" tambem na linha seguinte, mas SEM colunas de data -- so' nomes de parque).
    # Por isso nao paramos no primeiro match: só aceitamos o candidato cuja linha de
    # cabecalho realmente tem rótulos de data (formato "01/mar") nas colunas seguintes.
    col_to_date = {}
    header_row_idx, label_col = None, None
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            if val != "Soma de Nova2026":
                continue
            next_row = rows[i + 1] if i + 1 < len(rows) else []
            if cell(next_row, j) != "Rótulos de Linha":
                continue
            candidato = {}
            for c in range(j + 1, len(next_row)):
                dt = _parse_dd_mon(cell(next_row, c))
                if dt is not None:
                    candidato[c] = dt
            if candidato:
                header_row_idx, label_col, col_to_date = i + 1, j, candidato
                break
        if header_row_idx is not None:
            break
    if header_row_idx is None:
        return {}

    result = {mes: {} for mes in meses_meta}
    for mes in meses_meta:
        n_days = calendar.monthrange(2026, MONTH_NUMBER[mes])[1]
        for park in set(VISITACAO_META_PARK_NAMES.values()):
            result[mes][park] = [None] * n_days

    for r in range(header_row_idx + 1, len(rows)):
        label = cell(rows[r], label_col)
        if not isinstance(label, str):
            continue
        if label.strip().lower().startswith("total"):
            break
        park = VISITACAO_META_PARK_NAMES.get(label)
        if park is None:
            continue
        for c, dt in col_to_date.items():
            mes_en = next((m for m in meses_meta if MONTH_NUMBER[m] == dt.month), None)
            if mes_en is None:
                continue
            v = cell(rows[r], c)
            if isinstance(v, (int, float)):
                result[mes_en][park][dt.day - 1] = v
    return result


# ---------------------------------------------------------------------------
# Visitação Parques 2026.xlsx -> aba "Mapa Clima"
# Serie diaria (sequencial, ano inteiro) com o codigo de clima do dia, por regiao. O
# painel so' usa 1 "parque representante" por regiao (CLIMA_REGIOES no HTML): AquaRio
# representa Parques Rio, PNI representa Parques Foz, Vila Velha e Três Pescadores tem
# regiao propria. BUG EVITADO: essa aba nunca foi ligada ao pipeline -- por isso a aba
# Clima do painel sumia (ficava sem nenhum dado) assim que o painel passava a ler o
# data.json em vez dos numeros fixos.
# ---------------------------------------------------------------------------

# (coluna da Data, coluna do texto "Clima") por parque-representante, 0-indexed
CLIMA_REGIAO_COLS = {
    "AquaRio": (6, 7),
    "PNI": (15, 16),
    "Vila Velha": (24, 25),
    "Três Pescadores": (30, 31),
}

CLIMA_LABEL_TO_CODE = {
    "dia de sol": "sun", "nublado": "cloud", "nublado com chuva": "cloud-rain",
    "chuvoso": "rain", "sol + chuva": "sun-rain", "tempestade": "storm",
    "calor excessivo": "heat", "frio intenso": "cold", "sol + vento": "wind-sun",
    "chuva + vento": "wind-rain", "ventoso": "wind", "fechado": "closed", "alerta": "alert",
}


def build_clima_emoji(service, spreadsheet_id, sheet_name, meses_com_dados):
    rows = get_values(service, spreadsheet_id, sheet_name)
    result = {mes: {} for mes in meses_com_dados}
    for mes in meses_com_dados:
        month_number = MONTH_NUMBER[mes]
        n_days = calendar.monthrange(2026, month_number)[1]
        for park in CLIMA_REGIAO_COLS:
            result[mes][park] = [None] * n_days

    for row in rows:
        for park, (date_col, clima_col) in CLIMA_REGIAO_COLS.items():
            date_val = cell(row, date_col)
            dt = None
            if isinstance(date_val, datetime.datetime):
                dt = date_val.date()
            elif isinstance(date_val, (int, float)):
                dt = serial_to_date(date_val)
            if dt is None or dt.year != 2026:
                continue
            mes_en = next((m for m in meses_com_dados if MONTH_NUMBER[m] == dt.month), None)
            if mes_en is None:
                continue
            label = cell(row, clima_col)
            if not isinstance(label, str) or not label.strip():
                continue
            code = CLIMA_LABEL_TO_CODE.get(label.strip().lower())
            if code:
                result[mes_en][park][dt.day - 1] = code
    return result


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

# linha (0-indexed) onde comeca o bloco de cada parque nesta aba. "3P" (Três Pescadores)
# e "Vila Velha" ficam num grupo separado, abaixo do titulo "SOUL PARQUES" -- so' tem
# coluna 2026 preenchida (parques novos, sem historico 2025), mas o layout de linhas
# (Visitação/Ecommerce/Share/R$ em mídia) e' o mesmo dos outros blocos. As chaves "3P" e
# "Vila Velha" (em vez de "Três Pescadores") sao de proposito -- e' o nome que o HTML
# (INVEST_PARKS) ja espera pra esses dois parques nesta tabela especifica.
SHARE_ECOMMERCE_BLOCKS = {
    "BioParque": 0, "AquaRio": 9, "Paineiras": 17, "M3F": 25, "AquaFoz": 33,
    "3P": 44, "Vila Velha": 52,
}
MESES_PT = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho",
            "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]


def _find_labeled_row_idx(rows, r0, label, max_offset=10):
    """Acha o indice (em `rows`) da linha rotulada `label` na coluna A, procurando a
    partir de r0+1 ate r0+max_offset. Usado pra achar "Ecommerce (base TI)" cujo
    deslocamento varia por bloco (BioParque tem uma linha extra "Historico Pareto" que
    empurra o resto do bloco pra baixo, os outros parques nao tem)."""
    label_norm = label.strip().lower()
    limit = min(r0 + max_offset, len(rows) - 1)
    for idx in range(r0 + 1, limit + 1):
        val = cell(rows[idx], 0)
        if isinstance(val, str) and val.strip().lower() == label_norm:
            return idx
    return None


def build_investimento_midia(service, spreadsheet_id, sheet_name, meses_com_dados):
    """Monta SHARE.investimentoMidia.meses para os parques rastreados nesta aba
    (AquaRio, BioParque, Paineiras, M3F, AquaFoz, e os dois do bloco "SOUL PARQUES":
    "3P"/Três Pescadores e Vila Velha).

    Para meses fechados (todo mes exceto o corrente), tanto o lado 2025 quanto o 2026 sao
    o total do mes inteiro. Para o mes corrente (ainda em andamento), o lado 2026 e'
    parcial (só os dias já lançados na planilha) e o 2025 e' o mes inteiro do ano passado
    -- usado como referencia de "para onde estamos indo", nao como comparação dia-a-dia.

    "3P" e "Vila Velha" so' tem coluna 2026 (parques novos, sem historico 2025) -- os
    campos *2025 ficam None/0 automaticamente pra eles, sem precisar de tratamento
    especial (num() ja devolve None pra celula vazia).
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
            vis_row, inv_row = rows[r0 + 1], rows[r0 + 4]
            # BUG EVITADO: por pedido do usuario, o painel passou a usar "Ecommerce (base
            # TI)" em vez da linha "Ecommerce" simples -- mas o deslocamento dessa linha
            # varia por bloco (BioParque tem uma linha extra "Historico Investimento em
            # midia (Pareto)" que empurra o resto do bloco pra baixo), entao procuramos
            # pelo rotulo em vez de usar um indice fixo.
            ecom_ti_idx = _find_labeled_row_idx(rows, r0, "Ecommerce (base TI)")
            ecom_row = rows[ecom_ti_idx] if ecom_ti_idx is not None else rows[r0 + 2]
            share_row = rows[ecom_ti_idx + 1] if ecom_ti_idx is not None else rows[r0 + 3]

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
                "share2025": (ecom25 / vis25) if (vis25 and ecom25 is not None) else num(share_row, idx_2025),
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
        vis_row, inv_row = rows[r0 + 1], rows[r0 + 4]
        # mesma logica de build_investimento_midia: usa a Share de "Ecommerce (base TI)",
        # cujo deslocamento varia por bloco -- procura pelo rotulo em vez de indice fixo.
        ecom_ti_idx = _find_labeled_row_idx(rows, r0, "Ecommerce (base TI)")
        share_row = rows[ecom_ti_idx + 1] if ecom_ti_idx is not None else rows[r0 + 3]

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
# [2026] Mix OBZ e visitação.xlsx -> aba "ANÁLISE MIX DE ORIGEM 2026"
# Bloco "MIX PARCIAL | MES ATUAL": um mini-bloco por parque, com o nome do parque na
# coluna C e um cabecalho (2025/2026/Delta/Share) na mesma linha, seguido das linhas de
# categoria (Local/Brasileiros/Estrangeiros/... e "Total Mes" por ultimo). O nome do
# parque nesta aba as vezes difere do nome usado no resto do painel (ex.: "Aquario" sem
# acento/maiuscula) -- por isso o mapeamento MIX_ORIGEM_PARK_NAMES abaixo.
# BUG EVITADO (mesma classe dos outros): antes desta funcao, essa tabela inteira do
# painel (MIX_ORIGEM) nunca foi ligada ao pipeline -- ficava com numero fixo, travado
# na data em que foi colado a mao pela ultima vez, mesmo com o resto ja automatizado.
# ---------------------------------------------------------------------------

MIX_ORIGEM_PARK_NAMES = {
    "Aquario": "AquaRio",
    "BioParque": "BioParque",
    "Paineiras": "Paineiras",
    "M3F": "M3F",
    "AquaFoz": "AquaFoz",
    "Três Pescadores": "Três Pescadores",
    "Vila Velha": "Vila Velha",
    "PNI": "PNI",
}


def build_mix_origem(service, spreadsheet_id, sheet_name, col_offset=0):
    """col_offset=0 le' o bloco "MIX PARCIAL" (mes atual); col_offset=11 le' o bloco
    "MIX ACUMULADO ANO" que fica 11 colunas a direita, mesmas linhas, mesmo layout --
    usado pelo toggle Mensal/Acumulado do painel."""
    name_col = 2 + col_offset
    rows = get_values(service, spreadsheet_id, sheet_name)
    result = {}
    for i, row in enumerate(rows):
        nome_sheet = cell(row, name_col)
        if nome_sheet not in MIX_ORIGEM_PARK_NAMES:
            continue
        header_d = cell(row, name_col + 1)
        # confirma que e' mesmo a linha de cabecalho do bloco (coluna seguinte = 2025 ou
        # 2026), nao outra celula qualquer da planilha que por acaso tenha o mesmo nome
        if header_d not in (2025, 2026):
            continue
        formato = "full" if header_d == 2025 else "somente2026"
        park = MIX_ORIGEM_PARK_NAMES[nome_sheet]

        categorias = []
        j = i + 1
        while j < len(rows):
            label = cell(rows[j], name_col)
            if not label:
                break
            # BUG EVITADO: a planilha tem "Local " com espaco sobrando no final da celula;
            # sem o strip(), a cor da legenda (que faz lookup por nome exato) nao batia e
            # a barra/legenda de "Local" caia no cinza padrao.
            label = label.strip() if isinstance(label, str) else label
            if formato == "full":
                v2025, v2026 = cell(rows[j], name_col + 1), cell(rows[j], name_col + 2)
                delta = cell(rows[j], name_col + 3)
                share25, share26 = cell(rows[j], name_col + 4), cell(rows[j], name_col + 5)
            else:
                v2025, delta, share25 = None, None, None
                v2026, share26 = cell(rows[j], name_col + 1), cell(rows[j], name_col + 2)
            categorias.append({
                "label": label, "v2025": v2025, "v2026": v2026, "delta": delta,
                "share25": share25, "share26": share26,
            })
            is_total = isinstance(label, str) and label.strip().startswith("Total")
            j += 1
            if is_total:
                break
        result[park] = {"formato": formato, "categorias": categorias}
    return result


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
                # BUG EVITADO: nome de parque (ex.: "M3F", "PAINEIRAS") pode aparecer em mais
                # de um lugar na aba (ex.: legenda, celula de referencia). So' tratamos como
                # o INICIO de um bloco de verdade se a linha logo abaixo tiver o cabecalho
                # "MES" na mesma coluna -- senao ficamos vulneraveis a sobrescrever um bloco
                # bom com lixo de uma ocorrencia errada do nome mais abaixo na planilha.
                header_below = cell(rows[i + 1], c) if i + 1 < len(rows) else None
                if not (isinstance(header_below, str) and header_below.strip().upper() in ("MÊS", "MES")):
                    continue
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
# BUG EVITADO: "SETOR" (area: Performance/Marca & Experiencia/Conteudo/InHouse) e "CUSTO"
# (tipo de custo: Custo Fixo/Midia OFF/Midia ON/Eventos) sao DUAS colunas DIFERENTES que
# coexistem em Junho/Julho -- nao aliases uma da outra. Antes eram tratadas como a mesma
# coisa e a coluna CUSTO era sempre descartada (SETOR vinha primeiro na linha).
DETAIL_HEADER_ALIASES = {
    "parque": ["PARQUE"],
    "setor": ["SETOR"],
    "custo": ["CUSTO"],
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


def _clean_str(v):
    if v is None:
        return None
    s = str(v).strip()
    return s if s else None


def build_invest_mkt_detail(service, spreadsheet_id, meses):
    """Le a lista de campanhas de cada aba mensal. Retorna {mes: [ {parque, setor,
    custo, fornecedor, descricao, valor, observacao}, ... ]}.

    Lê TODOS os meses passados (não só os "meses_com_dados") -- a aba mensal de
    Investimento Marketing já recebe lançamentos de custo fixo com antecedência (ex.:
    Agosto já tinha linhas preenchidas em Julho), então mesmo mês "futuro" pode ter
    detalhe real pra mostrar, mesmo que o resumo geral daquele mês ainda esteja incompleto.
    Cada mês é lido isoladamente (try/except) -- se uma aba não existir ou vier vazia,
    esse mês fica com lista vazia, sem derrubar os outros meses nem o resto do pipeline.
    """
    detail = {}
    for mes in meses:
        try:
            rows = get_values(service, spreadsheet_id, mes)
        except Exception as e:
            print(f"AVISO: falha ao ler detalhe de Investimento Marketing de {mes} ({e}) -- mês fica vazio.", file=sys.stderr)
            detail[mes] = []
            continue
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
                "setor": _clean_str(cell(row, cols.get("setor"))) if "setor" in cols else None,
                "custo": _clean_str(cell(row, cols.get("custo"))) if "custo" in cols else None,
                "fornecedor": cell(row, cols.get("fornecedor")) if "fornecedor" in cols else None,
                "descricao": cell(row, cols.get("descricao")) if "descricao" in cols else None,
                "valor": _parse_valor(cell(row, cols.get("valor"))),
                "observacao": cell(row, cols.get("observacao")) if "observacao" in cols else None,
            })
        detail[mes] = items
    return detail


def build_eventos(service, spreadsheet_id, sheet_name):
    """Aba 'Eventos' de Visitação Parques 2026: calendário simples (uma linha por evento) com
    3 colunas -- Data | Parque | Evento -- cobrindo Abril/2025 em diante (sem separação por
    aba/mês como as outras abas dessa planilha). Retorna lista ordenada por data, com ano/mês
    (PT) já calculados pra facilitar o filtro no front-end. O campo "parque" é mantido como
    veio da planilha (ex.: "Aquario", "Parques Rio", "Todos os Parques") -- os nomes lá não
    seguem exatamente a lista canônica de parques do painel, então não tentamos remapear.
    """
    rows = get_values(service, spreadsheet_id, sheet_name)
    eventos = []
    for row in rows[1:]:  # linha 0 = cabeçalho (Data / Parque / Evento)
        serial = cell(row, 0)
        evento = _clean_str(cell(row, 2))
        if serial is None or not evento:
            continue
        d = serial_to_date(serial)
        if d is None:
            continue
        eventos.append({
            "data": d.isoformat(),
            "ano": d.year,
            "mes": MESES_PT[d.month - 1],
            "parque": _clean_str(cell(row, 1)) or "",
            "evento": evento,
        })
    eventos.sort(key=lambda e: e["data"])
    return eventos


# ---------------------------------------------------------------------------
# Visitação Parques HISTORICO_Grupo Cataratas.xlsx (uma aba por parque: AquaRio,
# BioParque, Paineiras, PNI, M3F, AquaFoz). Cada aba tem até 3 blocos empilhados na coluna
# B: o nome do parque (Visitação, ano a ano), "Share " (share %, só em AquaRio/BioParque/
# Paineiras) e "Visão semestral" (2 colunas: Jan-Jun / Jul-Dez). "Consolidado" (soma dos 6
# parques) é calculado aqui, não existe pronto na planilha.
# ---------------------------------------------------------------------------

HISTORICO_PARKS = ["AquaRio", "BioParque", "Paineiras", "PNI", "M3F", "AquaFoz"]
HISTORICO_PARKS_COM_SHARE = ["AquaRio", "BioParque", "Paineiras"]


def _historico_find_row(rows, label):
    """Acha a linha cujo rótulo (coluna B) bate com `label` (ex.: nome do parque, "Share ",
    "Visão semestral") -- ignora espaços sobrando, que a planilha tem em alguns rótulos."""
    target = label.strip().lower()
    for i, row in enumerate(rows):
        v = cell(row, 1)
        if isinstance(v, str) and v.strip().lower() == target:
            return i
    return None


def _historico_parse_block(rows, start_row, ncols):
    """A partir da linha de rótulo do bloco (`start_row`), lê as linhas seguintes no formato
    Ano | valor x ncols (coluna B = ano, colunas seguintes = Jan..Dez ou Sem1/Sem2), até achar
    uma linha sem ano válido na coluna B. Tolera UMA linha de sub-cabeçalho logo no início
    (ex.: "Jan-Jun"/"Jul-Dez" antes dos anos, no bloco Visão Semestral de Paineiras)."""
    anos = {}
    if start_row is None:
        return anos
    r = start_row + 1
    permitiu_subcabecalho = False
    while r < len(rows):
        row = rows[r]
        ano_raw = cell(row, 1)
        if ano_raw is None or ano_raw == "":
            if not anos and not permitiu_subcabecalho and any(
                isinstance(cell(row, 2 + i), str) for i in range(ncols)
            ):
                permitiu_subcabecalho = True
                r += 1
                continue
            break
        try:
            ano = int(float(ano_raw))
        except (TypeError, ValueError):
            break
        valores = [
            cell(row, 2 + i) if isinstance(cell(row, 2 + i), (int, float)) else None
            for i in range(ncols)
        ]
        anos[str(ano)] = valores
        r += 1
    return anos


def build_historico(service, spreadsheet_id):
    parques = {}
    for park in HISTORICO_PARKS:
        rows = get_values(service, spreadsheet_id, park)
        idx_park = _historico_find_row(rows, park)
        anos_visitacao = _historico_parse_block(rows, idx_park, 12)

        idx_share = _historico_find_row(rows, "Share") if park in HISTORICO_PARKS_COM_SHARE else None
        anos_share = _historico_parse_block(rows, idx_share, 12) if idx_share is not None else {}

        idx_sem = _historico_find_row(rows, "Visão semestral")
        anos_semestral = _historico_parse_block(rows, idx_sem, 2) if idx_sem is not None else {}

        parques[park] = {
            "anos": anos_visitacao,
            "anosShare": anos_share,
            "anosSemestral": anos_semestral,
        }

    # "Consolidado": soma mês a mês dos 6 parques por ano -- trata parque sem dado naquele
    # mês como 0 na soma (mesma convenção da aba "Grupo Cataratas" da planilha: um parque
    # que ainda não existia naquele ano simplesmente não entra na soma), mas o mês fica
    # None se NENHUM dos 6 parques tiver dado.
    def _consolida(campo, ncols):
        todos_anos = set()
        for p in HISTORICO_PARKS:
            todos_anos.update(parques[p][campo].keys())
        out = {}
        for ano in todos_anos:
            soma = []
            for m in range(ncols):
                valores_mes = [parques[p][campo].get(ano, [None] * ncols)[m] for p in HISTORICO_PARKS]
                validos = [v for v in valores_mes if v is not None]
                soma.append(sum(validos) if validos else None)
            out[ano] = soma
        return out

    parques["Consolidado"] = {
        "anos": _consolida("anos", 12),
        "anosShare": {},  # share é sempre por parque -- não faz sentido "somar" percentuais
        "anosSemestral": _consolida("anosSemestral", 2),
    }

    return {"parques": parques, "parquesComShare": HISTORICO_PARKS_COM_SHARE}


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

    service = get_drive_service()

    print("Lendo Visitação Parques 2026...", file=sys.stderr)
    visitacao = build_visitacao(service, cfg["visitacao_parques_id"], cfg["meses_com_dados"])

    # Meses sem resultado real ainda (Agosto em diante, tipicamente): entram no VISITACAO
    # só com a Meta (projecao "Nova2026") no lugar do OBZ -- Realizado fica None/vazio.
    # TEMPORARIO: usa o dado estatico (VISITACAO_META_ESTATICO) em vez de ler ao vivo do
    # Drive, ate a planilha "Visitação Diária Ponderada" ser compartilhada com a service
    # account. Pra trocar pra leitura ao vivo depois, troque a linha abaixo por:
    #   visitacao_meta = build_visitacao_meta(service, cfg["visitacao_diaria_ponderada_id"], cfg["sheet_names"]["visitacao_diaria_gc_v2"], meses_meta)
    meses_meta = [m for m in cfg["meses"] if m not in cfg["meses_com_dados"]]
    if meses_meta:
        print("Usando Meta estática (Ago-Dez) — planilha ainda não conectada ao Drive...", file=sys.stderr)
        visitacao_meta = VISITACAO_META_ESTATICO
        for mes in meses_meta:
            month_number = MONTH_NUMBER[mes]
            n_days = calendar.monthrange(2026, month_number)[1]
            daily = {}
            for park in cfg["parques"]:
                meta_arr = visitacao_meta.get(mes, {}).get(park)
                daily[park] = {
                    "Realizado 2026": [None] * n_days,
                    "Realizado 2025": [None] * n_days,
                    "OBZ 2026": meta_arr if meta_arr is not None else [None] * n_days,
                }
            visitacao[mes] = {
                "monthNumber": month_number,
                "nDays": n_days,
                "summary": {},
                "daily": daily,
                "atrativos": {"daily": {}, "accum": {}},
            }

    print("Lendo Mapa Clima...", file=sys.stderr)
    clima_emoji = build_clima_emoji(
        service, cfg["visitacao_parques_id"], cfg["sheet_names"]["mapa_clima"], cfg["meses_com_dados"]
    )
    for mes, parques in clima_emoji.items():
        for park, emoji_arr in parques.items():
            if park in visitacao[mes]["daily"]:
                visitacao[mes]["daily"][park]["Emoji"] = emoji_arr

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
    # Le TODOS os meses (nao so' meses_com_dados) -- meses futuros ja podem ter lancamentos
    # de custo fixo cadastrados com antecedencia na planilha (ver comentario na funcao).
    invest_mkt_detail = build_invest_mkt_detail(
        service, cfg["investimento_marketing_id"], cfg["meses"]
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

    print("Lendo Mix de Origem...", file=sys.stderr)
    mix_origem = build_mix_origem(
        service, cfg["mix_obz_visitacao_id"], cfg["sheet_names"]["mix_origem"]
    )
    mix_origem_acumulado = build_mix_origem(
        service, cfg["mix_obz_visitacao_id"], cfg["sheet_names"]["mix_origem"], col_offset=11
    )

    print("Lendo Eventos...", file=sys.stderr)
    eventos = build_eventos(
        service, cfg["visitacao_parques_id"], cfg["sheet_names"]["eventos"]
    )

    print("Lendo Histórico...", file=sys.stderr)
    historico_id = cfg.get("visitacao_historico_id", "")
    if historico_id and not historico_id.startswith("COLE_AQUI"):
        try:
            historico = build_historico(service, historico_id)
        except Exception as e:
            # NUNCA deixa a aba Histórico derrubar o resto do pipeline -- se der erro (id
            # errado, aba renomeada, planilha ainda não compartilhada etc.), essa aba fica
            # vazia neste ciclo e o resto do data.json sai normal.
            print(f"AVISO: falha ao ler Histórico ({e}) -- aba fica vazia neste ciclo.", file=sys.stderr)
            historico = {"parques": {}, "parquesComShare": []}
    else:
        print(
            "Histórico: visitacao_historico_id ainda não configurado (placeholder) "
            "-- aba fica vazia até o ID real ser colado no config.json.",
            file=sys.stderr,
        )
        historico = {"parques": {}, "parquesComShare": []}

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
        "MIX_ORIGEM": mix_origem,
        "MIX_ORIGEM_ACUMULADO": mix_origem_acumulado,
        "EVENTOS": eventos,
        "HISTORICO": historico,
    }

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"OK: {args.out} gerado.", file=sys.stderr)


if __name__ == "__main__":
    main()
