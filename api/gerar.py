# api/gerar.py
import json
import re
import openpyxl
from http.server import BaseHTTPRequestHandler
from pathlib import Path

# Planilha esperada em: dados/GABARITO.xlsx
PLANILHA = Path(__file__).parent.parent / "dados" / "GABARITO.xlsx"

MAX_WILDCARDS = 5         # limite de asteriscos
MAX_COMBINACOES = 20000   # limite de combinações retornadas

def carregar_gabarito():
    if not PLANILHA.exists():
        raise FileNotFoundError(f"Planilha não encontrada: {PLANILHA}")
    wb = openpyxl.load_workbook(PLANILHA)
    ws = wb.active
    dados = []
    # Colunas: B=MÊS, C=ANO, F=CÓD IDENTIFICADOR (conforme você informou)
    for row in ws.iter_rows(min_row=2, values_only=True):
        mes = (str(row[1]).zfill(2) if row[1] is not None else "")
        ano = (str(row[2]) if row[2] is not None else "")
        codigo = (str(row[5]) if row[5] is not None else "")
        if mes and ano and codigo:
            dados.append({"mes": mes, "ano": ano, "codigo": codigo})
    return dados

def buscar_codigos_por_mes_ano(mes, ano):
    dados = carregar_gabarito()
    return [d["codigo"] for d in dados if d["mes"] == mes and d["ano"] == ano]

def validar_formato(caf: str):
    if not caf:
        return False, 'Informe um CAF.'
    if not re.fullmatch(r'[A-Z0-9.*]+', caf):
        return False, 'Use apenas A–Z, 0–9, ponto (.) e asterisco (*).'
    if len(caf) < 8:
        return False, 'CAF muito curto.'
    return True, ''

def gerar_combinacoes(mask: str, codigos_validos):
    wc = mask.count('*')
    total = 10 ** wc
    if wc > MAX_WILDCARDS:
        raise ValueError(f'Excesso de curingas ({wc}). Máximo: {MAX_WILDCARDS}')
    if total > MAX_COMBINACOES:
        raise ValueError(f'Muitas combinações ({total:,}). Máximo: {MAX_COMBINACOES:,}')

    combos = []
    for i in range(total):
        digits = str(i).zfill(wc)
        p = 0
        out = []
        for ch in mask:
            if ch == '*':
                out.append(digits[p]); p += 1
            else:
                out.append(ch)
        code = ''.join(out)
        # aceita se contiver qualquer código válido da planilha
        for cod_val in codigos_validos:
            if cod_val in code:
                combos.append(code)
                break
    return combos

def _send_json(handler: BaseHTTPRequestHandler, status: int, payload: dict):
    body = json.dumps(payload, ensure_ascii=False).encode('utf-8')
    handler.send_response(status)
    handler.send_header('Content-Type', 'application/json; charset=utf-8')
    # CORS
    handler.send_header('Access-Control-Allow-Origin', '*')
    handler.send_header('Access-Control-Allow-Headers', 'Content-Type')
    handler.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
    handler.send_header('Content-Length', str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)

class handler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.end_headers()

    def do_POST(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            raw = self.rfile.read(length) if length else b'{}'
            data = json.loads(raw.decode('utf-8') or '{}')
            caf = (data.get('caf') or '').strip().upper()

            ok, msg = validar_formato(caf)
            if not ok:
                return _send_json(self, 400, {'erro': msg})

            mes = caf[2:4]
            ano = caf[4:8]
            if not (mes.isdigit() and ano.isdigit()):
                return _send_json(self, 400, {'erro': 'Não foi possível extrair MÊS/ANO do CAF.'})

            codigos_validos = buscar_codigos_por_mes_ano(mes, ano)
            if not codigos_validos:
                return _send_json(self, 404, {'erro': f'Nenhum código encontrado para {mes}/{ano}.'})

            combos = gerar_combinacoes(caf, codigos_validos)
            return _send_json(self, 200, {'combos': combos})

        except FileNotFoundError as e:
            return _send_json(self, 500, {'erro': str(e)})
        except Exception as e:
            return _send_json(self, 500, {'erro': f'Erro interno: {e}'})
