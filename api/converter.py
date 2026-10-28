"""Função serverless do Vercel: POST /api/converter

Corpo: o PDF (application/pdf). Query: nome, chunk, compacta, motor.
Resposta: a planilha .xlsx, com o resumo nos cabeçalhos X-*.
"""

import json
import sys
from http.server import BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import parse_qs, urlparse

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from pdf2vector.servico import LIMITE_RESPOSTA, ErroEntrada, processar  # noqa: E402

LIMITE_ENTRADA = 4_400_000  # o Vercel corta corpos acima de 4,5 MB


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        tamanho = int(self.headers.get("Content-Length") or 0)
        if tamanho > LIMITE_ENTRADA:
            return self._erro(413, "O PDF passa de 4,4 MB, o limite da versão online.")
        dados = self.rfile.read(tamanho)
        params = parse_qs(urlparse(self.path).query)
        try:
            xlsx, headers = processar(dados, params, limite_resposta=LIMITE_RESPOSTA)
        except ErroEntrada as exc:
            return self._erro(exc.status, str(exc))
        self.send_response(200)
        for k, v in headers.items():
            self.send_header(k, v)
        self.send_header("Content-Length", str(len(xlsx)))
        self.end_headers()
        self.wfile.write(xlsx)

    def do_GET(self):
        self._json(200, {"status": "ok", "uso": "POST com o PDF no corpo"})

    def _erro(self, status, mensagem):
        self._json(status, {"detail": mensagem})

    def _json(self, status, corpo):
        dados = json.dumps(corpo, ensure_ascii=False).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(dados)))
        self.end_headers()
        self.wfile.write(dados)
