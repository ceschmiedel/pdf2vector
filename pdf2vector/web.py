"""Servidor local: serve o mesmo site publicado no Vercel (que processa o PDF
no navegador) e, além dele, uma API HTTP para automações.

Rodar com ``uvicorn pdf2vector.web:app --reload`` e abrir http://localhost:8000.
"""

from __future__ import annotations

from pathlib import Path

from fastapi import FastAPI, HTTPException, Request
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import Response
from fastapi.staticfiles import StaticFiles

from . import __version__
from .servico import ErroEntrada, processar

LIMITE_MB = 2048
PACOTE = Path(__file__).resolve().parent
PUBLIC = PACOTE.parent / "public"

app = FastAPI(title="pdf2vector", version=__version__,
              description="Recebe um PDF, vetoriza e devolve um Excel com o conteúdo completo.")


@app.get("/saude")
def saude() -> dict:
    return {"status": "ok", "versao": __version__}


@app.post("/api/converter")
async def converter_pdf(request: Request) -> Response:
    """Corpo: o PDF. Query: ``nome``, ``chunk``, ``compacta``, ``motor``, ``ocr``."""
    dados = await request.body()
    if len(dados) > LIMITE_MB * 1024 * 1024:
        raise HTTPException(413, f"Arquivo maior que {LIMITE_MB} MB.")
    try:
        xlsx, headers = await run_in_threadpool(processar, dados, dict(request.query_params))
    except ErroEntrada as exc:
        raise HTTPException(exc.status, str(exc)) from exc
    tipo = headers.pop("Content-Type")
    return Response(xlsx, media_type=tipo, headers=headers)


# Os módulos que o navegador executa vêm direto do pacote, sem passo de build.
app.mount("/py", StaticFiles(directory=PACOTE), name="py")
if PUBLIC.is_dir():
    app.mount("/", StaticFiles(directory=PUBLIC, html=True), name="site")
