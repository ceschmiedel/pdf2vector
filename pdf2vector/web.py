"""Servidor local com o mesmo frontend e a mesma API do deploy no Vercel.

Rodar com ``uvicorn pdf2vector.web:app --reload`` e abrir http://localhost:8000.
"""

from __future__ import annotations

from pathlib import Path

from fastapi import FastAPI, HTTPException, Request
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import FileResponse, Response

from . import __version__
from .servico import ErroEntrada, processar

LIMITE_MB = 100
FRONTEND = Path(__file__).resolve().parents[1] / "public" / "index.html"

app = FastAPI(title="pdf2vector", version=__version__,
              description="Recebe um PDF, vetoriza e devolve um Excel com o conteúdo completo.")


@app.get("/", include_in_schema=False)
def index() -> FileResponse:
    if not FRONTEND.exists():
        raise HTTPException(404, "Frontend não encontrado (public/index.html).")
    return FileResponse(FRONTEND)


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
