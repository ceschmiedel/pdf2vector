"""Fluxo completo: PDF -> extração -> vetorização -> Excel."""

from __future__ import annotations

from pathlib import Path

from .excel import write_excel
from .extractor import DocumentData, extract
from .vectorizer import Vetorizacao, vectorize


def converter(
    pdf: str | Path | bytes,
    destino: str | Path | None = None,
    *,
    nome: str | None = None,
    motor: str = "hash",
    modelo: str = "",
    dimensao: int = 256,
    tamanho_chunk: int = 200,
    sobreposicao: int = 40,
    ocr: bool = False,
    vetorizar: bool = True,
    incluir_trechos: bool = True,
) -> tuple[bytes, DocumentData, Vetorizacao | None]:
    """Converte um PDF em planilha. Devolve (bytes do .xlsx, dados, vetorização)."""
    doc = extract(pdf, nome=nome, ocr=ocr)
    vet = None
    if vetorizar:
        vet = vectorize(doc.texto_por_pagina, motor=motor, modelo=modelo, dimensao=dimensao,
                        tamanho_chunk=tamanho_chunk, sobreposicao=sobreposicao)
    xlsx = write_excel(doc, vet, destino, incluir_trechos=incluir_trechos)
    return xlsx, doc, vet
