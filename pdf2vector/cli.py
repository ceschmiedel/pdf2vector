"""Linha de comando: ``pdf2vector documento.pdf [-o saida.xlsx]``."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .pipeline import converter


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        prog="pdf2vector",
        description="Extrai e vetoriza todo o conteúdo de PDFs e gera uma planilha Excel por arquivo.")
    ap.add_argument("pdfs", nargs="+", type=Path, help="um ou mais arquivos PDF")
    ap.add_argument("-o", "--saida", type=Path,
                    help="arquivo .xlsx (um PDF) ou pasta de saída (vários PDFs)")
    ap.add_argument("--motor", choices=["hash", "semantico"], default="hash",
                    help="como gerar os vetores (padrão: hash, sem dependências)")
    ap.add_argument("--modelo", default="", help="modelo sentence-transformers para --motor semantico")
    ap.add_argument("--dimensao", type=int, default=256, help="dimensão do vetor no motor hash")
    ap.add_argument("--chunk", type=int, default=200, help="palavras por trecho (padrão 200)")
    ap.add_argument("--sobreposicao", type=int, default=40, help="palavras repetidas entre trechos")
    ap.add_argument("--ocr", action="store_true", help="aplica OCR (Tesseract) em páginas escaneadas")
    ap.add_argument("--sem-vetores", action="store_true", help="só extrai, sem vetorizar")
    ap.add_argument("--sem-formatacao", action="store_true",
                    help="omite a aba de formatação por trecho (planilhas menores)")
    args = ap.parse_args(argv)

    varios = len(args.pdfs) > 1
    if varios and args.saida and args.saida.suffix.lower() == ".xlsx":
        ap.error("com vários PDFs, --saida deve ser uma pasta")

    erros = 0
    for pdf in args.pdfs:
        if args.saida is None:
            destino = pdf.with_suffix(".xlsx")
        elif varios or args.saida.is_dir():
            args.saida.mkdir(parents=True, exist_ok=True)
            destino = args.saida / f"{pdf.stem}.xlsx"
        else:
            destino = args.saida
        try:
            _, doc, vet = converter(
                pdf, destino, motor=args.motor, modelo=args.modelo, dimensao=args.dimensao,
                tamanho_chunk=args.chunk, sobreposicao=args.sobreposicao, ocr=args.ocr,
                vetorizar=not args.sem_vetores, incluir_trechos=not args.sem_formatacao)
        except Exception as exc:
            erros += 1
            print(f"✗ {pdf}: {exc}", file=sys.stderr)
            continue
        m = doc.metadados
        extra = f", {len(vet.chunks)} vetores de {vet.dimensao} dimensões" if vet else ""
        print(f"✓ {pdf} → {destino} ({m['paginas']} páginas, {m['total_palavras']} palavras, "
              f"{m['total_tabelas']} tabelas, {m['total_imagens']} imagens{extra})")
        for aviso in doc.avisos:
            print(f"  aviso: {aviso}", file=sys.stderr)
    return 1 if erros else 0


if __name__ == "__main__":
    raise SystemExit(main())
