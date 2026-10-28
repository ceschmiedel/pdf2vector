"""Conversão chamada pela API HTTP (Vercel e servidor local).

O PDF chega como corpo bruto da requisição e as opções na query string, assim
a função do Vercel não precisa de parser multipart nem de framework.
"""

from __future__ import annotations

from pathlib import Path
from urllib.parse import quote

from .pipeline import converter

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
# O Vercel recusa respostas acima de 4,5 MB; acima disso a planilha é refeita sem
# a aba de formatação por trecho, que é a maior delas.
LIMITE_RESPOSTA = 4_300_000


class ErroEntrada(ValueError):
    def __init__(self, status: int, mensagem: str):
        super().__init__(mensagem)
        self.status = status


def _opcao(params: dict, chave: str, padrao: str) -> str:
    v = params.get(chave, padrao)
    return (v[0] if isinstance(v, list) else v) or padrao


def processar(dados: bytes, params: dict, limite_resposta: int | None = None
              ) -> tuple[bytes, dict[str, str]]:
    """Valida, converte e devolve (xlsx, cabeçalhos HTTP)."""
    if not dados:
        raise ErroEntrada(400, "Nenhum arquivo recebido.")
    if not dados.lstrip()[:5].startswith(b"%PDF"):
        raise ErroEntrada(415, "O arquivo enviado não é um PDF.")
    try:
        chunk = int(_opcao(params, "chunk", "200"))
    except ValueError:
        raise ErroEntrada(422, "chunk deve ser um número inteiro.") from None
    if not 20 <= chunk <= 2000:
        raise ErroEntrada(422, "chunk deve estar entre 20 e 2000.")
    motor = _opcao(params, "motor", "hash")
    if motor not in ("hash", "semantico"):
        raise ErroEntrada(422, "motor deve ser 'hash' ou 'semantico'.")
    compacta = _opcao(params, "compacta", "0") in ("1", "true", "sim")
    ocr = _opcao(params, "ocr", "0") in ("1", "true", "sim")
    nome = Path(_opcao(params, "nome", "documento.pdf")).name

    opcoes = dict(nome=nome, motor=motor, tamanho_chunk=chunk,
                  sobreposicao=max(0, chunk // 5), ocr=ocr)
    try:
        xlsx, doc, vet = converter(dados, incluir_trechos=not compacta, **opcoes)
        if limite_resposta and len(xlsx) > limite_resposta and not compacta:
            compacta = True
            xlsx, doc, vet = converter(dados, incluir_trechos=False, **opcoes)
    except RuntimeError as exc:
        raise ErroEntrada(400, str(exc)) from exc
    except Exception as exc:
        raise ErroEntrada(422, f"Não foi possível ler o PDF: {exc}") from exc
    if limite_resposta and len(xlsx) > limite_resposta:
        raise ErroEntrada(413, "A planilha deste PDF passa do limite de 4,5 MB da hospedagem. "
                               "Use a linha de comando para documentos grandes.")

    m = doc.metadados
    resumo = {
        "X-Paginas": m["paginas"], "X-Palavras": m["total_palavras"],
        "X-Tabelas": m["total_tabelas"], "X-Imagens": m["total_imagens"],
        "X-Links": m["total_links"], "X-Desenhos": m["total_desenhos_vetoriais"],
        "X-Chunks": len(vet.chunks) if vet else 0, "X-Dimensao": vet.dimensao if vet else 0,
        "X-Compacta": "1" if compacta else "0",
    }
    saida = Path(nome).stem + ".xlsx"
    headers = {
        "Content-Type": XLSX,
        "Content-Disposition": f"attachment; filename*=UTF-8''{quote(saida)}",
        "Access-Control-Expose-Headers": ", ".join(resumo),
        **{k: str(v) for k, v in resumo.items()},
    }
    return xlsx, headers
