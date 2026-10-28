"""Extração completa do conteúdo de um PDF com PyMuPDF.

Tudo o que o documento carrega vira registros simples (dicts), prontos para
virar linhas de planilha: metadados, páginas, blocos/linhas/trechos de texto
com posição e fonte, tabelas, imagens, links, anotações, campos de formulário,
sumário (bookmarks) e os desenhos vetoriais (linhas, retângulos, curvas).
"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

import pymupdf


@dataclass
class DocumentData:
    arquivo: str
    metadados: dict[str, Any] = field(default_factory=dict)
    paginas: list[dict] = field(default_factory=list)
    blocos: list[dict] = field(default_factory=list)
    linhas: list[dict] = field(default_factory=list)
    trechos: list[dict] = field(default_factory=list)
    tabelas: list[dict] = field(default_factory=list)
    imagens: list[dict] = field(default_factory=list)
    links: list[dict] = field(default_factory=list)
    anotacoes: list[dict] = field(default_factory=list)
    campos: list[dict] = field(default_factory=list)
    sumario: list[dict] = field(default_factory=list)
    desenhos: list[dict] = field(default_factory=list)
    fontes: list[dict] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)

    @property
    def texto_por_pagina(self) -> list[str]:
        return [p["texto"] for p in self.paginas]


def _r(v: float) -> float:
    return round(float(v), 2)


def _bbox(rect) -> dict:
    x0, y0, x1, y1 = (rect.x0, rect.y0, rect.x1, rect.y1) if hasattr(rect, "x0") else rect
    return {"x0": _r(x0), "y0": _r(y0), "x1": _r(x1), "y1": _r(y1)}


def _cor(c) -> str:
    """Converte cor do PyMuPDF (int sRGB ou tupla 0..1) em #RRGGBB."""
    if c is None:
        return ""
    if isinstance(c, int):
        return f"#{c:06X}"
    if isinstance(c, (tuple, list)):
        if len(c) == 1:
            c = (c[0],) * 3
        if len(c) == 4:  # CMYK -> RGB aproximado
            cc, m, y, k = c
            c = ((1 - cc) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k))
        if len(c) == 3:
            return "#" + "".join(f"{max(0, min(255, round(v * 255))):02X}" for v in c)
    return str(c)


def _estilo(flags: int) -> str:
    partes = []
    if flags & 16:
        partes.append("negrito")
    if flags & 2:
        partes.append("itálico")
    if flags & 8:
        partes.append("mono")
    if flags & 1:
        partes.append("sobrescrito")
    return ", ".join(partes)


_TIPOS_LINK = {
    pymupdf.LINK_NONE: "nenhum",
    pymupdf.LINK_GOTO: "interno",
    pymupdf.LINK_URI: "url",
    pymupdf.LINK_LAUNCH: "arquivo",
    pymupdf.LINK_NAMED: "nomeado",
    pymupdf.LINK_GOTOR: "outro_pdf",
}

_OPERACOES = {"l": "linha", "re": "retângulo", "qu": "quadrilátero", "c": "curva"}


def extract(
    source: str | Path | bytes,
    *,
    nome: str | None = None,
    ocr: bool = False,
    ocr_idioma: str = "por+eng",
    tabelas: bool = True,
    desenhos: bool = True,
    max_desenhos_por_pagina: int = 5000,
    progresso: Callable[[int, int], None] | None = None,
) -> DocumentData:
    """Lê o PDF (caminho ou bytes) e devolve todo o conteúdo estruturado.

    ``ocr=True`` aplica OCR (Tesseract) às páginas sem camada de texto,
    típico de documentos escaneados. ``progresso(pagina, total)`` é chamado
    a cada página lida.
    """
    if isinstance(source, (bytes, bytearray)):
        doc = pymupdf.open(stream=bytes(source), filetype="pdf")
        dados = bytes(source)
        nome = nome or "documento.pdf"
    else:
        path = Path(source)
        doc = pymupdf.open(path)
        dados = path.read_bytes()
        nome = nome or path.name

    out = DocumentData(arquivo=nome)
    try:
        _metadados(doc, dados, out)
        _sumario(doc, out)
        fontes: dict[tuple, dict] = {}
        for page in doc:
            _pagina(page, out, fontes, ocr, ocr_idioma, tabelas, desenhos, max_desenhos_por_pagina)
            if progresso:
                progresso(page.number + 1, doc.page_count)
        out.fontes = sorted(fontes.values(), key=lambda f: (-f["ocorrencias"], f["fonte"]))
        out.metadados["total_caracteres"] = sum(p["caracteres"] for p in out.paginas)
        out.metadados["total_palavras"] = sum(p["palavras"] for p in out.paginas)
        out.metadados["total_tabelas"] = len(out.tabelas)
        out.metadados["total_imagens"] = len(out.imagens)
        out.metadados["total_links"] = len(out.links)
        out.metadados["total_desenhos_vetoriais"] = len(out.desenhos)
    finally:
        doc.close()
    return out


def _metadados(doc, dados: bytes, out: DocumentData) -> None:
    meta = {k: (v or "") for k, v in (doc.metadata or {}).items()}
    out.metadados = {
        "arquivo": out.arquivo,
        "tamanho_bytes": len(dados),
        "sha256": hashlib.sha256(dados).hexdigest(),
        "paginas": doc.page_count,
        "titulo": meta.get("title", ""),
        "autor": meta.get("author", ""),
        "assunto": meta.get("subject", ""),
        "palavras_chave": meta.get("keywords", ""),
        "criador": meta.get("creator", ""),
        "produtor": meta.get("producer", ""),
        "criado_em": meta.get("creationDate", ""),
        "modificado_em": meta.get("modDate", ""),
        "formato": meta.get("format", ""),
        "criptografado": bool(doc.is_encrypted),
        "e_formulario": bool(doc.is_form_pdf),
        "anexos_embutidos": doc.embfile_count(),
    }


def _sumario(doc, out: DocumentData) -> None:
    for nivel, titulo, pagina, *_ in doc.get_toc(simple=False):
        out.sumario.append({"nivel": nivel, "titulo": titulo, "pagina": pagina})


def _pagina(page, out, fontes, ocr, ocr_idioma, tabelas, desenhos, max_desenhos):
    n = page.number + 1
    textpage = None
    usou_ocr = False
    if ocr and not page.get_text("text").strip():
        try:
            textpage = page.get_textpage_ocr(language=ocr_idioma, full=True)
            usou_ocr = True
        except Exception as exc:  # Tesseract ausente ou idioma não instalado
            out.avisos.append(f"Página {n}: OCR indisponível ({exc})")

    texto = page.get_text("text", textpage=textpage, sort=True)
    d = page.get_text("dict", textpage=textpage, sort=True)

    for bi, bloco in enumerate(d["blocks"]):
        if bloco.get("type") != 0:
            continue
        linhas_bloco = []
        for li, linha in enumerate(bloco["lines"]):
            spans = linha["spans"]
            txt_linha = "".join(s["text"] for s in spans)
            linhas_bloco.append(txt_linha)
            out.linhas.append({
                "pagina": n, "bloco": bi, "linha": li, "texto": txt_linha,
                **_bbox(linha["bbox"]),
                "direcao": f"({_r(linha['dir'][0])}, {_r(linha['dir'][1])})",
            })
            for si, s in enumerate(spans):
                if not s["text"].strip():
                    continue
                out.trechos.append({
                    "pagina": n, "bloco": bi, "linha": li, "trecho": si,
                    "texto": s["text"], "fonte": s["font"], "tamanho": _r(s["size"]),
                    "estilo": _estilo(s["flags"]), "cor": _cor(s["color"]),
                    **_bbox(s["bbox"]),
                })
                chave = (s["font"], _r(s["size"]))
                f = fontes.setdefault(chave, {"fonte": s["font"], "tamanho": _r(s["size"]),
                                              "estilo": _estilo(s["flags"]), "ocorrencias": 0,
                                              "caracteres": 0})
                f["ocorrencias"] += 1
                f["caracteres"] += len(s["text"])
        out.blocos.append({
            "pagina": n, "bloco": bi, "texto": "\n".join(linhas_bloco),
            "linhas": len(linhas_bloco), **_bbox(bloco["bbox"]),
        })

    n_tabelas = _tabelas(page, n, out) if tabelas else 0
    n_imagens = _imagens(page, n, out)
    n_links = _links(page, n, out)
    _anotacoes(page, n, out)
    _campos(page, n, out)
    n_desenhos = _desenhos(page, n, out, max_desenhos) if desenhos else 0

    rect = page.rect
    out.paginas.append({
        "pagina": n,
        "largura_pt": _r(rect.width), "altura_pt": _r(rect.height),
        "largura_mm": _r(rect.width * 25.4 / 72), "altura_mm": _r(rect.height * 25.4 / 72),
        "rotacao": page.rotation,
        "caracteres": len(texto), "palavras": len(texto.split()),
        "blocos_texto": sum(1 for b in d["blocks"] if b.get("type") == 0),
        "tabelas": n_tabelas, "imagens": n_imagens, "links": n_links,
        "desenhos_vetoriais": n_desenhos,
        "ocr": "sim" if usou_ocr else "não",
        "sem_texto": "sim" if not texto.strip() else "não",
        "texto": texto,
    })


def _tabelas(page, n, out) -> int:
    try:
        achadas = page.find_tables()
    except Exception as exc:
        out.avisos.append(f"Página {n}: falha ao detectar tabelas ({exc})")
        return 0
    count = 0
    for t in achadas.tables:
        linhas = _celulas(page, t)
        if not linhas or not any(any(c.strip() for c in row) for row in linhas):
            continue
        cab = None
        try:
            if t.header and not t.header.external:
                cab = linhas[0]
        except Exception:
            cab = None
        count += 1
        out.tabelas.append({
            "id": len(out.tabelas) + 1, "pagina": n, "indice_na_pagina": count,
            "linhas": t.row_count, "colunas": t.col_count, **_bbox(t.bbox),
            "cabecalho": cab, "dados": linhas,
        })
    return count


def _celulas(page, tabela) -> list[list[str]]:
    """Lê o texto de cada célula pela sua área na página. É mais fiel que o
    ``extract()`` do PyMuPDF, que embaralha pontuação (ex.: "1.234,56")."""
    try:
        return [[("" if c is None else page.get_textbox(c).strip()) for c in row.cells]
                for row in tabela.rows]
    except Exception:
        return [[("" if c is None else str(c)) for c in row] for row in tabela.extract()]


def _imagens(page, n, out) -> int:
    count = 0
    for info in page.get_image_info(hashes=True, xrefs=True):
        count += 1
        out.imagens.append({
            "pagina": n, "indice": count, "xref": info.get("xref", 0),
            "largura_px": info.get("width"), "altura_px": info.get("height"),
            "bits_por_componente": info.get("bpc"),
            "espaco_cor": {1: "cinza", 3: "RGB", 4: "CMYK"}.get(info.get("colorspace"), str(info.get("colorspace"))),
            "dpi_x": info.get("xres"), "dpi_y": info.get("yres"),
            "bytes": info.get("size"),
            "md5": (info.get("digest") or b"").hex(),
            **_bbox(info["bbox"]),
        })
    return count


def _links(page, n, out) -> int:
    count = 0
    for lk in page.get_links():
        count += 1
        destino = lk.get("uri") or lk.get("file") or lk.get("name") or ""
        if lk.get("kind") == pymupdf.LINK_GOTO:
            destino = f"página {lk.get('page', -1) + 1}"
        texto = page.get_textbox(lk["from"]).strip()
        out.links.append({
            "pagina": n, "tipo": _TIPOS_LINK.get(lk.get("kind"), str(lk.get("kind"))),
            "destino": destino, "texto_ancora": texto, **_bbox(lk["from"]),
        })
    return count


def _anotacoes(page, n, out) -> None:
    for a in page.annots() or []:
        info = a.info
        out.anotacoes.append({
            "pagina": n, "tipo": a.type[1], "conteudo": info.get("content", ""),
            "autor": info.get("title", ""), "assunto": info.get("subject", ""),
            "criada_em": info.get("creationDate", ""), **_bbox(a.rect),
        })


def _campos(page, n, out) -> None:
    for w in page.widgets() or []:
        out.campos.append({
            "pagina": n, "nome": w.field_name, "tipo": w.field_type_string,
            "valor": "" if w.field_value is None else str(w.field_value),
            "rotulo": w.field_label or "", **_bbox(w.rect),
        })


def _desenhos(page, n, out, limite) -> int:
    try:
        paths = page.get_drawings()
    except Exception as exc:
        out.avisos.append(f"Página {n}: falha ao ler desenhos vetoriais ({exc})")
        return 0
    if len(paths) > limite:
        out.avisos.append(
            f"Página {n}: {len(paths)} desenhos vetoriais; apenas os primeiros {limite} foram listados")
    for i, p in enumerate(paths[:limite]):
        ops = [it[0] for it in p.get("items", [])]
        resumo = ", ".join(f"{ops.count(k)} {v}" for k, v in _OPERACOES.items() if ops.count(k))
        pontos = []
        for it in p.get("items", []):
            for obj in it[1:]:
                if isinstance(obj, pymupdf.Point):
                    pontos.append(f"{_r(obj.x)},{_r(obj.y)}")
                elif isinstance(obj, pymupdf.Rect):
                    pontos.append(f"{_r(obj.x0)},{_r(obj.y0)} {_r(obj.x1)},{_r(obj.y1)}")
                elif isinstance(obj, pymupdf.Quad):
                    pontos.append(" ".join(f"{_r(q.x)},{_r(q.y)}" for q in (obj.ul, obj.ur, obj.lr, obj.ll)))
        out.desenhos.append({
            "pagina": n, "indice": i + 1, "tipo": {"f": "preenchido", "s": "contorno",
                                                    "fs": "preenchido+contorno"}.get(p.get("type"), p.get("type")),
            "elementos": resumo, "cor_contorno": _cor(p.get("color")),
            "cor_preenchimento": _cor(p.get("fill")),
            "espessura": _r(p["width"]) if p.get("width") is not None else "",
            "fechado": "sim" if p.get("closePath") else "não",
            **_bbox(p["rect"]), "pontos": " | ".join(pontos),
        })
    return min(len(paths), limite)
