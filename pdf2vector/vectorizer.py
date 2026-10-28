"""Vetorização do texto: divide o documento em trechos (chunks) e gera um
vetor numérico (embedding) para cada um.

Dois motores:

* ``hash`` (padrão): vetor esparso de palavras e bigramas via *hashing trick*,
  normalizado (L2). Não precisa de internet nem de modelos baixados, é
  determinístico e serve para busca por similaridade de termos.
* ``semantico``: embeddings densos de um modelo ``sentence-transformers``
  (padrão: ``paraphrase-multilingual-MiniLM-L12-v2``, bom para português).
  Requer ``pip install pdf2vector[semantic]``.
"""

from __future__ import annotations

import hashlib
import math
import re
import unicodedata
from collections import Counter
from dataclasses import dataclass

import numpy as np

MODELO_SEMANTICO_PADRAO = "paraphrase-multilingual-MiniLM-L12-v2"

_PALAVRA = re.compile(r"\w+", re.UNICODE)

STOPWORDS = set("""
a à ao aos as às até com como da das de dela dele deles do dos e é ela elas ele eles em entre
era essa esse esta este eu foi for há isso isto já la lhe mais mas me mesmo meu minha muito na
nas nem no nos nós num numa o os ou para pela pelas pelo pelos por qual quando que quem se sem
ser seu seus sua suas são só também te tem têm um uma umas uns você
the of and to in is it for on that this with as are be by or an at from was were not
""".split())


@dataclass
class Chunk:
    id: int
    pagina_inicial: int
    pagina_final: int
    texto: str
    palavras: int
    caracteres: int
    palavras_chave: str = ""


@dataclass
class Vetorizacao:
    motor: str
    modelo: str
    dimensao: int
    chunks: list[Chunk]
    vetores: np.ndarray  # (n_chunks, dimensao), normalizados

    def buscar(self, consulta: str, k: int = 5) -> list[tuple[Chunk, float]]:
        """Retorna os ``k`` trechos mais parecidos com a consulta (cosseno)."""
        if not self.chunks:
            return []
        q = embed([consulta], motor=self.motor, modelo=self.modelo, dimensao=self.dimensao)[0]
        scores = self.vetores @ q
        ordem = np.argsort(-scores)[:k]
        return [(self.chunks[i], float(scores[i])) for i in ordem]


def _normalizar(texto: str) -> str:
    texto = unicodedata.normalize("NFKD", texto.lower())
    return "".join(c for c in texto if not unicodedata.combining(c))


def tokens(texto: str) -> list[str]:
    return [t for t in _PALAVRA.findall(_normalizar(texto)) if len(t) > 1 and t not in STOPWORDS]


def chunk_text(paginas: list[str], tamanho: int = 200, sobreposicao: int = 40) -> list[Chunk]:
    """Divide o texto em trechos de ~``tamanho`` palavras, com ``sobreposicao``
    palavras repetidas entre trechos vizinhos para não perder contexto.
    Cada trecho guarda de qual página até qual página ele vai."""
    if tamanho <= 0:
        raise ValueError("tamanho deve ser positivo")
    sobreposicao = max(0, min(sobreposicao, tamanho - 1))
    palavras: list[tuple[str, int]] = []
    for n, texto in enumerate(paginas, start=1):
        palavras.extend((w, n) for w in texto.split())

    chunks: list[Chunk] = []
    passo = tamanho - sobreposicao
    inicio = 0
    while inicio < len(palavras):
        janela = palavras[inicio:inicio + tamanho]
        texto = " ".join(w for w, _ in janela)
        chunks.append(Chunk(
            id=len(chunks) + 1, pagina_inicial=janela[0][1], pagina_final=janela[-1][1],
            texto=texto, palavras=len(janela), caracteres=len(texto),
        ))
        if inicio + tamanho >= len(palavras):
            break
        inicio += passo
    return chunks


def _hash_embed(textos: list[str], dimensao: int) -> np.ndarray:
    m = np.zeros((len(textos), dimensao), dtype=np.float32)
    for i, texto in enumerate(textos):
        toks = tokens(texto)
        feats = toks + [f"{a}_{b}" for a, b in zip(toks, toks[1:])]
        for f, c in Counter(feats).items():
            h = int.from_bytes(hashlib.blake2b(f.encode(), digest_size=8).digest(), "little")
            sinal = 1.0 if (h >> 63) & 1 else -1.0
            m[i, h % dimensao] += sinal * (1.0 + math.log(c))
    normas = np.linalg.norm(m, axis=1, keepdims=True)
    normas[normas == 0] = 1.0
    return m / normas


_modelos_cache: dict[str, object] = {}


def _semantic_embed(textos: list[str], modelo: str) -> np.ndarray:
    try:
        from sentence_transformers import SentenceTransformer
    except ImportError as exc:
        raise RuntimeError(
            "Motor 'semantico' requer sentence-transformers: pip install 'pdf2vector[semantic]'"
        ) from exc
    if modelo not in _modelos_cache:
        _modelos_cache[modelo] = SentenceTransformer(modelo)
    vet = _modelos_cache[modelo].encode(textos, normalize_embeddings=True, show_progress_bar=False)
    return np.asarray(vet, dtype=np.float32)


def embed(textos: list[str], *, motor: str = "hash", modelo: str = "", dimensao: int = 256) -> np.ndarray:
    if motor == "hash":
        return _hash_embed(textos, dimensao)
    if motor == "semantico":
        return _semantic_embed(textos, modelo or MODELO_SEMANTICO_PADRAO)
    raise ValueError(f"motor desconhecido: {motor!r} (use 'hash' ou 'semantico')")


def _palavras_chave(chunks: list[Chunk], k: int = 8) -> None:
    """Preenche as palavras-chave de cada trecho por TF-IDF."""
    contagens = [Counter(tokens(c.texto)) for c in chunks]
    df = Counter(t for cnt in contagens for t in cnt)
    n = len(chunks)
    for c, cnt in zip(chunks, contagens):
        pontos = {t: f * (math.log((1 + n) / (1 + df[t])) + 1) for t, f in cnt.items()}
        melhores = sorted(pontos.items(), key=lambda kv: (-kv[1], kv[0]))[:k]
        c.palavras_chave = ", ".join(t for t, _ in melhores)


def vectorize(
    paginas: list[str],
    *,
    motor: str = "hash",
    modelo: str = "",
    dimensao: int = 256,
    tamanho_chunk: int = 200,
    sobreposicao: int = 40,
) -> Vetorizacao:
    chunks = chunk_text(paginas, tamanho_chunk, sobreposicao)
    _palavras_chave(chunks)
    if motor == "semantico":
        modelo = modelo or MODELO_SEMANTICO_PADRAO
    vetores = embed([c.texto for c in chunks], motor=motor, modelo=modelo, dimensao=dimensao)
    if not chunks:
        vetores = np.zeros((0, dimensao), dtype=np.float32)
    return Vetorizacao(
        motor=motor, modelo=modelo or "hashing-trick (palavras + bigramas)",
        dimensao=int(vetores.shape[1]) if vetores.ndim == 2 and vetores.shape[0] else dimensao,
        chunks=chunks, vetores=vetores,
    )
