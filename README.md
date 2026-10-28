# pdf2vector

Recebe um documento PDF, **extrai todo o conteúdo**, **vetoriza** o texto e gera um
**arquivo Excel (.xlsx)** com as informações completas do documento.

Funciona por linha de comando, por uma página web (upload → download) ou como biblioteca Python.

## O que vai para o Excel

| Aba | Conteúdo |
|---|---|
| **Resumo** | Visão geral: páginas, palavras, tabelas, imagens, links, vetorização usada |
| **Metadados** | Título, autor, assunto, palavras-chave, criador, datas, tamanho, SHA-256, criptografia |
| **Páginas** | Dimensões (pt e mm), rotação, contagens por página e o texto integral de cada página |
| **Blocos de texto** / **Linhas** | Texto na ordem de leitura com a posição (x0, y0, x1, y1) na página |
| **Formatação** | Cada trecho de texto com fonte, tamanho, negrito/itálico, cor e posição |
| **Fontes** | Fontes usadas no documento e quantas vezes aparecem |
| **Tabelas** + **TabN pX** | Todas as tabelas detectadas; cada uma também em aba própria. Números no padrão brasileiro (`1.234,56`, `R$ 10,00`, `12%`) viram números de verdade no Excel |
| **Imagens** | Tamanho em pixels, DPI, espaço de cor, bytes, hash e posição |
| **Desenhos vetoriais** | Linhas, retângulos e curvas do PDF (gráficos, bordas, logotipos) com cores, espessura e pontos |
| **Links**, **Anotações**, **Formulário**, **Sumário** | Hiperlinks e destinos, comentários, campos de formulário preenchidos e bookmarks |
| **Chunks** | O texto dividido em trechos (com sobreposição), página de origem e palavras-chave (TF-IDF) |
| **Vetores** | O embedding de cada trecho, uma coluna por dimensão (`d0 … dN`) |
| **Avisos** | Problemas encontrados (ex.: OCR indisponível) |

### Vetorização

* `hash` (padrão) — vetores de 256 dimensões por *hashing trick* (palavras + bigramas,
  sem acentos e sem stopwords), normalizados. Não baixa nada, roda offline e é determinístico.
* `semantico` — embeddings de um modelo `sentence-transformers`
  (padrão `paraphrase-multilingual-MiniLM-L12-v2`, 384 dimensões, bom para português).
  Captura significado, não só palavras iguais. Requer `pip install -e ".[semantic]"`.

Os vetores da aba **Vetores** podem ser carregados em um banco vetorial (pgvector, Qdrant,
Chroma…) ou usados para busca por similaridade.

## Instalação

```bash
pip install -e ".[web,dev]"
```

Para OCR de PDFs escaneados, instale também o Tesseract (`apt install tesseract-ocr tesseract-ocr-por`).

## Uso

Linha de comando:

```bash
pdf2vector documento.pdf -o documento.xlsx
```

Vários arquivos de uma vez, saída em uma pasta:

```bash
pdf2vector contratos/*.pdf -o saida/
```

Opções úteis: `--motor semantico`, `--chunk 300 --sobreposicao 50`, `--ocr`,
`--sem-vetores`, `--sem-formatacao` (planilhas menores em documentos muito grandes).

Aplicação web:

```bash
uvicorn pdf2vector.web:app --reload
```

Abra http://localhost:8000, envie o PDF e a planilha é baixada. A API fica em
`POST /converter` (multipart, campo `arquivo`) e a documentação em `/docs`.

Como biblioteca:

```python
from pdf2vector import converter

xlsx_bytes, doc, vet = converter("documento.pdf", "documento.xlsx")
for trecho, score in vet.buscar("prazo de pagamento", k=3):
    print(score, trecho.pagina_inicial, trecho.texto[:80])
```

## Testes

```bash
python exemplos/gerar_pdf_exemplo.py exemplo.pdf
```

```bash
pytest
```

## Estrutura

```
pdf2vector/
  extractor.py   # leitura do PDF (PyMuPDF): texto, tabelas, imagens, vetores, links…
  vectorizer.py  # divisão em trechos + embeddings (hash ou semântico) + busca
  excel.py       # montagem da planilha (openpyxl)
  pipeline.py    # PDF → extração → vetorização → Excel
  cli.py         # linha de comando
  web.py         # aplicação web (FastAPI)
```

## Limites

* PDFs escaneados (só imagem) não têm texto: use `--ocr`.
* A detecção de tabelas funciona melhor em tabelas com bordas; tabelas sem linhas podem sair
  como texto nas abas de blocos/linhas.
* Uma célula do Excel aceita até 32.767 caracteres; textos maiores são truncados com aviso
  `…[truncado]` (o conteúdo completo continua dividido nas abas de blocos e chunks).
