"""Aplicação web: envia um PDF pelo navegador e baixa a planilha.

Rodar com ``uvicorn pdf2vector.web:app --reload`` e abrir http://localhost:8000.
"""

from __future__ import annotations

from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import HTMLResponse, Response

from . import __version__
from .pipeline import converter

LIMITE_MB = 100
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

app = FastAPI(title="pdf2vector", version=__version__,
              description="Recebe um PDF, vetoriza e devolve um Excel com o conteúdo completo.")

PAGINA = """<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>PDF para Excel</title>
<style>
 :root{--bg:#f5f7fa;--card:#fff;--fg:#1d2733;--muted:#5b6b7c;--accent:#1f4e78;--line:#d9e0e8}
 @media (prefers-color-scheme:dark){:root{--bg:#11161c;--card:#1a222b;--fg:#e6edf3;--muted:#93a1b0;--accent:#5b9bd5;--line:#2c3742}}
 *{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:16px/1.5 system-ui,sans-serif}
 main{max-width:640px;margin:48px auto;padding:0 16px}
 .card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:24px}
 h1{font-size:1.5rem;margin:0 0 4px}p{color:var(--muted);margin:0 0 20px}
 label{display:block;font-weight:600;margin:14px 0 6px}
 input[type=file],select,input[type=number]{width:100%;padding:10px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--fg)}
 .row{display:grid;grid-template-columns:1fr 1fr;gap:12px}
 .check{display:flex;gap:8px;align-items:center;font-weight:400;margin-top:14px}
 button{margin-top:22px;width:100%;padding:12px;border:0;border-radius:8px;background:var(--accent);color:#fff;font-size:1rem;font-weight:600;cursor:pointer}
 button:disabled{opacity:.6;cursor:wait}#status{margin-top:14px;min-height:1.5em;color:var(--muted)}
</style></head><body><main><div class="card">
<h1>PDF → vetores → Excel</h1>
<p>Envie um PDF. A planilha traz metadados, texto por página, blocos e linhas com posição,
formatação, tabelas, imagens, desenhos vetoriais, links, anotações, campos de formulário,
sumário e os vetores (embeddings) de cada trecho.</p>
<form id="f">
 <label for="pdf">Arquivo PDF</label><input id="pdf" name="arquivo" type="file" accept="application/pdf,.pdf" required>
 <div class="row">
  <div><label for="motor">Vetorização</label><select id="motor" name="motor">
   <option value="hash">Rápida (hash, sem modelo)</option>
   <option value="semantico">Semântica (sentence-transformers)</option></select></div>
  <div><label for="chunk">Palavras por trecho</label><input id="chunk" name="tamanho_chunk" type="number" value="200" min="20" max="2000"></div>
 </div>
 <label class="check"><input type="checkbox" name="ocr" value="true"> Aplicar OCR em páginas escaneadas</label>
 <button id="b" type="submit">Gerar Excel</button><div id="status" role="status"></div>
</form></div></main>
<script>
const f=document.getElementById('f'),b=document.getElementById('b'),st=document.getElementById('status');
f.addEventListener('submit',async e=>{e.preventDefault();b.disabled=true;st.textContent='Processando…';
 try{const r=await fetch('/converter',{method:'POST',body:new FormData(f)});
  if(!r.ok){const j=await r.json().catch(()=>({detail:r.statusText}));throw new Error(j.detail)}
  const blob=await r.blob(),a=document.createElement('a');
  const nome=(f.arquivo.files[0].name||'documento.pdf').replace(/\\.pdf$/i,'')+'.xlsx';
  a.href=URL.createObjectURL(blob);a.download=nome;a.click();URL.revokeObjectURL(a.href);
  st.textContent='Pronto: '+nome+' ('+(r.headers.get('X-Paginas')||'?')+' páginas).';
 }catch(err){st.textContent='Erro: '+err.message}finally{b.disabled=false}});
</script></body></html>"""


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return PAGINA


@app.get("/saude")
def saude() -> dict:
    return {"status": "ok", "versao": __version__}


@app.post("/converter")
async def converter_pdf(
    arquivo: UploadFile = File(..., description="PDF a converter"),
    motor: str = Form("hash"),
    tamanho_chunk: int = Form(200),
    sobreposicao: int = Form(40),
    ocr: bool = Form(False),
) -> Response:
    dados = await arquivo.read()
    if len(dados) > LIMITE_MB * 1024 * 1024:
        raise HTTPException(413, f"Arquivo maior que {LIMITE_MB} MB.")
    if not dados.startswith(b"%PDF"):
        raise HTTPException(415, "O arquivo enviado não é um PDF.")
    if motor not in ("hash", "semantico"):
        raise HTTPException(422, "motor deve ser 'hash' ou 'semantico'.")
    if not 20 <= tamanho_chunk <= 2000:
        raise HTTPException(422, "tamanho_chunk deve estar entre 20 e 2000.")

    nome = Path(arquivo.filename or "documento.pdf").name
    try:
        xlsx, doc, _ = await run_in_threadpool(
            converter, dados, nome=nome, motor=motor, tamanho_chunk=tamanho_chunk,
            sobreposicao=sobreposicao, ocr=ocr)
    except RuntimeError as exc:
        raise HTTPException(400, str(exc)) from exc
    except Exception as exc:
        raise HTTPException(422, f"Não foi possível ler o PDF: {exc}") from exc

    saida = Path(nome).stem + ".xlsx"
    return Response(xlsx, media_type=XLSX, headers={
        "Content-Disposition": f"attachment; filename*=UTF-8''{quote(saida)}",
        "X-Paginas": str(doc.metadados["paginas"]),
        "Access-Control-Expose-Headers": "X-Paginas",
    })
