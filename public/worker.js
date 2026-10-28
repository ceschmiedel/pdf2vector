// Roda o pdf2vector (Python) dentro do navegador com Pyodide.
// O PDF nunca sai da máquina do usuário e não há limite de tamanho do servidor:
// o que limita é a memória do navegador.

const MODULOS = ["__init__", "extractor", "vectorizer", "excel", "pipeline"];
const WHEELS = ["et_xmlfile-2.0.0-py3-none-any.whl", "openpyxl-3.1.5-py2.py3-none-any.whl"];

let pyodide = null;
let preparando = null;

function avisar(tipo, dados) {
  self.postMessage({ tipo, ...dados });
}

async function preparar(indexURL) {
  avisar("estado", { texto: "Carregando o motor Python (só na primeira vez)…" });
  importScripts(indexURL + "pyodide.js");
  pyodide = await loadPyodide({ indexURL });

  avisar("estado", { texto: "Carregando PyMuPDF e NumPy…" });
  await pyodide.loadPackage(["pymupdf", "numpy", "micropip"]);
  const micropip = pyodide.pyimport("micropip");
  const base = new URL("vendor/", self.location.href).href;
  await micropip.install(WHEELS.map((w) => base + w));

  avisar("estado", { texto: "Carregando o pdf2vector…" });
  pyodide.FS.mkdirTree("/app/pdf2vector");
  for (const m of MODULOS) {
    const r = await fetch(new URL(`py/${m}.py`, self.location.href));
    if (!r.ok) throw new Error(`Não foi possível baixar py/${m}.py (${r.status}).`);
    pyodide.FS.writeFile(`/app/pdf2vector/${m}.py`, await r.text());
  }
  pyodide.runPython(`
import sys
sys.path.insert(0, "/app")
from pdf2vector.extractor import extract
from pdf2vector.vectorizer import vectorize
from pdf2vector.excel import write_excel

def rodar(dados, nome, chunk, compacta, progresso, etapa):
    pdf = dados.to_bytes()
    doc = extract(pdf, nome=nome, progresso=progresso)
    del pdf
    etapa("Vetorizando o texto…")
    vet = vectorize(doc.texto_por_pagina, tamanho_chunk=chunk, sobreposicao=chunk // 5)
    etapa("Montando a planilha…")
    xlsx = write_excel(doc, vet, incluir_trechos=not compacta)
    m = doc.metadados
    resumo = {
        "paginas": m["paginas"], "palavras": m["total_palavras"], "tabelas": m["total_tabelas"],
        "imagens": m["total_imagens"], "links": m["total_links"],
        "desenhos": m["total_desenhos_vetoriais"], "chunks": len(vet.chunks),
        "dimensao": vet.dimensao, "avisos": len(doc.avisos),
    }
    return xlsx, resumo
`);
}

self.onmessage = async ({ data }) => {
  try {
    if (data.tipo === "preparar") {
      preparando = preparando || preparar(data.indexURL);
      await preparando;
      avisar("pronto", {});
      return;
    }
    if (data.tipo === "converter") {
      await preparando;
      const inicio = performance.now();
      let ultimo = 0;
      const progresso = (pagina, total) => {
        const agora = performance.now();
        if (pagina === total || agora - ultimo > 150) {
          ultimo = agora;
          avisar("progresso", { pagina, total, segundos: (agora - inicio) / 1000 });
        }
      };
      const etapa = (texto) => avisar("estado", { texto });
      etapa("Lendo o PDF…");
      const rodar = pyodide.globals.get("rodar");
      const ret = rodar(new Uint8Array(data.pdf), data.nome, data.chunk, data.compacta, progresso, etapa);
      const [xlsxPy, resumoPy] = ret.toJs({ dict_converter: Object.fromEntries });
      ret.destroy();
      rodar.destroy();
      const xlsx = xlsxPy.slice ? xlsxPy : new Uint8Array(xlsxPy);
      avisar("concluido", {
        xlsx, resumo: resumoPy, segundos: (performance.now() - inicio) / 1000,
      });
      return;
    }
  } catch (err) {
    avisar("erro", { texto: traduzir(err) });
  }
};

function traduzir(err) {
  const msg = String(err && err.message ? err.message : err);
  if (/memory|out of bounds|Aborted/i.test(msg)) {
    return "O navegador ficou sem memória para este PDF. Marque \"Planilha compacta\" e tente de novo, ou use a linha de comando do pdf2vector.";
  }
  const linhas = msg.trim().split("\n");
  return linhas[linhas.length - 1] || msg;
}
