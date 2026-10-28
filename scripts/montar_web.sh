#!/bin/sh
# Copia os módulos Python que o navegador executa (via Pyodide) para public/py/.
# Roda no build do Vercel; localmente o servidor (pdf2vector.web) serve direto da pasta.
set -e
cd "$(dirname "$0")/.."
mkdir -p public/py
for m in __init__ extractor vectorizer excel pipeline; do
  cp "pdf2vector/$m.py" "public/py/$m.py"
done
echo "public/py pronto"
