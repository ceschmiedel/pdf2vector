"""Gera um PDF de exemplo com texto, tabela, imagem, link, anotação, campo de
formulário, sumário e desenhos vetoriais — útil para testar o pdf2vector.

    python exemplos/gerar_pdf_exemplo.py exemplo.pdf
"""

import sys

import pymupdf


def gerar(destino: str) -> None:
    doc = pymupdf.open()
    p1 = doc.new_page()  # A4 retrato
    p1.insert_text((72, 72), "Relatório Trimestral de Vendas", fontsize=20, fontname="hebo")
    p1.insert_text((72, 100), "As vendas de café cresceram no Sul e no Sudeste.", fontsize=11)
    p1.insert_text((72, 116), "A margem operacional ficou estável no período.", fontsize=11, fontname="tiit")

    # Tabela desenhada com linhas (como a maioria dos PDFs gerados por sistemas)
    x0, y0, cw, rh = 72, 150, 130, 22
    dados = [["Região", "Vendas (R$)", "Variação"],
             ["Sul", "1.234,56", "12%"],
             ["Sudeste", "9.870,00", "-3,5%"],
             ["Nordeste", "450,10", "8%"]]
    for i, linha in enumerate(dados):
        for j, valor in enumerate(linha):
            r = pymupdf.Rect(x0 + j * cw, y0 + i * rh, x0 + (j + 1) * cw, y0 + (i + 1) * rh)
            p1.draw_rect(r, color=(0, 0, 0), width=0.8)
            p1.insert_text((r.x0 + 4, r.y1 - 7), valor, fontsize=10,
                           fontname="hebo" if i == 0 else "helv")

    # Gráfico vetorial simples
    p1.draw_line((72, 400), (472, 400), color=(0.2, 0.2, 0.2), width=1)
    for k, h in enumerate([60, 110, 30]):
        p1.draw_rect(pymupdf.Rect(110 + k * 120, 400 - h, 170 + k * 120, 400),
                     color=None, fill=(0.12, 0.31, 0.47))
    p1.draw_circle((520, 360), 25, color=(0.8, 0.1, 0.1), fill=(1, 0.8, 0.8))

    # Imagem raster
    pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 40, 30), False)
    pix.set_rect(pix.irect, (31, 78, 120))
    p1.insert_image(pymupdf.Rect(72, 430, 152, 490), pixmap=pix)

    # Link externo e anotação
    p1.insert_text((72, 520), "Mais informações no site", fontsize=11)
    p1.insert_link({"kind": pymupdf.LINK_URI, "from": pymupdf.Rect(72, 508, 200, 524),
                    "uri": "https://example.com/vendas"})
    annot = p1.add_text_annot((300, 520), "Conferir números do Nordeste")
    annot.set_info(title="Auditoria")

    # Campo de formulário
    w = pymupdf.Widget()
    w.field_type = pymupdf.PDF_WIDGET_TYPE_TEXT
    w.field_name = "responsavel"
    w.field_value = "Maria Souza"
    w.rect = pymupdf.Rect(72, 560, 272, 580)
    p1.add_widget(w)

    p2 = doc.new_page()
    p2.insert_text((72, 72), "Conclusões", fontsize=16, fontname="hebo")
    texto = ("O mercado de café especial continua em expansão. Recomenda-se ampliar a "
             "distribuição no Nordeste e revisar os custos logísticos no Sudeste. ") * 6
    p2.insert_textbox(pymupdf.Rect(72, 90, 520, 400), texto, fontsize=11)

    doc.set_toc([[1, "Relatório Trimestral de Vendas", 1], [2, "Conclusões", 2]])
    doc.set_metadata({"title": "Relatório Trimestral", "author": "Equipe Comercial",
                      "subject": "Vendas", "keywords": "café, vendas, trimestre"})
    doc.save(destino)


if __name__ == "__main__":
    gerar(sys.argv[1] if len(sys.argv) > 1 else "exemplo.pdf")
