#!/usr/bin/env python3
"""
Script para converter artigos .docx para HTML formatado
"""

import os
import re
from docx import Document

CSS_STYLE = """
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: "Times New Roman", serif; font-size: 12pt; line-height: 1.5; background-color: #f0f2f5; color: #1a1a1a; display: flex; flex-direction: column; align-items: center; padding: 40px 20px; }
.paper-page { background-color: #ffffff; width: 210mm; min-height: 297mm; height: auto; margin-bottom: 30px; box-shadow: 0 4px 15px rgba(0, 0, 0, 0.15); position: relative; padding: 25mm 20mm 20mm 30mm; display: flex; flex-direction: column; }
.page-content { flex-grow: 1; position: relative; z-index: 2; }
.page-footer { font-size: 10pt; text-align: right; padding-top: 10px; border-top: 1px solid #f0f2f5; margin-top: 10px; }
h1 { font-size: 14pt; font-weight: bold; text-align: center; margin: 0.5cm 0; line-height: 1.3; }
h2 { font-size: 12pt; font-weight: bold; margin: 0.6cm 0 0.3cm 0; }
p { text-indent: 1.25cm; margin: 0 0 0.25cm 0; text-align: justify; }
.section-header { font-weight: bold; margin: 0.8cm 0 0.4cm 0; }
.keywords-line { text-indent: 0; margin-top: 0.2cm; }
.control-panel { position: fixed; bottom: 30px; right: 30px; background-color: rgba(255, 255, 255, 0.95); padding: 15px 20px; border-radius: 50px; box-shadow: 0 10px 25px rgba(0, 0, 0, 0.2); display: flex; gap: 10px; z-index: 1000; border: 1px solid rgba(226, 232, 240, 0.8); }
button { padding: 10px 18px; border: none; border-radius: 25px; cursor: pointer; font-weight: bold; font-size: 10pt; background-color: #1A365D; color: white; }
button:hover { transform: translateY(-2px); }
@media print { body { background-color: transparent; color: #000000; display: block; padding: 0; } .control-panel { display: none !important; } .paper-page { width: 100% !important; height: auto !important; margin: 0 !important; box-shadow: none !important; padding: 0 !important; page-break-after: always; display: block; background: white; } @page { size: A4; margin: 2.5cm 2cm 2cm 3cm; } }
"""

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>{title}</title>
    <style>
{css}
    </style>
</head>
<body>
    <div class="control-panel no-print">
        <button onclick="window.print()">🖨️ Imprimir / Salvar PDF</button>
    </div>
{content}
</body>
</html>"""


def extract_text_from_docx(doc_path):
    doc = Document(doc_path)
    elements = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        style_name = para.style.name.lower() if para.style else ""
        if "heading 1" in style_name or "title" in style_name:
            elements.append(("h1", text))
        elif "heading 2" in style_name:
            elements.append(("h2", text))
        elif "heading 3" in style_name:
            elements.append(("h3", text))
        elif text.upper() in ["RESUMO", "ABSTRACT"]:
            elements.append(("section-header", text))
        elif text.lower().startswith("palavras-chave") or text.lower().startswith(
            "keywords"
        ):
            elements.append(("keywords", text))
        else:
            elements.append(("p", text))
    return elements


def group_into_pages(elements, chars_per_page=4500):
    pages = []
    current_page = []
    current_chars = 0
    for elem in elements:
        content_len = len(elem[1])
        if current_chars + content_len > chars_per_page and current_page:
            pages.append(current_page)
            current_page = []
            current_chars = 0
        current_page.append(elem)
        current_chars += content_len
    if current_page:
        pages.append(current_page)
    return pages


def elements_to_html(elements):
    html_parts = []
    for elem_type, text in elements:
        if elem_type == "h1":
            html_parts.append(f"<h1>{text}</h1>")
        elif elem_type == "h2":
            html_parts.append(f"<h2>{text}</h2>")
        elif elem_type == "h3":
            html_parts.append(f"<h3>{text}</h3>")
        elif elem_type == "section-header":
            html_parts.append(f'<p class="section-header"><strong>{text}</strong></p>')
        elif elem_type == "keywords":
            html_parts.append(f'<p class="keywords-line"><strong>{text}</strong></p>')
        elif elem_type == "p":
            html_parts.append(f"<p>{text}</p>")
    return "\n            ".join(html_parts)


def convert_article(doc_path, output_dir, article_name):
    print(f"Convertendo: {article_name}")
    try:
        elements = extract_text_from_docx(doc_path)
        if not elements:
            print(f"  [ERRO] Nenhum texto encontrado")
            return False

        title = elements[0][1] if elements else article_name
        pages = group_into_pages(elements)

        html_parts = []
        for i, page_elements in enumerate(pages):
            page_content = elements_to_html(page_elements)
            html_parts.append(f"""
    <section class="paper-page">
        <div class="page-content">
            {page_content}
        </div>
        <div class="page-footer">{i + 1}</div>
    </section>""")

        full_content = "\n".join(html_parts)
        html_doc = HTML_TEMPLATE.format(
            title=title, css=CSS_STYLE, content=full_content
        )

        safe_name = re.sub(r"[^\w\s-]", "", article_name)
        output_path = os.path.join(output_dir, f"{safe_name}.html")

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html_doc)

        print(f"  [OK] Salvo em: {output_path}")
        return True
    except Exception as e:
        print(f"  [ERRO] {str(e)}")
        return False


def main():
    base = "C:/Users/Renato/Documents/Artigos 2026/Prontos-Submissao"
    output_base = (
        "C:/Users/Renato/Documents/Artigos 2026/Prontos-Submissao/HTML_Convertidos"
    )
    os.makedirs(output_base, exist_ok=True)

    # Articles to convert - using direct file paths to avoid encoding issues
    articles = [
        (
            base
            + "/Artigo - Paradoxo da Produtividade IA/paradoxo_produtividade_ia_v2.docx",
            "Paradoxo_Produtividade_IA",
        ),
    ]

    # Add Gestao article using glob to find exact path
    import glob

    gestao_files = glob.glob(base + "/*Gestao*/*FUGA*.docx")
    if gestao_files:
        articles.append((gestao_files[0], "Escravidao_Digital_Fuga_Institucional"))

    # Add Messan articles
    messan_dir = base + "/Artigos a Submeter - Messan"
    articles.append(
        (
            messan_dir + "/Artigo (version final).docx",
            "Messan_Sistemas_Controle_Criatividade",
        )
    )
    articles.append(
        (messan_dir + "/Artigo 2_ a submeter pdf.docx", "Messan_Pensamento_Integrado")
    )

    print("=" * 60)
    print("CONVERSOR DOCX -> HTML (Formato Acadêmico)")
    print("=" * 60)

    results = []
    for doc_path, output_name in articles:
        if os.path.exists(doc_path):
            success = convert_article(doc_path, output_base, output_name)
            results.append((output_name, success))
        else:
            print(f"[AVISO] Arquivo não encontrado: {doc_path}")
            results.append((output_name, False))

    print("\n" + "=" * 60)
    print("RESUMO DA CONVERSÃO:")
    print("=" * 60)
    for name, success in results:
        status = "[OK]" if success else "[FALHOU]"
        print(f"  {status} {name}")

    print(f"\nArquivos HTML salvos em: {output_base}")


if __name__ == "__main__":
    main()
