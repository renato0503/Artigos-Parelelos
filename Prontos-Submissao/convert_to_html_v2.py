#!/usr/bin/env python3
"""
Script para converter artigos .docx para HTML formatado ( ABNT / APA )
Versão corrigida com formatação impecável
"""

import os
import re
from docx import Document

CSS_STYLE = """
* { margin: 0; padding: 0; box-sizing: border-box; }

body {
    font-family: "Times New Roman", Times, serif;
    font-size: 12pt;
    line-height: 1.6;
    background-color: #f0f2f5;
    color: #1a1a1a;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 40px 20px;
}

/* PÁGINA DE PAPEL A4 */
.paper-page {
    background-color: #ffffff;
    width: 210mm;
    min-height: 297mm;
    height: auto;
    margin-bottom: 30px;
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.12);
    position: relative;
    padding: 30mm 20mm 25mm 30mm;
    display: flex;
    flex-direction: column;
    border-radius: 4px;
}

/* Indicador visual de margens */
.paper-page::before {
    content: "";
    position: absolute;
    top: 25mm;
    left: 25mm;
    right: 15mm;
    bottom: 20mm;
    border: 1px dashed #d0d5dd;
    pointer-events: none;
    z-index: 1;
}

.page-content {
    flex-grow: 1;
    position: relative;
    z-index: 2;
}

.page-footer {
    font-size: 10pt;
    text-align: right;
    color: #666;
    padding-top: 8px;
    border-top: 1px solid #e8e8e8;
    margin-top: 15px;
}

/* TÍTULO DO ARTIGO */
.article-title {
    font-size: 14pt;
    font-weight: bold;
    text-align: center;
    margin: 0 0 0.6cm 0;
    line-height: 1.4;
    color: #1a1a1a;
}

/* SUBTÍTULO / ABSTRACT */
.subtitle {
    font-size: 11pt;
    font-style: italic;
    text-align: center;
    margin: 0 0 0.5cm 0;
    color: #444;
}

/* SEÇÕES DO ARTIGO */
h2 {
    font-size: 12pt;
    font-weight: bold;
    margin: 0.7cm 0 0.35cm 0;
    text-transform: uppercase;
    color: #1a1a1a;
    border-bottom: 1px solid #ddd;
    padding-bottom: 4px;
}

h3 {
    font-size: 11pt;
    font-weight: bold;
    margin: 0.5cm 0 0.25cm 0;
    color: #333;
}

/* TÍTULO DE SEÇÃO PRINCIPAL (1. INTRODUÇÃO, 2. REVISÃO...) */
.section-number {
    font-weight: bold;
}

.section-title {
    font-weight: bold;
    margin: 0.5cm 0 0.3cm 0;
}

/* PARÁGRAFO COMUM */
p {
    text-indent: 1.25cm;
    margin: 0 0 0.3cm 0;
    text-align: justify;
    line-height: 1.6;
}

/* RESUMO / ABSTRACT */
.abstract-section {
    margin: 0.4cm 0;
}

.abstract-header {
    font-weight: bold;
    font-size: 12pt;
    text-align: center;
    margin: 0 0 0.25cm 0;
    text-transform: uppercase;
}

.abstract-text {
    text-indent: 0;
    text-align: justify;
    line-height: 1.4;
    margin: 0 0 0.3cm 0;
    font-size: 11pt;
}

/* KEYWORDS */
.keywords-line {
    text-indent: 0;
    margin: 0.2cm 0 0.4cm 0;
    font-size: 11pt;
}

.keywords-line strong {
    font-weight: bold;
}

/* TABELAS ABNT */
table {
    width: 100%;
    border-collapse: collapse;
    margin: 0.5cm 0;
    font-size: 10pt;
    line-height: 1.3;
}

th {
    border-top: 1px solid black;
    border-bottom: 1px solid black;
    padding: 6px 8px;
    font-weight: bold;
    text-align: center;
    background-color: #f8f8f8;
}

td {
    padding: 5px 8px;
    text-align: left;
    border-bottom: 1px solid #ddd;
}

tr.last-row td {
    border-bottom: 1px solid black;
}

.table-title {
    font-size: 11pt;
    font-weight: bold;
    text-align: left;
    margin: 0.4cm 0 0.15cm 0;
}

.table-source {
    font-size: 9pt;
    font-style: italic;
    text-align: left;
    margin: 0.1cm 0 0.5cm 0;
    color: #555;
}

/* FIGURAS */
.figure-title {
    font-size: 11pt;
    font-weight: bold;
    text-align: left;
    margin: 0.4cm 0 0.15cm 0;
}

.figure-source {
    font-size: 9pt;
    font-style: italic;
    text-align: left;
    margin: 0.1cm 0 0.5cm 0;
    color: #555;
}

.figure-container {
    text-align: center;
    margin: 0.3cm 0;
}

/* REFERÊNCIAS BIBLIOGRÁFICAS */
.references-section {
    margin-top: 0.5cm;
}

.ref-entry {
    text-indent: -1.25cm;
    padding-left: 1.25cm;
    margin: 0 0 0.2cm 0;
    text-align: justify;
    font-size: 10pt;
    line-height: 1.4;
}

/* HIPÓTESES */
.hypothesis {
    text-indent: 0;
    margin: 0.25cm 0 0.25cm 1.25cm;
    font-style: italic;
}

.hypothesis-label {
    font-weight: bold;
    font-style: normal;
}

/* CITAÇÃO DIRETA LONGA */
.quote-block {
    margin: 0.4cm 1.5cm 0.4cm 1.5cm;
    font-size: 11pt;
    text-align: justify;
    border-left: 3px solid #1A365D;
    padding-left: 0.8cm;
    font-style: italic;
    background-color: #fafafa;
    padding: 0.3cm 0.5cm;
}

/* PAINEL DE CONTROLE */
.control-panel {
    position: fixed;
    bottom: 30px;
    right: 30px;
    background-color: rgba(255, 255, 255, 0.98);
    padding: 12px 18px;
    border-radius: 50px;
    box-shadow: 0 8px 30px rgba(0, 0, 0, 0.18);
    display: flex;
    gap: 10px;
    z-index: 1000;
    border: 1px solid rgba(226, 232, 240, 0.9);
    backdrop-filter: blur(10px);
}

button {
    padding: 10px 18px;
    border: none;
    border-radius: 25px;
    cursor: pointer;
    font-weight: 600;
    font-size: 10pt;
    background-color: #1A365D;
    color: white;
    display: flex;
    align-items: center;
    gap: 6px;
    transition: all 0.2s ease;
}

button:hover {
    background-color: #2c5282;
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(26, 54, 93, 0.3);
}

/* PRINT */
@media print {
    body {
        background-color: transparent;
        color: #000000;
        display: block;
        padding: 0;
    }

    .control-panel {
        display: none !important;
    }

    .paper-page {
        width: 100% !important;
        height: auto !important;
        margin: 0 !important;
        box-shadow: none !important;
        padding: 0 !important;
        page-break-after: always;
        display: block;
        background: white;
        border-radius: 0;
    }

    .paper-page::before {
        display: none !important;
    }

    @page {
        size: A4;
        margin: 2.5cm 2cm 2cm 3cm;
    }
}
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


def parse_docx_content(doc_path):
    """Extrai conteúdo estruturado do DOCX"""
    doc = Document(doc_path)
    elements = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        style = para.style.name.lower() if para.style else ""

        # Classificar o elemento
        elem = {"text": text, "style": style, "type": classify_text(text, style)}
        elements.append(elem)

    return elements


def classify_text(text, style):
    """Classifica o texto para formatacao adequada"""

    # Título do artigo (primeiro parágrafo Usually)
    if style in ["title", "heading 1"] or (
        len(text) > 20 and len(text) < 200 and text.isupper()
    ):
        return "article_title"

    # Section headers (RESUMO, ABSTRACT, INTRODUÇÃO, etc.)
    upper_text = text.upper()
    if upper_text in ["RESUMO", "ABSTRACT"]:
        return "abstract_header"
    if upper_text in ["KEYWORDS", "PALAVRAS-CHAVE"]:
        return "keywords"
    if upper_text == "REFERÊNCIAS":
        return "references_header"

    # Keywords line
    if text.lower().startswith("palavras-chave") or text.lower().startswith("keywords"):
        return "keywords_line"

    # Section numbers (1. INTRODUÇÃO, 2. REVISÃO, etc.)
    if re.match(r"^\d+\.\s+[A-ZÀ-Ú]", text):
        return "section_header"

    # Subsection (2.1, 2.2, etc.)
    if re.match(r"^\d+\.\d+\s+", text):
        return "subsection"

    # Hypothesis (H1a:, H1b:, etc.)
    if re.match(r"^H\d+[a-z]?:", text):
        return "hypothesis"

    # Table title pattern
    if text.startswith("Tabela") or re.match(r"^Tabela\s+\d+", text, re.IGNORECASE):
        return "table_title"

    # Figure title pattern
    if text.startswith("Figura") or re.match(r"^Figura\s+\d+", text, re.IGNORECASE):
        return "figure_title"

    # Table source
    if text.lower().startswith("fonte:") or text.lower().startswith("source:"):
        return "table_source"

    # Regular paragraph
    return "paragraph"


def convert_to_html(elements, title, filename):
    """Converte elementos para HTML formatado"""

    pages_content = []
    current_page_elements = []
    current_char_count = 0
    chars_per_page = 4200

    def process_page():
        html_parts = []
        prev_type = None

        for elem in current_page_elements:
            t = elem["type"]
            text = escape_html(elem["text"])

            if t == "article_title":
                html_parts.append(f'<h1 class="article-title">{text}</h1>')
            elif t == "abstract_header":
                html_parts.append(f'<p class="abstract-header">{text}</p>')
            elif t == "keywords_line":
                html_parts.append(
                    f'<p class="keywords-line"><strong>{text}</strong></p>'
                )
            elif t == "section_header":
                # Extrair número e título
                match = re.match(r"^(\d+\.)\s*(.*)", text)
                if match:
                    num = match.group(1)
                    title_text = match.group(2)
                    html_parts.append(
                        f'<h2><span class="section-number">{num}</span> {title_text}</h2>'
                    )
                else:
                    html_parts.append(f"<h2>{text}</h2>")
            elif t == "subsection":
                html_parts.append(f"<h3>{text}</h3>")
            elif t == "hypothesis":
                match = re.match(r"^(H\d+[a-z]?:)\s*(.*)", text)
                if match:
                    label = match.group(1)
                    content = match.group(2)
                    html_parts.append(
                        f'<p class="hypothesis"><span class="hypothesis-label">{label}</span> {content}</p>'
                    )
                else:
                    html_parts.append(f'<p class="hypothesis">{text}</p>')
            elif t == "table_title":
                html_parts.append(f'<p class="table-title">{text}</p>')
            elif t == "figure_title":
                html_parts.append(f'<p class="figure-title">{text}</p>')
            elif t == "table_source":
                html_parts.append(f'<p class="table-source">{text}</p>')
            elif t == "references_header":
                html_parts.append(f'<h2 class="references-section">{text}</h2>')
            elif t == "keywords":
                pass  # Ignorar, já processado junto com abstract
            else:
                html_parts.append(f"<p>{text}</p>")

        return "\n            ".join(html_parts)

    for elem in elements:
        text_len = len(elem["text"])

        # Quebrar página se necessário
        if current_char_count + text_len > chars_per_page and current_page_elements:
            page_html = process_page()
            pages_content.append(page_html)
            current_page_elements = []
            current_char_count = 0

        current_page_elements.append(elem)
        current_char_count += text_len

    # Última página
    if current_page_elements:
        page_html = process_page()
        pages_content.append(page_html)

    # Montar HTML das páginas
    pages_html = []
    for i, page_content in enumerate(pages_content):
        pages_html.append(f"""
    <section class="paper-page">
        <div class="page-content">
            {page_content}
        </div>
        <div class="page-footer">{i + 1}</div>
    </section>""")

    full_content = "\n".join(pages_html)

    return HTML_TEMPLATE.format(title=title, css=CSS_STYLE, content=full_content)


def escape_html(text):
    """Escapa caracteres HTML"""
    text = text.replace("&", "&amp;")
    text = text.replace("<", "&lt;")
    text = text.replace(">", "&gt;")
    text = text.replace('"', "&quot;")
    return text


def convert_article(doc_path, output_dir, article_name):
    """Converte um artigo completo"""
    print(f"Convertendo: {article_name}")

    try:
        elements = parse_docx_content(doc_path)

        if not elements:
            print(f"  [ERRO] Nenhum texto encontrado")
            return False

        # Título do artigo (primeiro elemento útil)
        title = elements[0]["text"] if elements else article_name

        html_content = convert_to_html(elements, title, article_name)

        safe_name = re.sub(r"[^\w\s-]", "", article_name)
        output_path = os.path.join(output_dir, f"{safe_name}.html")

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html_content)

        print(f"  [OK] Salvo em: {output_path}")
        return True

    except Exception as e:
        print(f"  [ERRO] {str(e)}")
        import traceback

        traceback.print_exc()
        return False


def main():
    base = "C:/Users/Renato/Documents/Artigos 2026/Prontos-Submissao"
    output_base = (
        "C:/Users/Renato/Documents/Artigos 2026/Prontos-Submissao/HTML_Corrigido"
    )

    os.makedirs(output_base, exist_ok=True)

    # Artigos para converter
    articles = [
        (
            base
            + "/Artigo - Paradoxo da Produtividade IA/paradoxo_produtividade_ia_v2.docx",
            "Paradoxo_Produtividade_IA",
        ),
    ]

    import glob

    gestao_files = glob.glob(base + "/*Gestao*/*FUGA*.docx")
    if gestao_files:
        articles.append((gestao_files[0], "Escravidao_Digital_Fuga_Institucional"))

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
    print("CONVERSOR DOCX -> HTML (ABNT Impecável)")
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
    print("RESUMO:")
    print("=" * 60)
    for name, success in results:
        status = "[OK]" if success else "[FALHOU]"
        print(f"  {status} {name}")

    print(f"\nSalvo em: {output_base}")


if __name__ == "__main__":
    main()
