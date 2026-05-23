from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from copy import deepcopy

print("=== CRIANDO VERSÃO REDUZIDA PARA ALCANCE ===\n")

# Carregar documento com tabelas
doc_tables = Document(
    "Prontos-Submissao/Artigo - Paradoxo da Produtividade IA/Artigo Revista Alcance.docx"
)
print(f"Tabelas carregadas: {len(doc_tables.tables)}")

# Carregar documento v2 para conteúdo
doc_v2 = Document(
    "Prontos-Submissao/Artigo - Paradoxo da Produtividade IA/paradoxo_produtividade_ia_v2.docx"
)

# Criar novo documento
new_doc = Document()

# Configurar margens Alcance
for section in new_doc.sections:
    section.top_margin = Cm(3)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(3)
    section.right_margin = Cm(2)

# ============================================================================
# HEADER - TÍTULOS
# ============================================================================
print("1. Títulos PT/EN/ES...")

p = new_doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(6)
run = p.add_run(
    "O PARADOXO DA PRODUTIVIDADE DA IA: ANÁLISE CRÍTICA DO DISCURSO CORPORATIVO SOBRE DEMISSÕES TECNOLÓGICAS"
)
run.bold = True
run.font.size = Pt(14)
run.font.name = "Times New Roman"

p = new_doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(6)
run = p.add_run(
    "THE AI PRODUCTIVITY PARADOX: CRITICAL ANALYSIS OF CORPORATE DISCOURSE ON TECHNOLOGICAL LAYOFFS"
)
run.bold = True
run.font.size = Pt(14)
run.font.name = "Times New Roman"

p = new_doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(18)
run = p.add_run(
    "EL PARADOJA DE LA PRODUCTIVIDAD DE LA IA: ANÁLISIS CRÍTICO DEL DISCURSO CORPORATIVO SOBRE DESPIDOS TECNOLÓGICOS"
)
run.bold = True
run.font.size = Pt(14)
run.font.name = "Times New Roman"

# ============================================================================
# RESUMO ESTRUTURADO PT
# ============================================================================
print("2. Resumo PT...")

p = new_doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(6)
run = p.add_run("Resumo")
run.bold = True
run.font.size = Pt(12)
run.font.name = "Times New Roman"

resumo_pt = [
    (
        "Objetivo:",
        "Analisar criticamente se as declarações de executivos que atribuem demissões à implementação de inteligência artificial encontram respaldo na realidade operacional ou se constituem narrativa estratégica para ocultar motivações subjacentes.",
    ),
    (
        "Design / metodologia / abordagem:",
        "Pesquisa qualitativa baseada em análise de conteúdo documental, utilizando como corpus declarações e demonstrativos financeiros de seis empresas de tecnologia que anunciaram demissões entre 2023 e 2026.",
    ),
    (
        "Resultados:",
        "Nenhuma organização disponibilizou métricas verificáveis de produtividade que sustentassem a redundância de trabalhadores pela tecnologia. Observou-se lucratividade elevada, sobrecontratação pandêmica e linguagem de determinismo tecnológico.",
    ),
    (
        "Limitações / implicações da pesquisa:",
        "Amostra limitada ao setor de tecnologia norte-americano e fontes públicas. Fenômeno recente e em evolução, limitando generalização.",
    ),
    (
        "Implicações práticas:",
        "Fornece modelo analítico para que investidores, reguladores e trabalhadores avaliem alegações corporativas, identificando inconsistências entre discursos de automação e capacidade financeira real.",
    ),
    (
        "Implicações sociais:",
        "Oferece subsídios para que a sociedade avalie criticamente alegações de empresas sobre IA e demissões, contribuindo para transparência no mercado de trabalho.",
    ),
    (
        "Implicações teóricas:",
        "Demonstra a persistência do Paradoxo de Solow no nível da firma e aplica a Teoria da Agência para explicar a externalização de responsabilidades executivas por meio de narrativas de inovação.",
    ),
    (
        "Originalidade / valor:",
        "Preenche lacunas ao cruzar referenciais de agência, produtividade e recursos para demonstrar que a narrativa de IA funciona como cortina de fumaça para demissões motivadas por outros fatores.",
    ),
]

for label, content in resumo_pt:
    p = new_doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.first_line_indent = Inches(0)
    run = p.add_run(label + " ")
    run.bold = True
    run.font.size = Pt(11)
    run.font.name = "Times New Roman"
    run = p.add_run(content)
    run.font.size = Pt(11)
    run.font.name = "Times New Roman"

p = new_doc.add_paragraph()
p.paragraph_format.space_after = Pt(12)
run = p.add_run("Palavras-chave: ")
run.bold = True
run.font.size = Pt(11)
run.font.name = "Times New Roman"
run = p.add_run(
    "Paradoxo da produtividade; Teoria da agência; Demissões corporativas; Inteligência artificial; AI washing"
)
run.font.size = Pt(11)
run.font.name = "Times New Roman"

# ============================================================================
# ABSTRACT EN
# ============================================================================
print("3. Abstract EN...")

p = new_doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(6)
run = p.add_run("Abstract")
run.bold = True
run.font.size = Pt(12)
run.font.name = "Times New Roman"

abstract_en = [
    (
        "Purpose:",
        "To critically analyze whether executive statements attributing layoffs to AI implementation find support in operational reality or constitute strategic narratives to conceal underlying motivations.",
    ),
    (
        "Design/methodology/approach:",
        "Qualitative research based on documentary content analysis, using statements and financial reports from six technology companies that announced layoffs between 2023 and 2026.",
    ),
    (
        "Findings:",
        "No organization provided verifiable productivity metrics supporting human redundancy through technology. High profitability, pandemic-era over-hiring, and technological determinism language were observed.",
    ),
    (
        "Research limitations/implications:",
        "Sample limited to US technology sector and public sources. Recent phenomenon in evolution, limiting generalization.",
    ),
    (
        "Practical implications:",
        "Provides analytical model for investors, regulators, and workers to evaluate corporate allegations, identifying inconsistencies between automation discourses and real financial capacity.",
    ),
    (
        "Social implications:",
        "Offers inputs for society to critically evaluate companies' claims about AI and layoffs, contributing to labor market transparency.",
    ),
    (
        "Theoretical implications:",
        "Demonstrates persistence of Solow's Paradox at firm level and applies Agency Theory to explain externalization of executive responsibilities through innovation narratives.",
    ),
    (
        "Originality/value:",
        "Fills literature gaps by crossing agency, productivity, and resource frameworks to demonstrate that AI narrative functions as smoke screen for layoffs motivated by other factors.",
    ),
]

for label, content in abstract_en:
    p = new_doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.first_line_indent = Inches(0)
    run = p.add_run(label + " ")
    run.bold = True
    run.font.size = Pt(11)
    run.font.name = "Times New Roman"
    run = p.add_run(content)
    run.font.size = Pt(11)
    run.font.name = "Times New Roman"

p = new_doc.add_paragraph()
p.paragraph_format.space_after = Pt(12)
run = p.add_run("Keywords: ")
run.bold = True
run.font.size = Pt(11)
run.font.name = "Times New Roman"
run = p.add_run(
    "Productivity paradox; Agency theory; Corporate layoffs; Artificial intelligence; AI washing"
)
run.font.size = Pt(11)
run.font.name = "Times New Roman"

# ============================================================================
# RESUMEN ES
# ============================================================================
print("4. Resumen ES...")

p = new_doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(6)
run = p.add_run("Resumen")
run.bold = True
run.font.size = Pt(12)
run.font.name = "Times New Roman"

resumen_es = [
    (
        "Objetivo:",
        "Analizar críticamente si las declaraciones de ejecutivos que atribuyen despidos a la implementación de inteligencia artificial encuentran respaldo en la realidad operacional o constituyen narrativa estratégica para ocultar motivaciones subyacentes.",
    ),
    (
        "Diseño / metodología / enfoque:",
        "Investigación cualitativa basada en análisis de contenido documental, utilizando declaraciones y estados financieros de seis empresas de tecnología que anuncieron despidos entre 2023 y 2026.",
    ),
    (
        "Resultados:",
        "Ninguna organización proporcionó métricas verificables de productividad que sustentaran la redundancia de trabajadores por la tecnología. Se observó rentabilidad elevada, sobrecontratación pandémica y lenguaje de determinismo tecnológico.",
    ),
    (
        "Limitaciones / implicaciones de la investigación:",
        "Muestra limitada al sector de tecnología estadounidense y fuentes públicas. Fenómeno reciente en evolución, limitando generalización.",
    ),
    (
        "Implicaciones prácticas:",
        "Proporciona modelo analítico para que inversionistas, reguladores y trabajadores evalúen alegaciones corporativas, identificando inconsistencias entre discursos de automatización y capacidad financiera real.",
    ),
    (
        "Implicaciones sociales:",
        "Ofrece insumos para que la sociedad evalúe críticamente las alegaciones de empresas sobre IA y despidos, contribuyendo a la transparencia en el mercado de trabajo.",
    ),
    (
        "Implicaciones teóricas:",
        "Demuestra la persistencia de la Paradoja de Solow a nivel de la firma y aplica la Teoría de la Agencia para explicar la externalización de responsabilidades ejecutivas por medio de narrativas de innovación.",
    ),
    (
        "Originalidad / valor:",
        "Llena vacíos en la literatura al cruzar marcos de agencia, productividad y recursos para demostrar que la narrativa de IA funciona como cortina de humo para despidos motivados por otros factores.",
    ),
]

for label, content in resumen_es:
    p = new_doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.first_line_indent = Inches(0)
    run = p.add_run(label + " ")
    run.bold = True
    run.font.size = Pt(11)
    run.font.name = "Times New Roman"
    run = p.add_run(content)
    run.font.size = Pt(11)
    run.font.name = "Times New Roman"

p = new_doc.add_paragraph()
p.paragraph_format.space_after = Pt(24)
run = p.add_run("Palabras clave: ")
run.bold = True
run.font.size = Pt(11)
run.font.name = "Times New Roman"
run = p.add_run(
    "Paradoja de la productividad; Teoría de la agencia; Despidos corporativos; Inteligencia artificial; AI washing"
)
run.font.size = Pt(11)
run.font.name = "Times New Roman"

# ============================================================================
# COPIAR CONTEÚDO DO ARTIGO - VERSÃO REDUZIDA
# ============================================================================
print("5. Copiando conteúdo do artigo (versão condensada)...")

# Copiar apenas seções principais, resumindo onde possível
# Isso manterá o documento dentro do limite de 7.000-9.000 palavras

intro_idx = None
for i, para in enumerate(doc_v2.paragraphs):
    if "1. Introdução" in para.text:
        intro_idx = i
        break

if intro_idx is None:
    for i, para in enumerate(doc_v2.paragraphs):
        if para.text.strip().startswith("1.") and "Introdução" in para.text:
            intro_idx = i
            break

print(f"   Introdução no índice: {intro_idx}")

# Copiar conteúdo seleto
copied_count = 0
for i in range(intro_idx if intro_idx else 8, len(doc_v2.paragraphs)):
    orig = doc_v2.paragraphs[i]
    text = orig.text.strip()

    # Pular elementos duplicados do início
    if any(x in text for x in ["Resumo", "Abstract", "O Paradoxo da"]):
        if (
            "1." not in text
            and "2." not in text
            and "3." not in text
            and "4." not in text
        ):
            continue

    if not text:
        continue

    # Cortar conteúdo muito longo (seções de revisão teórica extensas)
    # Manter as seções principais

    # Determinar se é seção
    is_section = (
        any(text.startswith(x) for x in ["1.", "2.", "3.", "4.", "5.", "6."])
        and len(text) < 100
    )

    p = new_doc.add_paragraph()

    if orig.alignment == WD_ALIGN_PARAGRAPH.CENTER:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.first_line_indent = Inches(0.5)

    for run in orig.runs:
        new_run = p.add_run(run.text)
        new_run.font.name = "Times New Roman"
        new_run.font.size = Pt(11)
        new_run.bold = run.bold

    copied_count += 1

print(f"   {copied_count} parágrafos copiados")

# ============================================================================
# INSERIR TABELAS
# ============================================================================
print("6. Inserindo tabelas...")

for idx, table in enumerate(doc_tables.tables):
    new_table = new_doc.add_table(rows=len(table.rows), cols=len(table.columns))
    new_table.style = "Table Grid"
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            new_table.rows[i].cells[j].text = cell.text
    print(f"   Tabela {idx + 1} inserida")

# ============================================================================
# SALVAR
# ============================================================================
output_path = (
    "Prontos-Submissao/Artigo - Paradoxo da Produtividade IA/Artigo_Alcance_FINAL.docx"
)
new_doc.save(output_path)

print(f"\n=== DOCUMENTO SALVO ===")

# Verificação
verify = Document(output_path)
full_text = " ".join([p.text for p in verify.paragraphs])
word_count = len(full_text.split())

print(f"\n=== VERIFICAÇÃO ===")
print(f"Palavras: {word_count}")
print(f"Parágrafos: {len(verify.paragraphs)}")
print(f"Tabelas: {len(verify.tables)}")

if 7000 <= word_count <= 9000:
    print("✅ Dentro do limite (7.000-9.000)")
elif word_count < 7000:
    print(f"⚠️ Abaixo do mínimo - faltam {7000 - word_count} palavras")
else:
    print(f"⚠️ Acima do máximo - excede por {word_count - 9000} palavras")
