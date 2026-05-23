from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from copy import deepcopy

print("=== CRIANDO ARTIGO COMPLETO PARA ALCANCE ===\n")

# Carregar documento com estrutura e tabelas
doc_existing = Document(
    "Prontos-Submissao/Artigo - Paradoxo da Produtividade IA/Artigo Revista Alcance.docx"
)

# Carregar v2 para extrair conteúdo adicional
doc_v2 = Document(
    "Prontos-Submissao/Artigo - Paradoxo da Produtividade IA/paradoxo_produtividade_ia_v2.docx"
)

print(
    f"Documento existente: {len(doc_existing.paragraphs)} parágrafos, {len(doc_existing.tables)} tabelas"
)
print(f"V2: {len(doc_v2.paragraphs)} parágrafos")

# Criar novo documento
new_doc = Document()

# Margens Alcance
for section in new_doc.sections:
    section.top_margin = Cm(3)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(3)
    section.right_margin = Cm(2)

# ============================================================================
# TÍTULOS
# ============================================================================
print("1. Adicionando títulos...")

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
# RESUMO ESTRUTURADO PT (8 itens)
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
        "Analisar criticamente se as declarações de executivos que atribuem demissões em massa à implementação de inteligência artificial encontram respaldo na realidade operacional pública ou se constituem narrativa estratégica para ocultar motivações gerenciais subjacentes.",
    ),
    (
        "Design / metodologia / abordagem:",
        "A pesquisa caracteriza-se como qualitativa e baseia-se em análise de conteúdo documental. O corpus foi composto por declarações e demonstrativos financeiros de seis empresas de tecnologia (Salesforce, Intuit, Dropbox, Block, Cisco e IBM) que anunciaram reduções de pessoal entre 2023 e 2026.",
    ),
    (
        "Resultados:",
        "Nenhuma organização disponibilizou publicamente métricas verificáveis de produtividade que sustentassem a redundância de trabalhadores pela tecnologia. Em contrapartida, observou-se lucratividade elevada, histórico de sobrecontratação pandêmica, linguagem de determinismo tecnológico e investimentos simultâneos na própria ferramenta. Dessa forma, caracterizou-se o argumento tecnológico como pretexto discursivo.",
    ),
    (
        "Limitações / implicações da pesquisa:",
        "A amostra foi limitada ao setor de tecnologia norte-americano e a fontes públicas. O fenômeno investigado é recente e ainda em evolução, o que limita a generalização dos achados e sugere cautela na interpretação dos resultados.",
    ),
    (
        "Implicações práticas:",
        "O trabalho fornece um modelo analítico para que investidores, reguladores e trabalhadores avaliem alegações corporativas sobre IA e demissões. Com esse propósito, possibilita-se a identificação de inconsistências entre discursos de automação e a capacidade financeira real das organizações.",
    ),
    (
        "Implicações sociais:",
        "O estudo oferece subsídios para que a sociedade avalie criticamente alegações de empresas sobre a relação entre implementação de IA e demissões, contribuindo para transparência no mercado de trabalho e proteção dos direitos dos trabalhadores.",
    ),
    (
        "Implicações teóricas:",
        "O estudo demonstra a persistência do Paradoxo de Solow no nível da firma. De forma complementar, aplica-se a Teoria da Agência para explicar a externalização de responsabilidades executivas por meio de narrativas de inovação tecnológica.",
    ),
    (
        "Originalidade / valor:",
        "O estudo preenche lacunas na literatura ao cruzar referenciais de agência, produtividade e recursos para demonstrar que a narrativa de IA funciona como cortina de fumaça para demissões motivadas por outros fatores, introduzindo o conceito de AI Scapegoating.",
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
# ABSTRACT EN (8 itens)
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
        "To critically analyze whether executive statements attributing mass layoffs to artificial intelligence implementation find support in public operational reality or constitute strategic narratives to conceal underlying managerial motivations.",
    ),
    (
        "Design/methodology/approach:",
        "This qualitative research was based on documentary content analysis. The corpus comprised statements and financial reports from six technology companies (Salesforce, Intuit, Dropbox, Block, Cisco, and IBM) that announced layoffs between 2023 and 2026.",
    ),
    (
        "Findings:",
        "No organization publicly provided verifiable productivity metrics supporting human redundancy through technology. In contrast, high profitability, pandemic-era over-hiring history, technological determinism language, and simultaneous investments in the same tool were observed. Thus, the technological argument was characterized as discursive pretext.",
    ),
    (
        "Research limitations/implications:",
        "The sample was limited to the US technology sector and public sources. The investigated phenomenon is recent and still evolving, which limits the generalization of findings and suggests caution in interpreting results.",
    ),
    (
        "Practical implications:",
        "The paper provides an analytical model for investors, regulators, and workers to evaluate corporate allegations about AI and layoffs. This enables the identification of inconsistencies between automation discourses and the organizations' real financial capacity.",
    ),
    (
        "Social implications:",
        "The study offers inputs for society to critically evaluate companies' claims about the relationship between AI implementation and layoffs, contributing to labor market transparency and worker rights protection.",
    ),
    (
        "Theoretical implications:",
        "The study demonstrates the persistence of Solow's Paradox at the firm level. Complementarily, it applies Agency Theory to explain the externalization of executive responsibilities through technological innovation narratives.",
    ),
    (
        "Originality/value:",
        "This article fills literature gaps by crossing agency, productivity, and resource frameworks to demonstrate that the AI narrative functions as a smoke screen for layoffs motivated by other factors, introducing the concept of AI Scapegoating.",
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
# RESUMEN ES (8 itens)
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
        "Analizar críticamente si las declaraciones de ejecutivos que atribuyen despidos en masa a la implementación de inteligencia artificial encuentran respaldo en la realidad operacional pública o constituyen narrativa estratégica para ocultar motivaciones gerenciales subyacentes.",
    ),
    (
        "Diseño / metodología / enfoque:",
        "La investigación se caracteriza como cualitativa y se basa en análisis de contenido documental. El corpus se componía de declaraciones y estados financieros de seis empresas de tecnología (Salesforce, Intuit, Dropbox, Block, Cisco e IBM) que anuncieron reducciones de personal entre 2023 y 2026.",
    ),
    (
        "Resultados:",
        "Ninguna organización disponibilizó públicamente métricas verificables de productividad que sustentaran la redundancia de trabajadores por la tecnología. En cambio, se observó rentabilidad elevada, historial de sobrecontratación pandémica, lenguaje de determinismo tecnológico e inversiones simultáneas en la propia herramienta. De esta forma, se caracterizó el argumento tecnológico como pretexto discursivo.",
    ),
    (
        "Limitaciones / implicaciones de la investigación:",
        "La muestra se limitó al sector de tecnología estadounidense y a fuentes públicas. El fenómeno investigado es reciente y aún en evolução, lo que limita la generalización de los hallazgos y sugiere cautela en la interpretación de los resultados.",
    ),
    (
        "Implicaciones prácticas:",
        "El trabajo proporciona un modelo analítico para que inversionistas, reguladores y trabajadores evalúen alegaciones corporativas sobre IA y despidos. Con ese propósito, se possibilita la identificación de inconsistencias entre discursos de automatización y la capacidad financiera real de las organizaciones.",
    ),
    (
        "Implicaciones sociales:",
        "El estudio ofrece insumos para que la sociedad evalúe críticamente las alegaciones de empresas sobre la relación entre implementación de IA y despidos, contribuyendo a la transparencia en el mercado de trabajo y la protección de los derechos de los trabajadores.",
    ),
    (
        "Implicaciones teóricas:",
        "El estudio demuestra la persistencia de la Paradoja de Solow a nivel de la firma. De forma complementaria, se aplica la Teoría de la Agencia para explicar la externalización de responsabilidades ejecutivas por medio de narrativas de innovación tecnológica.",
    ),
    (
        "Originalidad / valor:",
        "El estudio llena vacíos en la literatura al cruzar marcos de agencia, productividad y recursos para demostrar que la narrativa de IA funciona como cortina de humo para despidos motivados por otros factores, introduciendo el concepto de AI Scapegoating.",
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
# COPIAR CONTEÚDO DO ARTIGO (EXISTENTE)
# ============================================================================
print("5. Copiando conteúdo do artigo existente...")

# Copiar parágrafos da Introdução em diante
intro_idx = None
for i, para in enumerate(doc_existing.paragraphs):
    if "1. Introdução" in para.text:
        intro_idx = i
        break

print(f"   Introdução no índice: {intro_idx}")

copied = 0
for i in range(intro_idx if intro_idx else 0, len(doc_existing.paragraphs)):
    orig = doc_existing.paragraphs[i]
    text = orig.text.strip()

    p = new_doc.add_paragraph()

    if orig.alignment == WD_ALIGN_PARAGRAPH.CENTER:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Inches(0)
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Inches(0.5)

    p.paragraph_format.space_after = Pt(6)

    for run in orig.runs:
        new_run = p.add_run(run.text)
        new_run.font.name = "Times New Roman"
        new_run.font.size = Pt(11)
        new_run.bold = run.bold

    copied += 1

print(f"   {copied} parágrafos copiados")

# ============================================================================
# INSERIR TABELAS
# ============================================================================
print("6. Inserindo tabelas...")

for idx, table in enumerate(doc_existing.tables):
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
    print("✅ DENTRO DO LIMITE (7.000-9.000)")
elif word_count < 7000:
    print(f"⚠️ ABAIXO DO MÍNIMO - faltam {7000 - word_count} palavras")
else:
    print(f"⚠️ ACIMA DO MÁXIMO - excede por {word_count - 9000} palavras")
