from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from copy import deepcopy
import os

BASE = "Prontos-Submissao/Artigo - Paradoxo da Produtividade IA"
SRC = os.path.join(BASE, "Artigo Revista Alcance.docx")
FINAL = os.path.join(BASE, "Artigo_Alcance_FINAL.docx")
ANON = os.path.join(BASE, "Manuscrito_Alcance_Anônimo.docx")

if not os.path.exists(SRC):
    raise FileNotFoundError(f'Fonte não encontrada: {SRC}')

src = Document(SRC)
new = Document()

# Margens de acordo com Revista Alcance
section = new.sections[0]
section.top_margin = Cm(3)
section.left_margin = Cm(3)
section.bottom_margin = Cm(2)
section.right_margin = Cm(2)

# Helper para criar parágrafo com style
def add_paragraph(text='', bold=False, center=False, italic=False):
    p = new.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)
    return p

# Add titles in 3 languages
title_port = 'O PARADOXO DA PRODUTIVIDADE DA IA: ANÁLISE CRÍTICA DO DISCURSO CORPORATIVO SOBRE DEMISSÕES TECNOLÓGICAS'
title_eng = 'THE AI PRODUCTIVITY PARADOX: CRITICAL ANALYSIS OF CORPORATE DISCOURSE ON TECHNOLOGICAL LAYOFFS'
title_esp = 'EL PARADOJA DE LA PRODUCTIVIDAD DE LA IA: ANÁLISIS CRÍTICO DEL DISCURSO CORPORATIVO SOBRE DESPIDOS TECNOLÓGICOS'
add_paragraph(title_port, bold=True, center=True)
add_paragraph(title_eng, bold=True, center=True)
add_paragraph(title_esp, bold=True, center=True)

# Portuguese structured abstract
add_paragraph('Resumo', bold=True, center=True)
port = [
    ('Objetivo:', 'Analisar criticamente se as declarações de executivos que atribuem demissões em massa à implementação de inteligência artificial encontram respaldo na realidade operacional pública ou se constituem narrativa estratégica para ocultar motivações gerenciais subjacentes.'),
    ('Design / metodologia / abordagem:', 'A pesquisa caracteriza-se como qualitativa e baseia-se em análise de conteúdo documental. O corpus foi composto por declarações e demonstrativos financeiros de seis empresas de tecnologia (Salesforce, Intuit, Dropbox, Block, Cisco e IBM) que anunciaram reduções de pessoal entre 2023 e 2026.'),
    ('Resultados:', 'Nenhuma organização disponibilizou publicamente métricas verificáveis de produtividade que sustentassem a redundância de trabalhadores pela tecnologia. Em contrapartida, observou-se lucratividade elevada, histórico de sobrecontratação pandêmica, linguagem de determinismo tecnológico e investimentos simultâneos na própria ferramenta. Dessa forma, caracterizou-se o argumento tecnológico como pretexto discursivo.'),
    ('Limitações / implicações da pesquisa:', 'A amostra foi limitada ao setor de tecnologia norte-americano e a fontes públicas. O fenômeno investigado é recente e ainda em evolução, o que limita a generalização dos achados e sugere cautela na interpretação dos resultados.'),
    ('Implicações práticas:', 'O trabalho fornece um modelo analítico para que investidores, reguladores e trabalhadores avaliem alegações corporativas sobre IA e demissões. Com esse propósito, possibilita-se a identificação de inconsistências entre discursos de automação e a capacidade financeira real das organizações.'),
    ('Implicações sociais:', 'O estudo oferece subsídios para que a sociedade avalie criticamente alegações de empresas sobre a relação entre implementação de IA e demissões, contribuindo para transparência no mercado de trabalho e proteção dos direitos dos trabalhadores.'),
    ('Implicações teóricas:', 'O estudo demonstra a persistência do Paradoxo de Solow no nível da firma. De forma complementar, aplica-se a Teoria da Agência para explicar a externalização de responsabilidades executivas por meio de narrativas de inovação tecnológica.'),
    ('Originalidade / valor:', 'O estudo preenche lacunas na literatura ao cruzar referenciais de agência, produtividade e recursos para demonstrar que a narrativa de IA funciona como cortina de fumaça para demissões motivadas por outros fatores, introduzindo o conceito de AI Scapegoating.'),
]
for label, text in port:
    p = new.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1
    run = p.add_run(label + ' ')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)
    run2 = p.add_run(text)
    run2.font.name = 'Times New Roman'
    run2.font.size = Pt(12)

add_paragraph('Palavras-chave: Paradoxo da produtividade; Teoria da agência; Demissões corporativas; Inteligência artificial; AI washing')

# English structured abstract
add_paragraph('Abstract', bold=True, center=True)
eng = [
    ('Purpose:', 'To critically analyze whether executive statements attributing mass layoffs to artificial intelligence implementation find support in public operational reality or constitute strategic narratives to conceal underlying managerial motivations.'),
    ('Design/methodology/approach:', 'This qualitative research was based on documentary content analysis. The corpus comprised statements and financial reports from six technology companies (Salesforce, Intuit, Dropbox, Block, Cisco, and IBM) that announced layoffs between 2023 and 2026.'),
    ('Findings:', 'No organization publicly provided verifiable productivity metrics supporting human redundancy through technology. In contrast, high profitability, pandemic-era over-hiring history, technological determinism language, and simultaneous investments in the same tool were observed. Thus, the technological argument was characterized as discursive pretext.'),
    ('Research limitations/implications:', 'The sample was limited to the US technology sector and public sources. The investigated phenomenon is recent and still evolving, which limits the generalization of findings and suggests caution in interpreting results.'),
    ('Practical implications:', 'The paper provides an analytical model for investors, regulators, and workers to evaluate corporate allegations about AI and layoffs. This enables the identification of inconsistencies between automation discourses and the organizations’ real financial capacity.'),
    ('Social implications:', 'The study offers inputs for society to critically evaluate companies’ claims about the relationship between AI implementation and layoffs, contributing to labor market transparency and worker rights protection.'),
    ('Theoretical implications:', 'The study demonstrates the persistence of Solow’s Paradox at the firm level. Complementarily, it applies Agency Theory to explain the externalization of executive responsibilities through technological innovation narratives.'),
    ('Originality/value:', 'This article fills literature gaps by crossing agency, productivity, and resource frameworks to demonstrate that the AI narrative functions as a smoke screen for layoffs motivated by other factors, introducing the concept of AI Scapegoating.'),
]
for label, text in eng:
    p = new.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1
    run = p.add_run(label + ' ')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)
    run2 = p.add_run(text)
    run2.font.name = 'Times New Roman'
    run2.font.size = Pt(12)

add_paragraph('Keywords: Productivity paradox; Agency theory; Corporate layoffs; Artificial intelligence; AI washing')

# Spanish structured abstract
add_paragraph('Resumen', bold=True, center=True)
esp = [
    ('Objetivo:', 'Analizar críticamente si las declaraciones de ejecutivos que atribuyen despidos en masa a la implementación de inteligencia artificial encuentran respaldo en la realidad operacional pública o constituyen narrativas estratégicas para ocultar motivaciones gerenciales subyacentes.'),
    ('Diseño / metodología / enfoque:', 'La investigación se caracteriza como cualitativa y se basa en análisis de contenido documental. El corpus incluyó declaraciones y estados financieros de seis empresas tecnológicas (Salesforce, Intuit, Dropbox, Block, Cisco e IBM) que anunciaron reducciones de personal entre 2023 y 2026.'),
    ('Resultados:', 'Ninguna organización publicó métricas de productividad verificables que sustentaran la redundancia de trabajadores por la tecnología. En cambio, se observó alta rentabilidad, historial de sobrecontratación pandémica, lenguaje de determinismo tecnológico e inversiones simultáneas en la misma herramienta. Así, el argumento tecnológico se caracterizó como un pretexto discursivo.'),
    ('Limitaciones / implicaciones de la investigación:', 'La muestra se limitó al sector tecnológico de Estados Unidos y a fuentes públicas. El fenómeno investigado es reciente y aún evoluciona, lo que limita la generalización de los resultados y sugiere cautela en la interpretación.'),
    ('Implicaciones prácticas:', 'El trabajo ofrece un modelo analítico para que inversionistas, reguladores y trabajadores evalúen las alegaciones corporativas sobre IA y despidos. Esto permite identificar inconsistencias entre los discursos de automatización y la capacidad financiera real de las organizaciones.'),
    ('Implicaciones sociales:', 'El estudio aporta elementos para que la sociedad evalúe críticamente las afirmaciones de las empresas sobre la relación entre la implementación de IA y los despidos, contribuyendo a la transparencia del mercado laboral y a la protección de los derechos de los trabajadores.'),
    ('Implicaciones teóricas:', 'El estudio demuestra la persistencia de la Paradoja de Solow a nivel de empresa. Complementariamente, aplica la Teoría de la Agencia para explicar la externalización de responsabilidades ejecutivas mediante narrativas de innovación tecnológica.'),
    ('Originalidad / valor:', 'El estudio llena vacíos en la literatura al cruzar marcos de agencia, productividad y recursos para demostrar que la narrativa de la IA funciona como cortina de humo para despidos motivados por otros factores, introduciendo el concepto de AI Scapegoating.'),
]
for label, text in esp:
    p = new.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1
    run = p.add_run(label + ' ')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)
    run2 = p.add_run(text)
    run2.font.name = 'Times New Roman'
    run2.font.size = Pt(12)

add_paragraph('Palabras clave: Paradoja de la productividad; Teoría de la agencia; Despidos corporativos; Inteligencia artificial; AI washing')

# Find introduction starting point in source
intro_idx = None
for i, p in enumerate(src.paragraphs):
    if p.text.strip().startswith('1. Introdução') or p.text.strip().lower() == 'introdução' or p.text.strip().lower() == 'introducao':
        intro_idx = i
        break
if intro_idx is None:
    raise ValueError('Não encontrei o início da Introdução no documento fonte.')

# Copy everything from introduction onwards preserving media and tables
children = list(src.element.body.iterchildren())
start_child = None
# map paragraph to its element
para_map = {p._p: i for i, p in enumerate(src.paragraphs)}
for child in children:
    if child.tag.endswith('p') and child in para_map:
        if para_map[child] == intro_idx:
            start_child = child
            break
if start_child is None:
    # fallback: locate by exact text match from introduction paragraph
    intro_text = src.paragraphs[intro_idx].text
    for child in children:
        if child.tag.endswith('p') and child.text == intro_text:
            start_child = child
            break
if start_child is None:
    raise ValueError('Falha ao mapear o elemento de introdução para a árvore XML.')

start_index = children.index(start_child)
for child in children[start_index:]:
    new.element.body.append(deepcopy(child))

# Normalize formatting for all paragraphs outside tables
for p in new.paragraphs:
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1
    p.paragraph_format.first_line_indent = Pt(0)
    for run in p.runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)

# Ensure no empty plain paragraphs lingering at the top
empty_paras = [p for p in new.paragraphs if not p.text.strip()]
for p in empty_paras:
    if p._p.getparent() is not None:
        p._element.getparent().remove(p._element)

new.save(FINAL)
print('Gravado', FINAL)

# Create anonymous copy if needed by clearing author metadata
anon = Document(FINAL)
anon.core_properties.author = ''
anon.save(ANON)
print('Gravado', ANON)
