from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from copy import deepcopy
import os

BASE = "Prontos-Submissao/Artigo - Paradoxo da Produtividade IA"
SRC = os.path.join(BASE, "Artigo Revista Alcance.docx")
FINAL = os.path.join(BASE, "Artigo_Alcance_FINAL.docx")

print('Regenerating final document with media and preserving element order...')

if not os.path.exists(SRC):
    print(f'Fonte não encontrada: {SRC}')
    raise SystemExit(1)

src = Document(SRC)
new = Document()

# set margins
for section in new.sections:
    section.top_margin = Cm(3)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(3)
    section.right_margin = Cm(2)

# Add multilingual titles (keep as in create_final_v3)
titles = [
    "O PARADOXO DA PRODUTIVIDADE DA IA: ANÁLISE CRÍTICA DO DISCURSO CORPORATIVO SOBRE DEMISSÕES TECNOLÓGICAS",
    "THE AI PRODUCTIVITY PARADOX: CRITICAL ANALYSIS OF CORPORATE DISCOURSE ON TECHNOLOGICAL LAYOFFS",
    "EL PARADOJA DE LA PRODUCTIVIDAD DE LA IA: ANÁLISIS CRÍTICO DEL DISCURSO CORPORATIVO SOBRE DESPIDOS TECNOLÓGICOS",
]
for t in titles:
    p = new.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(t)
    run.bold = True
    run.font.size = Pt(14)
    run.font.name = "Times New Roman"

# Add placeholder for author block (will be removed later for anonymous)
# We copy the author block from source if present (paragraphs between title and 'Resumo')
stop_idx = None
for i, p in enumerate(src.paragraphs[:40]):
    t = p.text.strip().lower()
    if t.startswith('resumo') or t.startswith('abstract'):
        stop_idx = i
        break

if stop_idx and stop_idx > 3:
    for i in range(3, stop_idx):
        text = src.paragraphs[i].text.strip()
        if text:
            p = new.add_paragraph(text)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)

# Now add the summaries (try to copy from our helper if present in scripts, else skip)
# For simplicity, reuse the create_final_v3 content if needed — but here we will copy source summaries if present
# Copy entire body elements from 'Introdução' onwards preserving tables and pictures
body = src.element.body
children = list(body.iterchildren())

# find the child index corresponding to the paragraph that starts with '1. Introdução'
start_child_idx = 0
found = False
for idx, child in enumerate(children):
    # paragraph tag endswith 'p'
    if child.tag.endswith('p'):
        # get text
        # get the text by building a temporary paragraph
        try:
            p = src.paragraphs[0]
        except Exception:
            p = None
        # approach: match by searching through paragraphs and mapping to _p
        pass

# Map paragraph objects to their _p elements for quick lookup
p_to_idx = {}
for i, p in enumerate(src.paragraphs):
    p_to_idx[p._p] = i

# find paragraph index in src.paragraphs containing '1. Introdução'
intro_para_idx = None
for i, p in enumerate(src.paragraphs):
    if '1. Introdução' in p.text:
        intro_para_idx = i
        break

if intro_para_idx is None:
    # fallback: find 'Introdução' alone
    for i, p in enumerate(src.paragraphs):
        if p.text.strip().lower().startswith('introdução') or p.text.strip().lower().startswith('introducao'):
            intro_para_idx = i
            break

# Now find the corresponding child index in body
start_child_idx = 0
if intro_para_idx is not None:
    target_p = src.paragraphs[intro_para_idx]._p
    for idx, child in enumerate(children):
        if child is target_p:
            start_child_idx = idx
            found = True
            break

# If not found, start from the beginning of children
if not found:
    start_child_idx = 0

# Append all children from start_child_idx to the new document body
for child in children[start_child_idx:]:
    new.element.body.append(deepcopy(child))

# Post-process: remove empty paragraphs and collapse spaces
for p in new.paragraphs:
    # remove trailing/leading whitespace
    if p.text.strip() == '':
        try:
            p._element.getparent().remove(p._element)
        except Exception:
            pass
    else:
        # remove extra space after
        try:
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.left_indent = Cm(0)
            p.paragraph_format.first_line_indent = Inches(0.5)
        except Exception:
            pass

# Save final
new.save(FINAL)
print('Documento regenerado e salvo em:', FINAL)
