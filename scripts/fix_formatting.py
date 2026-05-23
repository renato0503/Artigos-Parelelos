from docx import Document
from docx.shared import Pt, Inches
import os

BASE = "Prontos-Submissao/Artigo - Paradoxo da Produtividade IA"
FINAL = os.path.join(BASE, "Artigo_Alcance_FINAL.docx")
ANON = os.path.join(BASE, "Manuscrito_Alcance_Anônimo.docx")
COMAUTH = os.path.join(BASE, "Manuscrito_Alcance_ComAutores.docx")

ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

if not os.path.exists(FINAL):
    print('Arquivo final não encontrado:', FINAL)
    raise SystemExit(1)

print('Carregando documento final para ajuste de formatação...')
doc = Document(FINAL)

def para_in_table(par):
    # check if paragraph has a table ancestor
    try:
        ancestors = par._p.xpath('ancestor::w:tc', namespaces=ns)
        return len(ancestors) > 0
    except Exception:
        return False

# Adjust paragraph formatting: do NOT apply to paragraphs inside table cells
print('Ajustando parágrafos (não afetando parágrafos em tabelas)...')
for p in doc.paragraphs:
    if para_in_table(p):
        # ensure no extra spacing inside table cells
        try:
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.first_line_indent = Inches(0)
        except Exception:
            pass
    else:
        # normal body paragraphs: small spacing after, first-line indent
        try:
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.first_line_indent = Inches(0.5)
        except Exception:
            pass

# For each table, ensure there is spacing paragraphs around it (add small spacing if needed)
print('Ajustando espaçamento em torno de tabelas...')
body = doc.element.body
children = list(body.iterchildren())
from copy import deepcopy

# Iterate over children and when a table element found, ensure previous sibling is a paragraph with spacing
for idx, child in enumerate(children):
    if child.tag.endswith('tbl'):
        # check previous element
        prev = children[idx - 1] if idx - 1 >= 0 else None
        next_el = children[idx + 1] if idx + 1 < len(children) else None
        from docx.oxml import OxmlElement
        if prev is None or not prev.tag.endswith('p'):
            # insert a paragraph before
            p = OxmlElement('w:p')
            body.insert(idx, p)
        if next_el is None or not next_el.tag.endswith('p'):
            p2 = OxmlElement('w:p')
            body.insert(idx + 2, p2)

# After ensuring paragraphs exist, set their spacing
for p in doc.paragraphs:
    if not para_in_table(p):
        # if paragraph is empty and adjacent to a table, give it spacing
        txt = p.text.strip()
        if txt == '':
            try:
                p.paragraph_format.space_after = Pt(6)
            except Exception:
                pass

# Save adjusted final
backup = FINAL.replace('.docx', '_FORMATTED.docx')
doc.save(backup)
print('Documento final formatado salvo em:', backup)

# Create anonymous copy by removing author block between titles and 'Resumo' or 'Abstract'
print('Criando versão anônima a partir do documento formatado...')
anon = Document(backup)
stop_idx = None
for i, p in enumerate(anon.paragraphs[:60]):
    t = p.text.strip().lower()
    if t.startswith('resumo') or t.startswith('abstract'):
        stop_idx = i
        break
paras_to_remove = []
if stop_idx and stop_idx > 3:
    for i in range(3, stop_idx):
        paras_to_remove.append(anon.paragraphs[i])
for p in paras_to_remove:
    try:
        p._element.getparent().remove(p._element)
    except Exception:
        pass
# clear author metadata
anon.core_properties.author = ''
anon.save(ANON)
print('Versão anônima salva em:', ANON)

# Also update the copy with authors to ensure formatting
print('Atualizando cópia com autores...')
if os.path.exists(COMAUTH):
    cdoc = Document(COMAUTH)
    # minimal formatting fix: normalize paragraphs outside tables
    for p in cdoc.paragraphs:
        if para_in_table(p):
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.first_line_indent = Inches(0)
        else:
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.first_line_indent = Inches(0.5)
    cdoc.save(COMAUTH)
    print('Cópia com autores atualizada:', COMAUTH)
else:
    print('Cópia com autores não encontrada, pulando.')

print('\nConcluído.')
