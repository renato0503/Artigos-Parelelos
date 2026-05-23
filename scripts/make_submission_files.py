from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl import load_workbook
import os

BASE = "Prontos-Submissao/Artigo - Paradoxo da Produtividade IA"
FINAL = os.path.join(BASE, "Artigo_Alcance_FINAL.docx")
ANON = os.path.join(BASE, "Manuscrito_Alcance_Anônimo.docx")
ANON_PDF = os.path.join(BASE, "Manuscrito_Alcance_Anonimo.pdf")
COMAUTH = os.path.join(BASE, "Manuscrito_Alcance_ComAutores.docx")
SUPP = os.path.join(BASE, "Documento suplementar - revista alcance.docx")
LETTER = os.path.join(BASE, "Carta_Editor.docx")
TABLES_XLSX = os.path.join(BASE, "tabelas.xlsx")

if not os.path.exists(FINAL):
    print(f"Arquivo base não encontrado: {FINAL}")
    raise SystemExit(1)

print(f"Carregando {FINAL}...")
doc = Document(FINAL)

# --------------------------------------------------
# Extrair bloco de autores: assumimos que está entre o título e o cabeçalho 'Resumo' ou 'Resumo'
# Procurar índice do parágrafo que contém 'Resumo' ou 'Resumo'
stop_idx = None
for i, p in enumerate(doc.paragraphs[:40]):
    if p.text.strip().lower().startswith('resumo') or p.text.strip().lower().startswith('abstract'):
        stop_idx = i
        break

# Autoras/afiliações: entre o título (parágrafo 0..5) e stop_idx
author_block = []
if stop_idx and stop_idx > 3:
    for i in range(3, stop_idx):
        text = doc.paragraphs[i].text.strip()
        if text:
            author_block.append(text)

print('Bloco de autores extraído:')
for a in author_block:
    print(' -', a)

# --------------------------------------------------
# Criar cópia com autores (arquivo interno)
print('Criando cópia com autores...')
doc.save(COMAUTH)

# --------------------------------------------------
# Remover parágrafos do bloco de autores para versão anônima
print('Criando versão anônima...')
anon_doc = Document(FINAL)
# Collect elements to remove
paras_to_remove = []
for i, p in enumerate(anon_doc.paragraphs[:40]):
    if p.text.strip().lower().startswith('resumo') or p.text.strip().lower().startswith('abstract'):
        break
    # mark paragraphs between title and resumo
    if i >= 3:  # skip the three title paragraphs added by script
        paras_to_remove.append(p)

# remove parasite paragraphs
for p in paras_to_remove:
    p._element.getparent().remove(p._element)

# Clear core properties author
anon_doc.core_properties.author = ''

# Ajustar espaçamento entre parágrafos: remover espaços em branco entre parágrafos
from docx.shared import Pt, Inches
for p in anon_doc.paragraphs:
    try:
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.first_line_indent = Inches(0.5)
    except Exception:
        pass

# Inserir tabelas a partir do xlsx
if os.path.exists(TABLES_XLSX):
    # Only insert tables from xlsx if the final doc does not already contain tables
    if len(anon_doc.tables) == 0:
        print('Inserindo tabelas a partir de tabelas.xlsx...')
        wb = load_workbook(TABLES_XLSX, data_only=True)
        for sheetname in wb.sheetnames:
            ws = wb[sheetname]
            rows = list(ws.rows)
            if not rows:
                continue
            # create a heading for the table
            h = anon_doc.add_paragraph()
            h.add_run(f"Tabela: {sheetname}").bold = True
            # build table
            table = anon_doc.add_table(rows=len(rows), cols=len(rows[0]))
            table.style = 'Table Grid'
            for i, row in enumerate(rows):
                for j, cell in enumerate(row):
                    val = '' if cell.value is None else str(cell.value)
                    table.rows[i].cells[j].text = val
            anon_doc.add_paragraph()  # spacing
    else:
        print('O documento final já contém tabelas; pulando inserção de tabelas do xlsx para evitar duplicação.')
else:
    print('Arquivo de tabelas não encontrado; pulando.')

anon_doc.save(ANON)
print(f'Versão anônima salva em: {ANON}')

# --------------------------------------------------
# Gerar carta ao editor a partir do documento suplementar
if os.path.exists(SUPP):
    print('Criando carta ao editor a partir do documento suplementar...')
    supp = Document(SUPP)
    # Inserir dados dos autores no topo
    insert_text = '\n'.join(author_block) if author_block else 'Autores: (não detectados)'
    para = supp.add_paragraph()
    para.alignment = 0
    para.add_run('Autores e afiliações:').bold = True
    supp.add_paragraph(insert_text)
    supp.save(LETTER)
    print(f'Carta ao editor salva em: {LETTER}')
else:
    print('Documento suplementar não encontrado; criando carta básica...')
    new = Document()
    new.add_paragraph('À Editoria da Revista Alcance')
    new.add_paragraph('Prezados,')
    new.add_paragraph('Submetemos o manuscrito intitulado "O PARADOXO DA PRODUTIVIDADE DA IA..." para avaliação.')
    new.add_paragraph('\nAutores e afiliações:')
    for a in author_block:
        new.add_paragraph(a)
    new.save(LETTER)
    print(f'Carta ao editor criada em: {LETTER}')

print('\nConcluído.')
print('Arquivos gerados:')
for p in [COMAUTH, ANON, LETTER]:
    print(' -', p if os.path.exists(p) else f'{p} (não criado)')
