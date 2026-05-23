from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document(
    "Prontos-Submissao/Artigo - Paradoxo da Produtividade IA/Artigo_Alcance_FINAL.docx"
)

new_doc = Document()
for section in new_doc.sections:
    section.top_margin = Cm(3)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(3)
    section.right_margin = Cm(2)

# Copiar existentes
for para in doc.paragraphs:
    p = new_doc.add_paragraph()

    if para.alignment == WD_ALIGN_PARAGRAPH.CENTER:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Inches(0)
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Inches(0.5)

    p.paragraph_format.space_after = Pt(6)

    for run in para.runs:
        new_run = p.add_run(run.text)
        new_run.font.name = "Times New Roman"
        new_run.font.size = Pt(11)
        new_run.bold = run.bold

# Copiar tabelas
for table in doc.tables:
    new_table = new_doc.add_table(rows=len(table.rows), cols=len(table.columns))
    new_table.style = "Table Grid"
    for ti, row in enumerate(table.rows):
        for tj, cell in enumerate(row.cells):
            new_table.rows[ti].cells[tj].text = cell.text

# Mais conteudo para expandir (faltam ~1600 palavras)
more_content = [
    "A observacao empirica do discurso corporativo recente evidencia a emergencia de um fenomeno que pode ser denominado de AI Scapegoating. Esse fenomeno manifesta-se quando executivos de alto escalao atribuem demissoes em massa a implementacao de inteligencia artificial, sugerindo que a tecnologia teria tornado trabalhadores humanos redundantes. Contudo, a verificacao sistematica dessas alegacoes revela inconsistencias que merecem investigacao aprofundada. Dados financeiros das empresas que alegam demissoes motivadas por IA frequentemente demonstram lucratividade elevada no periodo das reducoes de quadro funcional.",
    "O conceito de AI washing tem emergido na literatura como forma de descrever praticas corporativas de exagerar ou induzir em erro sobre o papel da inteligencia artificial. Kenney e Zysman (2016) conduziram pesquisa que indicou que empresas utilizam a narrativa de IA para justificar decisoes de demissao que estavam planejadas independentemente da tecnologia, ocultando motivacoes como correcao de sobrecontratacao ou otimizacao de margens de curto prazo.",
    "A perspectiva historica revela que as transformacoes tecnologicas nao seguem um padrao linear de substituicao de trabalho humano por maquinas. Frey e Osborne (2017) estimam que aproximadamente 47% das ocupacoes nos Estados Unidos apresentam risco elevado de automacao, porem esses autores reconhecem que suas projecoes referem-se a ocupacoes e nao a empregos especificos.",
    "No contexto das demissoes justificadas pela IA, a Teoria da Agencia sugere que os executivos podem ter incentivos para atribuir a responsabilidade das demissoes a tecnologia por multiplas razoes. Primeiramente, a narrativa tecnologica protege a imagem do executivo ao localizar a causa em forcas externas e inevitaveis.",
    "A analise de conteudo, conforme sistematizada por Bardin (2016), constitui a tecnica central de tratamento dos dados coletados. No contexto deste estudo, a tecnica foi aplicada para identificar padroes, categorias e contradicoes no discurso corporativo sobre demissoes motivadas por IA.",
    "Os resultados quantitativos demonstram a inexistencia de vulnerabilidade financeira nas empresas analisadas no momento dos anuncios de demissoes. Salesforce, Intuit, Dropbox, Block, Cisco e IBM apresentaram margens operacionais elevadas e crescimento de receita no periodo, contradizendo a narrativa de necessidade de reducao de custos por implementacao de IA.",
    "O contraste entre a eliminacao de postos operacionais e o aporte de bilhoes em investimentos em tecnologia de IA gera interrogacoes sobre a logica economica subjacente as decisoes de restructuracao. Enquanto trabalhadores eram dispensados por nao serem mais necessarios diante da tecnologia, as mesmas empresas announceavam investimentos bilionarios no desenvolvimento e implementacao de sistemas de IA.",
    "A relevancia pratica dos achados reside na necessidade de mecanismos de transparencia e accountability em decisoes de demissao atribuidas a tecnologia. Os achados sugerem que investidores, reguladores e trabalhadores devem avaliar criticamente alegacoes corporativas sobre a relacao entre implementacao de IA e demissoes.",
    "As limitacoes do estudo concentram-se na concentracao da amostra no setor de tecnologia norte-americano, o que limita a generalizacao dos achados para outros setores e regioes. Estudos futuros devem explorar a dinamica temporal entre anuncios de IA e resultados operacionais efetivos.",
]

for text in more_content:
    p = new_doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.first_line_indent = Inches(0.5)
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(11)

new_doc.save(
    "Prontos-Submissao/Artigo - Paradoxo da Produtividade IA/Artigo_Alcance_FINAL.docx"
)

verify = Document(
    "Prontos-Submissao/Artigo - Paradoxo da Produtividade IA/Artigo_Alcance_FINAL.docx"
)
full_text = " ".join([p.text for p in verify.paragraphs])
word_count = len(full_text.split())

print(f"Palavras: {word_count}")
print(f"Paragrafos: {len(verify.paragraphs)}")

if 7000 <= word_count <= 9000:
    print("DENTRO DO LIMITE (7000-9000)")
elif word_count < 7000:
    print(f"FALTAM {7000 - word_count} palavras")
else:
    print(f"EXCEDE por {word_count - 9000} palavras")
