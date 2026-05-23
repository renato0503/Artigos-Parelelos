from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

print("=== EXPANDINDO ARTIGO PARA ALCANCE ===\n")

# Carregar documento final atual
doc = Document(
    "Prontos-Submissao/Artigo - Paradoxo da Produtividade IA/Artigo_Alcance_FINAL.docx"
)

# Carregar v2 para extrair conteúdo adicional
doc_v2 = Document(
    "Prontos-Submissao/Artigo - Paradoxo da Produtividade IA/paradoxo_produtividade_ia_v2.docx"
)

# Criar novo documento expandido
new_doc = Document()

# Margens Alcance
for section in new_doc.sections:
    section.top_margin = Cm(3)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(3)
    section.right_margin = Cm(2)

# Copiar todos os parágrafos existentes primeiro
for para in doc.paragraphs:
    text = para.text.strip()

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
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            new_table.rows[i].cells[j].text = cell.text

# Agora adicionar conteúdo expandido do v2
print("Adicionando conteúdo expandido do v2...")

# Encontrar onde inserir conteúdo adicional
# Procurar índice da seção de Referencial Teórico no novo doc
ref_idx = None
for i, para in enumerate(new_doc.paragraphs):
    if "2. Referencial Teórico" in para.text:
        ref_idx = i
        break

print(f"Referencial Teórico encontrado no índice: {ref_idx}")

# Conteúdo adicional a ser inserido (expandir seções)
# Isso deve ser feito antes da seção 4 (Resultados)

# Buscar conteúdo adicional no v2
v2_paras = doc_v2.paragraphs

# Encontrar início do referencial teórico no v2
v2_ref_idx = None
for i, para in enumerate(v2_paras):
    if "2." in para.text and "Referencial" in para.text:
        v2_ref_idx = i
        break

print(f"Referencial Teórico v2 no índice: {v2_ref_idx}")

# Inserir conteúdo expandido antes da seção 4
# Parágrafos extras para adicionar (~2000 palavras)

# Texto adicional para Referencial Teórico
expand_refencial = """

A relação entre inovação tecnológica e mercado de trabalho constitui tema central na teoria econômica desde os primórdios da industrialização. A Revolução Industrial, iniciada na segunda metade do século XVIII, representou a primeira transformação radical nos processos produtivos, substituindo o trabalho artesanal por máquinas e fábricas. Esse período demonstrou que a introdução de tecnologias disruptivas não elimina necessariamente o trabalho humano, mas transforma fundamentalmente sua natureza e organização.

A Terceira Revolução Industrial, também denominada Revolução da Informação, emergiu na segunda metade do século XX com o desenvolvimento da computação digital e das tecnologias de comunicação. A informatização dos processos produtivos transformou profundamente a organização do trabalho, introduzindo sistemas computadorizados de gestão e controle de qualidade. Brynjolfsson e McAfee (2014) argumentam que essa revolução tecnológica diferiu das anteriores em velocidade e escala, uma vez que as tecnologias digitais apresentam capacidade de difusão exponencial e custo decrescente.

O período contemporâneo tem sido marcado pela Quarta Revolução Industrial, caracterizada pela convergência de tecnologias físicas, digitais e biológicas. A inteligência artificial representa o elemento central dessa nova fase de transformação, distinguindo-se das tecnologias anteriores por sua capacidade de aprendizado e adaptação. Autor (2015) analisa que as tecnologias de IA apresentam potencial para executar tarefas que antes eram consideradas prerrogativa exclusiva da inteligência humana, incluindo reconhecimento de linguagem, tomada de decisão em contextos complexos e geração de conteúdo criativo.

A perspectiva histórica revela que as transformações tecnológicas não seguem um padrão linear de substituição de trabalho humano por máquinas. Frey e Osborne (2017) estimam que aproximadamente 47% das ocupações nos Estados Unidos apresentam risco elevado de automação, porém esses autores reconhecem que suas projeções referem-se a ocupações e não a empregos específicos, uma vez que cada ocupação engloba múltiplas tarefas com diferentes níveis de automatizáveis.

A aplicação do paradoxo de Solow ao contexto da inteligência artificial tem sido objeto de investigação recente. Bäck et al. (2025) conduziram estudo abrangente sobre a relação entre adoção de IA e produtividade em nível de firma, utilizando dados de empresas europeias. Os resultados indicaram que a adoção de IA não está correlacionada de forma consistente com ganhos de produtividade medidos, sugerindo que os benefícios da tecnologia podem ser capturados de forma desproporcional por empresas líderes ou que os custos de implementação ainda superam os retornos no curto prazo.

No contexto das demissões justificadas pela IA, a Teoria da Agência sugere que os executivos podem ter incentivos para atribuir a responsabilidade das demissões à tecnologia por múltiplas razões. Primeiramente, a narrativa tecnológica protege a imagem do executivo ao localizar a causa em forças externas e inevitáveis. Em segundo lugar, a atribuição à IA justifica as demissões como parte de uma estratégia de modernização. Em terceiro lugar, a narrativa de automação pode obscurecer o fato de que as demissões foram motivadas por erros de gestão.

Kaplan e Minton (2012) investigaram a relação entre desempenho financeiro e rotatividade de CEOs, demonstrando que a pressão por resultados cria incentivos para que gestores adotem medidas dramáticas de redução de custos. Os autores encontraram que a rotatividade de CEOs aumenta significativamente após períodos de baixo desempenho financeiro, criando um ciclo no qual gestores recém-nomeados enfrentam pressão intensa para demonstrar resultados no curto prazo.

A Visão Baseada em Recursos (RBV) constitui teoria estratégica que postula que as vantagens competitivas sustentáveis das empresas derivam de recursos internos específicos. Barney (1991) sistematizou e formalizou a RBV, introduzindo o framework VRIN para análise de recursos estratégicos. Segundo esse framework, recursos valiosos são aqueles que permitem à empresa implementar estratégias que melhoram sua eficiência ou eficácia. Barney e Clark (2007) aprofundaram a análise dos recursos intangíveis, destacando que o capital humano organizacional representa categoria particularmente importante de recursos estratégicos.

A teoria da destruição criadora, desenvolvida por Schumpeter (1942), oferece perspectiva fundamental para compreender a relação entre inovação tecnológica e emprego. Schumpeter propôs que o desenvolvimento econômico é caracterizado por ondas de destruição criadora, nas quais novas tecnologias e organizações tornam obsoletas as estruturas econômicas anteriores. Nesse processo, setores inteiros podem entrar em declínio, gerando desemprego temporário, enquanto novos setores emergem, criando oportunidades de emprego em atividades que não existiam anteriormente.

A teoria do capital humano, desenvolvida por Becker (1964), oferece perspectiva complementar ao destacar que o valor dos trabalhadores deriva de suas competências acumuladas. Becker demonstrou que investimentos em educação e treinamento aumentam a produtividade dos trabalhadores e, consequentemente, seus salários. Essa teoria indica que a tecnologia não torna os trabalhadores obsoletos, mas transforma as habilidades demandadas pelo mercado de trabalho.
""".strip().split("\n\n")

# Inserir conteúdo expandido antes da seção 4 (Resultados)
# Precisa encontrar a posição correta

resultados_idx = None
for i, para in enumerate(new_doc.paragraphs):
    if "4. Resultados" in para.text or "4.1 Apresentação" in para.text:
        resultados_idx = i
        break

print(f"Seção de Resultados encontrada no índice: {resultados_idx}")

if resultados_idx:
    # Inserir parágrafos adicionais antes
    for j, text in enumerate(expand_refcional):
        p = new_doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.first_line_indent = Inches(0.5)
        run = p.add_run(text.strip())
        run.font.name = "Times New Roman"
        run.font.size = Pt(11)

    print(f"   {len(expand_refcional)} parágrafos de expansão inseridos")

# Adicionar mais conteúdo na introdução
introducao_idx = None
for i, para in enumerate(new_doc.paragraphs):
    if "1. Introdução" in para.text:
        introducao_idx = i
        break

expand_intro = """
A observação empírica do discurso corporativo revela um fenômeno que pode ser denominado de "AI Scapegoating" ou "Paradoxo da Produtividade da IA". Esse fenômeno manifesta-se quando executivos de alto escalão atribuem demissões em massa à implementação de inteligência artificial, sugerindo que a tecnologia teria tornado trabalhadores humanos redundantes. Contudo, a verificação sistemática dessas alegações revela inconsistências que merecem investigação aprofundada. Dados financeiros das empresas que alegam demissões motivadas por IA frequentemente demonstram lucratividade elevada no período das reduções de quadro funcional, o que sugere que a narrativa tecnológica pode estar ocultando motivações distintas.

A relevância dessa investigação reside no fato de que a aceitação acrítica da narrativa de automação pode obscurecer decisões de gestão de quadro funcional motivadas por outros fatores, como correção de sobrecontratação pandêmica, pressões por margem de lucro ou reestruturações estratégicas não relacionadas à tecnologia. O objetivo geral desta pesquisa consiste em analisar criticamente se as declarações de executivos que atribuem demissões à inteligência artificial encontram respaldo na realidade operacional ou se constituem narrativa estratégica para ocultar motivações subjacentes. Os objetivos específicos incluem: compilar e sistematizar declarações de CEOs de empresas que utilizaram a IA como justificativa para demissões; confrontar essas declarações com dados financeiros e operacionais das respectivas empresas; e identificar padrões que sugiram a utilização da narrativa de IA como cortina de fumaça para decisões motivadas por outros fatores.
""".strip().split("\n\n")

# Inserir após introdução
if introducao_idx:
    for text in expand_intro:
        p = new_doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.first_line_indent = Inches(0.5)
        run = p.add_run(text.strip())
        run.font.name = "Times New Roman"
        run.font.size = Pt(11)

    print(f"   {len(expand_intro)} parágrafos de expansão da introdução inseridos")

# Adicionar conteúdo extra na metodologia
metodologia_idx = None
for i, para in enumerate(new_doc.paragraphs):
    if "3. Procedimentos Metodológicos" in para.text or "3.1" in para.text:
        metodologia_idx = i
        break

expand_metodo = """
O presente estudo adota uma abordagem qualitativa de investigação, posicionando-se no campo da pesquisa exploratória e descritiva, com fundamentação epistemológica interpretativista-crítica. Essa perspectiva epistemológica reconhece que a realidade social é construída através de significados e interpretações dos atores sociais, e que o discurso corporativo constitui prática social que reflete e simultaneamente molda as relações de poder nas organizações. O posicionamento interpretativista-crítico orienta a análise para além da descrição literal das declarações, buscando identificar as estruturas ideológicas e os interesses que permeiam o discurso dos executivos.

A pesquisa qualitativa foi escolhida como abordagem central por ser mais adequada para investigações que buscam compreender fenômenos complexos em seus contextos naturais, conforme argumentam Creswell e Creswell (2018). Os autores destacam que a pesquisa qualitativa permite explorar fenômeno em profundidade, capturando significados e interpretações que não são acessíveis através de abordagens quantitativas. No caso em análise, o discurso corporativo sobre demissões constitui fenômeno que requer interpretação contextualizada.

O procedimento técnico adotado fundamenta-se na pesquisa documental, complementada pela análise de conteúdo. Vergara (2016) define pesquisa documental como aquela que utiliza materiais que não receberam tratamento analítico, como documentos institucionais, relatórios e comunicações. No presente estudo, as declarações de CEOs, transcrições de earnings calls e comunicados corporativos constituem documentos que foram submetidos à análise sistemática. A escolha por esse procedimento justifica-se pela necessidade de acessar fontes primárias do discurso corporativo.

A análise de conteúdo, conforme sistematizada por Bardin (2016), constitui a técnica central de tratamento dos dados coletados. A autora define análise de conteúdo como conjunto de técnicas de comunicação sistemática e objetiva que permite a descrição do conteúdo das mensagens. No contexto deste estudo, a técnica foi aplicada para identificar padrões, categorias e contradições no discurso corporativo sobre demissões motivadas por IA.

O corpus de análise é composto por declarações públicas de CEOs, transcrições de earnings calls, comunicados corporativos e reportagens especializadas em veículos de negócios e tecnologia. Essa composição do corpus justifica-se pela necessidade de triangulação de fontes, a qual fortalece a validade dos achados conforme orienta Yin (2018) para estudos de caso. As declarações diretas dos executivos foram complementadas por reportagens que contextualizaram os anúncios e por dados financeiros disponíveis em demonstrativos oficiais das empresas.

O período de coleta abrange os anos de 2023 a 2026, coincidindo com a intensificação da adoção de inteligência artificial generativa no ambiente corporativo e com o aumento significativo de anúncios de demissões no setor de tecnologia. A delimitação temporal justifica-se pela natureza do fenômeno estudado, o qual é contemporâneo e ainda em evolução. A escolha por esse período permite capturar a diversidade de situações e contextos nos quais as empresas utilizaram a narrativa de IA para justificar demissões.

Os dados foram coletados a partir de múltiplas fontes, classificadas em primárias e secundárias. As fontes primárias consistem em transcrições oficiais de earnings calls, disponíveis em plataformas especializadas como Seeking Alpha, The Motley Fool e nos próprios websites de relações com investidores das empresas analisadas. As fontes secundárias compreendem reportagens em veículos especializados como Bloomberg, Reuters, TechCrunch, The Verge, Business Insider, Folha de S.Paulo e Valor Econômico, as quais complementam as informações primárias com contextualização e análise.
""".strip().split("\n\n")

if metodologia_idx:
    for text in expand_metodo:
        p = new_doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.first_line_indent = Inches(0.5)
        run = p.add_run(text.strip())
        run.font.name = "Times New Roman"
        run.font.size = Pt(11)

    print(f"   {len(expand_metodo)} parágrafos de expansão da metodologia inseridos")

# Expandir conclusão
conclusao_idx = None
for i, para in enumerate(new_doc.paragraphs):
    if "5. Conclusão" in para.text or "5.1" in para.text:
        conclusao_idx = i
        break

expand_conc = """
Os resultados evidenciaram que nenhuma das seis empresas analisadas apresentou métricas concretas de produtividade que sustentassem a narrativa de que a IA tornou trabalhadores redundantes. Dessa forma, confirma-se a persistência do paradoxo de Solow (1987) no contexto contemporâneo da inteligência artificial, corroborando os achados de Abel et al. (2022). A ausência de dados operacionais concretos representa achado fundamental que enfraquece a premissa básica das demissões justificadas por IA.

Os achados demonstraram que os incentivos de agência, conforme teorizado por Jensen e Meckling (1976), explicam o comportamento dos executivos ao utilizarem a narrativa de IA como mecanismo de externalização de responsabilidade. Quatro das seis empresas analisadas empregaram linguagem de determinismo tecnológico, apresentando a IA como força externa e inevitável que determinaria as demissões, protegendo assim a imagem do agente perante acionistas e mercado.

A análise sob a ótica da Visão Baseada em Recursos revelou contradição entre a decisão de promover demissões de trabalhadores e a preservação de recursos estratégicos valiosos, raros, inimitáveis e não-substituíveis. Barney (1991) argumenta que recursos intangíveis como conhecimento tácito representam fontes de vantagem competitiva sustentável, e a eliminação desses recursos sem consideração das sinergias entre competências humanas e artificiais indica perspectiva de curto prazo.

O estudo evidenciou que o discurso corporativo apresenta características do fenômeno denominado AI Washing, análogo ao greenwashing no contexto ambiental. A narrativa serve dupla função: justifica demissões de forma mais palatável para o público e reguladores, e posiciona a empresa como referência em tecnologia perante investidores. A atribuição de demissões à IA funciona como estratégia de legitimação organizacional em consonância com o que propõem Kenney e Zysman (2016).

Os dados de headcount revelaram que a maioria das empresas analisadas sobrecontratou durante o período pandêmico. A Salesforce cresceu aproximadamente 46% entre 2019 e 2023, a Block aproximadamente 63%, e a Cisco aproximadamente 21%. Após as demissões, as empresas mantinham quadro funcional superior ao período pré-pandemia, indicando que as demissões representam correção de sobrecontratação mais do que resposta a implementação de IA. Esse achado dialoga com a teoria de destruição criadora de Schumpeter (1942) e com o modelo task-based de Autor et al. (2003).

Diante do exposto, conclui-se que as evidências analisadas indicam que o argumento da inteligência artificial está sendo utilizado, predominantemente, como cortina de fumaça para ocultar motivações subjacentes às demissões, tais como a correção de sobrecontratação durante o período pandêmico, as pressões por margens de lucro de curto prazo, os incentivos de agência dos executivos que buscam proteger sua imagem perante acionistas, e o posicionamento estratégico perante investidores que valorizam narrativas de eficiência tecnológica.

O estudo contribui para a atualização do paradoxo de Solow ao demonstrar sua manifestação no nível microeconômico no contexto da inteligência artificial. O estudo amplia a aplicação da Teoria da Agência ao analisar como executivos utilizam narrativas tecnológicas como mecanismo de proteção frente a decisões impopulares. O estudo introduz o conceito de AI Washing aplicado ao discurso sobre demissões, contribuindo para a literatura sobre governança corporativa e transparência.
""".strip().split("\n\n")

if conclusao_idx:
    for text in expand_conc:
        p = new_doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.first_line_indent = Inches(0.5)
        run = p.add_run(text.strip())
        run.font.name = "Times New Roman"
        run.font.size = Pt(11)

    print(f"   {len(expand_conc)} parágrafos de expansão da conclusão inseridos")

# Salvar
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
