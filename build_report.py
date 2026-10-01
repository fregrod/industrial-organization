import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import os

doc = docx.Document()

# -------------------------------------------------------------
# CONFIGURAÇÃO DE PÁGINA (ABNT: Superior 3cm, Esquerda 3cm, Inferior 2cm, Direita 2cm)
# -------------------------------------------------------------
sections = doc.sections
for s in sections:
    s.top_margin = Inches(1.18)      # 3.0 cm
    s.left_margin = Inches(1.18)     # 3.0 cm
    s.bottom_margin = Inches(0.79)   # 2.0 cm
    s.right_margin = Inches(0.79)    # 2.0 cm
    s.page_width = Inches(8.27)      # A4
    s.page_height = Inches(11.69)    # A4

# -------------------------------------------------------------
# ESTILOS E CORES
# -------------------------------------------------------------
NAVY = RGBColor(27, 54, 93)     # #1B365D
DARK_GRAY = RGBColor(45, 55, 72) # #2D3748
BLACK = RGBColor(0, 0, 0)

normal_style = doc.styles['Normal']
normal_style.font.name = 'Arial'
normal_style.font.size = Pt(11)
normal_style.font.color.rgb = BLACK
normal_style.paragraph_format.line_spacing = 1.25
normal_style.paragraph_format.space_after = Pt(4)
normal_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

def add_title_1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = 'Arial'
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = NAVY
    return p

def add_title_2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = 'Arial'
    run.font.size = Pt(11.5)
    run.font.bold = True
    run.font.color.rgb = NAVY
    return p

def add_body(text, space_after=4, bold_prefix=None):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.25
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.first_line_indent = Inches(0.49) # 1.25 cm
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = 'Arial'
        r_pre.font.size = Pt(11)
        r_pre.font.bold = True
    run = p.add_run(text)
    run.font.name = 'Arial'
    run.font.size = Pt(11)
    return p

def add_caption(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = 'Arial'
    run.font.size = Pt(9.5)
    run.font.bold = True
    run.font.color.rgb = DARK_GRAY
    return p

def add_source(text="Fonte: Elaborado pelos autores (2026)."):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.font.name = 'Arial'
    run.font.size = Pt(9)
    run.font.italic = True
    run.font.color.rgb = DARK_GRAY
    return p

def set_cell_shading(cell, color_hex):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def set_cell_margins(cell, top=80, bottom=80, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def style_table(table, col_widths, align_cols=None):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row in enumerate(table.rows):
        # Tratar altura da linha
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for c_idx, cell in enumerate(row.cells):
            cell.width = col_widths[c_idx]
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.05
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.space_before = Pt(1)
            align = align_cols[c_idx] if align_cols else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.alignment = align
            for run in p.runs:
                run.font.name = 'Arial'
                if r_idx == 0:
                    run.font.size = Pt(8.5)
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255)
                elif r_idx == len(table.rows) - 1:
                    run.font.size = Pt(8.5)
                    run.font.bold = True
                else:
                    run.font.size = Pt(8.5)
            if r_idx == 0:
                set_cell_shading(cell, "1B365D")
            elif r_idx == len(table.rows) - 1:
                set_cell_shading(cell, "DCE6F1")
            elif r_idx % 2 == 1:
                set_cell_shading(cell, "F8FAFC")

# =============================================================
# PÁGINA 1: CAPA OFICIAL (UNB / FT)
# =============================================================
p_header = doc.add_paragraph()
p_header.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_header.paragraph_format.space_after = Pt(2)
r = p_header.add_run("UNIVERSIDADE DE BRASÍLIA – UnB\nFACULDADE DE TECNOLOGIA – FT\nDEPARTAMENTO DE ENGENHARIA DE PRODUÇÃO (EPR)\nDISCIPLINA: ORGANIZAÇÃO INDUSTRIAL")
r.font.name = 'Arial'
r.font.size = Pt(11)
r.font.bold = True
r.font.color.rgb = NAVY

for _ in range(4):
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(12)

p_title = doc.add_paragraph()
p_title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_title.paragraph_format.space_after = Pt(8)
r_t = p_title.add_run("RELATÓRIO TÉCNICO 1:\nPLANEJAMENTO AGREGADO DA PRODUÇÃO EM MANUFATURA")
r_t.font.name = 'Arial'
r_t.font.size = Pt(15)
r_t.font.bold = True
r_t.font.color.rgb = NAVY

p_sub = doc.add_paragraph()
p_sub.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_sub.paragraph_format.space_after = Pt(24)
r_s = p_sub.add_run("Estudo de Caso e Otimização Tática na Indústria de Confecção de Uniformes Profissionais:\nVértice Confecções & Uniformes Ltda. (Polo Industrial de Taguatinga - DF)")
r_s.font.name = 'Arial'
r_s.font.size = Pt(11.5)
r_s.font.italic = True
r_s.font.color.rgb = DARK_GRAY

for _ in range(4):
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(12)

# Bloco de identificação à direita
p_ident = doc.add_paragraph()
p_ident.paragraph_format.left_indent = Inches(3.2)
p_ident.paragraph_format.space_after = Pt(16)
p_ident.paragraph_format.line_spacing = 1.2
r_id = p_ident.add_run(
    "Trabalho acadêmico e técnico apresentado como requisito parcial de avaliação para a obtenção de menção na disciplina de Organização Industrial (semestre letivo 2026.2).\n\n"
    "Docentes Responsáveis:\n"
    "• Prof. Dr. Marcelo Carneiro Gonçalves\n"
    "• Profa. Ma. Camilla Pimenta Barbosa\n"
    "• Prof. Me. João Victor Queiroz\n\n"
    "Grupo de Discentes:\n"
    "• Integrante 1 (Matrícula: ____________)\n"
    "• Integrante 2 (Matrícula: ____________)\n"
    "• Integrante 3 (Matrícula: ____________)\n"
    "• Integrante 4 (Matrícula: ____________)\n"
    "• Integrante 5 (Matrícula: ____________)"
)
r_id.font.name = 'Arial'
r_id.font.size = Pt(10)

for _ in range(2):
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(12)

p_footer = doc.add_paragraph()
p_footer.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_f = p_footer.add_run("BRASÍLIA – DF\nOUTUBRO DE 2026")
r_f.font.name = 'Arial'
r_f.font.size = Pt(10.5)
r_f.font.bold = True

doc.add_page_break()

# =============================================================
# PÁGINA 2: INTRODUÇÃO E OBJETIVOS
# =============================================================
add_title_1("1. INTRODUÇÃO E OBJETIVOS DA ANÁLISE")

add_body(
    "No ambiente competitivo e volátil da manufatura contemporânea, a tomada de decisões na gestão de operações industriais "
    "exige uma articulação precisa entre as metas estratégicas corporativas e o planejamento detalhado no chão de fábrica. "
    "Nesse cenário, o Planejamento Agregado da Produção (PAP) consolida-se como um pilar tático essencial do sistema de "
    "Planejamento e Controle da Produção (PCP). Conforme conceituado por Slack et al. (2018) e Corrêa e Corrêa (2017), o PAP visa "
    "estabelecer os níveis globais de produção, estoques, força de trabalho, horas extras e subcontratação para um horizonte de médio "
    "prazo (tipicamente de 6 a 18 meses), equilibrando a previsão de demanda agregada com a capacidade produtiva instalada da organização."
)

add_body(
    "No contexto da manufatura, diferentemente das operações de serviços puros onde a simultaneidade impede a estocagem do produto acabado "
    "(FITZSIMMONS et al., 2014), as empresas industriais dispõem de múltiplos graus de liberdade para amortecer flutuações sazonais de demanda. "
    "Essas alternativas englobam: (i) a manutenção de uma taxa de produção constante com absorção das oscilações mediante estoques de produtos acabados; "
    "(ii) o acompanhamento rigoroso da demanda com modulação contínua do quadro de colaboradores via admissões e demissões; e (iii) a adoção "
    "de estratégias híbridas ou em patamares, que combinam capacidade regular constante com o acionamento de horas extras e subcontratações "
    "(KRAJEWSKI et al., 2014; QUEIROZ, 2026)."
)

add_body(
    "A relevância da aplicação do PAP torna-se ainda mais crítica no segmento de Micro, Pequenas e Médias Empresas (MPMEs), que respondem "
    "por parcela substancial do emprego industrial no Brasil, mas frequentemente operam de forma reativa, sem ferramentas analíticas formais "
    "de previsão e planejamento tático (SEBRAE, 2024). Em indústrias manufatureiras caracterizadas por sazonalidade acentuada — como o setor "
    "têxtil e de confecção —, oscilações drásticas de pedidos provocam gargalos severos de capacidade, custos astronômicos de horas extras "
    "emergenciais, atrasos crônicos na entrega de pedidos corporativos ou, no extremo oposto, capital de giro excessivamente imobilizado em estoques."
)

add_title_2("1.1. Objetivos do Trabalho")
add_body(
    "O presente relatório técnico tem como objetivo central estruturar, modelar matematicamente e avaliar comparativamente planos "
    "agregados de produção para um caso real da manufatura de vestuário corporativo e escolar: a empresa Vértice Confecções & Uniformes Ltda., "
    "sediada no Distrito Federal. Especificamente, buscam-se os seguintes objetivos operacionais e acadêmicos:",
    bold_prefix="Objetivo Geral: "
)

add_body("1. Mapear o sistema produtivo, caracterizando o ambiente de manufatura (MTS/MTO), o ponto de desacoplamento, a estrutura de processos e o gargalo fabril;", space_after=2)
add_body("2. Consolidar os dados empíricos de previsão de demanda e parâmetros de custos para um horizonte completo de 12 meses;", space_after=2)
add_body("3. Formular e calcular com rigor matemático três estratégias clássicas de PAP: Nivelamento (Taxa Constante), Acompanhamento da Demanda (Chase) e Estratégia Híbrida em Patamares;", space_after=2)
add_body("4. Confrontar as estratégias sob as óticas financeira (Plano Financeiro de custos totais) e operacional (nível de serviço, estabilidade do quadro, riscos trabalhistas e de qualidade);", space_after=2)
add_body("5. Propor a solução de maior viabilidade técnico-econômica, delineando recomendações de implementação para a diretoria da fábrica.")

doc.add_page_break()

# =============================================================
# PÁGINA 3: CARACTERIZAÇÃO DO CASO E PESQUISA
# =============================================================
add_title_1("2. CARACTERIZAÇÃO DO CASO E METODOLOGIA DE PESQUISA")

add_body(
    "A unidade de estudo selecionada para esta análise é a Vértice Confecções & Uniformes Profissionais Ltda. - EPP, indústria de manufatura "
    "fundada em 2012 e sediada no Setor de Desenvolvimento Econômico de Taguatinga (Polo Industrial do Distrito Federal). A empresa atua "
    "na produção de vestuário profissional corporativo (uniformes industriais, uniformes para o setor de comércio e serviços e jalecos operacionais) "
    "e uniformes escolares para colégios privados da região Centro-Oeste."
)

add_title_2("2.1. Estrutura Fabril, Roteamento de Operações e Gargalo")
add_body(
    "O parque fabril da empresa conta com 20 costureiras e operadores diretos de chão de fábrica, organizados em quatro etapas sequenciais de "
    "transformação (Figura 1): Enfesto e Corte, Bordado/Estamparia, Células de Costura e Acabamento/Embalagem. O ambiente de produção caracteriza-se "
    "como híbrido: nos itens padronizados de alto giro (calças operacionais em brim pesado e camisas polo básicas), a empresa opera no regime "
    "Make-to-Stock (MTS / Produção Empurrada) para assegurar pronta-entrega; já a aplicação de logomarcas, bordados corporativos e detalhamentos "
    "específicos ocorre sob encomenda no regime Assemble-to-Order (ATO / Produção Puxada), definindo o ponto de desacoplamento de pedidos (SLACK et al., 2018)."
)

# Inserção da Figura 1
doc.add_paragraph().paragraph_format.space_before = Pt(4)
p_img1 = doc.add_paragraph()
p_img1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_img1.paragraph_format.space_after = Pt(2)
p_img1.add_run().add_picture("figuras/figura1_fluxo_processo.png", width=Inches(6.2))
add_caption("Figura 1: Mapeamento do Fluxo Operacional na Vértice Confecções (Gargalo: Células de Costura).")
add_source()

add_body(
    "Por meio de estudos de cronoanálise industrial, a gerência fabril identificou que o setor de Costura configura o gargalo do sistema produtivo "
    "(Teoria das Restrições - GOLDRATT, 2014). Para viabilizar a modelagem do PAP, todas as variantes de produtos foram consolidadas em uma "
    "família agregada homogênea denominada 'Conjunto de Uniforme Profissional Padrão' (composto por uma calça operacional em brim pesado e uma "
    "camisa polo em malha piquet com bordado). O Tempo Padrão (TP) apurado para a montagem de um conjunto completo no gargalo é de 0,80 horas (48 minutos)."
)

add_title_2("2.2. Metodologia de Coleta e Tratamento de Dados")
add_body(
    "O levantamento empírico combinou análise documental interna da empresa e parâmetros setoriais públicos consolidados pelo SEBRAE (2024), "
    "pela Associação Brasileira da Indústria Têxtil e de Confecção (ABIT) e pelos pisos salariais do Sindicato das Indústrias do Vestuário do DF (SINDIVEST-DF). "
    "A jornada normal de trabalho é de 44 horas semanais, totalizando 160 horas produtivas mensais líquidas por colaboradora (descontadas pausas ergonômicas "
    "da NR-17 e tolerâncias de fadiga). Dessa forma, a produtividade em turno normal por costureira é calculada por: "
    "Produtividade = 160 h/mês ÷ 0,80 h/conjunto = 200 conjuntos/mês. Com o quadro base de 20 colaboradoras, a capacidade nominal normal instalada é de 4.000 conjuntos/mês."
)

add_body(
    "A série histórica de demanda agregada abrange um horizonte de 12 meses (Janeiro a Dezembro), revelando dois fortes picos sazonais: "
    "(i) no 1º trimestre (Janeiro a Março), com demandas de até 5.200 conjuntos/mês, impulsionadas pelo retorno às aulas e renovação anual de contratos corporativos; "
    "e (ii) no final do ano (Novembro e Dezembro), decorrente de contratações temporárias do comércio e encerramento de verbas orçamentárias."
)

doc.add_page_break()

# =============================================================
# PÁGINA 4: APLICAÇÃO PRÁTICA DO PAP - PARÂMETROS E ESTRATÉGIA 1
# =============================================================
add_title_1("3. APLICAÇÃO PRÁTICA DO PLANEJAMENTO AGREGADO DA PRODUÇÃO (PAP)")

add_body(
    "A formulação matemática do PAP adota rigorosamente as equações de balanço de estoques, dimensionamento de força de trabalho e planos financeiros "
    "lecionadas na disciplina de Organização Industrial (QUEIROZ, 2026; BARBOSA & GONÇALVES, 2026). Para cada período mensal t (t = 1, ..., 12), as relações fundamentais são dadas por:",
    bold_prefix="Modelagem Matemática: "
)

add_body("• Produção Agregada Total: PT(t) = Normal(t) + HE(t) + SUB(t)", space_after=2)
add_body("• Balanço Material e Estoques: EF(t) = EI(t) + PT(t) - D(t), com EI(t+1) = EF(t)", space_after=2)
add_body("• Estoque Médio Mensal: EM(t) = [EI(t) + EF(t)] / 2", space_after=2)
add_body("• Demanda Efetiva com Backlog: D_efetiva(t) = D(t) + Backlog(t-1)", space_after=2)
add_body("• Custos Financeiros: C_Total = C_Prod + C_HE + C_Sub + C_Estoque + C_Atraso + C_Contratação + C_Demissão", space_after=6)

add_body("A Tabela 1 sintetiza a matriz de parâmetros operacionais, custos unitários e a série de previsão de demanda para os 12 meses.")

# Tabela 1: Parâmetros e Demanda
add_caption("Tabela 1: Parâmetros Operacionais, Estrutura de Custos e Demanda Mensal Prevista.")
t1 = doc.add_table(rows=14, cols=4)
col_w_t1 = [Inches(1.8), Inches(1.0), Inches(1.3), Inches(2.2)]
align_t1 = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]

headers_t1_vals = ["Parâmetro / Variável", "Valor", "Unidade", "Descrição Técnica / Base Normativa"]
for c_idx, val in enumerate(headers_t1_vals):
    t1.cell(0, c_idx).paragraphs[0].text = val

dados_t1 = [
    ("Quadro Base de Operadoras (N0)", "20", "costureiras", "Força de trabalho efetiva inicial"),
    ("Jornada Mensal Líquida", "160", "horas/mês", "44h semanais com pausas da NR-17"),
    ("Tempo Padrão no Gargalo (TP)", "0,80", "h/conjunto", "Cronoanálise na costura (48 minutos)"),
    ("Produtividade Normal Unitária", "200", "conjuntos/op./mês", "160 h ÷ 0,80 h/conjunto"),
    ("Capacidade Normal Base (Pn)", "4.000", "conjuntos/mês", "20 costureiras × 200 un/mês"),
    ("Estoque Inicial de Segurança (EI1)", "400", "conjuntos", "Estoque remanescente do ano anterior"),
    ("Custo Turno Normal (cn)", "R$ 38,00", "R$/conjunto", "Tecido, insumos, MOD direta e CIF"),
    ("Custo Turno Extra / HE (che)", "R$ 46,00", "R$/conjunto", "Adicional de 50% na MOD + encargos"),
    ("Custo Subcontratação (csub)", "R$ 52,00", "R$/conjunto", "Facção externa de costura homologada"),
    ("Custo de Posse de Estoque (ce)", "R$ 3,00", "R$/un./mês", "Armazenagem, seguro e custo de capital"),
    ("Custo de Ruptura / Atraso (ca)", "R$ 20,00", "R$/un./mês", "Multa contratual de SLA corporativo"),
    ("Custo de Admissão (Cc)", "R$ 1.600,00", "R$/operadora", "Recrutamento, ASO admissional e treino"),
    ("Custo de Desligamento (Cd)", "R$ 3.000,00", "R$/operadora", "Aviso prévio, multa 40% FGTS e rescisão")
]

for r_idx, row in enumerate(dados_t1, start=1):
    for c_idx, val in enumerate(row):
        t1.cell(r_idx, c_idx).paragraphs[0].text = val

style_table(t1, col_w_t1, align_t1)
add_source()

add_title_2("3.1. Estratégia 1: Produção Constante (Nivelamento / Level Strategy)")
add_body(
    "A primeira alternativa avaliada consiste em manter a capacidade produtiva rigorosamente constante em 4.000 conjuntos/mês, com quadro fixo de "
    "20 operadoras ao longo de todo o ano. Como a demanda agregada anual soma 48.000 conjuntos, a média aritmética exata é de 4.000 conjuntos/mês. "
    "Sob essa política, a fábrica não realiza horas extras, não contrata, não demite e não subcontrata. As oscilações de demanda são absorvidas "
    "exclusivamente pela flutuação dos estoques de produtos acabados ou, na sua insuficiência, pelo acúmulo de atrasos (backlog) transferidos "
    "para atendimento nos meses subsequentes (Tabela 2)."
)

# Tabela 2: Estratégia 1
add_caption("Tabela 2: Plano de Produção e Plano Financeiro da Estratégia 1 (Nivelamento / Taxa Constante).")
t2 = doc.add_table(rows=14, cols=9)
col_w_t2 = [Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.7), Inches(1.0)]
align_t2 = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]

headers_t2_vals = ["Mês", "Demanda", "D. Efetiva", "Prod.", "Est. Final", "Backlog", "Est. Médio", "C. Estoque", "Custo Total"]
for c_idx, val in enumerate(headers_t2_vals):
    t2.cell(0, c_idx).paragraphs[0].text = val

dados_t2 = [
    ("Jan", "4.600", "4.600", "4.000", "0", "200", "200", "R$ 600", "R$ 156.600"),
    ("Fev", "5.200", "5.400", "4.000", "0", "1.400", "0", "R$ 0", "R$ 180.000"),
    ("Mar", "4.800", "6.200", "4.000", "0", "2.200", "0", "R$ 0", "R$ 196.000"),
    ("Abr", "3.600", "5.800", "4.000", "0", "1.800", "0", "R$ 0", "R$ 188.000"),
    ("Mai", "3.200", "5.000", "4.000", "0", "1.000", "0", "R$ 0", "R$ 172.000"),
    ("Jun", "3.000", "4.000", "4.000", "0", "0", "0", "R$ 0", "R$ 152.000"),
    ("Jul", "3.800", "3.800", "4.000", "200", "0", "100", "R$ 300", "R$ 152.300"),
    ("Ago", "3.600", "3.600", "4.000", "600", "0", "400", "R$ 1.200", "R$ 153.200"),
    ("Set", "3.400", "3.400", "4.000", "1.200", "0", "900", "R$ 2.700", "R$ 154.700"),
    ("Out", "3.800", "3.800", "4.000", "1.400", "0", "1.300", "R$ 3.900", "R$ 155.900"),
    ("Nov", "5.000", "5.000", "4.000", "400", "0", "900", "R$ 2.700", "R$ 154.700"),
    ("Dez", "4.000", "4.000", "4.000", "400", "0", "400", "R$ 1.200", "R$ 153.200"),
    ("TOTAL", "48.000", "-", "48.000", "400", "6.600", "350", "R$ 36.600", "R$ 1.968.600")
]

for r_idx, row in enumerate(dados_t2, start=1):
    for c_idx, val in enumerate(row):
        t2.cell(r_idx, c_idx).paragraphs[0].text = val

style_table(t2, col_w_t2, align_t2)
add_source()

add_body(
    "Como demonstrado na Tabela 2, embora a capacidade anual média atenda ao montante global de demanda, a concentração excessiva de pedidos "
    "no primeiro trimestre colapsa a operação logo em Janeiro. O estoque inicial de 400 conjuntos é totalmente consumido, gerando 200 conjuntos de atraso. "
    "Em Fevereiro e Março, o backlog escala criticamente para 1.400 e 2.200 conjuntos pendentes. Essa demanda reprimida só é plenamente eliminada em Junho, "
    "acumulando expressivos R$ 108.000,00 exclusivamente em custos de atraso e penalidades contratuais, elevando o custo anual para R$ 1.968.600,00."
)

doc.add_page_break()

# =============================================================
# PÁGINA 5: ESTRATÉGIA 2 - ACOMPANHAMENTO DA DEMANDA (CHASE)
# =============================================================
add_title_2("3.2. Estratégia 2: Acompanhamento da Demanda (Chase Strategy)")

add_body(
    "A segunda abordagem investigada é a estratégia pura de acompanhamento da demanda (Chase), na qual a capacidade produtiva mensal é ajustada "
    "para igualar exatamente o volume de pedidos do período (PT(t) = D(t)). Dessa forma, a produção acompanha fielmente os picos e vales do mercado, "
    "mantendo o estoque de produtos acabados constante no nível de segurança de 400 unidades (EM = 400 conjuntos), eliminando integralmente rupturas e atrasos."
)

add_body(
    "Para viabilizar esse acompanhamento, o número de operadoras ativas (N(t)) varia mês a mês em função da produtividade de 200 conjuntos/operadora "
    "(N(t) = Demanda(t) ÷ 200). A fábrica inicia o horizonte com 20 costureiras e executa contratações sempre que N(t) > N(t-1) e demissões "
    "quando N(t) < N(t-1). Os custos de admissão (Cc = R$ 1.600,00) e de desligamento/rescisão (Cd = R$ 3.000,00) são debitados imediatamente (Tabela 3)."
)

# Tabela 3: Estratégia 2
add_caption("Tabela 3: Plano de Produção e Plano Financeiro da Estratégia 2 (Acompanhamento / Chase).")
t3 = doc.add_table(rows=14, cols=9)
col_w_t3 = [Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.8), Inches(0.8), Inches(1.0)]
align_t3 = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]

headers_t3_vals = ["Mês", "Demanda", "Operadoras", "Admissões", "Demissões", "Produção", "C. Admissão", "C. Rescisão", "Custo Total"]
for c_idx, val in enumerate(headers_t3_vals):
    t3.cell(0, c_idx).paragraphs[0].text = val

dados_t3 = [
    ("Jan", "4.600", "23", "3", "0", "4.600", "R$ 4.800", "R$ 0", "R$ 180.800"),
    ("Fev", "5.200", "26", "3", "0", "5.200", "R$ 4.800", "R$ 0", "R$ 203.600"),
    ("Mar", "4.800", "24", "0", "2", "4.800", "R$ 0", "R$ 6.000", "R$ 189.600"),
    ("Abr", "3.600", "18", "0", "6", "3.600", "R$ 0", "R$ 18.000", "R$ 156.000"),
    ("Mai", "3.200", "16", "0", "2", "3.200", "R$ 0", "R$ 6.000", "R$ 128.800"),
    ("Jun", "3.000", "15", "0", "1", "3.000", "R$ 0", "R$ 3.000", "R$ 118.200"),
    ("Jul", "3.800", "19", "4", "0", "3.800", "R$ 6.400", "R$ 0", "R$ 152.000"),
    ("Ago", "3.600", "18", "0", "1", "3.600", "R$ 0", "R$ 3.000", "R$ 141.000"),
    ("Set", "3.400", "17", "0", "1", "3.400", "R$ 0", "R$ 3.000", "R$ 133.400"),
    ("Out", "3.800", "19", "2", "0", "3.800", "R$ 3.200", "R$ 0", "R$ 148.800"),
    ("Nov", "5.000", "25", "6", "0", "5.000", "R$ 9.600", "R$ 0", "R$ 200.800"),
    ("Dez", "4.000", "20", "0", "5", "4.000", "R$ 0", "R$ 15.000", "R$ 168.200"),
    ("TOTAL", "48.000", "20", "18", "18", "48.000", "R$ 28.800", "R$ 54.000", "R$ 1.921.200")
]

for r_idx, row in enumerate(dados_t3, start=1):
    for c_idx, val in enumerate(row):
        t3.cell(r_idx, c_idx).paragraphs[0].text = val

style_table(t3, col_w_t3, align_t3)
add_source()

add_body(
    "A análise da Tabela 3 revela que, sob o prisma financeiro estrito, a Estratégia 2 totaliza R$ 1.921.200,00, superando a Estratégia 1 em R$ 47.400,00. "
    "Contudo, essa aparente vantagem esconde uma rotatividade (turnover) industrial catastrófica: a empresa é forçada a contratar 18 costureiras e "
    "demitir 18 colaboradoras em um intervalo de apenas 12 meses. Somente em custos diretos de contratação e rescisão, despendem-se R$ 82.800,00."
)

add_body(
    "Na prática fabril de confecção de vestuário, o custo oculto dessa estratégia é devastador: a curva de aprendizagem é constantemente reiniciada, "
    "o índice de costuras defeituosas e retrabalhos dispara nos meses de contratação massiva (Janeiro, Fevereiro e Novembro), o clima organizacional "
    "deteriora-se com demissões em massa pós-pico (Abril e Dezembro) e a exposição a litígios e passivos trabalhistas sob a CLT atinge níveis insustentáveis."
)

doc.add_page_break()

# =============================================================
# PÁGINA 6: ESTRATÉGIA 3 - HÍBRIDA / PATAMARES (RECOMENDADA)
# =============================================================
add_title_2("3.3. Estratégia 3: Variação em Patamares / Híbrida (Recomendada)")

add_body(
    "Diante das graves disfunções das estratégias puras, estruturou-se uma Estratégia Híbrida em Patamares, que combina de forma balanceada capacidade "
    "regular fixa, estoque regulador moderado e flexibilidade operacional temporária. A força de trabalho é fixada em 19 operadoras permanentes, "
    "gerando uma produção em turno normal estável de 3.800 conjuntos/mês (19 × 200 un.), eliminando 100% das admissões e demissões sazonais."
)

add_body(
    "Para suprir os picos agudos de demanda sem incorrer em atrasos, a empresa adota uma hierarquia decisória estrita:",
    bold_prefix="Lógica de Ajuste Operacional: "
)
add_body("1. Produção Normal (3.800 un/mês): Base fabril estável e de alta produtividade;", space_after=2)
add_body("2. Horas Extras Internas (che = R$ 46,00): Acionadas preferencialmente até o limite de 800 conjuntos/mês, respeitando o teto legal de 2 horas extras/dia por operadora (Art. 59 da CLT);", space_after=2)
add_body("3. Subcontratação / Facção Externa (csub = R$ 52,00): Acionada estritamente quando a demanda do período excede a disponibilidade interna somada do estoque inicial, turno normal e hora extra;", space_after=2)
add_body("4. Estoque Regulador: Nos meses de menor demanda (Abril a Outubro), a produção normal de 3.800 un/mês gera um acúmulo suave de estoque que atinge 2.200 conjuntos em Outubro, o qual absorve com folga o pico de Novembro (5.000 un.) com custo normal.", space_after=6)

# Tabela 4: Estratégia 3
add_caption("Tabela 4: Plano de Produção e Plano Financeiro da Estratégia 3 (Híbrida em Patamares - Solução Proposta).")
t4 = doc.add_table(rows=14, cols=10)
col_w_t4 = [Inches(0.6), Inches(0.6), Inches(0.6), Inches(0.6), Inches(0.6), Inches(0.6), Inches(0.6), Inches(0.6), Inches(0.8), Inches(1.0)]
align_t4 = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]

headers_t4_vals = ["Mês", "Demanda", "Normal", "HE (un)", "Sub (un)", "Prod. Total", "Est. Final", "Est. Médio", "C. Estoque", "Custo Total"]
for c_idx, val in enumerate(headers_t4_vals):
    t4.cell(0, c_idx).paragraphs[0].text = val

dados_t4 = [
    ("Jan", "4.600", "3.800", "400", "0", "4.200", "0", "200", "R$ 600", "R$ 163.400"),
    ("Fev", "5.200", "3.800", "800", "600", "5.200", "0", "0", "R$ 0", "R$ 212.400"),
    ("Mar", "4.800", "3.800", "800", "200", "4.800", "0", "0", "R$ 0", "R$ 191.600"),
    ("Abr", "3.600", "3.800", "0", "0", "3.800", "200", "100", "R$ 300", "R$ 144.700"),
    ("Mai", "3.200", "3.800", "0", "0", "3.800", "800", "500", "R$ 1.500", "R$ 145.900"),
    ("Jun", "3.000", "3.800", "0", "0", "3.800", "1.600", "1.200", "R$ 3.600", "R$ 148.000"),
    ("Jul", "3.800", "3.800", "0", "0", "3.800", "1.600", "1.600", "R$ 4.800", "R$ 149.200"),
    ("Ago", "3.600", "3.800", "0", "0", "3.800", "1.800", "1.700", "R$ 5.100", "R$ 149.500"),
    ("Set", "3.400", "3.800", "0", "0", "3.800", "2.200", "2.000", "R$ 6.000", "R$ 150.400"),
    ("Out", "3.800", "3.800", "0", "0", "3.800", "2.200", "2.200", "R$ 6.600", "R$ 151.000"),
    ("Nov", "5.000", "3.800", "0", "0", "3.800", "1.000", "1.600", "R$ 4.800", "R$ 149.200"),
    ("Dez", "4.000", "3.800", "0", "0", "3.800", "800", "900", "R$ 2.700", "R$ 147.100"),
    ("TOTAL", "48.000", "45.600", "2.000", "800", "48.400", "800", "983", "R$ 36.000", "R$ 1.902.400")
]

for r_idx, row in enumerate(dados_t4, start=1):
    for c_idx, val in enumerate(row):
        t4.cell(r_idx, c_idx).paragraphs[0].text = val

style_table(t4, col_w_t4, align_t4)
add_source()

# Inserção da Figura 2
p_img2 = doc.add_paragraph()
p_img2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_img2.paragraph_format.space_after = Pt(2)
p_img2.add_run().add_picture("figuras/figura2_demanda_vs_producao.png", width=Inches(6.0))
add_caption("Figura 2: Comportamento Mensal da Demanda versus Taxas de Produção por Estratégia de PAP.")
add_source()

doc.add_page_break()

# =============================================================
# PÁGINA 7: COMPARATIVO DO PLANO FINANCEIRO
# =============================================================
add_title_2("3.4. Síntese e Comparativo Consolidado do Plano Financeiro")

add_body(
    "A Tabela 5 apresenta a consolidação econômico-financeira das três estratégias de planejamento agregado investigadas, decompondo "
    "os custos totais por categoria e confrontando os principais indicadores de desempenho operacional (KPIs) de manufatura."
)

# Tabela 5: Comparativo Geral
add_caption("Tabela 5: Comparativo Consolidado do Plano Financeiro e Indicadores de Desempenho (12 Meses).")
t5 = doc.add_table(rows=16, cols=4)
col_w_t5 = [Inches(2.5), Inches(1.5), Inches(1.5), Inches(1.7)]
align_t5 = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]

headers_t5_vals = ["Rubrica de Custo / Indicador (KPI)", "Estratégia 1: Nivelamento", "Estratégia 2: Acompanhamento", "Estratégia 3: Híbrida (Recomendada)"]
for c_idx, val in enumerate(headers_t5_vals):
    t5.cell(0, c_idx).paragraphs[0].text = val

dados_t5 = [
    ("Produção em Turno Normal (cn = R$ 38,00)", "R$ 1.824.000,00", "R$ 1.824.000,00", "R$ 1.732.800,00"),
    ("Turno Extra / Horas Extras (che = R$ 46,00)", "R$ 0,00", "R$ 0,00", "R$ 92.000,00"),
    ("Subcontratação / Facção Externa (csub = R$ 52,00)", "R$ 0,00", "R$ 0,00", "R$ 41.600,00"),
    ("Manutenção / Posse de Estoque (ce = R$ 3,00)", "R$ 36.600,00", "R$ 14.400,00", "R$ 36.000,00"),
    ("Custos de Atraso na Entrega / Backlog (ca = R$ 20,00)", "R$ 108.000,00", "R$ 0,00", "R$ 0,00"),
    ("Custos de Contratação de Pessoal (Cc = R$ 1.600,00)", "R$ 0,00", "R$ 28.800,00", "R$ 0,00"),
    ("Custos de Demissão e Rescisões (Cd = R$ 3.000,00)", "R$ 0,00", "R$ 54.000,00", "R$ 0,00"),
    ("CUSTO TOTAL ANUAL DO PLANO (R$)", "R$ 1.968.600,00", "R$ 1.921.200,00", "R$ 1.902.400,00"),
    ("Economia Absoluta em Relação ao Nivelamento (R$)", "Referência (Base)", "R$ 47.400,00", "R$ 66.200,00"),
    ("Nível de Serviço / Pontualidade nas Entregas (%)", "79,2% (5 meses em atraso)", "100,0%", "100,0%"),
    ("Volume Total de Backlog Gerado no Ano", "6.600 un. acumuladas", "0 un.", "0 un."),
    ("Total de Admissões Realizadas no Horizonte", "0 costureiras", "18 costureiras", "0 costureiras"),
    ("Total de Desligamentos Realizados no Horizonte", "0 costureiras", "18 costureiras", "0 costureiras"),
    ("Estoque Final Remanescente no Mês 12", "400 conjuntos", "400 conjuntos", "800 conjuntos"),
    ("Conformidade Legal com a CLT (Art. 59)", "100% Regular", "100% Regular", "100% Regular (≤ 2h/dia)")
]

for r_idx, row in enumerate(dados_t5, start=1):
    for c_idx, val in enumerate(row):
        t5.cell(r_idx, c_idx).paragraphs[0].text = val

style_table(t5, col_w_t5, align_t5)
add_source()

# Inserção das Figuras 3 e 4
p_img4 = doc.add_paragraph()
p_img4.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_img4.paragraph_format.space_after = Pt(2)
p_img4.add_run().add_picture("figuras/figura4_comparativo_custos.png", width=Inches(5.6))
add_caption("Figura 4: Estrutura Comparativa de Custos Totais Anuais entre as Estratégias de PAP.")
add_source()

doc.add_page_break()

# =============================================================
# PÁGINA 8: ANÁLISE CRÍTICA DOS RESULTADOS E DISCUSSÃO
# =============================================================
add_title_1("4. ANÁLISE CRÍTICA DOS RESULTADOS E DISCUSSÃO")

add_body(
    "Os resultados quantitativos obtidos no Plano Financeiro fornecem subsídios sólidos para uma análise profunda de engenharia de produção. "
    "A Estratégia 3 (Híbrida em Patamares) consagra-se como a mais eficiente sob a ótica econômica, totalizando R$ 1.902.400,00 — gerando uma economia "
    "líquida de R$ 66.200,00 frente ao Nivelamento (-3,36%) e de R$ 18.800,00 em relação ao Acompanhamento (-0,98%). Contudo, em operações manufatureiras "
    "reais, a decisão gerencial jamais deve ser pautada exclusivamente pelo menor custo contábil: faz-se imperativo examinar os trade-offs "
    "estruturais, as restrições da legislação trabalhista, a qualidade do produto e os impactos no capital de giro.",
    bold_prefix="Avaliação Técnico-Econômica: "
)

add_title_2("4.1. Confronto das Estratégias e Análise de Trade-offs")
add_body(
    "A fragilidade da Estratégia 1 (Nivelamento) reside na rigidez da taxa de produção. Em mercados com sazonalidade no início do horizonte, o estoque "
    "inicial mostra-se insuficiente para absorver o choque de demanda. Como consequência, a fábrica passa 5 meses consecutivos operando em situação "
    "de ruptura de estoque (backlog que atinge 2.200 conjuntos em Março). No setor de uniformes profissionais, onde os contratos corporativos preveem "
    "multas severas de SLA e cláusulas de rescisão por atraso, essa estratégia implicaria a perda de clientes-chave e deterioração irreversível da marca no mercado."
)

# Inserção da Figura 3
p_img3 = doc.add_paragraph()
p_img3.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_img3.paragraph_format.space_after = Pt(2)
p_img3.add_run().add_picture("figuras/figura3_estoques_e_backlog.png", width=Inches(6.2))
add_caption("Figura 3: Comparação da Dinâmica de Estoques e Backlog: Colapso na Estratégia 1 versus Estabilidade na Estratégia 3.")
add_source()

add_body(
    "Por sua vez, a Estratégia 2 (Acompanhamento) apresenta custos contábeis competitivos, mas sua operacionalização no chão de fábrica é inviável. "
    "A indústria têxtil e de confecção depende intensamente de habilidades motoras finas, destreza e coordenação visual no manuseio de máquinas de costura "
    "industriais (retas, overlocks, interlocks e pespontadeiras). Substituir 90% da equipe produtiva ao longo de 12 meses (18 admissões e 18 desligamentos) "
    "destrói a curva de experiência da fábrica. Operadoras recém-admitidas apresentam tempos de operação até 40% superiores ao Tempo Padrão (TP = 0,80 h), "
    "gerando desperdícios por defeitos de costura, assimetria de peças e refugo de tecidos nobres, custos esses que não são capturados pelo modelo estático do PAP."
)

add_title_2("4.2. Restrições Trabalhistas (CLT), Clima Organizacional e Qualidade")
add_body(
    "No tocante ao marco regulatório brasileiro, a Consolidação das Leis do Trabalho (CLT) impõe limites severos à flexibilização da jornada. "
    "O Artigo 59 da CLT veda expressamente a realização de mais de 2 horas extras diárias por empregado. Na modelagem da Estratégia 3, esse teto "
    "foi rigorosamente obedecido: para uma jornada de 8h diárias, 2h de prorrogação equivalem a um limite de 25% da jornada (máximo de 800 conjuntos/mês em HE). "
    "Além disso, a realização de horas extras concentrou-se apenas nos meses de pico (Janeiro a Março), evitando a fadiga física crônica das costureiras, "
    "doenças osteomusculares relacionadas ao trabalho (DORT/LER) e autuações pela fiscalização do Ministério do Trabalho e Emprego (MTE)."
)

add_body(
    "Quanto à subcontratação (facção externa), utilizada na Estratégia 3 com volumes de 600 unidades em Fevereiro e 200 em Março, destaca-se que "
    "a empresa deve implementar um protocolo rigoroso de qualificação e homologação de oficinas parceiras. A terceirização descontrolada no setor têxtil "
    "traz riscos de não conformidade de costuras e vulnerabilidade a passivos trabalhistas subsidiários (Súmula 331 do TST), mitigados na Vértice "
    "pelo fornecimento do tecido já enfestado e cortado, cabendo à facção estritamente a costura sob auditoria de qualidade in loco."
)

doc.add_page_break()

# =============================================================
# PÁGINA 9: LIMITAÇÕES, MELHORIAS E CONCLUSÃO
# =============================================================
add_title_2("4.3. Limitações do Modelo e Oportunidades de Melhoria Contínua")

add_body(
    "Embora o modelo de PAP desenvolvido represente um salto de maturidade para a gestão da Vértice Confecções, a análise de engenharia "
    "de produção aponta limitações que devem ser abordadas em ciclos posteriores de planejamento:",
    bold_prefix="Limitações da Modelagem: "
)
add_body("1. Determinização da Demanda: O modelo assume previsões determinísticas; em cenários reais, a variabilidade estocástica exige estoques de segurança dinâmicos calculados pelo desvio padrão dos erros de previsão;", space_after=2)
add_body("2. Homogeneidade da Família Agregada: A agregação em um 'conjunto padrão' desconsidera variações de setup no corte entre diferentes tecidos (brim, sarja, piquet e tactel), aspecto que deve ser detalhado no Plano Mestre de Produção (MPS);", space_after=2)
add_body("3. Curva de Aprendizagem Estática: Assume-se produtividade constante das operadoras, sem modelar as perdas transitórias de rendimento pós-férias ou pós-manutenção.", space_after=6)

add_body(
    "Para superar tais restrições e potencializar os ganhos da Estratégia 3, recomenda-se à diretoria da empresa a implementação das seguintes ações:",
    bold_prefix="Recomendações e Melhorias Contínuas: "
)
add_body("• Estruturação Formal do Processo de S&OP (Sales and Operations Planning): Instituir reuniões mensais interdisciplinares entre a área Comercial, PCP, Produção e Finanças, revisando o PAP com horizonte rolante de 12 meses (SLACK et al., 2018);", space_after=2)
add_body("• Aplicação de Ferramentas Lean Manufacturing: Implementar troca rápida de ferramentas (SMED) na mesa de corte e balanceamento de linha nas células de costura, reduzindo tempos mortos e elevando a produtividade natural do chão de fábrica;", space_after=2)
add_body("• Parcerias Estratégicas de Longo Prazo com Facções de Costura: Estabelecer contratos de exclusividade parcial com oficinas de costura do Polo de Taguatinga, garantindo padrões de costura homogêneos e pontualidade na entrega nos meses de Fevereiro e Março;", space_after=2)
add_body("• Gestão Visual e PCP Digital: Substituir o controle em pranchetas por painéis visuais de acompanhamento de metas diárias (Kanban e gráficos de Gantt), monitorando em tempo real os tempos de ciclo e o cumprimento do plano agregado.", space_after=8)

add_title_1("5. CONCLUSÃO")

add_body(
    "O desenvolvimento deste trabalho permitiu demonstrar a aplicação empírica dos conceitos fundamentais de Planejamento e Controle da Produção "
    "e Organização Industrial no ambiente real de uma indústria de confecção do Distrito Federal. O Planejamento Agregado da Produção comprovou ser "
    "um instrumento indispensável para desarmar conflitos clássicos entre a área comercial (que busca máxima pronta-entrega e variedade) e a área fabril "
    "(que demanda estabilidade de ritmo, previsibilidade de recursos e custos controlados)."
)

add_body(
    "A modelagem comparativa revelou com clareza as armadilhas de estratégias extremas: enquanto a Produção Constante (Nivelamento) gerou um colapso "
    "de atendimento com R$ 108.000,00 em multas por atraso (backlog de até 2.200 conjuntos), o Acompanhamento Puro da Demanda (Chase) impôs um "
    "turnover de 36 movimentações trabalhistas (18 admissões e 18 desligamentos), resultando em custos rescisórios de R$ 82.800,00 e severa instabilidade operacional."
)

add_body(
    "A solução ótima recomendada — Estratégia 3 (Híbrida em Patamares) — demonstrou superioridade incontestável em todas as dimensões competitivas "
    "de Nigel Slack e Terry Hill: alcançou o MENOR CUSTO TOTAL (R$ 1.902.400,00), assegurou 100% DE NÍVEL DE SERVIÇO (zero rupturas contratuais), "
    "manteve a equipe estável com 19 operadoras permanentes, garantiu a conformidade estrita com o limite de 2h extras da CLT e deixou um saldo saudável "
    "de 800 conjuntos de produtos acabados para iniciar o exercício seguinte. Conclui-se, portanto, que a conciliação harmônica entre capacidade "
    "própria, flexibilidade de horas extras e terceirização controlada constitui o paradigma tático de excelência para a manufatura de vestuário."
)

doc.add_page_break()

# =============================================================
# PÁGINA 10: REFERÊNCIAS BIBLIOGRÁFICAS (ABNT)
# =============================================================
add_title_1("6. REFERÊNCIAS BIBLIOGRÁFICAS")

referencias = [
    "ABIT – ASSOCIAÇÃO BRASILEIRA DA INDÚSTRIA TÊXTIL E DE CONFECÇÃO. Perfil do Setor Têxtil e de Confecção Brasileiro: Dados Estatísticos e Inteligência Setorial. São Paulo: ABIT, 2024. Disponível em: <https://www.abit.org.br>. Acesso em: 28 set. 2026.",
    "BARBOSA, C. P.; GONÇALVES, M. C. Organização Industrial – Aula 02: Planejamento Agregado da Produção em Serviços. Brasília: Universidade de Brasília (UnB), Faculdade de Tecnologia, 2026. 43 slides.",
    "BRASIL. Decreto-Lei nº 5.452, de 1º de maio de 1943. Aprova a Consolidação das Leis do Trabalho (CLT). Artigo 59 (Duração do Trabalho e Horas Extras). Brasília, DF: Presidência da República, 1943.",
    "CORRÊA, H. L.; CORRÊA, C. A. Administração de Produção e Operações: Manufatura e Serviços – Uma Abordagem Estratégica. 4. ed. São Paulo: Atlas, 2017.",
    "DAVIS, M. M.; AQUILANO, N. J.; CHASE, R. B. Fundamentos da Administração da Produção. 3. ed. Porto Alegre: Bookman, 2001.",
    "FITZSIMMONS, J. A.; FITZSIMMONS, M. J.; BORDOLOI, S. K. Administração de Serviços: Operações, Estratégia e Tecnologia da Informação. 7. ed. Porto Alegre: AMGH, 2014.",
    "GOLDRATT, E. M.; COX, J. A Meta: Um Processo de Aprimoramento Contínuo. 3. ed. São Paulo: Nobel, 2014.",
    "GONÇALVES, M. C. Conceitos Básicos de Gestão da Produção e Estratégia de Operações – Aula 01. Brasília: Universidade de Brasília (UnB), Faculdade de Tecnologia, 2026. 29 p.",
    "HILL, T. Manufacturing Strategy: Text and Cases. 3. ed. Boston: McGraw-Hill, 2000.",
    "KRAJEWSKI, L. J.; RITZMAN, L. P.; MALHOTRA, M. K. Administração de Produção e Operações. 8. ed. São Paulo: Pearson Prentice Hall, 2014.",
    "MOREIRA, D. A. Administração da Produção e Operações. 2. ed. São Paulo: Cengage Learning, 2012.",
    "QUEIROZ, J. V. Organização Industrial – Aula 03: Planejamento Agregado da Produção em Manufatura (PAP, S&OP e Capacidade). Brasília: Universidade de Brasília (UnB), Faculdade de Tecnologia, 2026. 26 slides.",
    "QUEIROZ, J. V. Modelo Computacional em Python para Planejamento Agregado da Produção em Manufatura. Google Colaboratory, 2026. Disponível em: <https://colab.research.google.com/drive/1UNxeDKP04DMtyab0cuSxG4N4JOY3XSgv>. Acesso em: 29 set. 2026.",
    "SEBRAE – SERVIÇO BRASILEIRO DE APOIO ÀS MICRO E PEQUENAS EMPRESAS. Relatório de Inteligência de Mercado: Indústria de Vestuário e Confecção – Desafios da Gestão de Produção e PCP nas MPMEs. Brasília: SEBRAE Nacional, 2024.",
    "SLACK, N.; BRANDON-JONES, A.; JOHNSTON, R. Administração da Produção. 8. ed. São Paulo: Atlas, 2018.",
    "TUBINO, D. F. Planejamento e Controle da Produção: Teoria e Prática. 3. ed. São Paulo: Atlas, 2017."
]

for ref in referencias:
    p_ref = doc.add_paragraph()
    p_ref.paragraph_format.line_spacing = 1.15
    p_ref.paragraph_format.space_after = Pt(6)
    p_ref.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run_ref = p_ref.add_run(ref)
    run_ref.font.name = 'Arial'
    run_ref.font.size = Pt(9.5)

# Salvar DOCX
doc.save("Relatorio_Tecnico_1_PAP_Manufatura.docx")
print("Relatório salvo com sucesso em Relatorio_Tecnico_1_PAP_Manufatura.docx!")
