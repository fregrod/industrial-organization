import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import subprocess
import pypdf

def generate_doc(font_body=10, line_spacing=1.15, img_width=5.0, table_font=7.5):
    doc = docx.Document()
    
    # Margens ABNT ajustadas
    for s in doc.sections:
        s.top_margin = Inches(0.85)      # 2.2 cm
        s.bottom_margin = Inches(0.80)   # 2.0 cm
        s.left_margin = Inches(0.98)     # 2.5 cm
        s.right_margin = Inches(0.85)    # 2.2 cm
        s.page_width = Inches(8.27)
        s.page_height = Inches(11.69)
        
    NAVY = RGBColor(27, 54, 93)
    DARK_GRAY = RGBColor(55, 65, 81)
    BLACK = RGBColor(0, 0, 0)
    
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Arial'
    normal_style.font.size = Pt(font_body)
    normal_style.font.color.rgb = BLACK
    normal_style.paragraph_format.line_spacing = line_spacing
    normal_style.paragraph_format.space_after = Pt(2.5)
    normal_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    def add_title_1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(font_body + 1.5)
        run.font.bold = True
        run.font.color.rgb = NAVY
        return p

    def add_title_2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(font_body + 0.5)
        run.font.bold = True
        run.font.color.rgb = NAVY
        return p

    def add_body(text, space_after=2.5, bold_prefix=None):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = line_spacing
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.first_line_indent = Inches(0.40)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.font.name = 'Arial'
            r_pre.font.size = Pt(font_body)
            r_pre.font.bold = True
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(font_body)
        return p

    def add_caption(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(8.5)
        run.font.bold = True
        run.font.color.rgb = DARK_GRAY
        return p

    def add_source(text="Fonte: Elaborado pelos autores (2026)."):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(8)
        run.font.italic = True
        run.font.color.rgb = DARK_GRAY
        return p

    def set_cell_shading(cell, color_hex):
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
        cell._tc.get_or_add_tcPr().append(shading)

    def set_cell_margins(cell, top=30, bottom=30, left=50, right=50):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)

    def style_table(table, col_widths, align_cols=None):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for r_idx, row in enumerate(table.rows):
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            for c_idx, cell in enumerate(row.cells):
                cell.width = col_widths[c_idx]
                set_cell_margins(cell)
                p = cell.paragraphs[0]
                p.paragraph_format.line_spacing = 1.0
                p.paragraph_format.space_after = Pt(0.5)
                p.paragraph_format.space_before = Pt(0.5)
                align = align_cols[c_idx] if align_cols else WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.alignment = align
                for run in p.runs:
                    run.font.name = 'Arial'
                    if r_idx == 0:
                        run.font.size = Pt(table_font)
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(255, 255, 255)
                    elif r_idx == len(table.rows) - 1:
                        run.font.size = Pt(table_font)
                        run.font.bold = True
                    else:
                        run.font.size = Pt(table_font)
                if r_idx == 0:
                    set_cell_shading(cell, "1B365D")
                elif r_idx == len(table.rows) - 1:
                    set_cell_shading(cell, "DCE6F1")
                elif r_idx % 2 == 1:
                    set_cell_shading(cell, "F8FAFC")

    # -----------------------------
    # PÁGINA 1: CAPA
    # -----------------------------
    p_h = doc.add_paragraph()
    p_h.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_h.paragraph_format.space_after = Pt(4)
    r = p_h.add_run("UNIVERSIDADE DE BRASÍLIA – UnB\nFACULDADE DE TECNOLOGIA – FT\nDEPARTAMENTO DE ENGENHARIA DE PRODUÇÃO (EPR)\nDISCIPLINA: ORGANIZAÇÃO INDUSTRIAL (2026.2)")
    r.font.name = 'Arial'
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = NAVY

    for _ in range(3):
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_after = Pt(10)

    p_t = doc.add_paragraph()
    p_t.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t.paragraph_format.space_after = Pt(6)
    r_t = p_t.add_run("RELATÓRIO TÉCNICO 1:\nPLANEJAMENTO AGREGADO DA PRODUÇÃO EM MANUFATURA")
    r_t.font.name = 'Arial'
    r_t.font.size = Pt(14)
    r_t.font.bold = True
    r_t.font.color.rgb = NAVY

    p_s = doc.add_paragraph()
    p_s.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_s.paragraph_format.space_after = Pt(20)
    r_s = p_s.add_run("Estudo de Caso e Otimização Tática na Indústria de Confecção de Uniformes Profissionais:\nVértice Confecções & Uniformes Ltda. (Polo Industrial de Taguatinga - DF)")
    r_s.font.name = 'Arial'
    r_s.font.size = Pt(10.5)
    r_s.font.italic = True
    r_s.font.color.rgb = DARK_GRAY

    for _ in range(3):
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_after = Pt(10)

    p_id = doc.add_paragraph()
    p_id.paragraph_format.left_indent = Inches(3.2)
    p_id.paragraph_format.space_after = Pt(16)
    p_id.paragraph_format.line_spacing = 1.15
    r_id = p_id.add_run(
        "Trabalho acadêmico e técnico apresentado como requisito parcial de avaliação para a obtenção de menção na disciplina de Organização Industrial (2026.2).\n\n"
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
    r_id.font.size = Pt(9.5)

    for _ in range(2):
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_after = Pt(10)

    p_f = doc.add_paragraph()
    p_f.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_f = p_f.add_run("BRASÍLIA – DF\nOUTUBRO DE 2026")
    r_f.font.name = 'Arial'
    r_f.font.size = Pt(10)
    r_f.font.bold = True

    doc.add_page_break()

    # -----------------------------
    # SEÇÃO 1: INTRODUÇÃO
    # -----------------------------
    add_title_1("1. INTRODUÇÃO E OBJETIVOS")
    add_body(
        "No ambiente concorrencial da manufatura moderna, a gestão tática de operações exige equilíbrio preciso entre as metas estratégicas corporativas e a capacidade física do chão de fábrica. O Planejamento Agregado da Produção (PAP) consolida-se como o elemento central do sistema de Planejamento e Controle da Produção (PCP), conectando o plano de negócios e o S&OP (Sales and Operations Planning) aos programas mestres detalhados (SLACK et al., 2018; CORRÊA; CORRÊA, 2017). O PAP visa estabelecer para um horizonte de médio prazo (6 a 18 meses) as taxas globais de produção, níveis de estoques, força de trabalho, turnos extras e subcontratações, harmonizando a demanda agregada com a capacidade produtiva instalada."
    )
    add_body(
        "Diferentemente do setor de serviços — cuja simultaneidade entre produção e consumo impede o amortecimento por meio de estoques de produtos acabados (FITZSIMMONS et al., 2014; BARBOSA; GONÇALVES, 2026) —, as empresas de manufatura contam com três famílias clássicas de alternativas táticas: (i) Estratégia de Nivelamento (capacidade constante com estoques reguladores); (ii) Estratégia de Acompanhamento (ajuste contínuo da força de trabalho via admissões e demissões); e (iii) Estratégia Híbrida em Patamares (capacidade própria regular somada a horas extras e terceirização nos picos) (KRAJEWSKI et al., 2014; QUEIROZ, 2026)."
    )
    add_body(
        "A aplicação do PAP revela-se crítica no segmento de Micro, Pequenas e Médias Empresas (MPMEs) confeccionistas, que respondem por parcela vital do emprego industrial no Brasil, mas frequentemente operam sem instrumentos analíticos formais de previsão e planejamento (SEBRAE, 2024). Em setores marcados por forte sazonalidade anual, a ausência de planejamento agregado acarreta gargalos crônicos, gastos vultosos com horas extras imprevistas, atrasos nas entregas ou capital de giro excessivamente imobilizado em estoques."
    )
    add_title_2("1.1. Objetivos da Análise")
    add_body(
        "Este relatório técnico tem por objetivo modelar, mensurar quantitativamente e avaliar comparativamente as estratégias de PAP para a indústria de confecção Vértice Confecções & Uniformes Ltda., sediada em Taguatinga - DF. Busca-se: (a) mapear o fluxo produtivo e gargalo fabril; (b) estruturar a matriz de custos e a demanda para 12 meses; (c) calcular com rigor as três estratégias de PAP; (d) analisar trade-offs sob aspectos financeiros, trabalhistas (CLT) e operacionais; e (e) justificar a estratégia ótima para implementação imediata pela diretoria."
    )

    # -----------------------------
    # SEÇÃO 2: CARACTERIZAÇÃO DO CASO
    # -----------------------------
    add_title_1("2. CARACTERIZAÇÃO DO CASO E METODOLOGIA DE PESQUISA")
    add_body(
        "A Vértice Confecções & Uniformes Profissionais Ltda. - EPP é uma indústria manufatureira de médio porte fundada em 2012 no Setor Industrial de Taguatinga - DF. A empresa atua na produção de uniformes corporativos pesados (calças em brim e jalecos industriais) e uniformes escolares padronizados para instituições de ensino do Centro-Oeste."
    )
    add_body(
        "O sistema produtivo adota ambiente híbrido Make-to-Stock (MTS / Produção Empurrada) e Assemble-to-Order (ATO / Produção Puxada) (SLACK et al., 2018). Para calças operacionais e camisas polo básicas, adota-se produção antecipada; os bordados, logomarcas e detalhamentos corporativos são inseridos sob pedido, fixando ali o ponto de desacoplamento do processo (Figura 1)."
    )
    
    p_img1 = doc.add_paragraph()
    p_img1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img1.paragraph_format.space_after = Pt(1)
    p_img1.add_run().add_picture("figuras/figura1_fluxo_processo.png", width=Inches(img_width))
    add_caption("Figura 1: Mapeamento do Fluxo Operacional na Vértice Confecções (Gargalo: Células de Costura).")
    add_source()

    add_body(
        "Estudos de cronoanálise identificaram as Células de Costura como o gargalo operacional do sistema (GOLDRATT, 2014). Todas as variantes foram agrupadas na família agregada 'Conjunto de Uniforme Profissional Padrão' (calça em brim + camisa polo bordada). O Tempo Padrão (TP) de montagem no gargalo é de 0,80 horas (48 minutos). A fábrica opera em jornada normal de 44h semanais, totalizando 160h líquidas mensais por operadora (descontadas as pausas ergonômicas da NR-17). Com 20 costureiras efetivas, a produtividade individual é de 200 conjuntos/mês (160h ÷ 0,80h), estabelecendo uma capacidade nominal de 4.000 conjuntos/mês."
    )
    add_body(
        "O horizonte adotado é de 12 meses (Janeiro a Dezembro), com demanda agregada de 48.000 conjuntos (média de 4.000 un/mês). A sazonalidade concentra forte demanda no 1º trimestre (pico escolar e frotas corporativas: Jan=4.600, Fev=5.200, Mar=4.800) e no fim do ano (Nov=5.000 un.), com vales intermediários (Jun=3.000 un.)."
    )

    # -----------------------------
    # SEÇÃO 3: APLICAÇÃO PRÁTICA DO PAP
    # -----------------------------
    add_title_1("3. APLICAÇÃO PRÁTICA DO PLANEJAMENTO AGREGADO DA PRODUÇÃO (PAP)")
    add_body(
        "A modelagem matemática segue as equações de balanço de materiais da disciplina (QUEIROZ, 2026): Produção Agregada: PT(t) = Normal(t) + HE(t) + SUB(t); Balanço de Estoques: EF(t) = EI(t) + PT(t) - D(t); Estoque Médio: EM(t) = [EI(t) + EF(t)] / 2; Demanda Efetiva: D_efetiva(t) = D(t) + Backlog(t-1); e Custos Totais: C_Total = C_Prod + C_HE + C_Sub + C_Estoque + C_Atraso + C_Admissão + C_Demissão."
    )

    # Tabela 1
    add_caption("Tabela 1: Parâmetros Operacionais, Custos Unitários e Previsão de Demanda Mensal.")
    t1 = doc.add_table(rows=14, cols=4)
    col_w_t1 = [Inches(1.8), Inches(1.0), Inches(1.2), Inches(2.4)]
    align_t1 = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
    for c_idx, val in enumerate(["Parâmetro / Variável", "Valor", "Unidade", "Descrição Técnica"]):
        t1.cell(0, c_idx).paragraphs[0].text = val
    dados_t1 = [
        ("Quadro Base de Operadoras (N0)", "20", "costureiras", "Força de trabalho efetiva inicial"),
        ("Jornada Mensal Líquida", "160", "horas/mês", "44h semanais com pausas da NR-17"),
        ("Tempo Padrão no Gargalo (TP)", "0,80", "h/conjunto", "Cronoanálise na costura (48 minutos)"),
        ("Produtividade Normal Unitária", "200", "conj./op./mês", "160 h ÷ 0,80 h/conjunto"),
        ("Capacidade Normal Base (Pn)", "4.000", "conjuntos/mês", "20 costureiras × 200 un/mês"),
        ("Estoque Inicial de Segurança (EI1)", "400", "conjuntos", "Estoque de segurança remanescente"),
        ("Custo Turno Normal (cn)", "R$ 38,00", "R$/conjunto", "Tecido, insumos, MOD direta e CIF"),
        ("Custo Turno Extra / HE (che)", "R$ 46,00", "R$/conjunto", "Adicional 50% na MOD + encargos"),
        ("Custo Subcontratação (csub)", "R$ 52,00", "R$/conjunto", "Facção externa de costura homologada"),
        ("Custo Posse de Estoque (ce)", "R$ 3,00", "R$/un./mês", "Armazenagem e custo de capital (sobre EM)"),
        ("Custo Ruptura / Atraso (ca)", "R$ 20,00", "R$/un./mês", "Multa contratual de SLA corporativo"),
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
        "A Estratégia 1 mantém a produção rigorosamente nivelada na média aritmética anual de 4.000 conjuntos/mês, com equipe fixa de 20 operadoras sem horas extras, contratações ou subcontratação. As variações de demanda são absorvidas exclusivamente por estoques ou atrasos (backlog) (Tabela 2)."
    )

    # Tabela 2
    add_caption("Tabela 2: Plano de Produção e Plano Financeiro da Estratégia 1 (Nivelamento).")
    t2 = doc.add_table(rows=14, cols=9)
    col_w_t2 = [Inches(0.6), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.9)]
    align_t2 = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT,
                WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]
    for c_idx, val in enumerate(["Mês", "Demanda", "D. Efetiva", "Prod.", "Est. Final", "Backlog", "Est. Médio", "C. Estoque", "Custo Total"]):
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
        "Sob a Estratégia 1, o choque inicial consome o estoque no 1º mês, acumulando 5 meses seguidos de atraso crítico (backlog de 2.200 un. em Março). A penalidade por atraso soma R$ 108.000,00, elevando o custo total para R$ 1.968.600,00 e comprometendo o nível de serviço em 20,8% dos períodos."
    )

    add_title_2("3.2. Estratégia 2: Acompanhamento da Demanda (Chase Strategy)")
    add_body(
        "A Estratégia 2 iguala mensalmente a produção à demanda (PT(t) = D(t)), contratando e demitindo operadoras para modular a capacidade (N(t) = D(t) ÷ 200). Mantém-se o estoque fixo no nível de segurança de 400 unidades (Tabela 3)."
    )

    # Tabela 3
    add_caption("Tabela 3: Plano de Produção e Plano Financeiro da Estratégia 2 (Acompanhamento / Chase).")
    t3 = doc.add_table(rows=14, cols=9)
    col_w_t3 = [Inches(0.6), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.8), Inches(0.8), Inches(0.9)]
    align_t3 = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT,
                WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]
    for c_idx, val in enumerate(["Mês", "Demanda", "Operadoras", "Admissões", "Demissões", "Produção", "C. Admissão", "C. Rescisão", "Custo Total"]):
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
        "Apesar do custo total aparente de R$ 1.921.200,00 e ausência de atrasos, a política impõe a contratação de 18 costureiras e a demissão de 18 colaboradoras em 12 meses (R$ 82.800,00 em despesas rescisórias). Essa instabilidade destrói o ritmo operacional, gerando custos ocultos de curva de aprendizado e risco jurídico trabalhista sob a CLT."
    )

    add_title_2("3.3. Estratégia 3: Variação em Patamares / Híbrida (Recomendada)")
    add_body(
        "A Estratégia 3 estabelece uma base estável de 19 operadoras fixas (capacidade normal de 3.800 un/mês), sem rotatividade. Nos picos de demanda (Janeiro a Março), recorre-se a Horas Extras (HE) dentro do teto legal de 2h/dia (Art. 59 CLT = até 800 un/mês a R$ 46,00/un) e subcontratação com facções homologadas (R$ 52,00/un). Nos meses intermediários, o saldo suave estocado absorve integralmente o pico de Novembro sem necessidade de horas extras (Tabela 4)."
    )

    # Tabela 4
    add_caption("Tabela 4: Plano de Produção e Plano Financeiro da Estratégia 3 (Híbrida em Patamares - Recomendada).")
    t4 = doc.add_table(rows=14, cols=10)
    col_w_t4 = [Inches(0.55), Inches(0.65), Inches(0.65), Inches(0.65), Inches(0.65), Inches(0.65), Inches(0.65), Inches(0.65), Inches(0.75), Inches(0.95)]
    align_t4 = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT,
                WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]
    for c_idx, val in enumerate(["Mês", "Demanda", "Normal", "HE (un)", "Sub (un)", "Prod. Total", "Est. Final", "Est. Médio", "C. Estoque", "Custo Total"]):
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

    p_img2 = doc.add_paragraph()
    p_img2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img2.paragraph_format.space_after = Pt(1)
    p_img2.add_run().add_picture("figuras/figura2_demanda_vs_producao.png", width=Inches(img_width))
    add_caption("Figura 2: Demanda Mensal versus Taxas de Produção por Estratégia de PAP.")
    add_source()

    # -----------------------------
    # SEÇÃO 3.4: COMPARATIVO DO PLANO FINANCEIRO
    # -----------------------------
    add_title_2("3.4. Síntese e Comparativo Consolidado do Plano Financeiro")
    add_body(
        "A Tabela 5 consolida os custos anuais totais por rubrica e confronta os principais indicadores de desempenho (KPIs) de manufatura entre as alternativas."
    )

    # Tabela 5
    add_caption("Tabela 5: Comparativo Consolidado do Plano Financeiro e Indicadores de Desempenho (12 Meses).")
    t5 = doc.add_table(rows=15, cols=4)
    col_w_t5 = [Inches(2.5), Inches(1.3), Inches(1.3), Inches(1.5)]
    align_t5 = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]
    for c_idx, val in enumerate(["Rubrica de Custo / Indicador (KPI)", "E1: Nivelamento", "E2: Acompanhamento", "E3: Híbrida (Recomendada)"]):
        t5.cell(0, c_idx).paragraphs[0].text = val
    dados_t5 = [
        ("Produção em Turno Normal (cn = R$ 38,00)", "R$ 1.824.000,00", "R$ 1.824.000,00", "R$ 1.732.800,00"),
        ("Turno Extra / Horas Extras (che = R$ 46,00)", "R$ 0,00", "R$ 0,00", "R$ 92.000,00"),
        ("Subcontratação / Facção Externa (csub = R$ 52,00)", "R$ 0,00", "R$ 0,00", "R$ 41.600,00"),
        ("Manutenção / Posse de Estoque (ce = R$ 3,00)", "R$ 36.600,00", "R$ 14.400,00", "R$ 36.000,00"),
        ("Custos de Atraso na Entrega / Backlog (ca = R$ 20,00)", "R$ 108.000,00", "R$ 0,00", "R$ 0,00"),
        ("Custos de Admissão de Pessoal (Cc = R$ 1.600,00)", "R$ 0,00", "R$ 28.800,00", "R$ 0,00"),
        ("Custos de Desligamento / Rescisões (Cd = R$ 3.000,00)", "R$ 0,00", "R$ 54.000,00", "R$ 0,00"),
        ("CUSTO TOTAL ANUAL DO PLANO (R$)", "R$ 1.968.600,00", "R$ 1.921.200,00", "R$ 1.902.400,00"),
        ("Economia Absoluta em Relação ao Nivelamento", "Referência", "R$ 47.400,00", "R$ 66.200,00"),
        ("Nível de Serviço / Entregas no Prazo (%)", "79,2% (5 meses em atraso)", "100,0%", "100,0%"),
        ("Volume de Backlog Acumulado no Ano", "6.600 un. acumuladas", "0 un.", "0 un."),
        ("Movimentações Trabalhistas (Admissões + Demissões)", "0", "36 operadoras", "0"),
        ("Estoque Final Remanescente no Mês 12", "400 conjuntos", "400 conjuntos", "800 conjuntos"),
        ("Conformidade Legal com a CLT (Art. 59)", "100% Regular", "100% Regular", "100% Regular (≤ 2h/dia)")
    ]
    for r_idx, row in enumerate(dados_t5, start=1):
        for c_idx, val in enumerate(row):
            t5.cell(r_idx, c_idx).paragraphs[0].text = val
    style_table(t5, col_w_t5, align_t5)
    add_source()

    p_img4 = doc.add_paragraph()
    p_img4.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img4.paragraph_format.space_after = Pt(1)
    p_img4.add_run().add_picture("figuras/figura4_comparativo_custos.png", width=Inches(img_width * 0.95))
    add_caption("Figura 4: Estrutura Comparativa de Custos Totais Anuais do PAP.")
    add_source()

    # -----------------------------
    # SEÇÃO 4: ANÁLISE CRÍTICA DOS RESULTADOS
    # -----------------------------
    add_title_1("4. ANÁLISE CRÍTICA DOS RESULTADOS E DISCUSSÃO")
    add_body(
        "A análise comparativa comprova a superioridade econômico-operacional da Estratégia 3 (Híbrida em Patamares), que alcançou o menor custo total anual (R$ 1.902.400,00), gerando uma economia de R$ 66.200,00 frente ao Nivelamento (-3,36%) e de R$ 18.800,00 frente ao Acompanhamento (-0,98%). Contudo, na engenharia de operações, a avaliação deve transcender o critério financeiro, pesando os objetivos de desempenho de Slack (custo, rapidez, confiabilidade, flexibilidade e qualidade) e os qualificadores e ganhadores de pedidos de Terry Hill (HILL, 2000)."
    )

    p_img3 = doc.add_paragraph()
    p_img3.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img3.paragraph_format.space_after = Pt(1)
    p_img3.add_run().add_picture("figuras/figura3_estoques_e_backlog.png", width=Inches(img_width))
    add_caption("Figura 3: Comparação da Dinâmica de Estoques e Rupturas (E1 vs. E3).")
    add_source()

    add_title_2("4.1. Confronto das Estratégias e Análise de Trade-offs")
    add_body(
        "A Estratégia 1 (Nivelamento) mostrou-se operacionalmente inaceitável. Ao ignorar a concentração da demanda no 1º trimestre, esgota o estoque de segurança e gera 5 meses de backlog ininterrupto (Figura 3a). No setor de uniformes, onde contratos empresariais preveem multas diárias e cláusulas rescisórias, o atraso destruiria a confiabilidade da marca e acarretaria perda de clientes estratégicos."
    )
    add_body(
        "A Estratégia 2 (Acompanhamento) apresenta custos contábeis atraentes, mas é inviável na prática de confecção. Contratar 18 costureiras e demitir 18 operadoras em 12 meses destrói a curva de aprendizagem fabril. Operadoras recém-admitidas têm tempo de operação até 40% superior ao padrão nos primeiros meses, gerando costuras tortas, refugo de tecido e retrabalhos. O turnover extremo degrada o clima organizacional e expõe a empresa a passivos trabalhistas sob a CLT."
    )
    add_title_2("4.2. Restrições Trabalhistas (CLT), Gestão da Qualidade e Terceirização")
    add_body(
        "A Estratégia 3 respeita integralmente a CLT (Artigo 59), limitando as horas extras a 2 horas diárias (25% da jornada = até 800 un/mês), aplicadas somente nos picos críticos de Janeiro a Março, prevenindo a fadiga física crônica das costureiras. Para as 800 unidades subcontratadas com facções de costura em Fevereiro e Março, a Vértice fornece os tecidos já enfestados e cortados, retendo o controle de corte e estamparia e auditando a costura externa, o que mitiga riscos de qualidade e responsabilidade solidária/subsidiária (Súmula 331 do TST)."
    )
    add_title_2("4.3. Limitações e Oportunidades de Melhoria Contínua")
    add_body(
        "Dentre as limitações da modelagem, destacam-se a premissa de demanda determinística e a agregação em um produto padrão único, desconsiderando setups no corte. Recomenda-se à diretoria: (a) estruturar reuniões mensais de S&OP (Sales and Operations Planning) com horizonte rolante; (b) implementar práticas de Lean Manufacturing (troca rápida no corte e balanceamento das células de costura); (c) homologar contratos formais de parceria com facções locais; e (d) adotar gestão visual e PCP digital (Kanban de ordens de serviço)."
    )

    # -----------------------------
    # SEÇÃO 5: CONCLUSÃO
    # -----------------------------
    add_title_1("5. CONCLUSÃO")
    add_body(
        "O presente trabalho evidenciou a relevância tática do Planejamento Agregado da Produção na harmonização entre demanda de mercado e capacidade industrial. A aplicação ao caso da Vértice Confecções Ltda. demonstrou os perigos de estratégias extremas: o Nivelamento colapsou a entrega com R$ 108.000,00 em multas por atraso, enquanto o Acompanhamento impôs uma rotatividade destrutiva de 36 movimentações de pessoal (R$ 82.800,00 em custos rescisórios e severa perda de qualidade)."
    )
    add_body(
        "A Estratégia 3 (Híbrida em Patamares) consagrou-se como a melhor solução em todas as dimensões de desempenho: alcançou o MENOR CUSTO TOTAL (R$ 1.902.400,00, economia de R$ 66.200,00 frente a E1 e R$ 18.800,00 frente a E2), garantiu 100% DE NÍVEL DE SERVIÇO (zero rupturas contratuais), manteve a estabilidade de 19 operadoras permanentes, operou em plena conformidade com a CLT e entregou um saldo saudável de 800 conjuntos estocados no Mês 12 para inaugurar o ano seguinte. A estratégia consolida-se como o modelo de referência tática para a indústria de confecção."
    )

    # -----------------------------
    # SEÇÃO 6: REFERÊNCIAS
    # -----------------------------
    add_title_1("6. REFERÊNCIAS BIBLIOGRÁFICAS")
    referencias = [
        "ABIT – ASSOCIAÇÃO BRASILEIRA DA INDÚSTRIA TÊXTIL E DE CONFECÇÃO. Perfil do Setor Têxtil Brasileiro: Inteligência Setorial. São Paulo: ABIT, 2024. Disponível em: <https://www.abit.org.br>. Acesso em: 28 set. 2026.",
        "BARBOSA, C. P.; GONÇALVES, M. C. Organização Industrial – Aula 02: Planejamento Agregado da Produção em Serviços. Brasília: Universidade de Brasília (UnB), Faculdade de Tecnologia, 2026. 43 slides.",
        "BRASIL. Decreto-Lei nº 5.452, de 1º de maio de 1943. Consolidação das Leis do Trabalho (CLT). Artigo 59 (Duração do Trabalho e Horas Extras). Brasília, DF: Presidência da República, 1943.",
        "CORRÊA, H. L.; CORRÊA, C. A. Administração de Produção e Operações: Manufatura e Serviços – Uma Abordagem Estratégica. 4. ed. São Paulo: Atlas, 2017.",
        "DAVIS, M. M.; AQUILANO, N. J.; CHASE, R. B. Fundamentos da Administração da Produção. 3. ed. Porto Alegre: Bookman, 2001.",
        "FITZSIMMONS, J. A.; FITZSIMMONS, M. J.; BORDOLOI, S. K. Administração de Serviços: Operações, Estratégia e Tecnologia da Informação. 7. ed. Porto Alegre: AMGH, 2014.",
        "GOLDRATT, E. M.; COX, J. A Meta: Um Processo de Aprimoramento Contínuo. 3. ed. São Paulo: Nobel, 2014.",
        "GONÇALVES, M. C. Conceitos Básicos de Gestão da Produção e Estratégia de Operações – Aula 01. Brasília: UnB / FT, 2026. 29 p.",
        "HILL, T. Manufacturing Strategy: Text and Cases. 3. ed. Boston: McGraw-Hill, 2000.",
        "KRAJEWSKI, L. J.; RITZMAN, L. P.; MALHOTRA, M. K. Administração de Produção e Operações. 8. ed. São Paulo: Pearson, 2014.",
        "MOREIRA, D. A. Administração da Produção e Operações. 2. ed. São Paulo: Cengage Learning, 2012.",
        "QUEIROZ, J. V. Organização Industrial – Aula 03: Planejamento Agregado da Produção em Manufatura (PAP, S&OP e Capacidade). Brasília: UnB / FT, 2026. 26 slides.",
        "QUEIROZ, J. V. Modelo Computacional em Python para PAP em Manufatura. Google Colab, 2026. Disponível em: <https://colab.research.google.com/drive/1UNxeDKP04DMtyab0cuSxG4N4JOY3XSgv>. Acesso em: 29 set. 2026.",
        "SEBRAE – SERVIÇO BRASILEIRO DE APOIO ÀS MICRO E PEQUENAS EMPRESAS. Relatório de Inteligência: Desafios da Gestão de Produção e PCP nas MPMEs de Vestuário. Brasília: SEBRAE, 2024.",
        "SLACK, N.; BRANDON-JONES, A.; JOHNSTON, R. Administração da Produção. 8. ed. São Paulo: Atlas, 2018.",
        "TUBINO, D. F. Planejamento e Controle da Produção: Teoria e Prática. 3. ed. São Paulo: Atlas, 2017."
    ]
    for ref in referencias:
        p_ref = doc.add_paragraph()
        p_ref.paragraph_format.line_spacing = 1.10
        p_ref.paragraph_format.space_after = Pt(3.5)
        p_ref.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run_ref = p_ref.add_run(ref)
        run_ref.font.name = 'Arial'
        run_ref.font.size = Pt(8.5)

    doc.save("Relatorio_Tecnico_1_PAP_Manufatura.docx")
    
    # Converter para PDF
    cmd = "/opt/homebrew/bin/soffice --headless --convert-to pdf Relatorio_Tecnico_1_PAP_Manufatura.docx"
    subprocess.run(cmd, shell=True, check=True)
    
    reader = pypdf.PdfReader("Relatorio_Tecnico_1_PAP_Manufatura.pdf")
    n_pages = len(reader.pages)
    print(f"Resultado: {n_pages} páginas geradas.")
    return n_pages

generate_doc()
