import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import subprocess
import fitz

def build():
    doc = docx.Document()
    
    for s in doc.sections:
        s.top_margin = Inches(0.80)      # 2.0 cm
        s.bottom_margin = Inches(0.80)   # 2.0 cm
        s.left_margin = Inches(0.95)     # 2.4 cm
        s.right_margin = Inches(0.95)    # 2.4 cm
        s.page_width = Inches(8.27)
        s.page_height = Inches(11.69)
        
    NAVY = RGBColor(27, 54, 93)
    DARK_GRAY = RGBColor(55, 65, 81)
    BLACK = RGBColor(0, 0, 0)
    
    font_body = 9.5
    line_spacing = 1.15
    
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Arial'
    normal_style.font.size = Pt(font_body)
    normal_style.font.color.rgb = BLACK
    normal_style.paragraph_format.line_spacing = line_spacing
    normal_style.paragraph_format.space_after = Pt(2.5)
    normal_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    def add_title_1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(11.5)
        run.font.bold = True
        run.font.color.rgb = NAVY
        return p

    def add_title_2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = NAVY
        return p

    def add_body(text, space_after=2.5, bold_prefix=None):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = line_spacing
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.first_line_indent = Inches(0.35)
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
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(8)
        run.font.bold = True
        run.font.color.rgb = DARK_GRAY
        return p

    def add_source(text="Fonte: Elaborado pelos autores (2026)."):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(7.5)
        run.font.italic = True
        run.font.color.rgb = DARK_GRAY
        return p

    def set_cell_shading(cell, color_hex):
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
        cell._tc.get_or_add_tcPr().append(shading)

    def set_cell_margins(cell, top=20, bottom=20, left=40, right=40):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)

    def style_table(table, col_widths, align_cols=None, font_sz=6.8):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for r_idx, row in enumerate(table.rows):
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            for c_idx, cell in enumerate(row.cells):
                cell.width = col_widths[c_idx]
                set_cell_margins(cell)
                p = cell.paragraphs[0]
                p.paragraph_format.line_spacing = 1.0
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.space_before = Pt(0)
                align = align_cols[c_idx] if align_cols else WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.alignment = align
                for run in p.runs:
                    run.font.name = 'Arial'
                    if r_idx == 0:
                        run.font.size = Pt(font_sz)
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(255, 255, 255)
                    elif r_idx == len(table.rows) - 1:
                        run.font.size = Pt(font_sz)
                        run.font.bold = True
                    else:
                        run.font.size = Pt(font_sz)
                if r_idx == 0:
                    set_cell_shading(cell, "1B365D")
                elif r_idx == len(table.rows) - 1:
                    set_cell_shading(cell, "DCE6F1")
                elif r_idx % 2 == 1:
                    set_cell_shading(cell, "F8FAFC")

    # =========================================================
    # PÁGINA 1: CAPA
    # =========================================================
    p_h = doc.add_paragraph()
    p_h.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_h.paragraph_format.space_after = Pt(4)
    r = p_h.add_run("UNIVERSIDADE DE BRASÍLIA – UnB\nFACULDADE DE TECNOLOGIA – FT\nDEPARTAMENTO DE ENGENHARIA DE PRODUÇÃO (EPR)\nDISCIPLINA: ORGANIZAÇÃO INDUSTRIAL (2026.2)")
    r.font.name = 'Arial'
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = NAVY

    for _ in range(4):
        doc.add_paragraph().paragraph_format.space_after = Pt(10)

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

    for _ in range(4):
        doc.add_paragraph().paragraph_format.space_after = Pt(10)

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

    for _ in range(3):
        doc.add_paragraph().paragraph_format.space_after = Pt(10)

    p_f = doc.add_paragraph()
    p_f.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_f = p_f.add_run("BRASÍLIA – DF\nOUTUBRO DE 2026")
    r_f.font.name = 'Arial'
    r_f.font.size = Pt(10)
    r_f.font.bold = True

    doc.add_page_break()

    # =========================================================
    # PÁGINA 2: INTRODUÇÃO E OBJETIVOS
    # =========================================================
    add_title_1("1. INTRODUÇÃO E OBJETIVOS DA ANÁLISE")
    add_body(
        "No ambiente concorrencial da manufatura contemporânea, a tomada de decisões na gestão de operações industriais exige uma articulação rigorosa entre as diretrizes estratégicas corporativas e a capacidade física instalada do chão de fábrica. O Planejamento Agregado da Produção (PAP) consolida-se como o elemento central do nível tático do sistema de Planejamento e Controle da Produção (PCP), conectando o plano de negócios e o processo de S&OP (Sales and Operations Planning) aos programas mestres detalhados de produção (SLACK et al., 2018; CORRÊA; CORRÊA, 2017). O PAP visa estabelecer para um horizonte de médio prazo (6 a 18 meses) as taxas globais de produção, níveis de estoques, força de trabalho, horas extras e subcontratações, harmonizando a demanda agregada prevista com a capacidade produtiva instalada."
    )
    add_body(
        "Diferentemente das operações no setor de serviços — cuja simultaneidade e intangibilidade impedem a estocagem do produto final acabado (FITZSIMMONS et al., 2014; BARBOSA; GONÇALVES, 2026) —, as empresas de manufatura dispõem de múltiplos graus de liberdade para amortecer flutuações sazonais de demanda. Essas opções agrupam-se em três abordagens fundamentais: (i) Estratégia de Nivelamento (taxa de produção constante com absorção das oscilações mediante estoques de produtos acabados ou atrasos planejados); (ii) Estratégia de Acompanhamento (ajuste contínuo da força de trabalho via admissões e demissões sazonais); e (iii) Estratégia Híbrida em Patamares (capacidade própria regular combinada com horas extras e terceirização nos picos de demanda) (KRAJEWSKI et al., 2014; QUEIROZ, 2026)."
    )
    add_body(
        "A aplicação do PAP torna-se ainda mais crítica no segmento de Micro, Pequenas e Médias Empresas (MPMEs) de confecção, que respondem por parcela vital do emprego industrial no Brasil, mas frequentemente operam de maneira reativa, sem ferramentas analíticas formais de previsão e planejamento tático (SEBRAE, 2024). Em setores marcados por forte sazonalidade anual, a ausência de planejamento agregado acarreta gargalos crônicos, gastos vultosos com horas extras imprevistas, atrasos severos nas entregas de contratos corporativos ou, no extremo oposto, capital de giro excessivamente imobilizado em estoques."
    )
    add_title_2("1.1. Objetivos do Estudo")
    add_body(
        "Este relatório técnico tem por objetivo modelar matematicamente, mensurar custos e avaliar comparativamente planos agregados de produção para um caso real da manufatura de vestuário profissional: a empresa Vértice Confecções & Uniformes Ltda., sediada no Polo Industrial de Taguatinga - DF. Especificamente, busca-se:",
        bold_prefix="Objetivo Geral: "
    )
    add_body("1. Mapear o sistema produtivo da empresa, caracterizando o ambiente de manufatura (MTS/MTO), o ponto de desacoplamento, a estrutura de processos e o gargalo fabril;", space_after=2)
    add_body("2. Consolidar os dados empíricos de previsão de demanda e parâmetros de custos operacionais para um horizonte completo de 12 meses;", space_after=2)
    add_body("3. Formular e calcular com rigor matemático três estratégias clássicas de PAP: Nivelamento (Taxa Constante), Acompanhamento da Demanda (Chase) e Estratégia Híbrida em Patamares;", space_after=2)
    add_body("4. Confrontar as estratégias sob as óticas financeira (Plano Financeiro de custos totais) e operacional (nível de serviço, estabilidade do quadro, conformidade com a CLT e qualidade);", space_after=2)
    add_body("5. Propor a solução de maior viabilidade técnico-econômica, delineando recomendações práticas de implementação para a diretoria industrial.", space_after=3)

    doc.add_page_break()

    # =========================================================
    # PÁGINA 3: CARACTERIZAÇÃO DO CASO E PESQUISA
    # =========================================================
    add_title_1("2. CARACTERIZAÇÃO DO CASO E METODOLOGIA DE PESQUISA")
    add_body(
        "A unidade de estudo selecionada para esta análise é a Vértice Confecções & Uniformes Profissionais Ltda. - EPP, indústria de manufatura fundada em 2012 e sediada no Setor Industrial de Taguatinga (Polo Industrial do Distrito Federal). A empresa atua na confecção de uniformes profissionais corporativos (calças em brim pesado, camisas polo de alta durabilidade e jalecos industriais) e uniformes escolares padronizados para colégios privados da região Centro-Oeste."
    )
    add_body(
        "O sistema produtivo adota ambiente híbrido Make-to-Stock (MTS / Produção Empurrada) e Assemble-to-Order (ATO / Produção Puxada) (SLACK et al., 2018). Para calças operacionais e camisas polo básicas de alta rotatividade, adota-se produção antecipada; os bordados, logomarcas e detalhamentos específicos de cada cliente corporativo são customizados sob pedido firme, definindo o ponto de desacoplamento de pedidos (Figura 1)."
    )
    
    p_img1 = doc.add_paragraph()
    p_img1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img1.paragraph_format.space_after = Pt(1)
    p_img1.add_run().add_picture("figuras/figura1_fluxo_processo.png", width=Inches(5.5))
    add_caption("Figura 1: Mapeamento do Fluxo Operacional na Vértice Confecções (Gargalo: Células de Costura).")
    add_source()

    add_body(
        "Por meio de estudos de cronoanálise industrial, a gerência fabril identificou que o setor de Costura configura o gargalo do sistema produtivo (Teoria das Restrições - GOLDRATT, 2014). Para viabilizar a modelagem do PAP, todas as variantes foram agrupadas na família agregada representativa 'Conjunto de Uniforme Profissional Padrão' (composto por uma calça operacional em brim pesado 100% algodão e uma camisa polo mista com bordado corporativo). O Tempo Padrão (TP) apurado para a montagem de um conjunto completo no gargalo é de 0,80 horas (48 minutos)."
    )
    add_body(
        "A coleta de dados combinou análise documental interna da fábrica e parâmetros setoriais consolidados pelo SEBRAE (2024), pela Associação Brasileira da Indústria Têxtil e de Confecção (ABIT) e pelo Sindicato das Indústrias do Vestuário do DF (SINDIVEST-DF). A jornada normal de trabalho é de 44 horas semanais, totalizando 160 horas produtivas mensais líquidas por costureira (descontadas pausas ergonômicas da NR-17 e tolerâncias de fadiga). A produtividade em turno normal é de: Produtividade = 160 h/mês ÷ 0,80 h/conjunto = 200 conjuntos/mês por operadora. Com o quadro base de 20 colaboradoras, a capacidade nominal normal instalada é de 4.000 conjuntos/mês."
    )
    add_body(
        "O horizonte de planejamento adotado abrange 12 meses (Janeiro a Dezembro), com demanda agregada anual prevista de 48.000 conjuntos (média de 4.000 un/mês). A sazonalidade revela dois picos expressivos: no 1º trimestre (Janeiro a Março: Jan=4.600, Fev=5.200, Mar=4.800), decorrente de novos contratos corporativos e volta às aulas; e no fim de ano (Novembro: 5.000 un.), associado a compras do comércio e hotelaria para eventos de final de ano."
    )

    doc.add_page_break()

    # =========================================================
    # PÁGINA 4: APLICAÇÃO PRÁTICA DO PAP - PARÂMETROS E E1
    # =========================================================
    add_title_1("3. APLICAÇÃO PRÁTICA DO PLANEJAMENTO AGREGADO DA PRODUÇÃO (PAP)")
    add_body(
        "A modelagem matemática segue rigorosamente as relações teóricas lecionadas no curso (QUEIROZ, 2026): (i) Produção Total: PT(t) = Normal(t) + HE(t) + SUB(t); (ii) Balanço Material de Estoques: EF(t) = EI(t) + PT(t) - D(t); (iii) Estoque Médio Mensal: EM(t) = [EI(t) + EF(t)] / 2; (iv) Demanda Efetiva com Backlog: D_efetiva(t) = D(t) + Backlog(t-1); e (v) Plano Financeiro: C_Total = C_Prod + C_HE + C_Sub + C_Estoque + C_Atraso + C_Admissão + C_Demissão."
    )

    # Tabela 1
    add_caption("Tabela 1: Parâmetros Operacionais, Estrutura de Custos e Previsão de Demanda Mensal.")
    t1 = doc.add_table(rows=14, cols=4)
    col_w_t1 = [Inches(1.8), Inches(0.9), Inches(1.2), Inches(2.5)]
    align_t1 = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
    for c_idx, val in enumerate(["Parâmetro / Variável", "Valor", "Unidade", "Descrição Técnica / Base Normativa"]):
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
    style_table(t1, col_w_t1, align_t1, font_sz=6.8)
    add_source()

    add_title_2("3.1. Estratégia 1: Produção Constante (Nivelamento / Level Strategy)")
    add_body(
        "A Estratégia 1 estabelece capacidade constante em 4.000 conjuntos/mês (equipe fixa de 20 operadoras), sem horas extras, contratações ou terceirização. As variações de demanda são absorvidas exclusivamente por estoques ou atrasos (backlog) (Tabela 2)."
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
    style_table(t2, col_w_t2, align_t2, font_sz=6.8)
    add_source()

    add_body(
        "Sob a Estratégia 1, o pico inicial consome todo o estoque em Janeiro, provocando 5 meses seguidos de atraso crítico (backlog de até 2.200 conjuntos em Março). A penalidade contratual acumula R$ 108.000,00, elevando o custo total para R$ 1.968.600,00 e comprometendo o nível de serviço em 20,8% dos meses."
    )

    doc.add_page_break()

    # =========================================================
    # PÁGINA 5: ESTRATÉGIA 2 E ESTRATÉGIA 3
    # =========================================================
    add_title_2("3.2. Estratégia 2: Acompanhamento da Demanda (Chase Strategy)")
    add_body(
        "A Estratégia 2 iguala rigorosamente a produção à demanda mensal (PT(t) = D(t)), contratando e demitindo operadoras para modular a capacidade (N(t) = D(t) ÷ 200). Mantém-se o estoque constante no nível de segurança de 400 conjuntos (Tabela 3)."
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
    style_table(t3, col_w_t3, align_t3, font_sz=6.8)
    add_source()

    add_body(
        "Apesar do custo aparente de R$ 1.921.200,00 e zero atrasos, a política impõe a contratação de 18 costureiras e demissão de 18 colaboradoras em 12 meses (R$ 82.800,00 em despesas rescisórias). Essa instabilidade contínua reinicia a curva de experiência, degrada a qualidade da costura e eleva exponencialmente o risco trabalhista sob a CLT."
    )

    add_title_2("3.3. Estratégia 3: Variação em Patamares / Híbrida (Recomendada)")
    add_body(
        "A Estratégia 3 adota equipe estável de 19 costureiras fixas (produção normal de 3.800 un/mês), sem admissões ou demissões. Nos picos de demanda (Janeiro a Março), recorre-se a Horas Extras (HE) dentro do teto legal de 2h/dia (Art. 59 da CLT = até 800 un/mês a R$ 46,00/un) e subcontratação com facções homologadas (R$ 52,00/un). Nos meses intermediários, o estoque regulador suave absorve com folga o pico de Novembro com custo normal (Tabela 4)."
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
    style_table(t4, col_w_t4, align_t4, font_sz=6.8)
    add_source()

    doc.add_page_break()

    # =========================================================
    # PÁGINA 6: COMPARATIVO DO PLANO FINANCEIRO E FIGURA 2
    # =========================================================
    add_title_2("3.4. Síntese e Comparativo Consolidado do Plano Financeiro")
    add_body(
        "A Tabela 5 consolida os custos anuais totais por rubrica e confronta os principais indicadores de desempenho (KPIs) de manufatura entre as alternativas avaliadas, fornecendo a base técnico-econômica para a tomada de decisão."
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
    style_table(t5, col_w_t5, align_t5, font_sz=6.8)
    add_source()

    p_img2 = doc.add_paragraph()
    p_img2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img2.paragraph_format.space_after = Pt(1)
    p_img2.add_run().add_picture("figuras/figura2_demanda_vs_producao.png", width=Inches(5.0))
    add_caption("Figura 2: Demanda Mensal versus Taxas de Produção por Estratégia de PAP.")
    add_source()

    doc.add_page_break()

    # =========================================================
    # PÁGINA 7: ANÁLISE CRÍTICA, FIGURA 4 E TRADE-OFFS
    # =========================================================
    add_title_1("4. ANÁLISE CRÍTICA DOS RESULTADOS E DISCUSSÃO")
    add_body(
        "A análise comparativa comprova a superioridade econômico-operacional da Estratégia 3 (Híbrida em Patamares), que alcançou o menor custo total anual (R$ 1.902.400,00), gerando uma economia líquida de R$ 66.200,00 frente ao Nivelamento (-3,36%) e de R$ 18.800,00 frente ao Acompanhamento (-0,98%). Contudo, na engenharia de operações, a avaliação deve transcender o critério financeiro, pesando os objetivos de desempenho de Slack (custo, rapidez, confiabilidade, flexibilidade e qualidade) e os qualificadores e ganhadores de pedidos de Terry Hill (HILL, 2000)."
    )

    p_img4 = doc.add_paragraph()
    p_img4.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img4.paragraph_format.space_after = Pt(1)
    p_img4.add_run().add_picture("figuras/figura4_comparativo_custos.png", width=Inches(4.8))
    add_caption("Figura 4: Estrutura Comparativa de Custos Totais Anuais do PAP.")
    add_source()

    add_title_2("4.1. Confronto das Estratégias e Análise de Trade-offs")
    add_body(
        "A Estratégia 1 (Nivelamento) mostrou-se operacionalmente inaceitável. Ao ignorar a concentração da demanda no 1º trimestre, esgota o estoque de segurança e gera 5 meses consecutivos de backlog ininterrupto. No setor de uniformes, onde contratos empresariais preveem multas diárias severas e cláusulas rescisórias, o atraso destruiria a confiabilidade da marca e acarretaria a perda definitiva de clientes estratégicos."
    )
    add_body(
        "A Estratégia 2 (Acompanhamento) apresenta custos contábeis atraentes, mas sua viabilidade fabril é nula. A manufatura de vestuário depende intensamente de coordenação motora fina em células de costura. Contratar 18 costureiras e demitir 18 operadoras em 12 meses reinicia constantemente a curva de experiência. Operadoras recém-admitidas têm tempo de operação até 40% superior ao padrão nos primeiros meses, gerando costuras tortas, refugo de tecido e retrabalhos. O turnover extremo degrada o clima organizacional e expõe a empresa a pesados passivos trabalhistas sob a CLT."
    )

    doc.add_page_break()

    # =========================================================
    # PÁGINA 8: FIGURA 3, CLT, QUALIDADE E LIMITAÇÕES
    # =========================================================
    p_img3 = doc.add_paragraph()
    p_img3.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img3.paragraph_format.space_after = Pt(1)
    p_img3.add_run().add_picture("figuras/figura3_estoques_e_backlog.png", width=Inches(5.2))
    add_caption("Figura 3: Comparação da Dinâmica de Estoques e Rupturas (E1 vs. E3).")
    add_source()

    add_title_2("4.2. Restrições Trabalhistas (CLT), Gestão da Qualidade e Terceirização")
    add_body(
        "A Estratégia 3 respeita integralmente o marco legal da CLT (Artigo 59), limitando as horas extras a 2 horas diárias (25% da jornada = até 800 un/mês), aplicadas somente nos picos críticos de Janeiro a Março, prevenindo a fadiga física crônica das costureiras e doenças osteomusculares (DORT/LER). Para as 800 unidades subcontratadas com facções de costura em Fevereiro e Março, a Vértice fornece os tecidos já enfestados e cortados, retendo o controle de corte e estamparia e auditando a costura externa, o que mitiga riscos de qualidade e responsabilidade solidária/subsidiária (Súmula 331 do TST)."
    )
    add_title_2("4.3. Limitações do Modelo e Oportunidades de Melhoria Contínua")
    add_body(
        "Dentre as limitações da modelagem, destacam-se a premissa de demanda determinística e a agregação em um produto padrão único, desconsiderando setups no corte. Recomenda-se à diretoria industrial a implementação das seguintes ações de melhoria:",
        bold_prefix="Plano de Ação e Recomendações: "
    )
    add_body("• Estruturação Formal do Processo de S&OP (Sales and Operations Planning): Instituir reuniões mensais interdisciplinares entre Comercial, PCP e Finanças com horizonte rolante de 12 meses (SLACK et al., 2018);", space_after=1.5)
    add_body("• Aplicação de Ferramentas Lean Manufacturing: Implementar troca rápida de ferramentas (SMED) na mesa de corte e balanceamento de linha nas células de costura, eliminando desperdícios e elevando a produtividade natural do chão de fábrica;", space_after=1.5)
    add_body("• Homologação Estratégica de Facções de Costura: Formalizar contratos de parceria técnica com oficinas de costura do Polo de Taguatinga, garantindo padrões uniformes de costura e pontualidade na entrega nos meses de Fevereiro e Março;", space_after=1.5)
    add_body("• Gestão Visual e PCP Digital: Substituir o controle manual em planilhas por painéis visuais de acompanhamento diário (Kanban de ordens de serviço), monitorando em tempo real os tempos de ciclo e o cumprimento do plano agregado.", space_after=3)

    doc.add_page_break()

    # =========================================================
    # PÁGINA 9: 5. CONCLUSÃO
    # =========================================================
    add_title_1("5. CONCLUSÃO E DIRETRIZES FINAIS")
    add_body(
        "O presente trabalho evidenciou a relevância tática do Planejamento Agregado da Produção na harmonização entre demanda de mercado e capacidade industrial. A aplicação prática ao caso real da Vértice Confecções Ltda. demonstrou com clareza os perigos das estratégias extremas: o Nivelamento Puro gerou um colapso operacional com R$ 108.000,00 em multas por atraso (backlog de até 2.200 conjuntos), enquanto o Acompanhamento Puro impôs uma rotatividade destrutiva de 36 movimentações de pessoal (R$ 82.800,00 em custos rescisórios e severa perda de rendimento fabril)."
    )
    add_body(
        "A solução ótima recomendada — Estratégia 3 (Híbrida em Patamares) — consagrou-se como a melhor alternativa em todas as dimensões de desempenho: alcançou o MENOR CUSTO TOTAL (R$ 1.902.400,00, gerando economia de R$ 66.200,00 frente a E1 e R$ 18.800,00 frente a E2), garantiu 100% DE NÍVEL DE SERVIÇO (zero rupturas contratuais), manteve a estabilidade de 19 operadoras permanentes, operou em plena conformidade com a CLT e entregou um saldo saudável de 800 conjuntos estocados no Mês 12 para inaugurar o ano seguinte."
    )
    add_body(
        "A análise comprova que o PAP não é um exercício meramente matemático ou contábil, mas uma ferramenta estratégica integradora que alinha vendas, recursos humanos, finanças e engenharia de operações. A conciliação balanceada entre capacidade própria, flexibilidade legal de horas extras e terceirização controlada consolida-se como o modelo de excelência tática para a sustentabilidade e competitividade da indústria de confecção."
    )

    add_title_2("5.1. Quadro Síntese de Recomendações para a Diretoria")
    add_body(
        "Para a execução imediata das diretrizes estabelecidas no Plano de Produção Recomendado, sintetizam-se as quatro ações prioritárias:",
        bold_prefix="Diretrizes de Implementação: "
    )
    add_body("1. Fixação do Quadro de Costura em 19 Operadoras: Manter o quadro efetivo estável, eliminando contratações e demissões precipitadas ao longo do exercício fiscal;", space_after=1.5)
    add_body("2. Contratação Prévia de Horas Extras para o 1º Trimestre: Pactuar com a equipe e o sindicato a realização de até 2h extras/dia exclusivamente em Janeiro, Fevereiro e Março;", space_after=1.5)
    add_body("3. Auditoria e Contrato com Facções de Costura Homologadas: Reservar com antecedência a capacidade de 800 conjuntos externos para Fevereiro e Março;", space_after=1.5)
    add_body("4. Formação do Estoque Regulador Suave de Meio de Ano: Produzir em ritmo constante de 3.800 un/mês entre Abril e Outubro, assegurando o estoque de 2.200 conjuntos para abastecer o pico de Novembro sem custos extras.", space_after=3)

    doc.add_page_break()

    # =========================================================
    # PÁGINA 10: 6. REFERÊNCIAS BIBLIOGRÁFICAS
    # =========================================================
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
        p_ref.paragraph_format.line_spacing = 1.15
        p_ref.paragraph_format.space_after = Pt(5)
        p_ref.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run_ref = p_ref.add_run(ref)
        run_ref.font.name = 'Arial'
        run_ref.font.size = Pt(9)

    doc.save("Relatorio_Tecnico_1_PAP_Manufatura.docx")
    
    cmd = "/opt/homebrew/bin/soffice --headless --convert-to pdf Relatorio_Tecnico_1_PAP_Manufatura.docx"
    subprocess.run(cmd, shell=True, check=True)
    
    doc_pdf = fitz.open("Relatorio_Tecnico_1_PAP_Manufatura.pdf")
    print(f"Total de páginas no PDF compilado: {len(doc_pdf)}")
    for i, page in enumerate(doc_pdf):
        txt = page.get_text().strip().replace('\n', ' ')
        print(f"Página {i+1} ({len(txt)} chars): {txt[:80]}")

build()
