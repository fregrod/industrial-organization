import sys
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# ==============================================================================
# PALETA DE CORES MODERNA & MINIMALISTA (EXECUTIVE DESIGN SYSTEM)
# ==============================================================================
C_NAVY_DARK = RGBColor(15, 23, 42)     # #0F172A - Slate 900
C_NAVY_HEADER = RGBColor(30, 41, 59)   # #1E293B - Slate 800
C_PRIMARY = RGBColor(37, 99, 235)      # #2563EB - Royal Blue
C_PRIMARY_LIGHT = RGBColor(239, 246, 255) # #EFF6FF - Soft Blue Pill
C_SUCCESS = RGBColor(16, 185, 129)     # #10B981 - Emerald 500
C_SUCCESS_LIGHT = RGBColor(236, 253, 245) # #ECFDF5 - Soft Green Pill
C_WARNING = RGBColor(245, 158, 11)     # #F59E0B - Amber 500
C_DANGER = RGBColor(239, 68, 68)       # #EF4444 - Rose Red
C_DANGER_LIGHT = RGBColor(254, 242, 242) # #FEF2F2 - Soft Red Pill
C_PURPLE = RGBColor(139, 92, 246)      # #8B5CF6 - Violet
C_BG_SLATE = RGBColor(248, 250, 252)   # #F8FAFC - Light Slate Background
C_CARD_BG = RGBColor(255, 255, 255)    # #FFFFFF - Crisp Card
C_BORDER = RGBColor(226, 232, 240)     # #E2E8F0 - Subtle Slate Border
C_TEXT_DARK = RGBColor(15, 23, 42)     # #0F172A - Titulos / Numeros
C_TEXT_MUTED = RGBColor(100, 116, 139) # #64748B - Subtitulos
C_TEXT_BODY = RGBColor(51, 65, 85)     # #334155 - Textos explicativos
C_WHITE = RGBColor(255, 255, 255)

blank_layout = prs.slide_layouts[6]

# ==============================================================================
# COMPONENTES MODERNOS DE DESIGN
# ==============================================================================

def create_modern_slide(tag_text, title_text, subtitle_text, slide_num):
    slide = prs.slides.add_slide(blank_layout)
    
    # Fundo Slate Suave
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_BG_SLATE
    bg.line.fill.background()
    
    # Header minimalista (Container invisível com tipografia moderna)
    # Tag Pill / Badge de categoria
    tag_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(8.0), Inches(0.35))
    tf_tag = tag_box.text_frame
    tf_tag.word_wrap = True
    p_tag = tf_tag.paragraphs[0]
    p_tag.text = tag_text.upper()
    p_tag.font.name = 'Arial'
    p_tag.font.size = Pt(9.5)
    p_tag.font.bold = True
    p_tag.font.color.rgb = C_PRIMARY
    
    # Título do Slide
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.68), Inches(10.5), Inches(0.5))
    tf_title = title_box.text_frame
    tf_title.word_wrap = True
    p_t = tf_title.paragraphs[0]
    p_t.text = title_text
    p_t.font.name = 'Arial'
    p_t.font.size = Pt(20)
    p_t.font.bold = True
    p_t.font.color.rgb = C_TEXT_DARK
    
    # Subtítulo explicativo
    if subtitle_text:
        sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.15), Inches(10.5), Inches(0.35))
        tf_sub = sub_box.text_frame
        tf_sub.word_wrap = True
        p_s = tf_sub.paragraphs[0]
        p_s.text = subtitle_text
        p_s.font.name = 'Arial'
        p_s.font.size = Pt(11.5)
        p_s.font.color.rgb = C_TEXT_MUTED
        
    # Linha sutil separadora
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.52), Inches(11.733), Pt(1))
    line.fill.solid()
    line.fill.fore_color.rgb = C_BORDER
    line.line.fill.background()
    
    # Rodapé Institucional
    foot_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.08), Inches(10.0), Inches(0.3))
    tff = foot_box.text_frame
    pf = tff.paragraphs[0]
    pf.text = "UnB / FT / EPR – Organização Industrial (2026.2) | Relatório Técnico 1: PAP em Manufatura"
    pf.font.name = 'Arial'
    pf.font.size = Pt(9)
    pf.font.color.rgb = C_TEXT_MUTED
    
    # Número do slide minimalista
    num_box = slide.shapes.add_textbox(Inches(11.0), Inches(7.08), Inches(1.533), Inches(0.3))
    tfn = num_box.text_frame
    pn = tfn.paragraphs[0]
    pn.alignment = PP_ALIGN.RIGHT
    pn.text = f"{slide_num:02d} / 15"
    pn.font.name = 'Arial'
    pn.font.size = Pt(9.5)
    pn.font.bold = True
    pn.font.color.rgb = C_PRIMARY
    
    return slide

def add_card(slide, left, top, width, height, bg_color=C_CARD_BG, border_color=C_BORDER):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = border_color
    card.line.width = Pt(1.0)
    return card

def add_kpi_card(slide, left, top, width, height, number_str, label_str, subtext_str=None, num_color=C_PRIMARY, bg_color=C_CARD_BG, border_color=C_BORDER):
    card = add_card(slide, left, top, width, height, bg_color, border_color)
    tb = slide.shapes.add_textbox(left + Inches(0.18), top + Inches(0.15), width - Inches(0.36), height - Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = True
    
    # Número de destaque
    p1 = tf.paragraphs[0]
    p1.text = number_str
    p1.font.name = 'Arial'
    p1.font.size = Pt(28)
    p1.font.bold = True
    p1.font.color.rgb = num_color
    p1.space_after = Pt(2)
    
    # Rótulo em maiúsculas
    p2 = tf.add_paragraph()
    p2.text = label_str.upper()
    p2.font.name = 'Arial'
    p2.font.size = Pt(9.5)
    p2.font.bold = True
    p2.font.color.rgb = C_TEXT_MUTED
    p2.space_after = Pt(4)
    
    # Subtexto / Insight
    if subtext_str:
        p3 = tf.add_paragraph()
        p3.text = subtext_str
        p3.font.name = 'Arial'
        p3.font.size = Pt(11)
        p3.font.color.rgb = C_TEXT_BODY
        p3.line_spacing = 1.15
        
    return card

def add_bullet_list(text_frame, items, font_size=11.5, space_after=6):
    for i, item in enumerate(items):
        p = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
        p.space_after = Pt(space_after)
        p.line_spacing = 1.2
        if isinstance(item, tuple):
            title, desc = item
            r1 = p.add_run()
            r1.text = title + (" " if desc else "")
            r1.font.name = 'Arial'
            r1.font.bold = True
            r1.font.size = Pt(font_size)
            r1.font.color.rgb = C_TEXT_DARK
            if desc:
                r2 = p.add_run()
                r2.text = desc
                r2.font.name = 'Arial'
                r2.font.size = Pt(font_size)
                r2.font.color.rgb = C_TEXT_BODY
        else:
            r = p.add_run()
            r.text = item
            r.font.name = 'Arial'
            r.font.size = Pt(font_size)
            r.font.color.rgb = C_TEXT_BODY

# ==============================================================================
# SLIDE 1: CAPA EXECUTIVA MINIMALISTA (DARK NAVY SLEEK)
# ==============================================================================
slide1 = prs.slides.add_slide(blank_layout)
bg1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
bg1.fill.solid()
bg1.fill.fore_color.rgb = C_NAVY_DARK
bg1.line.fill.background()

# Pill Header Institucional
pill1 = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.0), Inches(4.8), Inches(0.42))
pill1.fill.solid()
pill1.fill.fore_color.rgb = RGBColor(30, 41, 59)
pill1.line.color.rgb = RGBColor(51, 65, 85)
tb_p1 = slide1.shapes.add_textbox(Inches(1.1), Inches(1.03), Inches(4.6), Inches(0.35))
tf_p1 = tb_p1.text_frame
p_pill = tf_p1.paragraphs[0]
p_pill.text = "UNIVERSIDADE DE BRASÍLIA  •  FACULDADE DE TECNOLOGIA"
p_pill.font.name = 'Arial'
p_pill.font.size = Pt(9.5)
p_pill.font.bold = True
p_pill.font.color.rgb = RGBColor(147, 197, 253)

# Título Principal
tb1_t = slide1.shapes.add_textbox(Inches(1.0), Inches(1.65), Inches(11.333), Inches(1.8))
tf1_t = tb1_t.text_frame
tf1_t.word_wrap = True
p1_1 = tf1_t.paragraphs[0]
p1_1.text = "Planejamento Agregado da Produção"
p1_1.font.name = 'Arial'
p1_1.font.size = Pt(36)
p1_1.font.bold = True
p1_1.font.color.rgb = C_WHITE

p1_2 = tf1_t.add_paragraph()
p1_2.text = "Modelagem Matemática de 12 Meses Aplicada à Indústria de Manufatura: Vértice Confecções Ltda."
p1_2.font.name = 'Arial'
p1_2.font.size = Pt(16)
p1_2.font.color.rgb = RGBColor(148, 163, 184)
p1_2.space_before = Pt(8)

# 3 Cards Executivos Inferiores
card1_a = add_card(slide1, Inches(1.0), Inches(3.9), Inches(3.6), Inches(2.2), RGBColor(30, 41, 59), RGBColor(51, 65, 85))
tb1_a = slide1.shapes.add_textbox(Inches(1.15), Inches(4.05), Inches(3.3), Inches(1.9))
tf1_a = tb1_a.text_frame
tf1_a.word_wrap = True
add_bullet_list(tf1_a, [
    ("DISCIPLINA", ""),
    ("Organização Industrial", ""),
    ("Semestre:", "2026.2"),
    ("Atividade:", "Relatório Técnico 1 & Seminário")
], font_size=11, space_after=4)
for p in tf1_a.paragraphs:
    for r in p.runs:
        r.font.color.rgb = RGBColor(226, 232, 240) if r.font.bold else RGBColor(148, 163, 184)

card1_b = add_card(slide1, Inches(4.85), Inches(3.9), Inches(3.6), Inches(2.2), RGBColor(30, 41, 59), RGBColor(51, 65, 85))
tb1_b = slide1.shapes.add_textbox(Inches(5.0), Inches(4.05), Inches(3.3), Inches(1.9))
tf1_b = tb1_b.text_frame
tf1_b.word_wrap = True
add_bullet_list(tf1_b, [
    ("CORPO DOCENTE", ""),
    ("Prof. Dr. Marcelo Carneiro Gonçalves", ""),
    ("Profa. Dra. Camilla Pimenta Barbosa", ""),
    ("Prof. Dr. João Victor Queiroz", "")
], font_size=11, space_after=4)
for p in tf1_b.paragraphs:
    for r in p.runs:
        r.font.color.rgb = RGBColor(226, 232, 240) if r.font.bold else RGBColor(148, 163, 184)

card1_c = add_card(slide1, Inches(8.7), Inches(3.9), Inches(3.6), Inches(2.2), RGBColor(30, 41, 59), RGBColor(51, 65, 85))
tb1_c = slide1.shapes.add_textbox(Inches(8.85), Inches(4.05), Inches(3.3), Inches(1.9))
tf1_c = tb1_c.text_frame
tf1_c.word_wrap = True
add_bullet_list(tf1_c, [
    ("EQUIPE DE ENGENHARIA", ""),
    ("Grupo de Alunos de Graduação", ""),
    ("Curso:", "Engenharia de Produção (FT/UnB)"),
    ("Data:", "01 de Outubro de 2026")
], font_size=11, space_after=4)
for p in tf1_c.paragraphs:
    for r in p.runs:
        r.font.color.rgb = RGBColor(226, 232, 240) if r.font.bold else RGBColor(148, 163, 184)

slide1.notes_slide.notes_text_frame.text = (
    "[TEMPO: 00:00 - 01:00 | Integrante 1]\n"
    "Bom dia professores Marcelo Gonçalves, Camilla Pimenta e João Victor Queiroz, e colegas de turma. "
    "Apresentamos o Relatório Técnico 1 e Seminário de Organização Industrial com a modelagem matemática completa de PAP "
    "para a empresa Vértice Confecções Ltda., localizada no Distrito Federal. "
    "Nosso objetivo é provar como o planejamento tático de operações equilibra custos, nível de serviço e legislação trabalhista."
)

# ==============================================================================
# SLIDE 2: CONTEXTO DA MPME E O DILEMA OPERACIONAL
# ==============================================================================
slide2 = create_modern_slide("Diagnóstico Operacional", "A Empresa e o Dilema da Sazonalidade", "Caracterização da fábrica de pequeno porte e o desafio tático de manufatura", 2)

# 3 Stat Cards no topo
add_kpi_card(slide2, Inches(0.8), Inches(1.7), Inches(3.7), Inches(1.6), "48.000 un", "Demanda Anual", "Média de 4.000 peças/mês com 2 grandes picos sazonais no ano.", C_PRIMARY)
add_kpi_card(slide2, Inches(4.8), Inches(1.7), Inches(3.7), Inches(1.6), "20 Costureiras", "Quadro Efetivo", "Equipe fixa instalada em Taguatinga-DF (jornada de 160h/mês).", C_NAVY_DARK)
add_kpi_card(slide2, Inches(8.8), Inches(1.7), Inches(3.7), Inches(1.6), "2 Picos Críticos", "Sazonalidade", "Jan-Mar (renovação escolar/frotas) e Nov (aquecimento industrial).", C_WARNING)

# 2 Cards Comparativos embaixo: Dilema vs Solução
card2_left = add_card(slide2, Inches(0.8), Inches(3.55), Inches(5.7), Inches(3.25), C_WHITE, C_BORDER)
tb2_l = slide2.shapes.add_textbox(Inches(1.0), Inches(3.7), Inches(5.3), Inches(2.95))
tf2_l = tb2_l.text_frame
tf2_l.word_wrap = True
add_bullet_list(tf2_l, [
    ("O Dilema Típico sem PAP Formal:", ""),
    ("• Resposta Reativa:", "Fábrica alternava entre atrasos e contratações desesperadas."),
    ("• Colapso de Entregas:", "No 1º trimestre, a demanda estoura em até 30% da capacidade."),
    ("• Multas Contratuais:", "Clientes corporativos cancelam pedidos e exigem ressarcimento."),
    ("• Instabilidade Fabril:", "Contratações sem treino aumentavam refugos e acidentes.")
], font_size=12, space_after=6)

card2_right = add_card(slide2, Inches(6.8), Inches(3.55), Inches(5.7), Inches(3.25), C_PRIMARY_LIGHT, C_PRIMARY)
tb2_r = slide2.shapes.add_textbox(Inches(7.0), Inches(3.7), Inches(5.3), Inches(2.95))
tf2_r = tb2_r.text_frame
tf2_r.word_wrap = True
add_bullet_list(tf2_r, [
    ("A Solução com o PAP Tático de Manufatura:", ""),
    ("• Visão de 12 Meses:", "Antecipar oscilações e estabilizar a mão de obra interna."),
    ("• Decisões Integradas:", "Balancear estoques, horas extras e terceirização (facção)."),
    ("• Nível de Serviço 100%:", "Atender rigorosamente os prazos contratuais firmados."),
    ("• Menor Custo Total:", "Minimizar o dispêndio anual respeitando a CLT.")
], font_size=12, space_after=6)

slide2.notes_slide.notes_text_frame.text = (
    "[TEMPO: 01:00 - 02:00 | Integrante 1]\n"
    "A Vértice Confecções opera com 20 costureiras e enfrenta forte sazonalidade: 48 mil conjuntos ao ano, "
    "com picos intensos no 1º trimestre e em novembro. Sem o PAP formal, a gestão oscilava de forma reativa: "
    "atrasava entregas no início do ano e demitia no meio do ano. "
    "A modelagem matemática do PAP estabelece previsibilidade para atender 100% dos pedidos ao menor custo."
)

# ==============================================================================
# SLIDE 3: CRONOANÁLISE E MAPEAMENTO DO GARGALO
# ==============================================================================
slide3 = create_modern_slide("Capacidade & Cronoanálise", "Mapeamento de Processo e Identificação do Gargalo", "Determinação do Tempo Padrão (TP) e capacidade nominal base da fábrica", 3)

# Imagem do Fluxograma (figura1_fluxo_processo.png)
card3_img = add_card(slide3, Inches(0.8), Inches(1.7), Inches(11.733), Inches(2.35), C_WHITE, C_BORDER)
if os.path.exists("figuras/figura1_fluxo_processo.png"):
    slide3.shapes.add_picture("figuras/figura1_fluxo_processo.png", Inches(0.95), Inches(1.8), Inches(11.433), Inches(2.15))

# 3 Stat Cards embaixo com os números da cronoanálise
add_kpi_card(slide3, Inches(0.8), Inches(4.3), Inches(3.7), Inches(2.5), "48 min", "Tempo Padrão no Gargalo", "TP = 0,80h apurado na Costura/Montagem, incluindo 15% de tolerâncias para fadiga e pausas da NR-17.", C_PRIMARY)
add_kpi_card(slide3, Inches(4.8), Inches(4.3), Inches(3.7), Inches(2.5), "200 un/mês", "Produtividade / Costureira", "Jornada mensal líquida de 160h ÷ 0,80h = 200 conjuntos por operadora em turno normal.", C_SUCCESS)
add_kpi_card(slide3, Inches(8.8), Inches(4.3), Inches(3.7), Inches(2.5), "4.000 un/mês", "Capacidade Nominal Base", "20 costureiras × 200 un/mês = 4.000 conjuntos/mês. Referência para o dimensionamento do PAP.", C_NAVY_DARK)

slide3.notes_slide.notes_text_frame.text = (
    "[TEMPO: 02:00 - 03:00 | Integrante 1]\n"
    "Conforme ensinado pelo Prof. Marcelo Carneiro na Aula 01, a capacidade da fábrica é ditada pelo gargalo. "
    "A cronoanálise identificou o gargalo na célula de costura e montagem. "
    "O Tempo Padrão foi mensurado em 0,80 horas (48 minutos), incluindo tolerâncias ergonômicas da NR-17. "
    "Isso resulta numa produtividade de 200 conjuntos por operadora e capacidade base de 4.000 peças/mês."
)

# ==============================================================================
# SLIDE 4: CURVA DE DEMANDA E HORIZONTE DE 12 MESES
# ==============================================================================
slide4 = create_modern_slide("Previsão de Demanda", "Perfil Sazonal da Demanda em 12 Meses", "Oscilação mensal de mercado e confronto direto com a capacidade instalada", 4)

# 3 Stat Cards no topo
add_kpi_card(slide4, Inches(0.8), Inches(1.7), Inches(3.7), Inches(1.6), "5.200 un", "Pico Máximo (Fevereiro)", "+30% acima da capacidade fabril instalada (+1.200 unidades).", C_DANGER)
add_kpi_card(slide4, Inches(4.8), Inches(1.7), Inches(3.7), Inches(1.6), "3.000 un", "Vale Mínimo (Junho)", "-25% abaixo da capacidade fabril instalada (-1.000 unidades).", C_SUCCESS)
add_kpi_card(slide4, Inches(8.8), Inches(1.7), Inches(3.7), Inches(1.6), "400 un", "Estoque de Segurança", "Estoque inicial de segurança para suportar o início do plano.", C_PRIMARY)

# Card visual com mini-tabela/distribuição dos 12 meses
card4_table = add_card(slide4, Inches(0.8), Inches(3.55), Inches(11.733), Inches(3.25), C_WHITE, C_BORDER)
tb4_t = slide4.shapes.add_textbox(Inches(1.0), Inches(3.7), Inches(11.333), Inches(2.95))
tf4_t = tb4_t.text_frame
tf4_t.word_wrap = True

add_bullet_list(tf4_t, [
    ("Distribuição Mensal da Demanda Consolidada (48.000 conjuntos):", ""),
    ("• 1º Trimestre (Pico Crítico):", "Jan: 4.600 un  |  Fev: 5.200 un  |  Mar: 4.800 un  (Demanda acumulada: 14.600 un vs. Capacidade: 12.000 un)."),
    ("• 2º Trimestre (Desaceleração):", "Abr: 3.600 un  |  Mai: 3.200 un  |  Jun: 3.000 un  (Vale sazonal; folga produtiva na fábrica)."),
    ("• 3º Trimestre (Recuperação):", "Jul: 3.800 un  |  Ago: 3.600 un  |  Set: 3.400 un  (Volume equilibrado; oportunidade de formação de estoque)."),
    ("• 4º Trimestre (2º Pico do Ano):", "Out: 3.800 un  |  Nov: 5.000 un  |  Dez: 4.000 un  (Aquecimento para encomendas industriais e de fim de ano).")
], font_size=12, space_after=6)

slide4.notes_slide.notes_text_frame.text = (
    "[TEMPO: 03:00 - 04:00 | Integrante 2]\n"
    "Analisando o horizonte de 12 meses estabelecido pelo Prof. João Victor na Aula 03, vemos claramente a montanha-russa de demanda. "
    "O primeiro trimestre é o período mais crítico: a demanda supera a capacidade em 2.600 peças. "
    "Depois temos um vale em junho com 3.000 peças e novo pico em novembro com 5.000 peças. "
    "O desafio do PAP é balancear essa sazonalidade sem gerar custos desnecessários."
)

# ==============================================================================
# SLIDE 5: AS 3 ESTRATÉGIAS CLÁSSICAS DE PAP
# ==============================================================================
slide5 = create_modern_slide("Fundamentação Teórica", "Os Modelos Clássicos de Decisão do PAP", "Conceituação teórica segundo Nigel Slack e Corrêa & Corrêa", 5)

card5_a = add_card(slide5, Inches(0.8), Inches(1.7), Inches(3.7), Inches(5.1), C_WHITE, C_BORDER)
tb5_a = slide5.shapes.add_textbox(Inches(1.0), Inches(1.9), Inches(3.3), Inches(4.7))
tf5_a = tb5_a.text_frame
tf5_a.word_wrap = True
add_bullet_list(tf5_a, [
    ("Estratégia 1: Nivelamento", ""),
    ("• Taxa Constante:", "Produção perfeitamente fixa em 4.000 un/mês com 20 costureiras."),
    ("• Mecanismo Amortecedor:", "Uso exclusivo de estoque pulmão e pedidos em atraso (backlog)."),
    ("• Foco Estrutural:", "Estabilidade fabril e previsibilidade operacional interna."),
    ("• Principal Risco:", "Rupturas graves em picos sazonais no início do ano.")
], font_size=12, space_after=8)

card5_b = add_card(slide5, Inches(4.8), Inches(1.7), Inches(3.7), Inches(5.1), C_WHITE, C_BORDER)
tb5_b = slide5.shapes.add_textbox(Inches(5.0), Inches(1.9), Inches(3.3), Inches(4.7))
tf5_b = tb5_b.text_frame
tf5_b.word_wrap = True
add_bullet_list(tf5_b, [
    ("Estratégia 2: Acompanhamento", ""),
    ("• Chase Strategy:", "Capacidade persegue rigorosamente a curva da demanda mensal (P = D)."),
    ("• Mecanismo Amortecedor:", "Contratações e demissões sistemáticas de costureiras."),
    ("• Foco Estrutural:", "Estoque final zero em todos os meses do ano."),
    ("• Principal Risco:", "Caos trabalhista, custo rescisório e queda drástica de qualidade.")
], font_size=12, space_after=8)

card5_c = add_card(slide5, Inches(8.8), Inches(1.7), Inches(3.7), Inches(5.1), C_PRIMARY_LIGHT, C_PRIMARY)
tb5_c = slide5.shapes.add_textbox(Inches(9.0), Inches(1.9), Inches(3.3), Inches(4.7))
tf5_c = tb5_c.text_frame
tf5_c.word_wrap = True
add_bullet_list(tf5_c, [
    ("Estratégia 3: Híbrida / Patamares", ""),
    ("• Solução Flexível:", "Base estável de 19 operadoras complementada por recursos elásticos."),
    ("• Mecanismo Amortecedor:", "Horas extras reguladas (CLT) + subcontratação de facções."),
    ("• Foco Estrutural:", "Menor custo global com garantia de 100% de pontualidade."),
    ("• Principal Vantagem:", "Equilíbrio sustentável entre fábrica, mercado e legislação.")
], font_size=12, space_after=8)

slide5.notes_slide.notes_text_frame.text = (
    "[TEMPO: 04:00 - 05:00 | Integrante 2]\n"
    "A teoria de Nigel Slack e Corrêa & Corrêa define três caminhos fundamentais: "
    "O Nivelamento mantém a fábrica fixa e deixa o estoque oscilar; "
    "O Acompanhamento modula os funcionários para zerar estoques; "
    "E a Estratégia Híbrida combina a estabilidade interna com recursos elásticos nos picos de demanda. "
    "Modelamos as três matematicamente para o caso da Vértice."
)

# ==============================================================================
# SLIDE 6: MATRIZ DE CUSTOS E LIMITES CLT
# ==============================================================================
slide6 = create_modern_slide("Parâmetros Financeiros", "Matriz de Custos Unitários e Limitações da CLT", "Dados coletados no polo têxtil do DF e exigências legais trabalhistas", 6)

# 4 Stat Cards de Custos
add_kpi_card(slide6, Inches(0.8), Inches(1.7), Inches(2.7), Inches(2.3), "R$ 38,00", "Produção Normal", "Custo unitário absorvido por conjunto em turno regular de trabalho.", C_PRIMARY)
add_kpi_card(slide6, Inches(3.8), Inches(1.7), Inches(2.7), Inches(2.3), "R$ 46,00", "Hora Extra (HE)", "Adicional constitucional de 50% sobre a mão de obra direta.", C_WARNING)
add_kpi_card(slide6, Inches(6.8), Inches(1.7), Inches(2.7), Inches(2.3), "R$ 52,00", "Facção Externa", "Custo do serviço terceirizado em oficinas parceiras homologadas.", C_PURPLE)
add_kpi_card(slide6, Inches(9.8), Inches(1.7), Inches(2.7), Inches(2.3), "R$ 3,00", "Posse de Estoque", "Custo mensal por unidade armazenada sobre o estoque médio.", C_NAVY_DARK)

# 2 Cards Inferiores: Custos de Pessoal e Regra CLT
card6_l = add_card(slide6, Inches(0.8), Inches(4.25), Inches(5.7), Inches(2.55), C_WHITE, C_BORDER)
tb6_l = slide6.shapes.add_textbox(Inches(1.0), Inches(4.4), Inches(5.3), Inches(2.25))
tf6_l = tb6_l.text_frame
tf6_l.word_wrap = True
add_bullet_list(tf6_l, [
    ("Custos de Movimentação de Pessoal e Penalidades:", ""),
    ("• Custo de Admissão / Treinamento:", "R$ 1.600,00 por costureira contratada."),
    ("• Custo de Demissão / Rescisão:", "R$ 3.000,00 (aviso prévio indenizado + FGTS 40%)."),
    ("• Penalidade por Atraso (Backlog):", "R$ 20,00 por conjunto/mês em atraso.")
], font_size=11.5, space_after=4)

card6_r = add_card(slide6, Inches(6.8), Inches(4.25), Inches(5.7), Inches(2.55), C_DANGER_LIGHT, C_DANGER)
tb6_r = slide6.shapes.add_textbox(Inches(7.0), Inches(4.4), Inches(5.3), Inches(2.25))
tf6_r = tb6_r.text_frame
tf6_r.word_wrap = True
add_bullet_list(tf6_r, [
    ("Restrição Jurídica Estrita – Artigo 59 da CLT:", ""),
    ("• Limite Máximo Legal:", "Até 2 horas extras por dia por operadora (máx. 25% da jornada)."),
    ("• Teto Operacional Fabril:", "Máximo de 800 conjuntos/mês produzidos via horas extras."),
    ("• Segurança Regulatória:", "Impossibilita absorver picos de 1.200 un apenas com hora extra.")
], font_size=11.5, space_after=4)

slide6.notes_slide.notes_text_frame.text = (
    "[TEMPO: 05:00 - 06:00 | Integrante 2]\n"
    "Os custos foram obtidos no polo têxtil do DF: R$ 38 no normal, R$ 46 na hora extra e R$ 52 na facção externa. "
    "A estocagem custa R$ 3 por peça ao mês, e atrasar pedidos acarreta multa contratual de R$ 20 por mês. "
    "Destacamos o Artigo 59 da CLT: o teto legal de 2h extras por dia limita a produção extra a 800 conjuntos por mês, "
    "tornando juridicamente impossível cobrir os picos de demanda apenas com horas extras."
)

# ==============================================================================
# SLIDE 7: ESTRATÉGIA 1: NIVELAMENTO (TAXA CONSTANTE)
# ==============================================================================
slide7 = create_modern_slide("Modelagem E1", "Estratégia 1: Produção Constante (Nivelamento)", "Produção fixa de 4.000 un/mês e o colapso no atendimento ao cliente", 7)

# 3 Red Alert Cards no Topo
add_kpi_card(slide7, Inches(0.8), Inches(1.7), Inches(3.7), Inches(1.7), "R$ 108.000", "Multas de Backlog", "Penalidades contratuais acumuladas ao longo de 5 meses de atrasos.", C_DANGER)
add_kpi_card(slide7, Inches(4.8), Inches(1.7), Inches(3.7), Inches(1.7), "2.200 un", "Pico de Atraso (Março)", "Pico máximo de pedidos acumulados sem entrega aos clientes corporativos.", C_DANGER)
add_kpi_card(slide7, Inches(8.8), Inches(1.7), Inches(3.7), Inches(1.7), "79,2%", "Nível de Entrega (OTD)", "On-Time Delivery inaceitável para contratos corporativos e licitações.", C_DANGER)

# Card Explicativo com diagnóstico
card7_diag = add_card(slide7, Inches(0.8), Inches(3.65), Inches(11.733), Inches(3.15), C_WHITE, C_BORDER)
tb7_d = slide7.shapes.add_textbox(Inches(1.0), Inches(3.8), Inches(11.333), Inches(2.85))
tf7_d = tb7_d.text_frame
tf7_d.word_wrap = True
add_bullet_list(tf7_d, [
    ("Diagnóstico Operacional da Estratégia de Nivelamento:", ""),
    ("• Esgotamento Imediato do Estoque:", "O estoque de segurança de 400 peças foi consumido integralmente já em Janeiro."),
    ("• Efeito Bola de Neve de Pedidos Reprimidos:", "Jan: -200 un  |  Fev: -1.400 un  |  Mar: -2.200 un  |  Abr: -1.800 un  |  Mai: -1.000 un."),
    ("• Regularização Tardia:", "A fábrica só conseguiu zerar os pedidos atrasados no final de Junho (5 meses de espera do cliente)."),
    ("• Custo Total Consolidado da Estratégia 1:", "R$ 1.968.600,00 (Produção Normal: R$ 1.824.000 + Posse: R$ 36.600 + Multas: R$ 108.000).")
], font_size=12, space_after=6)

slide7.notes_slide.notes_text_frame.text = (
    "[TEMPO: 06:00 - 07:00 | Integrante 3]\n"
    "Na Estratégia 1, fixamos 4.000 peças/mês com 20 costureiras. O resultado prático foi um desastre: "
    "como a demanda de Jan a Mar somou 14.600 peças e a capacidade foi de 12.000, o estoque de 400 peças acabou em janeiro. "
    "A empresa acumulou 2.200 peças em atraso em março e só normalizou as entregas em junho! "
    "Isso gerou R$ 108 mil em multas e derrubou o índice de pontualidade para 79%."
)

# ==============================================================================
# SLIDE 8: ESTRATÉGIA 2: ACOMPANHAMENTO (CHASE)
# ==============================================================================
slide8 = create_modern_slide("Modelagem E2", "Estratégia 2: Acompanhamento da Demanda (Chase)", "Variação mensal da força de trabalho e a armadilha do turnover excessivo", 8)

# 3 Stat Cards no topo
add_kpi_card(slide8, Inches(0.8), Inches(1.7), Inches(3.7), Inches(1.7), "36", "Movimentações de Pessoal", "18 contratações e 18 demissões realizadas ao longo de 12 meses.", C_WARNING)
add_kpi_card(slide8, Inches(4.8), Inches(1.7), Inches(3.7), Inches(1.7), "R$ 82.800", "Custos de Transição", "R$ 28.800 em admissões/treinamento + R$ 54.000 em rescisões contratuais.", C_WARNING)
add_kpi_card(slide8, Inches(8.8), Inches(1.7), Inches(3.7), Inches(1.7), "15 a 26", "Variação do Quadro", "Equipe oscila caoticamente: 26 costureiras em Fev e apenas 15 em Jun.", C_NAVY_DARK)

# Card de Impactos Ocultos
card8_diag = add_card(slide8, Inches(0.8), Inches(3.65), Inches(11.733), Inches(3.15), C_WHITE, C_BORDER)
tb8_d = slide8.shapes.add_textbox(Inches(1.0), Inches(3.8), Inches(11.333), Inches(2.85))
tf8_d = tb8_d.text_frame
tf8_d.word_wrap = True
add_bullet_list(tf8_d, [
    ("Impactos Ocultos e Perdas de Produtividade no Modelo Chase:", ""),
    ("• Destruição da Curva de Aprendizagem:", "Costureiras recém-admitidas têm rendimento 30% inferior no primeiro mês."),
    ("• Aumento Severo de Refugo e Retrabalho:", "Costuras tortas e retrabalho aumentam justamente nos meses de pico."),
    ("• Clima Organizacional Hostil:", "Insegurança constante no chão de fábrica e desmotivação geral das operadoras."),
    ("• Custo Total Consolidado da Estratégia 2:", "R$ 1.921.200,00 (Produção: R$ 1.824.000 + Posse: R$ 14.400 + Turnover: R$ 82.800).")
], font_size=12, space_after=6)

slide8.notes_slide.notes_text_frame.text = (
    "[TEMPO: 07:00 - 08:00 | Integrante 3]\n"
    "Na Estratégia 2, contratamos e demitimos para igualar a demanda mensal. "
    "Apesar do custo financeiro de R$ 1,92 milhão parecer menor que o Nivelamento, o custo social é inaceitável: "
    "foram 36 movimentações de pessoal (18 admissões e 18 demissões), custando R$ 82.800 em encargos. "
    "Além disso, a perda de costureiras experientes faz o índice de retrabalho disparar nos meses mais críticos."
)

# ==============================================================================
# SLIDE 9: SÍNTESE CRÍTICA: POR QUE OS EXTREMOS FALHAM?
# ==============================================================================
slide9 = create_modern_slide("Diagnóstico Crítico", "Por que as Estratégias Extremas Falham na Prática?", "Confronto entre as fragilidades do Nivelamento puro e do Acompanhamento puro", 9)

card9_left = add_card(slide9, Inches(0.8), Inches(1.7), Inches(5.7), Inches(5.1), C_DANGER_LIGHT, C_DANGER)
tb9_l = slide9.shapes.add_textbox(Inches(1.0), Inches(1.9), Inches(5.3), Inches(4.7))
tf9_l = tb9_l.text_frame
tf9_l.word_wrap = True
add_bullet_list(tf9_l, [
    ("FRAGILIDADES DO NIVELAMENTO PURO (E1)", ""),
    ("❌ Inflexibilidade Estrutural:", "A produção constante é incapaz de responder a picos iniciais de demanda."),
    ("❌ Esgotamento Precoce de Estoque:", "O estoque de segurança de 400 peças foi consumido nos primeiros 30 dias."),
    ("❌ Drenagem Financeira por Backlog:", "R$ 108.000,00 gastos exclusivamente em multas e indenizações."),
    ("❌ Perda de Reputação e Clientes:", "Tempo de espera de até 5 meses destrói a confiança comercial da fábrica.")
], font_size=12, space_after=8)

card9_right = add_card(slide9, Inches(6.8), Inches(1.7), Inches(5.7), Inches(5.1), C_WARNING, C_WARNING)
card9_right.fill.solid()
card9_right.fill.fore_color.rgb = RGBColor(254, 243, 199) # soft amber
tb9_r = slide9.shapes.add_textbox(Inches(7.0), Inches(1.9), Inches(5.3), Inches(4.7))
tf9_r = tb9_r.text_frame
tf9_r.word_wrap = True
add_bullet_list(tf9_r, [
    ("FRAGILIDADES DO ACOMPANHAMENTO PURO (E2)", ""),
    ("❌ Mão de Obra Tratada como Insumo:", "Despreza a especialização artesanal exigida na costura industrial."),
    ("❌ Custo Rescisório Desperdiçado:", "R$ 82.800,00 gastos em burocracia trabalhista e rescisões."),
    ("❌ Degradação da Qualidade:", "Equipes recém-contratadas aumentam o índice de retrabalho em peças pesadas."),
    ("❌ Risco Trabalhista e Passivos:", "Aumento de ações judiciais e despesas com homologações no sindicato.")
], font_size=12, space_after=8)

slide9.notes_slide.notes_text_frame.text = (
    "[TEMPO: 08:00 - 09:00 | Integrante 3]\n"
    "Este slide resume a conclusão fundamental: os extremos teóricos falham na vida real. "
    "O Nivelamento destrói a pontualidade e o relacionamento com os clientes por falta de capacidade flexível. "
    "O Acompanhamento destrói a qualidade fabril e gasta rios de dinheiro em rescisões rescisórias. "
    "Precisamos de uma terceira via: a Solução Híbrida."
)

# ==============================================================================
# SLIDE 10: ESTRATÉGIA 3: A SOLUÇÃO HÍBRIDA RECOMENDADA
# ==============================================================================
slide10 = create_modern_slide("Solução Recomendada", "Estratégia 3: Híbrida em Patamares (Ótima)", "Base própria estável combinada com horas extras reguladas e facção externa", 10)

# 4 Hero KPI Cards
add_kpi_card(slide10, Inches(0.8), Inches(1.7), Inches(2.7), Inches(2.2), "19", "Costureiras Fixas", "Quadro permanente com capacidade interna estável de 3.800 un/mês.", C_SUCCESS)
add_kpi_card(slide10, Inches(3.8), Inches(1.7), Inches(2.7), Inches(2.2), "1.800 un", "Horas Extras (1º Tri)", "Jan: 400 | Fev: 800 | Mar: 600 (dentro do teto do Art. 59 da CLT).", C_PRIMARY)
add_kpi_card(slide10, Inches(6.8), Inches(1.7), Inches(2.7), Inches(2.2), "1.000 un", "Facção Parceira", "Fev: 600 un e Mar: 400 un subcontratadas para absorver o pico.", C_PURPLE)
add_kpi_card(slide10, Inches(9.8), Inches(1.7), Inches(2.7), Inches(2.2), "Zero", "Demissões no Ano", "Estabilidade total da equipe e retenção integral de talentos.", C_SUCCESS)

# Card de Lógica de Engenharia
card10_bot = add_card(slide10, Inches(0.8), Inches(4.15), Inches(11.733), Inches(2.65), C_SUCCESS_LIGHT, C_SUCCESS)
tb10_b = slide10.shapes.add_textbox(Inches(1.0), Inches(4.25), Inches(11.333), Inches(2.45))
tf10_b = tb10_b.text_frame
tf10_b.word_wrap = True
add_bullet_list(tf10_b, [
    ("Lógica de Engenharia da Estratégia Híbrida:", ""),
    ("• 1º Trimestre (Pico Crítico):", "Combina capacidade própria (3.800 un) + HE no teto legal (até 800 un) + Facção externa (até 600 un)."),
    ("• 2º Trimestre (Vale):", "Demanda cai, mas a fábrica mantém 3.800 un/mês, absorvendo a mão de obra sem demissões."),
    ("• 3º Trimestre (Formação de Pulmão):", "Acumula estoque estrategicamente até 1.400 unidades em Outubro com custo de posse irrisório."),
    ("• 4º Trimestre (2º Pico):", "O pico de Novembro (5.000 un) é atendido com o estoque pulmão, sem gastar nem R$ 1,00 em hora extra!")
], font_size=11, space_after=4)

slide10.notes_slide.notes_text_frame.text = (
    "[TEMPO: 09:00 - 10:30 | Integrante 4]\n"
    "Desenvolvemos a Estratégia 3 com uma lógica refinada de engenharia: "
    "Fixamos 19 costureiras permanentes (3.800 peças/mês). "
    "Nos picos de Jan a Mar, acionamos horas extras dentro do teto legal da CLT e terceirizamos 1.000 peças com facções parceiras. "
    "No meio do ano, não demitimos: produzimos estoque pulmão que chega a 1.400 peças em outubro. "
    "Assim, atendemos o pico de novembro com estoque, com custo de horas extras ZERO no segundo semestre!"
)

# ==============================================================================
# SLIDE 11: DINÂMICA DE ESTOQUES E NÍVEL DE SERVIÇO
# ==============================================================================
slide11 = create_modern_slide("Balanço de Materiais", "Dinâmica de Estoques e Nível de Serviço", "Confronto da evolução dos estoques e eliminação total de atrasos (E1 vs. E3)", 11)

# Imagem dos Estoques (figura3_estoques_e_backlog.png)
card11_img = add_card(slide11, Inches(0.8), Inches(1.7), Inches(11.733), Inches(2.4), C_WHITE, C_BORDER)
if os.path.exists("figuras/figura3_estoques_e_backlog.png"):
    slide11.shapes.add_picture("figuras/figura3_estoques_e_backlog.png", Inches(0.95), Inches(1.8), Inches(11.433), Inches(2.2))

# 3 Stat Cards embaixo
add_kpi_card(slide11, Inches(0.8), Inches(4.35), Inches(3.7), Inches(2.45), "100%", "Nível de Entrega (OTD)", "Zero atrasos ao longo dos 12 meses. Atendimento pontual em 100% dos pedidos.", C_SUCCESS)
add_kpi_card(slide11, Inches(4.8), Inches(4.35), Inches(3.7), Inches(2.45), "1.400 un", "Estoque Pulmão Máximo", "Pico controlado em Outubro para bancar o choque de Novembro sem custos extras.", C_PRIMARY)
add_kpi_card(slide11, Inches(8.8), Inches(4.35), Inches(3.7), Inches(2.45), "800 un", "Estoque Final Seguro", "Atende com folga o estoque de segurança para a virada do próximo ano fiscal.", C_NAVY_DARK)

slide11.notes_slide.notes_text_frame.text = (
    "[TEMPO: 10:30 - 11:30 | Integrante 4]\n"
    "O gráfico demonstra o contraste gritante: enquanto o Nivelamento afunda no backlog em vermelho durante cinco meses, "
    "a Estratégia Híbrida mantém a barra verde de estoque sempre positiva. "
    "Atingimos 100% de nível de serviço, zero peças em atraso, e ainda entregamos 800 peças em estoque no final de dezembro, "
    "garantindo a segurança operacional para o ano seguinte."
)

# ==============================================================================
# SLIDE 12: QUADRO COMPARATIVO CONSOLIDADO DE CUSTOS
# ==============================================================================
slide12 = create_modern_slide("Resultado Econômico", "Quadro Comparativo Consolidado de Custos e Decisão", "Demonstração analítica da superioridade econômica da Estratégia Híbrida", 12)

# Imagem do Gráfico de Custos à esquerda (figura4_comparativo_custos.png)
card12_img = add_card(slide12, Inches(0.8), Inches(1.7), Inches(5.7), Inches(5.1), C_WHITE, C_BORDER)
if os.path.exists("figuras/figura4_comparativo_custos.png"):
    slide12.shapes.add_picture("figuras/figura4_comparativo_custos.png", Inches(0.95), Inches(1.85), Inches(5.4), Inches(4.8))

# 3 Scorecards de Decisão à direita
card12_e1 = add_card(slide12, Inches(6.8), Inches(1.7), Inches(5.7), Inches(1.5), C_WHITE, C_BORDER)
tb12_e1 = slide12.shapes.add_textbox(Inches(7.0), Inches(1.8), Inches(5.3), Inches(1.3))
tf12_e1 = tb12_e1.text_frame
tf12_e1.word_wrap = True
add_bullet_list(tf12_e1, [
    ("Estratégia 1 (Nivelamento):  R$ 1.968.600,00", ""),
    ("• Diferença:", "+R$ 66.200,00 mais cara que a Híbrida (+3,5%)."),
    ("• Causa:", "R$ 108.000,00 desperdiçados em multas de atraso.")
], font_size=11, space_after=2)

card12_e2 = add_card(slide12, Inches(6.8), Inches(3.45), Inches(5.7), Inches(1.5), C_WHITE, C_BORDER)
tb12_e2 = slide12.shapes.add_textbox(Inches(7.0), Inches(3.55), Inches(5.3), Inches(1.3))
tf12_e2 = tb12_e2.text_frame
tf12_e2.word_wrap = True
add_bullet_list(tf12_e2, [
    ("Estratégia 2 (Acompanhamento):  R$ 1.921.200,00", ""),
    ("• Diferença:", "+R$ 18.800,00 mais cara que a Híbrida (+1,0%)."),
    ("• Causa:", "R$ 82.800,00 gastos em 36 admissões e demissões.")
], font_size=11, space_after=2)

card12_e3 = add_card(slide12, Inches(6.8), Inches(5.2), Inches(5.7), Inches(1.6), C_SUCCESS_LIGHT, C_SUCCESS)
tb12_e3 = slide12.shapes.add_textbox(Inches(7.0), Inches(5.3), Inches(5.3), Inches(1.4))
tf12_e3 = tb12_e3.text_frame
tf12_e3.word_wrap = True
add_bullet_list(tf12_e3, [
    ("Estratégia 3 (Híbrida - CAMPEÃ):  R$ 1.902.400,00", ""),
    ("• Economia Absoluta:", "Poupa R$ 66.200 vs E1 e R$ 18.800 vs E2."),
    ("• Eficiência:", "Menor custo financeiro + Zero demissões + 100% On-Time Delivery!")
], font_size=11, space_after=2)

slide12.notes_slide.notes_text_frame.text = (
    "[TEMPO: 11:30 - 12:30 | Integrante 4]\n"
    "Os números finais comprovam: a Estratégia Híbrida vence em todas as frentes. "
    "Custa R$ 1.902.400,00, economizando R$ 66.200 em relação ao Nivelamento e R$ 18.800 em relação ao Acompanhamento. "
    "A Híbrida é a solução mais barata e a mais eficiente porque aloca os recursos certos na hora certa: "
    "mão de obra própria onde ela é barata, horas extras controladas e facção parceira apenas para absorver o topo dos picos."
)

# ==============================================================================
# SLIDE 13: ALINHAMENTO COM OS 5 OBJETIVOS DE SLACK
# ==============================================================================
slide13 = create_modern_slide("Gestão de Operações", "Alinhamento com os 5 Objetivos de Desempenho", "Avaliação multicritério segundo o modelo teórico de Nigel Slack", 13)

objectives = [
    ("1. CUSTO", "R$ 1.902.400", "Menor custo global anual; elimina desperdícios com multas e rescisões contratuais.", C_PRIMARY),
    ("2. CONFIABILIDADE", "100%", "100% dos pedidos corporativos faturados e entregues rigorosamente na data acordada.", C_SUCCESS),
    ("3. QUALIDADE", "< 1,5%", "Equipe estável de 19 costureiras experientes reduz retrabalho e refugos no gargalo.", C_PURPLE),
    ("4. FLEXIBILIDADE", "Elástica", "Combinação sinérgica de horas extras reguladas e facção externa nos picos sazonais.", C_WARNING),
    ("5. RAPIDEZ", "Imediata", "Estoque de segurança desacopla a demanda no 2º semestre, viabilizando pronta-entrega.", C_NAVY_DARK)
]

for idx, (obj_title, kpi_val, obj_desc, obj_color) in enumerate(objectives):
    top_pos = Inches(1.7 + idx * 1.02)
    card_obj = add_card(slide13, Inches(0.8), top_pos, Inches(11.733), Inches(0.88), C_WHITE, C_BORDER)
    
    # Pill de Título do Objetivo
    tb_t = slide13.shapes.add_textbox(Inches(1.0), top_pos + Inches(0.12), Inches(2.6), Inches(0.65))
    tf_t = tb_t.text_frame
    p_t = tf_t.paragraphs[0]
    p_t.text = obj_title
    p_t.font.name = 'Arial'
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = obj_color
    
    # KPI Value
    tb_k = slide13.shapes.add_textbox(Inches(3.7), top_pos + Inches(0.12), Inches(2.2), Inches(0.65))
    tf_k = tb_k.text_frame
    p_k = tf_k.paragraphs[0]
    p_k.text = kpi_val
    p_k.font.name = 'Arial'
    p_k.font.size = Pt(14)
    p_k.font.bold = True
    p_k.font.color.rgb = C_TEXT_DARK
    
    # Descrição do Objetivo
    tb_d = slide13.shapes.add_textbox(Inches(6.0), top_pos + Inches(0.12), Inches(6.3), Inches(0.65))
    tf_d = tb_d.text_frame
    tf_d.word_wrap = True
    p_d = tf_d.paragraphs[0]
    p_d.text = obj_desc
    p_d.font.name = 'Arial'
    p_d.font.size = Pt(11)
    p_d.font.color.rgb = C_TEXT_BODY

slide13.notes_slide.notes_text_frame.text = (
    "[TEMPO: 12:30 - 13:30 | Integrante 5]\n"
    "Analisando sob os cinco objetivos de desempenho de Nigel Slack, a Estratégia Híbrida atinge pontuação máxima: "
    "Custo mínimo de R$ 1,902 milhão; Confiabilidade de 100% nas entregas; Qualidade garantida pela estabilidade da mão de obra; "
    "Flexibilidade conferida pelas facções e rapidez viabilizada pelo estoque pulmão. "
    "É a confirmação de que uma decisão de PCP impacta todos os indicadores estratégicos da empresa."
)

# ==============================================================================
# SLIDE 14: PLANO DE AÇÃO PARA A DIRETORIA
# ==============================================================================
slide14 = create_modern_slide("Plano de Ação", "Diretrizes de Implantação para a Diretoria Industrial", "Roteiro executivo em 4 passos para viabilizar a Estratégia Híbrida na fábrica", 14)

steps = [
    ("Passo 1", "Estabilização da Equipe Base", "Fixar o quadro em 19 costureiras registradas, eliminando o turnover e retendo conhecimento técnico de costura pesada."),
    ("Passo 2", "Pactuação Prévia de Horas Extras", "Formalizar acordo com o sindicato laboral para até 2h extras/dia exclusivamente em Jan-Mar (respeito estrito ao Art. 59 da CLT)."),
    ("Passo 3", "Homologação de Facções com Corte Interno", "Contratar facções fornecendo o tecido cortado pela Vértice, garantindo padrão de qualidade e afastando riscos da Súmula 331 do TST."),
    ("Passo 4", "Implantação do Processo de S&OP", "Instituir reuniões mensais de Sales and Operations Planning entre Comercial e PCP, com horizonte rolante de 12 meses.")
]

for idx, (st_num, st_title, st_desc) in enumerate(steps):
    top_pos = Inches(1.7 + idx * 1.25)
    card_st = add_card(slide14, Inches(0.8), top_pos, Inches(11.733), Inches(1.1), C_WHITE, C_BORDER)
    
    # Número do Passo
    tb_n = slide14.shapes.add_textbox(Inches(1.0), top_pos + Inches(0.18), Inches(1.8), Inches(0.75))
    tf_n = tb_n.text_frame
    p_n = tf_n.paragraphs[0]
    p_n.text = st_num.upper()
    p_n.font.name = 'Arial'
    p_n.font.size = Pt(14)
    p_n.font.bold = True
    p_n.font.color.rgb = C_PRIMARY
    
    # Título do Passo
    tb_t = slide14.shapes.add_textbox(Inches(2.9), top_pos + Inches(0.18), Inches(3.8), Inches(0.75))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True
    p_t = tf_t.paragraphs[0]
    p_t.text = st_title
    p_t.font.name = 'Arial'
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = C_TEXT_DARK
    
    # Descrição do Passo
    tb_d = slide14.shapes.add_textbox(Inches(6.8), top_pos + Inches(0.15), Inches(5.5), Inches(0.8))
    tf_d = tb_d.text_frame
    tf_d.word_wrap = True
    p_d = tf_d.paragraphs[0]
    p_d.text = st_desc
    p_d.font.name = 'Arial'
    p_d.font.size = Pt(11)
    p_d.font.color.rgb = C_TEXT_BODY
    p_d.line_spacing = 1.15

slide14.notes_slide.notes_text_frame.text = (
    "[TEMPO: 13:30 - 14:15 | Integrante 5]\n"
    "Para transformar o modelo em realidade fabril, traçamos um roteiro em quatro passos práticos: "
    "Fixar 19 operadoras estáveis; pactuar com antecedência as horas extras no sindicato para Jan a Mar; "
    "homologar facções fornecendo o tecido cortado internamente para proteger a empresa contra passivos da Súmula 331 do TST; "
    "e instituir reuniões mensais de S&OP entre Vendas e Produção com horizonte rolante."
)

# ==============================================================================
# SLIDE 15: CONCLUSÃO E REFERÊNCIAS
# ==============================================================================
slide15 = create_modern_slide("Fechamento", "Conclusões Estratégicas e Referências Bibliográficas", "Síntese dos aprendizados do projeto e abertura para a arguição da banca", 15)

card15_l = add_card(slide15, Inches(0.8), Inches(1.7), Inches(5.7), Inches(5.1), C_WHITE, C_BORDER)
tb15_l = slide15.shapes.add_textbox(Inches(1.0), Inches(1.9), Inches(5.3), Inches(4.7))
tf15_l = tb15_l.text_frame
tf15_l.word_wrap = True
add_bullet_list(tf15_l, [
    ("Principais Conclusões do Trabalho:", ""),
    ("• O PAP como Ferramenta Vital:", "Harmonizou a área Comercial (que exige pronta-entrega) e a Manufatura (que busca estabilidade)."),
    ("• Armadilhas dos Extremos Evitadas:", "Nivelamento gera colapso de entregas (R$ 108 mil); Acompanhamento gera caos de rotatividade (36 movimentações)."),
    ("• Superioridade da Solução Híbrida:", "R$ 1.902.400,00, 100% de nível de serviço, zero demissões, plena regularidade com a CLT e 800 peças em estoque de segurança.")
], font_size=12, space_after=8)

card15_r = add_card(slide15, Inches(6.8), Inches(1.7), Inches(5.7), Inches(5.1), C_WHITE, C_BORDER)
tb15_r = slide15.shapes.add_textbox(Inches(7.0), Inches(1.9), Inches(5.3), Inches(4.7))
tf15_r = tb15_r.text_frame
tf15_r.word_wrap = True
add_bullet_list(tf15_r, [
    ("Principais Referências Bibliográficas:", ""),
    ("• SLACK, N. et al.:", "Administração da Produção. 8. ed. Atlas, 2018."),
    ("• CORRÊA, H. L.; CORRÊA, C. A.:", "Administração de Produção e Operações, 2017."),
    ("• QUEIROZ, J. V.:", "PAP em Manufatura – Aula 03. UnB / FT, 2026."),
    ("• GONÇALVES, M. C.:", "Conceitos Básicos de PCP – Aula 01. UnB / FT, 2026."),
    ("• BARBOSA, C. P.:", "PAP em Serviços – Aula 02. UnB / FT, 2026."),
    ("• SEBRAE:", "Inteligência Setorial: PCP no Vestuário, 2024."),
    ("• BRASIL:", "Consolidação das Leis do Trabalho (CLT) – Art. 59, 1943."),
    ("Agradecimentos:", "Muito obrigado a todos! Estamos à disposição para as perguntas da banca.")
], font_size=10.5, space_after=4)

slide15.notes_slide.notes_text_frame.text = (
    "[TEMPO: 14:15 - 15:00 | Integrante 5]\n"
    "Concluímos destacando que o PAP provou ser o coração tático de uma fábrica de manufatura. "
    "Agradecemos aos professores Marcelo Gonçalves, Camilla Pimenta e João Victor Queiroz pela condução da disciplina. "
    "Encerramos rigorosamente aos 14 minutos e 45 segundos e estamos à disposição da banca para responder às perguntas!"
)

# Salvar apresentação na raiz e na pasta Entregas/
prs.save("Apresentacao_Seminario_PAP_Manufatura.pptx")
prs.save("Entregas/Apresentacao_Seminario_PAP_Manufatura.pptx")
print("Apresentação moderna e minimalista gerada com sucesso em Entregas/Apresentacao_Seminario_PAP_Manufatura.pptx!")
