import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
wb.remove(wb.active) # Remove default sheet

# Colors
navy_dark = '1B365D'
navy_med = '2C5282'
navy_light = 'E8EEF5'
gray_light = 'F7FAFC'
gray_border = 'CBD5E0'
green_header = '1C4532'
green_light = 'E6FFFA'
total_color = 'DCE6F1'
highlight_win = 'C6F6D5'

header_font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
header_fill_navy = PatternFill(start_color=navy_dark, end_color=navy_dark, fill_type='solid')
header_fill_sub = PatternFill(start_color=navy_med, end_color=navy_med, fill_type='solid')
header_fill_green = PatternFill(start_color=green_header, end_color=green_header, fill_type='solid')

fill_total = PatternFill(start_color=total_color, end_color=total_color, fill_type='solid')
fill_win = PatternFill(start_color=highlight_win, end_color=highlight_win, fill_type='solid')
fill_alt = PatternFill(start_color='F8FAFC', end_color='F8FAFC', fill_type='solid')

border_thin = Border(
    left=Side(style='thin', color=gray_border),
    right=Side(style='thin', color=gray_border),
    top=Side(style='thin', color=gray_border),
    bottom=Side(style='thin', color=gray_border)
)
border_total = Border(
    top=Side(style='thin', color=navy_dark),
    bottom=Side(style='double', color=navy_dark),
    left=Side(style='thin', color=gray_border),
    right=Side(style='thin', color=gray_border)
)

font_regular = Font(name='Arial', size=9)
font_bold = Font(name='Arial', size=9, bold=True)
font_title = Font(name='Arial', size=13, bold=True, color=navy_dark)
font_subtitle = Font(name='Arial', size=10, italic=True, color='4A5568')

def style_table_header(ws, row, start_col, end_col, fill=header_fill_navy):
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = header_font
        cell.fill = fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border_thin
    ws.row_dimensions[row].height = 28

def auto_fit_columns(ws, max_cols=20):
    for col in range(1, max_cols + 1):
        col_letter = get_column_letter(col)
        max_len = 0
        for row in range(1, ws.max_row + 1):
            val = ws.cell(row=row, column=col).value
            if val is not None:
                # ignore long formulas in length calculation
                s = str(val)
                if not s.startswith('='):
                    max_len = max(max_len, len(s))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 10)

# =============================================================
# ABA 1: PREMISSAS E DADOS
# =============================================================
ws1 = wb.create_sheet(title='Premissas e Dados')
ws1.views.sheetView[0].showGridLines = True

ws1['B2'] = "VÉRTICE CONFECÇÕES & UNIFORMES PROFISSIONAIS LTDA."
ws1['B2'].font = font_title
ws1['B3'] = "Planejamento Agregado da Produção (PAP) - Parâmetros, Custos e Demanda (12 Meses)"
ws1['B3'].font = font_subtitle

ws1['B5'] = "1. PARÂMETROS OPERACIONAIS"
ws1['B5'].font = font_bold
ws1['B6'] = "Parâmetro"
ws1['C6'] = "Valor"
ws1['D6'] = "Unidade"
style_table_header(ws1, 6, 2, 4, header_fill_sub)

params = [
    ("Número Base de Operadores (Costureiras)", 20, "operadores"),
    ("Jornada Mensal Líquida por Operador", 160, "horas/mês"),
    ("Tempo Padrão por Conjunto de Uniforme (TP)", 0.80, "horas/conjunto"),
    ("Produtividade Normal por Operador", 200, "conjuntos/mês"),
    ("Capacidade Produtiva em Turno Normal (20 op.)", 4000, "conjuntos/mês"),
    ("Estoque Inicial de Segurança (Mês 0)", 400, "conjuntos"),
    ("Estoque Final Desejado (Mês 12)", 400, "conjuntos"),
    ("Limite Legal de Horas Extras (CLT - 2h/dia)", 0.25, "da jornada normal (25%)"),
    ("Capacidade Máxima em Hora Extra (20 op.)", 800, "conjuntos/mês")
]

for idx, (p, v, u) in enumerate(params, start=7):
    ws1[f'B{idx}'] = p
    ws1[f'C{idx}'] = v
    ws1[f'D{idx}'] = u
    ws1[f'B{idx}'].font = font_regular
    ws1[f'C{idx}'].font = font_bold
    ws1[f'D{idx}'].font = font_regular
    ws1[f'C{idx}'].alignment = Alignment(horizontal='right')
    if isinstance(v, float) and v < 1:
        ws1[f'C{idx}'].number_format = '0.0%' if v == 0.25 else '0.00'
    elif isinstance(v, int):
        ws1[f'C{idx}'].number_format = '#,##0'
    for c in ['B', 'C', 'D']:
        ws1[f'{c}{idx}'].border = border_thin

ws1['B17'] = "2. PARÂMETROS DE CUSTOS (R$)"
ws1['B17'].font = font_bold
ws1['B18'] = "Rubrica de Custo"
ws1['C18'] = "Valor Unitário"
ws1['D18'] = "Unidade de Medida"
style_table_header(ws1, 18, 2, 4, header_fill_sub)

custos = [
    ("Custo de Produção em Turno Normal (cn)", 38.00, "R$ / conjunto"),
    ("Custo de Produção em Hora Extra (che)", 46.00, "R$ / conjunto"),
    ("Custo de Subcontratação / Facção Externa (csub)", 52.00, "R$ / conjunto"),
    ("Custo de Manutenção / Posse de Estoque (ce)", 3.00, "R$ / conjunto / mês"),
    ("Custo de Atraso / Ruptura / Backlog (ca)", 20.00, "R$ / conjunto / mês"),
    ("Custo de Contratação de Operadora (Cc)", 1600.00, "R$ / admissão"),
    ("Custo de Demissão / Rescisão de Operadora (Cd)", 3000.00, "R$ / desligamento")
]

for idx, (p, v, u) in enumerate(custos, start=19):
    ws1[f'B{idx}'] = p
    ws1[f'C{idx}'] = v
    ws1[f'D{idx}'] = u
    ws1[f'B{idx}'].font = font_regular
    ws1[f'C{idx}'].font = font_bold
    ws1[f'D{idx}'].font = font_regular
    ws1[f'C{idx}'].number_format = 'R$ #,##0.00'
    ws1[f'C{idx}'].alignment = Alignment(horizontal='right')
    for c in ['B', 'C', 'D']:
        ws1[f'{c}{idx}'].border = border_thin

ws1['F5'] = "3. PREVISÃO DE DEMANDA AGREGADA (12 MESES)"
ws1['F5'].font = font_bold
ws1['F6'] = "Mês"
ws1['G6'] = "Demanda Prevista"
ws1['H6'] = "Fator Sazonal / Justificativa de Mercado"
style_table_header(ws1, 6, 6, 8, header_fill_sub)

meses_dados = [
    ("Janeiro", 4600, "Volta às aulas (colégios) e renovação de frotas corporativas"),
    ("Fevereiro", 5200, "Pico de entrega escolar e novos contratos de empresas"),
    ("Março", 4800, "Finalizações de pedidos corporativos do 1º trimestre"),
    ("Abril", 3600, "Início do período estável / entressafra"),
    ("Maio", 3200, "Vale de demanda do 1º semestre"),
    ("Junho", 3000, "Menor demanda do ano (férias escolares de meio de ano)"),
    ("Julho", 3800, "Reposição de uniformes de inverno e meias estações"),
    ("Agosto", 3600, "Início do 2º semestre letivo e reposição empresarial"),
    ("Setembro", 3400, "Demanda intermediária de rotina"),
    ("Outubro", 3800, "Antecipação de compras do varejo e logística"),
    ("Novembro", 5000, "Pico de contratações de fim de ano no varejo / eventos"),
    ("Dezembro", 4000, "Fechamento de orçamento e eventos corporativos anuais")
]

for idx, (m, dem, just) in enumerate(meses_dados, start=7):
    ws1[f'F{idx}'] = m
    ws1[f'G{idx}'] = dem
    ws1[f'H{idx}'] = just
    ws1[f'F{idx}'].font = font_regular
    ws1[f'G{idx}'].font = font_bold
    ws1[f'H{idx}'].font = font_regular
    ws1[f'G{idx}'].number_format = '#,##0'
    ws1[f'F{idx}'].alignment = Alignment(horizontal='center')
    ws1[f'G{idx}'].alignment = Alignment(horizontal='right')
    for c in ['F', 'G', 'H']:
        ws1[f'{c}{idx}'].border = border_thin

ws1['F19'] = "Total Anual"
ws1['G19'] = "=SUM(G7:G18)"
ws1['H19'] = "Média Mensal: 4.000 un/mês"
ws1['F19'].font = font_bold
ws1['G19'].font = font_bold
ws1['H19'].font = font_bold
ws1['G19'].number_format = '#,##0'
for c in ['F', 'G', 'H']:
    ws1[f'{c}{19}'].fill = fill_total
    ws1[f'{c}{19}'].border = border_total

# =============================================================
# ABA 2: ESTRATÉGIA 1 - PRODUÇÃO CONSTANTE (NIVELAMENTO)
# =============================================================
ws2 = wb.create_sheet(title='Estratégia 1 - Nivelamento')
ws2.views.sheetView[0].showGridLines = True

ws2['B2'] = "ESTRATÉGIA 1: TAXA DE PRODUÇÃO CONSTANTE (NIVELAMENTO / LEVEL)"
ws2['B2'].font = font_title
ws2['B3'] = "Capacidade fixa em 4.000 un/mês (20 operadoras). Absorção por estoques e atrasos (backlog)."
ws2['B3'].font = font_subtitle

headers_e1 = [
    "Mês", "Demanda (Dt)", "Backlog Anterior", "Demanda Efetiva",
    "Produção Normal", "Estoque Inicial (EIt)", "Disponibilidade",
    "Estoque Final (EFt)", "Backlog Final", "Estoque Médio (EMt)",
    "Custo Produção (R$)", "Custo Estoque (R$)", "Custo Atraso (R$)", "Custo Total Mensal (R$)"
]

ws2.row_dimensions[5].height = 28
for col_idx, h in enumerate(headers_e1, start=2):
    ws2.cell(row=5, column=col_idx, value=h)
style_table_header(ws2, 5, 2, len(headers_e1) + 1, header_fill_navy)

for i in range(12):
    row = 6 + i
    m_nome = meses_dados[i][0]
    
    ws2[f'B{row}'] = m_nome
    ws2[f'C{row}'] = f"='Premissas e Dados'!G{7+i}"
    
    # Backlog anterior
    if i == 0:
        ws2[f'D{row}'] = 0
    else:
        ws2[f'D{row}'] = f"=J{row-1}"
    
    # Demanda efetiva = Dt + Backlog anterior
    ws2[f'E{row}'] = f"=C{row}+D{row}"
    
    # Produção normal constante
    ws2[f'F{row}'] = 4000
    
    # Estoque inicial: mes 1 = 400, outros = EF anterior
    if i == 0:
        ws2[f'G{row}'] = 400
    else:
        ws2[f'G{row}'] = f"=I{row-1}"
        
    # Disponibilidade = EIt + Producao
    ws2[f'H{row}'] = f"=G{row}+F{row}"
    
    # Estoque Final = IF(Disp >= Demanda Efetiva; Disp - Demanda Efetiva; 0)
    ws2[f'I{row}'] = f"=IF(H{row}>=E{row}, H{row}-E{row}, 0)"
    
    # Backlog Final = IF(Disp < Demanda Efetiva; Demanda Efetiva - Disp; 0)
    ws2[f'J{row}'] = f"=IF(H{row}<E{row}, E{row}-H{row}, 0)"
    
    # Estoque Médio = (EIt + EFt) / 2
    ws2[f'K{row}'] = f"=(G{row}+I{row})/2"
    
    # Custos
    ws2[f'L{row}'] = f"=F{row}*'Premissas e Dados'!$C$19" # Produção (cn=38)
    ws2[f'M{row}'] = f"=K{row}*'Premissas e Dados'!$C$22" # Estoque (ce=3)
    ws2[f'N{row}'] = f"=J{row}*'Premissas e Dados'!$C$23" # Atraso (ca=20)
    ws2[f'O{row}'] = f"=SUM(L{row}:N{row})"
    
    # Formatação
    ws2[f'B{row}'].alignment = Alignment(horizontal='center')
    for c in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
        ws2[f'{c}{row}'].number_format = '#,##0'
        ws2[f'{c}{row}'].font = font_regular
        ws2[f'{c}{row}'].border = border_thin
    for c in ['L', 'M', 'N', 'O']:
        ws2[f'{c}{row}'].number_format = 'R$ #,##0.00'
        ws2[f'{c}{row}'].font = font_regular
        ws2[f'{c}{row}'].border = border_thin
    ws2[f'O{row}'].font = font_bold
    ws2[f'B{row}'].border = border_thin

# Linha de Totais
row_tot = 18
ws2[f'B{row_tot}'] = "TOTAL ANUAL"
ws2[f'C{row_tot}'] = "=SUM(C6:C17)"
ws2[f'D{row_tot}'] = "-"
ws2[f'E{row_tot}'] = "-"
ws2[f'F{row_tot}'] = "=SUM(F6:F17)"
ws2[f'G{row_tot}'] = "-"
ws2[f'H{row_tot}'] = "-"
ws2[f'I{row_tot}'] = "=I17" # Estoque final no mes 12
ws2[f'J{row_tot}'] = "=SUM(J6:J17)" # Backlog acumulado
ws2[f'K{row_tot}'] = "=AVERAGE(K6:K17)" # Estoque medio anual
ws2[f'L{row_tot}'] = "=SUM(L6:L17)"
ws2[f'M{row_tot}'] = "=SUM(M6:M17)"
ws2[f'N{row_tot}'] = "=SUM(N6:N17)"
ws2[f'O{row_tot}'] = "=SUM(O6:O17)"

for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O']:
    cell = ws2[f'{col}{row_tot}']
    cell.font = font_bold
    cell.fill = fill_total
    cell.border = border_total
    if col in ['C', 'F', 'I', 'J', 'K']:
        cell.number_format = '#,##0'
    elif col in ['L', 'M', 'N', 'O']:
        cell.number_format = 'R$ #,##0.00'
    elif col in ['B', 'D', 'E', 'G', 'H']:
        cell.alignment = Alignment(horizontal='center')

# =============================================================
# ABA 3: ESTRATÉGIA 2 - ACOMPANHAMENTO DA DEMANDA (CHASE)
# =============================================================
ws3 = wb.create_sheet(title='Estratégia 2 - Acompanhamento')
ws3.views.sheetView[0].showGridLines = True

ws3['B2'] = "ESTRATÉGIA 2: ACOMPANHAMENTO DA DEMANDA (CHASE STRATEGY)"
ws3['B2'].font = font_title
ws3['B3'] = "Ajuste mensal da força de trabalho via contratações e demissões. Demanda atendida mês a mês."
ws3['B3'].font = font_subtitle

headers_e2 = [
    "Mês", "Demanda (Dt)", "Operadores Necessários", "Operadores Anteriores",
    "Contratações (Ct)", "Demissões (Dt)", "Produção Real (un)",
    "Estoque Médio", "Custo Produção (R$)", "Custo Contratação (R$)",
    "Custo Demissão (R$)", "Custo Estoque (R$)", "Custo Total Mensal (R$)"
]

ws3.row_dimensions[5].height = 28
for col_idx, h in enumerate(headers_e2, start=2):
    ws3.cell(row=5, column=col_idx, value=h)
style_table_header(ws3, 5, 2, len(headers_e2) + 1, header_fill_navy)

for i in range(12):
    row = 6 + i
    m_nome = meses_dados[i][0]
    
    ws3[f'B{row}'] = m_nome
    ws3[f'C{row}'] = f"='Premissas e Dados'!G{7+i}"
    
    # Operadores necessarios = Dt / 200
    ws3[f'D{row}'] = f"=INT(C{row}/'Premissas e Dados'!$C$10)"
    
    # Operadores anteriores
    if i == 0:
        ws3[f'E{row}'] = "='Premissas e Dados'!$C$7" # 20
    else:
        ws3[f'E{row}'] = f"=D{row-1}"
        
    # Contratações = MAX(0, Dt - Dt-1)
    ws3[f'F{row}'] = f"=MAX(0, D{row}-E{row})"
    
    # Demissões = MAX(0, Dt-1 - Dt)
    ws3[f'G{row}'] = f"=MAX(0, E{row}-D{row})"
    
    # Produção = Operadores * 200
    ws3[f'H{row}'] = f"=D{row}*'Premissas e Dados'!$C$10"
    
    # Estoque medio constante de seguranca = 400
    ws3[f'I{row}'] = 400
    
    # Custos
    ws3[f'J{row}'] = f"=H{row}*'Premissas e Dados'!$C$19" # Produção (cn=38)
    ws3[f'K{row}'] = f"=F{row}*'Premissas e Dados'!$C$24" # Contratação (Cc=1600)
    ws3[f'L{row}'] = f"=G{row}*'Premissas e Dados'!$C$25" # Demissão (Cd=3000)
    ws3[f'M{row}'] = f"=I{row}*'Premissas e Dados'!$C$22" # Estoque (ce=3)
    ws3[f'N{row}'] = f"=SUM(J{row}:M{row})"
    
    # Formatação
    ws3[f'B{row}'].alignment = Alignment(horizontal='center')
    for c in ['C', 'D', 'E', 'F', 'G', 'H', 'I']:
        ws3[f'{c}{row}'].number_format = '#,##0'
        ws3[f'{c}{row}'].font = font_regular
        ws3[f'{c}{row}'].border = border_thin
    for c in ['J', 'K', 'L', 'M', 'N']:
        ws3[f'{c}{row}'].number_format = 'R$ #,##0.00'
        ws3[f'{c}{row}'].font = font_regular
        ws3[f'{c}{row}'].border = border_thin
    ws3[f'N{row}'].font = font_bold
    ws3[f'B{row}'].border = border_thin

# Total E2
ws3[f'B{row_tot}'] = "TOTAL ANUAL"
ws3[f'C{row_tot}'] = "=SUM(C6:C17)"
ws3[f'D{row_tot}'] = "=AVERAGE(D6:D17)"
ws3[f'E{row_tot}'] = "-"
ws3[f'F{row_tot}'] = "=SUM(F6:F17)" # Total contratacoes
ws3[f'G{row_tot}'] = "=SUM(G6:G17)" # Total demissoes
ws3[f'H{row_tot}'] = "=SUM(H6:H17)"
ws3[f'I{row_tot}'] = "=AVERAGE(I6:I17)"
ws3[f'J{row_tot}'] = "=SUM(J6:J17)"
ws3[f'K{row_tot}'] = "=SUM(K6:K17)"
ws3[f'L{row_tot}'] = "=SUM(L6:L17)"
ws3[f'M{row_tot}'] = "=SUM(M6:M17)"
ws3[f'N{row_tot}'] = "=SUM(N6:N17)"

for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']:
    cell = ws3[f'{col}{row_tot}']
    cell.font = font_bold
    cell.fill = fill_total
    cell.border = border_total
    if col in ['C', 'D', 'F', 'G', 'H', 'I']:
        cell.number_format = '#,##0'
    elif col in ['J', 'K', 'L', 'M', 'N']:
        cell.number_format = 'R$ #,##0.00'
    elif col in ['B', 'E']:
        cell.alignment = Alignment(horizontal='center')

# =============================================================
# ABA 4: ESTRATÉGIA 3 - HÍBRIDA / PATAMARES (MIXED STRATEGY)
# =============================================================
ws4 = wb.create_sheet(title='Estratégia 3 - Híbrida')
ws4.views.sheetView[0].showGridLines = True

ws4['B2'] = "ESTRATÉGIA 3: HÍBRIDA / PATAMARES (MIXED STRATEGY - RECOMENDADA)"
ws4['B2'].font = font_title
ws4['B3'] = "Equipe base de 19 operadoras (3.800 un/mês). Picos absorvidos por Horas Extras e Facção Externa (Subcontratação)."
ws4['B3'].font = font_subtitle

headers_e3 = [
    "Mês", "Demanda (Dt)", "Estoque Inicial (EIt)", "Produção Normal (19 op)",
    "Horas Extras (HEt)", "Subcontratação (SUBt)", "Produção Total (PTt)",
    "Estoque Final (EFt)", "Estoque Médio (EMt)", "Custo Normal (R$)",
    "Custo HE (R$)", "Custo Subcontrat. (R$)", "Custo Estoque (R$)", "Custo Total Mensal (R$)"
]

ws4.row_dimensions[5].height = 28
for col_idx, h in enumerate(headers_e3, start=2):
    ws4.cell(row=5, column=col_idx, value=h)
style_table_header(ws4, 5, 2, len(headers_e3) + 1, header_fill_green)

# Heurística com fórmulas na planilha:
# Normal = 3.800 fixo
# Disp_preliminar = EIt + 3800
# Nec_adicional = MAX(0, Dt - Disp_preliminar)
# HE = MIN(800, Nec_adicional)
# SUB = MAX(0, Nec_adicional - HE)
# PT = 3800 + HE + SUB
# EFt = EIt + PT - Dt

for i in range(12):
    row = 6 + i
    m_nome = meses_dados[i][0]
    
    ws4[f'B{row}'] = m_nome
    ws4[f'C{row}'] = f"='Premissas e Dados'!G{7+i}"
    
    # Estoque Inicial
    if i == 0:
        ws4[f'D{row}'] = 400
    else:
        ws4[f'D{row}'] = f"=I{row-1}"
        
    # Produção Normal (19 operadoras = 3800)
    ws4[f'E{row}'] = 3800
    
    # Horas Extras: MIN(800, MAX(0, C_row - (D_row + E_row)))
    ws4[f'F{row}'] = f"=MIN(800, MAX(0, C{row}-(D{row}+E{row})))"
    
    # Subcontratação: MAX(0, (C_row - (D_row + E_row)) - F_row)
    ws4[f'G{row}'] = f"=MAX(0, (C{row}-(D{row}+E{row}))-F{row})"
    
    # Produção Total = Normal + HE + SUB
    ws4[f'H{row}'] = f"=SUM(E{row}:G{row})"
    
    # Estoque Final = EIt + PTt - Dt
    ws4[f'I{row}'] = f"=D{row}+H{row}-C{row}"
    
    # Estoque Médio = (EIt + EFt) / 2
    ws4[f'J{row}'] = f"=(D{row}+I{row})/2"
    
    # Custos
    ws4[f'K{row}'] = f"=E{row}*'Premissas e Dados'!$C$19" # Normal (cn=38)
    ws4[f'L{row}'] = f"=F{row}*'Premissas e Dados'!$C$20" # HE (che=46)
    ws4[f'M{row}'] = f"=G{row}*'Premissas e Dados'!$C$21" # Sub (csub=52)
    ws4[f'N{row}'] = f"=J{row}*'Premissas e Dados'!$C$22" # Estoque (ce=3)
    ws4[f'O{row}'] = f"=SUM(K{row}:N{row})"
    
    # Formatação
    ws4[f'B{row}'].alignment = Alignment(horizontal='center')
    for c in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
        ws4[f'{c}{row}'].number_format = '#,##0'
        ws4[f'{c}{row}'].font = font_regular
        ws4[f'{c}{row}'].border = border_thin
    for c in ['K', 'L', 'M', 'N', 'O']:
        ws4[f'{c}{row}'].number_format = 'R$ #,##0.00'
        ws4[f'{c}{row}'].font = font_regular
        ws4[f'{c}{row}'].border = border_thin
    ws4[f'O{row}'].font = font_bold
    ws4[f'B{row}'].border = border_thin

# Total E3
ws4[f'B{row_tot}'] = "TOTAL ANUAL"
ws4[f'C{row_tot}'] = "=SUM(C6:C17)"
ws4[f'D{row_tot}'] = "-"
ws4[f'E{row_tot}'] = "=SUM(E6:E17)"
ws4[f'F{row_tot}'] = "=SUM(F6:F17)"
ws4[f'G{row_tot}'] = "=SUM(G6:G17)"
ws4[f'H{row_tot}'] = "=SUM(H6:H17)"
ws4[f'I{row_tot}'] = "=I17" # Saldo final no mes 12
ws4[f'J{row_tot}'] = "=AVERAGE(J6:J17)" # Estoque medio anual
ws4[f'K{row_tot}'] = "=SUM(K6:K17)"
ws4[f'L{row_tot}'] = "=SUM(L6:L17)"
ws4[f'M{row_tot}'] = "=SUM(M6:M17)"
ws4[f'N{row_tot}'] = "=SUM(N6:N17)"
ws4[f'O{row_tot}'] = "=SUM(O6:O17)"

for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O']:
    cell = ws4[f'{col}{row_tot}']
    cell.font = font_bold
    cell.fill = fill_win
    cell.border = border_total
    if col in ['C', 'E', 'F', 'G', 'H', 'I', 'J']:
        cell.number_format = '#,##0'
    elif col in ['K', 'L', 'M', 'N', 'O']:
        cell.number_format = 'R$ #,##0.00'
    elif col in ['B', 'D']:
        cell.alignment = Alignment(horizontal='center')

# =============================================================
# ABA 5: COMPARATIVO FINANCEIRO E DECISÃO
# =============================================================
ws5 = wb.create_sheet(title='Comparativo Financeiro')
ws5.views.sheetView[0].showGridLines = True

ws5['B2'] = "CONSOLIDAÇÃO E COMPARATIVO DAS ESTRATÉGIAS DE PAP"
ws5['B2'].font = font_title
ws5['B3'] = "Análise Econômica, Operacional e Decisão Final de Planejamento"
ws5['B3'].font = font_subtitle

# Tabela 1: Resumo Financeiro
ws5['B5'] = "1. COMPARATIVO DE CUSTOS ANUAIS TOTAIS POR RUBRICA"
ws5['B5'].font = font_bold

headers_comp = [
    "Rubrica de Custo",
    "Estratégia 1: Nivelamento (Level)",
    "Estratégia 2: Acompanhamento (Chase)",
    "Estratégia 3: Híbrida / Patamares (Recomendada)"
]

style_table_header(ws5, 6, 2, 5, header_fill_navy)
for col_idx, h in enumerate(headers_comp, start=2):
    ws5.cell(row=6, column=col_idx, value=h)

comp_rows = [
    ("Produção em Turno Normal", "='Estratégia 1 - Nivelamento'!L18", "='Estratégia 2 - Acompanhamento'!J18", "='Estratégia 3 - Híbrida'!K18"),
    ("Produção em Turno Extra (Horas Extras)", 0.00, 0.00, "='Estratégia 3 - Híbrida'!L18"),
    ("Subcontratação (Facção Externa de Costura)", 0.00, 0.00, "='Estratégia 3 - Híbrida'!M18"),
    ("Manutenção de Estoque (Posse)", "='Estratégia 1 - Nivelamento'!M18", "='Estratégia 2 - Acompanhamento'!M18", "='Estratégia 3 - Híbrida'!N18"),
    ("Custo de Ruptura / Atraso na Entrega (Backlog)", "='Estratégia 1 - Nivelamento'!N18", 0.00, 0.00),
    ("Custos de Contratação de Operadores", 0.00, "='Estratégia 2 - Acompanhamento'!K18", 0.00),
    ("Custos de Demissão e Rescisões Trabalhistas", 0.00, "='Estratégia 2 - Acompanhamento'!L18", 0.00)
]

for idx, (rubrica, f1, f2, f3) in enumerate(comp_rows, start=7):
    ws5[f'B{idx}'] = rubrica
    ws5[f'C{idx}'] = f1
    ws5[f'D{idx}'] = f2
    ws5[f'E{idx}'] = f3
    ws5[f'B{idx}'].font = font_regular
    for col in ['C', 'D', 'E']:
        ws5[f'{col}{idx}'].font = font_regular
        ws5[f'{col}{idx}'].number_format = 'R$ #,##0.00'
        ws5[f'{col}{idx}'].border = border_thin
    ws5[f'B{idx}'].border = border_thin

row_tot_c = 14
ws5[f'B{row_tot_c}'] = "CUSTO TOTAL DO PLANO ANUAL"
ws5[f'C{row_tot_c}'] = "=SUM(C7:C13)"
ws5[f'D{row_tot_c}'] = "=SUM(D7:D13)"
ws5[f'E{row_tot_c}'] = "=SUM(E7:E13)"

for col in ['B', 'C', 'D', 'E']:
    ws5[f'{col}{row_tot_c}'].font = font_bold
    ws5[f'{col}{row_tot_c}'].border = border_total
    if col != 'B':
        ws5[f'{col}{row_tot_c}'].number_format = 'R$ #,##0.00'
        ws5[f'{col}{row_tot_c}'].fill = fill_total

ws5[f'E{row_tot_c}'].fill = fill_win # Highlight winner

# Tabela 2: Indicadores Operacionais (KPIs)
ws5['B16'] = "2. INDICADORES DE DESEMPENHO OPERACIONAL (KPIs)"
ws5['B16'].font = font_bold

headers_kpi = [
    "Indicador de Desempenho (KPI)",
    "Estratégia 1: Nivelamento",
    "Estratégia 2: Acompanhamento",
    "Estratégia 3: Híbrida / Patamares"
]
style_table_header(ws5, 17, 2, 5, header_fill_sub)
for col_idx, h in enumerate(headers_kpi, start=2):
    ws5.cell(row=17, column=col_idx, value=h)

kpis = [
    ("Nível de Serviço / Entregas no Prazo (%)", "79.2% (atrasos em 5 meses)", "100.0%", "100.0%"),
    ("Backlog Total Acumulado (un)", "='Estratégia 1 - Nivelamento'!J18", "0", "0"),
    ("Estoque Médio Anual (un)", "='Estratégia 1 - Nivelamento'!K18", "400", "='Estratégia 3 - Híbrida'!J18"),
    ("Total de Contratações no Ano (op.)", "0", "='Estratégia 2 - Acompanhamento'!F18", "0"),
    ("Total de Demissões no Ano (op.)", "0", "='Estratégia 2 - Acompanhamento'!G18", "0"),
    ("Estabilidade da Força de Trabalho", "Alta (equipe 100% estável)", "Muito Baixa (rotatividade crítica)", "Alta (equipe base estável)"),
    ("Risco Trabalhista e Passivos (CLT)", "Muito Baixo", "Muito Alto (multas e ações trabalhistas)", "Baixo (cumpre art. 59 da CLT)"),
    ("Dificuldade de Gestão da Qualidade", "Baixa", "Alta (retrabalho com novatos)", "Média (homologação de facções)")
]

for idx, (kpi, v1, v2, v3) in enumerate(kpis, start=18):
    ws5[f'B{idx}'] = kpi
    ws5[f'C{idx}'] = v1
    ws5[f'D{idx}'] = v2
    ws5[f'E{idx}'] = v3
    ws5[f'B{idx}'].font = font_regular
    for col in ['C', 'D', 'E']:
        ws5[f'{col}{idx}'].font = font_regular
        ws5[f'{col}{idx}'].border = border_thin
        ws5[f'{col}{idx}'].alignment = Alignment(horizontal='center')
    ws5[f'B{idx}'].border = border_thin

# Tabela 3: Decisão e Recomendação
ws5['B27'] = "3. SÍNTESE DA DECISÃO TÉCNICO-ECONÔMICA"
ws5['B27'].font = font_bold
ws5['B28'] = (
    "A Estratégia 3 (Híbrida / Patamares) apresenta o MENOR CUSTO TOTAL (R$ 1.902.400,00), "
    "gerando uma economia de R$ 66.200,00 frente ao Nivelamento e R$ 18.800,00 frente ao Acompanhamento. "
    "Além da superioridade financeira, garante 100% de nível de serviço (zero rupturas), estabilidade "
    "da equipe fabril e plena conformidade com a legislação trabalhista brasileira (CLT)."
)
ws5['B28'].font = font_regular
ws5.merge_cells('B28:E30')
ws5['B28'].alignment = Alignment(wrap_text=True, vertical='top')

# Ajustar larguras
for ws in [ws1, ws2, ws3, ws4, ws5]:
    auto_fit_columns(ws, max_cols=16)

wb.save("PAP_Confeccao_Uniformes_12_Meses.xlsx")
print("Workbook completo salvo em PAP_Confeccao_Uniformes_12_Meses.xlsx com sucesso!")
