import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = openpyxl.Workbook()
wb.remove(wb.active) # Remove default sheet

# Paleta de Cores Executiva
navy_dark = '1B365D'
navy_med = '2C5282'
gray_border = 'CBD5E0'
green_header = '1C4532'
total_color = 'DCE6F1'
highlight_win = 'C6F6D5'
section_fill = 'EDF2F7'

font_title = Font(name='Arial', size=13, bold=True, color=navy_dark)
font_subtitle = Font(name='Arial', size=10, italic=True, color='4A5568')
font_sec = Font(name='Arial', size=10, bold=True, color=navy_dark)
font_hdr = Font(name='Arial', size=9.5, bold=True, color='FFFFFF')
font_bold = Font(name='Arial', size=9, bold=True)
font_regular = Font(name='Arial', size=9)

fill_hdr_navy = PatternFill(start_color=navy_dark, end_color=navy_dark, fill_type='solid')
fill_hdr_green = PatternFill(start_color=green_header, end_color=green_header, fill_type='solid')
fill_sec = PatternFill(start_color=section_fill, end_color=section_fill, fill_type='solid')
fill_total = PatternFill(start_color=total_color, end_color=total_color, fill_type='solid')
fill_win = PatternFill(start_color=highlight_win, end_color=highlight_win, fill_type='solid')

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

meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
cols_meses = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']
col_tot = 'O'

def setup_sheet_header(ws, title, subtitle):
    ws.views.sheetView[0].showGridLines = True
    ws['B2'] = "VÉRTICE CONFECÇÕES & UNIFORMES PROFISSIONAIS LTDA."
    ws['B2'].font = font_title
    ws['B3'] = title
    ws['B3'].font = font_title
    ws['B4'] = subtitle
    ws['B4'].font = font_subtitle

def apply_row_style(ws, row, is_header=False, is_section=False, is_total=False, is_win=False, num_fmt=None):
    ws.row_dimensions[row].height = 20 if not is_header else 25
    for col_idx in range(2, 16):
        c = ws.cell(row=row, column=col_idx)
        c.border = border_total if is_total else border_thin
        
        if is_header:
            c.font = font_hdr
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        elif is_section:
            c.font = font_sec
            c.fill = fill_sec
            if col_idx == 2:
                c.alignment = Alignment(horizontal='left', vertical='center')
            else:
                c.alignment = Alignment(horizontal='center', vertical='center')
        elif is_total:
            c.font = font_bold
            c.fill = fill_win if is_win else fill_total
            c.alignment = Alignment(horizontal='right' if col_idx > 2 else 'left', vertical='center')
            if num_fmt and col_idx > 2:
                c.number_format = num_fmt
        else:
            c.font = font_regular if not (col_idx == 15 and num_fmt) else font_bold
            if col_idx == 2:
                c.alignment = Alignment(horizontal='left', vertical='center')
            else:
                c.alignment = Alignment(horizontal='right', vertical='center')
            if num_fmt and col_idx > 2:
                c.number_format = num_fmt

# =============================================================
# ABA 1: PREMISSAS E DADOS
# =============================================================
ws1 = wb.create_sheet(title='Premissas e Dados')
setup_sheet_header(ws1, "Planejamento Agregado da Produção (PAP) - Parâmetros Gerais", "Base Cadastral, Custos Operacionais e Demanda Mensal")

ws1['B6'] = "PARÂMETRO OPERACIONAL"
ws1['C6'] = "VALOR"
ws1['D6'] = "UNIDADE"
ws1['E6'] = "DESCRIÇÃO / BASE LEGAL"
for col, h in enumerate(['B6', 'C6', 'D6', 'E6'], start=2):
    ws1[h].font = font_hdr
    ws1[h].fill = fill_hdr_navy
    ws1[h].alignment = Alignment(horizontal='center', vertical='center')
ws1.row_dimensions[6].height = 24

params = [
    ("Quadro Base de Operadoras (Costureiras)", 20, "costureiras", "Força de trabalho efetiva inicial"),
    ("Jornada Mensal Líquida por Operadora", 160, "horas/mês", "44h semanais com pausas da NR-17"),
    ("Tempo Padrão no Gargalo da Costura (TP)", 0.80, "horas/conjunto", "Cronoanálise na costura (48 minutos)"),
    ("Produtividade Normal por Costureira", 200, "conjuntos/mês", "160 h ÷ 0.80 h/conjunto"),
    ("Capacidade Normal Base da Fábrica (20 op.)", 4000, "conjuntos/mês", "20 operadoras × 200 un/mês"),
    ("Estoque Inicial de Segurança (Mês 0)", 400, "conjuntos", "Estoque de segurança do ano anterior"),
    ("Estoque Final Desejado (Mês 12)", 400, "conjuntos", "Estoque para inauguração do ano seguinte"),
    ("Limite Legal de Horas Extras (CLT Art. 59)", 0.25, "da jornada (25%)", "Até 2h extras/dia = máx. 800 un/mês"),
    ("Custo de Produção em Turno Normal (cn)", 38.00, "R$ / conjunto", "Tecido, aviamentos, MOD direta e CIF"),
    ("Custo em Turno Extra / Hora Extra (che)", 46.00, "R$ / conjunto", "Adicional constitucional de 50% na MOD"),
    ("Custo de Subcontratação / Facção Externa (csub)", 52.00, "R$ / conjunto", "Oficina externa de costura parceira"),
    ("Custo de Manutenção / Posse de Estoque (ce)", 3.00, "R$ / conjunto / mês", "Armazenagem física e custo de capital (sobre EM)"),
    ("Custo de Ruptura / Atraso na Entrega (ca)", 20.00, "R$ / conjunto / mês", "Multa contratual de SLA corporativo"),
    ("Custo de Contratação / Admissão (Cc)", 1600.00, "R$ / admissão", "Processo seletivo, ASO admissional e treino"),
    ("Custo de Demissão / Rescisão Trabalhista (Cd)", 3000.00, "R$ / demissão", "Aviso prévio, multa 40% FGTS e homologação")
]

for idx, (p, v, u, d) in enumerate(params, start=7):
    ws1[f'B{idx}'] = p
    ws1[f'C{idx}'] = v
    ws1[f'D{idx}'] = u
    ws1[f'E{idx}'] = d
    ws1[f'B{idx}'].font = font_regular
    ws1[f'C{idx}'].font = font_bold
    ws1[f'D{idx}'].font = font_regular
    ws1[f'E{idx}'].font = font_regular
    ws1[f'C{idx}'].alignment = Alignment(horizontal='right')
    if isinstance(v, float):
        if v < 1:
            ws1[f'C{idx}'].number_format = '0.0%' if v == 0.25 else '0.00'
        elif v in [38.0, 46.0, 52.0, 3.0, 20.0, 1600.0, 3000.0]:
            ws1[f'C{idx}'].number_format = 'R$ #,##0.00'
    elif isinstance(v, int):
        ws1[f'C{idx}'].number_format = '#,##0'
    for c in ['B', 'C', 'D', 'E']:
        ws1[f'{c}{idx}'].border = border_thin

# Tabela de Demanda Horizontal
ws1['B24'] = "2. SÉRIE HISTÓRICA E PREVISÃO DA DEMANDA AGREGADA (UNIDADES)"
ws1['B24'].font = font_sec

ws1['B25'] = "Item / Período"
for idx, m in enumerate(meses):
    ws1.cell(row=25, column=3+idx, value=m)
ws1['O25'] = "Total Anual"
for col_idx in range(2, 16):
    cell = ws1.cell(row=25, column=col_idx)
    cell.font = font_hdr
    cell.fill = fill_hdr_navy
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = border_thin

demanda_vals = [4600, 5200, 4800, 3600, 3200, 3000, 3800, 3600, 3400, 3800, 5000, 4000]
ws1['B26'] = "Demanda Prevista (un.)"
for idx, d in enumerate(demanda_vals):
    cell = ws1.cell(row=26, column=3+idx, value=d)
    cell.font = font_bold
    cell.number_format = '#,##0'
    cell.alignment = Alignment(horizontal='right')
    cell.border = border_thin
ws1['O26'] = "=SUM(C26:N26)"
ws1['O26'].font = font_bold
ws1['O26'].fill = fill_total
ws1['O26'].border = border_total
ws1['O26'].number_format = '#,##0'
ws1['B26'].border = border_thin

# =============================================================
# FUNÇÃO PARA CONSTRUIR ABA DE PAP (SEM O TEXTO DRE)
# =============================================================
def build_pap_sheet(title, subtitle, sheet_name, strat_num):
    ws = wb.create_sheet(title=sheet_name)
    setup_sheet_header(ws, title, subtitle)
    
    r_hdr = 6
    ws[f'B{r_hdr}'] = "Variável / Rubrica de Planejamento da Produção"
    for idx, m in enumerate(meses):
        ws.cell(row=r_hdr, column=3+idx, value=m)
    ws[f'O{r_hdr}'] = "Total Anual"
    
    header_fill = fill_hdr_green if strat_num == 3 else fill_hdr_navy
    for col_idx in range(2, 16):
        c = ws.cell(row=r_hdr, column=col_idx)
        c.font = font_hdr
        c.fill = header_fill
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = border_thin
    ws.row_dimensions[r_hdr].height = 26

    # ---------------------------------------------------------
    # ESTRATÉGIA 1: NIVELAMENTO
    # ---------------------------------------------------------
    if strat_num == 1:
        ws['B7'] = "1. DEMANDA E VENDAS (UNIDADES)"
        apply_row_style(ws, 7, is_section=True)
        
        ws['B8'] = "(=) Demanda Prevista (Dt)"
        for idx in range(12):
            col_letter = cols_meses[idx]
            ws[f'{col_letter}8'] = f"='Premissas e Dados'!{col_letter}26"
        ws['O8'] = "=SUM(C8:N8)"
        apply_row_style(ws, 8, num_fmt='#,##0')
        
        ws['B9'] = "(+) Backlog Anterior / Demanda Reprimida"
        ws['C9'] = 0
        for idx in range(1, 12):
            curr_col = cols_meses[idx]
            prev_col = cols_meses[idx-1]
            ws[f'{curr_col}9'] = f"={prev_col}16"
        ws['O9'] = "-"
        apply_row_style(ws, 9, num_fmt='#,##0')
        
        ws['B10'] = "(=) Demanda Efetiva do Período"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}10'] = f"={col}8+{col}9"
        ws['O10'] = "-"
        apply_row_style(ws, 10, num_fmt='#,##0')

        ws['B11'] = "2. PLANO DE PRODUÇÃO E BALANÇO DE ESTOQUES (UNIDADES)"
        apply_row_style(ws, 11, is_section=True)
        
        ws['B12'] = "(+) Estoque Inicial (EIt)"
        ws['C12'] = 400
        for idx in range(1, 12):
            curr_col = cols_meses[idx]
            prev_col = cols_meses[idx-1]
            ws[f'{curr_col}12'] = f"={prev_col}15"
        ws['O12'] = "-"
        apply_row_style(ws, 12, num_fmt='#,##0')
        
        ws['B13'] = "(+) Produção Normal Constante (20 costureiras)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}13'] = 4000
        ws['O13'] = "=SUM(C13:N13)"
        apply_row_style(ws, 13, num_fmt='#,##0')
        
        ws['B14'] = "(=) Disponibilidade Total (EIt + Pt)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}14'] = f"={col}12+{col}13"
        ws['O14'] = "-"
        apply_row_style(ws, 14, num_fmt='#,##0')
        
        ws['B15'] = "(=) Estoque Final (EFt)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}15'] = f"=IF({col}14>={col}10, {col}14-{col}10, 0)"
        ws['O15'] = "=N15"
        apply_row_style(ws, 15, num_fmt='#,##0')
        
        ws['B16'] = "(-) Backlog Final / Atrasos Acumulados"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}16'] = f"=IF({col}14<{col}10, {col}10-{col}14, 0)"
        ws['O16'] = "=SUM(C16:N16)"
        apply_row_style(ws, 16, num_fmt='#,##0')
        
        ws['B17'] = "(=) Estoque Médio Mensal (EMt)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}17'] = f"=({col}12+{col}15)/2"
        ws['O17'] = "=AVERAGE(C17:N17)"
        apply_row_style(ws, 17, num_fmt='#,##0')

        ws['B18'] = "3. PLANO FINANCEIRO / CUSTOS OPERACIONAIS DO PAP (R$)"
        apply_row_style(ws, 18, is_section=True)
        
        ws['B19'] = "(+) Custos de Produção em Turno Normal (cn = R$ 38,00)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}19'] = f"={col}13*'Premissas e Dados'!$C$15"
        ws['O19'] = "=SUM(C19:N19)"
        apply_row_style(ws, 19, num_fmt='R$ #,##0.00')
        
        ws['B20'] = "(+) Custos de Manutenção / Posse de Estoque (ce = R$ 3,00)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}20'] = f"={col}17*'Premissas e Dados'!$C$18"
        ws['O20'] = "=SUM(C20:N20)"
        apply_row_style(ws, 20, num_fmt='R$ #,##0.00')
        
        ws['B21'] = "(+) Custos de Penalidade por Atraso / Backlog (ca = R$ 20,00)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}21'] = f"={col}16*'Premissas e Dados'!$C$19"
        ws['O21'] = "=SUM(C21:N21)"
        apply_row_style(ws, 21, num_fmt='R$ #,##0.00')
        
        ws['B22'] = "(=) CUSTO TOTAL OPERACIONAL DO PAP (R$)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}22'] = f"=SUM({col}19:{col}21)"
        ws['O22'] = "=SUM(C22:N22)"
        apply_row_style(ws, 22, is_total=True, num_fmt='R$ #,##0.00')

    # ---------------------------------------------------------
    # ESTRATÉGIA 2: ACOMPANHAMENTO (CHASE)
    # ---------------------------------------------------------
    elif strat_num == 2:
        ws['B7'] = "1. DEMANDA E VENDAS (UNIDADES)"
        apply_row_style(ws, 7, is_section=True)
        
        ws['B8'] = "(=) Demanda Prevista (Dt)"
        for idx in range(12):
            col_letter = cols_meses[idx]
            ws[f'{col_letter}8'] = f"='Premissas e Dados'!{col_letter}26"
        ws['O8'] = "=SUM(C8:N8)"
        apply_row_style(ws, 8, num_fmt='#,##0')
        
        ws['B9'] = "2. DIMENSIONAMENTO DA FORÇA DE TRABALHO E ROTATIVIDADE"
        apply_row_style(ws, 9, is_section=True)
        
        ws['B10'] = "(=) Operadoras Necessárias no Mês (Nt = Dt ÷ 200)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}10'] = f"=INT({col}8/'Premissas e Dados'!$C$10)"
        ws['O10'] = "=AVERAGE(C10:N10)"
        apply_row_style(ws, 10, num_fmt='#,##0')
        
        ws['B11'] = "(-) Operadoras do Mês Anterior (Nt-1)"
        ws['C11'] = "='Premissas e Dados'!$C$7"
        for idx in range(1, 12):
            curr_col = cols_meses[idx]
            prev_col = cols_meses[idx-1]
            ws[f'{curr_col}11'] = f"={prev_col}10"
        ws['O11'] = "-"
        apply_row_style(ws, 11, num_fmt='#,##0')
        
        ws['B12'] = "(+) Contratações / Admissões Realizadas (Ct)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}12'] = f"=MAX(0, {col}10-{col}11)"
        ws['O12'] = "=SUM(C12:N12)"
        apply_row_style(ws, 12, num_fmt='#,##0')
        
        ws['B13'] = "(-) Demissões / Desligamentos Realizados (Dt)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}13'] = f"=MAX(0, {col}11-{col}10)"
        ws['O13'] = "=SUM(C13:N13)"
        apply_row_style(ws, 13, num_fmt='#,##0')

        ws['B14'] = "3. PLANO DE PRODUÇÃO E ESTOQUES (UNIDADES)"
        apply_row_style(ws, 14, is_section=True)
        
        ws['B15'] = "(=) Produção Real em Turno Normal (Pt = Nt × 200)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}15'] = f"={col}10*'Premissas e Dados'!$C$10"
        ws['O15'] = "=SUM(C15:N15)"
        apply_row_style(ws, 15, num_fmt='#,##0')
        
        ws['B16'] = "(=) Estoque Médio de Segurança (Fixo)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}16'] = 400
        ws['O16'] = 400
        apply_row_style(ws, 16, num_fmt='#,##0')

        ws['B17'] = "4. PLANO FINANCEIRO / CUSTOS OPERACIONAIS DO PAP (R$)"
        apply_row_style(ws, 17, is_section=True)
        
        ws['B18'] = "(+) Custos de Produção em Turno Normal (cn = R$ 38,00)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}18'] = f"={col}15*'Premissas e Dados'!$C$15"
        ws['O18'] = "=SUM(C18:N18)"
        apply_row_style(ws, 18, num_fmt='R$ #,##0.00')
        
        ws['B19'] = "(+) Custos de Admissão de Operadoras (Cc = R$ 1.600,00)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}19'] = f"={col}12*'Premissas e Dados'!$C$20"
        ws['O19'] = "=SUM(C19:N19)"
        apply_row_style(ws, 19, num_fmt='R$ #,##0.00')
        
        ws['B20'] = "(+) Custos de Demissão e Rescisão (Cd = R$ 3.000,00)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}20'] = f"={col}13*'Premissas e Dados'!$C$21"
        ws['O20'] = "=SUM(C20:N20)"
        apply_row_style(ws, 20, num_fmt='R$ #,##0.00')
        
        ws['B21'] = "(+) Custos de Manutenção de Estoque de Segurança (ce = R$ 3,00)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}21'] = f"={col}16*'Premissas e Dados'!$C$18"
        ws['O21'] = "=SUM(C21:N21)"
        apply_row_style(ws, 21, num_fmt='R$ #,##0.00')
        
        ws['B22'] = "(=) CUSTO TOTAL OPERACIONAL DO PAP (R$)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}22'] = f"=SUM({col}18:{col}21)"
        ws['O22'] = "=SUM(C22:N22)"
        apply_row_style(ws, 22, is_total=True, num_fmt='R$ #,##0.00')

    # ---------------------------------------------------------
    # ESTRATÉGIA 3: HÍBRIDA / PATAMARES
    # ---------------------------------------------------------
    elif strat_num == 3:
        ws['B7'] = "1. DEMANDA E VENDAS (UNIDADES)"
        apply_row_style(ws, 7, is_section=True)
        
        ws['B8'] = "(=) Demanda Prevista (Dt)"
        for idx in range(12):
            col_letter = cols_meses[idx]
            ws[f'{col_letter}8'] = f"='Premissas e Dados'!{col_letter}26"
        ws['O8'] = "=SUM(C8:N8)"
        apply_row_style(ws, 8, num_fmt='#,##0')
        
        ws['B9'] = "2. PLANO DE PRODUÇÃO E CAPACIDADE AGREGADA (UNIDADES)"
        apply_row_style(ws, 9, is_section=True)
        
        ws['B10'] = "(+) Produção Normal Estável (19 costureiras fixas)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}10'] = 3800
        ws['O10'] = "=SUM(C10:N10)"
        apply_row_style(ws, 10, num_fmt='#,##0')
        
        ws['B11'] = "(+) Produção em Horas Extras (HEt ≤ 800 un. CLT)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}11'] = f"=MIN(800, MAX(0, {col}8-({col}14+{col}10)))"
        ws['O11'] = "=SUM(C11:N11)"
        apply_row_style(ws, 11, num_fmt='#,##0')
        
        ws['B12'] = "(+) Produção Subcontratada / Facção Externa (SUBt)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}12'] = f"=MAX(0, ({col}8-({col}14+{col}10))-{col}11)"
        ws['O12'] = "=SUM(C12:N12)"
        apply_row_style(ws, 12, num_fmt='#,##0')
        
        ws['B13'] = "(=) Produção Agregada Total (PTt = Normal + HE + SUB)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}13'] = f"=SUM({col}10:{col}12)"
        ws['O13'] = "=SUM(C13:N13)"
        apply_row_style(ws, 13, num_fmt='#,##0')

        ws['B14'] = "(+) Estoque Inicial do Período (EIt)"
        ws['C14'] = 400
        for idx in range(1, 12):
            curr_col = cols_meses[idx]
            prev_col = cols_meses[idx-1]
            ws[f'{curr_col}14'] = f"={prev_col}15"
        ws['O14'] = "-"
        apply_row_style(ws, 14, num_fmt='#,##0')
        
        ws['B15'] = "(=) Estoque Final Regulador (EFt = EIt + PTt - Dt)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}15'] = f"={col}14+{col}13-{col}8"
        ws['O15'] = "=N15"
        apply_row_style(ws, 15, num_fmt='#,##0')
        
        ws['B16'] = "(=) Estoque Médio Mensal (EMt = [EIt + EFt] ÷ 2)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}16'] = f"=({col}14+{col}15)/2"
        ws['O16'] = "=AVERAGE(C16:N16)"
        apply_row_style(ws, 16, num_fmt='#,##0')

        ws['B17'] = "3. PLANO FINANCEIRO / CUSTOS OPERACIONAIS DO PAP (R$)"
        apply_row_style(ws, 17, is_section=True)
        
        ws['B18'] = "(+) Custos de Produção em Turno Normal (cn = R$ 38,00)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}18'] = f"={col}10*'Premissas e Dados'!$C$15"
        ws['O18'] = "=SUM(C18:N18)"
        apply_row_style(ws, 18, num_fmt='R$ #,##0.00')
        
        ws['B19'] = "(+) Custos de Produção em Horas Extras (che = R$ 46,00)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}19'] = f"={col}11*'Premissas e Dados'!$C$16"
        ws['O19'] = "=SUM(C19:N19)"
        apply_row_style(ws, 19, num_fmt='R$ #,##0.00')
        
        ws['B20'] = "(+) Custos de Subcontratação / Facção Externa (csub = R$ 52,00)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}20'] = f"={col}12*'Premissas e Dados'!$C$17"
        ws['O20'] = "=SUM(C20:N20)"
        apply_row_style(ws, 20, num_fmt='R$ #,##0.00')
        
        ws['B21'] = "(+) Custos de Manutenção / Posse de Estoque (ce = R$ 3,00)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}21'] = f"={col}16*'Premissas e Dados'!$C$18"
        ws['O21'] = "=SUM(C21:N21)"
        apply_row_style(ws, 21, num_fmt='R$ #,##0.00')
        
        ws['B22'] = "(=) CUSTO TOTAL OPERACIONAL DO PAP (R$)"
        for idx in range(12):
            col = cols_meses[idx]
            ws[f'{col}22'] = f"=SUM({col}18:{col}21)"
        ws['O22'] = "=SUM(C22:N22)"
        apply_row_style(ws, 22, is_total=True, is_win=True, num_fmt='R$ #,##0.00')

    ws.column_dimensions['B'].width = 44
    for idx in range(12):
        col_letter = cols_meses[idx]
        ws.column_dimensions[col_letter].width = 13
    ws.column_dimensions['O'].width = 16

build_pap_sheet("Estratégia 1: Taxa Constante (Nivelamento / Level)", "Planejamento Agregado da Produção: Meses em Colunas, Variáveis e Custos em Linhas", "Estratégia 1 - Nivelamento", 1)
build_pap_sheet("Estratégia 2: Acompanhamento da Demanda (Chase)", "Planejamento Agregado da Produção: Meses em Colunas, Variáveis e Custos em Linhas", "Estratégia 2 - Acompanhamento", 2)
build_pap_sheet("Estratégia 3: Híbrida / Patamares (Recomendada)", "Planejamento Agregado da Produção: Meses em Colunas, Variáveis e Custos em Linhas", "Estratégia 3 - Híbrida", 3)

# =============================================================
# ABA 5: COMPARATIVO FINANCEIRO CONSOLIDADO
# =============================================================
ws5 = wb.create_sheet(title='Comparativo Financeiro')
setup_sheet_header(ws5, "Comparativo do Plano Financeiro do PAP", "Comparativo Consolidado de Custos Operacionais Mensais e Anuais entre as Estratégias")

r_hdr = 6
ws5[f'B{r_hdr}'] = "Estratégia / Rubrica de Custo Operacional (R$)"
for idx, m in enumerate(meses):
    ws5.cell(row=r_hdr, column=3+idx, value=m)
ws5[f'O{r_hdr}'] = "Total Anual"
for col_idx in range(2, 16):
    c = ws5.cell(row=r_hdr, column=col_idx)
    c.font = font_hdr
    c.fill = fill_hdr_navy
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = border_thin
ws5.row_dimensions[r_hdr].height = 26

# E1
ws5['B7'] = "ESTRATÉGIA 1: PRODUÇÃO CONSTANTE (NIVELAMENTO / LEVEL)"
apply_row_style(ws5, 7, is_section=True)

ws5['B8'] = "(+) Custo de Produção em Turno Normal"
for idx in range(12):
    col = cols_meses[idx]
    ws5[f'{col}8'] = f"='Estratégia 1 - Nivelamento'!{col}19"
ws5['O8'] = "=SUM(C8:N8)"
apply_row_style(ws5, 8, num_fmt='R$ #,##0.00')

ws5['B9'] = "(+) Custo de Posse de Estoque"
for idx in range(12):
    col = cols_meses[idx]
    ws5[f'{col}9'] = f"='Estratégia 1 - Nivelamento'!{col}20"
ws5['O9'] = "=SUM(C9:N9)"
apply_row_style(ws5, 9, num_fmt='R$ #,##0.00')

ws5['B10'] = "(+) Custo de Penalidade por Atraso / Backlog"
for idx in range(12):
    col = cols_meses[idx]
    ws5[f'{col}10'] = f"='Estratégia 1 - Nivelamento'!{col}21"
ws5['O10'] = "=SUM(C10:N10)"
apply_row_style(ws5, 10, num_fmt='R$ #,##0.00')

ws5['B11'] = "(=) CUSTO TOTAL MENSAL - E1 NIVELAMENTO"
for idx in range(12):
    col = cols_meses[idx]
    ws5[f'{col}11'] = f"=SUM({col}8:{col}10)"
ws5['O11'] = "=SUM(C11:N11)"
apply_row_style(ws5, 11, is_total=True, num_fmt='R$ #,##0.00')

# E2
ws5['B13'] = "ESTRATÉGIA 2: ACOMPANHAMENTO DA DEMANDA (CHASE STRATEGY)"
apply_row_style(ws5, 13, is_section=True)

ws5['B14'] = "(+) Custo de Produção em Turno Normal"
for idx in range(12):
    col = cols_meses[idx]
    ws5[f'{col}14'] = f"='Estratégia 2 - Acompanhamento'!{col}18"
ws5['O14'] = "=SUM(C14:N14)"
apply_row_style(ws5, 14, num_fmt='R$ #,##0.00')

ws5['B15'] = "(+) Custo de Rotatividade (Admissões + Demissões)"
for idx in range(12):
    col = cols_meses[idx]
    ws5[f'{col}15'] = f"='Estratégia 2 - Acompanhamento'!{col}19+'Estratégia 2 - Acompanhamento'!{col}20"
ws5['O15'] = "=SUM(C15:N15)"
apply_row_style(ws5, 15, num_fmt='R$ #,##0.00')

ws5['B16'] = "(+) Custo de Posse de Estoque de Segurança"
for idx in range(12):
    col = cols_meses[idx]
    ws5[f'{col}16'] = f"='Estratégia 2 - Acompanhamento'!{col}21"
ws5['O16'] = "=SUM(C16:N16)"
apply_row_style(ws5, 16, num_fmt='R$ #,##0.00')

ws5['B17'] = "(=) CUSTO TOTAL MENSAL - E2 ACOMPANHAMENTO"
for idx in range(12):
    col = cols_meses[idx]
    ws5[f'{col}17'] = f"=SUM({col}14:{col}16)"
ws5['O17'] = "=SUM(C17:N17)"
apply_row_style(ws5, 17, is_total=True, num_fmt='R$ #,##0.00')

# E3
ws5['B19'] = "ESTRATÉGIA 3: HÍBRIDA / PATAMARES (SOLUÇÃO RECOMENDADA)"
apply_row_style(ws5, 19, is_section=True)

ws5['B20'] = "(+) Custo de Produção em Turno Normal (19 operadoras)"
for idx in range(12):
    col = cols_meses[idx]
    ws5[f'{col}20'] = f"='Estratégia 3 - Híbrida'!{col}18"
ws5['O20'] = "=SUM(C20:N20)"
apply_row_style(ws5, 20, num_fmt='R$ #,##0.00')

ws5['B21'] = "(+) Custo de Horas Extras e Facção Externa"
for idx in range(12):
    col = cols_meses[idx]
    ws5[f'{col}21'] = f"='Estratégia 3 - Híbrida'!{col}19+'Estratégia 3 - Híbrida'!{col}20"
ws5['O21'] = "=SUM(C21:N21)"
apply_row_style(ws5, 21, num_fmt='R$ #,##0.00')

ws5['B22'] = "(+) Custo de Posse de Estoque Regulador"
for idx in range(12):
    col = cols_meses[idx]
    ws5[f'{col}22'] = f"='Estratégia 3 - Híbrida'!{col}21"
ws5['O22'] = "=SUM(C22:N22)"
apply_row_style(ws5, 22, num_fmt='R$ #,##0.00')

ws5['B23'] = "(=) CUSTO TOTAL MENSAL - E3 HÍBRIDA (RECOMENDADA)"
for idx in range(12):
    col = cols_meses[idx]
    ws5[f'{col}23'] = f"=SUM({col}20:{col}22)"
ws5['O23'] = "=SUM(C23:N23)"
apply_row_style(ws5, 23, is_total=True, is_win=True, num_fmt='R$ #,##0.00')

# Diferença e Economia
ws5['B25'] = "ECONOMIA GERADA PELA E3 HÍBRIDA FRENTE AO NIVELAMENTO (R$)"
for idx in range(12):
    col = cols_meses[idx]
    ws5[f'{col}25'] = f"={col}11-{col}23"
ws5['O25'] = "=O11-O23"
apply_row_style(ws5, 25, is_total=True, is_win=True, num_fmt='R$ #,##0.00')

ws5.column_dimensions['B'].width = 46
for idx in range(12):
    col_letter = cols_meses[idx]
    ws5.column_dimensions[col_letter].width = 13
ws5.column_dimensions['O'].width = 16

ws1.column_dimensions['B'].width = 42
ws1.column_dimensions['C'].width = 14
ws1.column_dimensions['D'].width = 20
ws1.column_dimensions['E'].width = 45
for idx in range(12):
    col_letter = cols_meses[idx]
    ws1.column_dimensions[col_letter].width = 11
ws1.column_dimensions['O'].width = 15

wb.save("PAP_Confeccao_Uniformes_12_Meses.xlsx")
wb.save("Entregas/PAP_Confeccao_Uniformes_12_Meses.xlsx")
print("Planilha PAP 100% atualizada sem qualquer menção a 'DRE'!")
