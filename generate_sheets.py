import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
# remove default sheet
wb.remove(wb.active)

# Color palette
navy_dark = '1B365D'
navy_light = 'E8EEF5'
header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
header_fill = PatternFill(start_color=navy_dark, end_color=navy_dark, fill_type='solid')
sub_header_fill = PatternFill(start_color='2C5282', end_color='2C5282', fill_type='solid')
accent_fill = PatternFill(start_color='EDF2F7', end_color='EDF2F7', fill_type='solid')
total_fill = PatternFill(start_color='DCE6F1', end_color='DCE6F1', fill_type='solid')

border_thin = Border(
    left=Side(style='thin', color='CBD5E0'),
    right=Side(style='thin', color='CBD5E0'),
    top=Side(style='thin', color='CBD5E0'),
    bottom=Side(style='thin', color='CBD5E0')
)
border_total = Border(
    top=Side(style='thin', color='1B365D'),
    bottom=Side(style='double', color='1B365D')
)

font_regular = Font(name='Arial', size=10)
font_bold = Font(name='Arial', size=10, bold=True)
font_title = Font(name='Arial', size=14, bold=True, color='1B365D')
font_subtitle = Font(name='Arial', size=11, italic=True, color='4A5568')

# -------------------------------------------------------------
# ABA 1: PREMISSAS E DADOS GERAIS
# -------------------------------------------------------------
ws1 = wb.create_sheet(title='Premissas e Dados')
ws1.views.sheetView[0].showGridLines = True

ws1['B2'] = "VÉRTICE CONFECÇÕES & UNIFORMES PROFISSIONAIS LTDA."
ws1['B2'].font = font_title
ws1['B3'] = "Planejamento Agregado da Produção (PAP) - Horizonte de 12 Meses"
ws1['B3'].font = font_subtitle

# Parâmetros
ws1['B5'] = "1. PARÂMETROS OPERACIONAIS E DA FORÇA DE TRABALHO"
ws1['B5'].font = font_bold
ws1.merge_cells('B5:D5')

params = [
    ("Número Base de Operadores (Costureiras)", 20, "operadores"),
    ("Jornada Mensal Líquida por Operador", 160, "horas/mês"),
    ("Tempo Padrão por Conjunto de Uniforme (TP)", 0.80, "horas/conjunto"),
    ("Produtividade Normal por Operador", 200, "conjuntos/mês"),
    ("Capacidade Produtiva em Turno Normal (20 op.)", 4000, "conjuntos/mês"),
    ("Estoque Inicial de Segurança (Mês 0)", 400, "conjuntos"),
    ("Estoque Final Desejado (Mês 12)", 400, "conjuntos"),
    ("Limite Legal de Horas Extras (CLT - 2h/dia)", "25%", "da jornada normal"),
    ("Capacidade Máxima em Hora Extra (20 op.)", 800, "conjuntos/mês")
]

ws1['B6'] = "Parâmetro"
ws1['C6'] = "Valor"
ws1['D6'] = "Unidade"
for col, h in enumerate(['B6', 'C6', 'D6']):
    ws1[h].font = header_font
    ws1[h].fill = sub_header_fill
    ws1[h].alignment = Alignment(horizontal='center')

for idx, (param, val, unit) in enumerate(params, start=7):
    ws1[f'B{idx}'] = param
    ws1[f'C{idx}'] = val
    ws1[f'D{idx}'] = unit
    ws1[f'B{idx}'].font = font_regular
    ws1[f'C{idx}'].font = font_bold
    ws1[f'D{idx}'].font = font_regular
    ws1[f'C{idx}'].alignment = Alignment(horizontal='right')
    ws1[f'D{idx}'].alignment = Alignment(horizontal='left')
    for c in ['B', 'C', 'D']:
        ws1[f'{c}{idx}'].border = border_thin

# Custos
ws1['B17'] = "2. PARÂMETROS DE CUSTOS (R$)"
ws1['B17'].font = font_bold

custos = [
    ("Custo de Produção em Turno Normal (cn)", 38.00, "R$ / conjunto"),
    ("Custo de Produção em Hora Extra (che)", 46.00, "R$ / conjunto"),
    ("Custo de Subcontratação / Facção Externa (csub)", 52.00, "R$ / conjunto"),
    ("Custo de Manutenção / Posse de Estoque (ce)", 3.00, "R$ / conjunto / mês"),
    ("Custo de Atraso / Ruptura / Backlog (ca)", 20.00, "R$ / conjunto / mês"),
    ("Custo de Contratação de Operadora (Cc)", 1600.00, "R$ / admissão"),
    ("Custo de Demissão / Rescisão de Operadora (Cd)", 3000.00, "R$ / desligamento")
]

ws1['B18'] = "Rubrica de Custo"
ws1['C18'] = "Valor Unitário"
ws1['D18'] = "Unidade de Medida"
for col, h in enumerate(['B18', 'C18', 'D18']):
    ws1[h].font = header_font
    ws1[h].fill = sub_header_fill
    ws1[h].alignment = Alignment(horizontal='center')

for idx, (param, val, unit) in enumerate(custos, start=19):
    ws1[f'B{idx}'] = param
    ws1[f'C{idx}'] = val
    ws1[f'D{idx}'] = unit
    ws1[f'B{idx}'].font = font_regular
    ws1[f'C{idx}'].font = font_bold
    ws1[f'D{idx}'].font = font_regular
    ws1[f'C{idx}'].number_format = 'R$ #,##0.00'
    ws1[f'C{idx}'].alignment = Alignment(horizontal='right')
    ws1[f'D{idx}'].alignment = Alignment(horizontal='left')
    for c in ['B', 'C', 'D']:
        ws1[f'{c}{idx}'].border = border_thin

# Demanda
ws1['F5'] = "3. PREVISÃO DE DEMANDA AGREGADA (12 MESES)"
ws1['F5'].font = font_bold

ws1['F6'] = "Mês"
ws1['G6'] = "Demanda Prevista (un.)"
ws1['H6'] = "Sazonalidade / Justificativa de Mercado"
for h in ['F6', 'G6', 'H6']:
    ws1[h].font = header_font
    ws1[h].fill = sub_header_fill
    ws1[h].alignment = Alignment(horizontal='center')

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
    ws1[f'{c}{19}'].fill = total_fill
    ws1[f'{c}{19}'].border = border_total

wb.save("PAP_Confeccao_Uniformes_12_Meses.xlsx")
print("Aba 1 salva com sucesso.")
