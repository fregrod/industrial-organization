import os
import matplotlib.pyplot as plt
import numpy as np

os.makedirs('figuras', exist_ok=True)

# Set style
plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
plt.rcParams['font.sans-serif'] = 'Arial'
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.size'] = 10

meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
demanda = [4600, 5200, 4800, 3600, 3200, 3000, 3800, 3600, 3400, 3800, 5000, 4000]

# -------------------------------------------------------------
# FIGURA 1: DEMANDA VS PRODUÇÃO NAS 3 ESTRATÉGIAS
# -------------------------------------------------------------
prod_e1 = [4000] * 12
prod_e2 = demanda.copy()
prod_e3 = [4200, 5200, 4800, 3800, 3800, 3800, 3800, 3800, 3800, 3800, 3800, 3800] # PT E3

fig, ax = plt.subplots(figsize=(10, 5), dpi=300)
x = np.arange(len(meses))

ax.plot(x, demanda, marker='o', linewidth=2.5, color='#D9381E', label='Demanda Prevista (Dt)', zorder=5)
ax.plot(x, prod_e1, linestyle='--', linewidth=2, color='#2B6CB0', label='E1: Nivelamento (4.000 un/mês fixo)')
ax.plot(x, prod_e2, linestyle=':', linewidth=2, color='#D69E2E', label='E2: Acompanhamento (Chase)')
ax.plot(x, prod_e3, linestyle='-', linewidth=2.2, color='#2F855A', label='E3: Híbrida (Base 3.800 + HE + Sub)')

ax.set_title('Figura 2: Comparativo de Demanda vs. Taxas de Produção por Estratégia de PAP', fontsize=12, fontweight='bold', pad=15)
ax.set_xlabel('Mês do Horizonte de Planejamento', fontweight='bold', labelpad=10)
ax.set_ylabel('Quantidade de Conjuntos de Uniformes (un.)', fontweight='bold', labelpad=10)
ax.set_xticks(x)
ax.set_xticklabels(meses)
ax.set_ylim(2500, 5600)
ax.legend(frameon=True, facecolor='white', framealpha=0.9, loc='upper right')
ax.grid(True, linestyle='--', alpha=0.5)

plt.tight_layout()
plt.savefig('figuras/figura2_demanda_vs_producao.png')
plt.close()
print("Figura 2 salva.")

# -------------------------------------------------------------
# FIGURA 2: EVOLUÇÃO DE ESTOQUES E BACKLOG (E1 vs E3)
# -------------------------------------------------------------
est_e1 = [0, 0, 0, 0, 0, 0, 200, 600, 1200, 1400, 400, 400]
backlog_e1 = [200, 1400, 2200, 1800, 1000, 0, 0, 0, 0, 0, 0, 0]
est_e3 = [0, 0, 0, 200, 800, 1600, 1600, 1800, 2200, 2200, 1000, 800]

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 5), dpi=300)

# E1
ax1.plot(x, est_e1, marker='s', color='#2B6CB0', linewidth=2, label='Estoque Final (EF)')
ax1.bar(x, backlog_e1, color='#E53E3E', alpha=0.4, label='Backlog / Atraso Crítico')
ax1.set_title('(a) Estratégia 1: Nivelamento (Level)', fontsize=11, fontweight='bold')
ax1.set_xlabel('Mês', fontweight='bold')
ax1.set_ylabel('Unidades', fontweight='bold')
ax1.set_xticks(x)
ax1.set_xticklabels(meses)
ax1.set_ylim(0, 2500)
ax1.legend(frameon=True, facecolor='white', loc='upper left')
ax1.grid(True, linestyle='--', alpha=0.5)

# E3
ax2.plot(x, est_e3, marker='^', color='#2F855A', linewidth=2, label='Estoque Final Regulador (EF)')
ax2.axhline(0, color='black', linewidth=0.8)
ax2.set_title('(b) Estratégia 3: Híbrida / Patamares (Zero Backlog)', fontsize=11, fontweight='bold')
ax2.set_xlabel('Mês', fontweight='bold')
ax2.set_ylabel('Unidades', fontweight='bold')
ax2.set_xticks(x)
ax2.set_xticklabels(meses)
ax2.set_ylim(0, 2500)
ax2.legend(frameon=True, facecolor='white', loc='upper left')
ax2.grid(True, linestyle='--', alpha=0.5)

plt.suptitle('Figura 3: Comparação da Dinâmica de Estoques e Rupturas (E1 vs. E3)', fontsize=12, fontweight='bold', y=1.02)
plt.tight_layout()
plt.savefig('figuras/figura3_estoques_e_backlog.png')
plt.close()
print("Figura 3 salva.")

# -------------------------------------------------------------
# FIGURA 3: COMPARATIVO DE CUSTOS TOTAIS POR RUBRICA
# -------------------------------------------------------------
categorias = ['E1: Nivelamento', 'E2: Acompanhamento', 'E3: Híbrida']
c_prod_normal = [1824000.0, 1824000.0, 1732800.0]
c_he = [0.0, 0.0, 92000.0]
c_sub = [0.0, 0.0, 41600.0]
c_estoque = [36600.0, 14400.0, 36000.0]
c_atraso = [108000.0, 0.0, 0.0]
c_turnover = [0.0, 82800.0, 0.0]

fig, ax = plt.subplots(figsize=(9, 6), dpi=300)
w = 0.55

p1 = ax.bar(categorias, c_prod_normal, width=w, label='Produção em Turno Normal', color='#2B6CB0')
p2 = ax.bar(categorias, c_he, width=w, bottom=c_prod_normal, label='Horas Extras (HE)', color='#319795')
bottom_sub = np.array(c_prod_normal) + np.array(c_he)
p3 = ax.bar(categorias, c_sub, width=w, bottom=bottom_sub, label='Subcontratação (Facção)', color='#4FD1C5')
bottom_est = bottom_sub + np.array(c_sub)
p4 = ax.bar(categorias, c_estoque, width=w, bottom=bottom_est, label='Manutenção de Estoque', color='#D69E2E')
bottom_atr = bottom_est + np.array(c_estoque)
p5 = ax.bar(categorias, c_atraso, width=w, bottom=bottom_atr, label='Custo de Atraso (Backlog)', color='#E53E3E')
bottom_turn = bottom_atr + np.array(c_atraso)
p6 = ax.bar(categorias, c_turnover, width=w, bottom=bottom_turn, label='Rotatividade (Contratação/Demissão)', color='#805AD5')

# Valores totais sobre as barras
totais = [1968600.0, 1921200.0, 1902400.0]
for idx, tot in enumerate(totais):
    ax.text(idx, tot + 20000, f'R$ {tot:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.'),
            ha='center', va='bottom', fontweight='bold', fontsize=10)

ax.set_title('Figura 4: Estrutura Comparativa de Custos Totais Anuais do PAP', fontsize=12, fontweight='bold', pad=15)
ax.set_ylabel('Custo Total Anual (R$)', fontweight='bold', labelpad=10)
ax.set_ylim(0, 2200000)
ax.legend(frameon=True, facecolor='white', framealpha=0.9, loc='upper left', bbox_to_anchor=(1, 1))
ax.grid(True, linestyle='--', alpha=0.5, axis='y')

plt.tight_layout()
plt.savefig('figuras/figura4_comparativo_custos.png')
plt.close()
print("Figura 4 salva.")

# -------------------------------------------------------------
# FIGURA 1: FLUXOGRAMA DE PROCESSOS (MATRIZ VOLUME-VARIEDADE)
# -------------------------------------------------------------
fig, ax = plt.subplots(figsize=(10, 3.5), dpi=300)
ax.axis('off')

# Caixas de processo
boxes = [
    ("Enfesto & Corte", "Tecidos em Brim e Malha\nEnfestadeira e Faca Vertical\nLotes padronizados"),
    ("Bordado & Estamparia", "Personalização Corporativa\nMáquinas Multicabeças\nPonto de Desacoplamento"),
    ("Célula de Costura", "20 Operadoras Polivalentes\nInterlock, Reta e Pespontadeira\nGargalo Operacional"),
    ("Acabamento & Embalagem", "Caseamento, Botões, Revisão,\nPassadoria e Embalagem Final\nPronto para Expedição")
]

colors = ['#E2E8F0', '#E2E8F0', '#FEFCBF', '#E2E8F0'] # Amarelo no gargalo
edge_colors = ['#4A5568', '#4A5568', '#D69E2E', '#4A5568']

for i, (titulo, desc) in enumerate(boxes):
    x_pos = 0.05 + i * 0.24
    # Retângulo
    rect = plt.Rectangle((x_pos, 0.25), 0.19, 0.55, facecolor=colors[i], edgecolor=edge_colors[i],
                         linewidth=2 if i == 2 else 1.5, boxstyle="round,pad=0.03" if hasattr(plt, 'boxstyle') else None)
    ax.add_patch(rect)
    ax.text(x_pos + 0.095, 0.65, titulo, ha='center', va='center', fontweight='bold', fontsize=10, color='#1A202C')
    ax.text(x_pos + 0.095, 0.42, desc, ha='center', va='center', fontsize=7.5, color='#4A5568')
    
    # Seta entre as caixas
    if i < 3:
        ax.annotate('', xy=(x_pos + 0.235, 0.52), xytext=(x_pos + 0.195, 0.52),
                    arrowprops=dict(arrowstyle="->", color='#2B6CB0', lw=2.5))

ax.text(0.5, 0.92, "Figura 1: Mapeamento do Fluxo Operacional na Vértice Confecções (Gargalo: Células de Costura)",
        ha='center', va='center', fontweight='bold', fontsize=11, color='#1B365D')

ax.set_xlim(0, 1)
ax.set_ylim(0, 1)
plt.tight_layout()
plt.savefig('figuras/figura1_fluxo_processo.png')
plt.close()
print("Figura 1 salva.")
