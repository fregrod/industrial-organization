# Organização Industrial (UnB / FT / EPR)

Repositório de projetos, relatórios técnicos, scripts de automação e materiais da disciplina de **Organização Industrial** da Universidade de Brasília (**UnB**), Faculdade de Tecnologia (**FT**), Departamento de Engenharia de Produção (**EPR**).

---

## 📋 Projeto em Destaque: Planejamento Agregado de Produção (PAP)

Estudo de caso aplicado ao setor de manufatura (Confecção de Uniformes Profissionais em escala contínua com horizonte de 12 meses), com modelagem matemática, balanço de estoques, análise financeira e viabilidade sob a legislação trabalhista brasileira (CLT).

### 🎯 Estratégias de Produção Modeladas:
1. **Nivelamento (Level Strategy):** Força de trabalho constante (20 operadoras), capacidade fixa de 4.000 un/mês, absorção das flutuações via formação de estoques e ruptura com *backlog*.
2. **Acompanhamento (Chase Strategy):** Dimensionamento dinâmico mês a mês conforme a demanda prevista, com movimentações contratuais (admissões e demissões) e estoque nulo.
3. **Estratégia Híbrida (Recomendada):** Base fabril estável dimensionada na média básica (19 operadoras), acionamento de horas extras dentro dos limites legais do Art. 59 da CLT e subcontratação parceira nos períodos de pico.

---

## 📂 Estrutura do Repositório

```text
├── Entregas/                                  # Arquivos finais consolidados para submissão
│   ├── Relatorio_Tecnico_1_PAP_Manufatura.pdf # Relatório técnico completo (10 páginas)
│   ├── Relatorio_Tecnico_1_PAP_Manufatura.docx# Documento editável em Word
│   ├── PAP_Confeccao_Uniformes_12_Meses.xlsx  # Planilha com 5 abas e modelagem do PAP
│   ├── Apresentacao_Seminario_PAP_Manufatura.pptx # Apresentação executiva (15 slides)
│   ├── Apresentacao_Seminario_PAP_Manufatura.pdf  # Slides em formato PDF
│   ├── Roteiro_Apresentacao_15_Minutos.md     # Roteiro cronometrado de apresentação
│   └── README_Entregas.md                     # Checklist e orientações da entrega
│
├── figuras/                                   # Gráficos em alta resolução (Matplotlib)
│   ├── figura1_fluxo_processo.png
│   ├── figura2_demanda_vs_producao.png
│   ├── figura3_estoques_e_backlog.png
│   └── figura4_comparativo_custos.png
│
├── paginas_pdf/                               # Previews das 10 páginas do relatório técnico
├── slides_preview/                            # Previews dos 15 slides da apresentação
│
├── Scripts de Automação (Python):
│   ├── generate_clean_pap_workbook.py         # Geração da pasta de trabalho horizontal do PAP
│   ├── generate_clean_charts.py               # Geração dos gráficos executivos
│   ├── generate_slides.py                     # Criação dos slides formatados no PPTX
│   ├── build_perfect_10_pages.py              # Compilação e ajuste milimétrico do relatório DOCX
│   └── tune_report_pages.py                   # Ajuste de layout e paginação
│
└── Conteúdos/                                 # Slides e referências conceituais da disciplina
```

---

## 🛠️ Tecnologias e Ferramentas Utilizadas

- **Python 3**:
  - `python-docx`: Automação e diagramação profissional do relatório técnico em conformidade com as normas ABNT.
  - `python-pptx`: Geração automatizada da apresentação de slides executiva (Widescreen 16:9).
  - `matplotlib` & `seaborn`: Visualização de dados e gráficos comparativos de custo e produção.
  - `openpyxl`: Construção programática de planilhas financeiras com fórmulas matemáticas dinâmicas.
- **Microsoft Excel / Office 365**: Modelagem quantitativa, análise de cenários e balanço de massa.

---

## 👥 Autoria e Licença

Desenvolvido para as atividades acadêmicas e de engenharia da **Universidade de Brasília (UnB)**.
Distribuído sob licença aberta para fins educacionais.
