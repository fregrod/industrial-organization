# Roteiro Oficial de Apresentação – Seminário de PAP (15 Minutos)

**Disciplina:** Organização Industrial (UnB / FT / EPR – 2026.2)  
**Professores:** Dr. Marcelo Carneiro Gonçalves, Dra. Camilla Pimenta Barbosa, Dr. João Victor Queiroz  
**Trabalho:** Relatório Técnico 1 & Seminário de Planejamento Agregado da Produção (PAP)  
**Caso Prático:** *Vértice Confecções & Uniformes Profissionais Ltda. - EPP*  
**Duração Total:** Exatamente 15 minutos (com alarme programado aos 14m30s)  
**Divisão Sugerida:** 5 integrantes (~3 minutos cada) ou adaptável para qualquer formação de equipe.

---

## ⏱️ Cronômetro e Distribuição dos Papéis

| Bloco | Minutagem | Integrante | Slides | Foco Temático |
| :--- | :---: | :---: | :---: | :--- |
| **Bloco 1: Abertura e Diagnóstico** | 00:00 – 03:00 | Integrante 1 | 1, 2, 3 | Capa, Empresa Real, Gargalo e Cronoanálise |
| **Bloco 2: Metodologia e Demanda** | 03:00 – 06:00 | Integrante 2 | 4, 5, 6 | Horizonte 12 Meses, Estratégias Clássicas e Matriz de Custos |
| **Bloco 3: Modelagem das Estratégias 1 e 2** | 06:00 – 09:00 | Integrante 3 | 7, 8, 9 | Nivelamento (Backlog), Acompanhamento (Turnover) e Falhas |
| **Bloco 4: Estratégia Híbrida e Resultados** | 09:00 – 12:30 | Integrante 4 | 10, 11, 12 | Formulação Híbrida, Estoque e Comparativo Econômico |
| **Bloco 5: Trade-offs, Implantação e Fechamento** | 12:30 – 15:00 | Integrante 5 | 13, 14, 15 | Trade-offs de Slack, Plano de Ação Diretor, Conclusão e Banca |

---

## 🎙️ Roteiro de Fala Passo a Passo (Slide a Slide)

### BLOCO 1: ABERTURA, CONTEXTUALIZAÇÃO E CAPACIDADE (00:00 – 03:00)

#### Slide 1: Capa e Identificação (00:00 – 01:00)
* **Tempo:** 1 minuto
* **Falante:** Integrante 1
* **Fala Sugerida:**
  > "Bom dia professores Marcelo Gonçalves, Camilla Pimenta e João Victor Queiroz, e colegas. Nós somos o grupo responsável pela apresentação do Relatório Técnico 1 de Organização Industrial, com foco no Planejamento Agregado da Produção em Manufatura.
  > Nosso trabalho desenvolveu a modelagem matemática completa de 12 meses para uma empresa real de pequeno porte do Distrito Federal: a **Vértice Confecções e Uniformes Profissionais Ltda.** Nosso objetivo foi determinar a estratégia ótima de produção capaz de atender a uma demanda sazonal com menor custo total, zero atrasos e estrita conformidade trabalhista. Vamos começar entendendo o diagnóstico operacional da empresa."

#### Slide 2: Contextualização da Empresa e Problema Operacional (01:00 – 02:00)
* **Tempo:** 1 minuto
* **Falante:** Integrante 1
* **Fala Sugerida:**
  > "A Vértice opera em Taguatinga com 20 costureiras e produz uniformes profissionais em brim pesado. O produto agregado representativo é o 'Conjunto Padrão' (calça e camisa polo).
  > Historicamente, a empresa sofria com um clássico dilema de manufatura: a demanda tem dois picos acentuados no ano — a volta às aulas e renovação de frotas no 1º trimestre, e o aquecimento industrial no final do ano. Sem um PAP formal, a fábrica alternava entre atrasos severos com multas de clientes e contratações desesperadas com baixa produtividade. Decidimos estruturar o PCP tático da fábrica a partir de uma rigorosa cronoanálise."

#### Slide 3: Mapeamento de Processo e Cronoanálise do Gargalo (02:00 – 03:00)
* **Tempo:** 1 minuto
* **Falante:** Integrante 1
* **Fala Sugerida:**
  > "Como ensinado pelo professor Marcelo Carneiro na Aula 01, toda capacidade produtiva é ditada pelo gargalo da linha. No fluxo da Vértice — Corte, Costura, Estamparia, Acabamento e Embalagem —, a cronoanálise identificou que o gargalo está nas células de Costura e Montagem.
  > O Tempo Padrão (TP) apurado foi de **0,80 horas (48 minutos)** por conjunto, já considerando 15% de tolerâncias para fadiga e pausas ergonômicas da NR-17. Numa jornada mensal líquida de 160 horas por costureira, a produtividade apurada é de exatamente **200 peças por operadora/mês**. Com o quadro de 20 operadoras, a capacidade nominal da fábrica é de **4.000 conjuntos por mês**."

---

### BLOCO 2: DEMANDA, ESTRATÉGIAS E PARÂMETROS DE CUSTO (03:00 – 06:00)

#### Slide 4: Horizonte de Planejamento e Perfil da Demanda (03:00 – 04:00)
* **Tempo:** 1 minuto
* **Falante:** Integrante 2
* **Fala Sugerida:**
  > "Passando ao planejamento da demanda, conforme as diretrizes do professor João Victor na Aula 03, adotamos um horizonte de 12 meses, de Janeiro a Dezembro. A demanda anual consolidada é de **48.000 unidades**, com média de 4.000 por mês.
  > Notem no gráfico a forte oscilação: Janeiro atinge 4.600 peças, Fevereiro chega a 5.200 e Março a 4.800 unidades — bem acima da capacidade fabril. No meio do ano, a demanda desaba para 3.000 peças em Junho, voltando a subir para 5.000 em Novembro. O estoque inicial de segurança herdado foi de 400 peças. O desafio do PAP é balancear essa montanha-russa de demanda."

#### Slide 5: Fundamentação Teórica: As 3 Estratégias do PAP (04:00 – 05:00)
* **Tempo:** 1 minuto
* **Falante:** Integrante 2
* **Fala Sugerida:**
  > "A literatura de Nigel Slack e Corrêa & Corrêa estabelece três caminhos puros e mistos para o PAP:
  > 1. **Nivelamento (Level Strategy):** mantém a capacidade produtiva e a mão de obra perfeitamente constantes, usando estoques e backlog para amortecer as variações.
  > 2. **Acompanhamento (Chase Strategy):** ajusta a capacidade mês a mês para perseguir a demanda exata, contratando e demitindo operadoras conforme o volume.
  > 3. **Estratégia Híbrida ou por Patamares:** combina capacidade própria estável com recursos flexíveis — horas extras e facções parceiras — nos picos sazonais."

#### Slide 6: Matriz de Custos Unitários e Restrições Legais (05:00 – 06:00)
* **Tempo:** 1 minuto
* **Falante:** Integrante 2
* **Fala Sugerida:**
  > "Para alimentar o modelo quantitativo, levantamos os custos reais praticados no polo de confecções do DF:
  > - Custo de produção normal: **R$ 38,00/peça**.
  > - Hora extra: **R$ 46,00/peça** (adicional de 50% da CLT).
  > - Subcontratação (facção externa): **R$ 52,00/peça**.
  > - Custo de posse de estoque: **R$ 3,00/peça/mês**, calculado sobre o estoque médio para não distorcer o capital de giro.
  > - Multa por atraso (backlog): **R$ 20,00/peça/mês**, refletindo penalidades contratuais severas.
  > - Admissão e treinamento: **R$ 1.600,00** | Demissão e encargos rescisórios: **R$ 3.000,00**.
  > A restrição jurídica fundamental foi o Artigo 59 da CLT: o limite máximo de 2 horas extras diárias impõe um teto de 800 peças por mês em hora extra."

---

### BLOCO 3: MODELAGEM DAS ESTRATÉGIAS PURAS 1 E 2 (06:00 – 09:00)

#### Slide 7: Estratégia 1: Nivelamento (Taxa Constante) (06:00 – 07:00)
* **Tempo:** 1 minuto
* **Falante:** Integrante 3
* **Fala Sugerida:**
  > "Na Estratégia 1, fixamos a produção em exatamente 4.000 peças todos os meses com as 20 costureiras.
  > O resultado operacional foi desastroso no 1º semestre: como os pedidos de Jan a Mar superam a capacidade em 2.200 unidades, o estoque inicial de 400 peças foi consumido imediatamente em Janeiro. A fábrica entrou em colapso de entregas: gerou backlog de 200 peças em Jan, 1.400 em Fev e acumulou um pico de 2.200 peças atrasadas em Março!
  > A fábrica só conseguiu zerar os pedidos atrasados em Junho, mantendo clientes corporativos esperando por até 5 meses seguidos."

#### Slide 8: Estratégia 2: Acompanhamento da Demanda (Chase) (07:00 – 08:00)
* **Tempo:** 1 minuto
* **Falante:** Integrante 3
* **Fala Sugerida:**
  > "Partimos então para a Estratégia 2: variar o quadro de funcionárias para produzir exatamente a demanda do mês.
  > Em Janeiro, a empresa contratou 3 operadoras (passando para 23). Em Fevereiro, precisou de 26 costureiras (mais 3 contratações). Porém, em Março começa a queda: 2 demissões; em Abril, mais 6 demissões; em Maio, 2 demissões; e em Junho, mais 1 demissão, reduzindo a fábrica para apenas 15 costureiras.
  > No segundo semestre, o ciclo recomeça: contratações em Julho, demissões em Setembro e nova corrida de admissões para Novembro."

#### Slide 9: Análise Crítica dos Extremos (Por que falham?) (08:00 – 09:00)
* **Tempo:** 1 minuto
* **Falante:** Integrante 3
* **Fala Sugerida:**
  > "A análise crítica evidencia as armadilhas dos modelos teóricos puros:
  > No Nivelamento, a penalidade por atraso custou nada menos que **R$ 108.000,00**, derrubando o On-Time Delivery para 79,2%. A perda de clientes seria fatal.
  > No Acompanhamento, o custo direto foi menor, mas gerou **36 movimentações de pessoal** (18 admissões e 18 demissões), custando R$ 82.800,00. Além do custo financeiro, a curva de aprendizado derruba a qualidade, e o clima organizacional é destruído pelo medo constante de demissão. Os extremos são inviáveis."

---

### BLOCO 4: ESTRATÉGIA HÍBRIDA E RESULTADOS ECONÔMICOS (09:00 – 12:30)

#### Slide 10: Estratégia 3: Solução Híbrida por Patamares (Recomendada) (09:00 – 10:30)
* **Tempo:** 1 minuto e meio
* **Falante:** Integrante 4
* **Fala Sugerida:**
  > "Desenvolvemos então a **Estratégia 3: Híbrida por Patamares**, desenhada sob medida para a realidade de uma MPME.
  > Qual foi a nossa lógica de engenharia? Fixamos o quadro base em **19 costureiras permanentes**, gerando uma capacidade interna estável de 3.800 peças/mês.
  > Como atendemos os picos?
  > - No 1º trimestre: acionamos horas extras dentro do teto legal (400 un em Jan, 800 em Fev e 600 em Mar).
  > - O excedente crítico de Fevereiro e Março foi transferido para facções de costura parceiras homologadas (600 un em Fev e 400 em Mar).
  > - No 2º semestre: a demanda cai, mas a produção se mantém em 3.800 peças. Em vez de demitir, formamos estoque pulmão para absorver com tranquilidade o pico de 5.000 peças de Novembro, sem gastar nem um real com hora extra no final do ano!"

#### Slide 11: Dinâmica de Estoques e Nível de Serviço (10:30 – 11:30)
* **Tempo:** 1 minuto
* **Falante:** Integrante 4
* **Fala Sugerida:**
  > "Na dinâmica de estoques, a diferença é cristalina:
  > A Estratégia Híbrida eliminou 100% do backlog em todos os 12 meses. O estoque mínimo é mantido de forma segura nos primeiros meses, acumula gradativamente até 1.400 unidades em Outubro e desce suavemente para entregar um estoque final de 800 conjuntos em Dezembro — atendendo com folga a exigência da diretoria de iniciar o ano seguinte com estoque de segurança. O nível de serviço ao cliente é de **100% de pontualidade**."

#### Slide 12: Matriz Comparativa Consolidada de Custos e KPIs (11:30 – 12:30)
* **Tempo:** 1 minuto
* **Falante:** Integrante 4
* **Fala Sugerida:**
  > "Vejamos a síntese econômico-financeira:
  > - **Estratégia 1 (Nivelamento):** Custo de **R$ 1.968.600,00** (onerada pelas multas de R$ 108 mil).
  > - **Estratégia 2 (Acompanhamento):** Custo de **R$ 1.921.200,00** (onerada pelas rescisões de R$ 82,8 mil).
  > - **Estratégia 3 (Híbrida):** Custo total de **R$ 1.902.400,00**!
  > A Estratégia Híbrida é a **campeã absoluta**: economiza **R$ 66.200,00** em relação ao Nivelamento e **R$ 18.800,00** em relação ao Acompanhamento, com turnover ZERO e risco trabalhista nulo."

---

### BLOCO 5: TRADE-OFFS, PLANO DE AÇÃO E FECHAMENTO (12:30 – 15:00)

#### Slide 13: Discussão de Trade-offs e Objetivos de Desempenho (12:30 – 13:30)
* **Tempo:** 1 minuto
* **Falante:** Integrante 5
* **Fala Sugerida:**
  > "Analisando os resultados pelos 5 objetivos de desempenho operacional de Nigel Slack:
  > 1. **Custo:** A Híbrida minimiza o custo total e elimina desperdícios de rescisão e multas.
  > 2. **Confiabilidade:** 100% de pedidos entregues no prazo acordado.
  > 3. **Qualidade:** Manter uma equipe estável e experiente garante índice de retrabalho inferior a 1,5%.
  > 4. **Flexibilidade:** A combinação de HE e facção fornece elasticidade sem comprometer a estrutura física.
  > 5. **Rapidez:** O desacoplamento por estoque garante pronta-entrega imediata no 2º semestre."

#### Slide 14: Recomendações Práticas e Plano de Ação (13:30 – 14:15)
* **Tempo:** 45 segundos
* **Falante:** Integrante 5
* **Fala Sugerida:**
  > "Para a diretoria da Vértice, estruturamos um plano de implantação em 4 passos:
  > 1. Fixar a equipe fabril em 19 costureiras registradas.
  > 2. Celebrar acordo de prorrogação de jornada com o sindicato para até 2h extras/dia exclusivamente no primeiro trimestre.
  > 3. Homologar previamente 2 oficinas de facção externas, fornecendo os tecidos já cortados pela Vértice para mitigar riscos de vínculo da Súmula 331 do TST.
  > 4. Instituir o comitê mensal de S&OP entre Vendas e Produção para revisar os planos com horizonte rolante."

#### Slide 15: Conclusão, Referências e Abertura para a Banca (14:15 – 15:00)
* **Tempo:** 45 segundos
* **Falante:** Integrante 5
* **Fala Sugerida:**
  > "Concluímos destacando que o PAP provou ser a ferramenta vital para alinhar o setor Comercial e a Manufatura. A modelagem quantitativa evitou que a empresa tomasse decisões intuitivas perigosas.
  > Nossas referências fundamentaram-se em Slack, Corrêa & Corrêa, na legislação trabalhista e nas aulas dos professores Marcelo Gonçalves, Camilla Pimenta e João Victor Queiroz.
  > Encerramos rigorosamente aos 14 minutos e 45 segundos e colocamo-nos à inteira disposição da banca para responder às perguntas. Muito obrigado!"

---

## 🎯 Dicas de Ouro para a Apresentação Oral da Equipe

1. **Passagem de Bastão Suave:** Quando um integrante terminar seu bloco, faça a ponte para o próximo (ex: *"Passo agora a palavra ao [Nome], que apresentará a modelagem matemática das estratégias."*).
2. **Postura com os Dados:** Ao apontar tabelas e gráficos, cite os números-chave: *"R$ 1.902.400,00"*, *"19 operadoras estáveis"*, *"teto legal de 800 peças pela CLT"*. Isso demonstra domínio completo do projeto.
3. **Controle de Tempo Visual:** Coloque um celular em modo silencioso na mesa do apresentador rodando um cronômetro regressivo de 15 minutos. Se aos 12 minutos o Integrante 4 ainda estiver falando, o Integrante 5 deve sintetizar o fechamento para não estourar os 15 minutos.
