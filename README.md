# Vol monitor de crédito

Um arquivo só (`vol_monitor.py`). Para o gestor de um fundo de crédito, mostra:
- o que está acontecendo hoje;
- o que vem por aí;
- o que, historicamente, antecipa o preço de cada ativo da carteira.

Monitor dos fundos internos, um arquivo só (`vol_monitor.py`). O seletor no topo troca de fundo e as abas mostram o que importa para o gestor, sem repetir informação.

| aba | o que mostra |
|---|---|
| **Visão geral** | patrimônio e captação; retorno 12 meses em % do CDI; carrego da carteira de crédito (% do CDI, DI+ e IPCA+ médios); CS01 (R$ por bp de spread) e DV01 de juro real; VaR em R$; marcação esperada hoje. Gráficos: cota × CDI, retorno esperado para 1/5/10/20 dias com faixa de 80% contra o carrego, alocação (categoria, indexador, emissor, setor, prazo), PnL de 12 meses (carrego × taxa × resíduo), captação líquida diária e liquidez (dias para vender a 20% do volume da B3) |
| **Sinais** | como a ANBIMA deve fechar hoje (negócios do dia, DAP, ação); situações de 1 a 20 dias com ~70–80% de acerto histórico; mais fracos e mais fortes do modelo de 5 dias; fatos relevantes e avisos aos debenturistas |
| **Negócios** | todos os negócios de crédito do dia na B3, ranqueados pela diferença entre a taxa negociada e a taxa indicativa ANBIMA da véspera (filtros de volume, papéis do fundo, índice e busca), com o movimento esperado da ANBIMA no dia e o acerto histórico. Ao clicar num papel: gráfico da taxa ANBIMA com os negócios de 30 dias em bolhas (tamanho = volume) e os de hoje em destaque, com a hora |
| **Ativos** | tabela da carteira de crédito (taxa, % par, duration, Δ taxa, vol, % do risco, dias para vender, sinais). Ao clicar num ativo:<br>• projeção do % do PU par para 20 dias úteis (mais provável e faixa 10–90%), com a leitura em texto<br>• taxa indicativa com os negócios<br>• PnL contra a taxa<br>• ação × taxa<br>• fundamentos, agenda e documentos |
| **Risco** | vol, VaR e backtest; cenários (spread +100 bps, juro real, juro pré, default do maior emissor) em % e R$; quem concentra o risco; retorno diário × VaR; fatores; correlação |
| **Agenda** | amortizações e vencimentos em R$, juros, resultados estimados dos emissores, IPCA; Focus |
| **Mercado** | índices de crédito e renda fixa; DI1/DAP ao vivo; referências (Ibovespa, dólar, commodities, Treasury); notícias dos emissores do fundo |

**Carteiras:** por padrão vem a última carteira aberta de cada fundo na CVM, que é divulgada com atraso. Para usar a posição atual, salve `carteiras/<nome do fundo>.xlsx` com as colunas `ativo` (`CODIGO - AAAA-MM-DD`) e `perc_nav`. Cota, PL, captação e resgate vêm do informe diário da CVM. A lista de fundos internos sai do painel de fundos (`PAINEL_HTML`).

## Previsão da marcação ANBIMA de hoje (aba Hoje)
A previsão é recalculada a cada 15 minutos com os negócios do dia. Para cada ativo, estima quanto o % do PU par da ANBIMA de hoje vai andar, somando:
- **Negócios do próprio papel na B3:** a ANBIMA anda uma fração do gap entre o negócio e a ANBIMA de ontem, conforme o volume do dia:

  | Volume do papel no dia | Fração do gap | Acerto fora da amostra |
  |---|---|---|
  | > R$ 20 mi | 39% | 82% |
  | R$ 5–20 mi | 30% | 73% |
  | R$ 1–5 mi | 7% | 58% |
  | varejo | — | ruído |

- **Papéis IPCA:** o DAP ao vivo × duration. Sem DAP, o tom do mercado IPCA.
- **Papel sem negócio:**
  - outros papéis do mesmo emissor (6,5% do gap deles);
  - na falta deles, o tom do mercado (10% do gap médio).
- **Ação do emissor:** a partir de 2% no dia, 1,8 bps de % par por 1%.

A calibração vem de um estudo com 447 debêntures, CRI e CRA e com os negócios B3 de 04/2023 a 10/2026. O teste é fora da amostra, no último terço do tempo. O resultado completo está na aba "O que antecipa". Não mostraram poder:
- o momento da ANBIMA;
- o gap de ontem não fechado;
- o resultado trimestral;
- o fato relevante;
- a ação antecipando a semana seguinte.

## Sinais de 1 a 20 dias úteis (aba Hoje)
São regras testadas nas 100 debêntures mais líquidas (04/2023–10/2026). Todas ficaram estáveis em 2023-24 e em 2025-26. Cada uma mostra a mediana do movimento do % do PU par, a faixa 10–90% do que aconteceu e o acerto do sentido contra a taxa-base.

| Situação | 5 dias | 10 dias | 20 dias |
|---|---|---|---|
| DI+: semana muito pior que o mercado (≤ −2 desvios) → continua caindo contra o mercado | −7 bps, 69% (base 53%) | −12, 69% | −18, 69% |
| DI+: semana ≥ 50 bps melhor que o mercado → continua subindo contra o mercado | +14, 68% (47%) | +27, 70% | +43, 70% |
| DI+: negócio ≥ R$ 20 mi ≥ 30 bps acima da ANBIMA → % par total sobe | +35, 80% (51%) | +50, 81% | +54, 86% |
| Semana muito melhor com negócios acima da ANBIMA → % par total sobe | +12, 72% (54%) | +20, 71% | +27, 70% |
| IPCA: queda extrema na semana (≤ −4 desvios) → repique (só 2025-26) | +28, 73% (50%) | +20, 69% | — |

Motivo provável: a ANBIMA ajusta o spread de papéis DI+ em etapas, então a marcação tem inércia depois de um choque, e acompanha negócios grandes ao longo de vários dias.

Não tiveram poder: os fundos (resgates × posições do CDA, compras e vendas, fluxo agregado dos fundos de crédito), as notícias e o momentum do índice.

## Próximos 5 dias úteis: faixa de 80% e direção (aba Hoje)
O modelo é treinado uma vez por dia com todo o histórico ANBIMA em cache (~440 papéis). Dele saem duas peças:
- **Centro:** quanto o papel deve andar contra o mercado (LightGBM).
- **Faixa:** a vol prevista × um fator calibrado nos 6 meses anteriores, por faixa de vol (conformal). O alvo de cobertura é 80%.

Cada papel vem com os 3 fatores que mais pesaram na previsão.

Teste fora da amostra (walk-forward mensal, 04/2024–10/2026, 161 mil previsões):

| Medida | Resultado |
|---|---|
| Cobertura da faixa | 77% |
| Largura da faixa | 111 bps, contra 117 da banda pela vol do próprio papel |
| Correlação de postos por semana (IC) | 0,078 (t = 5,7) |
| 10% mais fracos | −5 bps contra o mercado em 5 dias; subiram em 46% das vezes |
| 10% mais fortes | +1 bps; subiram em 55% das vezes |

O que tem poder:
- reversão de curto prazo (o papel que caiu na semana contra o mercado volta);
- papel abaixo do par sobe pelo carrego;
- negócios do mês acima da marcação;
- ação do emissor em 5 a 20 dias.

Vol, volume, dias sem negócio e o lado dos negócios não ajudam.

## Rodar
```
pip install pandas numpy requests pyarrow openpyxl lxml selenium lightgbm
python vol_monitor.py modelo                    # cria carteira_modelo.xlsx (15 debêntures mais negociadas)
python vol_monitor.py                           # todos os fundos internos
python vol_monitor.py carteira.xlsx             # uma carteira avulsa
python vol_monitor.py carteira.xlsx rodar       # a cada 15 min no pregão
```
A carteira tem as colunas `ativo` (`CODIGO - AAAA-MM-DD`, ex.: `HAPV18 - 2030-01-27`) e `perc_nav` (% do PL).
O Microsoft Edge é usado em modo headless para abrir as páginas públicas da ANBIMA Data.

A primeira execução é lenta: baixa 3,5 anos de negócios da B3, a página da ANBIMA de cada ativo e 5 anos de ITR/DFP da CVM.
Da segunda em diante só atualiza o que mudou. Tudo fica em `dados/`, que não vai para o git.

## As contas
- **Retorno do dia** = (PU indicativo + pagamentos) / PU indicativo anterior − 1. Amortização e juros não são perda.
- **PnL contra a taxa indicativa:**
  - carrego = rendimento do PU par + (taxa − cupom);
  - Δ taxa = −duration modificada × variação da taxa indicativa;
  - resíduo = o que sobra.
- **Taxa antes do arquivo diário da ANBIMA:** é a taxa implícita no PU oficial (erro de ~1,5 bps).
- **Vol:** retornos diários oficiais em 12 meses. Na falta deles, negócios semanais sem o ruído; na falta de negócios, a mediana dos pares.
- **VaR 95%:** paramétrico, com correlação encolhida 30% e backtest de 250 dias.
- **Fundamentos (ITR/DFP consolidados):**
  - EBITDA = EBIT + D&A da DFC;
  - dívida = empréstimos e financiamentos;
  - caixa = caixa + aplicações;
  - data de entrega = 1ª entrega à CVM (point in time).
- **Previsibilidade:** correlação com intervalo de 95% e teste fora da amostra (ajuste nos 2/3 mais antigos, teste no último terço). Só os sinais que passam nos dois viram impacto estimado.

## Fontes (públicas)
- **ANBIMA:** ANBIMA Data (páginas públicas) e arquivo diário de debêntures.
- **B3:** boletim diário, com negócios de balcão e cadastro, e futuros.
- **SND:** debentures.com.br.
- **CVM:** ITR, DFP, IPE, FCA e informe das securitizadoras.
- **Mercado e macro:**
  - Yahoo Finance;
  - SmartWealth (índices; só leitura local, os dados não são redistribuídos);
  - BCB;
  - Tesouro Direto;
  - IBGE;
  - Google Notícias.
