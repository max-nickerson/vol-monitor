# Vol monitor de crédito

Um arquivo só (`vol_monitor.py`). Para o gestor de um fundo de crédito, mostra:
- o que está acontecendo hoje;
- o que vem por aí;
- o que, historicamente, antecipa o preço de cada ativo da carteira.

A página `saida/vol_monitor.html` tem sete abas:

| aba | o que mostra |
|---|---|
| **Hoje** | vol, VaR e backtest; PnL de 12 meses separado em carrego, Δ taxa indicativa e resíduo. Lista "o que olhar agora": negócio da B3 longe da ANBIMA, ação do emissor fora do normal, fato relevante, resultado chegando, amortização, IPCA. Cada item vem com o impacto estimado em bps do PL quando o sinal é estatisticamente útil |
| **O que vem** | agenda de juros, amortização e vencimento (ANBIMA); próximo resultado de cada emissor, estimado pelo histórico de entregas à CVM; IPCA (IBGE); Focus |
| **Mercado e notícias** | Ibovespa, dólar, commodities, Treasury; DI1/DAP ao vivo; Tesouro; índices de crédito IDA/JGP; fatos relevantes e comunicados (CVM); notícias por emissor |
| **Ativos** | tabela da carteira. Ao clicar num ativo:<br>• taxa indicativa (oficial e implícita) com negócios, resultados e fatos relevantes marcados<br>• PnL acumulado contra a taxa<br>• % do PU par<br>• ação × taxa, com correlação por defasagem e intervalo de 95%<br>• fundamentos trimestrais (dívida líquida/EBITDA, cobertura, margem)<br>• agenda e documentos |
| **O que antecipa** | cada sinal com n, correlação, IC 95% e acerto fora da amostra. Também o valor relativo (spread × alavancagem × duration) e uma tabela por emissor |
| **Risco** | retorno diário × VaR, contribuição para o risco, sensibilidade a fatores (IDA, JGP, IMA-B, IRF-M, Ibovespa, dólar, Tesouro) e matriz de correlação |
| **Como é calculado** | as contas |

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
python vol_monitor.py carteira.xlsx             # uma vez
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
