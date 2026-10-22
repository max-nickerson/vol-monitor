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

## Rodar
```
pip install pandas numpy requests pyarrow openpyxl lxml selenium
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
