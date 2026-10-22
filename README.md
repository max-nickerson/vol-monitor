# Vol monitor de crédito

Vol, VaR e curvas (% do PU par, PU indicativo, taxa indicativa) de cada ativo de uma carteira de crédito, com a marcação oficial da ANBIMA.

## Entrada
Excel com `codigo` no formato `CODIGO - yyyy-mm-dd` (ex.: `AMAGI - 2028-01-28`), `perc_nav` e `data`.
`python vol_monitor.py exemplo` gera um `carteira_exemplo.xlsx` a partir da carteira pública (CVM) de um fundo.

## Rodar
```
python vol_simples.py          # uma vez -> saida/vol_simples.html
python vol_simples.py rodar    # repete a cada 15 min
```

## Arquivos
| arquivo | o que faz |
|---|---|
| `vol_simples.py` | monitor principal: retornos diários oficiais, vol, correlação (shrink 30%), VaR e backtest, negócios do dia (B3) |
| `anbima_hist.py` | histórico diário de PU par e PU indicativo desde a emissão, da página pública do ativo no ANBIMA Data (só o que a própria página recebe) |
| `taxa_hist.py` | limpa a série oficial (picos, incorporação de juros, pagamentos) e reconstrói a taxa antes da 1ª taxa oficial pela duration |
| `vol_monitor.py` | POC: leitura da carteira, identificação dos ativos, notícias/eventos/mercados |
| `pagina_simples.html`, `pagina.html` | templates das páginas |
| `fontes_mercado.html`, `critica.html` | pesquisa de fontes de mercado e crítica do POC |

## Fontes
ANBIMA Data (PU/taxa oficiais), ANBIMA arquivos diários, SND (preços de negociação), B3 BDI (negócios), CVM CDA, BCB.

## Dependências
Python 3.11+, `pandas`, `numpy`, `requests`, `pyarrow`, `openpyxl`, Microsoft Edge (headless, para `anbima_hist.py`).
Usa dois módulos de fora da pasta (não estão no repositório): `carteira.py` (negócios e cadastro B3) e
`lookthrough_cvm.py` (SND, ANBIMA, CVM; o mesmo do repositório `painel-fundos`).

`dados/` e `saida/` são caches locais e não vão para o git.
