"""Taxa indicativa histórica reconstruída do PU indicativo oficial (ANBIMA): parte da taxa oficial mais recente e volta no tempo.
Δtaxa(t) = -[ret. PU indicativo - ret. PU par - ln((1+taxa)/(1+taxa contratada))·dias/252] / duration modificada(t);
os eventos (juros, amortização) entram nos dois retornos; duration(t) = duration oficial × prazo restante(t) / prazo restante hoje."""
import numpy as np
import pandas as pd


def serie_oficial(h):
    """Série oficial limpa: só dias com PU indicativo; tira pico isolado do % do PU par (> 8 p.p. da mediana de 11 dias, ex.: erro
    nos 1os dias de marcação); E = pagamentos entre um preço e o seguinte (incorporação de juros não é pagamento; se o PU par cai
    sem evento registrado, o pagamento é a queda)."""
    x = h.sort_values("data").reset_index(drop=True)
    pago = x.valor_evento.fillna(0).where(~x.evento.fillna("").str.upper().str.contains("INCORPORA"), 0)
    pago = pago.where(~x.evento.fillna("").str.upper().str.contains("INCORPORA") | x.evento.fillna("").str.contains(";"), 0)
    x["pago"] = pago
    v = x[x.pu_ind.notna() & x.pu_par.notna() & (x.pu_par > 0)].copy()
    casas = v.pu_ind.map(lambda u: next((k for k in range(7) if abs(round(u, k) - u) < 1e-9), 7))
    v = v[10.0 ** -casas.values / v.pu_ind.values < 5e-4]                  # PU publicado arredondado (ex.: 1,0 num papel de VNE 1): passo > 0,05%
    pct = v.pu_ind / v.pu_par * 100
    med = pct.rolling(11, center=True, min_periods=3).median()
    v = v[(pct - med).abs() <= 8]
    if len(v) < 2:
        return v.assign(E=0.0)
    grupo = np.searchsorted(v.index.values, x.index.values, side="left")          # cada dia do calendário -> próximo preço válido
    E = pd.Series(x.pago.values).groupby(grupo).sum()
    v["E"] = [float(E.get(k, 0.0)) for k in range(len(v))]
    queda = v.pu_par.shift() - v.pu_par
    infer = queda.where(queda > 0.002 * v.pu_par.shift(), 0).fillna(0)
    v["E"] = np.maximum(v.E.values, infer.values)
    v.iloc[0, v.columns.get_loc("E")] = 0.0
    return v.reset_index(drop=True)


def reconstroi(h, taxa0, data0, dur0, venc, cupom):
    """h: DataFrame[data, pu_par, pu_ind, valor_evento] (ANBIMA); taxa0/dur0 oficiais em data0. Devolve Series de taxa (% a.a.) por data."""
    x = serie_oficial(h)
    x = x[x.data <= data0].reset_index(drop=True)
    if len(x) < 2 or taxa0 is None or not np.isfinite(taxa0):
        return pd.Series(dtype=float)
    venc, T0 = pd.Timestamp(venc), max((pd.Timestamp(venc) - pd.Timestamp(data0)).days / 365.25, 0.05)
    E = x.E.values
    ri = np.log((x.pu_ind.values[1:] + E[1:]) / x.pu_ind.values[:-1])
    rp = np.log((x.pu_par.values[1:] + E[1:]) / x.pu_par.values[:-1])
    du = np.array([max(1, np.busday_count(a.date(), b.date())) for a, b in zip(x.data[:-1], x.data[1:])])
    y = np.full(len(x), np.nan)
    y[-1] = taxa0 / 100
    c = cupom / 100
    for k in range(len(x) - 1, 0, -1):
        T = max((venc - x.data[k]).days / 365.25, 0.02)
        D = dur0 * T / T0 / (1 + y[k])
        ex = ri[k - 1] - rp[k - 1] - np.log((1 + y[k]) / (1 + c)) * du[k - 1] / 252
        y[k - 1] = y[k] + ex / max(D, 0.02)
    return pd.Series(y * 100, index=x.data.values)
