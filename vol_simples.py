"""
Vol monitor simples (fundo Dual) — só preço de MERCADO.

Medida: variação diária do % do PU par (preço de mercado / PU da curva). Assim juros acumulados e pagamentos não viram "vol".
  Debêntures: SND, PU médio negociado do dia e % do PU da curva (3 anos), só dias com >= R$ 50 mil negociados.
  CRI/CRA:    B3 negócios de balcão (taxa média ponderada por volume, >= R$ 50 mil) -> preço pela duration.
  Sem negócio suficiente: ANBIMA indicativa (% PU par diário, histórico local) e, por último, CVM mensal (marcado como sem mercado).
  Ruído de negócio (bid/ask, corretagem do varejo): correção de Roll papel a papel.
  Vol do fundo: vols x correlação dos retornos semanais (encolhida 50% para a média); VaR paramétrico 95%.
  Ao vivo: negócios da B3 de hoje (~15 min) viram "% do par hoje" com o PU par do SND do dia.
Uso:  python vol_simples.py        (uma vez)      python vol_simples.py rodar   (a cada 15 min no pregão)
"""
import io
import json
import sys
import time
from datetime import datetime

import numpy as np
import pandas as pd
import requests

import anbima_hist as AH
import taxa_hist as TH
import vol_monitor as vm

ct, lt = vm.ct, vm.lt
SND_DIR = vm.DADOS / "snd"
SND_DIR.mkdir(exist_ok=True)
FUNDO, MIN_DIA, MIN_SEM, K_PRIOR = "Dual", 50_000, 20, 20
HOJE = pd.Timestamp.today().normalize()


# ============================================================== dados de mercado
def snd_hist(cod):
    """Preços de negociação do SND (debentures.com.br): por dia, quantidade, nº de negócios, PU mín/médio/máx e % do PU da curva."""
    f = SND_DIR / f"{cod}.parquet"
    velho = pd.read_parquet(f) if f.exists() else None
    if velho is not None and pd.Timestamp(f.stat().st_mtime, unit="s") >= HOJE:
        return velho
    ini = velho.data.max() - pd.Timedelta(days=10) if velho is not None and len(velho) else HOJE - pd.DateOffset(years=3)
    u = ("https://www.debentures.com.br/exploreosnd/consultaadados/mercadosecundario/precosdenegociacao_r.asp?op_exc=False&emissor=&isin="
         f"&ativo={cod}&dt_ini={ini:%d/%m/%Y}&dt_fim={HOJE:%d/%m/%Y}&Submit.x=1")
    try:
        t = requests.get(u, headers=vm.UA, timeout=120).content.decode("latin1")
        tb = [x for x in pd.read_html(io.StringIO(t), decimal=",", thousands=".") if x.shape[1] >= 19]
    except Exception:
        return velho if velho is not None else pd.DataFrame(columns=["data", "qtd", "neg", "pmin", "pmed", "pmax", "pct"])
    novo = pd.DataFrame(columns=["data", "qtd", "neg", "pmin", "pmed", "pmax", "pct"])
    if tb:
        x = tb[0][tb[0][0].astype(str).str.match(r"^\d{2}/\d{2}/\d{4}$")]
        n = lambda c: pd.to_numeric(x[c], errors="coerce")
        novo = pd.DataFrame({"data": pd.to_datetime(x[0], dayfirst=True), "qtd": n(8), "neg": n(10), "pmin": n(12), "pmed": n(14),
                             "pmax": n(16), "pct": n(18)})
    d = pd.concat([velho, novo]) if velho is not None else novo
    d = d.drop_duplicates("data", keep="last").sort_values("data").reset_index(drop=True)
    d.to_parquet(f)
    return d


def b3_dias(dias=380, cods=None):
    """Negócios B3 dos papéis `cods`: dias passados lidos do cache diário (só as linhas dos papéis), hoje baixado de novo (~15 min)."""
    ds = list(pd.bdate_range(HOJE - pd.Timedelta(days=dias), HOJE))
    ct.negocios(ds)                                          # garante o cache (só baixa o que falta e o dia de hoje)
    fil = [("cod", "in", list(cods))] if cods else None
    partes = []
    for d in ds:
        f = ct.CACHE / "b3" / f"neg_{d:%Y%m%d}.parquet"
        if f.exists():
            try:
                x = pd.read_parquet(f, filters=fil)
            except Exception:                                 # dia sem negócio: coluna vazia sem tipo
                x = pd.read_parquet(f)
                x = x[x.cod.isin(cods)] if cods else x
            if len(x):
                partes.append(x)
    n = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame(columns=["cod", "data", "taxa", "pu", "vol"])
    n["data"] = pd.to_datetime(n.data).dt.normalize()
    return n


# ============================================================== estatística
def retornos(datas, nivel):
    """Δ ln(nível) entre pregões com preço, dividido pela raiz do nº de dias úteis entre eles (vira retorno de 1 dia)."""
    s = pd.Series(np.log(np.asarray(nivel, float)), index=pd.to_datetime(pd.Series(datas)).values).dropna()
    s = s[~s.index.duplicated(keep="last")].sort_index()
    if len(s) < 3:
        return pd.Series(dtype=float)
    gap = np.array([max(1, np.busday_count(a.date(), b.date())) for a, b in zip(s.index[:-1], s.index[1:])])
    return pd.Series(np.diff(s.values) / np.sqrt(gap), index=s.index[1:])


def vol(z):
    """Vol anual (%) robusta: tira saltos > 6 desvios (MAD) e corrige o ruído de negócio (Roll). Devolve (vol 12m, vol EWMA, n)."""
    z = z.dropna()
    if len(z) < 5:
        return None, None, len(z)
    mad = np.median(np.abs(z - z.median())) * 1.4826 + 1e-12
    z = z[np.abs(z - z.median()) < 6 * mad]
    v = z.var(ddof=1)
    fator = 1.0
    if len(z) >= 10:
        c1 = np.cov(z.values[1:], z.values[:-1])[0, 1]
        vr = max(v + 2 * c1, 0.25 * v)
        fator, v = vr / v if v > 0 else 1, vr
    e = z.iloc[:20].var(ddof=1) if len(z) >= 20 else v
    for x in z.values:                                     # EWMA (RiskMetrics, λ = 0,94), mesma correção de ruído
        e = 0.94 * e + 0.06 * x * x
    return float(np.sqrt(252 * v) * 100), float(np.sqrt(252 * e * fator) * 100), len(z)


def ruido(datas, nivel, ah):
    """Variância do ruído de cada preço negociado (log): dispersão de ln(negócio) - ln(ANBIMA) nos dias com os dois (sem a média)."""
    if ah is None or len(ah) < 6:
        return None, 0
    m = pd.DataFrame({"data": pd.to_datetime(pd.Series(datas)).values, "n": np.log(np.asarray(nivel, float))}).merge(
        pd.DataFrame({"data": ah.data.values, "a": np.log(ah.pct_par.values.astype(float))}), on="data").dropna()
    if len(m) < 8:
        return None, len(m)
    d = m.n - m.a
    return float(d.var(ddof=1)), len(m)


def vol_semanal(datas, nivel, ruido_var=None):
    """Vol anual (%) com retornos semanais (último preço de cada semana): o ruído de negócio pesa ~5x menos que no diário.
    Devolve (vol 12m, vol atual EWMA λ 0,85 ~ meia-vida de 4 semanas, nº de semanas)."""
    s = pd.Series(np.log(np.asarray(nivel, float)), index=pd.to_datetime(pd.Series(datas)).values).dropna().sort_index()
    w = s.groupby(pd.Grouper(freq="W-FRI")).last().dropna()
    if len(w) < 6:
        return None, None, max(len(w) - 1, 0)
    gap = np.diff(w.index.values).astype("timedelta64[D]").astype(int) / 7
    z = pd.Series(np.diff(w.values) / np.sqrt(gap), index=w.index[1:])
    z = z[gap <= 4]                                        # mais de 4 semanas sem negócio: não usa o salto
    if len(z) < 5:
        return None, None, len(z)
    mad = np.median(np.abs(z - z.median())) * 1.4826 + 1e-12
    z = z[np.abs(z - z.median()) < 6 * mad]
    v0 = v = z.var(ddof=1)
    if ruido_var is not None:                              # ruído medido contra a ANBIMA: cada retorno tem o ruído das duas pontas
        v = max(v - 2 * ruido_var, 0.25 * v)
    elif len(z) >= 10:                                     # sem ANBIMA: Roll (autocorrelação negativa do ruído)
        c1 = np.cov(z.values[1:], z.values[:-1])[0, 1]
        v = max(v + 2 * c1, 0.25 * v)
    e = z.iloc[:8].var(ddof=1) if len(z) >= 8 else v0
    for x in z.values:
        e = 0.85 * e + 0.15 * x * x
    e *= v / v0 if v0 > 0 else 1
    return float(np.sqrt(52 * v) * 100), float(np.sqrt(52 * e) * 100), len(z)


def historico(cod, b3c, ah, ho=None, tcalc=None):
    """Histórico diário do ativo, em colunas: d data | u PU médio negociado | p % do PU par negociado (SND) | t taxa média negociada (B3)
    | v volume (R$ mi) | uA PU indicativo | pA % PU par | tA taxa indicativa (ANBIMA)."""
    h = {}
    f = SND_DIR / f"{cod}.parquet"
    if f.exists():
        for r in pd.read_parquet(f).itertuples():
            h.setdefault(r.data, {}).update(u=r.pmed, p=r.pct, v=r.qtd * r.pmed / 1e6)
    bt = b3c[(b3c.cod == cod) & (b3c.vol > 0)]
    if len(bt):
        g = bt.assign(tv=bt.taxa.where(bt.taxa > 0) * bt.vol, tw=bt.vol.where(bt.taxa > 0), pv=bt.pu * bt.vol).groupby("data").agg(
            tv=("tv", "sum"), tw=("tw", "sum"), pv=("pv", "sum"), vol=("vol", "sum"))
        for d, r in g.iterrows():
            e = h.setdefault(d, {})
            if r.tw > 0:
                e["t"] = r.tv / r.tw
            e.setdefault("u", r.pv / r.vol)
            e.setdefault("v", r.vol / 1e6)
    if ho is not None:                                       # ANBIMA oficial desde o início da marcação
        for r in TH.serie_oficial(ho).itertuples():
            h.setdefault(r.data, {}).update(uA=r.pu_ind, pA=r.pu_ind / r.pu_par * 100 if r.pu_par else None)
    if tcalc is not None:
        for d, v in tcalc.items():
            h.setdefault(pd.Timestamp(d), {}).update(tA=v, tc=1)
    for r in ah.itertuples():
        h.setdefault(r.data, {}).update(uA=r.pu, pA=r.pct_par, tA=r.taxa_ind, tc=0)
    ds = sorted(h)
    rd = {"u": 2, "p": 2, "t": 3, "v": 2, "uA": 2, "pA": 2, "tA": 4, "tc": 0}
    out = {"d": [d.strftime("%Y-%m-%d") for d in ds]}
    for k, c in rd.items():
        out[k] = [None if (v := h[d].get(k)) is None or v != v else round(float(v), c) for d in ds]
    return out


# ============================================================== cálculo
def calcula():
    global HOJE
    HOJE = pd.Timestamp.today().normalize()
    t0 = time.time()
    a = vm.identifica(vm.le_carteira())
    D = vm.painel()
    fundo = next(f for f in D["fundos"] if f["nome"] == FUNDO)
    M, CDI = D["meses"], D["cdi"]
    perf = {p["chave"]: p for p in D["perf"]}
    _, an = lt.anbima_debentures()
    an = an.assign(cod=an.codigo.str.strip())
    b3 = b3_dias((HOJE - pd.Timestamp("2023-04-03")).days, set(a.codigo.dropna()))   # B3: desde o 1o dia do arquivo público
    b3c = b3
    try:                                                    # PU par de hoje (SND) para converter os negócios de hoje em % do par
        pp = lt.snd_pu([HOJE], log=lambda m: None)
        pu_par_hoje = pp[pp.data == pp.data.max()].drop_duplicates("codigo").set_index("codigo").pu_par
    except Exception:
        pu_par_hoje = pd.Series(dtype=float)
    lim12 = HOJE - pd.Timedelta(days=365)
    cods_all = [c for c in a.codigo.dropna()]
    oficial = AH.coleta(cods_all, log=lambda m: None)
    resumo = AH.resumo([c for c in cods_all if c in oficial and not (an.cod == c).any()])
    ativos, niveis, diarios = [], {}, {}
    for x in a.to_dict("records"):
        cod = x["codigo"]
        x.update(coberto=bool(cod), fonte=None, vol=None, vol_ewma=None, n=0)
        if not cod:
            ativos.append(x)
            continue
        p, det = perf.get(cod, {}), D["det"].get(cod) or {}
        ah = an[an.cod == cod].sort_values("data")
        anos = max((pd.Timestamp(x["venc"]) - HOJE).days / 365.25, 0.05)
        dur = (float(ah.duration_anos.dropna().iloc[-1]) if ah.duration_anos.notna().any() else
               p.get("duration") or min(anos, anos * 0.8 + 0.2))
        x["duration"] = round(float(dur), 2)
        serie = {"d": [], "p": [], "v": []}
        # ---- mercado: SND (debêntures) ou B3 (CRI/CRA e o que não está no SND)
        vt = et = None
        n = 0
        s = snd_hist(cod) if (x["tipo"] in ("DEB", None) and not cod.startswith(("CRA", "CRI")) and not cod[:2].isdigit()) else pd.DataFrame()
        if len(s) and s.pct.notna().any():
            s = s[(s.qtd * s.pmed >= MIN_DIA) & s.pct.between(20, 150)]
            s12 = s[s.data >= lim12]
            x["vol_diaria"] = vol(retornos(s12.data, s12.pct))[0]
            rv, x["n_ruido"] = ruido(s.data, s.pct, ah)
            x["ruido_pct"] = float(np.sqrt(rv) * 100) if rv is not None else None
            vt, et, n = vol_semanal(s12.data, s12.pct, rv)
            x["vol_bruta"] = vol_semanal(s12.data, s12.pct, 0.0)[0]
            x["fonte"] = "SND (negócios)"
            serie = {"d": [d.strftime("%Y-%m-%d") for d in s.data], "p": [round(v, 3) for v in s.pct], "v": [round(q * pm / 1e6, 2) for q, pm in zip(s.qtd, s.pmed)]}
            niveis[cod] = pd.Series(s.pct.values, index=s.data.values)
            if len(s):
                x["ult_pct"], x["ult_data"] = float(s.pct.iloc[-1]), s.data.iloc[-1].strftime("%d/%m/%y")
                m1 = s[s.data <= s.data.iloc[-1] - pd.Timedelta(days=30)]
                x["var_1m"] = float((s.pct.iloc[-1] / m1.pct.iloc[-1] - 1) * 100) if len(m1) else None
            x["dias_neg_12m"] = int(len(s12))
            x["vol_fin_3m"] = float((s.qtd * s.pmed)[s.data >= HOJE - pd.Timedelta(days=91)].sum() / 63 / 1e6)
        else:
            bh = b3c[(b3c.cod == cod) & (b3c.vol >= MIN_DIA) & (b3c.taxa > 0)]       # B3 informa 0 quando não há taxa
            if len(bh):
                g = bh.assign(tv=bh.taxa * bh.vol).groupby("data").agg(tv=("tv", "sum"), vol=("vol", "sum")).reset_index()
                g["taxa"] = g.tv / g.vol
                g = g.dropna(subset=["taxa"])
                proxy = np.exp(-dur * g.taxa / 100) * 100          # nível de preço implícito na taxa (duration)
                g12 = g.data >= lim12
                x["vol_diaria"] = vol(retornos(g.data[g12], proxy[g12]))[0]
                rv, x["n_ruido"] = ruido(g.data, proxy, ah)
                x["ruido_pct"] = float(np.sqrt(rv) * 100) if rv is not None else None
                vt, et, n = vol_semanal(g.data[g12], proxy[g12], rv)
                x["vol_bruta"] = vol_semanal(g.data[g12], proxy[g12], 0.0)[0]
                x["fonte"] = "B3 (negócios)"
                serie = {"d": [d.strftime("%Y-%m-%d") for d in g.data], "t": [round(v, 3) for v in g.taxa], "v": [round(v / 1e6, 2) for v in g.vol]}
                niveis[cod] = pd.Series(proxy.values, index=g.data.values)
                x["ult_taxa"], x["ult_data"] = float(g.taxa.iloc[-1]), g.data.iloc[-1].strftime("%d/%m/%y")
                x["dias_neg_12m"] = int(len(g))
                x["vol_fin_3m"] = float(g.vol[g.data >= HOJE - pd.Timedelta(days=91)].sum() / 63 / 1e6)
        # ---- ANBIMA indicativa (% PU par)
        va = ea = None
        na = 0
        if len(ah) >= 6:
            va, ea, na = vol(retornos(ah.data, ah.pct_par))
            x["anbima"] = {"d": [d.strftime("%Y-%m-%d") for d in ah.data], "p": [round(v, 3) if v == v else None for v in ah.pct_par]}
            x["taxa_anbima"], x["pct_anbima"], x["pu_anbima"] = float(ah.taxa_ind.iloc[-1]), float(ah.pct_par.iloc[-1]), float(ah.pu.iloc[-1])
            x["data_anbima"] = ah.data.iloc[-1].strftime("%d/%m/%y")
        # ---- ANBIMA oficial (página do ativo): retorno total diário pelo PU indicativo + eventos; taxa antes de 14/09 calculada
        vo = eo = None
        no, tcalc, ho = 0, None, oficial.get(cod)
        if ho is not None and len(TH.serie_oficial(ho)) >= 20:
            hi = TH.serie_oficial(ho)
            gap = np.array([max(1, np.busday_count(p_.date(), q_.date())) for p_, q_ in zip(hi.data[:-1], hi.data[1:])])
            rr = (hi.pu_ind.values[1:] + hi.E.values[1:]) / hi.pu_ind.values[:-1] - 1
            z = pd.Series(rr / np.sqrt(gap), index=hi.data.values[1:])
            diarios[cod] = z
            z12 = z[z.index > lim12]
            if len(z12) >= 20:
                vo, no = float(z12.std(ddof=1) * np.sqrt(252) * 100), len(z12)
                e = z12.iloc[:20].var(ddof=1)
                for v_ in z12.values:
                    e = 0.94 * e + 0.06 * v_ * v_
                eo = float(np.sqrt(252 * e) * 100)
                m1 = z12[z12.index > z12.index.max() - pd.Timedelta(days=30)]
                x["ret_1m"] = float(((1 + m1).prod() - 1) * 100)
            u = hi.iloc[-1]
            x["pu_anbima"], x["pct_anbima"], x["data_anbima"] = float(u.pu_ind), float(u.pu_ind / u.pu_par * 100), u.data.strftime("%d/%m/%y")
            anc = None
            if len(ah) and ah.taxa_ind.notna().any():
                f0 = ah.dropna(subset=["taxa_ind"]).iloc[0]
                anc = (float(f0.taxa_ind), f0.data, float(f0.duration_anos) if f0.duration_anos == f0.duration_anos else x["duration"])
            elif cod in resumo:
                rs = resumo[cod]
                anc = (rs["taxa"], pd.Timestamp(rs["data"]), rs["duration"] or x["duration"])
                x["taxa_anbima"], x["duration"] = rs["taxa"], round(rs["duration"] or x["duration"], 2)
            if anc and anc[0] is not None:
                cupom = x.get("taxa_emissao") if x.get("taxa_emissao") == x.get("taxa_emissao") and x.get("taxa_emissao") is not None else anc[0]
                tc = TH.reconstroi(ho, anc[0], anc[1], anc[2], x["venc"], float(cupom))
                tcalc = tc[pd.DatetimeIndex(tc.index) < anc[1]] if len(tc) else None
                x["taxa_inicio"] = float(tc.iloc[0]) if len(tc) else None
        x["vol_oficial"], x["n_oficial"] = vo, no
        # ---- CVM mensal (sem mercado)
        r = det.get("r") or []
        ex = [r[i] - CDI[M[i]] for i in range(len(r)) if r[i] is not None and M[i] in CDI][-36:]
        vm_ = float(np.std(ex, ddof=1) * np.sqrt(12)) if len(ex) >= 6 else None
        # ---- combinação
        if vo is not None:
            est, ew, x["fonte"] = vo, eo, "ANBIMA (oficial)"
        elif vt is not None and n >= MIN_SEM:
            est, ew = vt, et
        elif (vt is not None and n >= 8) or (va is not None and na >= 6):
            parts = [(v, e, k) for v, e, k, mn in ((vt, et, n, 8), (va, ea, na, 6)) if v is not None and k >= mn]
            nk = sum(k for *_, k in parts)
            est, ew = sum(v * k for v, _, k in parts) / nk, sum(e * k for _, e, k in parts) / nk
            if vm_ is not None:
                est = (nk * est + K_PRIOR * vm_) / (nk + K_PRIOR)
            x["fonte"] = (x["fonte"] + " + " if x["fonte"] else "") + ("ANBIMA" if va is not None else "")
        elif vm_ is not None:
            est, ew, x["fonte"] = vm_, None, "CVM mensal (sem mercado)"
        else:
            est = ew = None
            x["fonte"] = None
        x.update(vol=est, vol_ewma=ew, n=n, n_anbima=na, vol_mercado=vt, vol_anbima=va, vol_cvm=vm_)
        x["hist"] = historico(cod, b3c, ah, ho, tcalc)
        # ---- ao vivo: negócios de hoje -> variação pela TAXA (mesma base em qualquer papel), referência = ANBIMA de ontem
        hj = b3[(b3.cod == cod) & (b3.data == HOJE) & (b3.taxa > 0)]
        if len(hj):
            gr = hj[hj.vol >= MIN_DIA]
            so_varejo = not len(gr)
            gr = hj if so_varejo else gr
            tx = float((gr.taxa * gr.vol).sum() / gr.vol.sum())
            pu = float((gr.pu * gr.vol).sum() / gr.vol.sum())
            ref = ref_pu = quando = dref = None
            if len(ah) and ah.data.iloc[-1] < HOJE and ah.taxa_ind.notna().any():
                u = ah.dropna(subset=["taxa_ind"]).iloc[-1]
                ref, ref_pu, dref, quando = float(u.taxa_ind), float(u.pu), u.data, "ANBIMA " + u.data.strftime("%d/%m")
            else:
                ant = b3c[(b3c.cod == cod) & (b3c.data < HOJE) & (b3c.vol >= MIN_DIA) & (b3c.taxa > 0)]
                if len(ant):
                    u = ant[ant.data == ant.data.max()]
                    ref, ref_pu = float((u.taxa * u.vol).sum() / u.vol.sum()), float((u.pu * u.vol).sum() / u.vol.sum())
                    dref, quando = u.data.iloc[0], "negócio " + u.data.iloc[0].strftime("%d/%m")
            var = None
            if ref_pu:                                         # efeito pelo PU (vale em qualquer nível de taxa), sem o rendimento do período
                du = max(1, np.busday_count(pd.Timestamp(dref).date(), HOJE.date()))
                carrego = 1.00055 ** du - 1                   # ~CDI + spread médio por dia útil
                var = (pu / ref_pu - 1 - carrego) * 100
                if -50 > var or var < -15 and abs(tx - ref) < 1:  # PU caiu sem a taxa mudar: pagamento (cupom/amortização) entre as datas
                    var = -x["duration"] * (tx - ref)
            ruido_ = (x.get("ruido_pct") or 0) / (len(gr) ** 0.5)            # ruído do preço médio de hoje
            vol_d = (est or 0) / np.sqrt(252)
            x["hoje"] = {"n": int(len(hj)), "vol": float(hj.vol.sum()), "taxa": tx, "ref": ref, "quando": quando, "varejo": so_varejo,
                         "bps": (tx - ref) * 100 if ref is not None else None, "var": var,
                         "sd": var / np.sqrt(vol_d ** 2 + ruido_ ** 2) if var is not None and (vol_d or ruido_) else None,
                         "verificar": var is not None and abs(var) > 15}
        ativos.append(x)
    A = pd.DataFrame(ativos)
    # sem histórico suficiente (papel novo, sem negócio): vol mediana dos pares do mesmo tipo e faixa de duration
    A["faixa"] = pd.cut(A.duration, [0, 1.5, 3, 5, 99], labels=False)
    ok = A.coberto & A.vol.notna() & A.fonte.str.contains("oficial", na=False)
    for i in A.index[A.coberto & A.vol.isna()]:
        par = A[ok & (A.tipo == A.tipo[i]) & (A.faixa == A.faixa[i])]
        par = par if len(par) >= 3 else A[ok & (A.faixa == A.faixa[i])]
        if len(par):
            A.loc[i, ["vol", "fonte"]] = [float(par.vol.median()), f"estimado (mediana de {len(par)} pares)"]
    cart, serie_fundo = risco(A, niveis, fundo, diarios)
    print(f"cálculo: {len(A)} linhas, {A.coberto.sum()} com código, {time.time() - t0:.0f}s")
    return A, cart, serie_fundo


def risco(A, niveis, fundo, diarios):
    """Σ = vols × correlação; correlação dos retornos diários oficiais (ANBIMA, 12m, ≥ 60 dias em comum), encolhida 30% para a média.
    Backtest: retorno diário da carteira (ativos com ANBIMA) nos últimos 250 dias × VaR 95% do modelo para esse mesmo conjunto."""
    lim = HOJE - pd.Timedelta(days=365)
    k = A[A.vol.notna()].copy()
    w, s = k.perc_nav.values / 100, k.vol.values / 100
    cods = k.codigo.tolist()
    Rd = pd.DataFrame({c: z[z.index > lim] for c, z in diarios.items()})
    C = Rd.corr(min_periods=60)
    vals = C.where(~np.eye(len(C), dtype=bool)).stack()
    rho = float(np.clip(vals.mean(), 0.02, 0.9)) if len(vals) else 0.3
    n = len(k)
    R = np.full((n, n), rho)
    for i in range(n):
        for j in range(i + 1, n):
            c = C.at[cods[i], cods[j]] if cods[i] in C.index and cods[j] in C.columns else np.nan
            if c == c:
                R[i, j] = R[j, i] = 0.7 * c + 0.3 * rho
    np.fill_diagonal(R, 1)
    S = np.outer(s, s) * R
    vp = float(np.sqrt(w @ S @ w))
    contrib = w * (S @ w) / vp if vp else 0 * w
    A.loc[k.index, "contrib_pct"] = contrib / vp * 100 if vp else 0
    pl = fundo.get("pl") or 0
    # backtest: só os ativos com série oficial
    of = np.array([c in diarios for c in cods])
    wo = w[of]
    so = float(np.sqrt(wo @ S[np.ix_(of, of)] @ wo)) / np.sqrt(252) if of.any() else np.nan
    Rt = pd.DataFrame({c: diarios[c] for c in np.array(cods)[of]}).sort_index()
    Rt = Rt[Rt.index > HOJE - pd.Timedelta(days=400)]
    wv = pd.Series(wo, index=Rt.columns)
    ret = (Rt.fillna(0) * wv).sum(axis=1)
    cobre = Rt.notna().mul(wv * 100).sum(axis=1)
    ret, cobre = ret[cobre > 0.8 * cobre.max()], cobre[cobre > 0.8 * cobre.max()]
    ult = ret.iloc[-250:]
    viol = float((ult < -1.645 * so).mean() * 100) if len(ult) else None
    real = float(ult.std() * np.sqrt(252) * 100) if len(ult) else None
    vol63 = ret.rolling(63, min_periods=40).std() * np.sqrt(252) * 100
    serie = {"d": [d.strftime("%Y-%m-%d") for d in vol63.index], "vol": [None if v != v else round(v, 4) for v in vol63],
             "ret": [round(v * 100, 4) for v in ret], "var": round(-1.645 * so * 100, 4), "cobre": [round(c_, 1) for c_ in cobre]}
    grupo = A.fonte.fillna("sem dado").map(lambda f: "oficial" if "oficial" in f else "mercado" if ("SND" in f or "B3" in f) else
                                           "CVM" if "CVM" in f else "estimado" if "estimado" in f else "sem dado")
    fonte = A[A.coberto].groupby(grupo[A.coberto]).perc_nav.sum().round(2).to_dict()
    return {"vol": vp * 100, "vol_sem_div": float((w * s).sum() * 100), "var95_1d": 1.645 * vp / np.sqrt(252) * 100,
            "var95_1m": 1.645 * vp * np.sqrt(21 / 252) * 100, "pl": pl, "cobertura": float(k.perc_nav.sum()), "rho": rho,
            "pl_codigo": float(A[A.coberto].perc_nav.sum()), "por_fonte": fonte, "n": n,
            "duration": float((k.perc_nav * k.duration).sum() / max(k.perc_nav.sum(), 1e-9)),
            "choque": float(-(k.perc_nav / 100 * k.duration).sum()),
            "backtest": {"dias": int(len(ult)), "violacoes": viol, "vol_modelo": float(so * np.sqrt(252) * 100), "vol_real": real,
                         "pl_oficial": float(wo.sum() * 100)}}, serie


def uma_vez():
    A, cart, sf = calcula()
    limpa = lambda o: None if isinstance(o, float) and not np.isfinite(o) else o
    dados = {"gerado": datetime.now().strftime("%d/%m/%Y %H:%M"), "fundo": FUNDO, "carteira": cart, "serie_fundo": sf,
             "ativos": [{k: limpa(v) for k, v in r.items()} for r in A.to_dict("records")]}
    t = (vm.PASTA / "pagina_simples.html").read_text(encoding="utf-8")
    h = t.replace("/*DADOS*/", json.dumps(dados, ensure_ascii=False, default=str).replace("</", "<\\/"))
    (vm.SAIDA / "vol_simples.html").write_text(h, encoding="utf-8")
    print("saida/vol_simples.html", f"{len(h) / 1e6:.1f} MB")


def rodar():
    while True:
        try:
            uma_vez()
        except Exception as e:
            print(datetime.now(), "erro:", type(e).__name__, e)
        h = datetime.now().hour + datetime.now().minute / 60
        time.sleep(900 if 9 <= h <= 18.5 and datetime.now().weekday() < 5 else 3600)


if __name__ == "__main__":
    rodar() if len(sys.argv) > 1 and sys.argv[1] == "rodar" else uma_vez()
