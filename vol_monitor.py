"""
Vol monitor de crédito (proof of concept) — um fundo, dados públicos grátis, o mais "ao vivo" possível.

Entrada: Excel com uma linha por ativo:  ativo = "CODIGO - AAAA-MM-DD" (ex.: "AMAG12 - 2028-01-28"), perc_nav, data.
Para cada ativo:
  * identifica (cadastro B3, SND, informe das securitizadoras): tipo, emissor, CNPJ, indexador, taxa de emissão;
  * vol por três caminhos que se completam:
      - B3 negócios de balcão (~15 min de atraso; histórico de ~6 meses no cache do carteira.py) -> vol da taxa e do preço;
      - ANBIMA taxa indicativa diária (debêntures; histórico local cresce a cada dia que roda);
      - CVM: retorno mensal do ativo nas carteiras dos fundos (painel de fundos) -> vol mensal longa (suavizada);
    estimativa final = diária encolhida para a mensal conforme o número de observações (quanto mais roda, mais precisa);
  * carteira: vol do fundo (correlação estimada pelos retornos mensais), contribuição de cada ativo, choque de +100 bps;
  * notícias e eventos: Google Notícias por emissor, fatos relevantes da CVM, agenda de pagamentos (ANBIMA),
    ação do emissor (Yahoo), câmbio/commodities/juros ao vivo (Yahoo, B3).
Uso:
  python vol_monitor.py exemplo     -> cria carteira_exemplo.xlsx (fundo Dual, carteira pública CVM jun/26)
  python vol_monitor.py             -> calcula uma vez e grava saida/vol_monitor.html
  python vol_monitor.py rodar       -> fica rodando: B3/mercado a cada 15 min, notícias a cada hora, ANBIMA no fim do dia
"""
import base64
import html
import json
import re
import sys
import time
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import requests

PASTA = Path(__file__).resolve().parent
DADOS, SAIDA = PASTA / "dados", PASTA / "saida"
DADOS.mkdir(exist_ok=True), SAIDA.mkdir(exist_ok=True)
sys.path[:0] = [str(PASTA.parent), str(Path.home() / "OneDrive/Documents/Personal/itau_database")]
import carteira as ct                       # noqa: E402  (B3: negócios, cadastro, ISINs)
import lookthrough_cvm as lt                # noqa: E402  (SND, ANBIMA, securitizadoras, BCB)

CARTEIRA = PASTA / (sys.argv[2] if len(sys.argv) > 2 else "carteira_exemplo.xlsx")
PAINEL = Path.home() / "OneDrive/Documents/painel-fundos/painel_fundos.html"
UA = {"User-Agent": "Mozilla/5.0"}
K_PRIOR = 20          # observações diárias que "valem" a estimativa mensal (encolhimento)
RHO0 = 0.30           # correlação de crédito quando não há meses em comum suficientes
ALERTAS = ["recuperação judicial", "recuperacao judicial", "default", "inadimpl", "rebaix", "downgrade", "vencimento antecipado",
           "fato relevante", "rating", "aquisição", "fusão", "renegocia", "covenant", "assembleia", "waiver", "prejuízo", "calote",
           "reestrutura", "rescis", "multa", "investigação", "cvm", "falência"]
MERCADO = {"Dólar (R$)": "BRL=X", "Ibovespa": "^BVSP", "Brent (US$)": "BZ=F", "Minério de ferro (US$/t)": "TIO=F", "Soja (¢/bu)": "ZS=F",
           "Milho (¢/bu)": "ZC=F", "Boi gordo CME (¢/lb)": "LE=F", "Açúcar (¢/lb)": "SB=F", "Ouro (US$)": "GC=F", "Treasury 10a (%)": "^TNX",
           "S&P 500": "^GSPC"}
CODIGO = re.compile(r"^\s*([A-Z0-9]{4,15})\s*[-–]\s*(\d{4}-\d{2}-\d{2})")


# ============================================================== entrada
def le_carteira(arq=CARTEIRA):
    d = pd.read_excel(arq, dtype=str)
    col = {c.lower().strip(): c for c in d.columns}
    a, p = col.get("ativo", d.columns[0]), next((col[k] for k in ("perc_nav", "%nav", "perc", "% nav") if k in col), d.columns[1])
    dt = col.get("data")
    out = []
    for _, r in d.iterrows():
        m = CODIGO.match(str(r[a]).upper())
        out.append({"entrada": str(r[a]), "codigo": m.group(1) if m else None, "venc": m.group(2) if m else None,
                    "perc_nav": float(str(r[p]).replace(",", ".")), "data": str(r[dt])[:10] if dt else None})
    return pd.DataFrame(out)


def painel():
    """Dados do painel de fundos (retornos mensais dos ativos nas carteiras CVM, duration, spread)."""
    s = PAINEL.read_text(encoding="utf-8")
    b = max(re.findall(r"<script>(.*?)</script>", s, re.S), key=len)
    return json.JSONDecoder().raw_decode(b[b.index("{"):])[0]


def exemplo(nome_fundo="Dual"):
    """Excel de teste: carteira pública (CVM, look-through) do fundo; ativos com código B3 no formato do usuário."""
    D = painel()
    f = next(x for x in D["fundos"] if x["nome"] == nome_fundo)
    m = f["abertos"][-1]
    A = pd.DataFrame(D["ativos"], columns=["id", "mes", "ch", "rot", "cat", "grupo", "emis", "perc", "rs", "sp", "du", "x"])
    A = A[(A.id == f["id"]) & (A.mes == m)].groupby(["ch", "rot", "cat"], as_index=False).perc.sum().sort_values("perc", ascending=False)
    snd = lt.snd_caracteristicas()
    sec, _ = lt.securitizadoras(lambda x: None)
    v_sec = sec.drop_duplicates("serie").set_index("serie").venc
    cad = ct.cadastro().drop_duplicates("cod").set_index("cod")
    linhas = []
    for r in A.itertuples():
        venc = None
        if r.ch in snd.index:
            venc = pd.Timestamp(snd.venc[r.ch]).strftime("%Y-%m-%d")
        elif r.ch in v_sec.index:
            venc = str(v_sec[r.ch])[:10]
        elif r.ch in cad.index:
            venc = pd.to_datetime(cad.venc[r.ch], dayfirst=True).strftime("%Y-%m-%d")
        linhas.append({"ativo": f"{r.ch} - {venc}" if venc else f"SEM CÓDIGO B3: {r.rot} ({r.cat})", "perc_nav": round(r.perc, 4),
                       "data": (pd.Timestamp(m) + pd.offsets.MonthEnd(0)).strftime("%Y-%m-%d")})
    pd.DataFrame(linhas).to_excel(PASTA / "carteira_exemplo.xlsx", index=False)
    print(f"carteira_exemplo.xlsx: {nome_fundo} {m}, {len(linhas)} linhas, "
          f"{sum(1 for l in linhas if not l['ativo'].startswith('SEM'))} com código")


# ============================================================== identificação
def identifica(c):
    snd = lt.snd_caracteristicas()
    cad = ct.cadastro().drop_duplicates("cod").set_index("cod")
    sec, dev = lt.securitizadoras(lambda x: None)
    s1 = sec.sort_values("mes").drop_duplicates("serie", keep="last").set_index("serie")
    out = []
    for r in c.itertuples():
        x = {"codigo": r.codigo, "venc": r.venc, "perc_nav": r.perc_nav, "entrada": r.entrada, "tipo": None, "emissor": None, "cnpj": None,
             "indexador": None, "taxa_emissao": None, "isin": None, "venc_cadastro": None}
        if r.codigo:
            if r.codigo in cad.index:
                k = cad.loc[r.codigo]
                x.update(tipo=k.tipo, emissor=k.emissor, isin=k.isn, indexador=k.idx, taxa_emissao=ct.num(pd.Series([k.taxa]))[0],
                         venc_cadastro=pd.to_datetime(k.venc, dayfirst=True).strftime("%Y-%m-%d"))
            if r.codigo in snd.index:
                k = snd.loc[r.codigo]
                x.update(tipo=x["tipo"] or "DEB", emissor=k.empresa, cnpj=k.cnpj, indexador=k.indice,
                         taxa_emissao=float(k.juros) if k.juros == k.juros and k.juros else (float(k.pct) if k.pct == k.pct else None),
                         venc_cadastro=pd.Timestamp(k.venc).strftime("%Y-%m-%d"))
            if r.codigo in s1.index:
                k = s1.loc[r.codigo]
                dv = dev[(dev.certificado == k.certificado) & dev.papel.eq("Devedor")].cnpj
                x.update(tipo=k.tipo, indexador=x["indexador"] or k.taxa, cnpj=x["cnpj"] or (dv.iloc[0] if len(dv) else k.cnpj_sec),
                         venc_cadastro=str(k.venc)[:10], isin=x["isin"] or k["isin"])
        out.append(x)
    return pd.DataFrame(out)


# ============================================================== histórico de preço/taxa
def b3_diario(cods, dias=180):
    """Negócios de balcão B3 -> por ativo e dia: taxa e PU médios ponderados pelo volume (só negócios >= R$ 50 mil)."""
    ds = pd.bdate_range(pd.Timestamp.today() - pd.Timedelta(days=dias), pd.Timestamp.today())
    n = ct.negocios(list(ds))
    n = n[n.cod.isin(cods)].copy()
    n["data"] = pd.to_datetime(n.data).dt.normalize()
    bruto = n.copy()
    n = n[n.vol >= 50_000]
    g = n.assign(tv=n.taxa * n.vol, pv=n.pu * n.vol).groupby(["cod", "data"])
    d = g.agg(tv=("tv", "sum"), pv=("pv", "sum"), vol=("vol", "sum"), n=("vol", "size")).reset_index()
    d["taxa"], d["pu"] = d.tv / d.vol, d.pv / d.vol
    return d[["cod", "data", "taxa", "pu", "vol", "n"]], bruto


def anbima_hist(cods):
    _, h = lt.anbima_debentures()
    h = h.assign(cod=h.codigo.str.strip())
    return h[h.cod.isin(cods)][["cod", "data", "taxa_ind", "pu", "pct_par", "duration_anos", "tipo"]]


def vol_taxa(s, datas):
    """Vol anualizada (p.p. a.a.) de uma série de taxas em dias úteis possivelmente espaçados: Δ / sqrt(intervalo), desvio robusto."""
    s, datas = np.asarray(s, float), pd.to_datetime(pd.Series(datas)).values
    if len(s) < 3:
        return np.nan, len(s) - 1 if len(s) else 0
    gap = np.array([max(1, np.busday_count(a.astype("datetime64[D]"), b.astype("datetime64[D]"))) for a, b in zip(datas[:-1], datas[1:])])
    z = np.diff(s) / np.sqrt(gap)
    z = z[np.abs(z - np.median(z)) < 6 * (np.median(np.abs(z - np.median(z))) * 1.4826 + 1e-9)]      # tira negócio fora da curva
    if len(z) < 3:
        return np.nan, len(z)
    v = np.var(z, ddof=1)
    if len(z) >= 10:                     # Roll: ruído de negócio (bid/ask, corretagem do varejo) gera autocorrelação negativa
        c1 = np.cov(z[1:], z[:-1])[0, 1]
        v = max(v + 2 * c1, 0.25 * v)
    return float(np.sqrt(v * 252)), len(z)


# ============================================================== cálculo
def calcula():
    t0 = time.time()
    c = le_carteira()
    a = identifica(c)
    cods = [x for x in a.codigo.dropna()]
    D = painel()
    M, CDI = D["meses"], D["cdi"]
    perf = {p["chave"]: p for p in D["perf"]}
    b3, b3_bruto = b3_diario(cods)
    an = anbima_hist(cods)
    hoje = pd.Timestamp.today().normalize()
    ativos, mens = [], {}
    for x in a.to_dict("records"):
        cod = x["codigo"]
        x["coberto"] = bool(cod)
        if not cod:
            ativos.append(x)
            continue
        p, det = perf.get(cod, {}), D["det"].get(cod) or {}
        # duration: ANBIMA > painel (fundos) > aproximação pelo prazo
        ah = an[an.cod == cod].sort_values("data")
        anos = max((pd.Timestamp(x["venc"]) - hoje).days / 365.25, 0.05)
        dur = (float(ah.duration_anos.dropna().iloc[-1]) if ah.duration_anos.notna().any() else
               p.get("duration") or (det.get("h") or {}).get("du", [None])[-1] or min(anos, anos * 0.8 + 0.2))
        x["duration"] = round(float(dur), 2)
        # 1) mensal (CVM): excesso sobre o CDI
        r = det.get("r") or []
        ex = [(M[i], r[i] - CDI[M[i]]) for i in range(len(r)) if r[i] is not None and M[i] in CDI][-36:]
        if len(ex) >= 6:
            v = np.array([e for _, e in ex])
            x["vol_mensal"], x["n_mensal"] = float(np.std(v, ddof=1) * np.sqrt(12)), len(v)
            mens[cod] = pd.Series(dict(ex))
        else:
            x["vol_mensal"], x["n_mensal"] = None, len(ex)
        # 2) ANBIMA diária (taxa indicativa)
        if len(ah) >= 3:
            vt, n = vol_taxa(ah.taxa_ind.values, ah.data.values)
            x["vol_taxa_anbima"], x["n_anbima"] = float(vt), n
            x["taxa_anbima"], x["pu_anbima"], x["pct_par"] = float(ah.taxa_ind.iloc[-1]), float(ah.pu.iloc[-1]), float(ah.pct_par.iloc[-1])
            x["data_anbima"] = ah.data.iloc[-1].strftime("%d/%m")
        else:
            x["vol_taxa_anbima"], x["n_anbima"] = None, 0
        # 3) B3 negócios
        bh = b3[b3.cod == cod].sort_values("data")
        if len(bh) >= 3:
            vt, n = vol_taxa(bh.taxa.values, bh.data.values)
            vp, _ = vol_taxa(np.log(bh.pu.values) * 100, bh.data.values)      # vol direta do PU (% a.a.): sem a aproximação pela duration
            x["vol_taxa_b3"], x["n_b3"], x["vol_pu_b3"] = float(vt), n, float(vp)
        else:
            x["vol_taxa_b3"], x["n_b3"], x["vol_pu_b3"] = None, max(len(bh) - 1, 0), None
        bb = b3_bruto[b3_bruto.cod == cod]
        if len(bb):
            u = bb.sort_values("data").iloc[-1]
            x["ult_negocio"] = {"data": pd.Timestamp(u.data).strftime("%d/%m"), "taxa": float(u.taxa) if u.taxa == u.taxa else None,
                                "pu": float(u.pu), "vol": float(u.vol)}
            x["neg_30d"] = int((bb.data >= hoje - pd.Timedelta(days=30)).sum())
            x["vol_fin_30d"] = float(bb[bb.data >= hoje - pd.Timedelta(days=30)].vol.sum())
        # combinação: diária (taxa -> preço pela duration) encolhida para a mensal pelo nº de observações
        b3_preco = None
        if x["vol_taxa_b3"] is not None and x["vol_taxa_b3"] == x["vol_taxa_b3"]:
            b3_preco = x["vol_taxa_b3"] * x["duration"]
            if x["vol_pu_b3"] is not None and x["vol_pu_b3"] == x["vol_pu_b3"]:
                b3_preco = min(b3_preco, x["vol_pu_b3"])          # taxa x duration explode em papel estressado (não linear)
        x["vol_preco_b3"] = b3_preco
        x["vol_preco_anbima"] = x["vol_taxa_anbima"] * x["duration"] if x["vol_taxa_anbima"] is not None else None
        x["serie_b3"] = {"d": [d.strftime("%Y-%m-%d") for d in bh.data], "t": [round(v, 4) for v in bh.taxa], "v": [round(v / 1e6, 2) for v in bh.vol]}
        x["serie_anbima"] = {"d": [d.strftime("%Y-%m-%d") for d in ah.data], "t": [round(v, 4) if v == v else None for v in ah.taxa_ind]}
        x["serie_mensal"] = {"m": [m for m, _ in ex], "e": [round(e, 3) for _, e in ex]}
        ativos.append(x)
    A = pd.DataFrame(ativos)
    A, k_b3 = combina(A)
    cart = carteira_risco(A, mens)
    cart["k_b3"] = k_b3
    print(f"cálculo: {len(A)} ativos, {time.time() - t0:.0f}s")
    return A, cart


def combina(A):
    """B3 calibrada pela ANBIMA (k = mediana vol ANBIMA / vol B3 nos ativos com as duas); diária = média por nº de obs.;
    final = diária encolhida para a mensal (CVM) com peso K_PRIOR."""
    for c in ("vol_preco_b3", "vol_preco_anbima", "vol_mensal", "n_b3", "n_anbima"):
        A[c] = pd.to_numeric(A.get(c), errors="coerce")
    ok = (A.n_b3 >= 10) & (A.n_anbima >= 5) & (A.vol_preco_b3 > 0) & (A.vol_preco_anbima > 0)
    k = float(np.clip((A.vol_preco_anbima[ok] / A.vol_preco_b3[ok]).median(), 0.3, 1.0)) if ok.sum() >= 10 else 1.0
    est, nd_, conf = [], [], []
    for r in A.itertuples():
        d = [(v, n) for v, n, mn in ((r.vol_preco_anbima, r.n_anbima, 5), (r.vol_preco_b3 * k if r.vol_preco_b3 == r.vol_preco_b3 else np.nan, r.n_b3, 10))
             if v == v and n == n and n >= mn]
        nd = sum(n for _, n in d)
        vd = sum(v * n for v, n in d) / nd if nd else None
        vm = r.vol_mensal if r.vol_mensal == r.vol_mensal else None
        e = (nd * vd + K_PRIOR * vm) / (nd + K_PRIOR) if vd is not None and vm is not None else (vd if vd is not None else vm)
        nm = r.n_mensal if r.n_mensal == r.n_mensal else 0
        est.append(e), nd_.append(nd), conf.append("alta" if nd >= 60 else "média" if nd >= 20 or nm >= 24 else "baixa" if e is not None else "sem dado")
    A["vol_est"], A["n_diario"], A["confianca"] = est, nd_, conf
    A["vol_b3_calib"] = A.vol_preco_b3 * k
    A.loc[~A.coberto, ["vol_est", "confianca"]] = [np.nan, None]
    return A, k


def carteira_risco(A, mens):
    """Vol do fundo: Σ = D(σ) R D(σ); R = correlação dos excessos mensais (≥12 meses em comum) encolhida para RHO0."""
    k = A[A.vol_est.notna()].copy()
    w, s = k.perc_nav.values / 100, k.vol_est.values
    n = len(k)
    R = np.full((n, n), RHO0)
    cods = k.codigo.tolist()
    for i in range(n):
        for j in range(i + 1, n):
            a, b = mens.get(cods[i]), mens.get(cods[j])
            if a is not None and b is not None:
                cm = a.index.intersection(b.index)
                if len(cm) >= 12 and a[cm].std() > 0 and b[cm].std() > 0:
                    R[i, j] = R[j, i] = 0.5 * np.corrcoef(a[cm], b[cm])[0, 1] + 0.5 * RHO0
    np.fill_diagonal(R, 1)
    S = np.outer(s, s) * R
    vp = float(np.sqrt(w @ S @ w))
    contrib = w * (S @ w) / vp if vp else w * 0
    k["contrib"], k["contrib_pct"] = contrib, contrib / vp * 100 if vp else 0
    A.loc[k.index, "contrib"], A.loc[k.index, "contrib_pct"] = k.contrib, k.contrib_pct
    choque = float(-(k.perc_nav / 100 * k.duration).sum() * 1.0)        # +100 bps de spread em tudo, % do PL
    em = A[A.coberto].groupby("emissor", dropna=False).perc_nav.sum().sort_values(ascending=False)
    return {"vol_pl": vp, "vol_sem_div": float((w * s).sum()), "cobertura_pl": float(k.perc_nav.sum()), "pl_total": float(A.perc_nav.sum()),
            "pl_com_codigo": float(A[A.coberto].perc_nav.sum()), "choque_100bps": choque, "n": n,
            "duration_media": float((k.perc_nav * k.duration).sum() / max(k.perc_nav.sum(), 1e-9)),
            "emissores": [{"emissor": e, "perc": float(v)} for e, v in em.head(12).items()]}


# ============================================================== mercado, ações, notícias, eventos
def yahoo(tk, rng="3mo"):
    try:
        j = requests.get(f"https://query1.finance.yahoo.com/v8/finance/chart/{tk}?range={rng}&interval=1d", headers=UA, timeout=20).json()
        r = j["chart"]["result"][0]
        c = pd.Series(r["indicators"]["quote"][0]["close"], index=pd.to_datetime(r["timestamp"], unit="s")).dropna()
        return {"preco": r["meta"]["regularMarketPrice"], "hora": datetime.fromtimestamp(r["meta"]["regularMarketTime"]).strftime("%d/%m %H:%M"),
                "d1": (r["meta"]["regularMarketPrice"] / c.iloc[-2] - 1) * 100 if len(c) > 1 else None,
                "m1": (c.iloc[-1] / c.iloc[-22] - 1) * 100 if len(c) > 22 else None,
                "vol3m": float(np.log(c).diff().std() * np.sqrt(252) * 100), "serie": [round(v, 4) for v in c.values[-66:]],
                "datas": [d.strftime("%Y-%m-%d") for d in c.index[-66:]]}
    except Exception:
        return None


def juros_b3():
    """DI1 e DAP ao vivo (B3, ~15 min): taxa corrente por vencimento."""
    out = {}
    for t in ("DI1", "DAP"):
        try:
            j = ct.get(f"https://cotacao.b3.com.br/mds/api/v1/DerivativeQuotation/{t}", timeout=30, tent=2).json()
        except Exception:
            continue
        xs = []
        for s in j.get("Scty", []):
            m, q = s.get("asset", {}).get("AsstSummry", {}).get("mtrtyCode"), s.get("SctyQtn", {})
            if m and q.get("curPrc"):
                xs.append(((pd.Timestamp(m) - pd.Timestamp.today()).days / 365.25, q["curPrc"], q.get("prvsDayAdjstmntPric")))
        xs = sorted(x for x in xs if x[0] > 0.2)
        pega = lambda a: min(xs, key=lambda x: abs(x[0] - a)) if xs else None
        out[t] = {f"{a}a": {"taxa": p[1], "var_bps": (p[1] - p[2]) * 100 if p and p[2] else None} for a in (1, 2, 3, 5, 10) if (p := pega(a))}
    return out


def acoes_emissores(cnpjs):
    """CNPJ do emissor -> empresa listada na B3 (pela raiz do CNPJ) -> ação no Yahoo."""
    f = DADOS / "b3_listadas.parquet"
    if not f.exists() or time.time() - f.stat().st_mtime > 7 * 86400:
        L = []
        for pg in range(1, 30):
            q = base64.b64encode(json.dumps({"language": "pt-br", "pageNumber": pg, "pageSize": 120}).encode()).decode()
            try:
                j = ct.get(f"https://sistemaswebb3-listados.b3.com.br/listedCompaniesProxy/CompanyCall/GetInitialCompanies/{q}", timeout=60).json()
            except Exception:
                break
            L += j.get("results", [])
            if pg >= j.get("page", {}).get("totalPages", 0):
                break
        pd.DataFrame(L).to_parquet(f)
    L = pd.read_parquet(f)
    L["raiz"] = L.cnpj.astype(str).str.replace(r"\D", "", regex=True).str.zfill(14).str[:8]
    out = {}
    for c in {re.sub(r"\D", "", str(x)).zfill(14)[:8] for x in cnpjs if x}:
        m = L[L.raiz == c]
        if not len(m):
            continue
        cod = m.issuingCompany.iloc[0]
        for suf in ("3", "4", "11", "5", "6"):
            y = yahoo(f"{cod}{suf}.SA")
            if y:
                out[c] = {"ticker": cod + suf, "empresa": m.companyName.iloc[0], **y}
                break
    return out


def noticias(nome, dias=14):
    q = re.sub(r"\b(S\.?A\.?|S/A|LTDA\.?|PARTICIPA[CÇ][OÕ]ES|HOLDING|COMPANHIA|CIA\.?|DE|DO|DA|E|INVESTIMENTOS?|SECURITIZADORA)\b", " ",
               nome.upper())
    q = " ".join(q.split()[:3]).title()
    try:
        r = requests.get("https://news.google.com/rss/search", params={"q": f'"{q}" when:{dias}d', "hl": "pt-BR", "gl": "BR", "ceid": "BR:pt-419"},
                         headers=UA, timeout=20)
        it = ET.fromstring(r.content).findall(".//item")
    except Exception:
        return q, []
    out = []
    for i in it[:8]:
        t = i.findtext("title") or ""
        out.append({"titulo": t, "link": i.findtext("link"), "data": pd.to_datetime(i.findtext("pubDate")).strftime("%d/%m"),
                    "fonte": (i.find("source").text if i.find("source") is not None else ""),
                    "alerta": [k for k in ALERTAS if k in t.lower()]})
    return q, out


def cvm_ipe(cnpjs, dias=45):
    f = DADOS / f"ipe_{date.today():%Y%m%d}.parquet"
    if not f.exists():
        import io
        import zipfile
        r = requests.get(f"https://dados.cvm.gov.br/dados/CIA_ABERTA/DOC/IPE/DADOS/ipe_cia_aberta_{date.today().year}.zip", headers=UA, timeout=180)
        z = zipfile.ZipFile(io.BytesIO(r.content))
        pd.read_csv(z.open(z.namelist()[0]), sep=";", encoding="latin1", dtype=str).to_parquet(f)
    d = pd.read_parquet(f)
    d["raiz"] = d.CNPJ_Companhia.str.replace(r"\D", "", regex=True).str[:8]
    raizes = {re.sub(r"\D", "", str(x)).zfill(14)[:8] for x in cnpjs if x}
    d = d[d.raiz.isin(raizes) & (pd.to_datetime(d.Data_Entrega) >= pd.Timestamp.today() - pd.Timedelta(days=dias))]
    txt = (d.Assunto.fillna("") + " " + d.Tipo.fillna("")).str.lower()
    d = d[d.Categoria.isin(["Fato Relevante", "Aviso aos Debenturistas"]) | d.Categoria.str.contains("Debentur", na=False)
          | (d.Categoria.eq("Comunicado ao Mercado") & ~txt.str.contains("período de silêncio|agenda divulga|calendário"))
          | (d.Categoria.eq("Assembleia") & txt.str.contains("agd|debentur|waiver|titulares"))]
    return d.sort_values("Data_Entrega", ascending=False)


def agenda_anbima(cods):
    """Próximos eventos (juros, amortização, vencimento) da ANBIMA Data — página pública, Edge headless; cache diário."""
    f = DADOS / f"agenda_{date.today():%Y%m%d}.json"
    cache = json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
    falta = [c for c in cods if c not in cache]
    if falta:
        try:
            from selenium import webdriver
            o = webdriver.EdgeOptions()
            o.add_argument("--headless=new")
            d = webdriver.Edge(options=o)
            try:
                for c in falta:
                    tipo = "certificado-de-recebiveis" if c.startswith(("CRA", "CRI")) or re.match(r"^\d{2}[A-L]", c) else "debentures"
                    d.get(f"https://data.anbima.com.br/{tipo}/{c}/agenda")
                    time.sleep(5)
                    t = d.find_element("tag name", "body").text
                    ln, ev = t.splitlines(), []
                    pad = re.compile(r"(\d{2}/\d{2}/\d{4}) (\d{2}/\d{2}/\d{4}) ([^\d/]+?) ([\d.,]+ %|-) (R\$ [\d.,]+|-)(?: -)?(?= \d{2}/\d{2}/\d{4} \d{2}/\d{2}/\d{4}|$)")
                    for i, l in enumerate(ln[:-1]):          # eventos "dd/mm/aaaa dd/mm/aaaa Evento x % R$ y|-"; status na linha seguinte
                        st = ln[i + 1].strip()
                        for m in pad.finditer(l.strip()):
                            ev.append({"data": m.group(1), "evento": m.group(3), "pct": m.group(4),
                                       "status": "LIQUIDADO" if m.group(5) != "-" else (st if st in ("PREVISTO", "LIQUIDADO") else "PREVISTO")})
                    cache[c] = ev
            finally:
                d.quit()
        except Exception as e:
            print("agenda ANBIMA:", type(e).__name__)
        f.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
    hoje = pd.Timestamp.today()
    return {c: [e for e in v if e["status"] == "PREVISTO" and hoje.normalize() <= pd.to_datetime(e["data"], dayfirst=True) <= hoje + pd.Timedelta(days=120)][:4]
            for c, v in cache.items() if c in cods}


def contexto(A):
    cob = A[A.coberto]
    em = cob.groupby(["emissor", "cnpj"], dropna=False).perc_nav.sum().reset_index().sort_values("perc_nav", ascending=False)
    top = em.head(15)
    news = {}
    for r in top.itertuples():
        if r.emissor:
            q, it = noticias(r.emissor)
            news[r.emissor] = {"busca": q, "itens": it, "perc": r.perc_nav}
    try:
        ipe = cvm_ipe(cob.cnpj.dropna().unique())
        ipe = [{"data": r.Data_Entrega[:10], "empresa": r.Nome_Companhia, "categoria": r.Categoria, "assunto": r.Assunto or r.Tipo or "",
                "link": r.Link_Download} for r in ipe.head(40).itertuples()]
    except Exception as e:
        print("CVM IPE:", type(e).__name__)
        ipe = []
    top_cods = cob.sort_values("perc_nav", ascending=False).codigo.head(25).tolist()
    return {"noticias": news, "ipe": ipe, "agenda": agenda_anbima(top_cods), "acoes": acoes_emissores(cob.cnpj.dropna().unique()),
            "mercado": {k: yahoo(v) for k, v in MERCADO.items()}, "juros": juros_b3()}


# ============================================================== página
def pagina(A, cart, ctx, gerado):
    limpa = lambda o: None if isinstance(o, float) and not np.isfinite(o) else o
    dados = {"gerado": gerado, "carteira": cart, "ativos": [{k: limpa(v) for k, v in r.items()} for r in A.to_dict("records")], **ctx,
             "param": {"K_PRIOR": K_PRIOR, "RHO0": RHO0}}
    t = (PASTA / "pagina.html").read_text(encoding="utf-8")
    crit = (PASTA / "critica.html").read_text(encoding="utf-8") if (PASTA / "critica.html").exists() else "<p>(crítica ainda não escrita)</p>"
    t = t.replace("/*CRITICA*/", crit)
    return t.replace("/*DADOS*/", json.dumps(dados, ensure_ascii=False, default=str).replace("</", "<\\/"))


def uma_vez():
    A, cart = calcula()
    ctx = contexto(A)
    h = pagina(A, cart, ctx, datetime.now().strftime("%d/%m/%Y %H:%M"))
    (SAIDA / "vol_monitor.html").write_text(h, encoding="utf-8")
    A.drop(columns=["serie_b3", "serie_anbima", "serie_mensal"], errors="ignore").to_parquet(DADOS / f"foto_{datetime.now():%Y%m%d_%H%M}.parquet")
    print("saida/vol_monitor.html", f"{len(h) / 1e6:.1f} MB")


def rodar():
    """Loop: recalcula a cada 15 min no pregão (9h-18h30), a cada hora fora; ANBIMA atualiza sozinha no 1o ciclo após 20h."""
    while True:
        try:
            uma_vez()
        except Exception as e:
            print(datetime.now(), "erro:", type(e).__name__, e)
        h = datetime.now().hour + datetime.now().minute / 60
        time.sleep(900 if 9 <= h <= 18.5 and datetime.now().weekday() < 5 else 3600)


if __name__ == "__main__":
    {"exemplo": exemplo, "rodar": rodar}.get(sys.argv[1] if len(sys.argv) > 1 else "", uma_vez)()
