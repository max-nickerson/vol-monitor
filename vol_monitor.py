"""
Vol monitor de crédito — um arquivo só.

Para o gestor: o que está acontecendo, o que vem por aí e o que historicamente antecipa o preço de cada ativo da carteira.

Entrada: Excel com uma linha por ativo — ativo "CODIGO - AAAA-MM-DD" (ex.: "HAPV18 - 2030-01-27"), perc_nav (% do PL) e data.
Uso:
    python vol_monitor.py carteira.xlsx           calcula uma vez -> saida/vol_monitor.html
    python vol_monitor.py carteira.xlsx rodar     recalcula a cada 15 min no pregão (1 h fora dele)
    python vol_monitor.py modelo                  cria carteira_modelo.xlsx (15 debêntures mais negociadas no último mês)

Fontes (todas públicas e gratuitas; cache em dados/):
    ANBIMA Data (página pública de cada ativo): PU par e PU indicativo diários desde a emissão, eventos, agenda de pagamentos
        e o topo da página (taxa, duration) — o Edge abre a página e guarda o que a própria página recebe (sem token, sem API).
    ANBIMA arquivo diário de debêntures: taxa indicativa, PU, % PU par e duration (o site guarda ~5 semanas; o cache cresce).
    B3 Boletim Diário: negócios de balcão de crédito (desde 03/04/2023; ~15 min de atraso) e cadastro de instrumentos.
    SND (debentures.com.br): características das debêntures e preços de negociação.
    CVM: ITR/DFP (fundamentos trimestrais; data de entrega = quando o mercado soube), IPE (fatos relevantes, comunicados),
        FCA (códigos de negociação das ações), informe mensal das securitizadoras (devedor dos CRI/CRA).
    Yahoo Finance: ações dos emissores (preço ajustado), Ibovespa, dólar, commodities, Treasury.
    SmartWealth (smartwealth.com.br): índices de crédito e de renda fixa (IDA, JGP, IMA) — só lidos localmente, não redistribuídos.
    BCB (CDI, Focus), Tesouro Direto (curvas pré e IPCA+), B3 (DI1/DAP ao vivo), IBGE (calendário do IPCA), Google Notícias.

Contas (o PnL é sempre lido contra a taxa indicativa):
    retorno do dia = (PU indicativo + pagamentos) / PU indicativo anterior − 1  (amortização e juros não são perda)
    PnL = carrego (PU par + diferença taxa × cupom) − duration modificada × Δ taxa indicativa + resíduo
    taxa antes do arquivo diário = taxa implícita no PU oficial (inverte o próprio cálculo da ANBIMA; erro ~1,5 bps)
    vol = desvio dos retornos diários oficiais (12 m); sem ANBIMA: negócios semanais com o ruído medido; sem nada: pares
    VaR 95% paramétrico com correlação dos retornos oficiais encolhida 30%; backtest de 250 dias
    previsibilidade: correlação com intervalo de 95%, n e taxa de acerto FORA da amostra (último terço do tempo)
"""
import html as H
import io
import json
import os
import re
import sys
import time
import xml.etree.ElementTree as ET
import zipfile
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime
from pathlib import Path

import numpy as np
import pandas as pd
import requests

PASTA = Path(__file__).resolve().parent
DADOS, SAIDA = PASTA / "dados", PASTA / "saida"
for _p in ("b3", "anbima", "anbima_hist", "snd", "cvm", "yahoo", "indices", "securit"):
    (DADOS / _p).mkdir(parents=True, exist_ok=True)
SAIDA.mkdir(exist_ok=True)
S = requests.Session()
S.headers["User-Agent"] = "Mozilla/5.0"
HOJE = pd.Timestamp.today().normalize()
MIN_DIA = 50_000            # negócio mínimo (R$) para entrar no preço do dia
NEG_INST = 1_000_000        # volume do dia (R$) para o sinal negócio x marcação (abaixo disso é ruído de varejo)
B3_INICIO = pd.Timestamp("2023-04-03")
ANOS_CVM = 5
MERCADO = {"Ibovespa": "^BVSP", "Dólar": "BRL=X", "Brent": "BZ=F", "Minério": "TIO=F", "Soja": "ZS=F", "Milho": "ZC=F",
           "Boi (CME)": "LE=F", "Açúcar": "SB=F", "Ouro": "GC=F", "Treasury 10a": "^TNX", "S&P 500": "^GSPC"}
INDICES = {"IDA DI": "credito/anbima-ida-di-ida-di-index", "IDA IPCA ex-infra": "credito/anbima-ida-ipca-ex-infraestrutura-ida-ipca-ex-infrastructure-index",
           "IDA Geral": "credito/anbima-ida-geral-ida-geral-index", "JGP CDI Geral": "credito/jgp-cdi-general-index",
           "JGP CDI Low Rated": "credito/jgp-cdi-low-rated-index", "JGP Infra": "credito/jgp-infrastructure-general-index",
           "IMA-B": "renda-fixa/anbima-ima-b-ima-b-index", "IMA-B 5": "renda-fixa/anbima-ima-b-5-ima-b-5-index",
           "IRF-M": "renda-fixa/anbima-irf-m-irf-m-index", "CDI": "taxas/cdi-certificado-deposito-interbancario",
           "IEE (energia)": "acoes/iee-indice-de-energia-eletrica-b3", "IMOB": "acoes/imob-indice-imobiliario-b3",
           "IMAT (materiais)": "acoes/imat-indice-de-materiais-basicos-b3", "ICON (consumo)": "acoes/icon-b3"}
# Estudo intradiário (447 debêntures/CRI/CRA, negócios B3 04/2023–10/2026, teste fora da amostra no último terço):
# fração do gap (negócio − ANBIMA de ontem, em % do PU par) que a ANBIMA do dia anda, por volume do papel no dia; acerto do sinal
FRACAO_GAP = [(1e6, 0.0, 52), (5e6, 0.07, 58), (20e6, 0.30, 73), (float("inf"), 0.39, 82)]     # (até R$, fração, acerto %)
B_EMISSOR, B_MERCADO, B_ACAO = 0.065, 0.10, 1.8      # outros papéis do emissor; tom do mercado (sem negócio); bps de % par por 1% da ação
ESTUDO = {"papeis": 447, "dias_b3": 814, "periodo": "04/2023 a 10/2026", "linhas": [
    ["Negócio ≥ R$ 20 mi no dia, longe da ANBIMA de ontem", "mesmo dia", "anda 39% do gap", 82, 52],
    ["Negócio R$ 5–20 mi", "mesmo dia", "anda 30% do gap", 73, 50],
    ["Negócio R$ 1–5 mi", "mesmo dia", "anda 7% do gap", 58, 51],
    ["Negócio < R$ 1 mi (varejo)", "mesmo dia", "ruído", 52, 51],
    ["Negócio ≥ R$ 1 mi até as 11h", "mesmo dia", "anda 24% do gap", 71, 52],
    ["Negócio depois das 17h", "dia seguinte", "+4% do gap", 55, 54],
    ["Origem 'Registro' (vs Voice)", "mesmo dia", "18% vs 11% do gap", 76, 57],
    ["Juros reais do dia (IMA-B) em papéis IPCA", "mesmo dia", "r = 0,68", 83, 56],
    ["Ação do emissor no dia", "mesmo dia", "1,8 bps de % par por 1%", 57, 51],
    ["Outros papéis do emissor negociados (papel sem negócio)", "mesmo dia / 5 dias", "6,5% do gap", 57, 54],
    ["Tom do mercado (gap médio, R$ 50 mi+)", "mesmo dia", "22% do gap médio", 60, 50],
    ["Gap de ontem não fechado / movimento de ontem da ANBIMA", "dia seguinte", "nada", 50, 51]]}
ESTUDO5 = {"texto": "Teste fora da amostra (walk-forward mensal, 04/2024–10/2026, 161 mil previsões, 437 papéis, embargo de 12 dias)",
           "cobertura": 77, "largura_mediana": 111, "largura_ingenua": 117, "ic": 0.078, "ic_t": 5.7,
           "decis": [["pior 10%", -5.3, 46], ["2º", -2.2, 47], ["3º", -2.1, 48], ["5º", 0.6, 52], ["8º", 2.3, 54], ["9º", 3.1, 55], ["melhor 10%", 1.0, 55]],
           "features": [["Semana contra o mercado / longe da média de 20 dias", "reverte (cai → volta)", -4.5],
                        ["% do par (abaixo do par)", "sobe (carrego)", -6.1], ["Negócios do mês acima da marcação", "ANBIMA sobe", 3.0],
                        ["Varejo da semana acima da marcação", "ANBIMA sobe", 2.8], ["Ação do emissor 5–20 dias", "mesmo sentido", 2.2],
                        ["Vol, volume, dias sem negócio, lado dos negócios", "nada", 0]]}
# Regras de 1 a 20 dias úteis (estudo nas 100 debêntures mais líquidas, 04/2023–10/2026, estáveis em 2023-24 e 2025-26):
# mediana, faixa 10–90% e acerto do sentido; 'contra o mercado' = papel − média do mesmo indexador
REGRAS_DIAS = {"di_pior": {"nome": "DI+: semana muito pior que o mercado (≤ −2 desvios)", "sentido": -1, "alvo": "contra o mercado", "papeis": 30, "h": {"1": {"n": 2017, "mediana": -1.4, "p10": -13.5, "p90": 8.7, "acerto": 62.3, "base": 52.1, "acerto_2324": 59.3, "acerto_2526": 64.6}, "5": {"n": 1876, "mediana": -6.8, "p10": -38.6, "p90": 18.5, "acerto": 68.8, "base": 53.1, "acerto_2324": 65.2, "acerto_2526": 71.5}, "10": {"n": 1731, "mediana": -11.5, "p10": -62.8, "p90": 35.9, "acerto": 69.4, "base": 53.6, "acerto_2324": 66.7, "acerto_2526": 71.5}, "20": {"n": 1524, "mediana": -18.1, "p10": -81.4, "p90": 82.9, "acerto": 68.6, "base": 55.5, "acerto_2324": 67.4, "acerto_2526": 69.7}}}, "di_melhor": {"nome": "DI+: semana muito melhor que o mercado (≥ +50 bps)", "sentido": 1, "alvo": "contra o mercado", "papeis": 26, "h": {"1": {"n": 1037, "mediana": 3.1, "p10": -32.8, "p90": 42.6, "acerto": 58.7, "base": 47.9, "acerto_2324": 61.5, "acerto_2526": 56.7}, "5": {"n": 964, "mediana": 14.0, "p10": -46.4, "p90": 129.5, "acerto": 67.6, "base": 46.9, "acerto_2324": 70.4, "acerto_2526": 65.5}, "10": {"n": 904, "mediana": 26.8, "p10": -69.6, "p90": 197.5, "acerto": 70.1, "base": 46.4, "acerto_2324": 71.6, "acerto_2526": 69.0}, "20": {"n": 797, "mediana": 42.9, "p10": -117.0, "p90": 284.8, "acerto": 70.1, "base": 44.5, "acerto_2324": 72.8, "acerto_2526": 68.1}}}, "ipca_extremo": {"nome": "IPCA: queda extrema na semana (≤ −4 desvios) — repique", "sentido": 1, "alvo": "contra o mercado", "papeis": 28, "h": {"1": {"n": 187, "mediana": -6.7, "p10": -31.2, "p90": 27.1, "acerto": 36.4, "base": 49.6, "acerto_2324": None, "acerto_2526": 36.8}, "5": {"n": 176, "mediana": 28.5, "p10": -63.6, "p90": 96.2, "acerto": 72.7, "base": 50.3, "acerto_2324": None, "acerto_2526": 74.7}, "10": {"n": 115, "mediana": 19.9, "p10": -65.5, "p90": 152.7, "acerto": 68.7, "base": 49.8, "acerto_2324": None, "acerto_2526": 71.8}, "20": {"n": 83, "mediana": 48.2, "p10": -195.0, "p90": 254.2, "acerto": 62.7, "base": 50.4, "acerto_2324": None, "acerto_2526": 66.7}}}, "di_negocio": {"nome": "DI+: negócio ≥ R$ 20 mi ≥ 30 bps acima da ANBIMA", "sentido": 1, "alvo": "% par total", "papeis": 19, "h": {"1": {"n": 54, "mediana": 8.1, "p10": -8.2, "p90": 39.9, "acerto": 74.1, "base": 50.9, "acerto_2324": 71.4, "acerto_2526": 78.9}, "5": {"n": 54, "mediana": 34.6, "p10": -10.4, "p90": 126.9, "acerto": 79.6, "base": 50.8, "acerto_2324": 82.9, "acerto_2526": 73.7}, "10": {"n": 53, "mediana": 50.0, "p10": -6.1, "p90": 238.6, "acerto": 81.1, "base": 51.9, "acerto_2324": 85.7, "acerto_2526": 72.2}, "20": {"n": 49, "mediana": 54.2, "p10": -11.0, "p90": 329.1, "acerto": 85.7, "base": 52.8, "acerto_2324": 90.6, "acerto_2526": 76.5}}}, "semana_acima": {"nome": "Semana muito melhor com negócios acima da ANBIMA", "sentido": 1, "alvo": "% par total", "papeis": 55, "h": {"1": {"n": 280, "mediana": 2.3, "p10": -11.1, "p90": 30.7, "acerto": 60.7, "base": 52.1, "acerto_2324": 57.9, "acerto_2526": 62.7}, "5": {"n": 264, "mediana": 12.0, "p10": -22.9, "p90": 68.5, "acerto": 72.3, "base": 53.6, "acerto_2324": 73.5, "acerto_2526": 71.5}, "10": {"n": 249, "mediana": 20.4, "p10": -18.3, "p90": 101.6, "acerto": 70.7, "base": 54.8, "acerto_2324": 68.0, "acerto_2526": 72.6}, "20": {"n": 212, "mediana": 26.7, "p10": -29.2, "p90": 162.6, "acerto": 69.8, "base": 56.1, "acerto_2324": 70.3, "acerto_2526": 69.4}}}}
ALERTAS = ["recuperação judicial", "recuperacao judicial", "default", "inadimpl", "rebaix", "downgrade", "vencimento antecipado",
           "rating", "aquisição", "fusão", "renegocia", "covenant", "waiver", "prejuízo", "reestrutura", "falência", "investigação",
           "assembleia", "multa", "rescis", "calote"]
PARADAS = {"S", "SA", "S/A", "LTDA", "CIA", "COMPANHIA", "DE", "DO", "DA", "DOS", "DAS", "E", "HOLDING", "PARTICIPACOES",
           "PARTICIPAÇÕES", "INVESTIMENTOS", "BRASIL", "GRUPO", "SECURITIZADORA", "SOCIEDADE", "ANONIMA", "EMPREENDIMENTOS"}


# ================================================================== utilidades
def get(url, metodo="GET", tent=3, timeout=120, **kw):
    for i in range(tent):
        try:
            r = S.request(metodo, url, timeout=timeout, **kw)
            r.raise_for_status()
            return r
        except requests.RequestException:
            if i == tent - 1:
                raise
            time.sleep(3 * (i + 1))


def velho(f, horas=20):
    return not f.exists() or time.time() - f.stat().st_mtime > horas * 3600


def num(s):
    return pd.to_numeric(pd.Series(s).astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce")


def raiz(cnpj):
    d = re.sub(r"\D", "", str(cnpj or ""))
    return d.zfill(14)[:8] if len(d) >= 8 and set(d) != {"0"} else None


def dias_uteis(a, b):
    return max(1, int(np.busday_count(pd.Timestamp(a).date(), pd.Timestamp(b).date())))


def fisher(r, n):
    """Intervalo de 95% da correlação (Fisher)."""
    if r is None or not np.isfinite(r) or n < 5:
        return None, None
    z, se = np.arctanh(np.clip(r, -0.999, 0.999)), 1 / np.sqrt(n - 3)
    return float(np.tanh(z - 1.96 * se)), float(np.tanh(z + 1.96 * se))


def corr(a, b, rank=False):
    """Correlação (Pearson; rank=True: Spearman, robusta a semanas extremas como um crash de ação)."""
    a, b = np.asarray(a, float), np.asarray(b, float)
    k = np.isfinite(a) & np.isfinite(b)
    if k.sum() < 8 or a[k].std() == 0 or b[k].std() == 0:
        return None, int(k.sum())
    x, y = (pd.Series(a[k]).rank().values, pd.Series(b[k]).rank().values) if rank else (a[k], b[k])
    return float(np.corrcoef(x, y)[0, 1]), int(k.sum())


def log(m):
    print(f"{datetime.now():%H:%M:%S} {m}", flush=True)


# ================================================================== entrada
CODIGO = re.compile(r"^\s*([A-Z0-9]{4,15})\s*[-–]\s*(\d{4}-\d{2}-\d{2})")


def le_carteira(arq):
    d = pd.read_excel(arq, dtype=str)
    col = {c.lower().strip(): c for c in d.columns}
    a = col.get("ativo", d.columns[0])
    p = next((col[k] for k in ("perc_nav", "%nav", "perc", "% nav", "% pl") if k in col), d.columns[1])
    out = []
    for _, r in d.iterrows():
        m = CODIGO.match(str(r[a]).upper())
        out.append({"entrada": str(r[a]), "codigo": m.group(1) if m else None, "venc": m.group(2) if m else None,
                    "perc_nav": float(str(r[p]).replace(",", "."))})
    return pd.DataFrame(out)


def modelo():
    """carteira_modelo.xlsx: as 15 debêntures com mais volume na B3 no último mês, pesos iguais."""
    n = negocios(list(pd.bdate_range(HOJE - pd.Timedelta(days=30), HOJE)))
    n = n[n.tipo.eq("DEB")]
    top = n.groupby("cod").vol.sum().sort_values(ascending=False).head(15).index
    snd = snd_caracteristicas()
    L = [{"ativo": f"{c} - {pd.Timestamp(snd.venc[c]):%Y-%m-%d}", "perc_nav": 2.0} for c in top if c in snd.index and pd.notna(snd.venc[c])]
    pd.DataFrame(L).to_excel(PASTA / "carteira_modelo.xlsx", index=False)
    log(f"carteira_modelo.xlsx: {len(L)} ativos")


# ================================================================== B3
def b3_csv(nome, dia):
    r = get("https://arquivos.b3.com.br/bdi/table/export/csv?lang=pt-BR", "POST", timeout=300,
            json={"Name": nome, "Date": f"{dia:%Y-%m-%d}", "FinalDate": f"{dia:%Y-%m-%d}", "ClientId": "", "Filters": {}})
    t = r.content.decode("latin1").splitlines()
    i = next((k for k, l in enumerate(t) if l.count(";") > 5), None)
    return pd.DataFrame() if i is None else pd.read_csv(io.StringIO("\n".join(t[i:])), sep=";", dtype=str)


def cadastro_b3():
    f = DADOS / "b3_cadastro.parquet"
    if velho(f):
        for k in range(8):
            try:
                d = b3_csv("InstrumentRegistration", HOJE - pd.Timedelta(days=k))
            except requests.RequestException:
                continue
            if len(d) > 1000 and (d.iloc[:, 3] == "DEB").sum() > 1000:
                d.columns = ["cod", "isn", "emissor", "tipo", "incent", "serie", "emissao_n", "idx", "pct", "taxa", "base", "venc", "qtd",
                             "pu_emis", "restrito", "temis", "oferta"][:d.shape[1]]
                d[d.tipo != "COE"].to_parquet(f)
                break
    return pd.read_parquet(f) if f.exists() else pd.DataFrame(columns=["cod", "isn", "emissor", "tipo", "idx", "taxa", "venc"])


def negocios(dias, cods=None):
    """Negócio a negócio do balcão B3 (crédito). Um arquivo por dia; o dia de hoje é rebaixado a cada 15 min."""
    cols = ["tipo", "emissor", "cod", "qtd", "pu", "vol", "taxa", "origem", "hora", "data", "id", "isn", "liq", "situacao"]

    def um(d):
        f = DADOS / "b3" / f"neg_{d:%Y%m%d}.parquet"
        fresco = f.exists() and (pd.Timestamp(f.stat().st_mtime, unit="s") > d + pd.Timedelta(days=1, hours=12)
                                 or (d == HOJE and time.time() - f.stat().st_mtime < 900))
        if not fresco:
            try:
                x = b3_csv("Trade", d)
            except requests.RequestException:
                return f if f.exists() else None
            if len(x):
                x.columns = cols[:x.shape[1]]
                x = x[x.tipo.isin(["DEB", "CRI", "CRA", "LF", "LFSN", "LFSC", "NC", "CFF"]) & ~x.situacao.fillna("").str.contains("Cancelado")]
                x = x.assign(qtd=num(x.qtd).values, pu=num(x.pu).values, vol=num(x.vol).values, taxa=num(x.taxa).values,
                             data=pd.to_datetime(x.data, dayfirst=True))
            else:
                x = pd.DataFrame(columns=cols)
            x[["tipo", "cod", "isn", "qtd", "pu", "vol", "taxa", "hora", "data"]].to_parquet(f)
        return f

    mes0 = HOJE.replace(day=1) - pd.offsets.MonthBegin(1)                # meses fechados (antes do mês passado) ficam num arquivo só
    fechados = sorted({d.strftime("%Y%m") for d in dias if d < mes0})
    for m in fechados:
        fm = DADOS / "b3" / f"mes_{m}.parquet"
        if not fm.exists():
            ds = [d for d in dias if d.strftime("%Y%m") == m]
            with ThreadPoolExecutor(6) as ex:
                fs_ = [f for f in ex.map(um, ds) if f is not None]
            ps = [pd.read_parquet(f) for f in fs_]
            ps = [p_.astype({"cod": str}) for p_ in ps if len(p_)]
            if ps:
                pd.concat(ps, ignore_index=True).to_parquet(fm)
    with ThreadPoolExecutor(6) as ex:
        fs = [f for f in ex.map(um, [d for d in dias if d >= mes0]) if f is not None]
    fs = [DADOS / "b3" / f"mes_{m}.parquet" for m in fechados if (DADOS / "b3" / f"mes_{m}.parquet").exists()] + fs
    fil = [("cod", "in", list(cods))] if cods else None
    partes = []
    for f in fs:
        try:
            x = pd.read_parquet(f, filters=fil)
        except Exception:                                     # dia sem negócio: coluna sem tipo
            x = pd.read_parquet(f)
            x = x[x.cod.isin(cods)] if cods else x
        if len(x):
            partes.append(x)
    n = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame(columns=["tipo", "cod", "data", "taxa", "pu", "vol", "hora"])
    n["data"] = pd.to_datetime(n.data).dt.normalize()
    return n


def juros_b3():
    """DI1 e DAP ao vivo (~15 min): taxa e variação do dia por prazo."""
    out = {}
    for t in ("DI1", "DAP"):
        try:
            j = get(f"https://cotacao.b3.com.br/mds/api/v1/DerivativeQuotation/{t}", timeout=30, tent=2).json()
        except Exception:
            continue
        xs = []
        for s in j.get("Scty", []):
            m, q = s.get("asset", {}).get("AsstSummry", {}).get("mtrtyCode"), s.get("SctyQtn", {})
            if m and q.get("curPrc"):
                xs.append(((pd.Timestamp(m) - HOJE).days / 365.25, q["curPrc"], q.get("prvsDayAdjstmntPric")))
        xs = sorted(x for x in xs if x[0] > 0.2)
        perto = lambda a: min(xs, key=lambda x: abs(x[0] - a)) if xs else None
        out[t] = [{"prazo": f"{a}a", "taxa": p[1], "bps": (p[1] - p[2]) * 100 if p[2] else None} for a in (1, 2, 3, 5, 10) if (p := perto(a))]
    return out


# ================================================================== SND
SND = "https://www.debentures.com.br/exploreosnd/consultaadados/"


def _tabela_snd(b, inicio):
    t = b.decode("latin1").splitlines()
    i = next((j for j, l in enumerate(t) if l.startswith(inicio)), None)
    if i is None:
        return pd.DataFrame()
    a = pd.read_csv(io.StringIO("\n".join(t[i:])), sep="\t", dtype=str, quoting=3)
    a.columns = [c.strip() for c in a.columns]
    return a.apply(lambda s: s.str.strip())


def snd_caracteristicas():
    partes = []
    for exc in ("False", "True"):
        f = DADOS / "snd" / f"caracteristicas_{exc}.txt"
        if velho(f, 24):
            try:
                f.write_bytes(get(SND + f"emissoesdedebentures/caracteristicas_e.asp?tip_deb=publicas&op_exc={exc}&ativo=", timeout=300).content)
            except requests.RequestException:
                pass
        if f.exists():
            partes.append(_tabela_snd(f.read_bytes(), "Codigo do Ativo"))
    a = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()
    if not len(a):
        return pd.DataFrame(columns=["empresa", "cnpj", "venc", "indice", "pct", "juros"])
    dt = lambda c: pd.to_datetime(a[c], format="%d/%m/%Y", errors="coerce")
    return pd.DataFrame({"codigo": a["Codigo do Ativo"], "empresa": a["Empresa"], "cnpj": a["CNPJ"],
                         "venc": dt("Data de Saida / Novo Vencimento").fillna(dt("Data de Vencimento")), "indice": a["indice"],
                         "pct": pd.to_numeric(a["Percentual Multiplicador/Rentabilidade"], errors="coerce"),
                         "juros": pd.to_numeric(a["Juros Criterio Novo - Taxa"].str.replace(",", "."), errors="coerce")}
                        ).drop_duplicates("codigo").set_index("codigo")


def snd_precos(cod):
    """Preços de negociação do SND: por dia, quantidade, PU médio e % do PU da curva (3 anos)."""
    f = DADOS / "snd" / f"{cod}.parquet"
    velho_ = pd.read_parquet(f) if f.exists() else None
    if velho_ is not None and not velho(f, 20):
        return velho_
    ini = velho_.data.max() - pd.Timedelta(days=10) if velho_ is not None and len(velho_) else HOJE - pd.DateOffset(years=3)
    u = (SND + "mercadosecundario/precosdenegociacao_r.asp?op_exc=False&emissor=&isin="
         f"&ativo={cod}&dt_ini={ini:%d/%m/%Y}&dt_fim={HOJE:%d/%m/%Y}&Submit.x=1")
    vazio = pd.DataFrame(columns=["data", "qtd", "pmed", "pct"])
    try:
        t = get(u, timeout=120).content.decode("latin1")
        tb = [x for x in pd.read_html(io.StringIO(t), decimal=",", thousands=".") if x.shape[1] >= 19]
    except Exception:
        return velho_ if velho_ is not None else vazio
    novo = vazio
    if tb:
        x = tb[0][tb[0][0].astype(str).str.match(r"^\d{2}/\d{2}/\d{4}$")]
        n_ = lambda c: pd.to_numeric(x[c], errors="coerce")
        novo = pd.DataFrame({"data": pd.to_datetime(x[0], dayfirst=True), "qtd": n_(8), "pmed": n_(14), "pct": n_(18)})
    d = pd.concat([velho_, novo]) if velho_ is not None else novo
    d = d.drop_duplicates("data", keep="last").sort_values("data").reset_index(drop=True)
    d.to_parquet(f)
    return d


# ================================================================== ANBIMA
def _le_anbima(f):
    linhas = f.read_bytes().decode("latin1").splitlines()
    i = next(j for j, l in enumerate(linhas) if l.startswith("Código@"))
    a = pd.read_csv(io.StringIO("\n".join(linhas[i:])), sep="@", dtype=str)
    a.columns = ["codigo", "nome", "venc", "indice", "taxa_compra", "taxa_venda", "taxa_ind", "desvio", "int_min",
                 "int_max", "pu", "pct_par", "duration_du", "pct_reune", "ref_ntnb"][:a.shape[1]]
    for c in ["taxa_ind", "pu", "pct_par", "duration_du"]:
        a[c] = pd.to_numeric(a[c].str.replace(",", ".", regex=False), errors="coerce")
    a["codigo"] = a.codigo.str.strip().str.upper()
    a["duration_anos"] = a.duration_du / 252
    return a.assign(data=pd.Timestamp(f"20{f.stem[-6:-4]}-{f.stem[-4:-2]}-{f.stem[-2:]}"))


def anbima_diario():
    """Arquivo diário de debêntures da ANBIMA (o site guarda ~5 semanas; cada execução guarda o que estiver lá)."""
    for dia in pd.bdate_range(HOJE - pd.Timedelta(days=35), HOJE):
        f = DADOS / "anbima" / f"db{dia:%y%m%d}.txt"
        if f.exists():
            continue
        try:
            r = S.get(f"https://www.anbima.com.br/informacoes/merc-sec-debentures/arqs/db{dia:%y%m%d}.txt", timeout=60)
        except requests.RequestException:
            continue
        if r.status_code == 200 and "Código@".encode("latin1") in r.content:
            f.write_bytes(r.content)
    fs = sorted((DADOS / "anbima").glob("db*.txt"))
    return pd.concat([_le_anbima(f) for f in fs], ignore_index=True) if fs else pd.DataFrame(
        columns=["codigo", "data", "taxa_ind", "pu", "pct_par", "duration_anos", "indice"])


GANCHO = r"""
window.__cap=[];
const of=window.fetch; window.fetch=async(...a)=>{const r=await of(...a);try{const u=(a[0].url||a[0])+"";if(/grafico-pu/.test(u)){r.clone().text().then(t=>window.__cap.push({u,t}))}}catch(e){}return r};
const ox=XMLHttpRequest.prototype.open,os=XMLHttpRequest.prototype.send;
XMLHttpRequest.prototype.open=function(m,u){this.__u=u;return ox.apply(this,arguments)};
XMLHttpRequest.prototype.send=function(){this.addEventListener("load",()=>{if(/grafico-pu/.test(this.__u))window.__cap.push({u:this.__u,t:this.responseText})});return os.apply(this,arguments)};
"""
AGENDA_RE = re.compile(r"(\d{2}/\d{2}/\d{4}) (\d{2}/\d{2}/\d{4}) ([^\d/]+?) ([\d.,]+ %|-) (R\$ [\d.,]+|-)(?: -)?(?= \d{2}/\d{2}/\d{4} \d{2}/\d{2}/\d{4}|$)")


def anbima_paginas(cods):
    """Página pública de cada ativo na ANBIMA Data (Edge headless, como um usuário): histórico diário de PU par/PU indicativo e
    eventos (gráfico "Total"), topo da página (taxa, duration, % PU par) e agenda de eventos. Cache diário."""
    tipo = lambda c: "certificado-de-recebiveis" if c.startswith(("CRA", "CRI")) or re.match(r"^\d{2}[A-L]\d", c) else "debentures"
    fr, fa = DADOS / "anbima_hist" / f"resumo_{HOJE:%Y%m%d}.json", DADOS / "anbima_hist" / f"agenda_{HOJE:%Y%m%d}.json"
    resumo = json.loads(fr.read_text(encoding="utf-8")) if fr.exists() else {}
    agenda = json.loads(fa.read_text(encoding="utf-8")) if fa.exists() else {}
    hist = {c: pd.read_parquet(DADOS / "anbima_hist" / f"{c}.parquet") for c in cods if (DADOS / "anbima_hist" / f"{c}.parquet").exists()}
    arq = lambda c: DADOS / "anbima_hist" / f"{c}.parquet"
    falta = [c for c in cods if c not in resumo or c not in agenda or (arq(c).exists() and velho(arq(c)))]      # sem histórico: 1 tentativa por dia
    if falta:
        try:
            from selenium import webdriver
            o = webdriver.EdgeOptions()
            o.add_argument("--headless=new")
            o.add_argument("--window-size=1400,1000")
            d = webdriver.Edge(options=o)
        except Exception as e:
            log(f"ANBIMA Data: navegador indisponível ({type(e).__name__}); usando o cache")
            falta, d = [], None
        n_ = lambda v: float(v.replace(".", "").replace(",", ".")) if v else None
        try:
            for k, c in enumerate(falta):
                try:
                    d.get(f"https://data.anbima.com.br/{tipo(c)}/{c}/precos")
                    for _ in range(24):
                        time.sleep(0.5)
                        if d.execute_script("return [...document.querySelectorAll('li')].some(e=>e.innerText.trim()=='Total')"):
                            break
                    t = d.find_element("tag name", "body").text
                    g = lambda p: (re.search(p, t) or [None, None])[1]
                    dr = g(r"Data de refer.ncia\n(\d{2}/\d{2}/\d{4})\nPU [Ii]ndicativo")
                    resumo[c] = {"data": pd.to_datetime(dr, dayfirst=True).strftime("%Y-%m-%d") if dr else None,
                                 "taxa": n_(g(r"Taxa indicativa\n([\d.,]+) ?%")), "duration": (n_(g(r"Duration\n([\d.,]+) dias")) or 0) / 252 or None,
                                 "pct": n_(g(r"% PU par\n([\d.,]+) ?%"))}
                    if velho(DADOS / "anbima_hist" / f"{c}.parquet"):
                        d.execute_script(GANCHO)
                        d.execute_script("[...document.querySelectorAll('li')].find(e=>e.innerText.trim()=='Total')?.click()")
                        cap = None
                        for _ in range(30):
                            time.sleep(0.5)
                            cap = d.execute_script("return window.__cap.length ? window.__cap[window.__cap.length-1].t : null")
                            if cap:
                                break
                        if cap:
                            x = pd.DataFrame([{"data": p["data"], "pu_par": p.get("valor_pu_historico"), "pu_ind": p.get("valor_pu_indicativo"),
                                               "evento": "; ".join(e.get("tipo", "") for e in p.get("eventos") or []),
                                               "valor_evento": sum(float(e.get("valor") or 0) for e in p.get("eventos") or [])}
                                              for p in json.loads(cap).get("pus", [])])
                            if len(x):
                                x["data"] = pd.to_datetime(x.data)
                                x["pu_par"], x["pu_ind"] = pd.to_numeric(x.pu_par, errors="coerce"), pd.to_numeric(x.pu_ind, errors="coerce")
                                x.to_parquet(DADOS / "anbima_hist" / f"{c}.parquet")
                                hist[c] = x
                    if c not in agenda:
                        d.get(f"https://data.anbima.com.br/{tipo(c)}/{c}/agenda")
                        time.sleep(4)
                        ln, ev = d.find_element("tag name", "body").text.splitlines(), []
                        for i, l in enumerate(ln[:-1]):
                            st = ln[i + 1].strip()
                            for m in AGENDA_RE.finditer(l.strip()):
                                ev.append({"data": m.group(1), "evento": m.group(3).strip(), "pct": m.group(4),
                                           "status": "LIQUIDADO" if m.group(5) != "-" else (st if st in ("PREVISTO", "LIQUIDADO") else "PREVISTO")})
                        agenda[c] = ev
                    if (k + 1) % 20 == 0:
                        log(f"   ANBIMA Data {k + 1}/{len(falta)}")
                except Exception as e:
                    log(f"   ANBIMA Data {c}: {type(e).__name__}")
        finally:
            if d is not None:
                d.quit()
        fr.write_text(json.dumps(resumo), encoding="utf-8")
        fa.write_text(json.dumps(agenda, ensure_ascii=False), encoding="utf-8")
    return hist, resumo, agenda


# ================================================================== CRI/CRA: securitizadoras (CVM)
def securitizadoras():
    """Série CETIP -> devedor (CNPJ), vencimento, taxa, ISIN (informe mensal das securitizadoras, últimos 2 anos)."""
    linhas, devs = [], []
    for t in ("CRA", "CRI"):
        for a in (HOJE.year - 1, HOJE.year):
            f = DADOS / "securit" / f"inf_mensal_{t.lower()}_{a}.zip"
            if velho(f, 24 * 7) and (a == HOJE.year or not f.exists()):
                try:
                    f.write_bytes(get(f"https://dados.cvm.gov.br/dados/SECURIT/DOC/INF_MENSAL_{t}/DADOS/inf_mensal_{t.lower()}_{a}.zip", timeout=300).content)
                except requests.RequestException:
                    pass
            if not f.exists():
                continue
            z = zipfile.ZipFile(f)
            le = lambda k: next((pd.read_csv(z.open(n), sep=";", encoding="latin1", dtype=str, on_bad_lines="skip") for n in z.namelist() if k in n), pd.DataFrame())
            c, dv = le("classe"), le("cedente_devedor")
            if len(c):
                linhas.append(c.assign(tipo=t))
            if len(dv):
                devs.append(dv)
    if not linhas:
        return {}
    c = pd.concat(linhas, ignore_index=True).sort_values("Data_Referencia").drop_duplicates("Codigo_CETIP", keep="last")
    dv = pd.concat(devs, ignore_index=True) if devs else pd.DataFrame(columns=["Codigo_Identificacao_Certificado", "CNPJ", "Tipo"])
    dv = dv[dv.Tipo.fillna("").str.contains("Devedor")].drop_duplicates("Codigo_Identificacao_Certificado").set_index("Codigo_Identificacao_Certificado").CNPJ
    return {str(r.Codigo_CETIP).strip().upper(): {"tipo": r.tipo, "venc": str(r.Data_Vencimento)[:10], "taxa": r.Taxa_Juros, "isin": r.Codigo_ISIN,
                                                    "cnpj": dv.get(r.Codigo_Identificacao_Certificado)} for r in c.itertuples()}


# ================================================================== CVM companhias abertas
CVM = "https://dados.cvm.gov.br/dados/CIA_ABERTA/"


def cvm_zip(doc, ano):
    f = DADOS / "cvm" / f"{doc}_cia_aberta_{ano}.zip"
    if not f.exists() or (ano >= HOJE.year - (1 if doc == "dfp" else 0) and velho(f, 24 * 3)):
        try:
            f.write_bytes(get(f"{CVM}DOC/{doc.upper()}/DADOS/{doc}_cia_aberta_{ano}.zip", timeout=600).content)
        except requests.RequestException:
            pass
    return zipfile.ZipFile(f) if f.exists() else None


def companhias():
    """Por raiz de CNPJ: nome, setor, situação e códigos de negociação em bolsa (FCA)."""
    f = DADOS / "cvm" / "cad_cia_aberta.csv"
    if velho(f, 24 * 7):
        try:
            f.write_bytes(get(CVM + "CAD/DADOS/cad_cia_aberta.csv", timeout=300).content)
        except requests.RequestException:
            pass
    cad = pd.read_csv(f, sep=";", encoding="latin1", dtype=str) if f.exists() else pd.DataFrame(columns=["CNPJ_CIA", "DENOM_SOCIAL", "SETOR_ATIV", "SIT"])
    cad["raiz"] = cad.CNPJ_CIA.map(raiz)
    cad = cad.sort_values("SIT").drop_duplicates("raiz")
    tk = []
    for a in (HOJE.year, HOJE.year - 1):
        z = cvm_zip("fca", a)
        if z is None:
            continue
        v = pd.read_csv(z.open(f"fca_cia_aberta_valor_mobiliario_{a}.csv"), sep=";", encoding="latin1", dtype=str)
        tk.append(v[v.Mercado.eq("Bolsa") & v.Codigo_Negociacao.notna() & v.Data_Fim_Negociacao.isna()])
    tk = pd.concat(tk) if tk else pd.DataFrame(columns=["CNPJ_Companhia", "Codigo_Negociacao"])
    tk["raiz"] = tk.CNPJ_Companhia.map(raiz)
    tks = tk.groupby("raiz").Codigo_Negociacao.apply(lambda s: sorted({x.strip().upper() for x in s if re.match(r"^[A-Z]{4}\d{1,2}$", x.strip().upper())}))
    out = cad.set_index("raiz")[["DENOM_SOCIAL", "SETOR_ATIV", "SIT"]].rename(columns={"DENOM_SOCIAL": "nome", "SETOR_ATIV": "setor", "SIT": "sit"})
    out["tickers"] = tks.reindex(out.index)
    return out


CONTAS = {"DRE": ["3.01", "3.05", "3.06", "3.06.02", "3.11"], "BPA": ["1.01.01", "1.01.02"], "BPP": ["2.01.04", "2.02.01"]}


def _extrai(z, doc, ano):
    out = []
    for dem in ("DRE", "BPA", "BPP", "DFC_MI"):
        for k in ("con", "ind"):
            n = f"{doc}_cia_aberta_{dem}_{k}_{ano}.csv"
            if n not in z.namelist():
                continue
            d = pd.read_csv(z.open(n), sep=";", encoding="latin1", dtype=str,
                            usecols=["CNPJ_CIA", "DT_REFER", "VERSAO", "ESCALA_MOEDA", "ORDEM_EXERC", "DT_INI_EXERC", "DT_FIM_EXERC",
                                     "CD_CONTA", "DS_CONTA", "VL_CONTA"] if dem != "BPA" and dem != "BPP" else
                            ["CNPJ_CIA", "DT_REFER", "VERSAO", "ESCALA_MOEDA", "ORDEM_EXERC", "DT_FIM_EXERC", "CD_CONTA", "DS_CONTA", "VL_CONTA"])
            d = d[d.ORDEM_EXERC.str.contains("LTIMO", na=False) & ~d.ORDEM_EXERC.str.contains("PEN", na=False)]
            if dem == "DFC_MI":
                d = d[d.CD_CONTA.str.startswith("6.01.01") & d.DS_CONTA.str.contains("eprecia|mortiza|xaust", case=False, na=False)
                      & ~d.DS_CONTA.str.contains("custo de transa|custos de transa|ágio|agio|juros", case=False, na=False)]
                d = d.assign(CD_CONTA="DA")
            else:
                if dem == "DRE":
                    fin = d[d.CD_CONTA.eq("3.01")].DS_CONTA.str.contains("Intermedia|Seguros|Prêmios", case=False, na=False)
                    d = d[d.CD_CONTA.isin(CONTAS[dem])]
                    d.loc[d.CNPJ_CIA.isin(d[d.CD_CONTA.eq("3.01")][fin.values].CNPJ_CIA), "CD_CONTA"] = "FIN"
                else:
                    d = d[d.CD_CONTA.isin(CONTAS[dem])]
            d["v"] = pd.to_numeric(d.VL_CONTA, errors="coerce") * np.where(d.ESCALA_MOEDA.eq("MIL"), 1e3, 1)
            d["k"], d["dem"] = k, dem
            out.append(d.drop(columns=["VL_CONTA", "DS_CONTA", "ESCALA_MOEDA", "ORDEM_EXERC"]))
    x = pd.concat(out, ignore_index=True) if out else pd.DataFrame()
    if not len(x):
        return x
    if "DT_INI_EXERC" not in x:
        x["DT_INI_EXERC"] = None
    x = x.groupby(["CNPJ_CIA", "DT_REFER", "VERSAO", "DT_INI_EXERC", "DT_FIM_EXERC", "CD_CONTA", "dem", "k"], dropna=False).v.sum().reset_index()
    idx = pd.read_csv(z.open(f"{doc}_cia_aberta_{ano}.csv"), sep=";", encoding="latin1", dtype=str)[["CNPJ_CIA", "DT_REFER", "DT_RECEB"]]
    ent = idx.groupby(["CNPJ_CIA", "DT_REFER"]).DT_RECEB.min().rename("entrega")         # 1a entrega = quando o mercado soube
    return x.join(ent, on=["CNPJ_CIA", "DT_REFER"]).assign(doc=doc)


def fundamentos_brutos():
    partes = []
    for doc in ("itr", "dfp"):
        for a in range(HOJE.year - ANOS_CVM, HOJE.year + 1):
            f = DADOS / "cvm" / f"fund_{doc}_{a}.parquet"
            zf = DADOS / "cvm" / f"{doc}_cia_aberta_{a}.zip"
            atual = a >= HOJE.year - (1 if doc == "dfp" else 0)
            if not f.exists() or (atual and velho(f, 24 * 3)):
                z = cvm_zip(doc, a)
                if z is None:
                    continue
                log(f"   CVM {doc.upper()} {a}: extraindo contas")
                _extrai(z, doc, a).to_parquet(f)
                z.close()
                if not atual:
                    zf.unlink(missing_ok=True)                    # guarda só as contas (pequeno)
            if f.exists():
                partes.append(pd.read_parquet(f))
    x = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()
    if len(x):
        x["raiz"] = x.CNPJ_CIA.map(raiz)
    return x


def trimestres(F):
    """Uma companhia: série trimestral (receita, EBITDA, dívida líquida, alavancagem LTM...) com a data de entrega à CVM."""
    if not len(F):
        return pd.DataFrame()
    F = F[F.k.eq("con")] if F.k.eq("con").any() else F
    F = F.sort_values("VERSAO").drop_duplicates(["DT_REFER", "DT_INI_EXERC", "DT_FIM_EXERC", "CD_CONTA", "dem", "doc"], keep="last")
    fin = F.CD_CONTA.eq("FIN").any()
    fl = F[F.dem.isin(["DRE", "DFC_MI"]) & F.DT_INI_EXERC.notna()].copy()
    fl["ini"], fl["fim"] = pd.to_datetime(fl.DT_INI_EXERC), pd.to_datetime(fl.DT_FIM_EXERC)
    fl = fl[(fl.fim - fl.ini).dt.days > 80]
    # acumulado no exercício (YTD): a linha que começa no início do exercício; trimestre = YTD − YTD anterior do mesmo exercício
    ini_ex = fl.groupby(["CD_CONTA", "fim"]).ini.transform("min")
    ytd = fl[fl.ini.eq(ini_ex)].sort_values("fim").drop_duplicates(["CD_CONTA", "fim"], keep="last")
    ytd = ytd.sort_values(["CD_CONTA", "ini", "fim"])
    ytd["tri"] = ytd.v - ytd.groupby(["CD_CONTA", "ini"]).v.shift().fillna(0)
    tres = fl[((fl.fim - fl.ini).dt.days < 100)].sort_values("fim").drop_duplicates(["CD_CONTA", "fim"], keep="last")   # linha de 3 meses do ITR
    T = ytd.pivot_table(index="fim", columns="CD_CONTA", values="tri", aggfunc="last")
    T3 = tres.pivot_table(index="fim", columns="CD_CONTA", values="v", aggfunc="last")
    T = T3.combine_first(T)
    for c in [c for c in ("3.01", "3.05") if c in T]:                  # trimestre fora da linha dos vizinhos (erro de envio): fica sem número
        m = T[c].abs().rolling(5, center=True, min_periods=3).median()
        T.loc[(T[c].abs() < 0.4 * m) | (T[c].abs() > 2.5 * m), c] = np.nan
    st = F[F.dem.isin(["BPA", "BPP"])].assign(fim=lambda d: pd.to_datetime(d.DT_FIM_EXERC))
    B = st.pivot_table(index="fim", columns="CD_CONTA", values="v", aggfunc="last")
    E = F.assign(fim=pd.to_datetime(F.DT_REFER)).groupby("fim").entrega.min()
    q = T.join(B, how="outer").join(E, how="left").sort_index()
    q = q[q.index.is_month_end | (q.index.day >= 28)]
    g = lambda c: q[c] if c in q else pd.Series(np.nan, index=q.index)
    out = pd.DataFrame({"receita": g("3.01"), "ebit": g("3.05"), "da": g("DA").abs(), "desp_fin": g("3.06.02").abs(), "lucro": g("3.11"),
                        "divida": g("2.01.04").fillna(0) + g("2.02.01").fillna(0), "caixa": g("1.01.01").fillna(0) + g("1.01.02").fillna(0),
                        "entrega": q.entrega}, index=q.index)
    out["ebitda"] = out.ebit + out.da.fillna(0)
    out["dl"] = out.divida - out.caixa
    ltm = lambda c: out[c].rolling(4, min_periods=4).sum()           # 12 meses: exige os 4 trimestres daquela conta
    out["ebitda_12m"], out["receita_12m"], out["desp_fin_12m"] = ltm("ebitda"), ltm("receita"), ltm("desp_fin")
    out["dl_ebitda"] = (out.dl / out.ebitda_12m).where(out.ebitda_12m > 0)
    out["cobertura"] = (out.ebitda_12m / out.desp_fin_12m).where(out.desp_fin_12m > 0)
    out["margem"] = (out.ebitda_12m / out.receita_12m * 100).where(out.receita_12m > 0)
    out["financeira"] = fin
    return out


def ipe(raizes, anos=3):
    partes = []
    for a in range(HOJE.year - anos + 1, HOJE.year + 1):
        f = DADOS / "cvm" / f"ipe_cia_aberta_{a}.zip"
        if not f.exists() or (a == HOJE.year and velho(f, 6)):
            try:
                f.write_bytes(get(f"{CVM}DOC/IPE/DADOS/ipe_cia_aberta_{a}.zip", timeout=300).content)
            except requests.RequestException:
                pass
        if f.exists():
            z = zipfile.ZipFile(f)
            partes.append(pd.read_csv(z.open(z.namelist()[0]), sep=";", encoding="latin1", dtype=str))
    d = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame(columns=["CNPJ_Companhia", "Categoria", "Data_Entrega"])
    d["raiz"] = d.CNPJ_Companhia.map(raiz)
    d = d[d.raiz.isin(raizes)].copy()
    d["data"] = pd.to_datetime(d.Data_Entrega, errors="coerce").dt.normalize()
    return d.sort_values("data")


# ================================================================== mercado: Yahoo, índices, BCB, Tesouro, IBGE, notícias
def yahoo(tk, rng="5y"):
    """Fechamento ajustado diário + último preço. Cache: histórico 1x por dia; cotação a cada execução."""
    f = DADOS / "yahoo" / f"{re.sub(r'[^A-Za-z0-9]', '_', tk)}.json"
    if velho(f, 0.25):
        try:
            j = get(f"https://query1.finance.yahoo.com/v8/finance/chart/{tk}", params={"range": rng, "interval": "1d"}, timeout=30, tent=2).json()
            f.write_text(json.dumps(j["chart"]["result"][0]), encoding="utf-8")
        except Exception:
            pass
    if not f.exists():
        return None
    r = json.loads(f.read_text(encoding="utf-8"))
    q = r.get("indicators", {})
    c = (q.get("adjclose") or [{}])[0].get("adjclose") or q["quote"][0]["close"]
    s = pd.Series(c, index=pd.to_datetime(r["timestamp"], unit="s").normalize(), dtype=float).dropna()
    s = s[~s.index.duplicated(keep="last")]
    vol = pd.Series(q["quote"][0].get("volume") or [], dtype=float)
    cl = pd.Series(q["quote"][0].get("close") or [], dtype=float)
    giro = float((vol * cl).tail(60).mean()) if len(vol) == len(cl) and len(vol) else 0.0
    m = r.get("meta", {})
    return {"serie": s, "preco": m.get("regularMarketPrice"), "hora": m.get("regularMarketTime"), "giro": giro,
            "ant": float(s.iloc[-2]) if len(s) > 1 else None}


def indices_credito():
    """Índices de crédito e renda fixa (SmartWealth): fechamento diário (desde 04/2025) e retorno mensal (desde o início)."""
    out = {}
    for nome, slug in INDICES.items():
        f = DADOS / "indices" / f"{slug.split('/')[-1]}.json"
        if velho(f, 12):
            try:
                t = get(f"https://smartwealth.com.br/indices/{slug}", timeout=60, tent=2).text.replace('\\"', '"')
                d = re.findall(r'"timestamp":"(\d{4}-\d{2}-\d{2})T[^"]*","close":([0-9.eE-]+)', t)
                m = re.findall(r'"date":"(\d{4}-\d{2}-\d{2})","month_return":([0-9.eE-]+)', t)
                if d:
                    f.write_text(json.dumps({"d": d, "m": m}), encoding="utf-8")
            except Exception:
                pass
        if f.exists():
            j = json.loads(f.read_text(encoding="utf-8"))
            s = pd.Series({pd.Timestamp(a): float(b) for a, b in j["d"]}).sort_index()
            s = s[(np.log(s / s.rolling(31, center=True, min_periods=5).median())).abs() < 0.04]    # ponto isolado fora da curva (erro da fonte)
            mm = pd.Series({pd.Timestamp(a): float(b) for a, b in j["m"]}).sort_index()
            out[nome] = {"serie": s[~s.index.duplicated()], "mensal": mm[~mm.index.duplicated()]}
    return out


def sgs(cod, desde):
    f = DADOS / f"sgs{cod}.json"
    if velho(f, 6):
        try:
            f.write_text(get(f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{cod}/dados",
                             params={"formato": "json", "dataInicial": f"{desde:%d/%m/%Y}"}, timeout=60).text)
        except requests.RequestException:
            pass
    if not f.exists():
        return pd.Series(dtype=float)
    d = pd.DataFrame(json.loads(f.read_text()))
    return pd.Series(d.valor.astype(float).values, index=pd.to_datetime(d.data, dayfirst=True)).sort_index()


def focus():
    out = {}
    for ind in ("IPCA", "Selic"):
        try:
            j = get("https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoAnuais?"
                    f"$filter=Indicador%20eq%20'{ind}'&$orderby=Data%20desc&$top=20&$format=json", timeout=60).json()["value"]
            j = [x for x in j if x.get("baseCalculo") == 0]
            ult = max(x["Data"] for x in j)
            out[ind] = {"data": ult, "anos": {x["DataReferencia"]: x["Mediana"] for x in j if x["Data"] == ult}}
        except Exception:
            pass
    return out


def tesouro():
    """Taxas do Tesouro Direto por dia: prefixado ~3 anos e IPCA+ ~5 anos (para os fatores de juros)."""
    f = DADOS / "tesouro.csv"
    if velho(f, 24):
        try:
            b = get("https://www.tesourotransparente.gov.br/ckan/dataset/df56aa42-484a-4a59-8184-7676580c81e3/resource/"
                    "796d2059-14e9-44e3-80c9-2d9e30b405c1/download/PrecoTaxaTesouroDireto.csv", timeout=300, tent=2).content
            if b[:20].count(b";") or b"Tipo Titulo" in b[:200]:
                f.write_bytes(b)
        except requests.RequestException:
            pass
    if not f.exists():
        return pd.DataFrame()
    d = pd.read_csv(f, sep=";", decimal=",")
    d = d[d["Tipo Titulo"].isin(["Tesouro Prefixado", "Tesouro IPCA+"])].copy()
    d["data"] = pd.to_datetime(d["Data Base"], dayfirst=True)
    d = d[d.data >= HOJE - pd.DateOffset(years=4)]
    d["anos"] = (pd.to_datetime(d["Data Vencimento"], dayfirst=True) - d.data).dt.days / 365.25
    d["taxa"] = d[["Taxa Compra Manha", "Taxa Venda Manha"]].replace(0, np.nan).mean(axis=1)
    out = {}
    for tipo, alvo, nome in (("Tesouro Prefixado", 3, "Pré 3a"), ("Tesouro IPCA+", 5, "IPCA+ 5a")):
        x = d[d["Tipo Titulo"].eq(tipo)].dropna(subset=["taxa"])
        x = x.assign(dist=(x.anos - alvo).abs()).sort_values("dist").drop_duplicates("data")
        out[nome] = x.set_index("data").taxa.sort_index()
    return pd.DataFrame(out)


def ipca_datas():
    try:
        j = get("https://servicodados.ibge.gov.br/api/v3/calendario/", params={"qtd": 400}, timeout=60).json()["items"]
        return sorted({pd.to_datetime(i["data_divulgacao"], dayfirst=True).normalize() for i in j
                       if i["titulo"].strip() in ("Índice Nacional de Preços ao Consumidor Amplo", "Índice Nacional de Preços ao Consumidor Amplo 15")
                       and pd.to_datetime(i["data_divulgacao"], dayfirst=True) >= HOJE})
    except Exception:
        return []


def noticias(nome, dias=14):
    q = " ".join([w for w in re.sub(r"[^\w\s]", " ", str(nome).upper()).split() if w not in PARADAS][:3]).title()
    if not q:
        return q, []
    try:
        r = get("https://news.google.com/rss/search", params={"q": f'"{q}" (ações OR debêntures OR resultado OR dívida OR lucro OR rating OR receita OR prejuízo OR emissão) when:{dias}d', "hl": "pt-BR", "gl": "BR", "ceid": "BR:pt-419"}, timeout=20, tent=2)
        it = ET.fromstring(r.content).findall(".//item")
    except Exception:
        return q, []
    chave = q.split()[0].lower()
    it = [i for i in it if chave in (i.findtext("title") or "").lower()]            # o nome tem de estar no título
    return q, [{"titulo": i.findtext("title") or "", "link": i.findtext("link"), "data": pd.to_datetime(i.findtext("pubDate")).strftime("%Y-%m-%d"),
                "fonte": (i.find("source").text if i.find("source") is not None else ""),
                "alerta": [k for k in ALERTAS if k in (i.findtext("title") or "").lower()]} for i in it[:8]]


# ================================================================== séries do ativo: oficial, taxa, PnL
def serie_oficial(h):
    """Série oficial limpa: dias com PU indicativo preciso (sem arredondamento > 0,05%), sem pico isolado de % PU par (> 8 p.p. da
    mediana de 11 dias); E = pagamentos até o preço seguinte (incorporação de juros não é pagamento; queda do PU par sem evento = pagamento)."""
    x = h.sort_values("data").reset_index(drop=True)
    ev = x.evento.fillna("").str.upper()
    x["pago"] = x.valor_evento.fillna(0).where(~ev.str.contains("INCORPORA") | ev.str.contains(";"), 0)
    sem_par = not (x.pu_par > 0).any()                         # ANBIMA sem PU par (alguns papéis): só retorno, sem decompor
    v = x[x.pu_ind.notna() & ((x.pu_par > 0) | sem_par)].copy()
    casas = v.pu_ind.map(lambda u: next((k for k in range(7) if abs(round(u, k) - u) < 1e-9), 7))
    v = v[10.0 ** -casas.values / v.pu_ind.values < 5e-4]
    if not sem_par:
        pct = v.pu_ind / v.pu_par * 100
        v = v[(pct - pct.rolling(11, center=True, min_periods=3).median()).abs() <= 8]
    if len(v) < 2:
        return v.assign(E=0.0)
    grupo = np.searchsorted(v.index.values, x.index.values, side="left")
    E = pd.Series(x.pago.values).groupby(grupo).sum()
    v["E"] = [float(E.get(k, 0.0)) for k in range(len(v))]
    queda = (v.pu_par.shift() - v.pu_par).fillna(0)
    v["E"] = np.maximum(v.E.values, queda.where(queda > 0.002 * v.pu_par.shift(), 0).fillna(0).values)
    v.iloc[0, v.columns.get_loc("E")] = 0.0
    return v.reset_index(drop=True)


def tipo_idx(indexador):
    s = str(indexador or "").upper().strip()
    if not s or s in ("NONE", "NAN", "SEM INDICE"):
        return "?"
    return "%DI" if "%" in s and "DI" in s else "DI+" if "DI" in s else "IPCA" if "IPCA" in s else "IGPM" if "IGP" in s else "PRE"


def indexadores(cods):
    """Indexador de cada código: SND (debêntures, inclusive vencidas) e, na falta, o cadastro da B3."""
    snd, cad = snd_caracteristicas(), cadastro_b3().drop_duplicates("cod").set_index("cod")
    return {c: tipo_idx(snd.indice.get(c) if c in snd.index else cad.idx.get(c)) for c in cods}


def taxa_e_pnl(v, of, venc, cupom, idx, dur0, cdi):
    """v: série oficial limpa; of: Series de taxa indicativa oficial (ANBIMA, % a.a.) por data; dur0: duration (anos) na última data.
    Taxa antes da 1ª oficial = implícita no PU oficial (de trás para frente). PnL diário em % do PU: carrego (PU par + (taxa − cupom))
    − D_mod × Δtaxa + resíduo. Para %DI as taxas viram spread equivalente sobre o CDI."""
    n = len(v)
    if n < 3:
        return None
    venc = pd.Timestamp(venc)
    dts = pd.DatetimeIndex(v.data)
    T = np.maximum((venc - dts).days / 365.25, 0.02)
    D = dur0 * T / T[-1] if dur0 else T * 0.8
    du = np.array([0] + [dias_uteis(a, b) for a, b in zip(dts[:-1], dts[1:])])
    cdi_aa = (cdi.reindex(dts, method="ffill").fillna(cdi.iloc[-1] if len(cdi) else 0.04).values / 100 + 1) ** 252 - 1 if len(cdi) else np.full(n, 0.12)
    conv = (lambda y: cdi_aa * (y / 100 - 1)) if idx == "%DI" else (lambda y: y / 100)
    ri = np.r_[np.nan, (v.pu_ind.values[1:] + v.E.values[1:]) / v.pu_ind.values[:-1] - 1]
    rp = np.r_[np.nan, (v.pu_par.values[1:] + v.E.values[1:]) / v.pu_par.values[:-1] - 1]
    y_of = of.reindex(dts).values if of is not None and len(of) else np.full(n, np.nan)
    y = np.full(n, np.nan)
    fonte = np.zeros(n, int)                                   # 0 oficial, 1 implícita no PU oficial
    ok = np.where(np.isfinite(y_of))[0]
    c = conv(np.full(n, cupom)) if cupom is not None and np.isfinite(cupom) else None
    if len(ok):
        y[ok] = conv(y_of)[ok] if idx == "%DI" else y_of[ok] / 100
        k0 = ok[0]
        if c is None:
            c = np.full(n, y[k0])
        for k in range(k0, 0, -1):                             # para trás: taxa implícita no PU oficial
            dm = D[k] / (1 + y[k]) if idx in ("IPCA", "PRE", "IGPM") else D[k]
            ex = np.log(1 + ri[k]) - np.log(1 + rp[k]) - np.log((1 + y[k]) / (1 + c[k])) * du[k] / 252
            y[k - 1] = y[k] + ex / max(dm, 0.02)
            fonte[k - 1] = 1
        for k in range(k0 + 1, n):                             # depois da 1ª oficial: dias sem arquivo -> implícita a partir do anterior
            if not np.isfinite(y[k]):
                dm = D[k - 1] / (1 + y[k - 1]) if idx in ("IPCA", "PRE", "IGPM") else D[k - 1]
                ex = np.log(1 + ri[k]) - np.log(1 + rp[k]) - np.log((1 + y[k - 1]) / (1 + c[k])) * du[k] / 252
                y[k] = y[k - 1] - ex / max(dm, 0.02)
                fonte[k] = 1
    else:
        return {"ret": ri, "taxa": np.full(n, np.nan), "fonte": fonte, "car": np.full(n, np.nan), "mtm": np.full(n, np.nan),
                "res": np.full(n, np.nan), "D": D}
    dm_ = np.r_[np.nan, np.where(np.isin(idx, ["IPCA", "PRE", "IGPM"]), D[:-1] / (1 + y[:-1]), D[:-1])]
    car = np.r_[np.nan, (1 + rp[1:]) * ((1 + y[:-1]) / (1 + c[1:])) ** (du[1:] / 252) - 1]
    mtm = -dm_ * np.r_[np.nan, np.diff(y)]
    res = ri - car - mtm
    res[fonte == 1] = np.nan                                   # taxa implícita: resíduo zero por construção, não é medida
    taxa_pub = np.where(fonte == 0, y_of, (y * 100) if idx != "%DI" else (y / cdi_aa + 1) * 100)
    return {"ret": ri, "taxa": taxa_pub, "fonte": fonte, "car": car, "mtm": mtm, "res": res, "D": D}


def vol_semanal(datas, nivel, ruido_var=None):
    """Vol anual (%) com retornos semanais de negócios; o ruído de cada preço (medido contra a ANBIMA) sai da variância."""
    s = pd.Series(np.log(np.asarray(nivel, float)), index=pd.to_datetime(pd.Series(datas)).values).dropna().sort_index()
    w = s.groupby(pd.Grouper(freq="W-FRI")).last().dropna()
    if len(w) < 6:
        return None, max(len(w) - 1, 0)
    gap = np.diff(w.index.values).astype("timedelta64[D]").astype(int) / 7
    z = pd.Series(np.diff(w.values) / np.sqrt(gap), index=w.index[1:])[gap <= 4]
    if len(z) < 5:
        return None, len(z)
    mad = np.median(np.abs(z - z.median())) * 1.4826 + 1e-12
    z = z[np.abs(z - z.median()) < 6 * mad]
    v = z.var(ddof=1)
    if ruido_var is not None:
        v = max(v - 2 * ruido_var, 0.25 * v)
    elif len(z) >= 10:
        v = max(v + 2 * np.cov(z.values[1:], z.values[:-1])[0, 1], 0.25 * v)
    return float(np.sqrt(52 * v) * 100), len(z)


# ================================================================== identificação
def identifica(c):
    snd, cad, sec = snd_caracteristicas(), cadastro_b3().drop_duplicates("cod").set_index("cod"), securitizadoras()
    out = []
    for r in c.itertuples():
        x = {"codigo": r.codigo, "venc": r.venc, "perc_nav": r.perc_nav, "entrada": r.entrada, "tipo": None, "emissor": None,
             "cnpj": None, "indexador": None, "cupom": None, "lastro": None}
        k = r.codigo
        if k and k in cad.index:
            q = cad.loc[k]
            x.update(tipo=q.tipo, emissor=q.emissor, indexador=q.idx, cupom=num([q.taxa]).iloc[0])
        if k and k in snd.index:
            q = snd.loc[k]
            x.update(tipo=x["tipo"] or "DEB", emissor=q.empresa, cnpj=q.cnpj, indexador=q.indice,
                     cupom=float(q.juros) if q.juros == q.juros and q.juros else (float(q.pct) if q.pct == q.pct else x["cupom"]))
        if k and k in sec:
            q = sec[k]
            x.update(tipo=q["tipo"], cnpj=q["cnpj"] or x["cnpj"], indexador=x["indexador"] or q["taxa"], lastro="devedor do " + q["tipo"])
        if x["emissor"] and re.search(r"SECURITIZ", str(x["emissor"]).upper()) and x["tipo"] == "DEB":
            x["lastro"] = "debênture de securitizadora: o risco é o lastro, não a securitizadora"
        x["raiz"] = raiz(x["cnpj"])
        x["idx"] = tipo_idx(x["indexador"])
        out.append(x)
    return pd.DataFrame(out)


def escolhe_acao(raizes, cias):
    """Raiz -> ação: a própria (mais líquida entre as classes) ou, sem ação, a do grupo pelo 1º nome distintivo (marcado)."""
    listadas = cias[cias.tickers.map(lambda t: isinstance(t, list) and len(t) > 0)]
    nome1 = lambda s: next((w for w in re.sub(r"[^\w\s]", " ", str(s).upper()).split() if w not in PARADAS and len(w) >= 4), None)
    l1 = listadas.nome.map(nome1)
    out = {}
    for r in raizes:
        if not r or r not in cias.index:
            continue
        tks, via = cias.tickers.get(r), "própria"
        if not isinstance(tks, list) or not tks:
            w = nome1(cias.nome[r])
            cand = listadas[(l1 == w) & (listadas.index != r)] if w else listadas.iloc[:0]
            if not len(cand):
                continue
            cand = cand.assign(n=cand.nome.str.len()).sort_values("n")
            tks, via = cand.tickers.iloc[0], f"grupo ({cand.nome.iloc[0][:40]})"
        best = None
        for t in tks[:4]:
            y = yahoo(f"{t}.SA")
            if y and len(y["serie"]) > 120 and (best is None or y["giro"] > best[1]["giro"]):
                best = (t, y)
        if best:
            out[r] = {"ticker": best[0], "via": via, **best[1]}
    return out


# ================================================================== análises
def semanal(s):
    return s.groupby(pd.Grouper(freq="W-FRI")).last().dropna()


def acao_x_credito(taxa_d, preco):
    """Semanal: retorno da ação (%) × Δtaxa indicativa (bps). Lags de −4 a +4 semanas (+k: ação k semanas antes)."""
    t, p = semanal(taxa_d), semanal(preco)
    dt, rp = t.diff() * 100, np.log(p).diff() * 100
    j = pd.concat([rp.rename("a"), dt.rename("t")], axis=1).dropna()
    if len(j) < 30:
        return None
    lags = {}
    for k in range(-4, 5):
        r, n = corr(j.a.values[:len(j) - k] if k >= 0 else j.a.values[-k:], j.t.values[k:] if k >= 0 else j.t.values[:len(j) + k], rank=True)
        lo, hi = fisher(r, n)
        lags[k] = {"r": r, "n": n, "lo": lo, "hi": hi}
    b = float(np.cov(j.a, j.t)[0, 1] / j.a.var()) if j.a.var() > 0 else None
    return {"n": len(j), "beta": b, "lags": lags, "metodo": "Spearman", "datas": [d.strftime("%Y-%m-%d") for d in j.index], "a": j.a.round(2).tolist(),
            "t": j.t.round(1).tolist(), "_j": j}


def reacao(taxa_d, datas, janela=5):
    """Δtaxa (bps) do último dia antes de cada data até `janela` dias úteis depois; z = Δ / desvio de Δ de mesma janela fora de eventos."""
    s = taxa_d.dropna()
    if len(s) < 60:
        return [], None
    idx = s.index
    dd = (s.shift(-janela) - s).dropna() * 100
    out = []
    for d in datas:
        i0 = idx.searchsorted(d) - 1
        if i0 < 0 or i0 + janela >= len(idx):
            continue
        out.append((d, float((s.iloc[i0 + janela] - s.iloc[i0]) * 100)))
    longe = dd[[min(abs((x - d).days) for d in datas) > 15 for x in dd.index]] if datas else dd
    sd = float(longe.std()) if len(longe) > 20 else None
    return out, sd


def negocio_x_marcacao(neg, taxa_d):
    """Dia com negócio (≥ R$ 50 mil): gap = taxa negociada − ANBIMA do dia anterior; mv0 = ANBIMA do dia − anterior; mv1 = dia seguinte."""
    s = taxa_d.dropna()
    if len(s) < 30 or not len(neg):
        return []
    g = neg.assign(tv=neg.taxa * neg.vol).groupby("data").agg(tv=("tv", "sum"), vol=("vol", "sum"))
    vd = g.vol
    g = (g.tv / g.vol).dropna()
    out = []
    for d, tx in g.items():
        i = s.index.searchsorted(d)
        if i == 0 or i >= len(s) - 1 or s.index[i] != d:
            continue
        gap, mv0, mv1 = (tx - s.iloc[i - 1]) * 100, (s.iloc[i] - s.iloc[i - 1]) * 100, (s.iloc[i + 1] - s.iloc[i]) * 100
        if abs(gap) < 300:
            out.append((d, gap, mv0, mv1, (tx - s.iloc[i]) * 100, float(vd[d])))
    return out


def regressao(x, y):
    x, y = np.asarray(x, float), np.asarray(y, float)
    k = np.isfinite(x) & np.isfinite(y)
    x, y = x[k], y[k]
    if len(x) < 15 or x.var() == 0:
        return None
    b = float(np.cov(x, y)[0, 1] / x.var())
    a = float(y.mean() - b * x.mean())
    e = y - a - b * x
    se = float(np.sqrt(e.var(ddof=2) / (x.var() * len(x))))
    return {"b": b, "a": a, "t": b / se if se > 0 else None, "r2": float(1 - e.var() / y.var()) if y.var() > 0 else None, "n": int(len(x))}


def fora_da_amostra(x, y, datas, limiar):
    """Ajusta a inclinação nos 2/3 mais antigos; no terço final, quando |x| ≥ limiar, acerto do sinal de y previsto (b·x)."""
    o = np.argsort(np.asarray(datas))
    x, y = np.asarray(x, float)[o], np.asarray(y, float)[o]
    k = int(len(x) * 2 / 3)
    f = regressao(x[:k], y[:k])
    if f is None:
        return None
    xt, yt = x[k:], y[k:]
    m = (np.abs(xt) >= limiar) & np.isfinite(yt) & (yt != 0)
    if m.sum() < 8:
        return {"b_treino": f["b"], "n_teste": int(m.sum()), "acerto": None}
    ac = float((np.sign(f["b"] * xt[m]) == np.sign(yt[m])).mean() * 100)
    return {"b_treino": f["b"], "n_teste": int(m.sum()), "acerto": ac, "base": float(max((yt[m] > 0).mean(), (yt[m] < 0).mean()) * 100)}


# ================================================================== cálculo principal
def calcula(c):
    t0 = time.time()
    a = identifica(c)
    cods = [c for c in a.codigo.dropna()]
    log(f"carteira: {len(a)} linhas, {len(cods)} com código")
    an = anbima_diario()
    of_hist, resumo, agenda = anbima_paginas(cods)
    neg = negocios(list(pd.bdate_range(B3_INICIO, HOJE)), set(cods))
    cdi = sgs(12, HOJE - pd.DateOffset(years=4))
    log("ANBIMA, B3 e CDI ok")
    # ---------------- por ativo
    ativos, diarios, taxas = [], {}, {}
    lim12 = HOJE - pd.Timedelta(days=365)
    for x in a.to_dict("records"):
        cod = x["codigo"]
        x["coberto"] = bool(cod)
        if not cod:
            ativos.append(x)
            continue
        ah = an[an.codigo.eq(cod)].sort_values("data")
        rs = resumo.get(cod) or {}
        anos = max((pd.Timestamp(x["venc"]) - HOJE).days / 365.25, 0.05)
        dur = (float(ah.duration_anos.dropna().iloc[-1]) if ah.duration_anos.notna().any() else rs.get("duration") or min(anos, anos * 0.8 + 0.2))
        x["duration"] = round(float(dur), 2)
        of = ah.set_index("data").taxa_ind.dropna()
        if not len(of) and rs.get("taxa") is not None and rs.get("data"):
            of = pd.Series({pd.Timestamp(rs["data"]): rs["taxa"]})
        x["taxa"] = float(of.iloc[-1]) if len(of) else None
        x["data_taxa"] = of.index[-1].strftime("%Y-%m-%d") if len(of) else None
        h = of_hist.get(cod)
        v = serie_oficial(h) if h is not None and len(h) else pd.DataFrame()
        p = taxa_e_pnl(v, of, x["venc"], x["cupom"], x["idx"], x["duration"], cdi) if len(v) >= 3 else None
        hist = {"d": [], "pA": [], "uA": [], "tA": [], "tc": [], "car": [], "mtm": [], "res": []}
        x["fonte"] = None
        if p is not None:
            dts = pd.DatetimeIndex(v.data)
            z = pd.Series(p["ret"], index=dts).iloc[1:]
            gap = np.array([dias_uteis(i, j) for i, j in zip(dts[:-1], dts[1:])])
            diarios[cod] = z / np.sqrt(gap)
            taxas[cod] = pd.Series(p["taxa"], index=dts)
            z12 = diarios[cod][diarios[cod].index > lim12]
            if len(z12) >= 20:
                x["vol"] = float(z12.std(ddof=1) * np.sqrt(252) * 100)
                x["fonte"] = "ANBIMA oficial"
            m12 = dts > lim12
            for k_ in ("car", "mtm", "res"):
                vals = p[k_][m12]
                x[f"pnl_{k_}"] = float(np.nansum(vals) * 1e4) if np.isfinite(vals).any() else None      # bps do PU, 12 m
            x["pnl_total"] = float(np.nansum(p["ret"][m12]) * 1e4)
            x["res_abs"] = float(np.nanmean(np.abs(p["res"][m12])) * 1e4) if np.isfinite(p["res"][m12]).any() else None
            u = v.iloc[-1]
            x["pct_par"], x["pu"], x["data_pu"] = float(u.pu_ind / u.pu_par * 100), float(u.pu_ind), u.data.strftime("%Y-%m-%d")
            m1 = dts > dts[-1] - pd.Timedelta(days=30)
            x["ret_1m"] = float((np.nanprod(1 + p["ret"][m1]) - 1) * 100)
            t1 = taxas[cod].dropna()
            ant = t1[t1.index <= t1.index[-1] - pd.Timedelta(days=30)] if len(t1) else t1
            x["dtaxa_1m"] = float((t1.iloc[-1] - ant.iloc[-1]) * 100) if len(ant) else None
            x["taxa_inicio"] = float(t1.iloc[0]) if len(t1) else None
            r4 = lambda arr, c: [None if not np.isfinite(q) else round(float(q), c) for q in arr]
            keep = dts >= HOJE - pd.DateOffset(years=4)
            hist = {"d": [d.strftime("%Y-%m-%d") for d in dts[keep]], "pA": r4((v.pu_ind / v.pu_par * 100).values[keep], 3),
                    "uA": r4(v.pu_ind.values[keep], 4), "tA": r4(p["taxa"][keep], 4), "tc": p["fonte"][keep].tolist(),
                    "car": r4(p["car"][keep] * 1e4, 2), "mtm": r4(p["mtm"][keep] * 1e4, 2), "res": r4(p["res"][keep] * 1e4, 2)}
        # negócios B3 (taxa) e SND (% da curva)
        nb = neg[neg.cod.eq(cod)]
        g = nb[(nb.vol >= MIN_DIA) & (nb.taxa > 0)]
        tr = g.assign(tv=g.taxa * g.vol).groupby("data").agg(tv=("tv", "sum"), vol=("vol", "sum"))
        hist["neg"] = {"d": [d.strftime("%Y-%m-%d") for d in tr.index], "t": [round(q, 3) for q in tr.tv / tr.vol], "v": [round(q / 1e6, 2) for q in tr.vol]}
        x["neg_30d"] = int((nb.data >= HOJE - pd.Timedelta(days=30)).sum())
        x["giro_30d"] = float(nb[nb.data >= HOJE - pd.Timedelta(days=30)].vol.sum() / 1e6)
        x["_neg"] = g
        if x.get("vol") is None:                                       # sem série oficial: negócios semanais
            s = snd_precos(cod) if x["tipo"] in ("DEB", None) and not cod.startswith(("CRA", "CRI")) else pd.DataFrame()
            s = s[(s.qtd * s.pmed >= MIN_DIA) & s.pct.between(20, 150) & (s.data > lim12)] if len(s) else s
            if len(s) >= 10:
                x["vol"], _ = vol_semanal(s.data, s.pct)
                x["fonte"] = "SND (negócios)" if x["vol"] else None
            elif len(tr) >= 10:
                px = np.exp(-x["duration"] * (tr.tv / tr.vol) / 100)
                x["vol"], _ = vol_semanal(tr.index[tr.index > lim12], px[tr.index > lim12])
                x["fonte"] = "B3 (negócios)" if x["vol"] else None
        # hoje: negócios contra a ANBIMA de ontem
        hj = nb[(nb.data == HOJE) & (nb.taxa > 0)]
        if len(hj) and cod in taxas:
            t_ant = taxas[cod].dropna()
            t_ant = t_ant[t_ant.index < HOJE]
            gr = hj[hj.vol >= MIN_DIA] if (hj.vol >= MIN_DIA).any() else hj
            tx = float((gr.taxa * gr.vol).sum() / gr.vol.sum())
            if len(t_ant):
                x["hoje"] = {"n": int(len(hj)), "vol": float(hj.vol.sum() / 1e6), "taxa": tx, "ref": float(t_ant.iloc[-1]),
                             "ref_data": t_ant.index[-1].strftime("%Y-%m-%d"), "bps": (tx - t_ant.iloc[-1]) * 100, "varejo": not (hj.vol >= MIN_DIA).any()}
        x["hist"] = hist
        x["agenda"] = [e for e in agenda.get(cod, []) if e["status"] == "PREVISTO"
                       and HOJE <= pd.to_datetime(e["data"], dayfirst=True) <= HOJE + pd.Timedelta(days=180)][:6]
        ativos.append(x)
    A = pd.DataFrame(ativos)
    A["faixa"] = pd.cut(A.duration, [0, 1.5, 3, 5, 99], labels=False)
    ok = A.coberto & A.vol.notna() & A.fonte.eq("ANBIMA oficial")
    for i in A.index[A.coberto & A.vol.isna()]:
        par = A[ok & (A.idx == A.idx[i]) & (A.faixa == A.faixa[i])]
        par = par if len(par) >= 3 else A[ok & (A.faixa == A.faixa[i])]
        if len(par):
            A.loc[i, ["vol", "fonte"]] = [float(par.vol.median()), f"pares ({len(par)})"]
    log(f"ativos: {A.fonte.eq('ANBIMA oficial').sum()} com série oficial ({time.time() - t0:.0f}s)")
    # ---------------- emissores: fundamentos, ações, documentos
    cias = companhias()
    raizes = sorted({r for r in A.raiz.dropna()})
    FB = fundamentos_brutos()
    fund = {r: trimestres(FB[FB.raiz.eq(r)]) for r in raizes} if len(FB) else {}
    acoes = escolhe_acao(raizes, cias)
    docs = ipe(set(raizes))
    log(f"emissores: {len(raizes)}, com ITR/DFP: {sum(1 for f in fund.values() if len(f))}, com ação: {len(acoes)} ({time.time() - t0:.0f}s)")
    # ---------------- previsibilidade (pooled entre emissores)
    emis = {}
    pool_lead, pool_ev, pool_fr, pool_neg = [], [], [], []
    for r in raizes:
        aa = A[A.raiz.eq(r) & A.codigo.isin(taxas.keys())].sort_values("perc_nav", ascending=False)
        ref = aa.codigo.iloc[0] if len(aa) else None
        td = taxas.get(ref) if ref else None
        e = {"raiz": r, "nome": cias.nome.get(r) if r in cias.index else A[A.raiz.eq(r)].emissor.iloc[0], "setor": cias.setor.get(r) if r in cias.index else None,
             "perc": float(A[A.raiz.eq(r)].perc_nav.sum()), "codigos": A[A.raiz.eq(r)].codigo.tolist(), "ref": ref}
        F = fund.get(r)
        if F is not None and len(F):
            Fq = F[F.index >= HOJE - pd.DateOffset(years=ANOS_CVM)]
            e["fund"] = {"ref": [d.strftime("%Y-%m-%d") for d in Fq.index], "entrega": [str(q)[:10] if isinstance(q, str) else None for q in Fq.entrega],
                         **{k: [None if not np.isfinite(q) else round(float(q) / (1e6 if k in ("receita", "ebitda", "dl", "lucro", "divida", "caixa") else 1), 2)
                                for q in Fq[k].astype(float)] for k in ("receita", "ebitda", "dl", "lucro", "dl_ebitda", "cobertura", "margem")}}
            e["financeira"] = bool(F.financeira.any())
            u = F.dropna(subset=["dl_ebitda"])
            if len(u):
                e["alav"], e["alav_ant"] = float(u.dl_ebitda.iloc[-1]), float(u.dl_ebitda.iloc[-5]) if len(u) >= 5 else None
                e["cob"] = float(u.cobertura.iloc[-1]) if np.isfinite(u.cobertura.iloc[-1]) else None
            ent = pd.to_datetime(F.entrega.dropna()).sort_values()
            lags = (pd.to_datetime(F.entrega) - F.index).dt.days.dropna()
            if len(ent):
                ult_ref = F.index[F.entrega.notna()].max()
                prox_ref = (ult_ref + pd.offsets.QuarterEnd(1)).normalize()
                mesmo = lags[(lags.index.month == prox_ref.month)]
                base = mesmo if len(mesmo) else lags
                e["prox_resultado"] = {"ref": prox_ref.strftime("%Y-%m-%d"), "esperado": (prox_ref + pd.Timedelta(days=float(base.median()))).strftime("%Y-%m-%d"),
                                       "min": (prox_ref + pd.Timedelta(days=float(base.min()))).strftime("%Y-%m-%d"),
                                       "max": (prox_ref + pd.Timedelta(days=float(base.max()))).strftime("%Y-%m-%d"), "n": int(len(base))}
            if td is not None:                                         # reação da taxa ao resultado e à alavancagem
                ev, sd = reacao(td, list(ent))
                lev = F.dl_ebitda.diff()
                for d, dlt in ev:
                    i_ = F.index[pd.to_datetime(F.entrega) == d]
                    dl_ = float(lev.get(i_[0], np.nan)) if len(i_) else np.nan
                    pool_ev.append({"raiz": r, "data": d, "d": dlt, "z": dlt / sd if sd else np.nan, "dlev": dl_})
                e["eventos_res"] = [{"data": d.strftime("%Y-%m-%d"), "bps": round(dlt, 1)} for d, dlt in ev]
                e["sd5"] = sd
        dr = docs[docs.raiz.eq(r)]
        fr = dr[dr.Categoria.eq("Fato Relevante")]
        if td is not None and len(fr):
            ev, sd = reacao(td, list(fr.data.dropna().unique()))
            pool_fr += [{"raiz": r, "data": d, "d": dlt, "z": dlt / sd if sd else np.nan} for d, dlt in ev]
        rel = dr[dr.Categoria.isin(["Fato Relevante", "Comunicado ao Mercado", "Aviso aos Debenturistas", "Escrituras e aditamentos de debêntures",
                                    "Informações de Companhias em Recuperação Judicial ou Extrajudicial", "Assembleia"]) & (dr.data >= HOJE - pd.Timedelta(days=60))]
        e["docs"] = [{"data": q.data.strftime("%Y-%m-%d"), "cat": q.Categoria, "assunto": str(q.Assunto if isinstance(q.Assunto, str) else q.Tipo or "")[:160],
                      "link": q.Link_Download} for q in rel.sort_values("data", ascending=False).head(12).itertuples()]
        e["fr_datas"] = [d.strftime("%Y-%m-%d") for d in fr.data.dropna()]
        ac = acoes.get(r)
        if ac:
            s = ac["serie"]
            e["acao"] = {"ticker": ac["ticker"], "via": ac["via"], "preco": ac["preco"], "d1": (ac["preco"] / ac["ant"] - 1) * 100 if ac["preco"] and ac["ant"] else None,
                         "s5": float(np.log(s.iloc[-1] / s.iloc[-6]) * 100) if len(s) > 6 else None,
                         "vol_s": float(np.log(semanal(s)).diff().std() * 100) if len(s) > 100 else None}
            if td is not None:
                ax = acao_x_credito(td, s)
                if ax:
                    j = ax.pop("_j")
                    e["acao_credito"] = ax
                    za, zt = (j.a - j.a.mean()) / j.a.std(), (j.t - j.t.mean()) / j.t.std()
                    pool_lead += [(d, za.iloc[i], zt.iloc[i + 1], j.a.iloc[i], j.t.iloc[i + 1], r) for i, d in enumerate(j.index[:-1])]
                    e["acao_credito"]["hoje_s"] = e["acao"]["s5"]
        emis[r] = e
    for cod, ng in zip(A.codigo, A["_neg"]):
        if cod in taxas and isinstance(ng, pd.DataFrame):
            pool_neg += [(d, g_, m0, m1, gc, vv, cod) for d, g_, m0, m1, gc, vv in negocio_x_marcacao(ng, taxas[cod])]
    prev = previsibilidade(pool_lead, pool_ev, pool_fr, pool_neg)
    A["prev_hoje"] = pd.Series([None] * len(A), index=A.index, dtype=object)
    juros = juros_b3()
    previsao_hoje(A, an, emis, juros)
    try:
        p5, info5 = modelo5(list(A.codigo.dropna()))
    except Exception as e:
        log(f"previsão 5d: erro {type(e).__name__} {e}")
        p5, info5 = {}, {}
    A["prev5"] = [p5.get(c) if p5.get(c, {}).get("mu") is not None else None for c in A.codigo]
    A["regras_dias"] = [p5.get(c, {}).get("regras") for c in A.codigo]
    A.attrs["info5"] = info5
    log(f"previsão 5d: {len(p5)} ativos ({time.time() - t0:.0f}s)")
    A.attrs["juros"] = juros
    log(f"previsibilidade ok ({time.time() - t0:.0f}s)")
    # ---------------- mercado e risco
    idx_c = indices_credito()
    mkt = {k: yahoo(v, "2y") for k, v in MERCADO.items()}
    td_ = tesouro()
    return A, emis, prev, idx_c, mkt, td_, diarios


def previsibilidade(pool_lead, pool_ev, pool_fr, pool_neg):
    out = []
    if len(pool_lead) >= 30:
        L = pd.DataFrame(pool_lead, columns=["data", "za", "zt1", "a", "t1", "raiz"])
        r, n = corr(L.za, L.zt1, rank=True)
        rg = regressao(L.a, L.t1)
        oos = fora_da_amostra(L.a, L.t1, L.data, 1.0 * L.a.std())
        lo, hi = fisher(r, n)
        out.append({"id": "acao_lidera", "sinal": "Ação do emissor na semana → Δ taxa da debênture na semana seguinte", "n": n, "r": r, "lo": lo, "hi": hi,
                    "b": rg["b"] if rg else None, "unid": "bps por 1% da ação", "oos": oos, "emissores": int(L.raiz.nunique())})
    if len(pool_neg) >= 30:
        N = pd.DataFrame(pool_neg, columns=["data", "gap", "mv0", "mv1", "gap_fech", "vol", "cod"])
        N = N[N.vol >= NEG_INST].assign(gap=lambda d: d.gap.clip(-150, 150), gap_fech=lambda d: d.gap_fech.clip(-150, 150))
        r0 = regressao(N.gap, N.mv0)
        r1 = regressao(N.gap_fech, N.mv1)
        lo, hi = fisher(*corr(N.gap, N.mv0))
        out.append({"id": "negocio_marcacao", "sinal": f"Negócio na B3 (≥ R$ {NEG_INST / 1e6:.0f} mi no dia) longe da ANBIMA de ontem → ANBIMA do dia anda na direção",
                    "n": r0["n"] if r0 else len(N), "r": corr(N.gap, N.mv0)[0], "lo": lo, "hi": hi, "b": r0["b"] if r0 else None,
                    "unid": "fração do gap que a ANBIMA anda no dia", "oos": fora_da_amostra(N.gap, N.mv0, N.data, 10), "ativos": int(N.cod.nunique())})
        if r1:
            lo, hi = fisher(*corr(N.gap_fech, N.mv1))
            out.append({"id": "negocio_dia_seguinte", "sinal": "Negócio longe da ANBIMA do próprio dia → ANBIMA do dia seguinte", "n": r1["n"],
                        "r": corr(N.gap_fech, N.mv1)[0], "lo": lo, "hi": hi, "b": r1["b"], "unid": "fração do gap restante no dia seguinte",
                        "oos": fora_da_amostra(N.gap_fech, N.mv1, N.data, 10)})
    if len(pool_ev) >= 15:
        E = pd.DataFrame(pool_ev)
        r, n = corr(E.dlev, E.d)
        lo, hi = fisher(r, n)
        out.append({"id": "resultado", "sinal": "Resultado trimestral (ITR/DFP): alta da dívida líquida/EBITDA → Δ taxa em 5 dias úteis", "n": n, "r": r, "lo": lo, "hi": hi,
                    "b": regressao(E.dlev, E.d)["b"] if regressao(E.dlev, E.d) else None, "unid": "bps por +1x de alavancagem",
                    "z_medio": float(E.z.abs().mean()), "eventos": len(E), "oos": fora_da_amostra(E.dlev, E.d, E.data, 0.3)})
    if len(pool_fr) >= 10:
        Fr = pd.DataFrame(pool_fr)
        out.append({"id": "fato_relevante", "sinal": "Fato relevante → |Δ taxa| em 5 dias úteis (vs. semana normal)", "n": len(Fr),
                    "z_medio": float(Fr.z.abs().mean()), "unid": "|z| médio (semana normal ≈ 0,8)", "r": None})
    return out


def risco(A, diarios, idx_c, mkt, td_):
    lim = HOJE - pd.Timedelta(days=365)
    k = A[A.vol.notna() & A.coberto].copy()
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
    Sg = np.outer(s, s) * R
    vp = float(np.sqrt(w @ Sg @ w))
    A.loc[k.index, "contrib_pct"] = w * (Sg @ w) / vp / vp * 100 if vp else 0
    of = np.array([c in diarios for c in cods])
    wo = w[of]
    so = float(np.sqrt(wo @ Sg[np.ix_(of, of)] @ wo)) / np.sqrt(252) if of.any() else np.nan
    Rt = pd.DataFrame({c: diarios[c] for c in np.array(cods)[of]}).sort_index()
    if not len(Rt.columns):
        Rt = pd.DataFrame(index=pd.DatetimeIndex([]))
    Rt = Rt[Rt.index > HOJE - pd.Timedelta(days=420)]
    wv = pd.Series(wo, index=Rt.columns)
    ret = (Rt.fillna(0) * wv).sum(axis=1)
    cobre = Rt.notna().mul(wv * 100).sum(axis=1)
    ret = ret[cobre > 0.8 * cobre.max()]
    ult = ret.iloc[-250:]
    # fatores: regressão do retorno diário da carteira oficial (excesso sobre a média) nas variações dos fatores
    fat = {}
    for nome, d in idx_c.items():
        if nome in ("IDA DI", "IDA IPCA ex-infra", "JGP CDI Low Rated", "IMA-B", "IRF-M"):
            fat[nome] = np.log(d["serie"]).diff() * 100
    for nome in ("Ibovespa", "Dólar"):
        if mkt.get(nome):
            fat[nome] = np.log(mkt[nome]["serie"]).diff() * 100
    for c_ in td_.columns if len(td_) else []:
        fat[c_ + " (bps)"] = td_[c_].diff() * 100
    fatores = []
    y = ret * 100
    for nome, f in fat.items():
        j = pd.concat([y.rename("y"), f.rename("f")], axis=1).dropna()
        j = j[j.index > lim]
        rg = regressao(j.f, j.y)
        if rg:
            fatores.append({"fator": nome, "beta": rg["b"], "t": rg["t"], "r2": rg["r2"], "n": rg["n"], "r": corr(j.f, j.y)[0]})
    top = k.sort_values("perc_nav", ascending=False).codigo.head(25).tolist()
    Ct = C.reindex(index=[c for c in top if c in C.index], columns=[c for c in top if c in C.columns])
    return {"vol": vp * 100, "vol_sem_div": float((w * s).sum() * 100), "var95_1d": 1.645 * vp / np.sqrt(252) * 100,
            "var95_1m": 1.645 * vp * np.sqrt(21 / 252) * 100, "cobertura": float(k.perc_nav.sum()), "pl_codigo": float(A[A.coberto].perc_nav.sum()),
            "rho": rho, "n": n, "duration": float((k.perc_nav * k.duration).sum() / max(k.perc_nav.sum(), 1e-9)),
            "dv01": float(-(k.perc_nav / 100 * k.duration).sum()),
            "backtest": {"dias": int(len(ult)), "violacoes": float((ult < -1.645 * so).mean() * 100) if len(ult) else None,
                         "vol_modelo": float(so * np.sqrt(252) * 100) if so == so else None, "vol_real": float(ult.std() * np.sqrt(252) * 100) if len(ult) else None},
            "serie": {"d": [d.strftime("%Y-%m-%d") for d in ret.index], "ret": [round(v * 100, 4) for v in ret], "var": round(-1.645 * so * 100, 4)},
            "fatores": fatores, "corr": {"cods": Ct.index.tolist(), "m": [[None if not np.isfinite(v) else round(float(v), 2) for v in row] for row in Ct.values]}}


def valor_relativo(A, emis):
    """Spread (DI+) × alavancagem e duration, entre os ativos da carteira: resíduo > 0 = paga mais que o perfil explica."""
    k = A[A.coberto & A.idx.eq("DI+") & A.taxa.notna()].copy()
    k["alav"] = k.raiz.map(lambda r: (emis.get(r) or {}).get("alav"))
    k = k[k.alav.notna() & k.alav.between(-2, 15) & k.taxa.between(0, 15)]
    if len(k) < 8:
        return None
    X = np.c_[np.ones(len(k)), k.alav.values, k.duration.values]
    b, *_ = np.linalg.lstsq(X, k.taxa.values, rcond=None)
    k["res_bps"] = (k.taxa - X @ b) * 100
    r2 = float(1 - k.res_bps.var() / (k.taxa * 100).var())
    return {"b": [round(float(q), 4) for q in b], "r2": r2, "n": len(k),
            "pontos": [{"cod": q.codigo, "emissor": q.emissor, "alav": round(q.alav, 2), "taxa": q.taxa, "dur": q.duration, "res": round(q.res_bps, 0)} for q in k.itertuples()]}


# ================================================================== previsão de 5 dias úteis (% do PU par) com intervalo
F5 = ["m5", "m20", "m60", "dev20", "dev60", "vol20", "vol60", "vol250", "vratio", "parado20", "maxabs20", "nivel", "rel5", "rel20", "anos", "idx_c",
      "press5", "press20", "varejo5", "varejo20", "lvol5", "lvol20", "imb5", "imb20", "dias_sem_neg"]
NOME5 = {"m5": "andou na semana", "m20": "andou no mês", "m60": "andou em 3 meses", "dev20": "longe da média de 20 dias", "dev60": "longe da média de 60 dias",
         "vol20": "vol 20 dias", "vol60": "vol 60 dias", "vol250": "vol 1 ano", "vratio": "vol subindo", "parado20": "marcação parada", "maxabs20": "maior salto no mês",
         "nivel": "% do par (abaixo do par sobe pelo carrego)", "rel5": "semana contra o mercado", "rel20": "mês contra o mercado", "anos": "prazo",
         "idx_c": "indexador", "press5": "negócios da semana vs marcação", "press20": "negócios do mês vs marcação", "varejo5": "varejo da semana vs marcação",
         "varejo20": "varejo do mês vs marcação", "lvol5": "volume da semana", "lvol20": "volume do mês", "imb5": "lado dos negócios (semana)",
         "imb20": "lado dos negócios (mês)", "dias_sem_neg": "dias sem negócio"}


def painel5(an_cods=None):
    """Painel (papel, dia) com features point in time, a partir de todo histórico ANBIMA em cache e dos negócios B3."""
    rows, pares = [], []
    for f in (DADOS / "anbima_hist").glob("*.parquet"):
        h = pd.read_parquet(f)
        if h.pu_ind.notna().sum() < 120 or not (h.pu_par > 0).any():
            continue
        v = serie_oficial(h)
        if len(v) < 120:
            continue
        rows.append(pd.DataFrame({"cod": f.stem, "data": v.data.values, "p": (v.pu_ind / v.pu_par).values, "E": v.E.values}))
        pares.append(pd.DataFrame({"cod": f.stem, "data": h.data.values, "pu_par": h.pu_par.values}).dropna())
    S = pd.concat(rows, ignore_index=True).sort_values(["cod", "data"]).reset_index(drop=True)
    S = S[S.p.between(0.2, 1.6)].reset_index(drop=True)
    g = S.groupby("cod")
    S["lp"] = np.log(S.p) * 1e4
    S["y0"] = S.lp - g.lp.shift()
    S.loc[(S.E > 0) | (S.y0.abs() > 1500), "y0"] = np.nan
    S["y5"] = g.lp.shift(-5) - S.lp
    pag5 = g.E.transform(lambda e: e[::-1].rolling(5, min_periods=1).sum()[::-1].shift(-1))
    S.loc[((g.data.shift(-5) - S.data).dt.days > 9) | (pag5 > 0) | (S.y5.abs() > 3000), "y5"] = np.nan
    r = lambda c, k, fn, mp: S[c].groupby(S.cod).transform(lambda x: getattr(x.rolling(k, min_periods=mp), fn)())
    y0f = S.y0.fillna(0).rename("y0f")
    S["y0f"] = y0f
    for k in (5, 20, 60):
        S[f"m{k}"] = r("y0f", k, "sum", int(k * .8))
    S["dev20"], S["dev60"] = S.lp - r("lp", 20, "mean", 15), S.lp - r("lp", 60, "mean", 45)
    S["vol20"], S["vol60"], S["vol250"] = r("y0", 20, "std", 12), r("y0", 60, "std", 40), r("y0", 250, "std", 120)
    S["vratio"] = S.vol20 / S.vol60
    S["parado"] = (S.y0.abs() < 0.5).astype(float)
    S["parado20"] = r("parado", 20, "mean", 12)
    S["absy"] = S.y0.abs()
    S["maxabs20"] = r("absy", 20, "max", 12)
    S["nivel"] = S.lp
    cad = cadastro_b3().drop_duplicates("cod").set_index("cod")
    S["idx"] = S.cod.map(indexadores(list(S.cod.unique())))
    S["idx_c"] = S.idx.map({"DI+": 0, "IPCA": 1, "PRE": 2, "%DI": 3, "IGPM": 4}).fillna(5)
    mk = S.groupby(["idx", "data"]).y0.mean().rename("mk0").reset_index().sort_values("data")
    mk["mk5"] = mk.groupby("idx").mk0.transform(lambda x: x.rolling(5, min_periods=4).sum())
    mk["mk20"] = mk.groupby("idx").mk0.transform(lambda x: x.rolling(20, min_periods=15).sum())
    S = S.merge(mk, on=["idx", "data"], how="left")
    S["rel5"], S["rel20"] = S.m5 - S.mk5, S.m20 - S.mk20
    S["anos"] = (pd.to_datetime(S.cod.map(snd_caracteristicas().venc)) - S.data).dt.days / 365.25
    # negócios B3 contra a ANBIMA do mesmo dia
    n = negocios(list(pd.bdate_range(B3_INICIO, HOJE)), set(S.cod))
    n = n[(n.vol > 0) & (n.pu > 0)]
    P = pd.concat(pares).drop_duplicates(["cod", "data"])
    n = n.merge(P, on=["cod", "data"]).merge(S[["cod", "data", "p"]], on=["cod", "data"])
    n["gf"] = np.log(n.pu / n.pu_par / n.p) * 1e4
    n = n[n.gf.abs() < 500]
    inst = n.vol >= MIN_DIA
    d = n[inst].assign(gv=lambda x: x.gf * x.vol, sv=lambda x: np.sign(x.gf) * x.vol).groupby(["cod", "data"])[["gv", "vol", "sv"]].sum()
    rv = n[~inst].assign(rgv=lambda x: x.gf * x.vol).groupby(["cod", "data"]).agg(rgv=("rgv", "sum"), rvol=("vol", "sum"))
    S = S.merge(d.join(rv, how="outer").reset_index(), on=["cod", "data"], how="left")
    for c in ("gv", "vol", "sv", "rgv", "rvol"):
        S[c] = S[c].fillna(0)
    for k in (5, 20):
        gv, vol = r("gv", k, "sum", 1), r("vol", k, "sum", 1)
        S[f"press{k}"], S[f"lvol{k}"] = (gv / vol).where(vol >= 1e6), np.log1p(vol / 1e6)
        S[f"imb{k}"] = (r("sv", k, "sum", 1) / vol).where(vol >= 1e6)
        rg, rvv = r("rgv", k, "sum", 1), r("rvol", k, "sum", 1)
        S[f"varejo{k}"] = (rg / rvv).where(rvv > 0)
    S["dias_sem_neg"] = S.groupby("cod").vol.transform(lambda x: (x > 0).astype(int).groupby((x > 0).cumsum()).cumcount())
    S.loc[S.data < B3_INICIO, ["press5", "press20", "lvol5", "lvol20", "imb5", "imb20", "varejo5", "varejo20", "dias_sem_neg"]] = np.nan
    S["mk_y5"] = S.groupby(["idx", "data"]).y5.transform("mean")
    S["y5i"] = S.y5 - S.mk_y5
    return S


def modelo5(cods):
    """Treina 1x por dia (cache): média idiossincrática de 5 dias (LightGBM huber) e escala (log|y5|); intervalo 8–92% = média + fator
    conformal × escala, fator por faixa de escala, calibrado nos 6 meses antes de um embargo de 12 dias. Devolve previsões para `cods`."""
    try:
        import lightgbm as lgb
    except ImportError:
        log("previsão 5d: instale lightgbm")
        return {}, {}
    S = painel5()
    hoje_ = S.groupby("cod").data.transform("max") == S.data
    ult = S[hoje_ & S.cod.isin(cods)]
    treino = S.dropna(subset=["y5"])
    treino = treino[treino.data >= B3_INICIO - pd.DateOffset(years=3)]
    corte = treino.data.max() - pd.Timedelta(days=12)
    cal_ini = corte - pd.Timedelta(days=200)
    prm = dict(n_estimators=250, learning_rate=0.03, num_leaves=15, min_child_samples=500, subsample=0.7, subsample_freq=1, colsample_bytree=0.7,
               reg_lambda=5, verbose=-1)
    fm = DADOS / f"modelo5_{HOJE:%Y%m%d}.json"
    if fm.exists():
        j = json.loads(fm.read_text())
        mu, sg = lgb.Booster(model_str=j["mu"]), lgb.Booster(model_str=j["sg"])
        cortes, kq, info = np.array(j["cortes"]), {int(k): v for k, v in j["kq"].items()}, j["info"]
    else:
        tr = treino[treino.data < cal_ini - pd.Timedelta(days=12)]
        mu = lgb.LGBMRegressor(objective="huber", alpha=15, **prm).fit(tr[F5], tr.y5i).booster_
        sg = lgb.LGBMRegressor(objective="l2", **prm).fit(tr[F5], np.log(np.abs(tr.y5) + 1)).booster_
        cal = treino[treino.data >= cal_ini]
        esc = np.exp(sg.predict(cal[F5]))
        z = (cal.y5 - mu.predict(cal[F5])) / esc
        cortes = np.quantile(esc, [0.2, 0.4, 0.6, 0.8])
        fx = np.digitize(esc, cortes)
        kq = {b: np.quantile(z[fx == b], [0.08, 0.92]).tolist() for b in range(5)}
        cob = float(np.mean([(lo <= zz <= hi) for zz, b in zip(z, fx) for lo, hi in [kq[b]]]) * 100)
        info = {"treino": int(len(tr)), "papeis": int(S.cod.nunique()), "ate": str(tr.data.max())[:10], "calibracao": f"{cal_ini:%d/%m/%y}–{corte:%d/%m/%y}",
                "cobertura_cal": cob}
        # depois de calibrar, retreina a média com tudo até o embargo (mais dado recente)
        tr2 = treino[treino.data < corte]
        mu = lgb.LGBMRegressor(objective="huber", alpha=15, **prm).fit(tr2[F5], tr2.y5i).booster_
        fm.write_text(json.dumps({"mu": mu.model_to_string(), "sg": sg.model_to_string(), "cortes": cortes.tolist(), "kq": kq, "info": info}))
        for velho_ in DADOS.glob("modelo5_*.json"):
            if velho_ != fm:
                velho_.unlink(missing_ok=True)
    # previsões de hoje + decil no universo de hoje
    U = S[hoje_ & (S.data >= S.data.max() - pd.Timedelta(days=5))]
    mu_u = pd.Series(mu.predict(U[F5]), index=U.cod.values)
    out = {}
    if len(ult):
        m = mu.predict(ult[F5])
        e = np.exp(sg.predict(ult[F5]))
        b = np.digitize(e, cortes)
        contrib = mu.predict(ult[F5], pred_contrib=True)
        for i, (c, row) in enumerate(zip(ult.cod, ult.itertuples())):
            lo, hi = m[i] + kq[b[i]][0] * e[i], m[i] + kq[b[i]][1] * e[i]
            pct = float((mu_u < m[i]).mean() * 100)
            top = sorted(zip(F5, contrib[i][:-1]), key=lambda t: -abs(t[1]))[:3]
            out[c] = {"mu": float(m[i]), "q08": float(lo), "q92": float(hi), "pct_univ": pct, "data": row.data.strftime("%Y-%m-%d"),
                      "sinal": "queda" if pct <= 10 else "alta" if pct >= 90 else None,
                      "porque": [{"f": NOME5.get(f, f), "bps": round(float(v), 1), "valor": None if not np.isfinite(getattr(row, f)) else round(float(getattr(row, f)), 2)} for f, v in top]}
    info["n_univ_hoje"] = int(len(U))
    # regras de 1 a 20 dias: gatilhos no último dia de cada papel da carteira
    for row in ult.itertuples():
        z5 = row.rel5 / (row.vol60 * np.sqrt(5)) if row.vol60 and row.vol60 == row.vol60 else np.nan
        own = row.gv / row.vol if row.vol else np.nan
        g_ = {"di_pior": row.idx == "DI+" and z5 <= -2, "di_melhor": row.idx == "DI+" and row.rel5 >= 50,
              "ipca_extremo": row.idx == "IPCA" and z5 <= -4, "di_negocio": row.idx == "DI+" and row.vol >= 20e6 and own >= 30,
              "semana_acima": z5 >= 2.5 and row.press5 >= 10 and np.expm1(row.lvol5) >= 1}
        ativ = [{"regra": k, "nome": REGRAS_DIAS[k]["nome"], "z5": None if z5 != z5 else round(float(z5), 1), "rel5": round(float(row.rel5), 1),
                 "gap": None if own != own else round(float(own), 1)} for k, v in g_.items() if v is True or (isinstance(v, np.bool_) and v)]
        if ativ:
            out.setdefault(row.cod, {})["regras"] = ativ
    return out, info


# ================================================================== previsão da marcação ANBIMA de hoje
def previsao_hoje(A, an, emis, juros):
    """Para cada ativo, o movimento esperado do % do PU par na ANBIMA de hoje (bps), somando: negócios do próprio papel na B3
    (fração do gap pelo volume), juros (DAP ao vivo × duration, papéis IPCA), outros papéis do emissor, tom do mercado e a ação."""
    if not len(an):
        return
    d1 = an[an.data < HOJE].data.max()
    a1 = an[an.data == d1].drop_duplicates("codigo").set_index("codigo")
    a2 = an[an.data == an[an.data < d1].data.max()].drop_duplicates("codigo").set_index("codigo")
    par1 = a1.pu / a1.pct_par * 100
    par2 = a2.pu / a2.pct_par * 100
    cresc = (par1 / par2.reindex(par1.index)).where(lambda g_: g_.between(0.999, 1.003)).fillna(1.0004)   # PU par: um dia de rendimento
    nh = negocios([HOJE])
    nh = nh[nh.cod.isin(a1.index) & (nh.vol >= MIN_DIA) & (nh.pu > 0)].copy()
    nh["gap"] = np.log(nh.pu.values / (par1.reindex(nh.cod).values * cresc.reindex(nh.cod).values) / (a1.pct_par.reindex(nh.cod).values / 100)) * 1e4
    nh = nh[nh.gap.abs() < 500]
    g = nh.assign(gv=nh.gap * nh.vol).groupby("cod").agg(gv=("gv", "sum"), vol=("vol", "sum"), n=("vol", "size"))
    g["gap"] = g.gv / g.vol
    cad = cadastro_b3().drop_duplicates("cod").set_index("cod")
    emis_cod = cad.emissor.str.upper().str.strip()
    g["emissor"] = [emis_cod.get(c) for c in g.index]
    ix = indexadores(list(g.index))
    g["idx"] = [ix[c] for c in g.index]
    grande = g[g.vol >= 1e6]
    mkt = {i: float(x.gv.sum() / x.vol.sum()) for i, x in grande.groupby("idx") if x.vol.sum() >= 50e6}
    dap = {int(v["prazo"][:-1]): v["bps"] for v in (juros or {}).get("DAP", []) if v.get("bps") is not None}
    for i in A.index[A.coberto]:
        x = A.loc[i]
        cod, comp = x.codigo, {}
        dur = x.duration if x.duration == x.duration else 2.0
        tx = x.taxa if isinstance(x.taxa, (int, float)) and x.taxa == x.taxa else 0.0
        dm = dur / (1 + tx / 100) if x.idx in ("IPCA", "PRE", "IGPM") else dur
        usa_dap = x.idx == "IPCA" and any(abs(v) > 0.05 for v in dap.values())
        juro = 0.0
        if usa_dap:
            k = min(dap, key=lambda a_: abs(a_ - dur))
            juro = -dm * dap[k]
            if abs(juro) >= 0.5:
                comp["juros"] = {"bps": juro, "txt": f"DAP {k}a {dap[k]:+.1f} bps × duration mod. {dm:.1f}"}
        if cod in g.index:
            r = g.loc[cod]
            fr, ac = next((f, h) for lim, f, h in FRACAO_GAP if r.vol < lim)
            comp["negocio"] = {"bps": fr * (r.gap - juro), "txt": f"{int(r.n)} negócios, R$ {r.vol / 1e6:.1f} mi a {r.gap:+.0f} bps do % par de ontem"
                               + (f" ({r.gap - juro:+.0f} além do juro)" if juro else "") + f" (acerto histórico {ac}%)"}
        else:
            em = str(emis_cod.get(cod, "") or "")
            o = grande[grande.emissor.eq(em)] if em else grande.iloc[:0]
            if len(o):
                gi = float(o.gv.sum() / o.vol.sum())
                comp["emissor"] = {"bps": B_EMISSOR * gi, "txt": f"outros papéis do emissor: R$ {o.vol.sum() / 1e6:.0f} mi a {gi:+.0f} bps"}
            elif x.idx in mkt and not usa_dap:                     # IPCA com DAP ao vivo: o juro já entra abaixo
                comp["mercado"] = {"bps": B_MERCADO * mkt[x.idx], "txt": f"mercado {x.idx}: gap médio {mkt[x.idx]:+.0f} bps"}
        ac_ = (emis.get(x.raiz) or {}).get("acao") if isinstance(x.raiz, str) else None
        if ac_ and ac_.get("d1") is not None and abs(ac_["d1"]) >= 2:
            comp["acao"] = {"bps": B_ACAO * ac_["d1"], "txt": f"{ac_['ticker']} {ac_['d1']:+.1f}% hoje"}
        if comp:
            tot = float(sum(c["bps"] for c in comp.values()))
            A.at[i, "prev_hoje"] = {"comp": comp, "total": tot, "dtaxa": -tot / max(dm, 0.05), "impacto_pl": tot * x.perc_nav / 100}
    A.attrs["mercado_hoje"] = {k: round(v, 1) for k, v in mkt.items()}
    A.attrs["data_ref"] = d1.strftime("%Y-%m-%d")


# ================================================================== o que olhar (sinais de hoje)
def sinais(A, emis, prev, ipca, agenda_tot):
    P = {p["id"]: p for p in prev}
    out = []
    for x in A[A.coberto].itertuples():                      # previsão da marcação de hoje: negócio relevante do próprio papel
        ph = getattr(x, "prev_hoje", None)
        c = (ph or {}).get("comp", {}).get("negocio") if isinstance(ph, dict) else None
        if c and abs(c["bps"]) >= 3:
            out.append({"tipo": "Negócio hoje", "ativo": x.codigo, "emissor": x.emissor, "quando": "hoje",
                        "texto": c["txt"] + f" → ANBIMA de hoje esperada {ph['total']:+.0f} bps de % par ({ph['dtaxa']:+.1f} bps de taxa)",
                        "impacto_bps_pl": ph["impacto_pl"], "peso": x.perc_nav})
    pl = P.get("acao_lidera")
    usa = pl and pl.get("lo") is not None and (pl["lo"] > 0 or pl["hi"] < 0) and (pl.get("oos") or {}).get("acerto") and pl["oos"]["acerto"] > 55
    for r, e in emis.items():
        ac = e.get("acao")
        if ac and ac.get("s5") is not None and ac.get("vol_s") and abs(ac["s5"]) >= 1.5 * ac["vol_s"]:
            dur = A[A.raiz.eq(r)].duration.mean()
            esp = pl["b"] * ac["s5"] if usa else None
            out.append({"tipo": "Ação do emissor", "ativo": ac["ticker"], "emissor": e["nome"], "quando": "semana",
                        "texto": f"{ac['ticker']} {ac['s5']:+.1f}% em 5 dias ({ac['s5'] / ac['vol_s']:+.1f} desvios semanais)"
                                 + (f"; histórico (n={pl['n']}): taxa {esp:+.0f} bps na semana seguinte" if esp is not None else "; sinal da ação não é estatisticamente útil para o crédito"),
                        "impacto_bps_pl": -dur * esp * e["perc"] / 100 if esp is not None else None, "peso": e["perc"]})
        vistos = set()
        for d in e.get("docs", []):
            if d["assunto"] in vistos:
                continue
            vistos.add(d["assunto"])
            if pd.Timestamp(d["data"]) >= HOJE - pd.Timedelta(days=7) and d["cat"] in ("Fato Relevante", "Aviso aos Debenturistas", "Informações de Companhias em Recuperação Judicial ou Extrajudicial"):
                out.append({"tipo": d["cat"], "ativo": ", ".join(e["codigos"][:3]), "emissor": e["nome"], "quando": d["data"], "texto": d["assunto"], "link": d["link"],
                            "impacto_bps_pl": None, "peso": e["perc"]})
        pr = e.get("prox_resultado")
        if pr and HOJE <= pd.Timestamp(pr["esperado"]) <= HOJE + pd.Timedelta(days=15):
            out.append({"tipo": "Resultado", "ativo": ", ".join(e["codigos"][:3]), "emissor": e["nome"], "quando": pr["esperado"],
                        "texto": f"resultado de {pr['ref'][5:7]}/{pr['ref'][:4]} esperado ~{pr['esperado'][8:]}/{pr['esperado'][5:7]} (entre {pr['min'][8:]}/{pr['min'][5:7]} e {pr['max'][8:]}/{pr['max'][5:7]})"
                                + (f"; alavancagem atual {e['alav']:.1f}x" if e.get("alav") is not None else ""),
                        "impacto_bps_pl": None, "peso": e["perc"]})
    for ev in agenda_tot:
        if pd.Timestamp(ev["data"]) <= HOJE + pd.Timedelta(days=10) and re.search(r"Amortiza|Vencimento|Resgate|Opção|Call", ev["evento"], re.I):
            out.append({"tipo": "Agenda", "ativo": ev["ativo"], "emissor": ev["emissor"], "quando": ev["data"], "texto": f"{ev['evento']} {ev['pct']}",
                        "impacto_bps_pl": None, "peso": ev["peso"]})
    for d in ipca[:1]:
        if d <= HOJE + pd.Timedelta(days=10):
            out.append({"tipo": "Macro", "ativo": "IPCA", "emissor": "", "quando": d.strftime("%Y-%m-%d"), "texto": "divulgação do IPCA (IBGE)",
                        "impacto_bps_pl": None, "peso": float(A[A.idx.eq("IPCA")].perc_nav.sum())})
    for x in A[A.coberto & A.res_abs.notna()].itertuples() if "res_abs" in A else []:
        if x.res_abs > 15:
            out.append({"tipo": "Dado", "ativo": x.codigo, "emissor": x.emissor, "quando": "12m",
                        "texto": f"resíduo médio do PnL {x.res_abs:.0f} bps/dia: PU e taxa oficiais não fecham (evento não informado ou marcação)",
                        "impacto_bps_pl": None, "peso": x.perc_nav})
    ordem = {"Fato Relevante": 0, "Informações de Companhias em Recuperação Judicial ou Extrajudicial": 0, "Aviso aos Debenturistas": 1, "Negócio hoje": 2,
             "Ação do emissor": 3, "Resultado": 4, "Agenda": 5, "Macro": 6, "Dado": 7}
    return sorted(out, key=lambda s: (ordem.get(s["tipo"], 8), -(abs(s["impacto_bps_pl"]) if s.get("impacto_bps_pl") is not None else 0), -(s.get("peso") or 0)))


# ================================================================== página

# ================================================================== fundos internos: carteiras, cota, fluxo
PAINEL_HTML = Path(os.environ.get("PAINEL_HTML") or Path.home() / "OneDrive/Documents/painel-fundos/painel_fundos.html")
COD_B3 = re.compile(r"^([A-Z]{4}\d{2}|[A-Z]{4}[A-Z0-9]\d|CR[AI][0-9A-Z]{6,12}|\d{2}[A-L]\d{7})$")


def painel_dados():
    s = PAINEL_HTML.read_text(encoding="utf-8")
    b = max(re.findall(r"<script>(.*?)</script>", s, re.S), key=len)
    return json.JSONDecoder().raw_decode(b[b.index("{"):])[0]


def carteiras_internas():
    """Fundos internos (painel de fundos): última carteira aberta (CVM, look-through) ou carteiras/<nome>.xlsx se existir."""
    D = painel_dados()
    snd, sec = snd_caracteristicas(), securitizadoras()
    cad = cadastro_b3().drop_duplicates("cod").set_index("cod")
    A = pd.DataFrame(D["ativos"], columns=["id", "mes", "ch", "rot", "cat", "grupo", "emis", "perc", "rs", "sp", "du", "x"])
    out = {}
    for f in D["fundos"]:
        if not f["nosso"]:
            continue
        m = f["abertos"][-1] if f["abertos"] else None
        manual = PASTA / "carteiras" / f"{f['nome']}.xlsx"
        if manual.exists():
            c = le_carteira(manual).assign(cat="manual", rotulo=lambda d: d.entrada)
            fonte, m = "planilha", pd.Timestamp(manual.stat().st_mtime, unit="s").strftime("%Y-%m")
        else:
            a = A[(A.id == f["id"]) & (A.mes == m)].groupby(["ch", "rot", "cat", "emis"], as_index=False, dropna=False).perc.sum()
            linhas = []
            for r in a.itertuples():
                cod = str(r.ch).strip().upper()
                venc = None
                if COD_B3.match(cod):
                    if cod in snd.index and pd.notna(snd.venc.get(cod)):
                        venc = pd.Timestamp(snd.venc[cod]).strftime("%Y-%m-%d")
                    elif cod in sec:
                        venc = str(sec[cod]["venc"])[:10]
                    elif cod in cad.index:
                        venc = pd.to_datetime(cad.venc[cod], dayfirst=True, errors="coerce")
                        venc = venc.strftime("%Y-%m-%d") if pd.notna(venc) else None
                linhas.append({"entrada": f"{cod} - {venc}" if venc else str(r.rot), "codigo": cod if venc else None, "venc": venc,
                               "perc_nav": float(r.perc), "cat": r.cat, "rotulo": str(r.rot), "emis_cvm": r.emis})
            c, fonte = pd.DataFrame(linhas, columns=["entrada", "codigo", "venc", "perc_nav", "cat", "rotulo", "emis_cvm"]), "CVM (carteira aberta)"
        cart = f["carteira"][-1] if f.get("carteira") else {}
        out[f["id"]] = {"id": f["id"], "nome": f["nome"], "cnpj": f["cnpj"], "pl": f.get("pl"), "mes": m, "fonte": fonte, "carteira": c,
                        "ret_m": f.get("ret", {}), "cats_cvm": cart.get("cats", {}) if cart else {}}
    return out, D.get("cdi", {})


def fundo_diario(cnpjs, meses=14):
    """Informe diário da CVM dos fundos: cota, PL, captação e resgate (últimos `meses`)."""
    alvo = {re.sub(r"\D", "", c) for c in cnpjs if c}
    partes = []
    for p in pd.period_range(HOJE - pd.DateOffset(months=meses), HOJE, freq="M"):
        f = DADOS / "cvm" / f"inf_diario_fi_{p.strftime("%Y%m")}.zip"
        if not f.exists() or (p >= (HOJE - pd.DateOffset(months=1)).to_period("M") and velho(f, 12)):
            try:
                f.write_bytes(get(f"{CVM.replace('CIA_ABERTA/', '')}FI/DOC/INF_DIARIO/DADOS/inf_diario_fi_{p.strftime("%Y%m")}.zip", timeout=300).content)
            except requests.RequestException:
                pass
        if not f.exists():
            continue
        z = zipfile.ZipFile(f)
        d = pd.read_csv(z.open(z.namelist()[0]), sep=";", encoding="latin1", dtype=str)
        col = "CNPJ_FUNDO_CLASSE" if "CNPJ_FUNDO_CLASSE" in d else "CNPJ_FUNDO"
        d = d[d[col].str.replace(r"\D", "", regex=True).isin(alvo)]
        partes.append(pd.DataFrame({"cnpj": d[col].str.replace(r"\D", "", regex=True).values, "data": pd.to_datetime(d.DT_COMPTC).values,
                                    "cota": pd.to_numeric(d.VL_QUOTA, errors="coerce").values, "pl": pd.to_numeric(d.VL_PATRIM_LIQ, errors="coerce").values,
                                    "capt": pd.to_numeric(d.CAPTC_DIA, errors="coerce").values, "resg": pd.to_numeric(d.RESG_DIA, errors="coerce").values,
                                    "cotistas": pd.to_numeric(d.NR_COTST, errors="coerce").values}))
    if not partes:
        return {}
    x = pd.concat(partes).sort_values(["cnpj", "data"]).drop_duplicates(["cnpj", "data"], keep="last")
    return {c: g.set_index("data") for c, g in x.groupby("cnpj")}


def yield_ativo(x, cdi_aa, ipca_e):
    """Taxa total esperada (% a.a.) do papel pela taxa indicativa: CDI+s, %CDI, IPCA+r, pré."""
    t = x.get("taxa")
    if t is None or t != t:
        return None
    i = x.get("idx")
    if i == "DI+":
        return ((1 + cdi_aa) * (1 + t / 100) - 1) * 100
    if i == "%DI":
        return cdi_aa * t
    if i == "IPCA":
        return ((1 + ipca_e) * (1 + t / 100) - 1) * 100
    if i == "PRE":
        return t
    return None


def cota_fluxo(out, f, fdiario):
    """Cota, retorno contra o CDI e captação líquida do informe diário da CVM."""
    fd = fdiario.get(re.sub(r"\D", "", f["cnpj"] or ""))
    if fd is not None and len(fd) > 20:
        cota = fd.cota.dropna()
        cdi_d = sgs(12, HOJE - pd.DateOffset(years=4)).reindex(cota.index).fillna(0) / 100
        idx_cdi = (1 + cdi_d).cumprod()
        r_ = lambda d0: (float(cota.iloc[-1] / cota[cota.index <= d0].iloc[-1] - 1) * 100, float(idx_cdi.iloc[-1] / idx_cdi[idx_cdi.index <= d0].iloc[-1] - 1) * 100) if (cota.index <= d0).any() else (None, None)
        per = {"mes": cota.index[-1].replace(day=1) - pd.Timedelta(days=1), "ano": pd.Timestamp(cota.index[-1].year - 1, 12, 31),
               "12m": cota.index[-1] - pd.DateOffset(years=1)}
        out["perf"] = {k: {"fundo": a, "cdi": b, "pct_cdi": a / b * 100 if a is not None and b else None} for k, (a, b) in ((k, r_(d)) for k, d in per.items())}
        base = cota[cota.index >= cota.index[-1] - pd.DateOffset(years=1)]
        ic = idx_cdi.reindex(base.index)
        out["cota_serie"] = {"d": [d.strftime("%Y-%m-%d") for d in base.index], "fundo": [round(v / base.iloc[0] * 100, 4) for v in base],
                             "cdi": [round(v / ic.iloc[0] * 100, 4) for v in ic]}
        fl = (fd.capt.fillna(0) - fd.resg.fillna(0))
        plx = fd.pl.dropna()
        out["fluxo"] = {k: float(fl.iloc[-k:].sum() / plx.iloc[-k - 1] * 100) if len(plx) > k else None for k in (5, 20, 60)}
        f6 = fl[fl.index >= fl.index[-1] - pd.DateOffset(months=6)]
        out["fluxo_serie"] = {"d": [d.strftime("%Y-%m-%d") for d in f6.index], "v": [round(v / 1e6, 2) for v in f6]}
        out["pl_atual"] = float(plx.iloc[-1] / 1e6)
        out["cotistas"] = int(fd.cotistas.dropna().iloc[-1]) if fd.cotistas.notna().any() else None
        out["data_cota"] = cota.index[-1].strftime("%Y-%m-%d")
        out["pl"] = out["pl_atual"]


def consolida_fundo(f, A, emis, diarios, idx_c, mkt, td_, fdiario, cdi_aa, ipca_e, ipca_d, cdi_m):
    """Visão de um fundo: alocação, carrego, sensibilidades, risco, liquidez, cenários, sinais, agenda em R$, projeção, cota e fluxo."""
    c = f["carteira"]
    pl = (f.get("pl") or 0) * 1e6                                              # R$
    cod_w = c.dropna(subset=["codigo"]).groupby("codigo").perc_nav.sum()
    Af = A[A.codigo.isin(cod_w.index)].copy()
    Af["perc_nav"] = Af.codigo.map(cod_w)
    Af["coberto"] = True
    resto = c[c.codigo.isna()]
    out = {"id": f["id"], "nome": f["nome"], "cnpj": f["cnpj"], "pl": f.get("pl"), "mes": f["mes"], "fonte": f["fonte"],
           "pesos": {k: round(float(v), 4) for k, v in cod_w.items()},
           "outros": [{"rotulo": r.rotulo, "cat": r.cat, "perc": round(float(r.perc_nav), 3)} for r in resto.sort_values("perc_nav", ascending=False).head(60).itertuples()]}
    # alocação
    cat = c.groupby("cat").perc_nav.sum().sort_values(ascending=False) if len(c) else pd.Series(dtype=float)
    out["aloc_cat"] = {k: round(float(v), 2) for k, v in cat.items()}
    if Af.empty:                                                       # fundo sem crédito com código: só alocação, cota e fluxo
        out.update({"credito_pct": 0.0, "ativos": [], "raizes": [], "risco": None, "hoje": {"impacto_bps": 0, "n": 0, "ativos": []}, "regras": [],
                    "agenda": [], "docs": [], "carrego": {}, "sens": {}, "cenarios": [], "projecao": []})
        cota_fluxo(out, f, fdiario)
        return out
    cred = Af[Af.taxa.notna()]
    out["aloc_idx"] = {k: round(float(v), 2) for k, v in Af.groupby("idx").perc_nav.sum().sort_values(ascending=False).items()}
    Af["emis_nome"] = [(emis.get(r) or {}).get("nome") or e for r, e in zip(Af.raiz, Af.emissor)]
    Af["setor"] = [(emis.get(r) or {}).get("setor") or "—" for r in Af.raiz]
    out["aloc_emissor"] = {str(k)[:40]: round(float(v), 2) for k, v in Af.groupby("emis_nome").perc_nav.sum().sort_values(ascending=False).head(15).items()}
    out["aloc_setor"] = {str(k)[:40]: round(float(v), 2) for k, v in Af.groupby("setor").perc_nav.sum().sort_values(ascending=False).head(12).items()}
    out["aloc_prazo"] = {k: round(float(v), 2) for k, v in Af.groupby(pd.cut(Af.duration, [-1, 1, 3, 5, 100], labels=["até 1 ano", "1–3 anos", "3–5 anos", "mais de 5"]),
                                                                       observed=True).perc_nav.sum().items()}
    # liquidez: dias para vender a posição a 20% do volume médio diário na B3 (30 dias)
    adv = Af.giro_30d.fillna(0) * 1e6 / 21
    pos = Af.perc_nav / 100 * pl
    dias = np.where(adv > 0, pos / (0.2 * adv.replace(0, np.nan)), np.inf)
    Af["dias_liq"] = dias
    out["liquidez"] = {k: round(float(Af.perc_nav[m].sum()), 2) for k, m in (("até 1 dia", dias <= 1), ("1–5 dias", (dias > 1) & (dias <= 5)),
                                                                             ("5–20 dias", (dias > 5) & (dias <= 20)), ("mais de 20 dias", (dias > 20) & np.isfinite(dias)),
                                                                             ("sem negócio em 30 dias", ~np.isfinite(dias)))}
    # carrego e sensibilidades
    y = np.array([yield_ativo(r, cdi_aa, ipca_e) or np.nan for r in Af.to_dict("records")])
    ok = np.isfinite(y)
    wy = Af.perc_nav.values[ok]
    out["carrego"] = {"taxa_media": float((y[ok] * wy).sum() / wy.sum()) if ok.any() else None, "cobre": float(wy.sum()),
                      "cdi": cdi_aa * 100, "pct_cdi": float((y[ok] * wy).sum() / wy.sum() / (cdi_aa * 100) * 100) if ok.any() else None,
                      "spread_di": float((cred[cred.idx == "DI+"].taxa * cred[cred.idx == "DI+"].perc_nav).sum() / max(cred[cred.idx == "DI+"].perc_nav.sum(), 1e-9)) if (cred.idx == "DI+").any() else None,
                      "taxa_ipca": float((cred[cred.idx == "IPCA"].taxa * cred[cred.idx == "IPCA"].perc_nav).sum() / max(cred[cred.idx == "IPCA"].perc_nav.sum(), 1e-9)) if (cred.idx == "IPCA").any() else None}
    dm = np.where(Af.idx.isin(["IPCA", "PRE", "IGPM"]), Af.duration / (1 + Af.taxa.fillna(0) / 100), Af.duration).astype(float)
    Af["dm"] = dm
    w = Af.perc_nav.values / 100
    out["sens"] = {"duration_spread": float(np.nansum(w * dm) / max(w.sum(), 1e-9)), "cs01": float(np.nansum(w * dm) * pl * 1e-4),
                   "dv01_ipca": float(np.nansum(w * dm * (Af.idx == "IPCA")) * pl * 1e-4), "dv01_pre": float(np.nansum(w * dm * (Af.idx == "PRE")) * pl * 1e-4)}
    emi_w = Af.groupby("raiz").perc_nav.sum().sort_values(ascending=False)
    maior = emi_w.index[0] if len(emi_w) else None
    out["cenarios"] = [
        {"nome": "Spread de crédito +100 bps (todos os papéis)", "pct": -float(np.nansum(w * dm)), "rs": -float(np.nansum(w * dm) * pl / 100)},
        {"nome": "Spread +50 bps só nos papéis DI+", "pct": -float(np.nansum(w * dm * (Af.idx == "DI+"))) / 2, "rs": -float(np.nansum(w * dm * (Af.idx == "DI+")) * pl / 200)},
        {"nome": "Juro real (NTN-B) +100 bps", "pct": -float(np.nansum(w * dm * (Af.idx == "IPCA"))), "rs": -float(np.nansum(w * dm * (Af.idx == "IPCA")) * pl / 100)},
        {"nome": "Juro pré +100 bps", "pct": -float(np.nansum(w * dm * (Af.idx == "PRE"))), "rs": -float(np.nansum(w * dm * (Af.idx == "PRE")) * pl / 100)},
        {"nome": f"Default do maior emissor ({str((emis.get(maior) or {}).get('nome') or maior)[:30]}), recuperação 40%", "pct": -float(emi_w.iloc[0] * 0.6) if len(emi_w) else 0,
         "rs": -float(emi_w.iloc[0] * 0.6 * pl / 100) if len(emi_w) else 0}]
    # risco (VaR, backtest, fatores, contribuição)
    Ar = Af.copy()
    out["risco"] = risco(Ar, diarios, idx_c, mkt, td_) if len(Ar) and Ar.vol.notna().any() else None
    Af["contrib_pct"] = Ar.get("contrib_pct")
    # hoje (marcação esperada) e regras / modelo 5 dias
    ph = [(r.codigo, r.prev_hoje, r.perc_nav) for r in Af.itertuples() if isinstance(r.prev_hoje, dict)]
    out["hoje"] = {"impacto_bps": float(sum(p["total"] * wv / 100 for _, p, wv in ph)), "n": len(ph),
                   "ativos": sorted([{"cod": c_, "bps": p["total"], "imp": p["total"] * wv / 100, "perc": wv,
                                      "txt": "; ".join(f"{k}: {v['txt']}" for k, v in p["comp"].items())} for c_, p, wv in ph], key=lambda d: -abs(d["imp"]))[:25]}
    regras = [{"cod": r.codigo, "perc": r.perc_nav, **g} for r in Af.itertuples() if isinstance(r.regras_dias, list) for g in r.regras_dias]
    out["regras"] = regras
    p5 = Af[Af.prev5.map(lambda p: isinstance(p, dict))]
    out["p5_fracos"] = [{"cod": r.codigo, "perc": r.perc_nav, "pct": r.prev5["pct_univ"], "mu": r.prev5["mu"], "q08": r.prev5["q08"], "q92": r.prev5["q92"]}
                        for r in p5.sort_values("perc_nav", ascending=False).itertuples() if r.prev5["pct_univ"] <= 10][:15]
    out["p5_fortes"] = [{"cod": r.codigo, "perc": r.perc_nav, "pct": r.prev5["pct_univ"], "mu": r.prev5["mu"], "q08": r.prev5["q08"], "q92": r.prev5["q92"]}
                        for r in p5.sort_values("perc_nav", ascending=False).itertuples() if r.prev5["pct_univ"] >= 90][:15]
    # projeção do fundo contra o carrego (bps do PL): mediana dos sinais (regra > modelo) e faixa pela vol do fundo
    sig = {h: 0.0 for h in (1, 5, 10, 20)}
    for r in Af.itertuples():
        wv = r.perc_nav / 100
        rg = r.regras_dias[0]["regra"] if isinstance(r.regras_dias, list) and r.regras_dias else None
        for h in sig:
            if rg and str(h) in REGRAS_DIAS[rg]["h"]:
                sig[h] += wv * REGRAS_DIAS[rg]["h"][str(h)]["mediana"]
            elif isinstance(r.prev5, dict):
                sig[h] += wv * r.prev5["mu"] * min(h, 5) / 5
            if h == 1 and isinstance(r.prev_hoje, dict):
                sig[h] += wv * r.prev_hoje["total"]
    sd_d = (out["risco"]["vol"] / np.sqrt(252) * 100) if out["risco"] else None           # bps do PL por dia
    tx = out["carrego"]["taxa_media"]
    out["projecao"] = [{"h": h, "sinais": sig[h], "carrego_vs_cdi": ((tx or cdi_aa * 100) - cdi_aa * 100) * 100 * h / 252 * (out["carrego"]["cobre"] / 100),
                        "lo": sig[h] - 1.2816 * sd_d * np.sqrt(h) if sd_d else None, "hi": sig[h] + 1.2816 * sd_d * np.sqrt(h) if sd_d else None} for h in (1, 5, 10, 20)]
    # PnL 12m (bps do PL, pesos de hoje)
    out["pnl12"] = {k: float(np.nansum(Af[f"pnl_{k}"].astype(float) * Af.perc_nav / 100)) for k in ("car", "mtm", "res")}
    dts, acc = None, {}
    for r in Af.itertuples():
        h_ = r.hist if isinstance(r.hist, dict) else {}
        if not h_.get("d"):
            continue
        s_ = pd.DataFrame({k: pd.Series(h_[k], index=pd.to_datetime(h_["d"]), dtype=float) for k in ("car", "mtm", "res")})
        s_ = s_[s_.index > HOJE - pd.Timedelta(days=365)].fillna(0) * r.perc_nav / 100
        for k in s_:
            acc[k] = acc[k].add(s_[k], fill_value=0) if k in acc else s_[k]
    if acc:
        P = pd.DataFrame(acc).sort_index().cumsum()
        P = P.iloc[::2]
        out["pnl_serie"] = {"d": [d.strftime("%Y-%m-%d") for d in P.index], **{k: [round(v, 1) for v in P[k]] for k in P}}
    # agenda em R$ (amortização/vencimento estimados pela % do valor nominal × posição)
    ag = []
    for r in Af.itertuples():
        for e in (r.agenda if isinstance(r.agenda, list) else []):
            pct = num([str(e.get("pct", "")).replace("%", "").strip()]).iloc[0]
            amort = bool(re.search(r"Amortiza|Vencimento|Resgate", e["evento"], re.I))
            ag.append({"data": pd.to_datetime(e["data"], dayfirst=True).strftime("%Y-%m-%d"), "cod": r.codigo, "evento": e["evento"], "pct": e.get("pct"),
                       "rs": float(pct / 100 * r.perc_nav / 100 * pl / ((r.pct_par or 100) / 100)) / 1e6 if amort and pct == pct else None, "perc": r.perc_nav})
        if r.venc and HOJE <= pd.Timestamp(r.venc) <= HOJE + pd.Timedelta(days=180):
            ag.append({"data": r.venc, "cod": r.codigo, "evento": "Vencimento", "pct": "100%", "rs": float(r.perc_nav / 100 * pl) / 1e6, "perc": r.perc_nav})
    for rz in set(Af.raiz.dropna()):
        e = emis.get(rz) or {}
        if e.get("prox_resultado"):
            pr = e["prox_resultado"]
            ag.append({"data": pr["esperado"], "cod": ", ".join(Af[Af.raiz == rz].codigo.head(3)), "evento": f"Resultado {pr['ref'][5:7]}/{pr['ref'][:4]} (estimado {pr['min'][8:]}/{pr['min'][5:7]}–{pr['max'][8:]}/{pr['max'][5:7]})",
                       "pct": "", "rs": None, "perc": float(Af[Af.raiz == rz].perc_nav.sum())})
    ag += [{"data": d.strftime("%Y-%m-%d"), "cod": "IPCA", "evento": "Divulgação do IPCA (IBGE)", "pct": "", "rs": None, "perc": float(Af[Af.idx == "IPCA"].perc_nav.sum())} for d in ipca_d[:3]]
    out["agenda"] = sorted([a for a in ag if HOJE.strftime("%Y-%m-%d") <= a["data"] <= (HOJE + pd.Timedelta(days=120)).strftime("%Y-%m-%d")], key=lambda a: a["data"])
    # documentos e notícias dos emissores do fundo (7 dias = alerta)
    relev = lambda d: d["cat"] in ("Fato Relevante", "Aviso aos Debenturistas", "Escrituras e aditamentos de debêntures",
                                   "Informações de Companhias em Recuperação Judicial ou Extrajudicial") or re.search(r"rating|waiver|resgate|recompra|vencimento antecipado|AGD|reestrutura", d.get("assunto") or "", re.I)
    vist = set()
    out["docs"] = [d for d in sorted([{**d, "emissor": (emis.get(rz) or {}).get("nome"), "perc": float(Af[Af.raiz == rz].perc_nav.sum())}
                                      for rz in set(Af.raiz.dropna()) for d in (emis.get(rz) or {}).get("docs", []) if relev(d)], key=lambda d: d["data"], reverse=True)
                   if not ((d["emissor"], d["assunto"]) in vist or vist.add((d["emissor"], d["assunto"])))][:25]
    out["raizes"] = sorted(set(Af.raiz.dropna()))
    cota_fluxo(out, f, fdiario)
    # tabela de ativos do fundo (só o que é do fundo; o resto do ativo vem de D.ativos)
    out["ativos"] = [{"cod": r.codigo, "perc": round(float(r.perc_nav), 4), "rs": round(float(r.perc_nav / 100 * (out["pl"] or 0)), 2),
                      "contrib": None if not isinstance(r.contrib_pct, float) or r.contrib_pct != r.contrib_pct else round(r.contrib_pct, 2),
                      "dias_liq": None if not np.isfinite(r.dias_liq) else round(float(r.dias_liq), 1), "setor": r.setor, "emis": r.emis_nome}
                     for r in Af.sort_values("perc_nav", ascending=False).itertuples()]
    out["credito_pct"] = float(Af.perc_nav.sum())
    return out


# ================================================================== negócios do dia ranqueados pela diferença de taxa
def negocios_dia(A, fundos_pesos):
    """Negócios B3 de crédito do último pregão com negócio, contra a taxa indicativa ANBIMA do dia útil anterior.
    Devolve a lista de negócios (ranqueável pela diferença em bps) e, por papel, a curva ANBIMA e os negócios de 30 dias (bolhas)."""
    an = anbima_diario()
    dia = HOJE
    for k in range(6):
        nh = negocios([dia])
        nh = nh[(nh.taxa > 0) & (nh.vol >= 10_000)] if len(nh) else nh
        if len(nh) > 50:
            break
        dia = (dia - pd.offsets.BDay(1)).normalize()
    if not len(nh):
        return {"dia": None, "negocios": [], "papeis": {}}
    # referência: ANBIMA do dia útil anterior (arquivo diário; CRI/CRA pela série do monitor)
    ant = an[an.data < dia]
    ref = ant[ant.data == ant.data.max()].drop_duplicates("codigo").set_index("codigo") if len(ant) else pd.DataFrame()
    refA = {}
    for r in A[A.coberto].itertuples():
        h = r.hist if isinstance(r.hist, dict) else {}
        d_, t_ = h.get("d", []), h.get("tA", [])
        i = max((j for j, d in enumerate(d_) if d < dia.strftime("%Y-%m-%d") and t_[j] is not None), default=None)
        if i is not None:
            refA[r.codigo] = (t_[i], d_[i], r.duration, r.idx)
    snd, cad = snd_caracteristicas(), cadastro_b3().drop_duplicates("cod").set_index("cod")
    ix = indexadores(list(nh.cod.unique()))
    linhas = []
    for r in nh.itertuples():
        if len(ref) and r.cod in ref.index and pd.notna(ref.taxa_ind.get(r.cod)):
            tr, dr, du_ = float(ref.taxa_ind[r.cod]), ref.data.iloc[0].strftime("%Y-%m-%d"), float(ref.duration_anos.get(r.cod) or np.nan)
        elif r.cod in refA:
            tr, dr, du_, _ = refA[r.cod]
        else:
            continue
        dif = (r.taxa - tr) * 100
        if abs(dif) > 1500:
            continue
        linhas.append({"cod": r.cod, "hora": "" if pd.isna(getattr(r, "hora", None)) else str(r.hora)[:5], "vol": float(r.vol), "taxa": float(r.taxa), "pu": float(r.pu), "ref": tr,
                       "ref_data": dr, "dif": float(dif), "idx": ix.get(r.cod), "dur": du_ if du_ == du_ else None, "tipo": r.tipo,
                       "emissor": str(snd.empresa.get(r.cod) if r.cod in snd.index else cad.emissor.get(r.cod, "") or "")[:40]})
    N = pd.DataFrame(linhas)
    if not len(N):
        return {"dia": dia.strftime("%Y-%m-%d"), "negocios": [], "papeis": {}}
    # por papel no dia: volume e diferença ponderada → movimento esperado da ANBIMA (fração do gap pelo volume do dia, estudo)
    g = N.assign(dv=N.dif * N.vol).groupby("cod").agg(dv=("dv", "sum"), vol=("vol", "sum"), n=("vol", "size"))
    g["dif"] = g.dv / g.vol
    g["frac"], g["acerto"] = zip(*[next((f, h) for lim, f, h in FRACAO_GAP if v < lim) for v in g.vol])
    g["esperado"] = g.frac * g.dif
    N = N.join(g[["vol", "dif", "esperado", "acerto", "n"]].rename(columns={"vol": "vol_dia", "dif": "dif_dia", "n": "n_dia"}), on="cod")
    emis_cod = N.groupby("cod").emissor.first().to_dict()
    # papéis no ranking: curva ANBIMA (arquivo diário + série do monitor) e negócios de 30 dias, um por bolha
    cods = list(dict.fromkeys(N.reindex(N.dif.abs().sort_values(ascending=False).index).cod))[:400]
    h30 = negocios(list(pd.bdate_range(dia - pd.Timedelta(days=45), dia)), set(cods))
    h30 = h30[(h30.taxa > 0) & (h30.vol >= 10_000)]
    papeis = {}
    for c in cods:
        a = an[an.codigo.eq(c)].dropna(subset=["taxa_ind"]).sort_values("data")
        curva = {"d": [d.strftime("%Y-%m-%d") for d in a.data], "t": [round(float(v), 4) for v in a.taxa_ind]}
        if c in A.codigo.values:
            h = A.loc[A.codigo.eq(c), "hist"].iloc[0]
            if isinstance(h, dict) and h.get("d"):
                corte = (dia - pd.Timedelta(days=60)).strftime("%Y-%m-%d")
                par = [(d, t) for d, t in zip(h["d"], h["tA"]) if d >= corte and t is not None and d not in curva["d"]]
                curva = {"d": sorted([d for d, _ in par] + curva["d"]), "t": [t for _, t in sorted(par + list(zip(curva["d"], curva["t"])))]}
        x = h30[h30.cod.eq(c)]
        papeis[c] = {"emissor": emis_cod.get(c), "curva": curva, "neg": {"d": [d.strftime("%Y-%m-%d") for d in x.data], "h": ["" if pd.isna(v) else str(v)[:5] for v in (x["hora"] if "hora" in x else [None] * len(x))],
                                             "t": [round(float(v), 4) for v in x.taxa], "v": [round(float(v) / 1e6, 3) for v in x.vol]}}
    N = N.sort_values("dif", key=lambda v: -v.abs())
    return {"dia": dia.strftime("%Y-%m-%d"), "ref_data": ref.data.iloc[0].strftime("%Y-%m-%d") if len(ref) else None,
            "negocios": json.loads(N.drop(columns=["emissor", "pu", "dur", "ref_data", "tipo"]).to_json(orient="records", double_precision=4)), "papeis": papeis,
            "estudo": [{"faixa": "≥ R$ 20 mi", "frac": 0.39, "acerto": 82}, {"faixa": "R$ 5–20 mi", "frac": 0.30, "acerto": 73},
                       {"faixa": "R$ 1–5 mi", "frac": 0.07, "acerto": 58}, {"faixa": "< R$ 1 mi", "frac": 0.0, "acerto": 52}]}


# ================================================================== página e execução
CAMPOS_ATIVO = ["codigo", "emissor", "raiz", "idx", "indexador", "cupom", "venc", "tipo", "lastro", "duration", "taxa", "data_taxa", "pct_par", "pu", "vol",
                "fonte", "dtaxa_1m", "ret_1m", "pnl_car", "pnl_mtm", "pnl_res", "pnl_total", "giro_30d", "neg_30d", "hist", "agenda", "prev_hoje", "prev5", "regras_dias"]


def pagina(F, A, emis, idx_c, mkt, td_, ctx):
    limpa = lambda o: None if isinstance(o, float) and not np.isfinite(o) else o
    corte = (HOJE - pd.DateOffset(years=2)).strftime("%Y-%m-%d")
    at = {}
    for r in A[A.coberto].to_dict("records"):
        x = {k: limpa(r.get(k)) for k in CAMPOS_ATIVO}
        h = x.get("hist") or {}
        if h.get("d"):
            i0 = next((i for i, d in enumerate(h["d"]) if d >= corte), len(h["d"]))
            x["hist"] = {k: (v[i0:] if isinstance(v, list) else v) for k, v in h.items() if k != "neg"}
            ng = h.get("neg") or {"d": []}
            j0 = next((i for i, d in enumerate(ng["d"]) if d >= corte), len(ng["d"]))
            x["hist"]["neg"] = {k: v[j0:] for k, v in ng.items()}
        at[r["codigo"]] = x
    m_ = {}
    for nome, y in mkt.items():
        if y and len(y["serie"]) > 30:
            s_ = y["serie"]
            lr = np.log(s_).diff().dropna()
            d1 = (y["preco"] / y["ant"] - 1) * 100 if y["preco"] and y["ant"] else None
            z = d1 / (lr.tail(250).std() * 100) if d1 is not None else None
            if z is not None and abs(z) > 8:
                d1 = z = None
            m_[nome] = {"preco": y["preco"], "d1": d1, "z": z, "m1": float((s_.iloc[-1] / s_.iloc[-22] - 1) * 100)}
    ic = {}
    for nome, d in idx_c.items():
        s_ = d["serie"]
        if len(s_) > 30:
            lr = np.log(s_).diff().dropna()
            ic[nome] = {"d1": float(lr.iloc[-1] * 100), "z": float(lr.iloc[-1] / lr.tail(250).std()) if lr.tail(250).std() > 0 else None,
                        "m1": float((s_.iloc[-1] / s_[s_.index <= s_.index[-1] - pd.Timedelta(days=30)].iloc[-1] - 1) * 100),
                        "a12": float((np.prod(1 + d["mensal"].tail(12)) - 1) * 100) if len(d["mensal"]) >= 12 else None,
                        "data": s_.index[-1].strftime("%Y-%m-%d"), "d": [q.strftime("%Y-%m-%d") for q in s_.index], "v": [round(v, 5) for v in s_.values]}
    raizes = {r for f in F for r in f.get("raizes", [])}
    em = {r: e for r, e in emis.items() if r in raizes}
    dados = {"gerado": datetime.now().strftime("%d/%m/%Y %H:%M"), "hoje": HOJE.strftime("%Y-%m-%d"), "fundos": F, "ativos": at, "emissores": em,
             "mercado": m_, "indices": ic, "regras_dias": REGRAS_DIAS, **ctx}
    txt = json.dumps(dados, ensure_ascii=False, default=lambda o: None if isinstance(o, float) and not np.isfinite(o) else (o.isoformat() if hasattr(o, "isoformat") else str(o)))
    txt = re.sub(r"\bNaN\b|\b-?Infinity\b", "null", txt).replace("</", "<\\/")
    return HTML.replace("/*DADOS*/", txt)


def uma_vez(arq=None):
    global HOJE
    HOJE = pd.Timestamp.today().normalize()
    t0 = time.time()
    if arq:
        c = le_carteira(arq).assign(cat="planilha", rotulo=lambda d: d.entrada)
        fundos = {"planilha": {"id": "planilha", "nome": Path(arq).stem, "cnpj": None, "pl": None, "mes": HOJE.strftime("%Y-%m"), "fonte": "planilha", "carteira": c}}
    else:
        fundos, _ = carteiras_internas()
    U = pd.concat([f["carteira"].dropna(subset=["codigo"]) for f in fundos.values()])
    U = U.groupby("codigo", as_index=False).agg(venc=("venc", "first"), perc_nav=("perc_nav", "max"), entrada=("entrada", "first"))
    log(f"fundos: {len(fundos)}; ativos com código (união): {len(U)}")
    A, emis, prev, idx_c, mkt, td_, diarios = calcula(U)
    cdi = sgs(12, HOJE - pd.DateOffset(years=4))
    cdi_aa = float((1 + cdi.iloc[-1] / 100) ** 252 - 1) if len(cdi) else 0.14
    fo = focus()
    ipca_e = float(list((fo.get("IPCA") or {}).get("anos", {}).values())[1]) / 100 if len((fo.get("IPCA") or {}).get("anos", {})) > 1 else 0.045
    ipca_d = ipca_datas()
    fd = fundo_diario([f["cnpj"] for f in fundos.values()])
    F = []
    for f in fundos.values():
        try:
            F.append(consolida_fundo(f, A, emis, diarios, idx_c, mkt, td_, fd, cdi_aa, ipca_e, ipca_d, None))
        except Exception as e:
            log(f"fundo {f['nome']}: erro {type(e).__name__} {e}")
    F.sort(key=lambda x: -(x.get("credito_pct") or 0))
    news = {}
    peso = pd.Series({r: sum((x.get("pl") or 0) * sum(w for c_, w in x["pesos"].items() if (A.set_index("codigo").raiz.get(c_) == r)) for x in F) for r in {r for x in F for r in x.get("raizes", [])}})
    for r in peso.sort_values(ascending=False).index[:30]:
        e = emis.get(r) or {}
        if e.get("nome") and not re.search(r"SECURITIZ", str(e["nome"]).upper()):
            q, it = noticias(e["nome"])
            if it:
                news[r] = {"busca": q, "itens": it}
    try:
        neg_dia = negocios_dia(A, {x["nome"]: x["pesos"] for x in F if x.get("pesos")})
    except Exception as e:
        log(f"negócios do dia: erro {type(e).__name__} {e}")
        neg_dia = {"dia": None, "negocios": [], "papeis": {}}
    log(f"negócios do dia {neg_dia['dia']}: {len(neg_dia['negocios'])} com referência ANBIMA")
    ctx = {"negocios_dia": neg_dia, "juros": A.attrs.get("juros") or juros_b3(), "focus": fo, "noticias": news, "info5": A.attrs.get("info5"), "data_ref": A.attrs.get("data_ref")}
    h = pagina(F, A, emis, idx_c, mkt, td_, ctx)
    (SAIDA / "vol_monitor.html").write_text(h, encoding="utf-8")
    log(f"saida/vol_monitor.html {len(h) / 1e6:.1f} MB, {len(F)} fundos ({time.time() - t0:.0f}s)")


def rodar(arq=None):
    while True:
        try:
            uma_vez(arq)
        except Exception as e:
            log(f"erro: {type(e).__name__} {e}")
        hr = datetime.now().hour + datetime.now().minute / 60
        time.sleep(900 if 9 <= hr <= 18.5 and datetime.now().weekday() < 5 else 3600)


HTML = r"""<title>Monitor de crédito</title>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<link rel="preconnect" href="https://fonts.googleapis.com"><link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/plotly.js/3.1.0/plotly.min.js"></script>
<script>if(!window.Plotly)document.write('<script src="https://cdn.jsdelivr.net/npm/plotly.js-dist-min@3.1.0/plotly.min.js"><\/script>')</script>
<style>
:root{--bg:#f5f6f7;--card:#fff;--ink:#121820;--mut:#5b6672;--line:#e1e5e9;--acc:#0b6e66;--acc2:#dff0ed;--up:#13803c;--dn:#b42318;--warn:#a85b06;--warn2:#fdf0dc;--crit2:#fde6e3;--chip:#eef1f4;
 --s1:#0b6e66;--s2:#c2410c;--s3:#5b4bd1;--s4:#7c8794;--band:rgba(11,110,102,.16);color-scheme:light}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#0f1316;--card:#171c20;--ink:#e6e9ec;--mut:#9aa4ae;--line:#283037;--acc:#3fc0ad;--acc2:#16332e;--up:#47c47c;--dn:#f0675b;--warn:#e3a33f;--warn2:#3a2c14;--crit2:#3c1c19;--chip:#232a30;
 --s1:#3fc0ad;--s2:#f08a4b;--s3:#a597ff;--s4:#93a0ad;--band:rgba(63,192,173,.18);color-scheme:dark}}
:root[data-theme="dark"]{--bg:#0f1316;--card:#171c20;--ink:#e6e9ec;--mut:#9aa4ae;--line:#283037;--acc:#3fc0ad;--acc2:#16332e;--up:#47c47c;--dn:#f0675b;--warn:#e3a33f;--warn2:#3a2c14;--crit2:#3c1c19;--chip:#232a30;
 --s1:#3fc0ad;--s2:#f08a4b;--s3:#a597ff;--s4:#93a0ad;--band:rgba(63,192,173,.18);color-scheme:dark}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.45 "IBM Plex Sans",system-ui,sans-serif}
.w{max-width:1360px;margin:0 auto;padding-inline:16px;padding-block:12px 48px}
.top{display:flex;flex-wrap:wrap;gap:8px 14px;align-items:center;margin-bottom:6px}.top h1{font-size:19px;margin:0;font-weight:600}
.top select{font:inherit;font-weight:600;font-size:15px;padding:6px 10px;border-radius:7px;border:1px solid var(--line);background:var(--card);color:var(--ink);max-width:100%}
.mut{color:var(--mut)}.nota{font-size:12px;color:var(--mut);margin:6px 0 0}td,.kv b,.num{font-variant-numeric:tabular-nums}
.aviso{background:var(--warn2);color:var(--warn);border-radius:6px;padding:6px 10px;font-size:13px;margin:6px 0}
nav{display:flex;gap:2px;flex-wrap:wrap;position:sticky;top:env(safe-area-inset-top,0px);z-index:9;background:var(--bg);border-bottom:1px solid var(--line);margin-bottom:12px}
nav button{border:0;border-bottom:2px solid transparent;background:none;color:var(--mut);padding:9px 12px;font:inherit;font-weight:500;cursor:pointer}
nav button.on{color:var(--ink);border-bottom-color:var(--acc)}
.kp{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px;margin-bottom:12px}
.kv{background:var(--card);border:1px solid var(--line);border-radius:9px;padding:10px 12px}.kv span{display:block;color:var(--mut);font-size:12px}.kv b{display:block;font-size:20px;font-weight:600;margin:2px 0}.kv i{font-style:normal;font-size:12px;color:var(--mut)}
section{background:var(--card);border:1px solid var(--line);border-radius:9px;padding:12px 14px;margin-bottom:12px;min-width:0}
h2{font-size:15px;margin:0 0 4px;font-weight:600}.sub{font-size:12px;color:var(--mut);margin:0 0 8px}h3{font-size:13px;margin:12px 0 6px;font-weight:600}
.g2{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:12px}.g3{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px}
@media(max-width:920px){.g2,.g3{grid-template-columns:minmax(0,1fr)}}
.sc{overflow-x:auto}table{border-collapse:collapse;width:100%;font-size:13px}th,td{padding:5px 7px;border-bottom:1px solid var(--line);text-align:right;white-space:nowrap}
th{color:var(--mut);font-weight:500;font-size:12px;cursor:pointer;position:sticky;top:0;background:var(--card)}th.l,td.l{text-align:left}td.wrap{white-space:normal;min-width:220px}
tr.cl{cursor:pointer}tr.cl:hover td{background:var(--chip)}
.up{color:var(--up)}.dn{color:var(--dn)}.tag{display:inline-block;border-radius:4px;padding:1px 6px;font-size:11px;font-weight:600}.t-dn{background:var(--crit2);color:var(--dn)}.t-up{background:var(--acc2);color:var(--acc)}.t-w{background:var(--warn2);color:var(--warn)}
.pl{height:290px}.pl.s{height:230px}.pl.l{height:360px}.ctl{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin-bottom:8px}
.ctl input,.ctl select{font:inherit;background:var(--card);color:var(--ink);border:1px solid var(--line);border-radius:6px;padding:5px 8px;max-width:100%}
.seg button{border:1px solid var(--line);background:var(--card);color:var(--mut);padding:4px 9px;font:inherit;font-size:12px;cursor:pointer}.seg button.on{background:var(--acc);border-color:var(--acc);color:#fff}
.seg button:first-child{border-radius:6px 0 0 6px}.seg button:last-child{border-radius:0 6px 6px 0}
.vazio{color:var(--mut);font-style:italic}ul.ls{list-style:none;padding:0;margin:0}ul.ls li{padding:6px 0;border-bottom:1px solid var(--line)}a{color:var(--acc)}
.leia{font-size:13px;line-height:1.5;margin:4px 0 0}.leia b{font-weight:600}
.bar{display:flex;align-items:center;gap:8px;margin:3px 0}.bar span:first-child{flex:0 0 44%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12.5px}.bar i{display:block;height:12px;border-radius:3px;background:var(--s1)}.bar em{font-style:normal;font-size:12px;color:var(--mut)}
[hidden]{display:none!important}
</style>
<div class="w">
<div class="top"><h1>Monitor de crédito</h1><select id="fundo" aria-label="fundo"></select><span class="mut" id="ger"></span><button id="tema" class="seg" style="margin-left:auto;border:1px solid var(--line);background:var(--card);color:var(--mut);border-radius:6px;padding:4px 9px;cursor:pointer">tema</button></div>
<div id="avisos"></div>
<nav id="nav"></nav>
<div id="t_geral"></div><div id="t_sinais" hidden></div><div id="t_negocios" hidden></div><div id="t_ativos" hidden></div><div id="t_risco" hidden></div><div id="t_agenda" hidden></div><div id="t_mercado" hidden></div>
</div>
<script>
const D=/*DADOS*/;
const $=s=>document.querySelector(s),css=v=>getComputedStyle(document.documentElement).getPropertyValue(v).trim();
const br=(v,c=2)=>v==null||!isFinite(v)?"—":Number(v).toLocaleString("pt-BR",{minimumFractionDigits:c,maximumFractionDigits:c});
const sg=(v,c=1,u="")=>v==null||!isFinite(v)?"—":`<span class="${v>0?"up":v<0?"dn":""}">${v>0?"+":""}${br(v,c)}${u}</span>`;
const esc=s=>String(s??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const dt=s=>s?s.slice(8,10)+"/"+s.slice(5,7)+"/"+s.slice(2,4):"—";
const mi=v=>v==null?"—":Math.abs(v)>=1e3?br(v/1e3,2)+" bi":br(v,1)+" mi";
const lay=(o={})=>Object.assign({margin:{l:46,r:12,t:8,b:32},paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"rgba(0,0,0,0)",font:{family:"IBM Plex Sans",size:11,color:css("--ink")},
  xaxis:{gridcolor:css("--line"),zerolinecolor:css("--line")},yaxis:{gridcolor:css("--line"),zerolinecolor:css("--mut")},showlegend:true,legend:{orientation:"h",y:1.13,x:0,font:{size:11}},hovermode:"x unified"},o);
const plot=(id,t,l)=>{if(window.Plotly&&document.getElementById(id))Plotly.react(id,t,lay(l),{displayModeBar:false,responsive:true})};
function tabela(el,cols,rows,o={}){let k=o.ord||null,dir=o.asc?1:-1;const go=()=>{let r=rows.slice();if(k)r.sort((a,b)=>{const x=a[k],y=b[k];return (x==null)-(y==null)||(typeof x=="string"?x.localeCompare(y):(x-y))*dir});
  el.innerHTML=`<div class="sc"${o.h?` style="max-height:${o.h}px;overflow:auto"`:""}><table><tr>${cols.map(c=>`<th class="${c[2]||""}" data-k="${c[0]}">${c[1]}</th>`).join("")}</tr>${r.map(x=>`<tr class="${o.click?"cl":""}" data-i="${rows.indexOf(x)}">${cols.map(c=>`<td class="${c[2]||""}">${c[3]?c[3](x[c[0]],x):esc(x[c[0]]??"—")}</td>`).join("")}</tr>`).join("")}</table></div>`;
  el.querySelectorAll("th").forEach(th=>th.onclick=()=>{dir=k==th.dataset.k?-dir:-1;k=th.dataset.k;go()});if(o.click)el.querySelectorAll("tr.cl").forEach(tr=>tr.onclick=()=>o.click(rows[+tr.dataset.i]))};go()}
const barras=(el,obj,u="%",n=12)=>{const e=Object.entries(obj||{}).filter(([,v])=>v!=null).slice(0,n),m=Math.max(...e.map(([,v])=>Math.abs(v)),1e-9);
  el.innerHTML=e.length?e.map(([k,v])=>`<div class="bar"><span title="${esc(k)}">${esc(k)}</span><i style="width:${Math.max(2,Math.abs(v)/m*46)}%;${v<0?"background:var(--s2)":""}"></i><em>${br(v,1)}${u}</em></div>`).join(""):'<p class="vazio">—</p>'};
// ---------------------------------------------------------------- estado
const F=D.fundos,AT=D.ativos,E=D.emissores,RD=D.regras_dias;
let FS=null,SEL=null;
const ABAS=[["geral","Visão geral"],["sinais","Sinais"],["negocios","Negócios"],["ativos","Ativos"],["risco","Risco"],["agenda","Agenda"],["mercado","Mercado"]];
$("#nav").innerHTML=ABAS.map(([k,t])=>`<button data-a="${k}">${t}</button>`).join("");
function aba(id){ABAS.forEach(([k])=>{$("#t_"+k).hidden=k!=id;$(`#nav [data-a="${k}"]`).classList.toggle("on",k==id)});try{localStorage.setItem("mc_aba",id)}catch(e){}
  if(window.Plotly)setTimeout(()=>document.querySelectorAll(`#t_${id} .js-plotly-plot`).forEach(g=>Plotly.Plots.resize(g)),40)}
$("#nav").onclick=e=>{if(e.target.dataset.a)aba(e.target.dataset.a)};
$("#fundo").innerHTML=F.map((f,i)=>`<option value="${i}">${esc(f.nome)}${f.credito_pct<5?" (pouco crédito)":""}</option>`).join("");
$("#fundo").onchange=()=>{FS=F[+$("#fundo").value];SEL=null;try{localStorage.setItem("mc_fundo",FS.id)}catch(e){};tudo()};
$("#tema").onclick=()=>{const r=document.documentElement,d=r.dataset.theme=="dark"||(!r.dataset.theme&&matchMedia("(prefers-color-scheme:dark)").matches);r.dataset.theme=d?"light":"dark";tudo()};
$("#ger").textContent=`atualizado ${D.gerado}`;
const A=c=>AT[c]||{};
const dmod=a=>["IPCA","PRE","IGPM"].includes(a.idx)?(a.duration||1)/(1+(a.taxa||0)/100):(a.duration||1);
const sgt=(v,c=0)=>v==null||!isFinite(v)?"—":`<span class="${v>0?"dn":v<0?"up":""}">${v>0?"+":""}${br(v,c)}</span>`;
const pf=c=>FS.pesos[c]||0;
// projeção de um ativo: mediana e faixa 10–90% em bps de % do PU par para 0,1,5,10,20 dias úteis
function proj(a){const sd=(a.vol||0)/Math.sqrt(252)*100,hs=[0,1,5,10,20],rg=(a.regras_dias||[])[0],R=rg?RD[rg.regra]:null;
  return hs.map(h=>{if(!h)return{h,med:0,lo:0,hi:0,fonte:""};let med=0,lo=-1.2816*sd*Math.sqrt(h),hi=-lo,fonte="vol do papel";
    if(a.prev5){const k=Math.sqrt(h/5),c=a.prev5.mu*Math.min(h,5)/5;med=c;lo=c+(a.prev5.q08-a.prev5.mu)*k;hi=c+(a.prev5.q92-a.prev5.mu)*k;fonte="modelo 5 dias"}
    if(R&&R.h[h]){med=R.h[h].mediana;lo=R.h[h].p10;hi=R.h[h].p90;fonte="regra: "+R.nome}
    if(h==1&&a.prev_hoje){med+=a.prev_hoje.total;lo+=a.prev_hoje.total;hi+=a.prev_hoje.total;fonte+=" + negócios de hoje"}
    return{h,med,lo,hi,fonte}})}
const diasUteis=(d0,n)=>{const d=new Date(d0+"T12:00:00");let k=0;while(k<n){d.setDate(d.getDate()+1);if(d.getDay()%6)k++}return d.toISOString().slice(0,10)};
// ---------------------------------------------------------------- visão geral
function geral(){const f=FS,c=f.carrego||{},s=f.sens||{},r=f.risco||{},p=f.perf||{},fl=f.fluxo||{},h=f.hoje||{};
  const pl=f.pl||0,cr=(f.credito_pct||0)>=1;
  if(!cr){$("#t_geral").innerHTML=`<div class="kp"><div class="kv"><span>Patrimônio</span><b>R$ ${mi(pl)}</b><i>${f.cotistas?br(f.cotistas,0)+" cotistas · ":""}fluxo 20 dias ${sg(fl[20],2,"% do PL")}</i></div>
     <div class="kv"><span>Retorno 12 meses</span><b>${p["12m"]?br(p["12m"].pct_cdi,0)+"% do CDI":"—"}</b><i>${p["12m"]?br(p["12m"].fundo,2)+"% vs CDI "+br(p["12m"].cdi,2)+"%":""}</i></div>
     <div class="kv"><span>Retorno no mês / no ano</span><b>${p.mes?br(p.mes.pct_cdi,0)+"% do CDI":"—"}</b><i>ano ${p.ano?br(p.ano.pct_cdi,0)+"% do CDI":"—"}</i></div></div>
    <div class="g2"><section><h2>Retorno do fundo × CDI (12 meses)</h2><div id="g_cota" class="pl"></div></section><section><h2>Alocação (CVM)</h2><div id="g_aloc"></div></section></div>
    <section><h2>Captação líquida (6 meses, R$ mi por dia)</h2><div id="g_fl" class="pl s"></div></section>`;
   const cs=f.cota_serie;if(cs)plot("g_cota",[{x:cs.d,y:cs.fundo,name:f.nome,mode:"lines",line:{color:css("--s1"),width:2}},{x:cs.d,y:cs.cdi,name:"CDI",mode:"lines",line:{color:css("--s4"),dash:"dot"}}],{yaxis:{gridcolor:css("--line")}});
   barras($("#g_aloc"),f.aloc_cat);const fs=f.fluxo_serie;if(fs)plot("g_fl",[{type:"bar",x:fs.d,y:fs.v,marker:{color:fs.v.map(v=>v<0?css("--s2"):css("--s1"))}}],{showlegend:false,yaxis:{gridcolor:css("--line")}});return}
  $("#t_geral").innerHTML=`<div class="kp">
   <div class="kv"><span>Patrimônio</span><b>R$ ${mi(pl)}</b><i>${f.cotistas?br(f.cotistas,0)+" cotistas · ":""}fluxo 20 dias ${sg(fl[20],2,"% do PL")}</i></div>
   <div class="kv"><span>Retorno 12 meses</span><b>${p["12m"]?br(p["12m"].pct_cdi,0)+"% do CDI":"—"}</b><i>${p["12m"]?br(p["12m"].fundo,2)+"% vs CDI "+br(p["12m"].cdi,2)+"%":""} · mês ${p.mes?br(p.mes.pct_cdi,0)+"% CDI":"—"}</i></div>
   <div class="kv"><span>Carrego da carteira de crédito</span><b>${c.pct_cdi?br(c.pct_cdi,0)+"% do CDI":"—"}</b><i>${c.spread_di!=null?"DI+ "+br(c.spread_di,2)+"%":""}${c.taxa_ipca!=null?" · IPCA+ "+br(c.taxa_ipca,2)+"%":""}</i></div>
   <div class="kv"><span>Sensibilidade a spread (CS01)</span><b>R$ ${br((s.cs01||0)/1e3,0)} mil/bp</b><i>duration ${br(s.duration_spread,2)} · juro real R$ ${br((s.dv01_ipca||0)/1e3,0)} mil/bp</i></div>
   <div class="kv"><span>VaR 95% (1 dia / 1 mês)</span><b>R$ ${br((r.var95_1d||0)/100*pl,2)} mi</b><i>1 mês R$ ${br((r.var95_1m||0)/100*pl,1)} mi · vol ${br(r.vol,2)}% a.a.</i></div>
   <div class="kv"><span>Marcação esperada hoje</span><b>${sg(h.impacto_bps,2," bps")}</b><i>R$ ${br((h.impacto_bps||0)/1e4*pl*1e3,0)} mil · ${h.n||0} ativos com sinal</i></div></div>
  <div class="g2"><section><h2>Retorno do fundo × CDI (12 meses)</h2><p class="sub">Cota oficial (informe diário CVM), base 100.</p><div id="g_cota" class="pl"></div></section>
   <section><h2>Próximos dias: retorno esperado contra o carrego</h2><p class="sub">Mediana dos sinais testados (bps do PL) e faixa de 80% pela vol do fundo. A linha tracejada é o carrego acima do CDI.</p><div id="g_proj" class="pl"></div></section></div>
  <div class="g2"><section><h2>Alocação</h2><div class="ctl seg" id="al_b"></div><div id="g_aloc"></div></section>
   <section><h2>De onde veio o resultado (12 meses, ativos com marcação ANBIMA)</h2><p class="sub">Carrego (taxa e indexador) × variação da taxa indicativa × resíduo, em bps do PL, com os pesos de hoje.</p><div id="g_pnl" class="pl s"></div></section></div>
  <div class="g2"><section><h2>Captação líquida (6 meses, R$ mi por dia)</h2><div id="g_fl" class="pl s"></div><p class="nota">5 dias ${sg(fl[5],2,"%")} · 20 dias ${sg(fl[20],2,"%")} · 60 dias ${sg(fl[60],2,"%")} do PL.</p></section>
   <section><h2>Liquidez da carteira de crédito</h2><p class="sub">Dias para vender cada posição a 20% do volume médio diário da B3 (30 dias).</p><div id="g_liq"></div></section></div>`;
  const cs=f.cota_serie;if(cs)plot("g_cota",[{x:cs.d,y:cs.fundo,name:f.nome,mode:"lines",line:{color:css("--s1"),width:2}},{x:cs.d,y:cs.cdi,name:"CDI",mode:"lines",line:{color:css("--s4"),dash:"dot"}}],{yaxis:{gridcolor:css("--line")}});
  else $("#g_cota").innerHTML='<p class="vazio">sem informe diário</p>';
  const P=f.projecao||[],hs=[0].concat(P.map(x=>x.h));
  if(P.length&&P[0].lo!=null){const md=[0].concat(P.map(x=>x.sinais)),lo=[0].concat(P.map(x=>x.lo)),hi=[0].concat(P.map(x=>x.hi)),ca=[0].concat(P.map(x=>x.carrego_vs_cdi));
   plot("g_proj",[{x:hs.concat(hs.slice().reverse()),y:hi.concat(lo.slice().reverse()),fill:"toself",mode:"lines",fillcolor:css("--band"),line:{width:0},name:"faixa de 80%",hoverinfo:"skip"},
    {x:hs,y:md,name:"mediana dos sinais",mode:"lines+markers",line:{color:css("--s1"),width:2}},{x:hs,y:ca,name:"carrego acima do CDI",mode:"lines",line:{color:css("--s2"),dash:"dash"}}],
    {xaxis:{title:{text:"dias úteis"},tickvals:hs,gridcolor:css("--line")},yaxis:{title:{text:"bps do PL"},gridcolor:css("--line")},hovermode:"x"})}
  else $("#g_proj").innerHTML='<p class="vazio">sem ativos com marcação suficiente</p>';
  const AL={"Categoria":f.aloc_cat,"Indexador":f.aloc_idx,"Emissor (top 15)":f.aloc_emissor,"Setor":f.aloc_setor,"Prazo (duration)":f.aloc_prazo};
  $("#al_b").innerHTML=Object.keys(AL).map((k,i)=>`<button class="${i?"":"on"}" data-k="${k}">${k}</button>`).join("");
  barras($("#g_aloc"),AL["Categoria"]);$("#al_b").onclick=e=>{if(!e.target.dataset.k)return;$("#al_b").querySelectorAll("button").forEach(b=>b.classList.toggle("on",b==e.target));barras($("#g_aloc"),AL[e.target.dataset.k])};
  const ps=f.pnl_serie;if(ps)plot("g_pnl",[["car","carrego","--s1"],["mtm","Δ taxa indicativa","--s2"],["res","resíduo","--s4"]].map(([k,n,c])=>({x:ps.d,y:ps[k],name:n,mode:"lines",line:{color:css(c),width:1.8}})),{yaxis:{title:{text:"bps do PL"},gridcolor:css("--line")}});
  else $("#g_pnl").innerHTML='<p class="vazio">—</p>';
  const fs=f.fluxo_serie;if(fs)plot("g_fl",[{type:"bar",x:fs.d,y:fs.v,marker:{color:fs.v.map(v=>v<0?css("--s2"):css("--s1"))},name:"captação líquida"}],{showlegend:false,yaxis:{gridcolor:css("--line")}});
  barras($("#g_liq"),f.liquidez)}
// ---------------------------------------------------------------- sinais
function sinais(){const f=FS,h=f.hoje||{};
  const cel=(R,hh)=>{const x=R.h[hh];return x?`${sg(x.mediana,0)} <span class="mut">(${br(x.p10,0)} a ${br(x.p90,0)})</span><br><b>${br(x.acerto,0)}%</b> <span class="mut">vs ${br(x.base,0)}%</span>`:"—"};
  $("#t_sinais").innerHTML=`<section><h2>Hoje: como a marcação ANBIMA deve fechar</h2><p class="sub">Negócios do dia na B3 contra a ANBIMA de ontem, juro real (DAP) nos papéis IPCA, ação do emissor. Impacto esperado: <b>${sg(h.impacto_bps,2," bps do PL")}</b>.</p><div id="s_hoje"></div>
   <p class="nota">Acerto histórico do sentido no mesmo dia: negócio ≥ R$ 20 mi 82%, R$ 5–20 mi 73%, R$ 1–5 mi 58% (447 papéis, 2023–2026, fora da amostra).</p></section>
  <section><h2>Próximos 1 a 20 dias: situações com histórico de ~70–80% de acerto</h2><p class="sub">Mediana do movimento do % do PU par (faixa 10–90% do que aconteceu) e acerto do sentido contra a base, nas 100 debêntures mais líquidas, estável em 2023-24 e 2025-26.</p><div id="s_reg"></div></section>
  <div class="g2"><section><h2>Mais fracos para 5 dias (modelo)</h2><p class="sub">10% piores do universo hoje: historicamente −5 bps contra o mercado e só 46% sobem.</p><div id="s_fr"></div></section>
   <section><h2>Mais fortes para 5 dias (modelo)</h2><p class="sub">10% melhores: historicamente +1 bps e 55% sobem (sinal mais fraco; muitas vezes é carrego de papel abaixo do par).</p><div id="s_fo"></div></section></div>
  <section><h2>Fatos relevantes, avisos aos debenturistas e eventos de crédito (60 dias)</h2><div id="s_doc"></div></section>`;
  h.ativos&&h.ativos.length?tabela($("#s_hoje"),[["cod","Ativo","l"],["emis","Emissor","l",(v,x)=>esc((A(x.cod).emissor||"").slice(0,26))],["bps","% par (bps)","",v=>sg(v,1)],["dtx","Δ taxa (bps)","",(v,x)=>sgt(-x.bps/Math.max(dmod(A(x.cod)),.3),1)],
    ["imp","Impacto (bps PL)","",v=>sg(v,2)],["perc","% PL","",v=>br(v,2)],["txt","Por quê","l wrap",v=>esc(v)]],h.ativos,{h:360,click:x=>abrir(x.cod)}):$("#s_hoje").innerHTML='<p class="vazio">Nada relevante até agora.</p>';
  const RA=(f.regras||[]).map(r=>({...r,R:RD[r.regra]}));
  RA.length?tabela($("#s_reg"),[["cod","Ativo","l"],["nome","Situação","l wrap",(v,x)=>`${esc(x.R.nome)} <span class="mut">(${esc(x.R.alvo)})</span>`],["h1","1 dia","",(v,x)=>cel(x.R,1)],["h5","5 dias","",(v,x)=>cel(x.R,5)],
    ["h10","10 dias","",(v,x)=>cel(x.R,10)],["h20","20 dias","",(v,x)=>cel(x.R,20)],["imp","Impacto 10d","",(v,x)=>x.R.h[10]?sg(x.R.h[10].mediana*x.perc/100,2," bps"):"—"],["perc","% PL","",v=>br(v,2)]],RA,{h:420,click:x=>abrir(x.cod)}):$("#s_reg").innerHTML='<p class="vazio">Nenhum papel do fundo está numa dessas situações hoje.</p>';
  const t5=el=>(rows)=>rows.length?tabela(el,[["cod","Ativo","l"],["mu","Centro (bps)","",v=>sg(v,1)],["q08","Faixa 80%","",(v,x)=>`${br(x.q08,0)} a ${br(x.q92,0)}`],["pct","Posição","",v=>br(v,0)+"%"],["perc","% PL","",v=>br(v,2)]],rows,{click:x=>abrir(x.cod)}):el.innerHTML='<p class="vazio">nenhum</p>';
  t5($("#s_fr"))(f.p5_fracos||[]);t5($("#s_fo"))(f.p5_fortes||[]);
  const dc=f.docs||[];$("#s_doc").innerHTML=dc.length?`<ul class="ls">${dc.map(d=>`<li><span class="tag ${/Fato|Recupera/.test(d.cat)?"t-dn":"t-up"}">${esc(d.cat)}</span> <span class="mut">${dt(d.data)} · ${esc((d.emissor||"").slice(0,32))} · ${br(d.perc,2)}% PL</span><br>${d.link?`<a href="${esc(d.link)}" target="_blank" rel="noopener">${esc(d.assunto||"documento")}</a>`:esc(d.assunto)}</li>`).join("")}</ul>`:'<p class="vazio">nenhum</p>'}
// ---------------------------------------------------------------- negócios do dia
let NSEL=null;
function negocios(){const N=D.negocios_dia||{negocios:[]},L=N.negocios||[];
  $("#t_negocios").innerHTML=`<section><h2>Negócios de ${dt(N.dia)} na B3 ranqueados pela diferença contra a ANBIMA</h2>
   <p class="sub">Taxa do negócio − taxa indicativa ANBIMA de ${dt(N.ref_data)}. Positivo = negociado a taxa maior (preço abaixo da marcação). O "esperado" é quanto a ANBIMA tende a andar no dia pelo volume do papel: ≥ R$ 20 mi anda 39% da diferença (acerto 82%), R$ 5–20 mi 30% (73%), R$ 1–5 mi 7% (58%), varejo é ruído.</p>
   <div class="ctl"><span class="seg" id="ng_v"><button data-v="0">todos</button><button class="on" data-v="50000">≥ R$ 50 mil</button><button data-v="1000000">≥ R$ 1 mi</button><button data-v="5000000">≥ R$ 5 mi</button></span>
    <span class="seg" id="ng_f"><button class="on" data-f="">mercado</button><button data-f="1">papéis do fundo</button></span>
    <select id="ng_i"><option value="">todos os índices</option>${[...new Set(L.map(x=>x.idx).filter(Boolean))].map(i=>`<option>${esc(i)}</option>`).join("")}</select><input id="ng_b" placeholder="buscar código ou emissor" aria-label="buscar"></div>
   <div id="ng_t"></div></section><section id="ng_g" hidden><h2 id="ng_h"></h2><p class="sub" id="ng_s"></p><div id="ng_c" class="pl l"></div></section>`;
  let vmin=50000,soF="";
  const go=()=>{const q=$("#ng_b").value.toUpperCase(),i=$("#ng_i").value;
    const R=L.filter(x=>x.vol>=vmin&&(!soF||FS.pesos[x.cod]!=null)&&(!i||x.idx==i)&&(!q||(x.cod+" "+((N.papeis[x.cod]||{}).emissor||"")).toUpperCase().includes(q))).map((x,k)=>({...x,rank:k+1,abs:Math.abs(x.dif),pl:FS.pesos[x.cod]}));
    R.length?tabela($("#ng_t"),[["rank","#",""],["hora","Hora","l"],["cod","Ativo","l"],["emissor","Emissor","l",(v,x)=>esc(((N.papeis[x.cod]||{}).emissor||(A(x.cod).emissor||"")).slice(0,26))],["idx","Índice","l"],["vol","R$ mil","",v=>br(v/1e3,0)],
      ["taxa","Taxa negócio","",v=>br(v,4)],["ref","ANBIMA","",v=>br(v,4)],["dif","Diferença (bps)","",v=>sgt(v,1)+(Math.abs(v)>500?' <span class="tag t-w" title="diferença muito grande: papel estressado ou taxa em outra convenção">verificar</span>':"")],["vol_dia","Papel no dia","",(v,x)=>`R$ ${br(v/1e6,1)} mi · ${x.n_dia} neg.`],
      ["esperado","ANBIMA esperada (bps)","",(v,x)=>v?`${sgt(v,1)} <span class="mut">${x.acerto}%</span>`:'<span class="mut">ruído</span>'],["pl","% PL do fundo","",v=>v==null?"":br(v,2)]],R,{ord:"abs",h:560,click:x=>grafNeg(x.cod)})
      :$("#ng_t").innerHTML='<p class="vazio">nenhum negócio com esse filtro</p>'};
  $("#ng_v").onclick=e=>{if(e.target.dataset.v==null)return;vmin=+e.target.dataset.v;$("#ng_v").querySelectorAll("button").forEach(b=>b.classList.toggle("on",b==e.target));go()};
  $("#ng_f").onclick=e=>{if(e.target.dataset.f==null)return;soF=e.target.dataset.f;$("#ng_f").querySelectorAll("button").forEach(b=>b.classList.toggle("on",b==e.target));go()};
  $("#ng_i").onchange=go;$("#ng_b").oninput=go;go();if(NSEL)grafNeg(NSEL)}
function grafNeg(c){NSEL=c;const N=D.negocios_dia,P=(N.papeis||{})[c],hoje=N.negocios.filter(x=>x.cod==c);if(!P)return;
  const x0=hoje[0]||{},emi=P.emissor||(A(c).emissor||"");
  $("#ng_g").hidden=false;$("#ng_h").textContent=`${c} · ${emi}`;
  $("#ng_s").innerHTML=`${dt(N.dia)}: ${x0.n_dia||0} negócios, R$ ${br((x0.vol_dia||0)/1e6,1)} mi, diferença média ${sgt(x0.dif_dia,1)} bps contra a ANBIMA (${br(x0.ref,4)}%). ANBIMA esperada no dia: ${x0.esperado?sgt(x0.esperado,1)+" bps de taxa (acerto histórico "+x0.acerto+"%)":"sem efeito (volume de varejo)"}.${AT[c]?` <a href="#" id="ng_ab">abrir o ativo</a>`:""}`;
  const pas=P.neg.d.map((d,i)=>i).filter(i=>P.neg.d[i]<N.dia);
  const bolha=v=>6+Math.sqrt(Math.max(v,0))*9;
  plot("ng_c",[{x:P.curva.d,y:P.curva.t,name:"taxa indicativa ANBIMA",mode:"lines+markers",marker:{size:4},line:{color:css("--s1"),width:2}},
    {x:pas.map(i=>P.neg.d[i]),y:pas.map(i=>P.neg.t[i]),mode:"markers",name:"negócios anteriores",marker:{size:pas.map(i=>bolha(P.neg.v[i])),color:css("--s4"),opacity:.45,line:{width:0}},
     text:pas.map(i=>`${P.neg.d[i]} ${P.neg.h[i]||""}<br>R$ ${br(P.neg.v[i]*1e3,0)} mil a ${br(P.neg.t[i],4)}%`),hovertemplate:"%{text}<extra></extra>"},
    {x:hoje.map(()=>N.dia),y:hoje.map(x=>x.taxa),mode:"markers",name:"negócios de "+dt(N.dia),marker:{size:hoje.map(x=>bolha(x.vol/1e6)),color:hoje.map(x=>x.dif>0?css("--dn"):css("--up")),opacity:.75,line:{width:1,color:css("--card")}},
     text:hoje.map(x=>`${x.hora} · R$ ${br(x.vol/1e3,0)} mil<br>${br(x.taxa,4)}% (${x.dif>0?"+":""}${br(x.dif,1)} bps)`),hovertemplate:"%{text}<extra></extra>"},
    ...(x0.esperado?[{x:[N.ref_data,N.dia],y:[x0.ref,x0.ref+x0.esperado/100],mode:"lines",name:"ANBIMA esperada",line:{color:css("--s2"),dash:"dot",width:2}}]:[])],
    {yaxis:{title:{text:"taxa (% a.a.)"},gridcolor:css("--line")},hovermode:"closest"});
  const ab=$("#ng_ab");if(ab)ab.onclick=e=>{e.preventDefault();abrir(c)};
  $("#ng_g").scrollIntoView({behavior:"smooth",block:"start"})}
// ---------------------------------------------------------------- ativos
function ativos(){const f=FS;
  const rows=(f.ativos||[]).map(x=>{const a=A(x.cod);return{...x,emissor:a.emissor,idx:a.idx,taxa:a.taxa,par:a.pct_par,dur:a.duration,dtx:a.dtaxa_1m,vol:a.vol,
    sinal:(a.prev_hoje&&Math.abs(a.prev_hoje.total)>=3?"H":"")+((a.regras_dias||[]).length?"R":"")+(a.prev5&&a.prev5.sinal?(a.prev5.sinal=="queda"?"Q":"A"):"")}});
  $("#t_ativos").innerHTML=`<section><h2>Ativos de crédito do fundo</h2><div class="ctl"><input id="bus" placeholder="buscar código ou emissor" aria-label="buscar"><select id="fidx"><option value="">todos os indexadores</option>${[...new Set(rows.map(r=>r.idx))].map(i=>`<option>${esc(i)}</option>`).join("")}</select><span class="mut">clique para ver a projeção, curvas, PnL, fundamentos</span></div><div id="a_t"></div>
   <p class="nota">Sinal: <span class="tag t-w">H</span> marcação de hoje · <span class="tag t-up">R</span> regra de 1–20 dias · <span class="tag t-dn">Q</span>/<span class="tag t-up">A</span> modelo 5 dias (10% mais fracos/fortes). ${f.outros&&f.outros.length?`Fora desta lista (sem código B3/ANBIMA): ${f.outros.slice(0,6).map(o=>esc(o.cat)+" "+br(o.perc,1)+"%").join(", ")}…`:""}</p></section><div id="a_det"></div>`;
  const go=()=>{const q=$("#bus").value.toUpperCase(),i=$("#fidx").value;tabela($("#a_t"),[["cod","Ativo","l"],["emissor","Emissor","l",v=>esc((v||"").slice(0,28))],["idx","Índice","l"],["perc","% PL","",v=>br(v,2)],["rs","R$ mi","",v=>br(v,1)],
    ["taxa","Taxa","",v=>br(v,2)],["par","% par","",v=>br(v,2)],["dur","Duration","",v=>br(v,2)],["dtx","Δ taxa 1m (bps)","",v=>sgt(v,0)],["vol","Vol","",v=>br(v,2)],["contrib","% risco","",v=>br(v,1)],
    ["dias_liq","Dias p/ vender","",v=>v==null?"sem neg.":v>60?"> 60":br(v,1)],["sinal","Sinal","",v=>(v||"").split("").map(s=>`<span class="tag ${s=="Q"?"t-dn":s=="H"?"t-w":"t-up"}">${s}</span>`).join(" ")]],
    rows.filter(r=>(!q||(r.cod+" "+(r.emissor||"")).toUpperCase().includes(q))&&(!i||r.idx==i)),{ord:"perc",h:480,click:x=>abrir(x.cod)})};
  $("#bus").oninput=go;$("#fidx").onchange=go;go();if(SEL)detalhe(SEL)}
function abrir(c){aba("ativos");detalhe(c)}
function detalhe(c){SEL=c;const a=A(c),e=E[a.raiz]||{},h=a.hist||{d:[]},w=pf(c),pr=proj(a),fu=e.fund,ac=e.acao_credito;
  const p0=h.pA&&h.pA.length?h.pA[h.pA.length-1]:a.pct_par,d0=h.d&&h.d.length?h.d[h.d.length-1]:D.hoje,dm_=dmod(a);
  const p20=pr[4],p5=pr[2];
  const leitura=`Em 5 dias úteis o % do PU par deve estar entre <b>${br(p0*Math.exp(p5.lo/1e4),2)}</b> e <b>${br(p0*Math.exp(p5.hi/1e4),2)}</b> (80%), centro ${br(p0*Math.exp(p5.med/1e4),2)} — taxa entre ${br(a.taxa-p5.hi/dm_/100,2)}% e ${br(a.taxa-p5.lo/dm_/100,2)}%. Em 20 dias: ${br(p0*Math.exp(p20.lo/1e4),2)} a ${br(p0*Math.exp(p20.hi/1e4),2)}. Base: ${esc(p20.fonte||p5.fonte)}.`;
  $("#a_det").innerHTML=`<section><h2>${esc(c)} · ${esc(a.emissor||"")} <span class="mut">${esc(a.indexador||"")} ${a.cupom!=null?"· cupom "+br(a.cupom,2):""} · venc. ${dt(a.venc)} · ${br(w,2)}% do PL</span></h2>
   ${a.lastro?`<p class="aviso">${esc(a.lastro)}</p>`:""}
   <div class="kp"><div class="kv"><span>Taxa indicativa</span><b>${br(a.taxa,2)}%</b><i>${dt(a.data_taxa)} · Δ 1 mês ${a.dtaxa_1m==null?"—":br(a.dtaxa_1m,0)+" bps"}</i></div><div class="kv"><span>% do PU par</span><b>${br(a.pct_par,2)}%</b><i>PU ${br(a.pu,2)}</i></div>
    <div class="kv"><span>Vol 12m</span><b>${br(a.vol,2)}%</b><i>${esc(a.fonte||"")}</i></div><div class="kv"><span>PnL 12m (bps do PU)</span><b>${br(a.pnl_total,0)}</b><i>carrego ${br(a.pnl_car,0)} · taxa ${br(a.pnl_mtm,0)}</i></div>
    ${e.alav!=null?`<div class="kv"><span>Dívida líq./EBITDA</span><b>${br(e.alav,2)}x</b><i>há 1 ano ${br(e.alav_ant,2)}x · cobertura ${br(e.cob,1)}x</i></div>`:""}
    ${e.acao?`<div class="kv"><span>Ação ${esc(e.acao.ticker)}</span><b>${sg(e.acao.d1,1,"%")}</b><i>5 dias ${sg(e.acao.s5,1,"%")}</i></div>`:""}</div>
   <h3>Projeção do % do PU par (6 meses + próximos 20 dias úteis)</h3><p class="leia">${leitura}</p><div id="d_fan" class="pl l"></div>
   <div class="g2"><div><h3>Taxa indicativa (cheia = ANBIMA, pontilhada = implícita no PU; pontos = negócios B3)</h3><div id="d_tx" class="pl"></div></div><div><h3>PnL acumulado contra a taxa indicativa (bps do PU)</h3><div id="d_pnl" class="pl"></div></div></div>
   <div class="g2"><div><h3>${ac?"Ação "+esc(e.acao.ticker)+" × taxa (semanal)":"Ação do emissor"}</h3>${ac?'<div id="d_ac" class="pl s"></div>':'<p class="vazio">emissor sem ação listada</p>'}</div>
    <div><h3>Fundamentos (CVM)</h3>${fu?'<div id="d_fu" class="pl s"></div>':'<p class="vazio">sem ITR/DFP na CVM</p>'}</div></div>
   <div class="g2"><div><h3>Agenda</h3>${(a.agenda||[]).length?`<ul class="ls">${a.agenda.map(g=>`<li>${esc(g.data)} · ${esc(g.evento)} ${esc(g.pct)}</li>`).join("")}</ul>`:'<p class="vazio">sem eventos em 6 meses</p>'}${e.prox_resultado?`<p>Próximo resultado: ~${dt(e.prox_resultado.esperado)} (entre ${dt(e.prox_resultado.min)} e ${dt(e.prox_resultado.max)})</p>`:""}</div>
    <div><h3>Documentos CVM (60 dias)</h3>${(e.docs||[]).length?`<ul class="ls">${e.docs.slice(0,8).map(d=>`<li><span class="mut">${dt(d.data)} ${esc(d.cat)}</span> ${d.link?`<a href="${esc(d.link)}" target="_blank" rel="noopener">${esc(d.assunto||"documento")}</a>`:esc(d.assunto)}</li>`).join("")}</ul>`:'<p class="vazio">nenhum</p>'}</div></div></section>`;
  // leque
  const z0=new Date(d0+"T12:00:00");z0.setDate(z0.getDate()-180);const i0=h.d.findIndex(d=>d>=z0.toISOString().slice(0,10));
  const hx=h.d.slice(Math.max(i0,0)),hy=(h.pA||[]).slice(Math.max(i0,0)),fx=pr.map(p=>p.h?diasUteis(d0,p.h):d0);
  const fy=k=>pr.map(p=>p0*Math.exp(p[k]/1e4));
  plot("d_fan",[{x:hx,y:hy,name:"% PU par ANBIMA",mode:"lines",line:{color:css("--s1"),width:2}},
    {x:fx.concat(fx.slice().reverse()),y:fy("hi").concat(fy("lo").reverse()),fill:"toself",mode:"lines",fillcolor:css("--band"),line:{width:0},name:"faixa 10–90%",hoverinfo:"skip"},
    {x:fx,y:fy("med"),name:"mais provável",mode:"lines+markers",line:{color:css("--s2"),dash:"dot",width:2},text:pr.map(p=>p.h?`${p.h}d: ${br(p.med,0)} bps (taxa ${br(-p.med/dm_,0)} bps)`:"hoje"),hovertemplate:"%{text}<extra></extra>"}],
    {yaxis:{title:{text:"% do PU par"},gridcolor:css("--line")}});
  const T=h.tA||[],tc=h.tc||[],ng=h.neg||{d:[],t:[],v:[]};
  plot("d_tx",[{x:h.d,y:T.map((v,i)=>tc[i]===1?v:null),name:"implícita",mode:"lines",line:{color:css("--s1"),dash:"dot",width:1.4}},{x:h.d,y:T.map((v,i)=>tc[i]===0?v:null),name:"ANBIMA",mode:"lines",line:{color:css("--s1"),width:2}},
    {x:ng.d,y:ng.t,mode:"markers",name:"negócios",marker:{size:ng.v.map(v=>4+Math.sqrt(v)*2),color:css("--s2"),opacity:.5}}],{yaxis:{title:{text:"% a.a."},gridcolor:css("--line")}});
  const cs=k=>{let s=0;return(h[k]||[]).map(v=>v==null?s:(s+=v))};
  plot("d_pnl",[["car","carrego","--s1"],["mtm","Δ taxa","--s2"],["res","resíduo","--s4"]].map(([k,n,c])=>({x:h.d,y:cs(k),name:n,mode:"lines",line:{color:css(c),width:1.8}})),{yaxis:{title:{text:"bps"},gridcolor:css("--line")}});
  if(ac){const cum=v=>v.reduce((r,x)=>{r.push((r.length?r[r.length-1]:0)+x);return r},[]);
    plot("d_ac",[{x:ac.datas,y:cum(ac.a),name:"ação (%)",mode:"lines",line:{color:css("--s3")}},{x:ac.datas,y:cum(ac.t),name:"taxa (bps)",yaxis:"y2",mode:"lines",line:{color:css("--s2")}}],{yaxis2:{overlaying:"y",side:"right",showgrid:false},yaxis:{gridcolor:css("--line")}})}
  if(fu)plot("d_fu",[{x:fu.ref,y:fu.dl_ebitda,name:"dívida líq./EBITDA",mode:"lines+markers",line:{color:css("--s2")}},{x:fu.ref,y:fu.cobertura,name:"EBITDA/desp. fin.",mode:"lines+markers",yaxis:"y2",line:{color:css("--s1")}}],{yaxis:{title:{text:"x"},gridcolor:css("--line")},yaxis2:{overlaying:"y",side:"right",showgrid:false}});
  $("#a_det").scrollIntoView({behavior:"smooth",block:"start"})}
// ---------------------------------------------------------------- risco
function risco(){const f=FS,r=f.risco,pl=f.pl||0;
  if(!r){$("#t_risco").innerHTML='<section><p class="vazio">Fundo sem ativos de crédito com marcação suficiente.</p></section>';return}
  const bt=r.backtest||{};
  $("#t_risco").innerHTML=`<div class="kp"><div class="kv"><span>Vol anual (crédito marcado)</span><b>${br(r.vol,2)}%</b><i>sem diversificação ${br(r.vol_sem_div,2)}%</i></div>
   <div class="kv"><span>VaR 95% 1 dia</span><b>R$ ${br(r.var95_1d/100*pl,2)} mi</b><i>${br(r.var95_1d,3)}% do PL</i></div><div class="kv"><span>Backtest (250 dias)</span><b>${br(bt.violacoes,1)}%</b><i>esperado 5% · real ${br(bt.vol_real,2)}% vs modelo ${br(bt.vol_modelo,2)}%</i></div>
   <div class="kv"><span>Correlação média entre papéis</span><b>${br(r.rho,2)}</b><i>${r.n} papéis, ${br(r.cobertura,1)}% do PL</i></div></div>
  <div class="g2"><section><h2>Cenários</h2><div id="r_cen"></div><p class="nota">Choques instantâneos sobre a duration de cada papel; default com recuperação de 40%.</p></section>
   <section><h2>Quem concentra o risco</h2><div id="r_con" class="pl"></div></section></div>
  <div class="g2"><section><h2>Retorno diário × VaR 95%</h2><div id="r_bt" class="pl"></div></section><section><h2>Sensibilidade a fatores (12 meses)</h2><div id="r_fat"></div><p class="nota">β em % do PL por 1% (ou 1 bp) do fator; |t| ≥ 2 é sensibilidade real.</p></section></div>
  <section><h2>Correlação entre os maiores papéis</h2><div id="r_cor" class="pl l"></div></section>`;
  tabela($("#r_cen"),[["nome","Cenário","l wrap"],["pct","% do PL","",v=>sg(v,2,"%")],["rs","R$ mi","",v=>sg(v/1e6,1)]],f.cenarios||[]);
  const C=(f.ativos||[]).filter(x=>x.contrib!=null).sort((a,b)=>b.contrib-a.contrib).slice(0,15).reverse();
  plot("r_con",[{type:"bar",orientation:"h",y:C.map(x=>x.cod),x:C.map(x=>x.contrib),marker:{color:css("--s1")},name:"% do risco"}],{showlegend:false,margin:{l:80,r:12,t:8,b:30},xaxis:{title:{text:"% do risco"},gridcolor:css("--line")}});
  const S=r.serie||{d:[],ret:[]};plot("r_bt",[{x:S.d,y:S.ret,type:"bar",name:"retorno diário (%)",marker:{color:S.ret.map(v=>v<S.var?css("--dn"):css("--s4"))}},{x:[S.d[0],S.d[S.d.length-1]],y:[S.var,S.var],mode:"lines",name:"VaR 95%",line:{color:css("--dn"),dash:"dash"}}],{yaxis:{gridcolor:css("--line")}});
  tabela($("#r_fat"),[["fator","Fator","l"],["beta","β","",v=>br(v,4)],["t","t","",v=>br(v,1)],["r2","R²","",v=>br(v,2)]],r.fatores||[],{ord:"r2"});
  const M=r.corr;plot("r_cor",[{type:"heatmap",z:M.m,x:M.cods,y:M.cods,zmin:-1,zmax:1,colorscale:[[0,css("--s2")],[.5,css("--card")],[1,css("--s1")]]}],{showlegend:false,margin:{l:76,r:10,t:8,b:76},xaxis:{tickfont:{size:9}},yaxis:{tickfont:{size:9}}})}
// ---------------------------------------------------------------- agenda
function agenda(){const f=FS,ag=f.agenda||[];const caixa=ag.filter(a=>a.rs!=null&&a.data<=diasUteis(D.hoje,63)).reduce((s,a)=>s+a.rs,0);
  $("#t_agenda").innerHTML=`<div class="kp"><div class="kv"><span>Caixa de amortizações e vencimentos (3 meses)</span><b>R$ ${br(caixa,1)} mi</b><i>${br(caixa/(f.pl||1)*100,2)}% do PL (estimado)</i></div>
   <div class="kv"><span>Resultados de emissores (30 dias)</span><b>${ag.filter(a=>/Resultado/.test(a.evento)&&a.data<=diasUteis(D.hoje,21)).length}</b><i>datas estimadas pelo histórico de entregas à CVM</i></div>
   <div class="kv"><span>Focus</span><b>${D.focus&&D.focus.Selic?"Selic "+br(Object.values(D.focus.Selic.anos)[0],2)+"%":"—"}</b><i>${D.focus&&D.focus.IPCA?"IPCA "+Object.entries(D.focus.IPCA.anos).slice(0,2).map(([a,v])=>a+" "+br(v,2)+"%").join(" · "):""}</i></div></div>
  <section><h2>Próximos 120 dias</h2><div class="ctl seg" id="ag_f"><button class="on" data-k="">tudo</button><button data-k="Amortiza|Vencimento|Resgate">caixa</button><button data-k="juros">juros</button><button data-k="Resultado">resultados</button><button data-k="IPCA">macro</button></div><div id="ag_t"></div>
   <p class="nota">Amortização e vencimento em R$ = % do valor nominal × posição (estimado). Pagamento não é perda: o PU cai e o caixa entra.</p></section>`;
  const go=k=>tabela($("#ag_t"),[["data","Data","l",dt],["evento","Evento","l"],["pct","%","l"],["cod","Ativo","l"],["rs","R$ mi","",v=>v==null?"":br(v,2)],["perc","% PL","",v=>br(v,2)]],ag.filter(a=>!k||new RegExp(k,"i").test(a.evento+" "+a.cod)),{h:620,click:x=>AT[x.cod]&&abrir(x.cod)});
  go("");$("#ag_f").onclick=e=>{if(e.target.dataset.k==null)return;$("#ag_f").querySelectorAll("button").forEach(b=>b.classList.toggle("on",b==e.target));go(e.target.dataset.k)}}
// ---------------------------------------------------------------- mercado
function mercado(){const f=FS;
  $("#t_mercado").innerHTML=`<div class="g2"><section><h2>Índices de crédito e renda fixa</h2><div id="m_idx"></div><div id="m_ic" class="pl s"></div><p class="nota">IDA (ANBIMA) e JGP via SmartWealth; dia em desvios = variação / desvio diário de 1 ano.</p></section>
   <section><h2>Juros (B3, ~15 min) e referências</h2><div id="m_j"></div><div id="m_ref"></div></section></div>
  <section><h2>Notícias dos emissores do fundo (14 dias)</h2><div id="m_n" class="g3"></div></section>`;
  const I=Object.entries(D.indices).map(([k,v])=>({k,...v}));
  tabela($("#m_idx"),[["k","Índice","l"],["data","Data","",dt],["d1","Dia","",v=>sg(v,3,"%")],["z","Desvios","",v=>sg(v,1)],["m1","1 mês","",v=>sg(v,2,"%")],["a12","12 meses","",v=>sg(v,1,"%")]],I.filter(x=>/IDA|JGP|IMA-B|IRF/.test(x.k)));
  plot("m_ic",I.filter(x=>/IDA|JGP/.test(x.k)).map(v=>{const n=Math.max(0,v.d.findIndex(d=>d>=diasUteis(D.hoje,-1).replace(/^(\d{4})/,y=>y-1)));const b=v.v[n];return{x:v.d.slice(n),y:v.v.slice(n).map(q=>q/b*100),name:v.k,mode:"lines",line:{width:1.4}}}),{yaxis:{gridcolor:css("--line")}});
  const J=D.juros||{};$("#m_j").innerHTML=Object.keys(J).length?`<div class="sc"><table><tr><th class="l">Prazo</th>${Object.keys(J).map(k=>`<th>${k}</th><th>Δ bps</th>`).join("")}</tr>${(J.DI1||J.DAP||[]).map((_,i)=>`<tr><td class="l">${(J.DI1||J.DAP)[i].prazo}</td>${Object.values(J).map(L=>L[i]?`<td>${br(L[i].taxa,2)}</td><td>${sgt(L[i].bps,1)}</td>`:"<td></td><td></td>").join("")}</tr>`).join("")}</table></div>`:'<p class="vazio">fora do pregão</p>';
  tabela($("#m_ref"),[["k","","l"],["preco","Último","",v=>br(v,2)],["d1","Dia","",v=>sg(v,2,"%")],["z","Desvios","",v=>sg(v,1)],["m1","1 mês","",v=>sg(v,1,"%")]],Object.entries(D.mercado).map(([k,v])=>({k,...v})));
  const N=(f.raizes||[]).filter(r=>D.noticias[r]).map(r=>[r,D.noticias[r]]);
  $("#m_n").innerHTML=N.length?N.map(([r,n])=>`<div><h3>${esc((E[r]||{}).nome||n.busca)}</h3><ul class="ls">${n.itens.slice(0,6).map(i=>`<li>${i.alerta.length?`<span class="tag t-dn">${esc(i.alerta[0])}</span> `:""}<a href="${esc(i.link)}" target="_blank" rel="noopener">${esc(i.titulo)}</a> <span class="mut">${dt(i.data)}</span></li>`).join("")}</ul></div>`).join(""):'<p class="vazio">sem notícias</p>'}
// ---------------------------------------------------------------- tudo
function avisos(){const f=FS,m=f.mes,idade=m?Math.round((new Date(D.hoje)-new Date(m+"-28"))/864e5/30):null;
  $("#avisos").innerHTML=(f.fonte!="planilha"&&idade>1?`<div class="aviso">Carteira de ${esc(m)} (${esc(f.fonte)}): a CVM divulga as carteiras com atraso. Para usar a posição de hoje, salve a planilha em <b>carteiras/${esc(f.nome)}.xlsx</b> (colunas ativo e perc_nav).</div>`:"")+
   (f.credito_pct<1?`<div class="aviso">Fundo sem crédito privado com marcação ANBIMA: aparecem só patrimônio, retorno, alocação e captação.</div>`:"")}
function tudo(){avisos();geral();sinais();negocios();ativos();risco();agenda();mercado();let a="geral";try{a=localStorage.getItem("mc_aba")||a}catch(e){};aba(ABAS.some(x=>x[0]==a)?a:"geral")}
let ini=0;try{const id=localStorage.getItem("mc_fundo");const i=F.findIndex(f=>f.id==id);if(i>=0)ini=i;else ini=F.findIndex(f=>f.nome=="Dual")}catch(e){};ini=Math.max(ini,0);
$("#fundo").value=ini;FS=F[ini];tudo();
</script>
"""


if __name__ == "__main__":
    args = sys.argv[1:]
    if args and args[0] == "modelo":
        modelo()
    else:
        arq = next((Path(a) for a in args if a.lower().endswith((".xlsx", ".xls"))), None)
        rodar(arq) if "rodar" in args else uma_vez(arq)
