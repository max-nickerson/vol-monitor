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
            s = s[(np.log(s / s.rolling(7, center=True, min_periods=3).median())).abs() < 0.05]    # ponto isolado fora da curva (erro da fonte)
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
        r = get("https://news.google.com/rss/search", params={"q": f'"{q}" when:{dias}d', "hl": "pt-BR", "gl": "BR", "ceid": "BR:pt-419"}, timeout=20, tent=2)
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
    s = str(indexador or "").upper()
    return "%DI" if "%" in s and "DI" in s else "DI+" if "DI" in s else "IPCA" if "IPCA" in s else "IGPM" if "IGP" in s else "PRE"


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
def calcula(arq):
    t0 = time.time()
    a = identifica(le_carteira(arq))
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
    A["prev5"] = [p5.get(c) for c in A.codigo]
    A.attrs["info5"] = info5
    log(f"previsão 5d: {len(p5)} ativos ({time.time() - t0:.0f}s)")
    A.attrs["juros"] = juros
    log(f"previsibilidade ok ({time.time() - t0:.0f}s)")
    # ---------------- mercado e risco
    idx_c = indices_credito()
    mkt = {k: yahoo(v, "2y") for k, v in MERCADO.items()}
    td_ = tesouro()
    risco_ = risco(A, diarios, idx_c, mkt, td_)
    valor_rel = valor_relativo(A, emis)
    return A, emis, prev, risco_, idx_c, mkt, td_, valor_rel


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
    S["idx"] = S.cod.map(cad.idx).fillna("").map(tipo_idx)
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
    g["idx"] = [tipo_idx(cad.idx.get(c)) for c in g.index]
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
def pagina(A, emis, prev, R, idx_c, mkt, td_, vr, ctx):
    limpa = lambda o: None if isinstance(o, float) and not np.isfinite(o) else o
    ativos = [{k: limpa(v) for k, v in r.items() if not k.startswith("_")} for r in A.drop(columns=["faixa"], errors="ignore").to_dict("records")]
    m_ = {}
    for nome, y in mkt.items():
        if y and len(y["serie"]) > 30:
            s = y["serie"]
            lr = np.log(s).diff().dropna()
            d1 = (y["preco"] / y["ant"] - 1) * 100 if y["preco"] and y["ant"] else None
            z = d1 / (lr.tail(250).std() * 100) if d1 is not None else None
            if z is not None and abs(z) > 8:                       # cotação impossível (erro da fonte): não mostra
                d1 = z = None
            m_[nome] = {"preco": y["preco"], "d1": d1, "z": z,
                        "m1": float((s.iloc[-1] / s.iloc[-22] - 1) * 100), "d": [d.strftime("%Y-%m-%d") for d in s.index[-130:]], "v": [round(v, 4) for v in s.values[-130:]]}
    ic = {}
    for nome, d in idx_c.items():
        s = d["serie"]
        if len(s) > 30:
            lr = np.log(s).diff().dropna()
            ic[nome] = {"d1": float(lr.iloc[-1] * 100), "z": float(lr.iloc[-1] / lr.tail(250).std()) if lr.tail(250).std() > 0 else None,
                        "m1": float((s.iloc[-1] / s[s.index <= s.index[-1] - pd.Timedelta(days=30)].iloc[-1] - 1) * 100),
                        "a12": float((np.prod(1 + d["mensal"].tail(12)) - 1) * 100) if len(d["mensal"]) >= 12 else None,
                        "data": s.index[-1].strftime("%Y-%m-%d"), "d": [q.strftime("%Y-%m-%d") for q in s.index], "v": [round(v, 6) for v in s.values]}
    tj = {c: {"d": [d.strftime("%Y-%m-%d") for d in td_.index[-500:]], "v": [None if v != v else round(v, 3) for v in td_[c].values[-500:]]} for c in td_.columns} if len(td_) else {}
    dados = {"gerado": datetime.now().strftime("%d/%m/%Y %H:%M"), "hoje": HOJE.strftime("%Y-%m-%d"), "ativos": ativos,
             "emissores": json.loads(json.dumps(emis, default=lambda o: None if isinstance(o, float) else str)), "prev": prev, "risco": R,
             "mercado": m_, "indices": ic, "tesouro": tj, "valor_rel": vr, **ctx}
    txt = json.dumps(dados, ensure_ascii=False, default=lambda o: None if isinstance(o, float) and not np.isfinite(o) else (o.isoformat() if hasattr(o, "isoformat") else str(o)))
    txt = re.sub(r"\bNaN\b|\b-?Infinity\b", "null", txt).replace("</", "<\\/")
    return HTML.replace("/*DADOS*/", txt)


def uma_vez(arq):
    global HOJE
    HOJE = pd.Timestamp.today().normalize()
    A, emis, prev, R, idx_c, mkt, td_, vr = calcula(arq)
    ag = []
    for x in A[A.coberto].itertuples():
        for e in x.agenda if isinstance(x.agenda, list) else []:
            ag.append({"data": pd.to_datetime(e["data"], dayfirst=True).strftime("%Y-%m-%d"), "ativo": x.codigo, "emissor": x.emissor, "evento": e["evento"],
                       "pct": e["pct"], "peso": x.perc_nav})
    ipca = ipca_datas()
    for r, e in emis.items():
        if e.get("prox_resultado"):
            ag.append({"data": e["prox_resultado"]["esperado"], "ativo": ", ".join(e["codigos"][:3]), "emissor": e["nome"],
                       "evento": f"Resultado {e['prox_resultado']['ref'][5:7]}/{e['prox_resultado']['ref'][:4]} (estimado)", "pct": "", "peso": e["perc"]})
    for x in A[A.coberto].itertuples():
        if x.venc and HOJE <= pd.Timestamp(x.venc) <= HOJE + pd.Timedelta(days=180):
            ag.append({"data": x.venc, "ativo": x.codigo, "emissor": x.emissor, "evento": "Vencimento", "pct": "100%", "peso": x.perc_nav})
    ag += [{"data": d.strftime("%Y-%m-%d"), "ativo": "IPCA", "emissor": "IBGE", "evento": "Divulgação do IPCA", "pct": "", "peso": float(A[A.idx.eq("IPCA")].perc_nav.sum())} for d in ipca[:4]]
    ag = sorted([g for g in ag if g["data"] >= HOJE.strftime("%Y-%m-%d")], key=lambda g: g["data"])
    news = {}
    for e in [e for e in sorted(emis.values(), key=lambda e: -e["perc"]) if not re.search(r"SECURITIZ", str(e["nome"]).upper())][:20]:
        q, it = noticias(e["nome"])
        if it:
            news[e["raiz"]] = {"busca": q, "itens": it}
    ctx = {"agenda": ag, "sinais": sinais(A, emis, prev, ipca, ag), "noticias": news, "juros": A.attrs.get("juros") or juros_b3(), "focus": focus(),
           "estudo": ESTUDO, "mercado_hoje": A.attrs.get("mercado_hoje"), "data_ref": A.attrs.get("data_ref"), "info5": A.attrs.get("info5"), "estudo5": ESTUDO5}
    h = pagina(A, emis, prev, R, idx_c, mkt, td_, vr, ctx)
    (SAIDA / "vol_monitor.html").write_text(h, encoding="utf-8")
    log(f"saida/vol_monitor.html {len(h) / 1e6:.1f} MB")


def rodar(arq):
    while True:
        try:
            uma_vez(arq)
        except Exception as e:
            log(f"erro: {type(e).__name__} {e}")
        hr = datetime.now().hour + datetime.now().minute / 60
        time.sleep(900 if 9 <= hr <= 18.5 and datetime.now().weekday() < 5 else 3600)


HTML = r"""<title>Vol monitor de crédito</title>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<link rel="preconnect" href="https://fonts.googleapis.com"><link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/plotly.js/3.1.0/plotly.min.js"></script>
<script>if(!window.Plotly)document.write('<script src="https://cdn.jsdelivr.net/npm/plotly.js-dist-min@3.1.0/plotly.min.js"><\/script>')</script>
<style>
:root{--bg:#f6f7f8;--card:#fff;--ink:#14191d;--mut:#5d6873;--line:#dfe3e7;--acc:#0f766e;--acc2:#d9efec;--warn:#b45309;--warn2:#fdf0dc;--crit:#b42318;--crit2:#fde5e2;--ok:#15803d;--chip:#eef1f3;--s1:#0f766e;--s2:#c2410c;--s3:#6d28d9;--s4:#64748b;color-scheme:light}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#111416;--card:#1a1e21;--ink:#e6e8ea;--mut:#98a2ab;--line:#2b3136;--acc:#3fbfa9;--acc2:#173731;--warn:#e0a03a;--warn2:#3a2c14;--crit:#f2665a;--crit2:#3d1b18;--ok:#4cc38a;--chip:#252b2f;--s1:#3fbfa9;--s2:#f08a4b;--s3:#a78bfa;--s4:#94a3b8;color-scheme:dark}}
:root[data-theme="dark"]{--bg:#111416;--card:#1a1e21;--ink:#e6e8ea;--mut:#98a2ab;--line:#2b3136;--acc:#3fbfa9;--acc2:#173731;--warn:#e0a03a;--warn2:#3a2c14;--crit:#f2665a;--crit2:#3d1b18;--ok:#4cc38a;--chip:#252b2f;--s1:#3fbfa9;--s2:#f08a4b;--s3:#a78bfa;--s4:#94a3b8;color-scheme:dark}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.45 "IBM Plex Sans",system-ui,sans-serif}
.w{max-width:1320px;margin:0 auto;padding-inline:16px;padding-block:14px 40px}
header{display:flex;flex-wrap:wrap;align-items:baseline;gap:8px 16px;margin-bottom:10px}h1{font-size:20px;margin:0;font-weight:600}
.mut{color:var(--mut)}.num,td,.k b{font-variant-numeric:tabular-nums}
nav{display:flex;gap:4px;flex-wrap:wrap;position:sticky;top:env(safe-area-inset-top,0px);background:var(--bg);padding:6px 0;z-index:5;border-bottom:1px solid var(--line);margin-bottom:12px}
nav button{border:1px solid var(--line);background:var(--card);color:var(--ink);border-radius:6px;padding:6px 11px;font:inherit;cursor:pointer}
nav button.on{background:var(--acc);border-color:var(--acc);color:#fff}
.kp{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:12px}
.k{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:9px 11px}.k span{display:block;color:var(--mut);font-size:12px}.k b{font-size:19px;font-weight:600}.k i{display:block;font-style:normal;color:var(--mut);font-size:12px}
section{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:12px 14px;margin-bottom:12px;min-width:0}
h2{font-size:15px;margin:0 0 8px;font-weight:600}h3{font-size:13px;margin:10px 0 6px;font-weight:600}
.g2{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:12px}.g3{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px}
@media(max-width:900px){.g2,.g3{grid-template-columns:minmax(0,1fr)}}
.sc{overflow-x:auto}table{border-collapse:collapse;width:100%;font-size:13px}th,td{padding:5px 7px;border-bottom:1px solid var(--line);text-align:right;white-space:nowrap}
th{color:var(--mut);font-weight:500;cursor:pointer;position:sticky;top:0;background:var(--card)}th.l,td.l{text-align:left}td.wrap{white-space:normal;min-width:240px}
tr.cl{cursor:pointer}tr.cl:hover td{background:var(--chip)}tr.sel td{background:var(--acc2)}
.pos{color:var(--ok)}.neg{color:var(--crit)}.chip{display:inline-block;background:var(--chip);border-radius:4px;padding:1px 6px;font-size:12px;margin-right:4px}
.tag{display:inline-block;border-radius:4px;padding:1px 6px;font-size:12px;font-weight:500}.t-crit{background:var(--crit2);color:var(--crit)}.t-warn{background:var(--warn2);color:var(--warn)}.t-ok{background:var(--acc2);color:var(--acc)}
.pl{height:290px}.pl.s{height:230px}input,select{font:inherit;background:var(--card);color:var(--ink);border:1px solid var(--line);border-radius:6px;padding:5px 8px;max-width:100%}
.ctl{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:8px}a{color:var(--acc)}
.nota{font-size:12px;color:var(--mut);margin:6px 0 0}.vazio{color:var(--mut);font-style:italic}
ul.ls{list-style:none;padding:0;margin:0}ul.ls li{padding:6px 0;border-bottom:1px solid var(--line)}
[hidden]{display:none!important}
</style>
<div class="w">
<header><h1>Vol monitor de crédito</h1><span class="mut" id="ger"></span><button id="tema" style="margin-left:auto" class="chip">tema</button></header>
<nav id="nav"></nav>
<div id="p_resumo"></div><div id="p_agenda" hidden></div><div id="p_mercado" hidden></div><div id="p_ativos" hidden></div>
<div id="p_prev" hidden></div><div id="p_risco" hidden></div><div id="p_metodo" hidden></div>
</div>
<script>
const D=/*DADOS*/;
const $=s=>document.querySelector(s),css=v=>getComputedStyle(document.documentElement).getPropertyValue(v).trim();
const br=(v,c=2)=>v==null||!isFinite(v)?"—":Number(v).toLocaleString("pt-BR",{minimumFractionDigits:c,maximumFractionDigits:c});
const sg=(v,c=1,u="")=>v==null||!isFinite(v)?"—":`<span class="${v>0?"pos":v<0?"neg":""}">${v>0?"+":""}${br(v,c)}${u}</span>`;
const sgi=(v,c=1,u="")=>v==null||!isFinite(v)?"—":`<span class="${v>0?"neg":v<0?"pos":""}">${v>0?"+":""}${br(v,c)}${u}</span>`;
const dt=s=>s?s.slice(8,10)+"/"+s.slice(5,7)+"/"+s.slice(2,4):"—",esc=s=>String(s??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const A=D.ativos.filter(a=>a.coberto),E=D.emissores,R=D.risco;
const lay=(o={})=>Object.assign({margin:{l:48,r:12,t:10,b:34},paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"rgba(0,0,0,0)",font:{family:"IBM Plex Sans",size:11,color:css("--ink")},
  xaxis:{gridcolor:css("--line"),zerolinecolor:css("--line")},yaxis:{gridcolor:css("--line"),zerolinecolor:css("--mut")},showlegend:true,legend:{orientation:"h",y:1.14,x:0}},o);
const plot=(id,t,l)=>{if(window.Plotly&&document.getElementById(id))Plotly.react(id,t,lay(l),{displayModeBar:false,responsive:true})};
function tabela(el,cols,rows,o={}){let k=o.ord||null,dir=-1;const go=()=>{let r=rows.slice();if(k){r.sort((a,b)=>{const x=a[k],y=b[k];return (x==null)-(y==null)||(typeof x=="string"?x.localeCompare(y):(x-y))*dir})}
  el.innerHTML=`<div class="sc"${o.h?` style="max-height:${o.h}px;overflow:auto"`:""}><table><tr>${cols.map(c=>`<th class="${c[2]||""}" data-k="${c[0]}">${c[1]}</th>`).join("")}</tr>${r.map((x,i)=>`<tr class="${o.click?"cl":""}" data-i="${rows.indexOf(x)}">${cols.map(c=>`<td class="${c[2]||""}">${c[3]?c[3](x[c[0]],x):esc(x[c[0]]??"—")}</td>`).join("")}</tr>`).join("")}</table></div>`;
  el.querySelectorAll("th").forEach(th=>th.onclick=()=>{dir=k==th.dataset.k?-dir:-1;k=th.dataset.k;go()});
  if(o.click)el.querySelectorAll("tr.cl").forEach(tr=>tr.onclick=()=>o.click(rows[+tr.dataset.i],tr))};go()}
// ------------------------------------------------------------ navegação
const ABAS=[["resumo","Hoje"],["agenda","O que vem"],["mercado","Mercado e notícias"],["ativos","Ativos"],["prev","O que antecipa"],["risco","Risco"],["metodo","Como é calculado"]];
function aba(id){ABAS.forEach(([k])=>{$("#p_"+k).hidden=k!=id;$(`#nav [data-a="${k}"]`).classList.toggle("on",k==id)});try{localStorage.setItem("vm_aba",id)}catch(e){};
  if(window.Plotly)setTimeout(()=>document.querySelectorAll(`#p_${id} .js-plotly-plot`).forEach(g=>Plotly.Plots.resize(g)),30)}
$("#nav").innerHTML=ABAS.map(([k,t])=>`<button data-a="${k}">${t}</button>`).join("");$("#nav").onclick=e=>{if(e.target.dataset.a)aba(e.target.dataset.a)};
$("#tema").onclick=()=>{const r=document.documentElement,d=r.dataset.theme=="dark"||(!r.dataset.theme&&matchMedia("(prefers-color-scheme:dark)").matches);r.dataset.theme=d?"light":"dark";tudo()};
$("#ger").textContent=`atualizado ${D.gerado} · ${A.length} ativos com código, ${br(R.pl_codigo,1)}% do PL`;
// ------------------------------------------------------------ hoje
function resumo(){const bt=R.backtest||{};const pP=D.prev.find(p=>p.id=="negocio_marcacao");
  const pnl=k=>A.reduce((s,a)=>s+(a[k]!=null?a[k]*a.perc_nav/100:0),0);
  $("#p_resumo").innerHTML=`<div class="kp">
   <div class="k"><span>Vol anual (cobertos)</span><b>${br(R.vol,2)}%</b><i>sem diversificação ${br(R.vol_sem_div,2)}%</i></div>
   <div class="k"><span>VaR 95% 1 dia</span><b>${br(R.var95_1d,3)}%</b><i>1 mês ${br(R.var95_1m,2)}% do PL</i></div>
   <div class="k"><span>Backtest VaR (250 d)</span><b>${br(bt.violacoes,1)}%</b><i>esperado 5% · real ${br(bt.vol_real,2)}% vs modelo ${br(bt.vol_modelo,2)}%</i></div>
   <div class="k"><span>Duration média</span><b>${br(R.duration,2)}</b><i>+100 bps na taxa: ${br(R.dv01,2)}% do PL</i></div>
   <div class="k"><span>PnL 12m: carrego</span><b>${br(pnl("pnl_car")/100,2)}%</b><i>do PL, ativos com série oficial</i></div>
   <div class="k"><span>PnL 12m: taxa indicativa</span><b>${sg(pnl("pnl_mtm")/100,2,"%")}</b><i>resíduo ${br(pnl("pnl_res")/100,2)}%</i></div></div>
  <section><h2>Previsão da marcação ANBIMA de hoje</h2><div id="ph"></div></section>
  <section><h2>Próximos 5 dias úteis: onde o % do PU par deve estar</h2><div id="p5"></div></section>
  <section><h2>O que olhar agora</h2><div id="sin"></div><p class="nota">Ordenado pelo impacto estimado em bps do PL e pelo peso. Impacto só aparece quando o sinal é estatisticamente útil (aba "O que antecipa").</p></section>
  <div class="g2"><section><h2>Mercado de crédito hoje</h2><div id="r_idx"></div></section><section><h2>Ativos que mais andaram em 1 mês (taxa indicativa)</h2><div id="r_mov"></div></section></div>`;
  const PH=A.filter(a=>a.prev_hoje).map(a=>({cod:a.codigo,emissor:a.emissor,perc:a.perc_nav,...a.prev_hoje,txt:Object.entries(a.prev_hoje.comp).map(([k,c])=>`${k}: ${sg(c.bps,1)} (${esc(c.txt)})`).join("<br>")}));
  const tot=PH.reduce((s,p)=>s+p.impacto_pl,0);
  $("#ph").innerHTML=`<p>Esperado para o fundo: <b>${sg(tot,2," bps do PL")}</b> na marcação de hoje (ANBIMA de ${dt(D.data_ref)} → hoje), ${PH.length} ativos com sinal.${D.mercado_hoje&&Object.keys(D.mercado_hoje).length?" Tom do mercado: "+Object.entries(D.mercado_hoje).map(([k,v])=>`${k} ${sg(v,0," bps")}`).join(", "):""}</p><div id="ph_t"></div>
   <p class="nota">bps de % do PU par (positivo = PU sobe). Δ taxa ≈ −% par / duration modificada. Regras do estudo intradiário (aba "O que antecipa"): negócio do papel conta pelo volume (≥ R$ 20 mi: a ANBIMA anda 39% do gap, acerto 82%); papel IPCA anda com o DAP; ação ≥ 2% no dia.</p>`;
  PH.length?tabela($("#ph_t"),[["cod","Ativo","l"],["emissor","Emissor","l",v=>esc((v||"").slice(0,26))],["total","% par (bps)","",v=>sg(v,1)],["dtaxa","Δ taxa (bps)","",v=>sgi(v,1)],["impacto_pl","Impacto (bps PL)","",v=>sg(v,2)],["perc","% PL","",v=>br(v,2)],["txt","Por quê","l wrap",v=>v]],
    PH.sort((a,b)=>Math.abs(b.impacto_pl)-Math.abs(a.impacto_pl)),{h:380}):$("#ph_t").innerHTML='<p class="vazio">Sem negócios relevantes, juros ou ações fora do normal até agora.</p>';
  const P5=A.filter(a=>a.prev5).map(a=>({cod:a.codigo,emissor:a.emissor,perc:a.perc_nav,par:a.pct_par,...a.prev5,dur:a.duration,
     porq:a.prev5.porque.map(x=>`${esc(x.f)} ${sg(x.bps,1)}`).join(" · ")}));
  const I5=D.info5||{};
  $("#p5").innerHTML=P5.length?`<p>Faixa de 80% para o % do PU par daqui a 5 dias úteis (8%–92%) e o centro esperado do papel contra o mercado. <span class="tag t-crit">queda</span> / <span class="tag t-ok">alta</span> = 10% mais fracos/fortes do universo hoje (${I5.n_univ_hoje||"—"} papéis): historicamente ${D.estudo5.decis[0][1]} e +${D.estudo5.decis[6][1]} bps em 5 dias, ${D.estudo5.decis[0][2]}% e ${D.estudo5.decis[6][2]}% de alta.</p><div id="p5_t"></div>
   <p class="nota">Modelo treinado hoje com ${I5.papeis||"—"} papéis (até ${dt(I5.ate)}); intervalo calibrado em ${esc(I5.calibracao||"")} (cobertura ${br(I5.cobertura_cal,0)}%). Teste fora da amostra: cobertura ${D.estudo5.cobertura}%, IC ${br(D.estudo5.ic,3)} (t ${br(D.estudo5.ic_t,1)}). O centro é pequeno de propósito: em 5 dias a maior parte do movimento é ruído; a faixa é o que mais importa.</p>`:'<p class="vazio">sem previsão (instale lightgbm)</p>';
  if(P5.length)tabela($("#p5_t"),[["cod","Ativo","l"],["emissor","Emissor","l",v=>esc((v||"").slice(0,24))],["par","% par hoje","",v=>br(v,2)],["mu","Centro (bps)","",v=>sg(v,1)],
    ["q08","Faixa (bps)","",(v,x)=>`${sg(x.q08,0)} a ${sg(x.q92,0)}`],["q92","Faixa em taxa (bps)","",(v,x)=>`${sgi(-x.q92/Math.max(x.dur||1,.3),0)} a ${sgi(-x.q08/Math.max(x.dur||1,.3),0)}`],
    ["pct_univ","Posição no universo","",(v,x)=>`${br(v,0)}%${x.sinal?` <span class="tag ${x.sinal=="queda"?"t-crit":"t-ok"}">${x.sinal}</span>`:""}`],["porq","Por quê (bps)","l wrap",v=>v],["perc","% PL","",v=>br(v,2)]],P5,{ord:"pct_univ",h:380});
  const S=D.sinais.map(s=>({...s,imp:s.impacto_bps_pl}));
  const cls=t=>/Fato|Recupera|Dado/.test(t)?"t-crit":/Negócio|Ação/.test(t)?"t-warn":"t-ok";
  S.length?tabela($("#sin"),[["tipo","Tipo","l",v=>`<span class="tag ${cls(v)}">${esc(v)}</span>`],["ativo","Ativo","l"],["emissor","Emissor","l",v=>esc((v||"").slice(0,32))],["quando","Quando","l",v=>/\d{4}-/.test(v)?dt(v):esc(v)],
    ["texto","O que","l wrap",(v,s)=>s.link?`<a href="${esc(s.link)}" target="_blank" rel="noopener">${esc(v)}</a>`:esc(v)],["imp","Impacto (bps PL)","",v=>sg(v,1)],["peso","% PL","",v=>br(v,2)]],S,{h:460}):$("#sin").innerHTML='<p class="vazio">Nada fora do normal agora.</p>';
  const I=Object.entries(D.indices).filter(([k])=>/IDA|JGP|IMA|IRF/.test(k)).map(([k,v])=>({k,...v}));
  tabela($("#r_idx"),[["k","Índice","l"],["data","Data","",dt],["d1","Dia","",v=>sg(v,3,"%")],["z","Dia em desvios","",v=>sg(v,1)],["m1","1 mês","",v=>sg(v,2,"%")],["a12","12 meses","",v=>sg(v,2,"%")]],I);
  const M=A.filter(a=>a.dtaxa_1m!=null).sort((a,b)=>Math.abs(b.dtaxa_1m)-Math.abs(a.dtaxa_1m)).slice(0,12);
  tabela($("#r_mov"),[["codigo","Ativo","l"],["emissor","Emissor","l",v=>esc((v||"").slice(0,28))],["taxa","Taxa","",v=>br(v,2)],["dtaxa_1m","Δ 1m (bps)","",v=>sgi(v,0)],["ret_1m","Retorno 1m","",v=>sg(v,2,"%")],["perc_nav","% PL","",v=>br(v,2)]],M,{click:a=>{aba("ativos");detalhe(a)}})}
// ------------------------------------------------------------ agenda
function agenda(){$("#p_agenda").innerHTML=`<section><h2>Próximos eventos</h2><div class="ctl"><select id="ag_f"><option value="">tudo</option><option>Resultado</option><option>Juros</option><option>Amortiza</option><option>Vencimento</option><option>IPCA</option></select></div><div id="ag_t"></div>
  <p class="nota">Juros/amortização/vencimento: agenda pública da ANBIMA. Resultado: data estimada pela mediana dos atrasos de entrega da companhia à CVM no mesmo trimestre (entre o mais cedo e o mais tarde já vistos). Pagamento não é perda: o PU cai e o caixa entra.</p></section>
  <section><h2>Expectativas (Focus)</h2><div id="ag_fo"></div></section>`;
  const f=()=>{const q=$("#ag_f").value;tabela($("#ag_t"),[["data","Data","l",dt],["evento","Evento","l"],["pct","%","l"],["ativo","Ativo","l"],["emissor","Emissor","l",v=>esc((v||"").slice(0,34))],["peso","% PL","",v=>br(v,2)]],
    D.agenda.filter(a=>!q||(a.evento||"").includes(q)||a.ativo==q),{h:620})};$("#ag_f").onchange=f;f();
  const F=D.focus||{};$("#ag_fo").innerHTML=Object.entries(F).map(([k,v])=>`<span class="chip">${k} (${dt(v.data)}): ${Object.entries(v.anos).sort().map(([a,m])=>`${a} ${br(m,2)}%`).join(" · ")}</span>`).join(" ")||'<p class="vazio">indisponível</p>'}
// ------------------------------------------------------------ mercado
function mercado(){$("#p_mercado").innerHTML=`<div class="g2"><section><h2>Ativos de referência</h2><div id="mk_t"></div><p class="nota">Dia em desvios = variação do dia / desvio diário de 1 ano. |z| ≥ 2 é dia anormal.</p></section>
  <section><h2>Juros futuros (B3, ~15 min)</h2><div id="mk_j"></div><h3>Tesouro Direto (histórico)</h3><div id="mk_td" class="pl s"></div></section></div>
  <div class="g2"><section><h2>Índices de crédito (base 100 em 12 meses)</h2><div id="mk_ic" class="pl"></div><p class="nota">IDA (ANBIMA) e JGP, via SmartWealth.</p></section><section><h2>Fatos relevantes e comunicados (60 dias)</h2><div id="mk_doc" style="max-height:330px;overflow:auto"></div></section></div>
  <section><h2>Notícias dos maiores emissores (14 dias)</h2><div id="mk_n" class="g3"></div></section>`;
  const M=Object.entries(D.mercado).map(([k,v])=>({k,...v}));
  tabela($("#mk_t"),[["k","","l"],["preco","Último","",v=>br(v,2)],["d1","Dia","",v=>sg(v,2,"%")],["z","Desvios","",v=>sg(v,1)],["m1","1 mês","",v=>sg(v,1,"%")]],M);
  const J=D.juros||{};$("#mk_j").innerHTML=Object.keys(J).length?`<div class="sc"><table><tr><th class="l">Prazo</th>${Object.keys(J).map(k=>`<th>${k}</th><th>Δ bps</th>`).join("")}</tr>${(J.DI1||J.DAP||[]).map((_,i)=>`<tr><td class="l">${(J.DI1||J.DAP)[i].prazo}</td>${Object.values(J).map(L=>L[i]?`<td>${br(L[i].taxa,2)}</td><td>${sgi(L[i].bps,1)}</td>`:"<td></td><td></td>").join("")}</tr>`).join("")}</table></div>`:'<p class="vazio">indisponível agora</p>';
  plot("mk_td",Object.entries(D.tesouro).map(([k,v],i)=>({x:v.d,y:v.v,name:k,mode:"lines",line:{color:css(i?"--s2":"--s1"),width:1.6}})),{yaxis:{title:{text:"% a.a."},gridcolor:css("--line")}});
  plot("mk_ic",Object.entries(D.indices).filter(([k])=>/IDA|JGP/.test(k)).map(([k,v],i)=>{const n=v.d.findIndex(d=>d>=v.d[v.d.length-1].replace(/^(\d{4})/,y=>y-1));const b=v.v[Math.max(n,0)];
    return{x:v.d.slice(Math.max(n,0)),y:v.v.slice(Math.max(n,0)).map(q=>q/b*100),name:k,mode:"lines",line:{width:1.5}}}),{yaxis:{gridcolor:css("--line")}});
  const docs=Object.values(E).flatMap(e=>(e.docs||[]).map(d=>({...d,emp:e.nome,perc:e.perc}))).sort((a,b)=>b.data.localeCompare(a.data));
  $("#mk_doc").innerHTML=docs.length?`<ul class="ls">${docs.slice(0,60).map(d=>`<li><span class="tag ${/Fato|Recupera/.test(d.cat)?"t-crit":"t-ok"}">${esc(d.cat)}</span> <span class="mut">${dt(d.data)} · ${esc(d.emp.slice(0,30))} · ${br(d.perc,2)}% PL</span><br>${d.link?`<a href="${esc(d.link)}" target="_blank" rel="noopener">${esc(d.assunto||"documento")}</a>`:esc(d.assunto)}</li>`).join("")}</ul>`:'<p class="vazio">nenhum</p>';
  $("#mk_n").innerHTML=Object.entries(D.noticias).map(([r,n])=>`<div><h3>${esc((E[r]||{}).nome||n.busca)} <span class="mut">${br((E[r]||{}).perc,2)}% PL</span></h3><ul class="ls">${n.itens.map(i=>`<li>${i.alerta.length?`<span class="tag t-crit">${esc(i.alerta[0])}</span> `:""}<a href="${esc(i.link)}" target="_blank" rel="noopener">${esc(i.titulo)}</a> <span class="mut">${dt(i.data)} ${esc(i.fonte)}</span></li>`).join("")}</ul></div>`).join("")||'<p class="vazio">sem notícias</p>'}
// ------------------------------------------------------------ ativos
let SEL=null;
function ativos(){$("#p_ativos").innerHTML=`<section><h2>Carteira</h2><div class="ctl"><input id="busca" placeholder="buscar código ou emissor" aria-label="buscar"><span class="mut">clique num ativo para ver curvas, PnL, fundamentos e ação</span></div><div id="at_t"></div></section><div id="det"></div>`;
  const C=[["codigo","Ativo","l"],["emissor","Emissor","l",v=>esc((v||"").slice(0,30))],["idx","Índice","l"],["perc_nav","% PL","",v=>br(v,2)],["taxa","Taxa ind.","",v=>br(v,2)],["pct_par","% PU par","",v=>br(v,2)],
   ["duration","Duration","",v=>br(v,2)],["vol","Vol 12m","",v=>br(v,2)],["fonte","Fonte da vol","l"],["contrib_pct","% do risco","",v=>br(v,1)],["dtaxa_1m","Δ taxa 1m (bps)","",v=>sgi(v,0)],
   ["pnl_car","Carrego 12m (bps)","",v=>br(v,0)],["pnl_mtm","Taxa 12m (bps)","",v=>sg(v,0)],["pnl_res","Resíduo 12m (bps)","",v=>sg(v,0)],["giro_30d","Giro B3 30d (R$ mi)","",v=>br(v,1)]];
  const f=()=>{const q=$("#busca").value.toUpperCase();tabela($("#at_t"),C,A.filter(a=>!q||(a.codigo+" "+(a.emissor||"")).toUpperCase().includes(q)),{ord:"perc_nav",h:520,click:(a)=>detalhe(a)})};$("#busca").oninput=f;f();
  if(SEL)detalhe(SEL)}
function detalhe(a){SEL=a;const e=E[a.raiz]||{},h=a.hist||{d:[]};const ac=e.acao_credito,fu=e.fund;
  $("#det").innerHTML=`<section><h2>${esc(a.codigo)} · ${esc(a.emissor||"")} <span class="mut">${esc(a.indexador||"")} · cupom ${br(a.cupom,2)} · venc. ${dt(a.venc)} · ${br(a.perc_nav,2)}% PL</span></h2>
   ${a.lastro?`<p><span class="tag t-warn">${esc(a.lastro)}</span></p>`:""}
   <div class="kp"><div class="k"><span>Taxa indicativa</span><b>${br(a.taxa,2)}%</b><i>${dt(a.data_taxa)}</i></div><div class="k"><span>% PU par</span><b>${br(a.pct_par,2)}%</b><i>PU ${br(a.pu,2)}</i></div>
    <div class="k"><span>Vol 12m</span><b>${br(a.vol,2)}%</b><i>${esc(a.fonte||"")}</i></div><div class="k"><span>PnL 12m (bps do PU)</span><b>${br(a.pnl_total,0)}</b><i>carrego ${br(a.pnl_car,0)} · taxa ${br(a.pnl_mtm,0)} · resíduo ${br(a.pnl_res,0)}</i></div>
    ${a.hoje?`<div class="k"><span>Hoje (B3)</span><b>${sgi(a.hoje.bps,0," bps")}</b><i>${br(a.hoje.taxa,2)}% · R$ ${br(a.hoje.vol,1)} mi${a.hoje.varejo?" · varejo":""}</i></div>`:""}
    ${e.alav!=null?`<div class="k"><span>Dívida líq./EBITDA</span><b>${br(e.alav,2)}x</b><i>há 1 ano ${br(e.alav_ant,2)}x · cobertura ${br(e.cob,1)}x</i></div>`:""}</div>
   <div class="g2"><div><h3>Taxa indicativa (cheia = ANBIMA, pontilhada = implícita no PU oficial; pontos = negócios B3; ▲ resultado, ◆ fato relevante)</h3><div id="dt_t" class="pl"></div></div>
    <div><h3>PnL acumulado contra a taxa indicativa (bps do PU)</h3><div id="dt_p" class="pl"></div></div></div>
   <div class="g2"><div><h3>% do PU par</h3><div id="dt_par" class="pl s"></div></div><div><h3>Ação ${ac?esc(e.acao.ticker)+" ("+esc(e.acao.via)+")":""} × taxa (semanal)</h3>${ac?`<div id="dt_ac" class="pl s"></div>`:`<p class="vazio">${e.acao?"série curta para comparar":"emissor sem ação listada"}</p>`}</div></div>
   ${ac?`<div class="g2"><div><h3>Quem anda primeiro? Correlação ação × Δ taxa por defasagem (barras = intervalo de 95%)</h3><div id="dt_lag" class="pl s"></div><p class="nota">+k: a ação anda k semanas antes da taxa. Correlação de postos (Spearman): negativa = ação cai, taxa sobe. n=${ac.n} semanas; β contemporâneo ${br(ac.beta,2)} bps por 1% da ação.</p></div>
     <div><h3>Ação na semana × Δ taxa na semana</h3><div id="dt_sc" class="pl s"></div></div></div>`:""}
   ${fu?`<h3>Fundamentos (CVM, consolidado; data = trimestre de referência, entregue em ▲)</h3><div class="g2"><div id="dt_f1" class="pl s"></div><div id="dt_ft"></div></div>`:`<p class="vazio">${a.raiz?"emissor sem ITR/DFP na CVM (companhia fechada, securitizadora ou categoria sem demonstrações)":"sem CNPJ do emissor"}</p>`}
   <div class="g2"><div><h3>Agenda</h3>${(a.agenda||[]).length?`<ul class="ls">${a.agenda.map(g=>`<li>${esc(g.data)} · ${esc(g.evento)} ${esc(g.pct)}</li>`).join("")}</ul>`:'<p class="vazio">sem eventos previstos em 6 meses</p>'}${e.prox_resultado?`<p>Próximo resultado (${e.prox_resultado.ref.slice(5,7)}/${e.prox_resultado.ref.slice(0,4)}): ~${dt(e.prox_resultado.esperado)}, entre ${dt(e.prox_resultado.min)} e ${dt(e.prox_resultado.max)}</p>`:""}</div>
    <div><h3>Documentos CVM (60 dias)</h3>${(e.docs||[]).length?`<ul class="ls">${e.docs.map(d=>`<li><span class="mut">${dt(d.data)} ${esc(d.cat)}</span> ${d.link?`<a href="${esc(d.link)}" target="_blank" rel="noopener">${esc(d.assunto||"documento")}</a>`:esc(d.assunto)}</li>`).join("")}</ul>`:'<p class="vazio">nenhum</p>'}</div></div></section>`;
  const T=h.tA||[],tc=h.tc||[];
  const mk=(dts,ys,sym,cor,nome)=>({x:dts,y:ys,mode:"markers",name:nome,marker:{symbol:sym,size:9,color:css(cor)},hovertemplate:"%{x|%d/%m/%y}<extra>"+nome+"</extra>"});
  const em=(lista)=>{const o=[],p=[];(lista||[]).forEach(d=>{const i=h.d.findIndex(x=>x>=d);if(i>=0){o.push(h.d[i]);p.push(T[i])}});return[o,p]};
  const [rx,ry]=em((fu?fu.entrega:[]).filter(Boolean)),[fx,fy]=em(e.fr_datas);
  plot("dt_t",[{x:h.d,y:T.map((v,i)=>tc[i]===1?v:null),name:"implícita",mode:"lines",line:{color:css("--s1"),dash:"dot",width:1.4}},{x:h.d,y:T.map((v,i)=>tc[i]===0?v:null),name:"ANBIMA",mode:"lines",line:{color:css("--s1"),width:2}},
    {x:(h.neg||{}).d,y:(h.neg||{}).t,mode:"markers",name:"negócios",marker:{size:(h.neg||{v:[]}).v.map(v=>4+Math.sqrt(v)*2),color:css("--s2"),opacity:.5}},mk(rx,ry,"triangle-up","--s3","resultado"),mk(fx,fy,"diamond","--crit","fato relevante")],{yaxis:{title:{text:"% a.a."},gridcolor:css("--line")}});
  const cs=k=>{let s=0;return(h[k]||[]).map(v=>v==null?s:(s+=v))};
  plot("dt_p",[["car","carrego","--s1"],["mtm","Δ taxa indicativa","--s2"],["res","resíduo","--s4"]].map(([k,n,c])=>({x:h.d,y:cs(k),name:n,mode:"lines",line:{color:css(c),width:1.8}})),{yaxis:{title:{text:"bps"},gridcolor:css("--line")}});
  const tr5=[{x:h.d,y:h.pA,name:"% PU par ANBIMA",mode:"lines",line:{color:css("--s1"),width:1.8}}];
  if(a.prev5&&h.pA.length){const p0=h.pA[h.pA.length-1],d0=new Date(h.d[h.d.length-1]+"T12:00:00");let d5=new Date(d0),k=0;while(k<5){d5.setDate(d5.getDate()+1);if(d5.getDay()%6)k++}
    const f=b=>p0*Math.exp(b/1e4),ds=[h.d[h.d.length-1],d5.toISOString().slice(0,10)];
    tr5.push({x:ds.concat(ds.slice().reverse()),y:[p0,f(a.prev5.q92),f(a.prev5.q08),p0],fill:"toself",fillcolor:"rgba(194,65,12,.15)",line:{width:0},name:"faixa 80% em 5 dias",hoverinfo:"skip"},
             {x:ds,y:[p0,f(a.prev5.mu)],mode:"lines",line:{color:css("--s2"),dash:"dot"},name:"centro"})}
  plot("dt_par",tr5,{yaxis:{gridcolor:css("--line")},showlegend:false});
  if(ac){const L=Object.entries(ac.lags).map(([k,v])=>({k:+k,...v}));
    plot("dt_ac",[{x:ac.datas,y:ac.a.reduce((r,v)=>{r.push((r.length?r[r.length-1]:0)+v);return r},[]),name:"ação (log %, acum.)",mode:"lines",line:{color:css("--s3")}},{x:ac.datas,y:ac.t.reduce((r,v)=>{r.push((r.length?r[r.length-1]:0)+v);return r},[]),name:"taxa (bps, acum.)",yaxis:"y2",mode:"lines",line:{color:css("--s2")}}],
      {yaxis2:{overlaying:"y",side:"right",showgrid:false,title:{text:"bps"}},yaxis:{title:{text:"%"},gridcolor:css("--line")}});
    plot("dt_lag",[{type:"bar",x:L.map(l=>l.k),y:L.map(l=>l.r),marker:{color:L.map(l=>l.lo!=null&&(l.lo>0||l.hi<0)?css("--s2"):css("--s4"))},error_y:{type:"data",symmetric:false,array:L.map(l=>l.hi-l.r),arrayminus:L.map(l=>l.r-l.lo),color:css("--mut")},name:"correlação"}],{showlegend:false,xaxis:{title:{text:"defasagem (semanas)"},dtick:1},yaxis:{range:[-1,1],gridcolor:css("--line")}});
    plot("dt_sc",[{x:ac.a,y:ac.t,mode:"markers",marker:{size:5,color:css("--s1"),opacity:.6},name:"semanas"}],{showlegend:false,xaxis:{title:{text:"ação (%)"}},yaxis:{title:{text:"Δ taxa (bps)"},gridcolor:css("--line")}})}
  if(fu){plot("dt_f1",[{x:fu.ref,y:fu.dl_ebitda,name:"DL/EBITDA 12m",mode:"lines+markers",line:{color:css("--s2")}},{x:fu.ref,y:fu.cobertura,name:"EBITDA/desp. fin.",mode:"lines+markers",yaxis:"y2",line:{color:css("--s1")}}],{yaxis:{title:{text:"x"},gridcolor:css("--line")},yaxis2:{overlaying:"y",side:"right",showgrid:false}});
    const rows=fu.ref.map((r,i)=>({ref:r,ent:fu.entrega[i],rec:fu.receita[i],ebitda:fu.ebitda[i],lucro:fu.lucro[i],dl:fu.dl[i],alav:fu.dl_ebitda[i],mg:fu.margem[i]})).reverse().slice(0,12);
    tabela($("#dt_ft"),[["ref","Trimestre","l",dt],["ent","Entregue","l",dt],["rec","Receita (R$ mi)","",v=>br(v,0)],["ebitda","EBITDA","",v=>br(v,0)],["lucro","Lucro","",v=>sg(v,0)],["dl","Dív. líq.","",v=>br(v,0)],["alav","DL/EBITDA","",v=>br(v,2)],["mg","Margem 12m","",v=>br(v,1)]],rows,{h:240})}
  document.querySelectorAll("#at_t tr").forEach(tr=>tr.classList.toggle("sel",A[+tr.dataset.i]===a));$("#det").scrollIntoView({behavior:"smooth",block:"start"})}
// ------------------------------------------------------------ previsibilidade
function prev(){const P=D.prev;$("#p_prev").innerHTML=`<section><h2>O que historicamente antecipa o preço (todos os emissores da carteira juntos)</h2><div id="pv_t"></div>
  <p class="nota">r = correlação com intervalo de 95%. Útil = intervalo não cruza zero E acerto do sinal fora da amostra (último terço do tempo, inclinação estimada só com o passado) acima de 55% e acima da taxa-base.
  Sem isso o sinal é mostrado, mas não entra no impacto da aba "Hoje".</p></section>
  <section><h2>Sinais intradiários: o que move a ANBIMA do dia (estudo amplo)</h2><div id="pv_es"></div><p class="nota" id="pv_esn"></p></section>
  <div class="g2"><section><h2>Valor relativo: spread × alavancagem e duration (DI+)</h2><div id="pv_vr" class="pl"></div><div id="pv_vrt"></div></section>
  <section><h2>Por emissor: a ação antecipa a taxa?</h2><div id="pv_em"></div></section></div>`;
  const util=p=>p.lo!=null&&(p.lo>0||p.hi<0)&&p.oos&&p.oos.acerto!=null&&p.oos.acerto>55&&p.oos.acerto>p.oos.base;
  tabela($("#pv_t"),[["sinal","Sinal","l wrap"],["n","n",""],["r","r","",v=>br(v,2)],["lo","IC 95%","",(v,p)=>p.lo==null?"—":`${br(p.lo,2)} a ${br(p.hi,2)}`],["b","Efeito","",(v,p)=>v==null?(p.z_medio!=null?`|z| ${br(p.z_medio,2)}`:"—"):`${br(v,2)} <span class="mut">${esc(p.unid)}</span>`],
    ["oos","Acerto fora da amostra","",v=>v&&v.acerto!=null?`${br(v.acerto,0)}% <span class="mut">(base ${br(v.base,0)}%, n=${v.n_teste})</span>`:"—"],["id","Útil?","",(v,p)=>p.r==null?(p.z_medio>1.2?'<span class="tag t-warn">move mais que o normal</span>':'<span class="tag">não</span>'):util(p)?'<span class="tag t-ok">sim</span>':'<span class="tag">não</span>']],P);
  const Es=D.estudo;if(Es){tabela($("#pv_es"),[["0","Sinal","l wrap"],["1","Quando a ANBIMA anda","l"],["2","Tamanho","l"],["3","Acerto fora da amostra","",v=>br(v,0)+"%"],["4","Taxa-base","",v=>br(v,0)+"%"]],Es.linhas.map(l=>({...l})));
    $("#pv_esn").textContent=`${Es.papeis} debêntures/CRI/CRA mais negociadas, negócios B3 ${Es.periodo} (${Es.dias_b3} pregões). Alvo: % do PU par da ANBIMA (PU indicativo / PU par). Acerto = sinal certo no último terço do tempo, com a inclinação estimada só com o passado. Momento da ANBIMA, gap de ontem, resultado trimestral e fato relevante não mostraram poder.`}
  const E5=D.estudo5;if(E5){$("#pv_es").insertAdjacentHTML("beforebegin",`<h3>Próximos 5 dias: o que tem poder (IC = correlação de postos por semana; t)</h3><div id="pv_5"></div><p class="nota">${esc(E5.texto)}. Faixa 80%: cobertura ${E5.cobertura}%, largura mediana ${E5.largura_mediana} bps (ingênua pela vol: ${E5.largura_ingenua}).</p><h3>Mesmo dia</h3>`);
    tabela($("#pv_5"),[["0","Sinal","l wrap"],["1","Efeito em 5 dias","l"],["2","t","",v=>v?br(v,1):"—"]],E5.features.map(l=>({...l})))}
  const V=D.valor_rel;if(V){plot("pv_vr",[{x:V.pontos.map(p=>p.alav),y:V.pontos.map(p=>p.taxa),text:V.pontos.map(p=>p.cod),mode:"markers+text",textposition:"top center",textfont:{size:9},marker:{size:8,color:V.pontos.map(p=>p.res),colorscale:[[0,css("--s1")],[.5,css("--s4")],[1,css("--s2")]],showscale:false}}],
    {showlegend:false,xaxis:{title:{text:"dívida líquida / EBITDA (x)"}},yaxis:{title:{text:"spread CDI+ (%)"},gridcolor:css("--line")}});
    tabela($("#pv_vrt"),[["cod","Ativo","l"],["emissor","Emissor","l",v=>esc((v||"").slice(0,26))],["alav","DL/EBITDA","",v=>br(v,2)],["dur","Duration","",v=>br(v,2)],["taxa","Spread","",v=>br(v,2)],["res","Paga a mais (bps)","",v=>sg(v,0)]],V.pontos,{ord:"res",h:260});
    $("#pv_vr").insertAdjacentHTML("afterend",`<p class="nota">spread = ${br(V.b[0],2)} + ${br(V.b[1],3)} × alavancagem + ${br(V.b[2],3)} × duration; R² ${br(V.r2,2)}, n=${V.n}. Resíduo positivo = paga mais do que o perfil explica (ou o mercado vê algo que o balanço não mostra).</p>`)}
  else $("#pv_vr").innerHTML='<p class="vazio">poucos ativos CDI+ com alavancagem conhecida</p>';
  const L=Object.values(E).filter(e=>e.acao_credito).map(e=>{const l=e.acao_credito.lags;return{nome:e.nome,tk:e.acao.ticker,n:e.acao_credito.n,r0:l[0].r,r1:l[1].r,s1:l[1].lo!=null&&(l[1].lo>0||l[1].hi<0),beta:e.acao_credito.beta,perc:e.perc}});
  tabela($("#pv_em"),[["nome","Emissor","l",v=>esc((v||"").slice(0,28))],["tk","Ação","l"],["n","Semanas",""],["r0","r mesma semana","",v=>br(v,2)],["r1","r ação → taxa sem. seguinte","",(v,x)=>`${br(v,2)}${x.s1?' <span class="tag t-ok">sig.</span>':""}`],["beta","bps por 1%","",v=>br(v,2)],["perc","% PL","",v=>br(v,2)]],L,{ord:"perc",h:420})}
// ------------------------------------------------------------ risco
function risco(){const bt=R.backtest||{};$("#p_risco").innerHTML=`<div class="g2"><section><h2>Retorno diário da carteira oficial × VaR 95%</h2><div id="rk_b" class="pl"></div><p class="nota">${br(bt.violacoes,1)}% dos ${bt.dias} dias abaixo do VaR (esperado 5%). Ativos com série oficial ANBIMA, pesos de hoje.</p></section>
  <section><h2>Contribuição para o risco</h2><div id="rk_c" class="pl"></div></section></div>
  <div class="g2"><section><h2>Sensibilidade a fatores (12 meses, diário)</h2><div id="rk_f"></div><p class="nota">β = % do PL por 1% (ou 1 bp) do fator. |t| ≥ 2 é sensibilidade real. Correlação média entre ativos ${br(R.rho,2)}.</p></section>
  <section><h2>Correlação entre os maiores ativos (retornos diários oficiais)</h2><div id="rk_m" class="pl"></div></section></div>`;
  const S=R.serie||{d:[],ret:[]};plot("rk_b",[{x:S.d,y:S.ret,type:"bar",name:"retorno diário (%)",marker:{color:S.ret.map(v=>v<S.var?css("--crit"):css("--s4"))}},{x:[S.d[0],S.d[S.d.length-1]],y:[S.var,S.var],mode:"lines",name:"VaR 95%",line:{color:css("--crit"),dash:"dash"}}],{yaxis:{title:{text:"%"},gridcolor:css("--line")}});
  const C=A.filter(a=>a.contrib_pct!=null).sort((a,b)=>b.contrib_pct-a.contrib_pct).slice(0,20).reverse();
  plot("rk_c",[{type:"bar",orientation:"h",y:C.map(a=>a.codigo),x:C.map(a=>a.contrib_pct),marker:{color:css("--s1")},name:"% do risco"}],{showlegend:false,margin:{l:80,r:12,t:10,b:30},xaxis:{title:{text:"% do risco"},gridcolor:css("--line")}});
  tabela($("#rk_f"),[["fator","Fator","l"],["beta","β","",v=>br(v,4)],["t","t","",v=>br(v,1)],["r","r","",v=>br(v,2)],["r2","R²","",v=>br(v,2)],["n","dias",""]],R.fatores||[],{ord:"r2"});
  const M=R.corr;plot("rk_m",[{type:"heatmap",z:M.m,x:M.cods,y:M.cods,zmin:-1,zmax:1,colorscale:[[0,css("--s1")],[.5,css("--card")],[1,css("--s2")]]}],{showlegend:false,margin:{l:70,r:10,t:10,b:70},xaxis:{tickfont:{size:9}},yaxis:{tickfont:{size:9}}})}
// ------------------------------------------------------------ método
function metodo(){$("#p_metodo").innerHTML=`<section><h2>Como cada número é calculado</h2><ul class="ls">
 <li><b>Retorno do dia</b> = (PU indicativo + pagamentos) / PU indicativo anterior − 1. Amortização e juros entram como pagamento: o PU cai, mas não é perda.</li>
 <li><b>PnL contra a taxa indicativa</b>: carrego = rendimento do PU par (contrato, com o índice) + diferença entre a taxa indicativa e o cupom; Δ taxa = −duration modificada × variação da taxa indicativa; resíduo = o que sobra (convexidade, evento não informado, marcação).</li>
 <li><b>Taxa antes do arquivo diário da ANBIMA</b> = taxa implícita no PU oficial (o cálculo da ANBIMA invertido, dia a dia, para trás). Nesses dias o resíduo não é medido.</li>
 <li><b>Vol</b> = desvio dos retornos diários oficiais em 12 meses × √252. Sem série oficial: negócios semanais (SND ou B3) com o ruído de negócio descontado; sem nada: mediana dos pares de mesmo índice e duration.</li>
 <li><b>VaR</b> 95% paramétrico: vols × correlação dos retornos oficiais (≥ 60 dias em comum, encolhida 30% para a média). Backtest com os pesos de hoje nos últimos 250 dias.</li>
 <li><b>Fundamentos</b>: ITR/DFP da CVM (consolidado). EBITDA = EBIT + depreciação/amortização da DFC. Dívida = empréstimos e financiamentos (circulante + não circulante); caixa = caixa + aplicações. Data de entrega = 1ª entrega à CVM (o que o mercado sabia).</li>
 <li><b>Previsibilidade</b>: correlações com intervalo de 95% (Fisher) e teste fora da amostra (ajuste nos 2/3 mais antigos, teste no terço final). Só sinais que passam nos dois viram impacto estimado.</li>
 <li><b>Fontes</b>: ANBIMA Data (páginas públicas), ANBIMA arquivo diário, B3 (negócios de balcão, futuros), SND, CVM (ITR, DFP, IPE, FCA, securitizadoras), Yahoo Finance, SmartWealth (índices IDA/JGP/IMA, leitura local), BCB (CDI, Focus), Tesouro Direto, IBGE, Google Notícias.</li></ul></section>`}
function tudo(){resumo();agenda();mercado();ativos();prev();risco();metodo();let a="resumo";try{a=localStorage.getItem("vm_aba")||a}catch(e){};aba(ABAS.some(x=>x[0]==a)?a:"resumo")}
tudo();
</script>"""


if __name__ == "__main__":
    args = sys.argv[1:]
    if args and args[0] == "modelo":
        modelo()
    else:
        arq = Path(args[0]) if args and args[0].lower().endswith((".xlsx", ".xls")) else PASTA / "carteira_modelo.xlsx"
        if not arq.exists():
            sys.exit(f"carteira não encontrada: {arq}  (crie uma com: python vol_monitor.py modelo)")
        rodar(arq) if "rodar" in args else uma_vez(arq)
