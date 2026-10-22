"""
Histórico oficial ANBIMA por ativo (desde o início da marcação): PU par (PU histórico), PU indicativo e eventos.
Fonte: a página pública de cada ativo na ANBIMA Data, gráfico "PU" com o filtro "Total" — o Edge abre a página como um usuário,
clica em "Total" e guarda o que a própria página recebeu para desenhar o gráfico (nenhuma chamada fora da página, nenhum token reusado).
Cache: dados/anbima_hist/<codigo>.parquet (refaz se tiver mais de 1 dia).
Uso: python anbima_hist.py CODIGO [CODIGO ...]
"""
import json
import re
import sys
import time
from pathlib import Path

import pandas as pd

PASTA = Path(__file__).resolve().parent / "dados" / "anbima_hist"
PASTA.mkdir(parents=True, exist_ok=True)
GANCHO = r"""
window.__cap=[];
const of=window.fetch; window.fetch=async(...a)=>{const r=await of(...a);try{const u=(a[0].url||a[0])+"";if(/grafico-pu/.test(u)){r.clone().text().then(t=>window.__cap.push({u,t}))}}catch(e){}return r};
const ox=XMLHttpRequest.prototype.open,os=XMLHttpRequest.prototype.send;
XMLHttpRequest.prototype.open=function(m,u){this.__u=u;return ox.apply(this,arguments)};
XMLHttpRequest.prototype.send=function(){this.addEventListener("load",()=>{if(/grafico-pu/.test(this.__u))window.__cap.push({u:this.__u,t:this.responseText})});return os.apply(this,arguments)};
"""


def tipo_pagina(cod):
    return "certificado-de-recebiveis" if cod.startswith(("CRA", "CRI")) or re.match(r"^\d{2}[A-L]\d", cod) else "debentures"


def coleta(cods, log=print):
    """{codigo: DataFrame[data, pu_par, pu_ind, evento, valor_evento]} — usa o cache do dia quando existe."""
    out, falta = {}, []
    for c in cods:
        f = PASTA / f"{c}.parquet"
        if f.exists() and time.time() - f.stat().st_mtime < 20 * 3600:
            out[c] = pd.read_parquet(f)
        else:
            falta.append(c)
    if not falta:
        return out
    from selenium import webdriver
    o = webdriver.EdgeOptions()
    o.add_argument("--headless=new")
    o.add_argument("--window-size=1400,1000")
    d = webdriver.Edge(options=o)
    try:
        for k, c in enumerate(falta):
            try:
                d.get(f"https://data.anbima.com.br/{tipo_pagina(c)}/{c}/precos")
                for _ in range(20):                                  # espera o gráfico aparecer
                    time.sleep(0.5)
                    if d.execute_script("return [...document.querySelectorAll('li')].some(e=>e.innerText.trim()=='Total')"):
                        break
                d.execute_script(GANCHO)
                d.execute_script("[...document.querySelectorAll('li')].find(e=>e.innerText.trim()=='Total').click()")
                t = None
                for _ in range(30):
                    time.sleep(0.5)
                    t = d.execute_script("return window.__cap.length ? window.__cap[window.__cap.length-1].t : null")
                    if t:
                        break
                if not t:
                    log(f"   {c}: sem gráfico")
                    continue
                pus = json.loads(t).get("pus", [])
                x = pd.DataFrame([{"data": p["data"], "pu_par": p.get("valor_pu_historico"), "pu_ind": p.get("valor_pu_indicativo"),
                                   "evento": "; ".join(e.get("tipo", "") for e in p.get("eventos") or []),
                                   "valor_evento": sum(float(e.get("valor") or 0) for e in p.get("eventos") or [])} for p in pus])
                if not len(x):
                    continue
                x["data"] = pd.to_datetime(x.data)
                x["pu_par"], x["pu_ind"] = pd.to_numeric(x.pu_par, errors="coerce"), pd.to_numeric(x.pu_ind, errors="coerce")
                x.to_parquet(PASTA / f"{c}.parquet")
                out[c] = x
                log(f"   {k + 1}/{len(falta)} {c}: {len(x)} dias, {x.pu_ind.notna().sum()} com PU indicativo")
            except Exception as e:
                log(f"   {c}: erro {type(e).__name__}")
    finally:
        d.quit()
    return out


def resumo(cods, log=print):
    """Topo da página de preços (último dia): taxa indicativa, duration, % PU par e PU indicativo — para ativos fora do arquivo diário
    de debêntures (CRI/CRA). Cache do dia em dados/anbima_hist/resumo_AAAAMMDD.json."""
    f = PASTA / f"resumo_{time.strftime('%Y%m%d')}.json"
    out = json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
    falta = [c for c in cods if c not in out]
    if falta:
        from selenium import webdriver
        o = webdriver.EdgeOptions()
        o.add_argument("--headless=new")
        d = webdriver.Edge(options=o)
        num = lambda v: float(v.replace(".", "").replace(",", ".")) if v else None
        try:
            for c in falta:
                try:
                    d.get(f"https://data.anbima.com.br/{tipo_pagina(c)}/{c}/precos")
                    time.sleep(6)
                    t = d.find_element("tag name", "body").text
                    g = lambda p: (re.search(p, t) or [None, None])[1]
                    out[c] = {"data": pd.to_datetime(g(r"Data de refer.ncia\n(\d{2}/\d{2}/\d{4})\nPU [Ii]ndicativo"), dayfirst=True).strftime("%Y-%m-%d")
                              if g(r"Data de refer.ncia\n(\d{2}/\d{2}/\d{4})\nPU [Ii]ndicativo") else None,
                              "taxa": num(g(r"Taxa indicativa\n([\d.,]+) ?%")), "duration": (num(g(r"Duration\n([\d.,]+) dias")) or 0) / 252 or None,
                              "pct": num(g(r"% PU par\n([\d.,]+) ?%"))}
                except Exception as e:
                    log(f"   resumo {c}: {type(e).__name__}")
        finally:
            d.quit()
        f.write_text(json.dumps(out), encoding="utf-8")
    return {c: v for c, v in out.items() if c in cods and v.get("taxa") is not None and v.get("data")}


if __name__ == "__main__":
    r = coleta(sys.argv[1:])
    for c, x in r.items():
        print(c, len(x), x.pu_ind.notna().sum(), x.data.min().date(), x.data.max().date())
