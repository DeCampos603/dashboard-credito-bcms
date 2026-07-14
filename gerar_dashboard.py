# -*- coding: utf-8 -*-
"""
Gera um dashboard HTML autocontido de CRÉDITO DISPONÍVEL do BCMS
(UASG 160329 - OGU e 167329 - Fundo do Exército) a partir do export do
Tesouro Gerencial 'CRÉDITO DISP.xlsx' publicado no Google Drive.

- Baixa o xlsx público do Drive (ou usa --local para testar com arquivo local).
- Lê a aba 'CRÉDITO DISP' (cabeçalho linha 8), valida o layout por nome de coluna,
  filtra BCMS e calcula KPIs, quebras por Ação/ND e detalhe por NC.
- Acumula um snapshot diário em data/history.json (para o gráfico de tendência).
- Escreve site/index.html (self-contained: CSS + SVG + tabela) e site/data/history.json.

Uso:
    python gerar_dashboard.py                 # baixa do Drive (FileId padrão / env DRIVE_FILE_ID)
    python gerar_dashboard.py --local X.xlsx  # usa arquivo local (teste)
    python gerar_dashboard.py --date 2026-07-14  # força a data do snapshot (default: hoje)
"""
import os, sys, json, argparse, datetime, urllib.request, tempfile, html, math
import openpyxl

HDR_ROW, DATA_ROW = 8, 9
ALVOS = [("160329", "BCMS · OGU"), ("167329", "BCMS · Fundo do Exército")]
DEFAULT_FILE_ID = "1Jv546wpWQSFAlep3oLRAg29hVy86iJxJ"
HERE = os.path.dirname(os.path.abspath(__file__))
SITE = os.path.join(HERE, "site")
DATA = os.path.join(HERE, "data")
HISTFILE = os.path.join(DATA, "history.json")

# ---------------- leitura ----------------
def norm(s): return str(s).strip().upper() if s is not None else ""

def to_num(v):
    if v is None: return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace("'", "")
    if s in ("", "-", "-9", "NAO SE APLICA", "NÃO SE APLICA"): return 0.0
    try: return float(s)
    except ValueError:
        try: return float(s.replace(".", "").replace(",", "."))
        except ValueError: return 0.0

def disp(v):
    s = "" if v is None else str(v).strip().replace("'", "")
    return "" if s in ("-9", "NAO SE APLICA", "NÃO SE APLICA") else s

def baixar(file_id):
    url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (dashboard-bcms)"})
    tmp = os.path.join(tempfile.gettempdir(), "credito_disp_download.xlsx")
    with urllib.request.urlopen(req, timeout=120) as r, open(tmp, "wb") as f:
        f.write(r.read())
    if os.path.getsize(tmp) < 1000:
        raise SystemExit("Download muito pequeno — verifique o compartilhamento público do arquivo.")
    return tmp

# ---------------- ETL ----------------
def etl(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["CRÉDITO DISP"] if "CRÉDITO DISP" in wb.sheetnames else wb.active
    hdr = {}
    for c in range(1, ws.max_column + 1):
        n = norm(ws.cell(HDR_ROW, c).value)
        if n and n not in hdr: hdr[n] = c
    def col(name):
        c = hdr.get(name)
        if not c: raise SystemExit(f"Layout mudou: coluna '{name}' não encontrada.")
        return c
    C = dict(prov=col("PROVISAO RECEBIDA"), cred=col("CREDITO DISPONIVEL"),
             emp=col("DESPESAS EMPENHADAS"), liq=col("DESPESAS LIQUIDADAS"), pag=col("DESPESAS PAGAS"))
    res = {}
    for cod, nome in ALVOS:
        res[cod] = dict(nome=nome, prov=0.0, cred=0.0, emp=0.0, liq=0.0, pag=0.0,
                        n=0, por_acao={}, por_nd={}, nd_nome={}, linhas=[])
    for r in range(DATA_ROW, ws.max_row + 1):
        cod = norm(ws.cell(r, 3).value)
        if cod not in res: continue
        d = res[cod]
        vals = {k: to_num(ws.cell(r, C[k]).value) for k in C}
        for k in C: d[k] += vals[k]
        d["n"] += 1
        acao = disp(ws.cell(r, 6).value) or "(s/ ação)"
        ndc = disp(ws.cell(r, 9).value) or "(s/ ND)"
        d["por_acao"][acao] = d["por_acao"].get(acao, 0.0) + vals["cred"]
        d["por_nd"][ndc] = d["por_nd"].get(ndc, 0.0) + vals["cred"]
        d["nd_nome"][ndc] = disp(ws.cell(r, 10).value)
        if round(vals["cred"], 2) != 0:
            d["linhas"].append(dict(
                emit=disp(ws.cell(r, 2).value), nc=disp(ws.cell(r, 5).value), acao=acao,
                pi=disp(ws.cell(r, 7).value), nd=ndc, nd_desc=disp(ws.cell(r, 10).value),
                obj=disp(ws.cell(r, 11).value), op=disp(ws.cell(r, 12).value),
                cred=vals["cred"], emp=vals["emp"], liq=vals["liq"], pag=vals["pag"]))
    for cod in res:
        res[cod]["linhas"].sort(key=lambda x: x["cred"], reverse=True)
    return res

# ---------------- histórico ----------------
def atualizar_historico(res, data_str):
    hist = []
    if os.path.exists(HISTFILE):
        try:
            with open(HISTFILE, encoding="utf-8") as f: hist = json.load(f)
        except Exception: hist = []
    tot = {k: sum(res[c][k] for c, _ in ALVOS) for k in ("prov", "cred", "emp", "liq", "pag")}
    snap = {"data": data_str,
            "total": {k: round(tot[k], 2) for k in tot},
            **{c: {k: round(res[c][k], 2) for k in ("prov", "cred", "emp", "liq", "pag")} for c, _ in ALVOS}}
    hist = [h for h in hist if h.get("data") != data_str]
    hist.append(snap)
    hist.sort(key=lambda h: h["data"])
    os.makedirs(DATA, exist_ok=True)
    with open(HISTFILE, "w", encoding="utf-8") as f:
        json.dump(hist, f, ensure_ascii=False, indent=1)
    return hist

# ---------------- formatação ----------------
def brl(v):
    s = f"{abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-R$ " if v < 0 else "R$ ") + s

def pct(a, b): return (100.0 * a / b) if b else 0.0

def esc(s): return html.escape(str(s))

# ---------------- SVG charts ----------------
PAL = ["#2563eb", "#0891b2", "#7c3aed", "#059669", "#d97706", "#dc2626", "#4f46e5", "#0d9488"]

def barras_h(itens, titulo, cor="#2563eb", max_itens=10):
    itens = [(k, v) for k, v in itens if round(v, 2) != 0]
    itens = sorted(itens, key=lambda x: abs(x[1]), reverse=True)[:max_itens]
    if not itens:
        return f'<div class="chart"><h3>{esc(titulo)}</h3><p class="vazio">sem valores</p></div>'
    vmax = max(abs(v) for _, v in itens) or 1
    row_h, lbl_w, bar_w, pad = 30, 130, 300, 8
    W = lbl_w + bar_w + 90
    H = pad * 2 + row_h * len(itens)
    rows = []
    for i, (k, v) in enumerate(itens):
        y = pad + i * row_h
        w = max(2, abs(v) / vmax * bar_w)
        c = "#dc2626" if v < 0 else cor
        rows.append(
            f'<text x="{lbl_w-6}" y="{y+row_h/2+4}" text-anchor="end" class="blbl">{esc(k)}</text>'
            f'<rect x="{lbl_w}" y="{y+5}" width="{w:.1f}" height="{row_h-12}" rx="3" fill="{c}"><title>{esc(k)}: {brl(v)}</title></rect>'
            f'<text x="{lbl_w+w+6:.1f}" y="{y+row_h/2+4}" class="bval">{brl(v)}</text>')
    return (f'<div class="chart"><h3>{esc(titulo)}</h3>'
            f'<svg viewBox="0 0 {W} {H}" role="img" preserveAspectRatio="xMinYMin meet">{"".join(rows)}</svg></div>')

def barras_grupo(cats, series, titulo):
    # series: list of (nome, cor, {cat:val})
    if not cats:
        return ""
    vmax = max([abs(s[2].get(c, 0)) for s in series for c in cats] + [1])
    gw, pad_l, pad_b, pad_t = 46, 46, 40, 10
    W = pad_l + gw * len(cats) + 20
    H = 210
    plot_h = H - pad_b - pad_t
    nser = len(series)
    bw = (gw - 10) / nser
    el = []
    # eixo y (3 marcas)
    for t in range(4):
        val = vmax * t / 3
        y = pad_t + plot_h - (val / vmax * plot_h)
        el.append(f'<line x1="{pad_l}" y1="{y:.1f}" x2="{W-6}" y2="{y:.1f}" class="grid"/>')
        el.append(f'<text x="{pad_l-4}" y="{y+3:.1f}" text-anchor="end" class="axis">{val/1000:.0f}k</text>')
    for ci, cat in enumerate(cats):
        gx = pad_l + ci * gw
        for si, (nome, cor, d) in enumerate(series):
            v = d.get(cat, 0)
            bh = abs(v) / vmax * plot_h
            x = gx + 5 + si * bw
            y = pad_t + plot_h - bh
            el.append(f'<rect x="{x:.1f}" y="{y:.1f}" width="{bw-2:.1f}" height="{bh:.1f}" fill="{cor}" rx="1.5"><title>{esc(cat)} — {esc(nome)}: {brl(v)}</title></rect>')
        el.append(f'<text x="{gx+gw/2:.1f}" y="{H-pad_b+16}" text-anchor="middle" class="axis">{esc(cat)}</text>')
    leg = " ".join(f'<span class="lg"><i style="background:{c}"></i>{esc(n)}</span>' for n, c, _ in series)
    return (f'<div class="chart"><h3>{esc(titulo)}</h3>'
            f'<svg viewBox="0 0 {W} {H}" role="img" preserveAspectRatio="xMinYMin meet">{"".join(el)}</svg>'
            f'<div class="legenda">{leg}</div></div>')

def linha_tempo(hist, titulo):
    pts = [(h["data"], h["total"]["cred"]) for h in hist]
    if len(pts) < 1:
        return ""
    W, H, pad_l, pad_b, pad_t, pad_r = 640, 220, 60, 34, 14, 16
    plot_w, plot_h = W - pad_l - pad_r, H - pad_b - pad_t
    vals = [v for _, v in pts]
    vmin, vmax = min(vals + [0]), max(vals + [1])
    rng = (vmax - vmin) or 1
    n = len(pts)
    def X(i): return pad_l + (plot_w * (i / (n - 1)) if n > 1 else plot_w / 2)
    def Y(v): return pad_t + plot_h - ((v - vmin) / rng * plot_h)
    el = []
    for t in range(4):
        val = vmin + rng * t / 3
        y = Y(val)
        el.append(f'<line x1="{pad_l}" y1="{y:.1f}" x2="{W-pad_r}" y2="{y:.1f}" class="grid"/>')
        el.append(f'<text x="{pad_l-6}" y="{y+3:.1f}" text-anchor="end" class="axis">{val/1000:.0f}k</text>')
    if n > 1:
        d = "M" + " L".join(f"{X(i):.1f},{Y(v):.1f}" for i, (_, v) in enumerate(pts))
        el.append(f'<path d="{d}" fill="none" stroke="#2563eb" stroke-width="2.5"/>')
        area = f"M{X(0):.1f},{Y(vmin):.1f} L" + " L".join(f"{X(i):.1f},{Y(v):.1f}" for i, (_, v) in enumerate(pts)) + f" L{X(n-1):.1f},{Y(vmin):.1f} Z"
        el.append(f'<path d="{area}" fill="#2563eb" opacity="0.08"/>')
    step = max(1, n // 6)
    for i, (dt, v) in enumerate(pts):
        el.append(f'<circle cx="{X(i):.1f}" cy="{Y(v):.1f}" r="3.2" fill="#2563eb"><title>{esc(dt)}: {brl(v)}</title></circle>')
        if i % step == 0 or i == n - 1:
            lab = dt[8:10] + "/" + dt[5:7]
            el.append(f'<text x="{X(i):.1f}" y="{H-pad_b+16}" text-anchor="middle" class="axis">{lab}</text>')
    nota = "" if n > 1 else '<p class="vazio">o histórico começa hoje — a curva aparece a partir do 2º dia.</p>'
    return (f'<div class="chart wide"><h3>{esc(titulo)}</h3>'
            f'<svg viewBox="0 0 {W} {H}" role="img" preserveAspectRatio="xMinYMin meet">{"".join(el)}</svg>{nota}</div>')

# ---------------- HTML ----------------
def card(label, valor, sub="", classe=""):
    return (f'<div class="kpi {classe}"><div class="kpi-l">{esc(label)}</div>'
            f'<div class="kpi-v">{esc(valor)}</div>'
            f'{f"<div class=\"kpi-s\">{esc(sub)}</div>" if sub else ""}</div>')

def barra_exec(label, valor, base, cor):
    p = pct(valor, base)
    return (f'<div class="exec"><div class="exec-top"><span>{esc(label)}</span><span>{p:.1f}%</span></div>'
            f'<div class="exec-bar"><div style="width:{min(p,100):.1f}%;background:{cor}"></div></div></div>')

def tabela(cod, d):
    linhas = d["linhas"]
    body = []
    for L in linhas:
        neg = ' class="neg"' if L["cred"] < 0 else ''
        body.append(
            f'<tr><td>{esc(L["acao"])}</td><td>{esc(L["nd"])}</td>'
            f'<td class="obj" title="{esc(L["obj"])}">{esc(L["obj"][:80])}</td>'
            f'<td class="num">{brl(L["emp"])}</td><td class="num">{brl(L["liq"])}</td>'
            f'<td class="num">{brl(L["pag"])}</td><td class="num strong"{neg}>{brl(L["cred"])}</td></tr>')
    return (f'<div class="tab-wrap" id="tab-{cod}" style="{"" if cod==ALVOS[0][0] else "display:none"}">'
            f'<table class="det"><thead><tr><th>Ação</th><th>ND</th><th>Objeto (NC)</th>'
            f'<th class="num">Empenhado</th><th class="num">Liquidado</th><th class="num">Pago</th>'
            f'<th class="num">Crédito Disp.</th></tr></thead><tbody>{"".join(body)}</tbody></table>'
            f'<p class="tab-nota">{len(linhas)} linhas com Crédito Disponível ≠ 0 · Total {brl(d["cred"])}</p></div>')

def gerar_html(res, hist, data_str):
    tot = {k: sum(res[c][k] for c, _ in ALVOS) for k in ("prov", "cred", "emp", "liq", "pag", "n")}
    ger = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    dbr = data_str[8:10] + "/" + data_str[5:7] + "/" + data_str[0:4]

    # KPIs topo (total)
    kpis = (card("Provisão Recebida", brl(tot["prov"])) +
            card("Empenhado", brl(tot["emp"]), f'{pct(tot["emp"],tot["prov"]):.1f}% do recebido') +
            card("Liquidado", brl(tot["liq"])) +
            card("Pago", brl(tot["pag"])) +
            card("Crédito Disponível", brl(tot["cred"]), "saldo não empenhado", "destaque"))

    # cards por UASG
    blocos = []
    for cod, nome in ALVOS:
        d = res[cod]
        exec_html = (barra_exec("Empenhado / Recebido", d["emp"], d["prov"], "#2563eb") +
                     barra_exec("Liquidado / Empenhado", d["liq"], d["emp"], "#0891b2") +
                     barra_exec("Pago / Liquidado", d["pag"], d["liq"], "#059669"))
        blocos.append(
            f'<div class="uasg-card"><div class="uasg-h"><b>{esc(cod)}</b> · {esc(nome)}</div>'
            f'<div class="uasg-kpis">'
            f'<div><span>Recebido</span>{brl(d["prov"])}</div>'
            f'<div><span>Empenhado</span>{brl(d["emp"])}</div>'
            f'<div class="dest"><span>Crédito Disp.</span>{brl(d["cred"])}</div></div>'
            f'<div class="execs">{exec_html}</div></div>')

    # charts crédito por ação/ND (combinado)
    acao_comb, nd_comb, nd_nome = {}, {}, {}
    for cod, _ in ALVOS:
        for k, v in res[cod]["por_acao"].items(): acao_comb[k] = acao_comb.get(k, 0) + v
        for k, v in res[cod]["por_nd"].items(): nd_comb[k] = nd_comb.get(k, 0) + v
        nd_nome.update(res[cod]["nd_nome"])
    nd_lbl = {f'{k} {nd_nome.get(k,"")[:16]}'.strip(): v for k, v in nd_comb.items()}
    ch_acao = barras_h(acao_comb.items(), "Crédito Disponível por Ação (R$)", "#2563eb")
    ch_nd = barras_h(nd_lbl.items(), "Crédito Disponível por Natureza de Despesa (R$)", "#7c3aed")
    # estágios por UASG
    series = [
        ("Empenhado", "#2563eb", {res[c]["nome"].split("·")[1].strip(): res[c]["emp"] for c, _ in ALVOS}),
        ("Liquidado", "#0891b2", {res[c]["nome"].split("·")[1].strip(): res[c]["liq"] for c, _ in ALVOS}),
        ("Pago", "#059669", {res[c]["nome"].split("·")[1].strip(): res[c]["pag"] for c, _ in ALVOS}),
    ]
    cats = [res[c]["nome"].split("·")[1].strip() for c, _ in ALVOS]
    ch_estagios = barras_grupo(cats, series, "Estágios da despesa por fonte (R$)")
    ch_trend = linha_tempo(hist, "Tendência — Crédito Disponível total (R$)")

    # abas tabela
    abas = "".join(f'<button class="aba{ " on" if i==0 else ""}" onclick="mostrar(\'{cod}\')">{esc(cod)} {esc(nome.split("·")[1].strip())}</button>'
                   for i, (cod, nome) in enumerate(ALVOS))
    tabs = "".join(tabela(cod, res[cod]) for cod, _ in ALVOS)

    css = """
:root{--bg:#f4f6fb;--card:#fff;--ink:#0f172a;--mut:#64748b;--bd:#e2e8f0;--dest:#fef9c3;--destbd:#eab308}
@media(prefers-color-scheme:dark){:root{--bg:#0b1220;--card:#111a2e;--ink:#e7edf7;--mut:#93a4bd;--bd:#22304d;--dest:#3b3410;--destbd:#a3860f}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.45 -apple-system,Segoe UI,Roboto,Arial,sans-serif}
.wrap{max-width:1180px;margin:0 auto;padding:20px 18px 60px}
header{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:flex-end;gap:10px;margin-bottom:6px}
h1{font-size:22px;margin:0}.sub{color:var(--mut);font-size:13px}
.pos{background:#2563eb;color:#fff;padding:6px 12px;border-radius:8px;font-size:13px;font-weight:600}
h2{font-size:15px;text-transform:uppercase;letter-spacing:.04em;color:var(--mut);margin:26px 0 10px;border-bottom:1px solid var(--bd);padding-bottom:6px}
.kpis{display:grid;grid-template-columns:repeat(5,1fr);gap:12px}
.kpi{background:var(--card);border:1px solid var(--bd);border-radius:12px;padding:14px}
.kpi-l{color:var(--mut);font-size:12px}.kpi-v{font-size:19px;font-weight:700;margin-top:4px}.kpi-s{color:var(--mut);font-size:11px;margin-top:2px}
.kpi.destaque{background:var(--dest);border-color:var(--destbd)}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.uasg-card{background:var(--card);border:1px solid var(--bd);border-radius:12px;padding:14px}
.uasg-h{font-size:15px;margin-bottom:10px}.uasg-kpis{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:10px}
.uasg-kpis>div{display:flex;flex-direction:column;font-weight:700}.uasg-kpis span{color:var(--mut);font-size:11px;font-weight:400}
.uasg-kpis .dest{color:#a16207}
.execs{display:flex;flex-direction:column;gap:7px}.exec-top{display:flex;justify-content:space-between;font-size:12px;color:var(--mut)}
.exec-bar{height:8px;background:var(--bd);border-radius:5px;overflow:hidden}.exec-bar>div{height:100%}
.charts{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.chart{background:var(--card);border:1px solid var(--bd);border-radius:12px;padding:14px}.chart.wide{grid-column:1/-1}
.chart h3{font-size:14px;margin:0 0 10px}.chart svg{width:100%;height:auto}
.blbl{font-size:11px;fill:var(--mut)}.bval{font-size:11px;fill:var(--ink);font-weight:600}
.axis{font-size:10px;fill:var(--mut)}.grid{stroke:var(--bd);stroke-width:1}
.legenda{display:flex;gap:14px;margin-top:6px;font-size:12px;color:var(--mut)}.lg i{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:4px}
.vazio{color:var(--mut);font-size:12px;font-style:italic}
.abas{display:flex;gap:8px;margin-bottom:10px}.aba{background:var(--card);border:1px solid var(--bd);color:var(--ink);padding:7px 12px;border-radius:8px;cursor:pointer;font-size:13px}
.aba.on{background:#2563eb;color:#fff;border-color:#2563eb}
.tab-wrap{overflow-x:auto;background:var(--card);border:1px solid var(--bd);border-radius:12px}
table.det{border-collapse:collapse;width:100%;font-size:13px}
.det th,.det td{padding:7px 10px;border-bottom:1px solid var(--bd);text-align:left}
.det th{background:var(--bg);position:sticky;top:0}.det .num{text-align:right;white-space:nowrap}
.det .strong{font-weight:700}.det .neg{color:#dc2626}.det .obj{max-width:340px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--mut)}
.tab-nota{color:var(--mut);font-size:12px;padding:8px 10px;margin:0}
footer{margin-top:34px;color:var(--mut);font-size:12px;text-align:center;line-height:1.7}
@media(max-width:860px){.kpis{grid-template-columns:repeat(2,1fr)}.grid2,.charts{grid-template-columns:1fr}}
"""
    js = """
function mostrar(cod){document.querySelectorAll('.tab-wrap').forEach(function(t){t.style.display='none'});
document.getElementById('tab-'+cod).style.display='block';
document.querySelectorAll('.aba').forEach(function(b){b.classList.remove('on')});event.target.classList.add('on');}
"""
    return f"""<!doctype html><html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Crédito Disponível — BCMS</title><style>{css}</style></head>
<body><div class="wrap">
<header><div><h1>Crédito Disponível — BCMS</h1>
<div class="sub">UASG 160329 (OGU) e 167329 (Fundo do Exército) · Fonte: Tesouro Gerencial / SIAFI</div></div>
<div class="pos">Posição {dbr}</div></header>

<h2>Consolidado BCMS</h2>
<div class="kpis">{kpis}</div>

<h2>Por fonte</h2>
<div class="grid2">{"".join(blocos)}</div>

<h2>Onde está o crédito disponível</h2>
<div class="charts">{ch_acao}{ch_nd}{ch_estagios}{ch_trend}</div>

<h2>Detalhe por NC (crédito disponível ≠ 0)</h2>
<div class="abas">{abas}</div>
{tabs}

<footer>Gerado automaticamente em {ger} · Crédito Disponível = Provisão Recebida − Despesas Empenhadas<br>
Fonte: CRÉDITO DISP.xlsx (Google Drive, atualizado diariamente) · Valores negativos = anulações de descentralização</footer>
</div><script>{js}</script></body></html>"""

# ---------------- main ----------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--local", help="caminho de um xlsx local (teste)")
    ap.add_argument("--date", help="data do snapshot YYYY-MM-DD (default: hoje)")
    ap.add_argument("--file-id", default=(os.environ.get("DRIVE_FILE_ID") or DEFAULT_FILE_ID))
    args = ap.parse_args()

    data_str = args.date or datetime.date.today().isoformat()
    path = args.local if args.local else baixar(args.file_id)
    print("Fonte:", path)
    res = etl(path)
    hist = atualizar_historico(res, data_str)
    html_out = gerar_html(res, hist, data_str)

    os.makedirs(SITE, exist_ok=True)
    os.makedirs(os.path.join(SITE, "data"), exist_ok=True)
    with open(os.path.join(SITE, "index.html"), "w", encoding="utf-8") as f:
        f.write(html_out)
    # expõe o histórico junto do site
    with open(os.path.join(SITE, "data", "history.json"), "w", encoding="utf-8") as f:
        json.dump(hist, f, ensure_ascii=False, indent=1)

    tot = {k: sum(res[c][k] for c, _ in ALVOS) for k in ("prov", "cred", "emp", "liq", "pag")}
    print(f"OK -> {os.path.join(SITE,'index.html')}")
    print(f"Crédito Disponível total = {brl(tot['cred'])} | histórico: {len(hist)} dia(s)")

if __name__ == "__main__":
    main()
