# -*- coding: utf-8 -*-
"""
Gera um dashboard HTML autocontido de CRÉDITO DISPONÍVEL do BCMS
(UASG 160329 - OGU e 167329 - Fundo do Exército) a partir do export do
Tesouro Gerencial 'CRÉDITO DISP.xlsx' publicado no Google Drive.

- Baixa o xlsx público do Drive (ou usa --local para testar com arquivo local).
- Lê a aba 'CRÉDITO DISP' (cabeçalho linha 8), valida o layout por nome de coluna,
  filtra BCMS e calcula KPIs, quebras por Ação/ND e detalhe por NC.
- VALIDAÇÃO anti-falha: Crédito Disponível tem de fechar com Recebido − Concedido − Empenhado.
- Acumula um snapshot diário em data/history.json (para o gráfico de tendência).
- Escreve site/index.html (self-contained: CSS + SVG + tabela) e site/data/history.json.

Design (UI/UX): número-herói + equação Recebido−Empenhado=Disponível, waterfall,
barras divergentes, funil de estágios, tendência, tabela com busca/ordenação, tema claro/escuro.

Uso:
    python gerar_dashboard.py                 # baixa do Drive (FileId padrão / env DRIVE_FILE_ID)
    python gerar_dashboard.py --local X.xlsx  # usa arquivo local (teste)
    python gerar_dashboard.py --date 2026-07-14  # força a data do snapshot (default: hoje)
"""
import os, sys, json, argparse, datetime, urllib.request, tempfile, html, math, re
import openpyxl

HDR_ROW, DATA_ROW = 8, 9
ALVOS = [("160329", "BCMS · OGU"), ("167329", "BCMS · Fundo do Exército")]
FONTE_CURTA = {"160329": "OGU", "167329": "FEx"}
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
    ws = None
    for nm in wb.sheetnames:
        if norm(nm).startswith("CREDITO DISP") or "CRÉDITO DISP" in nm:
            ws = wb[nm]; break
    if ws is None:
        ws = wb.active
    # LINHA do cabeçalho de métricas detectada dinamicamente (o export do TG varia entre linha 7 e 8)
    hdr_row = None
    for r in range(1, 16):
        rowvals = {norm(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)}
        if "CREDITO DISPONIVEL" in rowvals or "PROVISAO RECEBIDA" in rowvals:
            hdr_row = r; break
    if not hdr_row:
        raise SystemExit("Não encontrei o cabeçalho (PROVISAO RECEBIDA / CREDITO DISPONIVEL) — "
                         "a fonte no Drive pode ter mudado de formato/relatório.")
    hdr = {}
    for c in range(1, ws.max_column + 1):
        n = norm(ws.cell(hdr_row, c).value)
        if n and n not in hdr: hdr[n] = c
    def col(name, req=True):
        c = hdr.get(name)
        if not c and req:
            raise SystemExit(f"Coluna '{name}' não encontrada (aba '{ws.title}') — a fonte mudou de layout.")
        return c
    C = dict(prov=col("PROVISAO RECEBIDA"), cred=col("CREDITO DISPONIVEL"),
             emp=col("DESPESAS EMPENHADAS"), liq=col("DESPESAS LIQUIDADAS"), pag=col("DESPESAS PAGAS"),
             conc=col("PROVISAO CONCEDIDA", req=False))  # ausente em alguns relatórios (CRO usa DESTAQUE)
    # primeira linha de dados: 1ª após o cabeçalho cujo col 3 (UG Executora) é um código numérico
    data_row = None
    for r in range(hdr_row + 1, min(hdr_row + 8, ws.max_row + 1)):
        if norm(ws.cell(r, 3).value).replace("'", "").isdigit():
            data_row = r; break
    if not data_row:
        data_row = hdr_row + 1
    # período real do relatório (ex.: "JUL/2026")
    periodo = None
    for r in range(max(1, hdr_row - 2), hdr_row + 3):
        for c in range(1, ws.max_column + 1):
            v = str(ws.cell(r, c).value or "").strip().upper()
            if re.match(r"^[A-Z]{3}/\d{4}$", v):
                periodo = v; break
        if periodo:
            break
    res = {}
    for cod, nome in ALVOS:
        res[cod] = dict(nome=nome, prov=0.0, conc=0.0, cred=0.0, emp=0.0, liq=0.0, pag=0.0,
                        n=0, por_acao={}, por_nd={}, nd_nome={}, linhas=[], celulas={})
    for r in range(data_row, ws.max_row + 1):
        cod = norm(ws.cell(r, 3).value)
        if cod not in res: continue
        d = res[cod]
        vals = {k: (to_num(ws.cell(r, C[k]).value) if C[k] else 0.0) for k in C}
        for k in C: d[k] += vals[k]
        d["n"] += 1
        ncv = disp(ws.cell(r, 5).value)          # "" quando NC = não se aplica (linha de empenho)
        acao = disp(ws.cell(r, 6).value) or "(s/ ação)"
        pic = disp(ws.cell(r, 7).value) or "—"
        pin = disp(ws.cell(r, 8).value)
        ndc = disp(ws.cell(r, 9).value) or "(s/ ND)"
        ndn = disp(ws.cell(r, 10).value)
        d["por_acao"][acao] = d["por_acao"].get(acao, 0.0) + vals["cred"]
        d["por_nd"][ndc] = d["por_nd"].get(ndc, 0.0) + vals["cred"]
        d["nd_nome"][ndc] = ndn
        # AGREGAÇÃO POR CÉLULA (Ação+PI+ND): o "crédito em tela" é o SALDO LÍQUIDO da célula.
        # Decompõe-se o próprio Crédito Disponível (col.17) em seus componentes, para que a linha
        # feche exata (Recebido líq − Empenhado = Disponível): linhas com NC real (descentralização,
        # alteração de ND, anulação) somam em "aloc"; linhas de empenho (NC = não se aplica) somam
        # em "emp". Assim a alteração de ND NÃO é somada em dobro com a NC original.
        ck = (acao, pic, ndc)
        cl = d["celulas"].get(ck)
        if cl is None:
            cl = dict(acao=acao, pi=pic, pi_nome=pin, nd=ndc, nd_nome=ndn,
                      aloc=0.0, emp=0.0, liq=0.0, pag=0.0, cred=0.0)
            d["celulas"][ck] = cl
        cl["cred"] += vals["cred"]
        cl["liq"] += vals["liq"]                  # liquidado da célula (estágio da despesa)
        cl["pag"] += vals["pag"]                  # pago da célula
        if ncv == "":
            cl["emp"] += -vals["cred"]           # empenho (col.17 negativo → empenho positivo)
        else:
            cl["aloc"] += vals["cred"]           # crédito recebido/alterado/anulado (líquido)
        if ndn and not cl["nd_nome"]:
            cl["nd_nome"] = ndn
        if round(vals["cred"], 2) != 0:
            d["linhas"].append(dict(
                emit=disp(ws.cell(r, 2).value), nc=ncv, acao=acao,
                pi=pic, nd=ndc, nd_desc=ndn,
                obj=disp(ws.cell(r, 11).value), op=disp(ws.cell(r, 12).value),
                dia=disp(ws.cell(r, 13).value),
                cred=vals["cred"], emp=vals["emp"], liq=vals["liq"], pag=vals["pag"]))
    for cod in res:
        res[cod]["linhas"].sort(key=lambda x: x["cred"], reverse=True)

    # GUARD: se NENHUMA UASG do BCMS aparece, a fonte no Drive trocou de relatório.
    # Falha claro (o site publicado anterior permanece no ar; nada de painel vazio/errado).
    if sum(res[c]["n"] for c, _ in ALVOS) == 0:
        ugs = sorted({norm(ws.cell(r, 3).value) for r in range(data_row, ws.max_row + 1)
                      if norm(ws.cell(r, 3).value).replace("'", "").isdigit()})
        raise SystemExit(
            f"Nenhuma linha das UASGs {'/'.join(c for c, _ in ALVOS)} (BCMS) na aba '{ws.title}'. "
            f"UGs presentes: {', '.join(ugs) or '(nenhuma)'}. A fonte no Drive parece ter mudado de "
            f"relatório — aponte o link correto do 'CRÉDITO DISP' do BCMS (UASG 160329/167329).")

    # VALIDAÇÃO ANTI-FALHA: o Crédito Disponível do TG tem de fechar com
    # Provisão Recebida − Provisão Concedida − Empenhado (identidade contábil do SIAFI).
    alertas = []
    for cod, _ in ALVOS:
        d = res[cod]
        esperado = d["prov"] - d["conc"] - d["emp"]
        dif = d["cred"] - esperado
        if abs(dif) > 0.02:
            alertas.append(
                f"UASG {cod}: Crédito Disponível somado (R$ {d['cred']:,.2f}) diverge de "
                f"Recebido−Concedido−Empenhado (R$ {esperado:,.2f}); diferença R$ {dif:,.2f}. "
                f"Verifique se o layout da planilha mudou.")
        if d["n"] == 0:
            alertas.append(f"UASG {cod}: nenhuma linha encontrada — confira o filtro/planilha.")
    return res, periodo, alertas

# ---------------- histórico ----------------
def atualizar_historico(res, data_str):
    hist = []
    if os.path.exists(HISTFILE):
        try:
            with open(HISTFILE, encoding="utf-8") as f: hist = json.load(f)
        except Exception: hist = []
    tot = {k: sum(res[c][k] for c, _ in ALVOS) for k in ("prov", "conc", "cred", "emp", "liq", "pag")}
    snap = {"data": data_str,
            "total": {k: round(tot[k], 2) for k in tot},
            **{c: {k: round(res[c][k], 2) for k in ("prov", "conc", "cred", "emp", "liq", "pag")} for c, _ in ALVOS}}
    hist = [h for h in hist if h.get("data") != data_str]
    hist.append(snap)
    hist.sort(key=lambda h: h["data"])
    os.makedirs(DATA, exist_ok=True)
    with open(HISTFILE, "w", encoding="utf-8") as f:
        json.dump(hist, f, ensure_ascii=False, indent=1)
    return hist

# ---------------- formatação ----------------
def _fmt(v):
    return f"{abs(v):,.2f}".replace(",", "\x00").replace(".", ",").replace("\x00", ".")

def brl(v):  # com prefixo R$
    return ("−R$ " if v < 0 else "R$ ") + _fmt(v)

def num(v):  # sem prefixo
    return ("−" if v < 0 else "") + _fmt(v)

def pct(a, b): return (100.0 * a / b) if b else 0.0

def abrev(v):
    a = abs(v); s = "−" if v < 0 else ""
    if a >= 1e6: return s + f"{a/1e6:.1f}".replace(".", ",") + " mi"
    if a >= 1e3: return s + f"{a/1e3:.0f} mil"
    return s + f"{a:.0f}"

def esc(s): return html.escape(str(s))

# ---------------- SVG ----------------
def _r(x, y, w, h, var, extra=""):
    return f'<rect x="{x:.1f}" y="{y:.1f}" width="{max(0,w):.1f}" height="{h:.1f}" style="fill:var(--{var})" {extra}/>'

def svg_util(recebido, empenhado, disponivel):
    """Barra empilhada de utilização (herói): Recebido = Empenhado + Disponível."""
    if recebido <= 0:
        return '<p class="vazio">sem provisão recebida</p>'
    x0, x1, y, h, W, H = 4, 636, 40, 34, 640, 92
    plot = x1 - x0
    fe = max(0.0, min(1.0, empenhado / recebido))
    we = plot * fe
    wd = plot - we
    pe, pd = pct(empenhado, recebido), pct(disponivel, recebido)
    return f'''<svg viewBox="0 0 {W} {H}" class="svg" role="img" aria-label="Utilização do crédito recebido">
<title>Provisão Recebida {brl(recebido)}: Empenhado {brl(empenhado)} ({pe:.1f}%) + Crédito Disponível {brl(disponivel)} ({pd:.1f}%)</title>
<text x="{x0}" y="20" class="s-lbl">PROVISÃO RECEBIDA · {esc(brl(recebido))}</text>
<line x1="{x0}" y1="26" x2="{x1}" y2="26" class="s-brk"/><line x1="{x0}" y1="26" x2="{x0}" y2="32" class="s-brk"/><line x1="{x1}" y1="26" x2="{x1}" y2="32" class="s-brk"/>
{_r(x0, y, we, h, "warning", 'rx="2"')}
{_r(x0+we, y, wd, h, "success", 'rx="2"')}
<line x1="{x0+we:.1f}" y1="{y}" x2="{x0+we:.1f}" y2="{y+h}" style="stroke:var(--surface);stroke-width:2"/>
<text x="{x0+6}" y="{y+h+16}" class="s-seg">Empenhado {esc(brl(empenhado))} · {pe:.1f}%</text>
<text x="{x1-6}" y="{y+h+16}" text-anchor="end" class="s-seg s-seg-ok">Crédito Disponível {esc(brl(disponivel))} · {pd:.1f}%</text>
</svg>'''

def svg_waterfall(recebido, empenhado, disponivel, mini=False):
    """Waterfall horizontal: Recebida → (−)Empenhado → (=)Disponível."""
    if recebido <= 0:
        return '<p class="vazio">sem dados</p>'
    if mini:
        W, H, xL, rh, gap, fs = 360, 118, 112, 26, 10, 10
        L1, L2, L3 = "Recebido", "(−) Empenhado", "(=) Disponível"
    else:
        W, H, xL, rh, gap, fs = 720, 196, 190, 42, 16, 13
        L1, L2, L3 = "Provisão Recebida", "(−) Empenhado", "(=) Crédito Disponível"
    xR = W - 24
    plot = xR - xL
    R = recebido
    def X(v): return xL + (v / R) * plot
    y1 = 14; y2 = y1 + rh + gap; y3 = y2 + rh + gap
    xd = X(max(disponivel, 0))
    parts = [f'<svg viewBox="0 0 {W} {H}" class="svg" role="img" aria-label="Composição do crédito">',
             f'<title>Provisão Recebida {brl(recebido)} menos Empenhado {brl(empenhado)} igual a Crédito Disponível {brl(disponivel)}</title>',
             f'<desc>{esc(brl(recebido))} − {esc(brl(empenhado))} = {esc(brl(disponivel))}</desc>']
    # linha 1 — Recebida
    parts.append(_r(xL, y1, plot, rh, "primary", 'rx="3"'))
    parts.append(f'<text x="{xL-8}" y="{y1+rh/2+4:.0f}" text-anchor="end" class="s-cat">{esc(L1)}</text>')
    parts.append(f'<text x="{xR-6}" y="{y1+rh/2+4:.0f}" text-anchor="end" class="s-val s-on">{esc(brl(recebido))}</text>')
    # linha 2 — Empenhado (flutuante, à direita, de X(D) a X(R))
    we = plot - (xd - xL)
    parts.append(_r(xd, y2, we, rh, "warning", 'rx="3"'))
    parts.append(f'<text x="{xL-8}" y="{y2+rh/2+4:.0f}" text-anchor="end" class="s-cat">{esc(L2)}</text>')
    parts.append(f'<text x="{xd+8:.1f}" y="{y2+rh/2+4:.0f}" class="s-val s-on">−{esc(brl(empenhado).replace("R$ ","R$ "))}</text>')
    # linha 3 — Disponível
    parts.append(_r(xL, y3, xd - xL, rh, "success", 'rx="3"'))
    parts.append(f'<text x="{xL-8}" y="{y3+rh/2+4:.0f}" text-anchor="end" class="s-cat s-cat-ok">{esc(L3)}</text>')
    parts.append(f'<text x="{xd+8:.1f}" y="{y3+rh/2+4:.0f}" class="s-val s-ok">{esc(brl(disponivel))}</text>')
    # conectores tracejados
    parts.append(f'<line x1="{xR:.1f}" y1="{y1+rh}" x2="{xR:.1f}" y2="{y2}" class="s-conn"/>')
    parts.append(f'<line x1="{xd:.1f}" y1="{y2+rh}" x2="{xd:.1f}" y2="{y3}" class="s-conn"/>')
    parts.append('</svg>')
    return "".join(parts)

def svg_diverg(itens, titulo, max_itens=8):
    itens = [(k, v) for k, v in itens if round(v, 2) != 0]
    itens = sorted(itens, key=lambda x: abs(x[1]), reverse=True)
    resto = sum(v for _, v in itens[max_itens:])
    itens = itens[:max_itens]
    if round(resto, 2) != 0:
        itens.append(("Outras", resto))
    if not itens:
        return f'<div class="card chart"><div class="eyebrow">{esc(titulo)}</div><p class="vazio">sem valores</p></div>'
    vmax = max(abs(v) for _, v in itens) or 1
    labW, rh, pad = 150, 30, 8
    zx = labW + 246  # eixo zero
    half = 230
    W = zx + half + 66
    H = pad * 2 + rh * len(itens) + 16
    el = [f'<line x1="{zx}" y1="{pad}" x2="{zx}" y2="{pad+rh*len(itens):.0f}" class="s-zero"/>']
    for i, (k, v) in enumerate(itens):
        y = pad + i * rh
        w = abs(v) / vmax * half
        lbl = k if len(k) <= 24 else k[:23] + "…"
        el.append(f'<text x="{labW-6}" y="{y+rh/2+4:.0f}" text-anchor="end" class="s-cat">{esc(lbl)}</text>')
        if v >= 0:
            el.append(_r(zx, y+5, w, rh-12, "success", f'rx="2"><title>{esc(k)}: {esc(brl(v))}</title></rect'.replace("/>", ">")))
            el.append(f'<text x="{zx+w+6:.1f}" y="{y+rh/2+4:.0f}" class="s-num s-ok">{esc(num(v))}</text>')
        else:
            el.append(_r(zx-w, y+5, w, rh-12, "danger", f'rx="2"><title>{esc(k)}: {esc(brl(v))}</title></rect'.replace("/>", ">")))
            el.append(f'<text x="{zx-w-6:.1f}" y="{y+rh/2+4:.0f}" text-anchor="end" class="s-num s-neg">{esc(num(v))}</text>')
    svg = f'<svg viewBox="0 0 {W} {H}" class="svg" role="img" aria-label="{esc(titulo)}">{"".join(el)}</svg>'
    return f'<div class="card chart"><div class="eyebrow">{esc(titulo)}</div>{svg}</div>'

def svg_funil(cod, d):
    emp, liq, pag = d["emp"], d["liq"], d["pag"]
    base = emp or 1
    W, rh, gap, xL = 340, 24, 12, 108
    plot = W - xL - 70
    H = 18 + 3 * (rh + gap)
    rows = [("Empenhado", emp, "stg1", ""), ("Liquidado", liq, "stg2", f"{pct(liq,emp):.0f}% do emp."),
            ("Pago", pag, "stg3", f"{pct(pag,liq):.0f}% do liq.")]
    el = []
    for i, (nome, val, cls, conv) in enumerate(rows):
        y = 12 + i * (rh + gap)
        w = max(2, abs(val) / base * plot)
        el.append(f'<text x="{xL-8}" y="{y+rh/2+4:.0f}" text-anchor="end" class="s-cat">{nome}</text>')
        el.append(f'<rect x="{xL}" y="{y}" width="{w:.1f}" height="{rh}" rx="2" style="fill:var(--{cls})"><title>{nome}: {esc(brl(val))}</title></rect>')
        el.append(f'<text x="{xL+w+6:.1f}" y="{y+rh/2+4:.0f}" class="s-num s-on2">{esc(num(val))}</text>')
        if conv:
            el.append(f'<text x="{W-4}" y="{y+rh/2+4:.0f}" text-anchor="end" class="s-conv">{conv}</text>')
    svg = f'<svg viewBox="0 0 {W} {H}" class="svg" role="img" aria-label="Estágios {cod}">{"".join(el)}</svg>'
    return f'<div class="card chart"><div class="eyebrow">Estágios · {esc(cod)} {esc(FONTE_CURTA[cod])}</div>{svg}</div>'

def svg_tendencia(hist):
    # agrega por SEMANA ISO — mantém o último snapshot de cada semana
    wk, order = {}, []
    for h in hist:
        try:
            y, m, dd = (int(x) for x in h["data"].split("-"))
            key = datetime.date(y, m, dd).isocalendar()[:2]
        except Exception:
            key = h.get("data")
        if key not in wk:
            order.append(key)
        wk[key] = h
    semanas = [wk[k] for k in order]
    pts = [(h["data"], h["total"]["cred"]) for h in semanas]
    W, H, pl, pb, pt, pr = 720, 210, 66, 34, 16, 74
    pw, ph = W - pl - pr, H - pb - pt
    vals = [v for _, v in pts]
    vmin, vmax = min(vals + [0]), max(vals + [1])
    rng = (vmax - vmin) or 1
    n = len(pts)
    def X(i): return pl + (pw * (i / (n - 1)) if n > 1 else pw / 2)
    def Y(v): return pt + ph - ((v - vmin) / rng * ph)
    el = []
    for t in range(4):
        val = vmin + rng * t / 3; y = Y(val)
        el.append(f'<line x1="{pl}" y1="{y:.1f}" x2="{W-pr}" y2="{y:.1f}" class="s-grid"/>')
        el.append(f'<text x="{pl-8}" y="{y+3:.1f}" text-anchor="end" class="s-ax">{esc(abrev(val))}</text>')
    if n > 1:
        line = "M" + " L".join(f"{X(i):.1f},{Y(v):.1f}" for i, (_, v) in enumerate(pts))
        area = f"M{X(0):.1f},{Y(vmin):.1f} L" + " L".join(f"{X(i):.1f},{Y(v):.1f}" for i, (_, v) in enumerate(pts)) + f" L{X(n-1):.1f},{Y(vmin):.1f} Z"
        el.append(f'<path d="{area}" class="s-area"/>')
        el.append(f'<path d="{line}" class="s-line"/>')
    step = max(1, n // 6)
    for i, (dt, v) in enumerate(pts):
        show = (i % step == 0 or i == n - 1)
        el.append(f'<circle cx="{X(i):.1f}" cy="{Y(v):.1f}" r="3" class="s-dot"><title>{esc(dt)}: {esc(brl(v))}</title></circle>')
        if show:
            el.append(f'<text x="{X(i):.1f}" y="{H-pb+16}" text-anchor="middle" class="s-ax">{dt[8:10]}/{dt[5:7]}</text>')
    # callout no último
    if n >= 1:
        dt, v = pts[-1]
        el.append(f'<text x="{X(n-1):.1f}" y="{Y(v)-10:.1f}" text-anchor="end" class="s-num s-ok">{esc(brl(v))}</text>')
    nota = "" if n > 1 else '<p class="vazio">a curva semanal aparece a partir da 2ª semana de histórico.</p>'
    svg = f'<svg viewBox="0 0 {W} {H}" class="svg" role="img" aria-label="Tendência semanal do Crédito Disponível">{"".join(el)}</svg>'
    return f'<div class="card chart wide"><div class="eyebrow">Tendência semanal · Crédito Disponível consolidado</div>{svg}{nota}</div>'

# ---------------- brasão BCMS ----------------
def _cog(cx, cy, ro, ri, teeth):
    pts = []
    for i in range(teeth * 2):
        ang = math.pi * i / teeth - math.pi / 2
        r = ro if i % 2 == 0 else ri
        pts.append(f"{cx + r*math.cos(ang):.1f},{cy + r*math.sin(ang):.1f}")
    return "M" + " L".join(pts) + " Z"

def brasao_svg():
    GOLD, RED, BLUE, CREAM, GRAY = "#C8901E", "#CE2B2B", "#1E6FD0", "#FBF3DA", "#CBD0D6"
    g1, g2 = _cog(46, 62, 13, 9.5, 10), _cog(72, 62, 13, 9.5, 10)
    star = _cog(60, 120, 13, 5.5, 9)
    return (f'<svg class="brasao" viewBox="0 0 120 152" role="img" aria-label="Brasão do BCMS">'
            f'<path d="M12,20 L108,20 L108,86 C108,118 86,136 60,148 C34,136 12,118 12,86 Z" fill="#fff" stroke="{GOLD}" stroke-width="3"/>'
            f'<path d="M27,58 L93,58 L104,88 L82,124 L60,136 L38,124 L16,88 Z" fill="{GRAY}" opacity="0.5"/>'
            f'<g stroke="{GOLD}" stroke-width="1.5" stroke-linecap="round">'
            f'<rect x="30" y="88" width="60" height="8" rx="4" fill="{CREAM}" transform="rotate(-32 60 92)"/>'
            f'<rect x="30" y="88" width="60" height="8" rx="4" fill="{CREAM}" transform="rotate(32 60 92)"/></g>'
            f'<path d="{g1}" fill="{CREAM}" stroke="{GOLD}" stroke-width="1.6"/><circle cx="46" cy="62" r="4" fill="none" stroke="{GOLD}" stroke-width="1.3"/>'
            f'<path d="{g2}" fill="{CREAM}" stroke="{GOLD}" stroke-width="1.6"/><circle cx="72" cy="62" r="4" fill="none" stroke="{GOLD}" stroke-width="1.3"/>'
            f'<path d="{star}" fill="{CREAM}" stroke="{GOLD}" stroke-width="1.5"/>'
            f'<rect x="3" y="2" width="114" height="34" rx="3" fill="{GOLD}"/>'
            f'<rect x="6" y="5" width="108" height="14" fill="{RED}"/><rect x="6" y="19" width="108" height="14" fill="{BLUE}"/>'
            f'<text x="60" y="26" text-anchor="middle" font-family="Georgia,serif" font-weight="700" font-size="19" fill="#fff">BCMS</text>'
            f'</svg>')

def brasao_html():
    import base64, glob
    mimes = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".svg": "image/svg+xml"}
    adir = os.path.join(HERE, "assets")
    # 1) nomes preferenciais; 2) qualquer imagem na pasta assets/
    candidatos = [os.path.join(adir, "brasao" + e) for e in mimes]
    if os.path.isdir(adir):
        for f in sorted(glob.glob(os.path.join(adir, "*"))):
            if os.path.splitext(f)[1].lower() in mimes and f not in candidatos:
                candidatos.append(f)
    for p in candidatos:
        ext = os.path.splitext(p)[1].lower()
        if ext in mimes and os.path.exists(p):
            with open(p, "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
            return f'<img class="brasao" src="data:{mimes[ext]};base64,{b64}" alt="Brasão do BCMS — Batalhão Central de Manutenção e Suprimento">'
    return brasao_svg()

# ---------------- componentes HTML ----------------
def kpi_tile(label, valor, chip, cls):
    chip_html = f'<span class="chip">{esc(chip)}</span>' if chip else ""
    return (f'<div class="kpi kpi-{cls}"><div class="kpi-l">{esc(label)}</div>'
            f'<div class="kpi-v num">{esc(valor)}</div>{chip_html}</div>')

def uasg_card(cod, d):
    barp = pct(d["emp"], d["prov"])
    return (f'<div class="card uasg">'
            f'<div class="uasg-h"><span class="uasg-cod">{esc(cod)}</span>'
            f'<span class="uasg-fonte">{esc(FONTE_CURTA[cod])}</span>'
            f'<span class="uasg-nome">{esc(d["nome"].split("·")[1].strip())}</span></div>'
            f'<div class="uasg-disp"><span class="uasg-disp-l">Crédito Disponível</span>'
            f'<span class="uasg-disp-v num">{esc(brl(d["cred"]))}</span></div>'
            f'<div class="uasg-eq num">{esc(brl(d["prov"]))} <i>−</i> {esc(brl(d["emp"]))}</div>'
            f'{svg_waterfall(d["prov"], d["emp"], d["cred"], mini=True)}'
            f'<div class="uasg-exec"><div class="exec-l"><span>Empenhado / Recebido</span><span class="num">{barp:.1f}%</span></div>'
            f'<div class="exec-track"><div class="exec-fill" style="width:{min(barp,100):.1f}%"></div></div></div>'
            f'</div>')

def tabela_html(tid, celulas, com_fonte, ativo):
    """Relação de crédito EM TELA por célula orçamentária (saldo líquido positivo)."""
    cols = (["Fonte"] if com_fonte else []) + ["Ação", "PI", "ND", "Aplicação", "Recebido (líq)", "Empenhado", "Crédito Disp."]
    ths = []
    for c in cols:
        numc = c in ("Recebido (líq)", "Empenhado", "Crédito Disp.")
        cls = ' class="num"' if numc else ''
        ths.append(f'<th{cls} tabindex="0" role="button" aria-sort="none" onclick="bcmsSort(this)" onkeydown="if(event.key==\'Enter\'||event.key==\' \'){{event.preventDefault();bcmsSort(this)}}">{esc(c)}<span class="sort"></span></th>')
    body = []
    tot = sum(c["cred"] for c in celulas)
    for c in celulas:
        fonte = f'<td><span class="pill-fonte">{esc(FONTE_CURTA.get(c.get("uasg",""),""))}</span></td>' if com_fonte else ''
        aplic = c.get("nd_nome") or c.get("pi_nome") or ""
        cid = esc(c.get("cid", ""))
        body.append(
            f'<tr class="cel-row" tabindex="0" role="button" data-cel="{cid}" title="Ver as notas de crédito desta célula (descrição completa)" '
            f'onclick="bcmsCel(this)" onkeydown="if(event.key==\'Enter\'||event.key==\' \'){{event.preventDefault();bcmsCel(this)}}">'
            f'{fonte}<td>{esc(c["acao"])}</td><td class="mono2">{esc(c["pi"])}</td><td class="mono2">{esc(c["nd"])}</td>'
            f'<td class="obj" title="{esc(aplic)}">{esc(aplic[:60])}</td>'
            f'<td class="num" data-sort="{c["aloc"]:.2f}">{esc(brl(c["aloc"]))}</td>'
            f'<td class="num" data-sort="{c["emp"]:.2f}">{esc(brl(c["emp"]))}</td>'
            f'<td class="num anchor" data-sort="{c["cred"]:.2f}">{esc(brl(c["cred"]))}<i class="chev" aria-hidden="true">›</i></td></tr>')
    ncols = len(cols)
    tfoot = (f'<tfoot><tr><td colspan="{ncols-1}">TOTAL · {len(celulas)} célula(s) com crédito em tela</td>'
             f'<td class="num anchor">{esc(brl(tot))}</td></tr></tfoot>')
    disp_style = "" if ativo else ' style="display:none"'
    return (f'<div class="tabpanel" id="{tid}" role="tabpanel"{disp_style}>'
            f'<div class="tbl-tools"><label class="visually-hidden" for="q-{tid}">Buscar</label>'
            f'<input type="search" id="q-{tid}" class="tbl-search" placeholder="Buscar por ação, PI, ND ou aplicação…" oninput="bcmsSearch(this,\'{tid}\')">'
            f'<span class="tbl-count" id="cnt-{tid}" data-unit="célula(s)" aria-live="polite">{len(celulas)} células</span></div>'
            f'<div class="tbl-scroll"><table class="det"><thead><tr>{"".join(ths)}</tr></thead>'
            f'<tbody>{"".join(body)}</tbody>{tfoot}</table></div></div>')

# ---------------- página ----------------
def gerar_html(res, hist, data_str, periodo=None, alertas=None):
    tot = {k: sum(res[c][k] for c, _ in ALVOS) for k in ("prov", "conc", "cred", "emp", "liq", "pag", "n")}
    ger = datetime.datetime.now().strftime("%d/%m/%Y às %H:%M")
    posicao = periodo if periodo else (data_str[8:10] + "/" + data_str[5:7] + "/" + data_str[0:4])

    # delta vs. dia anterior
    delta_html = ""
    if len(hist) >= 2:
        dv = hist[-1]["total"]["cred"] - hist[-2]["total"]["cred"]
        if round(dv, 2) != 0:
            seta = "▲" if dv > 0 else "▼"
            cls = "up" if dv > 0 else "down"
            delta_html = f'<div class="delta {cls}"><span>{seta}</span> {esc(num(dv))} <small>vs. dia anterior</small></div>'
        else:
            delta_html = '<div class="delta flat">sem variação vs. dia anterior</div>'
    else:
        delta_html = '<div class="delta flat">1º dia de histórico</div>'

    # alerta banner
    banner = ""
    if alertas:
        itens = "".join(f"<li>{esc(a)}</li>" for a in alertas)
        banner = f'<div class="banner" role="alert"><b>⚠ Verificação:</b><ul>{itens}</ul></div>'

    # KPIs secundários
    kpis = (kpi_tile("Provisão Recebida", brl(tot["prov"]), "", "prov") +
            kpi_tile("Empenhado", brl(tot["emp"]), f'{pct(tot["emp"],tot["prov"]):.1f}% do recebido', "emp") +
            kpi_tile("Liquidado", brl(tot["liq"]), f'{pct(tot["liq"],tot["emp"]):.1f}% do empenhado', "liq") +
            kpi_tile("Pago", brl(tot["pag"]), f'{pct(tot["pag"],tot["liq"]):.1f}% do liquidado', "pag"))

    cards_uasg = "".join(uasg_card(cod, res[cod]) for cod, _ in ALVOS)

    # combinado por Ação / ND
    acao, nd, nd_nome = {}, {}, {}
    for cod, _ in ALVOS:
        for k, v in res[cod]["por_acao"].items(): acao[k] = acao.get(k, 0) + v
        for k, v in res[cod]["por_nd"].items(): nd[k] = nd.get(k, 0) + v
        nd_nome.update(res[cod]["nd_nome"])
    nd_lbl = {(f'{k} {nd_nome.get(k,"")[:18]}').strip(): v for k, v in nd.items()}
    ch_acao = svg_diverg(acao.items(), "Crédito Disponível por Ação (R$)")
    ch_nd = svg_diverg(nd_lbl.items(), "Crédito Disponível por Natureza de Despesa (R$)")
    funis = "".join(svg_funil(cod, res[cod]) for cod, _ in ALVOS)
    trend = svg_tendencia(hist)

    # tabelas — crédito EM TELA por célula (saldo líquido positivo)
    # + dados de drill-down: as NCs (com descrição completa) que compõem cada célula
    lin_idx = {}
    for cod, _ in ALVOS:
        for L in res[cod]["linhas"]:
            lin_idx.setdefault((cod, L["acao"], L["pi"], L["nd"]), []).append(L)
    celdata = {}
    cid_map = {}
    cid_seq = [0]
    def cid_for(uasg, c):
        key = (uasg, c["acao"], c["pi"], c["nd"])
        cid = cid_map.get(key)
        if cid is None:
            cid_seq[0] += 1
            cid = "c%d" % cid_seq[0]
            cid_map[key] = cid
            ncs = sorted(([L["nc"], L["op"], round(L["cred"], 2), L["obj"], L.get("emit", ""), L.get("dia", "")]
                          for L in lin_idx.get(key, [])), key=lambda x: -x[2])
            celdata[cid] = {"t": f'{c["acao"]} · PI {c["pi"]} · ND {c["nd"]}', "nome": c.get("nd_nome", ""),
                            "u": FONTE_CURTA.get(uasg, ""), "uasg": uasg,
                            "acao": c["acao"], "pi": c["pi"], "pinome": c.get("pi_nome", ""),
                            "nd": c["nd"], "ndnome": c.get("nd_nome", ""),
                            "r": round(c["aloc"], 2), "e": round(c["emp"], 2), "l": round(c.get("liq", 0.0), 2),
                            "p": round(c.get("pag", 0.0), 2), "d": round(c["cred"], 2), "ncs": ncs}
        c["cid"] = cid
        return cid
    def celulas_pos(cod):
        cl = [c for c in res[cod]["celulas"].values() if c["cred"] > 0.005]
        cl.sort(key=lambda x: x["cred"], reverse=True)
        for c in cl:
            cid_for(cod, c)
        return cl
    cons_cel = []
    for cod, _ in ALVOS:
        for c in celulas_pos(cod):
            cons_cel.append({**c, "uasg": cod, "cid": cid_for(cod, c)})
    cons_cel.sort(key=lambda x: x["cred"], reverse=True)
    abas = (f'<button class="tab on" role="tab" aria-selected="true" tabindex="0" onclick="bcmsTab(this,\'tab-cons\')" onkeydown="bcmsTabKey(event,this)">Consolidado</button>'
            f'<button class="tab" role="tab" aria-selected="false" tabindex="-1" onclick="bcmsTab(this,\'tab-160329\')" onkeydown="bcmsTabKey(event,this)">160329 · OGU</button>'
            f'<button class="tab" role="tab" aria-selected="false" tabindex="-1" onclick="bcmsTab(this,\'tab-167329\')" onkeydown="bcmsTabKey(event,this)">167329 · FEx</button>')
    tabs = (tabela_html("tab-cons", cons_cel, True, True) +
            tabela_html("tab-160329", celulas_pos("160329"), False, False) +
            tabela_html("tab-167329", celulas_pos("167329"), False, False))

    # ===== ABA RESUMO: movimentação de NC do dia + resumo semanal =====
    movs = []
    for cod, _ in ALVOS:
        for L in res[cod]["linhas"]:
            if not L["nc"]:
                continue  # só NC real (exclui linhas de empenho)
            try:
                dd, mm, yy = L.get("dia", "").split("/")
                dt = datetime.date(int(yy), int(mm), int(dd))
            except Exception:
                dt = None
            movs.append({**L, "uasg": cod, "dt": dt})
    ncdata = {}
    for i, m in enumerate(movs, 1):
        nid = "n%d" % i
        m["nid"] = nid
        ncdata[nid] = {"nc": m["nc"], "u": FONTE_CURTA.get(m["uasg"], ""), "acao": m["acao"],
                       "pi": m["pi"], "nd": m["nd"], "ndn": m.get("nd_desc", ""), "op": m["op"],
                       "dia": m.get("dia", ""), "val": round(m["cred"], 2), "obj": m["obj"]}
    datas = sorted({m["dt"] for m in movs if m["dt"]}, reverse=True)
    max_date = datas[0] if datas else None
    fmt_d = lambda d: d.strftime("%d/%m/%Y") if d else "—"
    daily = sorted([m for m in movs if m["dt"] == max_date] if max_date else [], key=lambda x: x["cred"], reverse=True)
    rec_d = sum(m["cred"] for m in daily if m["cred"] > 0)
    red_d = sum(m["cred"] for m in daily if m["cred"] < 0)

    def _th(h, numc, sortable=True):
        cls = ' class="num"' if numc else ''
        if not sortable:
            return f'<th{cls}>{esc(h)}</th>'
        return (f'<th{cls} tabindex="0" role="button" aria-sort="none" onclick="bcmsSort(this)" '
                f'onkeydown="if(event.key===\'Enter\'||event.key===\' \'){{event.preventDefault();bcmsSort(this)}}">{esc(h)}<span class="sort"></span></th>')

    def mov_row(m):
        neg = ' cell-neg' if m["cred"] < 0 else ''
        return (f'<tr class="cel-row" tabindex="0" role="button" data-nc="{esc(m["nid"])}" title="Detalhar a NC" '
                f'onclick="bcmsNC(this)" onkeydown="if(event.key===\'Enter\'||event.key===\' \'){{event.preventDefault();bcmsNC(this)}}">'
                f'<td class="mono2">{esc(m["nc"])}</td><td><span class="pill-fonte">{esc(FONTE_CURTA[m["uasg"]])}</span></td>'
                f'<td>{esc(m["acao"])}</td><td class="mono2">{esc(m["nd"])}</td>'
                f'<td class="obj" title="{esc(m["op"])}">{esc(m["op"][:28])}</td>'
                f'<td class="num anchor{neg}" data-sort="{m["cred"]:.2f}">{esc(brl(m["cred"]))}<i class="chev" aria-hidden="true">›</i></td></tr>')

    def mov_tabela(tid, lst):
        if not lst:
            return '<p class="vazio">Sem movimentação de NC neste período.</p>'
        ths = _th("NC", False) + _th("Fonte", False) + _th("Ação", False) + _th("ND", False) + _th("Operação", False) + _th("Valor", True)
        body = "".join(mov_row(m) for m in lst)
        return (f'<div class="tbl-tools"><label class="visually-hidden" for="q-{tid}">Buscar</label>'
                f'<input type="search" id="q-{tid}" class="tbl-search" placeholder="Buscar por NC, ação, ND ou operação…" oninput="bcmsSearch(this,\'{tid}\')">'
                f'<span class="tbl-count" id="cnt-{tid}" data-unit="NC(s)" aria-live="polite">{len(lst)} NC(s)</span></div>'
                f'<div class="tbl-scroll" id="{tid}"><table class="det"><thead><tr>{ths}</tr></thead><tbody>{body}</tbody></table></div>')

    # resumo semanal — últimos 7 dias com movimentação
    week_days = datas[:7]
    week_rows, tot_rec, tot_red, n_week = [], 0.0, 0.0, 0
    daydata = {}
    for idx, d in enumerate(week_days):
        dm = [m for m in movs if m["dt"] == d]
        rec = sum(m["cred"] for m in dm if m["cred"] > 0)
        red = sum(m["cred"] for m in dm if m["cred"] < 0)
        tot_rec += rec; tot_red += red; n_week += len(dm)
        dk = "d%d" % idx
        daydata[dk] = {"d": fmt_d(d), "n": len(dm), "rec": round(rec, 2), "red": round(red, 2), "liq": round(rec + red, 2),
                       "ncs": sorted([[m["nc"], FONTE_CURTA[m["uasg"]], m["op"], round(m["cred"], 2), m["obj"]] for m in dm],
                                     key=lambda x: -x[3])}
        week_rows.append((d, len(dm), rec, red, rec + red, dk))

    def semana_tabela(rows):
        if not rows:
            return '<p class="vazio">Sem movimentação na última semana.</p>'
        heads = _th("Dia", False, False) + _th("Nº NC", True, False) + _th("Recebido (+)", True, False) + _th("Reduções (−)", True, False) + _th("Líquido", True, False)
        body = ""
        for d, n, rec, red, liq, dk in rows:
            body += (f'<tr class="cel-row" tabindex="0" role="button" data-day="{dk}" title="Ver as NCs deste dia" '
                     f'onclick="bcmsDay(this)" onkeydown="if(event.key===\'Enter\'||event.key===\' \'){{event.preventDefault();bcmsDay(this)}}">'
                     f'<td class="mono2">{fmt_d(d)}</td><td class="num">{n}</td>'
                     f'<td class="num col-pos">{esc(brl(rec))}</td><td class="num col-neg">{esc(brl(red))}</td>'
                     f'<td class="num anchor">{esc(brl(liq))}<i class="chev" aria-hidden="true">›</i></td></tr>')
        foot = (f'<tfoot><tr><td>TOTAL</td><td class="num">{n_week}</td>'
                f'<td class="num">{esc(brl(tot_rec))}</td><td class="num">{esc(brl(tot_red))}</td>'
                f'<td class="num anchor">{esc(brl(tot_rec + tot_red))}</td></tr></tfoot>')
        return f'<div class="tbl-scroll"><table class="det"><thead><tr>{heads}</tr></thead><tbody>{body}</tbody>{foot}</table></div>'

    # MÓDULO "Créditos em tela — resumo": o que são, valores e há quantos dias estão em tela (envelhecimento)
    asof = max_date or datetime.date.today()
    rec_por_cel = {}
    for cod, _ in ALVOS:
        for L in res[cod]["linhas"]:
            if L["cred"] > 0 and L.get("dia"):
                try:
                    dd, mm, yy = L["dia"].split("/"); dt = datetime.date(int(yy), int(mm), int(dd))
                except Exception:
                    dt = None
                if dt:
                    rec_por_cel.setdefault((cod, L["acao"], L["pi"], L["nd"]), []).append((dt, L["cred"]))
    def idade_celula(cod, c):
        recs = sorted(rec_por_cel.get((cod, c["acao"], c["pi"], c["nd"]), []), reverse=True)
        if not recs:
            return None, None
        acc, ref = 0.0, recs[0][0]
        for dt, v in recs:
            acc += v; ref = dt
            if acc >= c["cred"] - 0.005:
                break
        return (asof - ref).days, ref
    emtela = []
    for cod, _ in ALVOS:
        for c in celulas_pos(cod):
            dias, ref = idade_celula(cod, c)
            emtela.append({**c, "uasg": cod, "dias": dias, "ref": ref})
    emtela.sort(key=lambda x: x["cred"], reverse=True)
    tot_emtela = sum(c["cred"] for c in emtela)
    idades = [c["dias"] for c in emtela if c["dias"] is not None]
    idade_media = round(sum(idades) / len(idades)) if idades else 0
    idade_max = max(idades) if idades else 0

    def et_row(c):
        aplic = c.get("nd_nome") or c.get("pi_nome") or ""
        oque = f'{c["acao"]} · ND {c["nd"]}' + (f' — {aplic}' if aplic else '')
        dias = c["dias"]
        if dias is None:
            dcls, dtxt, dsort = "", "—", -1
        else:
            dcls = "col-neg" if dias > 60 else ("col-warn" if dias > 30 else "")
            dtxt = f'{dias} dia' + ('s' if dias != 1 else '')
            dsort = dias
        refd = c["ref"].strftime("%d/%m/%Y") if c["ref"] else "—"
        cid = esc(c.get("cid", ""))
        return (f'<tr class="cel-row" tabindex="0" role="button" data-cel="{cid}" title="Detalhar a célula" '
                f'onclick="bcmsCel(this)" onkeydown="if(event.key===\'Enter\'||event.key===\' \'){{event.preventDefault();bcmsCel(this)}}">'
                f'<td><span class="pill-fonte">{esc(FONTE_CURTA[c["uasg"]])}</span></td>'
                f'<td class="obj" title="{esc(oque)}">{esc(oque[:64])}</td>'
                f'<td class="mono2">{esc(refd)}</td>'
                f'<td class="num anchor" data-sort="{c["cred"]:.2f}">{esc(brl(c["cred"]))}</td>'
                f'<td class="num {dcls}" data-sort="{dsort}">{dtxt}<i class="chev" aria-hidden="true">›</i></td></tr>')
    et_ths = _th("Fonte", False) + _th("O que é (Ação · ND)", False) + _th("Recebido em", False) + _th("Crédito Disponível", True) + _th("Dias em tela", True)
    emtela_html = (
        '<section class="sec"><div class="eyebrow">Créditos em tela — resumo</div>'
        '<div class="et-head">'
        f'<div class="et-kpi et-hero"><span>Crédito Disponível em tela</span><b class="num">{esc(brl(tot_emtela))}</b></div>'
        f'<div class="et-kpi"><span>Células</span><b>{len(emtela)}</b></div>'
        f'<div class="et-kpi"><span>Idade média</span><b>{idade_media} dias</b></div>'
        f'<div class="et-kpi"><span>Mais antigo</span><b>{idade_max} dias</b></div>'
        f'<div class="et-meta">Posição {esc(posicao)}<br><span class="rh-delay">⏱ dados com ~24h de defasagem</span></div></div>'
        '<p class="sec-nota">O que está disponível para empenhar, por célula (Ação · PI · ND): o que é, o valor e '
        '<b>há quantos dias está em tela</b> (desde o recebimento que compõe o saldo). '
        '<b>Clique numa linha</b> para detalhar. Cor dos dias: verde ≤30 · âmbar 31–60 · vermelho &gt;60.</p>'
        f'<div class="tbl-scroll"><table class="det"><thead><tr>{et_ths}</tr></thead>'
        f'<tbody>{"".join(et_row(c) for c in emtela)}</tbody></table></div></section>'
    )

    resumo_html = (
        emtela_html
        + f'<section class="sec"><div class="eyebrow">Movimentação de NC — {fmt_d(max_date)} (dia anterior)</div>'
        f'<p class="sec-nota">Notas de crédito com lançamento em <b>{fmt_d(max_date)}</b> (último dia com movimento — dados com ~24h de defasagem): '
        f'<b>{len(daily)}</b> NC(s) · Recebido <b>{esc(brl(rec_d))}</b> · Reduções <b>{esc(brl(red_d))}</b> · Líquido <b>{esc(brl(rec_d + red_d))}</b>. '
        'Clique em uma NC para detalhá-la.</p>'
        + mov_tabela("mov-dia", daily) + '</section>'
        '<section class="sec"><div class="eyebrow">Resumo semanal — últimos 7 dias com movimentação</div>'
        '<p class="sec-nota">Movimentação de NC por dia: recebimentos (+), reduções/anulações (−) e líquido. <b>Clique em um dia</b> para ver as NCs daquele dia (e clique numa NC para a descrição completa).</p>'
        + semana_tabela(week_rows) + '</section>'
    )

    css = CSS
    js = JS
    celdata_json = json.dumps(celdata, ensure_ascii=False).replace("</", "<\\/")
    ncdata_json = json.dumps(ncdata, ensure_ascii=False).replace("</", "<\\/")
    daydata_json = json.dumps(daydata, ensure_ascii=False).replace("</", "<\\/")
    brasao = brasao_html()
    hero_eq = (f'<span class="eq-t eq-prov">{esc(brl(tot["prov"]))}</span>'
               f'<i class="eq-op">−</i>'
               f'<span class="eq-t eq-emp">{esc(brl(tot["emp"]))}</span>'
               f'<i class="eq-op">=</i>'
               f'<span class="eq-t eq-disp">{esc(brl(tot["cred"]))}</span>')
    return f"""<!doctype html><html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="color-scheme" content="light dark">
<title>Crédito Disponível — BCMS</title><style>{css}</style></head>
<body>
<div class="bcms-bar" aria-hidden="true"></div>
<header class="topbar">
  <div class="brand">
    {brasao}
    <div><h1>Crédito Disponível — BCMS</h1>
    <p class="subtitle">Batalhão Central de Manutenção e Suprimento · Tesouro Gerencial / SIAFI · UASGs 160329 (OGU) e 167329 (Fundo do Exército)</p></div>
  </div>
  <div class="topbar-r">
    <div class="selo-wrap"><span class="selo">Posição {esc(posicao)}</span><span class="selo-delay">⏱ dados com ~24h de defasagem</span></div>
    <button class="theme" id="themeBtn" aria-pressed="false" aria-label="Alternar tema claro/escuro" onclick="bcmsTheme()">
      <svg viewBox="0 0 24 24" class="ic-sun" aria-hidden="true"><circle cx="12" cy="12" r="4.5" style="fill:currentColor"/><g style="stroke:currentColor;stroke-width:1.6"><path d="M12 2v3M12 19v3M2 12h3M19 12h3M4.5 4.5l2 2M17.5 17.5l2 2M19.5 4.5l-2 2M6.5 17.5l-2 2"/></g></svg>
      <svg viewBox="0 0 24 24" class="ic-moon" aria-hidden="true"><path d="M20 14.5A8 8 0 019.5 4 8 8 0 1020 14.5z" style="fill:currentColor"/></svg>
    </button>
  </div>
</header>

<main class="wrap">
  {banner}

  <div class="toptabs" role="tablist" aria-label="Visões do painel">
    <button class="toptab on" role="tab" aria-selected="true" onclick="bcmsView(this,'resumo')">Resumo</button>
    <button class="toptab" role="tab" aria-selected="false" onclick="bcmsView(this,'completo')">Detalhamento completo</button>
  </div>

  <div id="view-resumo">
  {resumo_html}
  </div>

  <div id="view-completo" style="display:none">
  <section class="hero">
    <div class="hero-l">
      <div class="eyebrow">Crédito Disponível · Consolidado BCMS</div>
      <div class="hero-num num">{esc(brl(tot["cred"]))}</div>
      <div class="hero-eq">{hero_eq}</div>
      <div class="hero-eq-lbl"><span>Provisão Recebida</span><span>Empenhado</span><span>Disponível</span></div>
    </div>
    <div class="hero-r">
      {delta_html}
      {svg_util(tot["prov"], tot["emp"], tot["cred"])}
    </div>
  </section>

  <section class="sec">
    <div class="eyebrow">Composição — a subtração, visualmente</div>
    <div class="card">{svg_waterfall(tot["prov"], tot["emp"], tot["cred"])}</div>
  </section>

  <section class="sec">
    <div class="eyebrow">Indicadores de execução</div>
    <div class="kpis">{kpis}</div>
  </section>

  <section class="sec">
    <div class="eyebrow">Por fonte de recurso</div>
    <div class="grid2">{cards_uasg}</div>
  </section>

  <section class="sec">
    <div class="eyebrow">Onde está o crédito disponível</div>
    <div class="grid2">{ch_acao}{ch_nd}</div>
  </section>

  <section class="sec">
    <div class="eyebrow">Estágios da despesa por fonte</div>
    <div class="grid2">{funis}</div>
  </section>

  <section class="sec">{trend}</section>

  <section class="sec">
    <div class="eyebrow">Crédito Disponível em tela — por célula orçamentária</div>
    <p class="sec-nota">Saldo <b>líquido</b> por célula (Ação · PI · ND): <b>Recebido (líq) − Empenhado = Crédito Disponível</b> em cada linha. "Recebido (líq)" já compensa alterações de ND, detalhamentos e anulações (a alteração <b>não é somada</b> com a NC original). Só aparecem células com saldo &gt; 0; a soma fecha com o total consolidado. <b>Clique em uma linha</b> para ver as notas de crédito da célula com a descrição completa.</p>
    <div class="tabs" role="tablist" aria-label="Crédito em tela por UASG">{abas}</div>
    {tabs}
  </section>
  </div>
</main>

<div class="modal" id="modal" role="dialog" aria-modal="true" aria-labelledby="modal-title" aria-hidden="true" onclick="if(event.target===this)bcmsCelClose()">
  <div class="modal-panel">
    <button class="modal-x" aria-label="Fechar" onclick="bcmsCelClose()">✕</button>
    <div id="modal-body"></div>
  </div>
</div>

<footer class="rodape">
  <p class="rodape-brand">⚙ BCMS · Batalhão Central de Manutenção e Suprimento — Exército Brasileiro</p>
  <p><b>Metodologia:</b> Crédito Disponível = Provisão Recebida − Provisão Concedida − Despesas Empenhadas (saldo não empenhado "em tela" do Tesouro Gerencial). O detalhe é o <b>saldo líquido por célula orçamentária</b> (Ação · PI · ND): descentralizações, alterações de ND, detalhamentos, anulações e empenho são compensados dentro de cada célula, de modo que a alteração de ND <b>não é somada</b> à NC original. A soma das células reconcilia com o total consolidado.</p>
  <p>Fonte: CRÉDITO DISP.xlsx (Google Drive) · <b>⏱ Os dados do Tesouro Gerencial têm defasagem de aproximadamente 24 horas.</b> · Painel gerado em {esc(ger)}</p>
</footer>
<script>var CELDATA={celdata_json};var NCDATA={ncdata_json};var DAYDATA={daydata_json};</script>
<script>{js}</script>
</body></html>"""

# ============ CSS / JS (constantes) ============
CSS = r"""
*{box-sizing:border-box;margin:0;padding:0}
:root{
--bg:#EEF1F6;--surface:#FFFFFF;--surface-2:#E9EDF3;--ink:#0F1B2A;--ink-muted:#566374;
--border:#D8DEE7;--border-strong:#C2CBD6;--primary:#1C4A73;--primary-strong:#143A5C;
--success:#0F7A5A;--success-strong:#0B5E45;--warning:#B5822B;--warning-ink:#8A631C;
--danger:#B23A2E;--focus:#0B5E45;--hero-soft:#E7F1EC;--track:#E4E8EF;--gold:#C8901E;
--stg1:#1C4A73;--stg2:#3E6E9B;--stg3:#6D9AC0;
--shadow:0 1px 2px rgba(15,27,42,.06);--shadow-h:0 3px 10px rgba(15,27,42,.10);
--serif:Georgia,"Times New Roman","Nimbus Roman",serif;
--sans:-apple-system,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif;}
@media(prefers-color-scheme:dark){:root{
--bg:#0C1420;--surface:#131E2E;--surface-2:#1A2838;--ink:#E7EDF4;--ink-muted:#97A6B8;
--border:#26374C;--border-strong:#35495F;--primary:#4E86BD;--primary-strong:#6BA0D4;
--success:#35C08F;--success-strong:#4FD0A2;--warning:#D6A24A;--warning-ink:#E0B463;
--danger:#E5705E;--focus:#4FD0A2;--hero-soft:#10241C;--track:#1F2E3F;
--stg1:#4E86BD;--stg2:#6BA0D4;--stg3:#94BFE0;--shadow:none;--shadow-h:0 3px 10px rgba(0,0,0,.3);}}
:root[data-theme=dark]{
--bg:#0C1420;--surface:#131E2E;--surface-2:#1A2838;--ink:#E7EDF4;--ink-muted:#97A6B8;
--border:#26374C;--border-strong:#35495F;--primary:#4E86BD;--primary-strong:#6BA0D4;
--success:#35C08F;--success-strong:#4FD0A2;--warning:#D6A24A;--warning-ink:#E0B463;
--danger:#E5705E;--focus:#4FD0A2;--hero-soft:#10241C;--track:#1F2E3F;
--stg1:#4E86BD;--stg2:#6BA0D4;--stg3:#94BFE0;--shadow:none;--shadow-h:0 3px 10px rgba(0,0,0,.3);}
:root[data-theme=light]{
--bg:#EEF1F6;--surface:#FFFFFF;--surface-2:#E9EDF3;--ink:#0F1B2A;--ink-muted:#566374;
--border:#D8DEE7;--border-strong:#C2CBD6;--primary:#1C4A73;--primary-strong:#143A5C;
--success:#0F7A5A;--success-strong:#0B5E45;--warning:#B5822B;--warning-ink:#8A631C;
--danger:#B23A2E;--focus:#0B5E45;--hero-soft:#E7F1EC;--track:#E4E8EF;--gold:#C8901E;
--stg1:#1C4A73;--stg2:#3E6E9B;--stg3:#6D9AC0;--shadow:0 1px 2px rgba(15,27,42,.06);--shadow-h:0 3px 10px rgba(15,27,42,.10);}
html{transition:background-color .15s,color .15s}
body{background:var(--bg);color:var(--ink);font:15px/1.5 var(--sans);-webkit-font-smoothing:antialiased}
.num{font-variant-numeric:tabular-nums lining-nums;font-feature-settings:"tnum" 1,"lnum" 1}
.visually-hidden{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0)}
.wrap{max-width:1200px;margin:0 auto;padding:0 24px 64px}
.eyebrow{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.14em;color:var(--ink-muted);margin-bottom:12px}
.sec-nota{font-size:12.5px;color:var(--ink-muted);margin:-4px 0 12px;max-width:900px;line-height:1.55}
.sec-nota b{color:var(--ink)}
.sec{margin-top:30px}
/* abas de topo */
.toptabs{display:flex;gap:6px;margin:20px 0 6px;background:var(--surface-2);border:1px solid var(--border);border-radius:11px;padding:5px}
.toptab{flex:1;background:none;border:none;color:var(--ink-muted);font:650 14px var(--sans);padding:10px 14px;border-radius:8px;cursor:pointer}
.toptab.on{background:var(--surface);color:var(--ink);box-shadow:var(--shadow)}
.toptab:hover:not(.on){color:var(--ink)}
/* resumo hero */
.resumo-hero{display:flex;justify-content:space-between;align-items:center;gap:16px;flex-wrap:wrap;background:var(--hero-soft);border:1px solid var(--border);border-left:4px solid var(--success);border-radius:12px;padding:16px 22px;margin-top:18px}
.rh-l{display:block;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:var(--ink-muted)}
.rh-v{display:block;font-size:30px;font-weight:700;color:var(--success-strong);line-height:1.1;margin-top:3px}
.rh-meta{font-size:12.5px;color:var(--ink-muted);text-align:right}
.rh-delay{color:var(--warning-ink);font-weight:600}
.col-pos{color:var(--success-strong)}.col-neg{color:var(--danger)}.col-warn{color:var(--warning-ink)}
/* módulo créditos em tela */
.et-head{display:flex;flex-wrap:wrap;gap:12px;align-items:stretch;margin-bottom:12px}
.et-kpi{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:11px 16px;display:flex;flex-direction:column;gap:2px;min-width:130px;box-shadow:var(--shadow)}
.et-kpi span{font-size:11px;text-transform:uppercase;letter-spacing:.04em;color:var(--ink-muted)}
.et-kpi b{font-size:19px;font-weight:700}
.et-hero{background:var(--hero-soft);border-left:4px solid var(--success)}
.et-hero b{font-size:24px;color:var(--success-strong)}
.et-meta{margin-left:auto;align-self:center;font-size:12px;color:var(--ink-muted);text-align:right;line-height:1.5}
.m-kpis b.neg{color:var(--danger)}.m-kpis b.op{font-size:12px;font-weight:600;line-height:1.3}
.card{background:var(--surface);border:1px solid var(--border);border-radius:12px;box-shadow:var(--shadow);padding:16px}
/* topbar */
.topbar{max-width:1200px;margin:0 auto;padding:18px 24px;display:flex;justify-content:space-between;align-items:center;gap:14px;border-bottom:1px solid var(--border)}
.brand{display:flex;align-items:center;gap:12px}
.brasao{width:44px;height:auto;flex:none;filter:drop-shadow(0 1px 1px rgba(0,0,0,.12))}
.bcms-bar{height:5px;background:linear-gradient(to bottom,#CE2B2B 50%,#1E6FD0 50%)}
h1{font-family:var(--serif);font-size:26px;font-weight:600;letter-spacing:-.01em;line-height:1.15}
.subtitle{font-size:12.5px;color:var(--ink-muted);margin-top:2px}
.topbar-r{display:flex;align-items:center;gap:10px}
.selo-wrap{display:flex;flex-direction:column;align-items:flex-end;gap:3px}
.selo{background:var(--surface-2);color:var(--ink-muted);border:1px solid var(--border);border-radius:999px;padding:5px 13px;font-size:12.5px;font-weight:600;white-space:nowrap}
.selo-delay{font-size:10.5px;color:var(--warning-ink);font-weight:600;white-space:nowrap}
.theme{width:36px;height:36px;border:1px solid var(--border);background:var(--surface);border-radius:8px;color:var(--ink-muted);cursor:pointer;display:grid;place-items:center}
.theme:hover{border-color:var(--border-strong);color:var(--ink)}
.theme svg{width:18px;height:18px}
.ic-moon{display:none}
:root[data-theme=dark] .ic-sun,html:not([data-theme]) .ic-sun{display:block}
:root[data-theme=dark] .ic-moon{display:block}:root[data-theme=dark] .ic-sun{display:none}
@media(prefers-color-scheme:dark){html:not([data-theme]) .ic-sun{display:none}html:not([data-theme]) .ic-moon{display:block}}
/* banner */
.banner{background:color-mix(in srgb,var(--danger) 12%,var(--surface));border:1px solid var(--danger);border-radius:10px;padding:12px 16px;margin-top:20px;font-size:13px}
.banner ul{margin:6px 0 0 18px}
/* hero */
.hero{margin-top:24px;background:var(--hero-soft);border:1px solid var(--border);border-left:4px solid var(--success);border-radius:12px;padding:22px 26px;display:grid;grid-template-columns:1.35fr 1fr;gap:24px;align-items:center}
.hero-num{font-size:48px;font-weight:700;letter-spacing:-.02em;color:var(--success-strong);line-height:1.05;margin:6px 0 12px}
.hero-eq{display:flex;flex-wrap:wrap;align-items:baseline;gap:10px;font-size:15px}
.hero-eq .eq-op{font-style:normal;font-size:24px;font-weight:300;color:var(--ink-muted)}
.eq-t{font-variant-numeric:tabular-nums}
.eq-prov{color:var(--ink);font-weight:600}.eq-emp{color:var(--warning-ink);font-weight:600}
.eq-disp{color:var(--success-strong);font-weight:700}
.hero-eq-lbl{display:flex;gap:10px;font-size:11px;color:var(--ink-muted);margin-top:4px}
.hero-eq-lbl span:nth-child(1){flex:0 0 auto}.hero-eq-lbl span{padding-right:24px}
.delta{display:inline-flex;align-items:center;gap:6px;font-size:13px;font-weight:600;margin-bottom:10px;padding:4px 10px;border-radius:999px;background:var(--surface-2);border:1px solid var(--border)}
.delta.up{color:var(--success-strong)}.delta.down{color:var(--danger)}.delta.flat{color:var(--ink-muted);font-weight:500}
.delta small{font-weight:400;color:var(--ink-muted)}
/* svg text classes */
.svg{width:100%;height:auto;display:block}
.s-lbl{font-size:11px;font-weight:700;letter-spacing:.06em;fill:var(--ink-muted)}
.s-seg{font-size:11px;fill:var(--warning-ink);font-weight:600}
.s-seg-ok{fill:var(--success-strong)}
.s-brk{stroke:var(--border-strong);stroke-width:1}
.s-cat{font-size:12px;fill:var(--ink-muted)}.s-cat-ok{fill:var(--success-strong);font-weight:600}
.s-val{font-size:13px;font-weight:700}.s-on{fill:#fff}.s-ok{fill:var(--success-strong)}.s-warn{fill:var(--warning-ink)}
.s-on2{fill:var(--ink)}
.s-num{font-size:11px;font-weight:600;font-variant-numeric:tabular-nums}.s-neg{fill:var(--danger)}
.s-conn{stroke:var(--border-strong);stroke-width:1;stroke-dasharray:4 3}
.s-zero{stroke:var(--border-strong);stroke-width:1.5}
.s-grid{stroke:var(--border);stroke-width:1}.s-ax{font-size:11px;fill:var(--ink-muted)}
.s-line{fill:none;stroke:var(--success);stroke-width:2.4}.s-area{fill:var(--success);opacity:.10}
.s-dot{fill:var(--success)}.s-conv{font-size:10.5px;fill:var(--ink-muted)}
/* kpis */
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}
.kpi{background:var(--surface);border:1px solid var(--border);border-left:3px solid var(--border);border-radius:10px;padding:15px 16px;box-shadow:var(--shadow)}
.kpi-prov{border-left-color:var(--primary)}.kpi-emp{border-left-color:var(--warning)}
.kpi-liq{border-left-color:var(--stg2)}.kpi-pag{border-left-color:var(--success)}
.kpi-l{font-size:12.5px;font-weight:600;letter-spacing:.03em;color:var(--ink-muted);text-transform:uppercase}
.kpi-v{font-size:24px;font-weight:700;margin:5px 0}
.chip{display:inline-block;font-size:11px;color:var(--ink-muted);background:var(--surface-2);border-radius:999px;padding:2px 9px}
/* grids */
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
/* uasg */
.uasg{display:flex;flex-direction:column;gap:8px}
.uasg-h{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.uasg-cod{font-size:16px;font-weight:700;letter-spacing:.02em}
.uasg-fonte{font-size:11px;font-weight:700;color:var(--primary);border:1px solid var(--primary);border-radius:5px;padding:1px 6px}
.uasg-nome{font-size:12.5px;color:var(--ink-muted)}
.uasg-disp{display:flex;justify-content:space-between;align-items:baseline;margin-top:2px}
.uasg-disp-l{font-size:12px;color:var(--ink-muted)}
.uasg-disp-v{font-size:22px;font-weight:700;color:var(--success-strong)}
.uasg-eq{font-size:12px;color:var(--ink-muted)}.uasg-eq i{font-style:normal;color:var(--warning-ink)}
.uasg-exec{margin-top:2px}
.exec-l{display:flex;justify-content:space-between;font-size:11.5px;color:var(--ink-muted);margin-bottom:4px}
.exec-track{height:7px;background:var(--track);border-radius:4px;overflow:hidden}
.exec-fill{height:100%;background:var(--primary);border-radius:4px}
/* chart */
.chart .eyebrow{margin-bottom:8px}.chart.wide{grid-column:1/-1}
.vazio{color:var(--ink-muted);font-size:12px;font-style:italic;padding:8px 0}
/* tabs + table */
.tabs{display:flex;gap:4px;border-bottom:1px solid var(--border);margin-bottom:12px;overflow-x:auto}
.tab{background:none;border:none;border-bottom:2px solid transparent;color:var(--ink-muted);font:600 13px var(--sans);padding:9px 14px;cursor:pointer;white-space:nowrap}
.tab.on{color:var(--success-strong);border-bottom-color:var(--success)}
.tab:hover{color:var(--ink)}
.tbl-tools{display:flex;align-items:center;gap:12px;margin-bottom:10px;flex-wrap:wrap}
.tbl-search{flex:1;min-width:220px;background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:9px 12px;color:var(--ink);font:14px var(--sans)}
.tbl-count{font-size:12px;color:var(--ink-muted)}
.tbl-scroll{overflow-x:auto;border:1px solid var(--border);border-radius:12px;background:var(--surface)}
table.det{border-collapse:collapse;width:100%;font-size:14px}
.det th{position:sticky;top:0;background:var(--surface-2);color:var(--ink-muted);font-size:11.5px;text-transform:uppercase;letter-spacing:.04em;text-align:left;padding:10px 12px;cursor:pointer;white-space:nowrap;border-bottom:1px solid var(--border)}
.det th.num,.det td.num{text-align:right}
.det th .sort{display:inline-block;width:10px;color:var(--success)}
.det td{padding:9px 12px;border-bottom:1px solid var(--border)}
.det td.num{font-variant-numeric:tabular-nums;font-weight:500}
.det td.anchor{font-weight:700}
.det tbody tr:hover{background:var(--surface-2)}
.det .obj{max-width:340px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--ink-muted)}
.det .mono2{font-variant-numeric:tabular-nums;font-size:12.5px;color:var(--ink-muted)}
.cell-neg{color:var(--danger);box-shadow:inset 3px 0 0 var(--danger);background:color-mix(in srgb,var(--danger) 6%,transparent)}
.pill-fonte{font-size:10.5px;font-weight:700;color:var(--primary);border:1px solid var(--border-strong);border-radius:5px;padding:1px 6px}
.det tfoot td{padding:10px 12px;font-weight:700;background:var(--surface-2);border-top:1px solid var(--border-strong)}
.cel-row{cursor:pointer}.cel-row .chev{float:right;margin-left:8px;color:var(--ink-muted);font-weight:400;transition:transform .12s,color .12s}
.cel-row:hover .chev{color:var(--success-strong);transform:translateX(2px)}
/* modal drill-down */
.modal{position:fixed;inset:0;z-index:50;display:none;align-items:flex-start;justify-content:center;padding:40px 16px;background:rgba(8,14,24,.55);overflow-y:auto}
.modal.open{display:flex}
.modal-panel{position:relative;background:var(--surface);border:1px solid var(--border);border-radius:14px;box-shadow:0 20px 60px rgba(0,0,0,.35);max-width:760px;width:100%;padding:22px 24px}
.modal-x{position:absolute;top:12px;right:12px;width:32px;height:32px;border:1px solid var(--border);background:var(--surface-2);border-radius:8px;color:var(--ink-muted);cursor:pointer;font-size:14px;line-height:1}
.modal-x:hover{color:var(--ink);border-color:var(--border-strong)}
#modal-body h3{font-size:18px;font-weight:650;margin-bottom:2px;padding-right:36px}
.m-sub{font-size:13px;color:var(--ink-muted);margin-bottom:14px}
.m-ficha{display:flex;flex-wrap:wrap;gap:9px 18px;margin:2px 0 14px;padding:12px 14px;background:var(--surface-2);border:1px solid var(--border);border-radius:10px}
.m-ficha span{display:flex;flex-direction:column;gap:1px;font-size:10.5px;text-transform:uppercase;letter-spacing:.04em;color:var(--ink-muted);min-width:110px}
.m-ficha span.wide{flex-basis:100%}
.m-ficha b{font-size:13px;font-weight:600;color:var(--ink);text-transform:none;letter-spacing:0}
.m-kpis{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:16px}
.m-kpis span{flex:1;min-width:120px;display:flex;flex-direction:column;gap:2px;font-size:11px;color:var(--ink-muted);text-transform:uppercase;letter-spacing:.03em;background:var(--surface-2);border:1px solid var(--border);border-radius:9px;padding:9px 11px}
.m-kpis b{font-size:17px;font-weight:700;color:var(--ink);font-variant-numeric:tabular-nums}
.m-kpis .ok b{color:var(--success-strong)}.m-kpis .ok{border-left:3px solid var(--success)}
.m-formula{font-size:11.5px;color:var(--ink-muted);margin:-8px 0 14px;font-style:italic}
.m-ncs-h{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--ink-muted);margin-bottom:8px}
.m-ncs{display:flex;flex-direction:column;gap:8px;max-height:48vh;overflow-y:auto}
.m-nc{border:1px solid var(--border);border-radius:10px;padding:10px 12px;background:var(--surface)}
.m-nc-h{display:flex;justify-content:space-between;align-items:baseline;gap:10px}
.m-nc-num{font-variant-numeric:tabular-nums;font-weight:700;font-size:13px}
.m-nc-val{font-variant-numeric:tabular-nums;font-weight:700;font-size:13px;color:var(--success-strong);white-space:nowrap}
.m-nc-val.neg{color:var(--danger)}
.m-nc-op{font-size:11px;text-transform:uppercase;letter-spacing:.03em;color:var(--ink-muted);margin-top:3px}
.m-nc-desc{font-size:12.5px;color:var(--ink);margin-top:5px;line-height:1.5;white-space:pre-wrap}
/* footer */
.rodape{max-width:1200px;margin:40px auto 0;padding:20px 24px;border-top:1px solid var(--border);color:var(--ink-muted);font-size:12px;line-height:1.7}
.rodape b{color:var(--ink)}
.rodape-brand{color:var(--warning-ink);font-weight:700;font-size:13px;letter-spacing:.02em;margin-bottom:8px}
/* focus + motion */
:focus-visible{outline:2px solid var(--focus);outline-offset:2px;border-radius:4px}
@media(prefers-reduced-motion:reduce){*{transition:none!important;animation:none!important}}
/* responsive */
@media(max-width:1023px){.grid2{grid-template-columns:1fr}.hero{grid-template-columns:1fr}}
@media(max-width:640px){
.wrap{padding:0 16px 48px}.topbar{padding:14px 16px;flex-wrap:wrap}
h1{font-size:22px}.subtitle{display:none}
.hero-num{font-size:34px}.kpis{grid-template-columns:repeat(2,1fr)}
.det td.mono2,.det th:first-child,.det td:first-child{position:sticky;left:0;background:var(--surface)}
.det th:first-child{background:var(--surface-2)}}
"""

JS = r"""
(function(){var s=localStorage.getItem('bcms-theme');if(s){document.documentElement.setAttribute('data-theme',s);}})();
function bcmsTheme(){var h=document.documentElement;var cur=h.getAttribute('data-theme');
 var dark=cur?cur==='dark':window.matchMedia('(prefers-color-scheme:dark)').matches;
 var next=dark?'light':'dark';h.setAttribute('data-theme',next);localStorage.setItem('bcms-theme',next);
 document.getElementById('themeBtn').setAttribute('aria-pressed',next==='dark');}
function bcmsTab(btn,id){
 var list=btn.parentNode.querySelectorAll('.tab');
 list.forEach(function(b){b.classList.remove('on');b.setAttribute('aria-selected','false');b.tabIndex=-1;});
 btn.classList.add('on');btn.setAttribute('aria-selected','true');btn.tabIndex=0;
 document.querySelectorAll('.tabpanel').forEach(function(p){p.style.display='none';});
 document.getElementById(id).style.display='block';}
function bcmsTabKey(e,btn){var t=Array.prototype.slice.call(btn.parentNode.querySelectorAll('.tab'));
 var i=t.indexOf(btn);if(e.key==='ArrowRight'){e.preventDefault();t[(i+1)%t.length].focus();t[(i+1)%t.length].click();}
 else if(e.key==='ArrowLeft'){e.preventDefault();t[(i-1+t.length)%t.length].focus();t[(i-1+t.length)%t.length].click();}}
function bcmsView(btn,which){
 btn.parentNode.querySelectorAll('.toptab').forEach(function(b){b.classList.remove('on');b.setAttribute('aria-selected','false');});
 btn.classList.add('on');btn.setAttribute('aria-selected','true');
 document.getElementById('view-resumo').style.display=which==='resumo'?'':'none';
 document.getElementById('view-completo').style.display=which==='completo'?'':'none';
 window.scrollTo(0,0);}
function bcmsSearch(inp,tid){var q=inp.value.toLowerCase();var tb=document.getElementById(tid).querySelector('tbody');
 var rows=tb.querySelectorAll('tr');var n=0;
 rows.forEach(function(r){var ok=r.textContent.toLowerCase().indexOf(q)>-1;r.style.display=ok?'':'none';if(ok)n++;});
 var cnt=document.getElementById('cnt-'+tid);var u=cnt.getAttribute('data-unit')||'linha(s)';cnt.textContent=n+' '+u;}
function bcmsNC(row){var d=NCDATA[row.getAttribute('data-nc')];if(!d)return;var neg=d.val<0;
 var h='<h3 id="modal-title">NC '+bcmsEsc(d.nc)+'</h3>';
 h+='<p class="m-sub">'+bcmsEsc((d.u?d.u+' · ':'')+'Ação '+d.acao+' · PI '+d.pi+' · ND '+d.nd+(d.ndn?' — '+d.ndn:''))+'</p>';
 h+='<div class="m-kpis"><span>Data<b>'+bcmsEsc(d.dia||'—')+'</b></span><span>Operação<b class="op">'+bcmsEsc(d.op||'—')+'</b></span><span class="'+(neg?'':'ok')+'">Valor<b class="'+(neg?'neg':'')+'">'+bcmsBRL(d.val)+'</b></span></div>';
 h+='<div class="m-ncs-h">Descrição completa</div>';
 h+='<div class="m-nc"><div class="m-nc-desc">'+bcmsEsc(d.obj||'(sem descrição)')+'</div></div>';
 document.getElementById('modal-body').innerHTML=h;
 var m=document.getElementById('modal');m.classList.add('open');m.setAttribute('aria-hidden','false');
 var x=document.querySelector('.modal-x');if(x)x.focus();}
function bcmsSort(th){var table=th.closest('table');var idx=Array.prototype.indexOf.call(th.parentNode.children,th);
 var dir=th.getAttribute('aria-sort')==='ascending'?'descending':'ascending';
 th.parentNode.querySelectorAll('th').forEach(function(h){h.setAttribute('aria-sort','none');h.querySelector('.sort').textContent='';});
 th.setAttribute('aria-sort',dir);th.querySelector('.sort').textContent=dir==='ascending'?' ▲':' ▼';
 var tb=table.querySelector('tbody');var rows=Array.prototype.slice.call(tb.querySelectorAll('tr'));
 rows.sort(function(a,b){var ca=a.children[idx],cb=b.children[idx];
  var da=ca.getAttribute('data-sort'),db=cb.getAttribute('data-sort');
  var va,vb;if(da!==null&&db!==null){va=parseFloat(da);vb=parseFloat(db);}else{va=ca.textContent.trim().toLowerCase();vb=cb.textContent.trim().toLowerCase();}
  if(va<vb)return dir==='ascending'?-1:1;if(va>vb)return dir==='ascending'?1:-1;return 0;});
 rows.forEach(function(r){tb.appendChild(r);});}
function bcmsEsc(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
function bcmsBRL(v){var neg=v<0,s=Math.abs(v).toFixed(2).split('.');var i=s[0].replace(/\B(?=(\d{3})+(?!\d))/g,'.');return (neg?'−R$ ':'R$ ')+i+','+s[1];}
function bcmsCel(row){var d=CELDATA[row.getAttribute('data-cel')];if(!d)return;
 var h='<h3 id="modal-title">'+bcmsEsc(d.t)+'</h3>';
 var fonte=d.u==='OGU'?'OGU (Orçamento Geral da União)':(d.u==='FEx'?'FEx (Fundo do Exército)':(d.u||'—'));
 h+='<div class="m-ficha">'
   +'<span>UASG (Executora)<b>'+bcmsEsc(d.uasg||'—')+'</b></span>'
   +'<span>Fonte<b>'+bcmsEsc(fonte)+'</b></span>'
   +'<span>Ação Governo<b>'+bcmsEsc(d.acao||'—')+'</b></span>'
   +'<span class="wide">PI (Plano Interno)<b>'+bcmsEsc(d.pi||'—')+(d.pinome?' — '+bcmsEsc(d.pinome):'')+'</b></span>'
   +'<span class="wide">ND (Natureza de Despesa)<b>'+bcmsEsc(d.nd||'—')+(d.ndnome?' — '+bcmsEsc(d.ndnome):'')+'</b></span>'
   +'</div>';
 h+='<div class="m-kpis"><span>Recebido (líq)<b>'+bcmsBRL(d.r)+'</b></span><span>Empenhado<b>'+bcmsBRL(d.e)+'</b></span><span>Liquidado<b>'+bcmsBRL(d.l||0)+'</b></span><span>Pago<b>'+bcmsBRL(d.p||0)+'</b></span><span class="ok">Crédito Disponível<b>'+bcmsBRL(d.d)+'</b></span></div>';
 h+='<p class="m-formula">Recebido (líq) − Empenhado = Crédito Disponível · Empenhado ≥ Liquidado ≥ Pago</p>';
 var itens='';
 d.ncs.forEach(function(n){if(!n[0])return; var neg=n[2]<0;
  var meta=[]; if(n[4])meta.push('Emitente '+bcmsEsc(n[4])); if(n[1])meta.push(bcmsEsc(n[1])); if(n[5])meta.push('em '+bcmsEsc(n[5]));
  itens+='<div class="m-nc"><div class="m-nc-h"><span class="m-nc-num">'+bcmsEsc(n[0])+'</span><span class="m-nc-val'+(neg?' neg':'')+'">'+bcmsBRL(n[2])+'</span></div>';
  if(meta.length)itens+='<div class="m-nc-op">'+meta.join(' · ')+'</div>';
  if(n[3])itens+='<div class="m-nc-desc">'+bcmsEsc(n[3])+'</div>';
  itens+='</div>';});
 var nq=d.ncs.filter(function(n){return n[0];}).length;
 h+='<div class="m-ncs-h">Notas de crédito da célula ('+nq+')</div>';
 h+='<div class="m-ncs">'+(itens||'<p class="vazio">Sem notas de crédito para detalhar.</p>')+'</div>';
 document.getElementById('modal-body').innerHTML=h;
 var m=document.getElementById('modal');m.classList.add('open');m.setAttribute('aria-hidden','false');
 var x=document.querySelector('.modal-x');if(x)x.focus();}
function bcmsDay(row){var d=DAYDATA[row.getAttribute('data-day')];if(!d)return;
 var h='<h3 id="modal-title">Movimentação de '+bcmsEsc(d.d)+'</h3>';
 h+='<div class="m-kpis"><span>Nº de NC<b>'+d.n+'</b></span><span>Recebido<b class="col-pos">'+bcmsBRL(d.rec)+'</b></span><span>Reduções<b class="col-neg">'+bcmsBRL(d.red)+'</b></span><span class="ok">Líquido<b>'+bcmsBRL(d.liq)+'</b></span></div>';
 h+='<div class="m-ncs-h">Notas de crédito do dia ('+d.ncs.length+')</div>';
 var itens='';
 d.ncs.forEach(function(n){var neg=n[3]<0;
  itens+='<div class="m-nc"><div class="m-nc-h"><span class="m-nc-num">'+bcmsEsc(n[0])+' <span class="pill-fonte">'+bcmsEsc(n[1])+'</span></span><span class="m-nc-val'+(neg?' neg':'')+'">'+bcmsBRL(n[3])+'</span></div>';
  if(n[2])itens+='<div class="m-nc-op">'+bcmsEsc(n[2])+'</div>';
  if(n[4])itens+='<div class="m-nc-desc">'+bcmsEsc(n[4])+'</div>';
  itens+='</div>';});
 h+='<div class="m-ncs">'+(itens||'<p class="vazio">Sem NC neste dia.</p>')+'</div>';
 document.getElementById('modal-body').innerHTML=h;
 var m=document.getElementById('modal');m.classList.add('open');m.setAttribute('aria-hidden','false');
 var x=document.querySelector('.modal-x');if(x)x.focus();}
function bcmsCelClose(){var m=document.getElementById('modal');if(m){m.classList.remove('open');m.setAttribute('aria-hidden','true');}}
document.addEventListener('keydown',function(e){if(e.key==='Escape')bcmsCelClose();});
"""

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
    res, periodo, alertas = etl(path)
    for a in alertas:
        print("[ALERTA]", a)
    hist = atualizar_historico(res, data_str)
    html_out = gerar_html(res, hist, data_str, periodo, alertas)

    os.makedirs(SITE, exist_ok=True)
    os.makedirs(os.path.join(SITE, "data"), exist_ok=True)
    with open(os.path.join(SITE, "index.html"), "w", encoding="utf-8") as f:
        f.write(html_out)
    with open(os.path.join(SITE, "data", "history.json"), "w", encoding="utf-8") as f:
        json.dump(hist, f, ensure_ascii=False, indent=1)

    tot = {k: sum(res[c][k] for c, _ in ALVOS) for k in ("prov", "cred", "emp", "liq", "pag")}
    print(f"OK -> {os.path.join(SITE,'index.html')}")
    print(f"Periodo={periodo} | Credito Disp total={brl(tot['cred'])} | historico {len(hist)} dia(s)")

if __name__ == "__main__":
    main()
