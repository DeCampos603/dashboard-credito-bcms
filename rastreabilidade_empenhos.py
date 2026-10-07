# -*- coding: utf-8 -*-
"""
Módulo de Integração, Rastreabilidade e Auditoria de Empenhos (SICOF-BaApLog)
Responsável pelo ETL da planilha de empenhos do Tesouro Gerencial, conciliação
relacional com as Notas de Crédito (NCs), Prova da Verdade de integridade orçamentária,
geração dos componentes visuais e scripts interativos alinhados 100% ao Design System do Dashboard.
"""

import os, sys, json, re, html, datetime, urllib.request, tempfile
import openpyxl

DEFAULT_EMPENHOS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSTGXyEQ2v4UYga0nkP7ypbyumOZd68m32UZwqVVQhx-bHBACkBSj7v4RiNWZydIQ/pub?output=xlsx"

def baixar_empenhos(url=None):
    """Baixa o extrato de empenhos emitidos publicado no Google Sheets / Drive com retry e fallback."""
    target_url = url or DEFAULT_EMPENHOS_URL
    tmp = os.path.join(tempfile.gettempdir(), "empenhos_baaplog_download.xlsx")
    req = urllib.request.Request(target_url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
    
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=120) as r:
                content = bytearray()
                while True:
                    chunk = r.read(65536)
                    if not chunk:
                        break
                    content.extend(chunk)
                if len(content) > 1000:
                    with open(tmp, "wb") as f:
                        f.write(content)
                    return tmp
        except Exception as e:
            if attempt == 2 and os.path.exists(tmp) and os.path.getsize(tmp) > 1000:
                print(f"[AVISO] Download de empenhos falhou ({e}), usando cache existente em {tmp}")
                return tmp
            import time
            time.sleep(1)

    if os.path.exists(tmp) and os.path.getsize(tmp) > 1000:
        return tmp
    raise RuntimeError("Não foi possível baixar a base de empenhos após 3 tentativas.")

def _fmt_brl(v):
    if v is None: return "R$ 0,00"
    s = f"{abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("−R$ " if v < 0 else "R$ ") + s

def _esc(s):
    return html.escape(str(s or ""))

def etl_empenhos(path, res_credito, todas_ncs_dict=None):
    """Processa a planilha de empenhos, executa a Prova da Verdade e vincula às Notas de Crédito."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb['EMPENHO'] if 'EMPENHO' in wb.sheetnames else wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    nc_pat = re.compile(r'(202\dNC\d{6})', re.IGNORECASE)
    proc_pat = re.compile(r'\b(PE\s*\d+[/_\-]\d+|PREG[ÃA]O\s*(?:ELETR[ÔO]NICO)?\s*(?:N[º°])?\s*\d+[/_\-]\d+|DISPENSA\s*(?:N[º°])?\s*\d+[/_\-]\d+|INEX\w*\s*(?:N[º°])?\s*\d+[/_\-]\d+|DIEX\s*(?:N[º°])?\s*\d+[/_\-]\d+)', re.IGNORECASE)

    # 1. Mapeamento do Acervo de Créditos para a PROVA DA VERDADE e LIQUIDAÇÃO SIAFI
    cred_audit_db = {} # nc_code / nc_raw -> list of dicts
    cel_to_ncs = {}    # (ug, pi, nd) -> list of nc_codes
    cel_liq_db = {}    # (ug, pi, nd) -> dict com liq, pag, emp, prov
    cel_elem_liq_db = {} # (ug, pi, nd_elem) -> dict
    pi_liq_db = {}     # (ug, pi) -> dict
    global_pi_liq_db = {} # pi -> dict
    tot_liq_global = 0.0
    tot_pag_global = 0.0

    for cod, d in res_credito.items():
        tot_liq_global += d.get("liq", 0.0)
        tot_pag_global += d.get("pag", 0.0)
        for (acao, pi_c, nd_c), cel in d.get("celulas", {}).items():
            k_cel = (cod, pi_c, nd_c)
            cel_liq_db[k_cel] = {
                "liq": cel.get("liq", 0.0),
                "pag": cel.get("pag", 0.0),
                "emp": cel.get("emp", 0.0),
                "prov": cel.get("prov", 0.0),
                "cred": cel.get("cred", 0.0),
                "acao": acao
            }
            nd_elem = nd_c[:4] + "00" if len(nd_c) >= 6 else nd_c
            k_elem = (cod, pi_c, nd_elem)
            if k_elem not in cel_elem_liq_db:
                cel_elem_liq_db[k_elem] = {"liq": 0.0, "pag": 0.0, "emp": 0.0}
            cel_elem_liq_db[k_elem]["liq"] += cel.get("liq", 0.0)
            cel_elem_liq_db[k_elem]["pag"] += cel.get("pag", 0.0)
            cel_elem_liq_db[k_elem]["emp"] += cel.get("emp", 0.0)

            k_pi = (cod, pi_c)
            if k_pi not in pi_liq_db:
                pi_liq_db[k_pi] = {"liq": 0.0, "pag": 0.0, "emp": 0.0}
            pi_liq_db[k_pi]["liq"] += cel.get("liq", 0.0)
            pi_liq_db[k_pi]["pag"] += cel.get("pag", 0.0)
            pi_liq_db[k_pi]["emp"] += cel.get("emp", 0.0)

            if pi_c not in global_pi_liq_db:
                global_pi_liq_db[pi_c] = {"liq": 0.0, "pag": 0.0, "emp": 0.0}
            global_pi_liq_db[pi_c]["liq"] += cel.get("liq", 0.0)
            global_pi_liq_db[pi_c]["pag"] += cel.get("pag", 0.0)
            global_pi_liq_db[pi_c]["emp"] += cel.get("emp", 0.0)

        for L in d.get("linhas", []):
            nc_raw = L.get("nc")
            if not nc_raw: continue
            m = nc_pat.search(nc_raw)
            nc_code = m.group(1).upper() if m else nc_raw
            
            c_entry = {
                "ug": cod,
                "nc_full": nc_raw,
                "emit": L.get("emit") or "",
                "emit_nome": L.get("emit_nome") or "",
                "pi": L.get("pi") or "",
                "nd": L.get("nd") or "",
                "prov": L.get("prov", 0.0) or L.get("cred", 0.0),
                "obj": L.get("obj") or ""
            }
            cred_audit_db.setdefault(nc_code, []).append(c_entry)
            cred_audit_db.setdefault(nc_raw, []).append(c_entry)

            k_cel = (cod, L.get("pi"), L.get("nd"))
            cel_to_ncs.setdefault(k_cel, []).append(nc_code)
            cel_to_ncs.setdefault(k_cel, []).append(nc_raw)
            # Elemento genérico (ex: 339000 para 339039)
            nd_elem = (L.get("nd") or "")[:4] + "00"
            k_elem = (cod, L.get("pi"), nd_elem)
            cel_to_ncs.setdefault(k_elem, []).append(nc_code)
            cel_to_ncs.setdefault(k_elem, []).append(nc_raw)

    ug_cur = ""
    ug_nome_cur = ""
    nes_list = []
    nc_to_nes = {}
    top_forn = {}
    top_procs = {}
    tot_val = 0.0

    inconsistencias_citacao = []
    estat_prova = {
        "CONFIRMADO": 0,
        "CELULA_SIAFI": 0,
        "DIVERGENCIA_PI": 0,
        "DIVERGENCIA_UG": 0,
        "DIVERGENCIA_ND": 0,
        "EXERCICIO_ANTERIOR": 0,
        "NC_NAO_ENCONTRADA": 0
    }

    def _add_to_nc_to_nes(k, item):
        if not k: return
        lst = nc_to_nes.setdefault(k, [])
        if not any(x["ne"] == item["ne"] for x in lst):
            lst.append(item)

    for r in rows[5:]:
        if not r: continue
        if r[0] is not None and str(r[0]).strip() != "":
            ug_cur = str(r[0]).strip().replace("'", "")
        if r[1] is not None and str(r[1]).strip() != "":
            ug_nome_cur = str(r[1]).strip().replace("'", "")

        ne_full = str(r[3] or "").strip().replace("'", "")
        if not ne_full or "NE" not in ne_full or ne_full == "-9":
            continue

        dia = str(r[2] or "").strip()
        fav_doc = str(r[4] or "").strip().replace("'", "")
        fav_nome = str(r[5] or "").strip()
        val_raw = r[6]
        try:
            val = float(str(val_raw).replace("'", "").strip())
        except Exception:
            val = 0.0

        nd = str(r[7] or "").strip()
        nd_desc = str(r[8] or "").strip()
        pi = str(r[9] or "").strip()
        pi_desc = str(r[10] or "").strip()
        desc = str(r[11] or "").strip()

        # Saldo das contas contábeis
        try: mov_emp = float(str(r[14] or 0).replace("'", "").strip())
        except Exception: mov_emp = 0.0
        try: a_liquidar = float(str(r[15] or 0).replace("'", "").strip())
        except Exception: a_liquidar = 0.0
        try: liquidado_pago = float(str(r[21] or 0).replace("'", "").strip())
        except Exception: liquidado_pago = 0.0

        # Identificação da Liquidação na Célula Contábil do SIAFI
        nd_elem = nd[:4] + "00" if len(nd) >= 6 else nd
        c_liq_cand = (cel_liq_db.get((ug_cur, pi, nd)) or 
                      cel_elem_liq_db.get((ug_cur, pi, nd_elem)) or 
                      pi_liq_db.get((ug_cur, pi)) or 
                      global_pi_liq_db.get(pi) or {})
        
        cel_liq = c_liq_cand.get("liq", 0.0)
        cel_pag = c_liq_cand.get("pag", 0.0)
        cel_emp = c_liq_cand.get("emp", 0.0)
        pct_liq = (cel_liq / cel_emp * 100.0) if cel_emp > 0 else 0.0

        # Liquidação atribuída à NE
        ne_liq = 0.0
        if mov_emp > a_liquidar and mov_emp > 0:
            ne_liq = round(mov_emp - a_liquidar, 2)
        elif cel_liq > 0 and val > 0:
            ne_liq = min(val, cel_liq)

        ne_num = ne_full[-12:] if len(ne_full) >= 12 else ne_full

        # ================= PROVA DA VERDADE =================
        m_nc = nc_pat.search(desc)
        nc_citada = m_nc.group(1).upper() if m_nc else None

        status_prova = "VINCULO_CELULA"
        status_slug = "info"
        nc_efetiva = ""
        motivo_divergencia = ""

        if nc_citada:
            if nc_citada not in cred_audit_db:
                ano = nc_citada[:4]
                if ano < "2026":
                    status_prova = "Exercício Anterior (" + ano + ")"
                    status_slug = "muted"
                    estat_prova["EXERCICIO_ANTERIOR"] += 1
                else:
                    status_prova = "NC Não Encontrada no Extrato"
                    status_slug = "muted"
                    estat_prova["NC_NAO_ENCONTRADA"] += 1
                nc_efetiva = nc_citada
                motivo_divergencia = f"NC {nc_citada} citada no texto não consta no arquivo do exercício de 2026."
            else:
                recs = cred_audit_db[nc_citada]
                mesma_ug = [x for x in recs if x["ug"] == ug_cur]
                if not mesma_ug:
                    status_prova = "Divergência de UG"
                    status_slug = "danger"
                    estat_prova["DIVERGENCIA_UG"] += 1
                    nc_efetiva = nc_citada
                    ug_dest = recs[0]["ug"]
                    motivo_divergencia = f"NE é da UG {ug_cur}, mas a NC {nc_citada} citada foi descentralizada para a UG {ug_dest}."
                else:
                    mesmo_pi = [x for x in mesma_ug if x["pi"] == pi]
                    if not mesmo_pi:
                        status_prova = "Divergência de PI"
                        status_slug = "warn"
                        estat_prova["DIVERGENCIA_PI"] += 1
                        nc_efetiva = nc_citada
                        pi_dest = mesma_ug[0]["pi"]
                        motivo_divergencia = f"NE pertence ao PI {pi}, mas a NC {nc_citada} citada é do PI {pi_dest}."
                    else:
                        mesma_nd = [x for x in mesmo_pi if x["nd"] == nd or x["nd"][:4] == nd[:4] or x["nd"] == "339000" or x["nd"] == "449000"]
                        if mesma_nd:
                            status_prova = "Auditado & Confirmado (UG+PI+ND)"
                            status_slug = "ok"
                            estat_prova["CONFIRMADO"] += 1
                            nc_efetiva = nc_citada
                        else:
                            status_prova = "Divergência de ND"
                            status_slug = "warn"
                            estat_prova["DIVERGENCIA_ND"] += 1
                            nc_efetiva = nc_citada
                            motivo_divergencia = f"NE é da ND {nd}, mas a NC {nc_citada} citada é da ND {mesmo_pi[0]['nd']}."
        else:
            # Vínculo por Célula SIAFI
            nd_elem = nd[:4] + "00" if len(nd) >= 6 else nd
            cands = cel_to_ncs.get((ug_cur, pi, nd), []) or cel_to_ncs.get((ug_cur, pi, nd_elem), [])
            if cands:
                nc_efetiva = cands[0]
                status_prova = "Célula Orçamentária SIAFI"
                status_slug = "info"
                estat_prova["CELULA_SIAFI"] += 1
            else:
                nc_efetiva = ""
                status_prova = "Crédito Avulso"
                status_slug = "muted"

        if status_slug in ("danger", "warn", "muted") and nc_citada:
            inconsistencias_citacao.append({
                "ne": ne_num,
                "ne_full": ne_full,
                "dia": dia,
                "ug": ug_cur,
                "fav": fav_nome,
                "doc": fav_doc,
                "val": round(val, 2),
                "val_fmt": _fmt_brl(val),
                "pi_ne": pi,
                "nd_ne": nd,
                "nc_citada": nc_citada,
                "status_prova": status_prova,
                "status_slug": status_slug,
                "motivo": motivo_divergencia,
                "desc": desc
            })

        # Processo / Pregão
        m_proc = proc_pat.search(desc)
        proc_str = m_proc.group(0).upper().strip() if m_proc else ""

        ne_item = {
            "ne": ne_num,
            "ne_full": ne_full,
            "ug": ug_cur,
            "ug_nome": ug_nome_cur,
            "dia": dia,
            "doc": fav_doc,
            "fav": fav_nome,
            "val": round(val, 2),
            "nd": nd,
            "nd_desc": nd_desc,
            "pi": pi,
            "pi_desc": pi_desc,
            "nc": nc_efetiva or (nc_citada or ""),
            "nc_citada": nc_citada or "",
            "prova": status_prova,
            "prova_slug": status_slug,
            "motivo": motivo_divergencia,
            "proc": proc_str,
            "desc": desc,
            "a_liquidar": round(a_liquidar, 2),
            "pago": round(cel_pag if cel_pag > 0 else (ne_liq if ne_liq > 0 else liquidado_pago), 2),
            "cel_liq": round(cel_liq, 2),
            "cel_pag": round(cel_pag, 2),
            "cel_emp": round(cel_emp, 2),
            "pct_liq": round(pct_liq, 1),
            "ne_liq": round(ne_liq, 2)
        }
        nes_list.append(ne_item)
        tot_val += val

        # Indexa por NC para drill-down e relacionamentos
        ne_summary = {
            "ne": ne_num,
            "ne_full": ne_full,
            "dia": dia,
            "ug": ug_cur,
            "fav": fav_nome,
            "doc": fav_doc,
            "val": round(val, 2),
            "nd": nd,
            "pi": pi,
            "proc": proc_str,
            "prova": status_prova,
            "prova_slug": status_slug,
            "motivo": motivo_divergencia,
            "desc": desc,
            "cel_liq": round(cel_liq, 2),
            "cel_pag": round(cel_pag, 2),
            "ne_liq": round(ne_liq, 2)
        }
        if nc_efetiva:
            _add_to_nc_to_nes(nc_efetiva, ne_summary)
            for full_cand in cred_audit_db.get(nc_efetiva, []):
                _add_to_nc_to_nes(full_cand.get("nc_full"), ne_summary)
        if nc_citada and nc_citada != nc_efetiva:
            _add_to_nc_to_nes(nc_citada, ne_summary)
            for full_cand in cred_audit_db.get(nc_citada, []):
                _add_to_nc_to_nes(full_cand.get("nc_full"), ne_summary)

        # Agrega fornecedores
        if fav_nome and fav_nome != "NAO SE APLICA":
            top_forn[fav_nome] = top_forn.get(fav_nome, 0.0) + val

        # Agrega processos
        if proc_str:
            top_procs[proc_str] = top_procs.get(proc_str, 0.0) + val

    # Top fornecedores
    sorted_forn = sorted(top_forn.items(), key=lambda x: x[1], reverse=True)[:30]
    top_forn_list = [{"nome": f, "val": round(v, 2), "fmt": _fmt_brl(v)} for f, v in sorted_forn]

    # Top processos
    sorted_procs = sorted(top_procs.items(), key=lambda x: x[1], reverse=True)[:25]
    top_procs_list = [{"proc": p, "val": round(v, 2), "fmt": _fmt_brl(v)} for p, v in sorted_procs]

    return {
        "total_val": round(tot_val, 2),
        "total_qtd": len(nes_list),
        "qtd_fav": len(top_forn),
        "tot_liq_global": round(tot_liq_global, 2),
        "tot_pag_global": round(tot_pag_global, 2),
        "estat_prova": estat_prova,
        "inconsistencias": inconsistencias_citacao,
        "top_fornecedores": top_forn_list,
        "top_processos": top_procs_list,
        "nc_to_nes": nc_to_nes,
        "celulas_liq": {f"{k[0]}_{k[1]}_{k[2]}": v for k, v in cel_liq_db.items()},
        "pi_liq": global_pi_liq_db,
        "nes": nes_list
    }


def secao_rastreabilidade_empenhos(res, emp_data, histdata, data_str, periodo):
    """Gera o HTML da nova aba 'Crédito ➔ Empenho' 100% integrado ao Design System do Dashboard."""
    posicao = periodo if periodo else (data_str[8:10] + "/" + data_str[5:7] + "/" + data_str[0:4])
    
    tot_cred_disp = sum(res[c]["cred"] for c in res)
    tot_prov_rec = sum(res[c]["prov"] for c in res)
    tot_emp = emp_data["total_val"]
    qtd_nes = emp_data["total_qtd"]
    qtd_fav = emp_data["qtd_fav"]
    tot_liq = emp_data.get("tot_liq_global", 0.0)
    tot_pag = emp_data.get("tot_pag_global", 0.0)
    pct_liq_total = (tot_liq / tot_emp * 100.0) if tot_emp > 0 else 0.0
    
    taxa_exec = (tot_emp / tot_prov_rec * 100.0) if tot_prov_rec > 0 else 0.0

    ep = emp_data.get("estat_prova", {})
    qtd_conf = ep.get("CONFIRMADO", 0) + ep.get("CELULA_SIAFI", 0)
    pct_conf = (qtd_conf / qtd_nes * 100.0) if qtd_nes > 0 else 0.0
    qtd_alertas = len(emp_data.get("inconsistencias", []))

    return f"""
<section class="unidade unidade-rastreio" data-key="RASTREIO" style="display:none" id="secao-RASTREIO">
  <div class="ranking-header-card" style="border-top-color:#059669;margin-top:24px;">
    <div class="rh-tag" style="color:#059669;">SISTEMA INTEGRADO DE CONTROLE ORÇAMENTÁRIO & FINANCEIRO · SICOF</div>
    <h2 class="rh-title">🔗 Rastreabilidade Integral: Crédito ➔ Empenho Emitido</h2>
    <p class="rh-desc">Controle orçamentário completo da Base de Apoio Logístico do Exército e OMDS subordinadas — Da provisão recebida (Nota de Crédito) à emissão do empenho (NE), com conciliação contábil por Célula SIAFI e Prova Real de integridade entre dotações, fornecedores e processos licitatórios.</p>
  </div>

  <section class="sec" style="margin-top:24px;">
    <div class="eyebrow">Indicadores Globais de Execução e Conformidade</div>
    <div class="kpis">
      <div class="kpi" style="border-left-color:#2563EB;">
        <span class="kpi-lbl">TOTAL EMPENHADO (NEs)</span>
        <div class="kpi-val num" style="font-size:1.625rem;font-weight:800;color:var(--ink);">{_fmt_brl(tot_emp)}</div>
        <div class="kpi-sub"><b>{qtd_nes:,} NEs</b> ({taxa_exec:.1f}% do crédito descentralizado)</div>
      </div>
      <div class="kpi" style="border-left-color:#F59E0B;">
        <span class="kpi-lbl">DESPESAS LIQUIDADAS NO SIAFI</span>
        <div class="kpi-val num" style="font-size:1.625rem;font-weight:800;color:var(--ink);">{_fmt_brl(tot_liq)}</div>
        <div class="kpi-sub"><b>{_fmt_brl(tot_pag)} pagos</b> ({pct_liq_total:.1f}% do empenhado)</div>
      </div>
      <div class="kpi" style="border-left-color:var(--success);">
        <span class="kpi-lbl">CRÉDITO DISPONÍVEL LÍQUIDO</span>
        <div class="kpi-val num" style="font-size:1.625rem;font-weight:800;color:var(--success-strong);">{_fmt_brl(tot_cred_disp)}</div>
        <div class="kpi-sub">Saldo livre em caixa para novas contratações</div>
      </div>
      <div class="kpi" style="border-left-color:#059669;">
        <span class="kpi-lbl">PROVA REAL: INTEGRIDADE SIAFI</span>
        <div class="kpi-val num" style="font-size:1.625rem;font-weight:800;color:#059669;">{pct_conf:.1f}%</div>
        <div class="kpi-sub"><b>{qtd_conf:,} NEs</b> com lastro contábil comprovado</div>
      </div>
      <div class="kpi" style="border-left-color:#8B5CF6;">
        <span class="kpi-lbl">FORNECEDORES & CREDORES</span>
        <div class="kpi-val num" style="font-size:1.625rem;font-weight:800;color:var(--ink);">{qtd_fav:,}</div>
        <div class="kpi-sub">Pessoas jurídicas e supridos no exercício</div>
      </div>
    </div>
  </section>

  <!-- Segmented Control de Sub-abas -->
  <div class="toptabs" role="tablist" style="margin:24px 0 16px;">
    <button type="button" class="toptab on" id="r-subtab-btn-nc" onclick="bcmsRastreioSubAba('nc', this)" role="tab" aria-selected="true">
      🏢 Visão por Nota de Crédito (Top-Down)
    </button>
    <button type="button" class="toptab" id="r-subtab-btn-ne" onclick="bcmsRastreioSubAba('ne', this)" role="tab" aria-selected="false">
      📄 Relação de Empenhos ({qtd_nes:,} NEs)
    </button>
    <button type="button" class="toptab" id="r-subtab-btn-forn" onclick="bcmsRastreioSubAba('forn', this)" role="tab" aria-selected="false">
      🏢 Fornecedores & Pregões
    </button>
    <button type="button" class="toptab" id="r-subtab-btn-prova" onclick="bcmsRastreioSubAba('prova', this)" role="tab" aria-selected="false">
      🛡️ Prova Real & Auditoria ({qtd_alertas} Alertas)
    </button>
    <button type="button" class="toptab" id="r-subtab-btn-risco" onclick="bcmsRastreioSubAba('risco', this)" role="tab" aria-selected="false">
      ⚠️ Créditos Parados (Risco)
    </button>
  </div>

  <!-- ================= SUB-ABA 1: TOP-DOWN (CRÉDITO > EMPENHO) ================= -->
  <div class="r-subtab-pane" id="pane-rastreio-nc">
    <div class="tbl-tools">
      <label class="visually-hidden" for="rBuscaNC">Buscar Créditos</label>
      <input type="search" id="rBuscaNC" class="tbl-search" placeholder="Buscar por NC, Fornecedor, Objeto, PI, ND ou Concedente…" oninput="bcmsFiltraRastreioNC()">
      <button type="button" class="btn-excel btn-excel-lg" onclick="bcmsExportRastreioExcel('nc')" title="Baixar visão consolidada de Crédito e Empenho formatada para Excel"><span class="btn-excel-ic">📊</span> Exportar Crédito ➔ Empenho (Excel)</button>
      <span class="tbl-count" id="rContadorNC" aria-live="polite">Carregando dados…</span>
    </div>
    <div class="tbl-filtros" role="group" aria-label="Filtros de Crédito e Empenho">
      <span class="flt-lbl">Filtrar:</span>
      <select class="flt" id="rSelOM" aria-label="Filtrar por Organização Militar" onchange="bcmsFiltraRastreioNC()">
        <option value="">Unidade: todas as OMDS</option>
        <option value="160329">BCMS · OGU</option>
        <option value="167329">BCMS · FEx</option>
        <option value="160238">Ba Ap Log · OGU</option>
        <option value="167238">Ba Ap Log · FEx</option>
        <option value="160246">D C Mun · OGU</option>
        <option value="167246">D C Mun · FEx</option>
        <option value="160304">BMSA · OGU</option>
        <option value="167304">BMSA · FEx</option>
        <option value="160307">1º D Sup · OGU</option>
        <option value="167307">1º D Sup · FEx</option>
        <option value="160321">ECT · OGU</option>
        <option value="167321">ECT · FEx</option>
      </select>
      <select class="flt" id="rSelNDNC" aria-label="Filtrar por Natureza de Despesa" onchange="bcmsFiltraRastreioNC()">
        <option value="">ND: todas as despesas</option>
        <option value="339030">339030 · Material de Consumo</option>
        <option value="339039">339039 · Outros Serviços Terceiros PJ</option>
        <option value="449052">449052 · Equipamentos e Mat. Permanente</option>
        <option value="339015">339015 · Diárias - Militar</option>
        <option value="339033">339033 · Passagens e Locomoção</option>
        <option value="449051">449051 · Obras e Instalações</option>
      </select>
      <select class="flt" id="rSelFonteNC" aria-label="Filtrar por Fonte de Recursos" onchange="bcmsFiltraRastreioNC()">
        <option value="">Fonte: todas (OGU &amp; FEx)</option>
        <option value="160">Fonte 160 · OGU (Orçamento Geral)</option>
        <option value="167">Fonte 167 · FEx (Fundo do Exército)</option>
      </select>
      <select class="flt" id="rSelQueima" aria-label="Filtrar por ritmo de queima" onchange="bcmsFiltraRastreioNC()">
        <option value="">Ritmo de Queima: todos</option>
        <option value="alta">🟢 Queima Alta (&gt; 80% empenhado)</option>
        <option value="media">🟡 Queima Média (30% a 80%)</option>
        <option value="baixa">🔴 Queima Baixa (&lt; 30% empenhado)</option>
        <option value="zero">⚪ Sem Empenhos (100% disponível)</option>
      </select>
      <button type="button" class="flt-limpa" onclick="bcmsLimpaFiltrosRastreioNC()" title="Limpar todos os filtros">✕ Limpar</button>
    </div>

    <div class="tbl-scroll" id="scroll-rastreio-nc">
      <table class="det det-compact" id="tabRastreioNC">
        <thead>
          <tr>
            <th style="width:48px;text-align:center;">NEs</th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNC('nc')" title="Ordenar por Nota de Crédito">Nota de Crédito <span class="sort" id="sort-nc-nc"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNC('dia')" title="Ordenar por Data">Recebido em <span class="sort" id="sort-nc-dia"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNC('emit_nome')" title="Ordenar por Órgão Concedente">Órgão Concedente <span class="sort" id="sort-nc-emit"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNC('om_sigla')" title="Ordenar por OM Favorecida">Unidade (OMDS) <span class="sort" id="sort-nc-uasg"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNC('pi')" title="Ordenar por Plano Interno">Plano Interno <span class="sort" id="sort-nc-pi"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNC('nd')" title="Ordenar por Natureza de Despesa">ND <span class="sort" id="sort-nc-nd"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNC('prov')" class="num" title="Ordenar por Provisão Recebida">Crédito Recebido <span class="sort" id="sort-nc-prov">▼</span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNC('emp')" class="num" title="Ordenar por Total Empenhado">Total Empenhado <span class="sort" id="sort-nc-emp"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNC('liq')" class="num" title="Ordenar por Despesas Liquidadas">Liquidado (SIAFI) <span class="sort" id="sort-nc-liq"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNC('cred')" class="num anchor" title="Ordenar por Saldo Disponível">Saldo Disponível <span class="sort" id="sort-nc-cred"></span></th>
            <th style="width:130px;text-align:center;">Execução</th>
            <th style="width:84px;text-align:center;">Ações</th>
          </tr>
        </thead>
        <tbody id="corpoRastreioNC">
        </tbody>
      </table>
    </div>
    <div class="tbl-tools" style="margin-top:10px;justify-content:space-between;" id="paginacaoRastreioNC"></div>
  </div>

  <!-- ================= SUB-ABA 2: BOTTOM-UP (RELAÇÃO DE EMPENHOS) ================= -->
  <div class="r-subtab-pane" id="pane-rastreio-ne" style="display:none">
    <div class="tbl-tools">
      <label class="visually-hidden" for="rBuscaNE">Buscar Empenhos</label>
      <input type="search" id="rBuscaNE" class="tbl-search" placeholder="Buscar por Razão Social, CNPJ, Número da NE, Pregão ou NC…" oninput="bcmsFiltraRastreioNE()">
      <button type="button" class="btn-excel btn-excel-lg" onclick="bcmsExportRastreioExcel('ne')" title="Baixar relação de Notas de Empenho formatada para Excel"><span class="btn-excel-ic">📊</span> Exportar Notas de Empenho (Excel)</button>
      <span class="tbl-count" id="rContadorNE" aria-live="polite">Carregando empenhos…</span>
    </div>
    <div class="tbl-filtros" role="group" aria-label="Filtros de Empenhos">
      <span class="flt-lbl">Filtrar:</span>
      <select class="flt" id="rSelOMNE" aria-label="Filtrar por Unidade Gestora" onchange="bcmsFiltraRastreioNE()">
        <option value="">Unidade: todas as UGs</option>
        <option value="160329">160329 · BCMS (OGU)</option>
        <option value="167329">167329 · BCMS (FEx)</option>
        <option value="160238">160238 · Ba Ap Log (OGU)</option>
        <option value="167238">167238 · Ba Ap Log (FEx)</option>
        <option value="160246">160246 · D C Mun (OGU)</option>
        <option value="167246">167246 · D C Mun (FEx)</option>
        <option value="160304">160304 · BMSA (OGU)</option>
        <option value="167304">167304 · BMSA (FEx)</option>
        <option value="160307">160307 · 1º D Sup (OGU)</option>
        <option value="167307">167307 · 1º D Sup (FEx)</option>
        <option value="160321">160321 · ECT (OGU)</option>
        <option value="167321">167321 · ECT (FEx)</option>
      </select>
      <select class="flt" id="rSelNDNE" aria-label="Filtrar por Natureza de Despesa" onchange="bcmsFiltraRastreioNE()">
        <option value="">ND: todas as despesas</option>
        <option value="339030">339030 · Material de Consumo</option>
        <option value="339039">339039 · Outros Serviços Terceiros PJ</option>
        <option value="449052">449052 · Equipamentos e Mat. Permanente</option>
        <option value="339015">339015 · Diárias - Militar</option>
        <option value="339093">339093 · Indenizações e Restituições</option>
        <option value="339033">339033 · Passagens e Locomoção</option>
        <option value="449051">449051 · Obras e Instalações</option>
      </select>
      <select class="flt" id="rSelProvaNE" aria-label="Filtrar por Selo da Prova Real" onchange="bcmsFiltraRastreioNE()">
        <option value="">Selo Prova Real: todos</option>
        <option value="ok">🟢 Auditado &amp; Confirmado</option>
        <option value="info">🔵 Célula SIAFI (Sem Citação)</option>
        <option value="warn">🟡 Divergência de PI / ND</option>
        <option value="danger">🔴 Divergência de UG</option>
        <option value="muted">⚪ Exercício Anterior</option>
      </select>
      <button type="button" class="flt-limpa" onclick="bcmsLimpaFiltrosRastreioNE()" title="Limpar todos os filtros">✕ Limpar</button>
    </div>

    <div class="tbl-scroll" id="scroll-rastreio-ne">
      <table class="det det-compact" id="tabRastreioNE">
        <thead>
          <tr>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNE('ne')" title="Ordenar por Nota de Empenho">Nota de Empenho <span class="sort" id="sort-ne-ne"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNE('dia')" title="Ordenar por Data de Emissão">Emissão <span class="sort" id="sort-ne-dia">▼</span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNE('ug')" title="Ordenar por Unidade">UG <span class="sort" id="sort-ne-ug"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNE('fav')" title="Ordenar por Favorecido">Favorecido / Fornecedor <span class="sort" id="sort-ne-fav"></span></th>
            <th>CNPJ / CPF</th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNE('proc')" title="Ordenar por Processo / Pregão">Processo / Pregão <span class="sort" id="sort-ne-proc"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNE('val')" class="num anchor" title="Ordenar por Valor Empenhado">Valor da NE <span class="sort" id="sort-ne-val"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNE('cel_liq')" class="num" title="Ordenar por Liquidação da Célula">Liquidado (SIAFI) <span class="sort" id="sort-ne-liq"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNE('nd')" title="Ordenar por Natureza de Despesa">ND <span class="sort" id="sort-ne-nd"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNE('pi')" title="Ordenar por Plano Interno">Plano Interno <span class="sort" id="sort-ne-pi"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNE('nc')" title="Ordenar por NC Citada">NC de Origem <span class="sort" id="sort-ne-nc"></span></th>
            <th style="text-align:center;">Selo Prova Real</th>
            <th style="text-align:center;">Ações</th>
          </tr>
        </thead>
        <tbody id="corpoRastreioNE">
        </tbody>
      </table>
    </div>
    <div class="tbl-tools" style="margin-top:10px;justify-content:space-between;" id="paginacaoRastreioNE"></div>
  </div>

  <!-- ================= SUB-ABA 3: RAIO-X FORNECEDORES & PREGÕES ================= -->
  <div class="r-subtab-pane" id="pane-rastreio-forn" style="display:none">
    <div class="tbl-tools" style="margin-top:10px;">
      <label class="visually-hidden" for="rBuscaForn">Buscar Fornecedor ou Processo</label>
      <input type="search" id="rBuscaForn" class="tbl-search" placeholder="Buscar por Razão Social, CNPJ ou Processo / Pregão…" oninput="bcmsFiltraForn()">
      <span class="tbl-count" id="rContadorForn" aria-live="polite">Carregando credores…</span>
    </div>
    <div class="tbl-filtros" role="group" aria-label="Filtros de Fornecedores e Pregões">
      <span class="flt-lbl">Filtrar:</span>
      <select class="flt" id="rSelOMForn" aria-label="Filtrar por Unidade Gestora" onchange="bcmsFiltraForn()">
        <option value="">Unidade: todas as UGs</option>
        <option value="160329">160329 · BCMS (OGU)</option>
        <option value="167329">167329 · BCMS (FEx)</option>
        <option value="160238">160238 · Ba Ap Log (OGU)</option>
        <option value="167238">167238 · Ba Ap Log (FEx)</option>
        <option value="160246">160246 · D C Mun (OGU)</option>
        <option value="167246">167246 · D C Mun (FEx)</option>
        <option value="160304">160304 · BMSA (OGU)</option>
        <option value="167304">167304 · BMSA (FEx)</option>
        <option value="160307">160307 · 1º D Sup (OGU)</option>
        <option value="167307">167307 · 1º D Sup (FEx)</option>
        <option value="160321">160321 · ECT (OGU)</option>
        <option value="167321">167321 · ECT (FEx)</option>
      </select>
      <select class="flt" id="rSelNDForn" aria-label="Filtrar por Natureza de Despesa" onchange="bcmsFiltraForn()">
        <option value="">ND: todas as despesas</option>
        <option value="339030">339030 · Material de Consumo</option>
        <option value="339039">339039 · Outros Serviços Terceiros PJ</option>
        <option value="449052">449052 · Equipamentos e Mat. Permanente</option>
        <option value="339015">339015 · Diárias - Militar</option>
        <option value="339033">339033 · Passagens e Locomoção</option>
        <option value="449051">449051 · Obras e Instalações</option>
      </select>
      <button type="button" class="flt-limpa" onclick="bcmsLimpaFiltrosForn()" title="Limpar todos os filtros">✕ Limpar</button>
    </div>

    <div class="grid2" style="margin-top:14px;">
      <div class="card" style="padding:24px;">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
          <h3 style="font-family:var(--serif);font-size:1.15rem;font-weight:700;color:var(--ink);margin:0;">🏢 Maiores Fornecedores por Volume Empenhado</h3>
          <span class="pill-fonte" style="font-weight:700;">Top Credores (Clique para detalhar)</span>
        </div>
        <p style="font-size:0.8125rem;color:var(--ink-muted);margin:0 0 18px 0;line-height:1.5;">Concentração de despesa empenhada no âmbito da Base de Apoio Logístico do Exército em 2026. Clique em qualquer fornecedor para ver todas as NEs.</p>
        <div class="ranking-lista" id="listaTopFornecedores"></div>
      </div>

      <div class="card" style="padding:24px;">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
          <h3 style="font-family:var(--serif);font-size:1.15rem;font-weight:700;color:var(--ink);margin:0;">📜 Principais Pregões & Processos Licitatórios</h3>
          <span class="pill-fonte" style="font-weight:700;">Demandas SISLOG (Clique para detalhar)</span>
        </div>
        <p style="font-size:0.8125rem;color:var(--ink-muted);margin:0 0 18px 0;line-height:1.5;">Processos administrativos e pregões eletrônicos mais referenciados nas emissões de empenho. Clique para abrir a lista completa de empenhos.</p>
        <div class="ranking-lista" id="listaTopProcessos"></div>
      </div>
    </div>
  </div>

  <!-- ================= SUB-ABA 4: PROVA REAL & AUDITORIA DE CITAÇÃO ================= -->
  <div class="r-subtab-pane" id="pane-rastreio-prova" style="display:none">
    <div class="card" style="border-left:5px solid #2563EB;background:var(--hero-soft);margin-bottom:20px;padding:20px 24px;">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:6px;">
        <span style="font-size:1.25rem;">🛡️</span>
        <h4 style="margin:0;font-size:1.05rem;font-weight:700;color:var(--ink);">Motor de Auditoria: Prova Real de Dupla Chave (Crédito SIAFI vs Citação Textual)</h4>
      </div>
      <p style="margin:0;font-size:0.875rem;color:var(--ink);line-height:1.6;">No SIAFI, a execução financeira se dá estritamente por <b>Célula Orçamentária (UG + PI + ND + Fonte)</b>. O campo "Observação" da NE é texto livre digitado manualmente pelo operador e frequentemente contém <b>erros de digitação ou cópia de textos de processos anteriores</b>. O algoritmo cruza a citação com a dotação real de destino para validar o lastro contábil autêntico.</p>
    </div>

    <div class="kpis" style="margin-bottom:24px;">
      <div class="kpi" style="border-left-color:var(--success);">
        <span class="kpi-lbl">VÍNCULOS CONFIRMADOS</span>
        <div class="kpi-val num" style="font-size:1.625rem;font-weight:800;color:var(--success-strong);">{qtd_conf:,} NEs</div>
        <div class="kpi-sub"><b>{pct_conf:.1f}%</b> com lastro contábil comprovado</div>
      </div>
      <div class="kpi" style="border-left-color:var(--warning);">
        <span class="kpi-lbl">DIVERGÊNCIA DE PI (PLANO INTERNO)</span>
        <div class="kpi-val num" style="font-size:1.625rem;font-weight:800;color:var(--warning-ink);">{ep.get('DIVERGENCIA_PI', 0)} NEs</div>
        <div class="kpi-sub">Operador colou texto de outro PI</div>
      </div>
      <div class="kpi" style="border-left-color:var(--danger);">
        <span class="kpi-lbl">DIVERGÊNCIA DE UG (UNIDADE)</span>
        <div class="kpi-val num" style="font-size:1.625rem;font-weight:800;color:var(--danger);">{ep.get('DIVERGENCIA_UG', 0)} NEs</div>
        <div class="kpi-sub">Operador citou NC de outra OM</div>
      </div>
      <div class="kpi" style="border-left-color:#6366F1;">
        <span class="kpi-lbl">EXERCÍCIO ANTERIOR / NÃO LOCALIZADA</span>
        <div class="kpi-val num" style="font-size:1.625rem;font-weight:800;color:var(--ink);">{ep.get('EXERCICIO_ANTERIOR', 0) + ep.get('NC_NAO_ENCONTRADA', 0)} NEs</div>
        <div class="kpi-sub">Citação residual 2025 ou erro de dígito</div>
      </div>
    </div>

    <div class="card" style="padding:22px;margin-bottom:24px;">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;flex-wrap:wrap;gap:8px;">
        <h3 style="font-family:var(--serif);font-size:1.15rem;font-weight:700;color:var(--ink);margin:0;">📋 Registro de Inconsistências Detectadas na Observação</h3>
        <span class="pill-status status-canc" style="font-weight:700;">Auditoria Preventiva</span>
      </div>
      <p style="font-size:0.8125rem;color:var(--ink-muted);margin:0 0 16px 0;">Relação de Notas de Empenho onde a NC citada no texto diverge da célula orçamentária autêntica debitada no SIAFI. Clique em qualquer linha para auditar.</p>

      <div class="tbl-tools" style="margin-bottom:12px;">
        <label class="visually-hidden" for="rBuscaProva">Buscar Inconsistências</label>
        <input type="search" id="rBuscaProva" class="tbl-search" placeholder="Buscar por NE, Favorecido, NC citada ou motivo…" oninput="bcmsFiltraProva()">
        <span class="tbl-count" id="rContadorProva" aria-live="polite">Carregando inconsistências…</span>
      </div>
      <div class="tbl-filtros" style="margin-bottom:16px;" role="group" aria-label="Filtros de Inconsistências">
        <span class="flt-lbl">Filtrar:</span>
        <select class="flt" id="rSelOMProva" aria-label="Filtrar por UG" onchange="bcmsFiltraProva()">
          <option value="">Unidade: todas as UGs</option>
          <option value="160329">160329 · BCMS</option>
          <option value="160238">160238 · Ba Ap Log</option>
          <option value="160246">160246 · D C Mun</option>
          <option value="160304">160304 · BMSA</option>
          <option value="160307">160307 · 1º D Sup</option>
          <option value="160321">160321 · ECT</option>
        </select>
        <select class="flt" id="rSelStatusProva" aria-label="Filtrar por Tipo de Inconsistência" onchange="bcmsFiltraProva()">
          <option value="">Tipo: todas as divergências</option>
          <option value="warn">Divergência de PI / ND</option>
          <option value="danger">Divergência de UG</option>
          <option value="muted">Exercício Anterior / Não Localizada</option>
        </select>
        <button type="button" class="flt-limpa" onclick="bcmsLimpaFiltrosProva()" title="Limpar filtros">✕ Limpar</button>
      </div>
      
      <div class="tbl-scroll">
        <table class="det det-compact" id="tabInconsistencias">
          <thead>
            <tr>
              <th>Número da NE</th>
              <th>Emissão</th>
              <th>UG</th>
              <th>Favorecido / Fornecedor</th>
              <th class="num">Valor da NE</th>
              <th>PI Real da NE</th>
              <th>ND da NE</th>
              <th>NC Citada no Texto</th>
              <th>Diagnóstico da Inconsistência</th>
              <th style="text-align:center;">Ação</th>
            </tr>
          </thead>
          <tbody id="corpoInconsistencias">
          </tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- ================= SUB-ABA 5: MONITOR DE CRÉDITOS PARADOS (RISCO) ================= -->
  <div class="r-subtab-pane" id="pane-rastreio-risco" style="display:none">
    <div class="card" style="border-left:5px solid var(--warning);background:var(--warning-bg);margin-bottom:20px;padding:20px 24px;">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:6px;">
        <span style="font-size:1.25rem;">⚠️</span>
        <h4 style="margin:0;font-size:1.05rem;font-weight:700;color:var(--warning-ink);">Painel Preventivo de Gestão de Riscos Orçamentários (Fim de Exercício)</h4>
      </div>
      <p style="margin:0;font-size:0.875rem;color:var(--warning-ink);line-height:1.6;">Créditos orçamentários descentralizados que possuem <b>saldo disponível considerável (&gt; R$ 15.000)</b> e encontram-se <b>sem emissão de novos empenhos há mais de 45 dias</b>. O monitoramento contínuo evita o risco de estorno ou recolhimento pelo escalão superior no encerramento do exercício financeiro.</p>
    </div>

    <div class="tbl-tools" style="margin-bottom:12px;">
      <label class="visually-hidden" for="rBuscaRisco">Buscar Créditos em Risco</label>
      <input type="search" id="rBuscaRisco" class="tbl-search" placeholder="Buscar por NC, Favorecida, Órgão Concedente, PI ou ND…" oninput="bcmsFiltraRisco()">
      <span class="tbl-count" id="rContadorRisco" aria-live="polite">Carregando créditos parados…</span>
    </div>
    <div class="tbl-filtros" style="margin-bottom:16px;" role="group" aria-label="Filtros de Créditos Parados">
      <span class="flt-lbl">Filtrar:</span>
      <select class="flt" id="rSelOMRisco" aria-label="Filtrar por OMDS" onchange="bcmsFiltraRisco()">
        <option value="">Unidade: todas as OMDS</option>
        <option value="160329">BCMS · OGU</option>
        <option value="167329">BCMS · FEx</option>
        <option value="160238">Ba Ap Log · OGU</option>
        <option value="167238">Ba Ap Log · FEx</option>
        <option value="160246">D C Mun · OGU</option>
        <option value="167246">D C Mun · FEx</option>
        <option value="160304">BMSA · OGU</option>
        <option value="167304">BMSA · FEx</option>
        <option value="160307">1º D Sup · OGU</option>
        <option value="167307">1º D Sup · FEx</option>
        <option value="160321">ECT · OGU</option>
        <option value="167321">ECT · FEx</option>
      </select>
      <select class="flt" id="rSelFaixaRisco" aria-label="Filtrar por Faixa de Saldo" onchange="bcmsFiltraRisco()">
        <option value="">Saldo Parado: &gt; R$ 15.000</option>
        <option value="50k">&gt; R$ 50.000</option>
        <option value="100k">&gt; R$ 100.000</option>
        <option value="500k">&gt; R$ 500.000</option>
      </select>
      <button type="button" class="flt-limpa" onclick="bcmsLimpaFiltrosRisco()" title="Limpar filtros">✕ Limpar</button>
    </div>

    <div class="tbl-scroll">
      <table class="det det-compact" id="tabRastreioRisco">
        <thead>
          <tr>
            <th>Nota de Crédito</th>
            <th>Data de Aporte</th>
            <th>Idade</th>
            <th>Órgão Concedente</th>
            <th>OMDS Favorecida</th>
            <th>Plano Interno</th>
            <th>Natureza Despesa</th>
            <th class="num">Crédito Recebido</th>
            <th class="num anchor" style="color:var(--warning-ink);">Saldo Disponível Parado</th>
            <th style="text-align:center;">Ação Imediata</th>
          </tr>
        </thead>
        <tbody id="corpoRastreioRisco">
        </tbody>
      </table>
    </div>
  </div>
</section>
"""

CSS_RASTREIO = r"""
/* ==================== ESTILOS DA ABA CRÉDITO > EMPENHO (SICOF) ==================== */
.unidade-rastreio .det th {
  user-select: none;
}
.omds-rastreio {
  background: var(--bg-subtle);
  border-color: #059669;
  color: #059669;
}
.omds-rastreio:hover {
  border-color: #059669;
  color: var(--ink);
}
.omds-rastreio.on {
  background: linear-gradient(135deg, #065F46 0%, #059669 100%) !important;
  color: #FFFFFF !important;
  border-color: #059669 !important;
  box-shadow: 0 4px 14px rgba(5, 150, 105, 0.4) !important;
}
.omds-rastreio.on span { color: #FFFFFF !important; }
.rastreio-icon { font-size: 1.125rem; }

.queima-bar-wrap {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  width: 100%;
}
.queima-track {
  flex: 1;
  height: 6px;
  background: var(--track, rgba(0,0,0,0.08));
  border-radius: 999px;
  overflow: hidden;
}
.queima-fill {
  height: 100%;
  border-radius: 999px;
  transition: width 0.3s ease;
}
.queima-fill.alta  { background: #10B981; }
.queima-fill.media { background: #F59E0B; }
.queima-fill.baixa { background: #EF4444; }
.queima-fill.zero  { background: transparent; }
.queima-lbl {
  font-size: 0.75rem;
  font-weight: 700;
  font-family: var(--mono);
  color: var(--ink-muted);
  min-width: 36px;
  text-align: right;
}

/* Acordeão de NEs Vinculadas */
.sub-ne-wrap {
  padding: 16px 20px;
  margin: 6px 10px;
  background: var(--bg-surface);
  border: 1px solid var(--border);
  border-left: 4px solid var(--accent, #059669);
  border-radius: 12px;
  box-shadow: var(--shadow-sm);
}
.sub-ne-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 12px;
  flex-wrap: wrap;
  gap: 8px;
}
.sub-ne-title {
  font-size: 0.875rem;
  font-weight: 700;
  color: var(--ink);
}
.sub-ne-meta {
  font-size: 0.75rem;
  font-weight: 600;
  color: var(--ink-muted);
}
.btn-exp-ne {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 4px;
  padding: 3px 8px;
  font-size: 0.75rem;
  font-weight: 700;
  font-family: var(--mono);
  border-radius: 6px;
  background: var(--bg-subtle);
  border: 1px solid var(--border);
  color: var(--ink);
  cursor: pointer;
  transition: all .15s ease;
}
.btn-exp-ne:hover, .btn-exp-ne.aberto {
  background: var(--accent, #059669);
  color: #FFFFFF;
  border-color: var(--accent, #059669);
}

/* Ranking de Fornecedores & Pregões */
.ranking-lista {
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.rank-item {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 10px 14px;
  border-radius: 10px;
  background: var(--bg-subtle);
  border: 1px solid var(--border);
  cursor: pointer;
  user-select: none;
  transition: transform .2s var(--ease-spring), border-color .2s ease, background-color .2s ease;
}
.rank-item:hover {
  transform: translateX(4px);
  border-color: #059669;
  background: var(--bg-surface);
}
.rank-info {
  display: flex;
  flex-direction: column;
  gap: 4px;
  max-width: 65%;
}
.rank-pos {
  font-size: 0.6875rem;
  font-weight: 800;
  color: var(--primary);
  letter-spacing: 0.04em;
}
.rank-nome {
  font-size: 0.8125rem;
  font-weight: 600;
  color: var(--ink);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.rank-val {
  font-size: 0.875rem;
  font-weight: 700;
  font-family: var(--mono);
  color: var(--ink);
}
.kpi-bar-wrap {
  width: 140px;
  height: 6px;
  background: var(--track, rgba(0,0,0,0.08));
  border-radius: 999px;
  overflow: hidden;
}
.kpi-bar-fill {
  height: 100%;
  background: linear-gradient(90deg, #10B981, #059669);
  border-radius: 999px;
  transition: width 0.4s ease;
}

/* Links interativos para drill-down */
.link-drill {
  color: var(--primary);
  text-decoration: none;
  font-weight: 700;
  cursor: pointer;
  transition: color .15s ease;
}
.link-drill:hover {
  color: #059669;
  text-decoration: underline;
}
"""

JS_RASTREIO = r"""
/* ==================== SCRIPTS DA ABA CRÉDITO > EMPENHO (SICOF) ==================== */
var RASTREIO_INITIALIZED = false;
var RASTREIO_NC_ITEMS = [];
var RASTREIO_NE_ITEMS = [];
var RASTREIO_NC_FILTRADOS = [];
var RASTREIO_NE_FILTRADOS = [];
var RASTREIO_NC_PAGE = 1;
var RASTREIO_NE_PAGE = 1;
var RASTREIO_SORT_NC = { col: 'prov', asc: false };
var RASTREIO_SORT_NE = { col: 'dia', asc: false };
var BCMS_MODAL_STACK = [];

function bcmsModalPush(fn){
  if(typeof fn === 'function'){
    BCMS_MODAL_STACK.push(fn);
  }
}

function bcmsModalBack(){
  if(BCMS_MODAL_STACK.length > 1){
    BCMS_MODAL_STACK.pop(); // remove o modal atual
    var prevFn = BCMS_MODAL_STACK.pop(); // retira o anterior para que ao executar seja reinserido
    if(prevFn) prevFn();
  } else {
    bcmsCelClose();
  }
}

function bcmsGetNesDaNC(ncStr){
  if(!EMPENHODATA || !EMPENHODATA.nc_to_nes || !ncStr) return [];
  var s = String(ncStr).trim();
  if(EMPENHODATA.nc_to_nes[s]) return EMPENHODATA.nc_to_nes[s];
  var m = s.match(/(202\dNC\d{6})/i);
  if(m && EMPENHODATA.nc_to_nes[m[1].toUpperCase()]){
    return EMPENHODATA.nc_to_nes[m[1].toUpperCase()];
  }
  var sShort = s.slice(-10).toUpperCase();
  if(EMPENHODATA.nc_to_nes[sShort]){
    return EMPENHODATA.nc_to_nes[sShort];
  }
  return [];
}

function bcmsInitRastreio(){
  if(RASTREIO_INITIALIZED) return;
  if(typeof EMPENHODATA === 'undefined' || !EMPENHODATA) return;

  // 1. Constrói RASTREIO_NC_ITEMS enriquecendo as NCs do HISTDATA com as NEs reais
  var allNcs = [];
  if(typeof HISTDATA !== 'undefined' && HISTDATA && HISTDATA.items){
    allNcs = HISTDATA.items;
  }
  
  RASTREIO_NC_ITEMS = [];
  for(var i = 0; i < allNcs.length; i++){
    var it = allNcs[i];
    var nesVinculadas = bcmsGetNesDaNC(it.nc);
    var totEmpNE = 0;
    for(var j = 0; j < nesVinculadas.length; j++){
      totEmpNE += (nesVinculadas[j].val || 0);
    }
    
    var provVal = it.prov || 0;
    // O saldo disponível oficial é rigorosamente o saldo registrado no SIAFI (it.cred)
    var credOficial = (typeof it.cred !== 'undefined' && it.cred !== null) ? it.cred : Math.max(0, provVal - totEmpNE);
    var consumido = Math.max(0, provVal - credOficial);
    var taxaQueima = provVal > 0 ? Math.min(100, Math.max(0, (consumido / provVal) * 100)) : 0;
    var faixa = 'zero';
    if(credOficial <= 0.01 && provVal > 0){
      faixa = 'alta';
    } else if(totEmpNE > 0 || consumido > 0){
      if(taxaQueima >= 80) faixa = 'alta';
      else if(taxaQueima >= 30) faixa = 'media';
      else faixa = 'baixa';
    } else {
      faixa = 'zero';
    }

    var celKey = (it.fav_cod || it.uasg) + '_' + it.pi + '_' + it.nd;
    var cLiq = (EMPENHODATA.celulas_liq && EMPENHODATA.celulas_liq[celKey]) ? EMPENHODATA.celulas_liq[celKey] :
               (EMPENHODATA.pi_liq && EMPENHODATA.pi_liq[it.pi]) ? EMPENHODATA.pi_liq[it.pi] : null;
    var ncLiq = cLiq ? (cLiq.liq || 0) : (it.liq || 0);
    var ncPag = cLiq ? (cLiq.pag || 0) : (it.pag || 0);

    var ncObj = {
      hid: it.hid,
      nc: it.nc,
      dia: it.dia || '',
      emit: it.emit_cod || it.emit || '',
      emit_nome: it.emit_nome || '',
      uasg: it.fav_cod || '',
      uasg_nome: it.fav_nome || '',
      om_sigla: it.om_sigla || '',
      pi: it.pi || '',
      pi_nome: it.pi_nome || '',
      nd: it.nd || '',
      nd_desc: it.nd_desc || '',
      obj: it.obj || '',
      prov: provVal,
      emp: totEmpNE > 0 ? totEmpNE : consumido,
      liq: ncLiq,
      pag: ncPag,
      cred: credOficial,
      taxa_queima: taxaQueima,
      faixa_queima: faixa,
      nes: nesVinculadas
    };
    RASTREIO_NC_ITEMS.push(ncObj);
  }

  // 2. Base Geral de Empenhos
  RASTREIO_NE_ITEMS = EMPENHODATA.nes || [];

  RASTREIO_INITIALIZED = true;
  bcmsFiltraRastreioNC();
  bcmsFiltraRastreioNE();
  bcmsFiltraForn();
  bcmsFiltraProva();
  bcmsFiltraRisco();
}

function bcmsRastreioSubAba(subtab, btn){
  var tabs = document.querySelectorAll('#secao-RASTREIO .toptab');
  for(var i = 0; i < tabs.length; i++){
    tabs[i].classList.remove('on');
    tabs[i].setAttribute('aria-selected', 'false');
  }
  if(btn){
    btn.classList.add('on');
    btn.setAttribute('aria-selected', 'true');
  } else {
    var targetBtn = document.getElementById('r-subtab-btn-' + subtab);
    if(targetBtn){
      targetBtn.classList.add('on');
      targetBtn.setAttribute('aria-selected', 'true');
    }
  }
  var panes = document.querySelectorAll('#secao-RASTREIO .r-subtab-pane');
  for(var j = 0; j < panes.length; j++){
    panes[j].style.display = (panes[j].id === 'pane-rastreio-' + subtab) ? '' : 'none';
  }
}

/* ================= SUB-ABA 1: VISÃO TOP-DOWN (NOTAS DE CRÉDITO) ================= */
function bcmsFiltraRastreioNC(){
  var q = (document.getElementById('rBuscaNC') ? document.getElementById('rBuscaNC').value : '').toLowerCase().trim();
  var om = document.getElementById('rSelOM') ? document.getElementById('rSelOM').value : '';
  var nd = document.getElementById('rSelNDNC') ? document.getElementById('rSelNDNC').value : '';
  var fonte = document.getElementById('rSelFonteNC') ? document.getElementById('rSelFonteNC').value : '';
  var qm = document.getElementById('rSelQueima') ? document.getElementById('rSelQueima').value : '';

  RASTREIO_NC_FILTRADOS = RASTREIO_NC_ITEMS.filter(function(it){
    if(om && it.uasg !== om) return false;
    if(nd && it.nd !== nd) return false;
    if(fonte && ((fonte === '160' && it.uasg.indexOf('160') !== 0) || (fonte === '167' && it.uasg.indexOf('167') !== 0))) return false;
    if(qm && it.faixa_queima !== qm) return false;
    if(q){
      var match = (it.nc && it.nc.toLowerCase().indexOf(q) !== -1) ||
                  (it.pi && it.pi.toLowerCase().indexOf(q) !== -1) ||
                  (it.nd && it.nd.toLowerCase().indexOf(q) !== -1) ||
                  (it.emit_nome && it.emit_nome.toLowerCase().indexOf(q) !== -1) ||
                  (it.obj && it.obj.toLowerCase().indexOf(q) !== -1);
      if(!match && it.nes){
        for(var n = 0; n < it.nes.length; n++){
          if(it.nes[n].fav.toLowerCase().indexOf(q) !== -1 || it.nes[n].ne.toLowerCase().indexOf(q) !== -1){
            match = true; break;
          }
        }
      }
      if(!match) return false;
    }
    return true;
  });

  RASTREIO_NC_PAGE = 1;
  bcmsSortApplyNC();
  bcmsRenderRastreioNC();
}

function bcmsSortRastreioNC(col){
  if(RASTREIO_SORT_NC.col === col){
    RASTREIO_SORT_NC.asc = !RASTREIO_SORT_NC.asc;
  } else {
    RASTREIO_SORT_NC.col = col;
    RASTREIO_SORT_NC.asc = (col === 'nc' || col === 'emit_nome' || col === 'om_sigla');
  }

  var sortIds = ['nc', 'dia', 'emit', 'uasg', 'pi', 'nd', 'prov', 'emp', 'liq', 'cred'];
  for(var s = 0; s < sortIds.length; s++){
    var elS = document.getElementById('sort-nc-' + sortIds[s]);
    if(elS){
      var match = (sortIds[s] === col || (sortIds[s] === 'emit' && col === 'emit_nome') || (sortIds[s] === 'uasg' && col === 'om_sigla'));
      elS.textContent = match ? (RASTREIO_SORT_NC.asc ? '▲' : '▼') : '';
    }
  }

  bcmsSortApplyNC();
  bcmsRenderRastreioNC();
}

function bcmsSortApplyNC(){
  var col = RASTREIO_SORT_NC.col;
  var asc = RASTREIO_SORT_NC.asc;
  RASTREIO_NC_FILTRADOS.sort(function(a, b){
    var va = a[col] != null ? a[col] : '';
    var vb = b[col] != null ? b[col] : '';
    if(typeof va === 'number' && typeof vb === 'number'){
      return asc ? (va - vb) : (vb - va);
    }
    return asc ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));
  });
}

function bcmsRenderRastreioNC(){
  var tbody = document.getElementById('corpoRastreioNC');
  var cont = document.getElementById('rContadorNC');
  if(!tbody) return;

  var total = RASTREIO_NC_FILTRADOS.length;
  if(cont){
    cont.textContent = 'Exibindo ' + total.toLocaleString('pt-BR') + ' de ' + RASTREIO_NC_ITEMS.length.toLocaleString('pt-BR') + ' Notas de Crédito';
  }

  var porPag = 25;
  var totalPags = Math.max(1, Math.ceil(total / porPag));
  if(RASTREIO_NC_PAGE > totalPags) RASTREIO_NC_PAGE = totalPags;
  var ini = (RASTREIO_NC_PAGE - 1) * porPag;
  var fim = Math.min(total, ini + porPag);
  var paginaItens = RASTREIO_NC_FILTRADOS.slice(ini, fim);

  if(paginaItens.length === 0){
    tbody.innerHTML = '<tr><td colspan="13" style="text-align:center;padding:24px;color:var(--ink-muted);">Nenhuma Nota de Crédito encontrada para os filtros aplicados.</td></tr>';
    document.getElementById('paginacaoRastreioNC').innerHTML = '';
    return;
  }

  var html = '';
  for(var i = 0; i < paginaItens.length; i++){
    var it = paginaItens[i];
    var temNE = it.nes && it.nes.length > 0;
    var queimaCls = it.faixa_queima;
    var queimaFillW = Math.min(100, Math.round(it.taxa_queima));
    var emitCurto = it.emit_nome ? (it.emit_nome.length > 22 ? it.emit_nome.slice(0, 22) + '…' : it.emit_nome) : it.emit;

    html += '<tr class="cel-row" id="linha-nc-' + it.hid + '">' +
      '<td style="text-align:center;">' +
        (temNE ? '<button type="button" class="btn-exp-ne" id="btn-exp-' + it.hid + '" onclick="bcmsToggleAcordeaoNC(\'' + it.hid + '\')" title="Expandir ' + it.nes.length + ' empenhos emitidos">▼ ' + it.nes.length + '</button>' : '<span style="color:var(--ink-soft);font-size:0.75rem;">—</span>') +
      '</td>' +
      '<td class="mono2"><a href="javascript:void(0)" onclick="bcmsDetalheNC(\'' + it.hid + '\')" class="link-drill" title="Abrir ficha cadastral da NC">' + bcmsEsc(it.nc) + '</a></td>' +
      '<td class="mono2">' + bcmsEsc(it.dia || '—') + '</td>' +
      '<td title="' + bcmsEsc(it.emit_nome) + '"><span class="ug-pill emit">' + bcmsEsc(it.emit) + '</span> <small>' + bcmsEsc(emitCurto) + '</small></td>' +
      '<td><span class="ug-pill fav">' + bcmsEsc(it.uasg) + '</span> <b>' + bcmsEsc(it.om_sigla) + '</b></td>' +
      '<td class="mono2">' + bcmsEsc(it.pi || '—') + '</td>' +
      '<td class="mono2">' + bcmsEsc(it.nd || '—') + '</td>' +
      '<td class="num font-mono">' + bcmsBRL(it.prov) + '</td>' +
      '<td class="num font-mono font-bold" style="color:var(--success-strong);">' + bcmsBRL(it.emp) + '</td>' +
      '<td class="num font-mono" style="color:#059669;font-weight:600;">' + (it.liq > 0.005 ? bcmsBRL(it.liq) : '<span style="color:var(--ink-soft);">R$ 0,00</span>') + '</td>' +
      '<td class="num font-mono anchor">' + bcmsBRL(it.cred) + '</td>' +
      '<td>' +
        '<div class="queima-bar-wrap" title="' + it.taxa_queima.toFixed(1) + '% empenhado">' +
          '<div class="queima-track"><div class="queima-fill ' + queimaCls + '" style="width:' + queimaFillW + '%"></div></div>' +
          '<span class="queima-lbl">' + it.taxa_queima.toFixed(0) + '%</span>' +
        '</div>' +
      '</td>' +
      '<td style="text-align:center;"><button type="button" class="tbl-action-btn" onclick="bcmsDetalheNC(\'' + it.hid + '\')">Ficha ↗</button></td>' +
    '</tr>' +
    '<tr id="acordeao-' + it.hid + '" class="sub-row-detalhe" style="display:none">' +
      '<td colspan="13" style="padding:0;">' +
        '<div class="sub-ne-wrap" id="acordeao-cont-' + it.hid + '"></div>' +
      '</td>' +
    '</tr>';
  }
  tbody.innerHTML = html;

  var pagHtml = '<span class="pag-info" style="font-size:0.8125rem;color:var(--ink-muted);font-weight:600;">Página ' + RASTREIO_NC_PAGE + ' de ' + totalPags + ' (Exibindo ' + (ini + 1) + '–' + fim + ' de ' + total.toLocaleString('pt-BR') + ')</span>' +
    '<div style="display:flex;gap:8px;align-items:center;">' +
      '<button type="button" class="flt-limpa" onclick="bcmsPaginaRastreioNC(' + (RASTREIO_NC_PAGE - 1) + ')" ' + (RASTREIO_NC_PAGE <= 1 ? 'disabled' : '') + '>‹ Anterior</button>' +
      '<button type="button" class="flt-limpa" onclick="bcmsPaginaRastreioNC(' + (RASTREIO_NC_PAGE + 1) + ')" ' + (RASTREIO_NC_PAGE >= totalPags ? 'disabled' : '') + '>Próxima ›</button>' +
    '</div>';
  document.getElementById('paginacaoRastreioNC').innerHTML = pagHtml;
}

function bcmsPaginaRastreioNC(p){
  if(p < 1) return;
  RASTREIO_NC_PAGE = p;
  bcmsRenderRastreioNC();
  var scr = document.getElementById('scroll-rastreio-nc');
  if(scr) scr.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function bcmsToggleAcordeaoNC(hid){
  var tr = document.getElementById('acordeao-' + hid);
  var btn = document.getElementById('btn-exp-' + hid);
  var cont = document.getElementById('acordeao-cont-' + hid);
  if(!tr || !cont) return;

  var estaAberto = (tr.style.display !== 'none');
  if(estaAberto){
    tr.style.display = 'none';
    if(btn) btn.classList.remove('aberto');
    return;
  }

  var item = null;
  for(var i = 0; i < RASTREIO_NC_ITEMS.length; i++){
    if(RASTREIO_NC_ITEMS[i].hid === hid){
      item = RASTREIO_NC_ITEMS[i]; break;
    }
  }
  if(!item || !item.nes) return;

  var nes = item.nes;
  var miniHtml = '<div class="sub-ne-header">' +
    '<div class="sub-ne-title">📦 ' + nes.length + ' Nota(s) de Empenho vinculadas à NC ' + bcmsEsc(item.nc) + '</div>' +
    '<div class="sub-ne-meta">Total Empenhado: <b style="color:var(--success-strong);">' + bcmsBRL(item.emp) + '</b> de <b>' + bcmsBRL(item.prov) + '</b> (' + item.taxa_queima.toFixed(1) + '% executado) · Liquidado no SIAFI: <b style="color:#059669;">' + bcmsBRL(item.liq||0) + '</b> (Pago: <b>' + bcmsBRL(item.pag||0) + '</b>)</div>' +
  '</div>' +
  '<table class="det det-compact">' +
    '<thead>' +
      '<tr>' +
        '<th>Número NE</th>' +
        '<th>Data Emissão</th>' +
        '<th>Favorecido / Fornecedor</th>' +
        '<th>CNPJ / CPF</th>' +
        '<th>Processo / Pregão</th>' +
        '<th class="num">Valor da NE</th>' +
        '<th style="text-align:center;">Selo Prova Real</th>' +
        '<th style="text-align:center;">Ação</th>' +
      '</tr>' +
    '</thead>' +
    '<tbody>';

  for(var j = 0; j < nes.length; j++){
    var nItem = nes[j];
    var sCls = nItem.prova_slug === 'ok' ? 'status-disp' :
              (nItem.prova_slug === 'info' ? 'status-parcial' :
              (nItem.prova_slug === 'warn' ? 'status-canc' :
              (nItem.prova_slug === 'danger' ? 'status-canc' : 'status-zerada')));
    var badgeProva = '<span class="pill-status ' + sCls + '">' + bcmsEsc(nItem.prova || 'Auditado') + '</span>';
    var pregaoTag = nItem.proc ? '<a href="javascript:void(0)" onclick="bcmsDetalheProcesso(\'' + bcmsEsc(nItem.proc) + '\')" class="badge-pregao" style="cursor:pointer;">' + bcmsEsc(nItem.proc) + '</a>' : '—';

    miniHtml += '<tr class="cel-row">' +
      '<td class="mono2 font-bold"><a href="javascript:void(0)" onclick="bcmsDetalheNE(\'' + bcmsEsc(nItem.ne) + '\')" class="link-drill">' + bcmsEsc(nItem.ne) + '</a></td>' +
      '<td class="mono2">' + bcmsEsc(nItem.dia || '—') + '</td>' +
      '<td><a href="javascript:void(0)" onclick="bcmsDetalheFornecedor(\'' + bcmsEsc(nItem.fav) + '\')" class="link-drill" style="color:var(--ink);">' + bcmsEsc(nItem.fav) + '</a></td>' +
      '<td class="mono2">' + bcmsEsc(nItem.doc || '—') + '</td>' +
      '<td>' + pregaoTag + '</td>' +
      '<td class="num font-mono font-bold" style="color:var(--success-strong);">' + bcmsBRL(nItem.val) + '</td>' +
      '<td style="text-align:center;">' + badgeProva + '</td>' +
      '<td style="text-align:center;"><button type="button" class="tbl-action-btn" onclick="bcmsDetalheNE(\'' + bcmsEsc(nItem.ne) + '\')">Detalhar ↗</button></td>' +
    '</tr>';
  }
  miniHtml += '</tbody></table>';

  cont.innerHTML = miniHtml;
  tr.style.display = '';
  if(btn) btn.classList.add('aberto');
}

function bcmsLimpaFiltrosRastreioNC(){
  if(document.getElementById('rBuscaNC')) document.getElementById('rBuscaNC').value = '';
  if(document.getElementById('rSelOM')) document.getElementById('rSelOM').value = '';
  if(document.getElementById('rSelNDNC')) document.getElementById('rSelNDNC').value = '';
  if(document.getElementById('rSelFonteNC')) document.getElementById('rSelFonteNC').value = '';
  if(document.getElementById('rSelQueima')) document.getElementById('rSelQueima').value = '';
  bcmsFiltraRastreioNC();
}

/* ================= SUB-ABA 2: RELAÇÃO DE EMPENHOS (BOTTOM-UP) ================= */
function bcmsFiltraRastreioNE(){
  var q = (document.getElementById('rBuscaNE') ? document.getElementById('rBuscaNE').value : '').toLowerCase().trim();
  var om = document.getElementById('rSelOMNE') ? document.getElementById('rSelOMNE').value : '';
  var nd = document.getElementById('rSelNDNE') ? document.getElementById('rSelNDNE').value : '';
  var prova = document.getElementById('rSelProvaNE') ? document.getElementById('rSelProvaNE').value : '';

  RASTREIO_NE_FILTRADOS = RASTREIO_NE_ITEMS.filter(function(it){
    if(om && it.ug !== om) return false;
    if(nd && it.nd !== nd) return false;
    if(prova && it.prova_slug !== prova) return false;
    if(q){
      var match = (it.ne && it.ne.toLowerCase().indexOf(q) !== -1) ||
                  (it.fav && it.fav.toLowerCase().indexOf(q) !== -1) ||
                  (it.doc && it.doc.indexOf(q) !== -1) ||
                  (it.nc && it.nc.toLowerCase().indexOf(q) !== -1) ||
                  (it.proc && it.proc.toLowerCase().indexOf(q) !== -1) ||
                  (it.desc && it.desc.toLowerCase().indexOf(q) !== -1);
      if(!match) return false;
    }
    return true;
  });

  RASTREIO_NE_PAGE = 1;
  bcmsSortApplyNE();
  bcmsRenderRastreioNE();
}

function bcmsSortRastreioNE(col){
  if(RASTREIO_SORT_NE.col === col){
    RASTREIO_SORT_NE.asc = !RASTREIO_SORT_NE.asc;
  } else {
    RASTREIO_SORT_NE.col = col;
    RASTREIO_SORT_NE.asc = (col === 'ne' || col === 'fav' || col === 'ug');
  }

  var sortIds = ['ne', 'dia', 'ug', 'fav', 'proc', 'val', 'liq', 'nd', 'pi', 'nc'];
  for(var s = 0; s < sortIds.length; s++){
    var elS = document.getElementById('sort-ne-' + sortIds[s]);
    if(elS){
      var match = (sortIds[s] === col || (sortIds[s] === 'liq' && col === 'cel_liq'));
      elS.textContent = match ? (RASTREIO_SORT_NE.asc ? '▲' : '▼') : '';
    }
  }

  bcmsSortApplyNE();
  bcmsRenderRastreioNE();
}

function bcmsSortApplyNE(){
  var col = RASTREIO_SORT_NE.col;
  var asc = RASTREIO_SORT_NE.asc;
  RASTREIO_NE_FILTRADOS.sort(function(a, b){
    var va = a[col] != null ? a[col] : '';
    var vb = b[col] != null ? b[col] : '';
    if(typeof va === 'number' && typeof vb === 'number'){
      return asc ? (va - vb) : (vb - va);
    }
    return asc ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));
  });
}

function bcmsRenderRastreioNE(){
  var tbody = document.getElementById('corpoRastreioNE');
  var cont = document.getElementById('rContadorNE');
  if(!tbody) return;

  var total = RASTREIO_NE_FILTRADOS.length;
  if(cont){
    cont.textContent = 'Exibindo ' + total.toLocaleString('pt-BR') + ' de ' + RASTREIO_NE_ITEMS.length.toLocaleString('pt-BR') + ' Notas de Empenho';
  }

  var porPag = 50;
  var totalPags = Math.max(1, Math.ceil(total / porPag));
  if(RASTREIO_NE_PAGE > totalPags) RASTREIO_NE_PAGE = totalPags;
  var ini = (RASTREIO_NE_PAGE - 1) * porPag;
  var fim = Math.min(total, ini + porPag);
  var paginaItens = RASTREIO_NE_FILTRADOS.slice(ini, fim);

  if(paginaItens.length === 0){
    tbody.innerHTML = '<tr><td colspan="13" style="text-align:center;padding:24px;color:var(--ink-muted);">Nenhum empenho encontrado para os filtros aplicados.</td></tr>';
    document.getElementById('paginacaoRastreioNE').innerHTML = '';
    return;
  }

  var html = '';
  for(var i = 0; i < paginaItens.length; i++){
    var it = paginaItens[i];
    var sCls = it.prova_slug === 'ok' ? 'status-disp' :
              (it.prova_slug === 'info' ? 'status-parcial' :
              (it.prova_slug === 'warn' ? 'status-canc' :
              (it.prova_slug === 'danger' ? 'status-canc' : 'status-zerada')));
    var badgeProva = '<span class="pill-status ' + sCls + '">' + bcmsEsc(it.prova || 'Célula SIAFI') + '</span>';
    var pregaoTag = it.proc ? '<a href="javascript:void(0)" onclick="bcmsDetalheProcesso(\'' + bcmsEsc(it.proc) + '\')" class="badge-pregao" style="cursor:pointer;">' + bcmsEsc(it.proc) + '</a>' : '—';
    var ncLink = it.nc ? '<a href="javascript:void(0)" onclick="bcmsDetalheNC(\'' + it.nc + '\')" class="link-drill">' + bcmsEsc(it.nc) + '</a>' : '<span style="color:var(--ink-muted);">—</span>';

    var liqCelExib = (it.cel_liq > 0.005 ? bcmsBRL(it.cel_liq) : '<span style="color:var(--ink-soft);">R$ 0,00</span>');
    if(it.cel_liq > 0.005 && it.pct_liq > 0){
      liqCelExib += ' <small style="display:block;font-size:0.7rem;color:#059669;">' + it.pct_liq.toFixed(0) + '% Célula</small>';
    }

    html += '<tr class="cel-row">' +
      '<td class="mono2 font-bold"><a href="javascript:void(0)" onclick="bcmsDetalheNE(\'' + bcmsEsc(it.ne) + '\')" class="link-drill">' + bcmsEsc(it.ne) + '</a></td>' +
      '<td class="mono2">' + bcmsEsc(it.dia || '—') + '</td>' +
      '<td><span class="ug-pill fav">' + bcmsEsc(it.ug) + '</span></td>' +
      '<td title="' + bcmsEsc(it.fav) + '"><a href="javascript:void(0)" onclick="bcmsDetalheFornecedor(\'' + bcmsEsc(it.fav) + '\')" class="link-drill" style="color:var(--ink);">' + bcmsEsc(it.fav) + '</a></td>' +
      '<td class="mono2">' + bcmsEsc(it.doc || '—') + '</td>' +
      '<td>' + pregaoTag + '</td>' +
      '<td class="num font-mono font-bold anchor" style="color:var(--success-strong);">' + bcmsBRL(it.val) + '</td>' +
      '<td class="num font-mono" style="color:#059669;font-weight:700;">' + liqCelExib + '</td>' +
      '<td class="mono2">' + bcmsEsc(it.nd) + '</td>' +
      '<td class="mono2">' + bcmsEsc(it.pi) + '</td>' +
      '<td class="mono2">' + ncLink + '</td>' +
      '<td style="text-align:center;">' + badgeProva + '</td>' +
      '<td style="text-align:center;"><button type="button" class="tbl-action-btn" onclick="bcmsDetalheNE(\'' + bcmsEsc(it.ne) + '\')">Detalhar ↗</button></td>' +
    '</tr>';
  }
  tbody.innerHTML = html;

  var pagHtml = '<span class="pag-info" style="font-size:0.8125rem;color:var(--ink-muted);font-weight:600;">Página ' + RASTREIO_NE_PAGE + ' de ' + totalPags + ' (Exibindo ' + (ini + 1) + '–' + fim + ' de ' + total.toLocaleString('pt-BR') + ')</span>' +
    '<div style="display:flex;gap:8px;align-items:center;">' +
      '<button type="button" class="flt-limpa" onclick="bcmsPaginaRastreioNE(' + (RASTREIO_NE_PAGE - 1) + ')" ' + (RASTREIO_NE_PAGE <= 1 ? 'disabled' : '') + '>‹ Anterior</button>' +
      '<button type="button" class="flt-limpa" onclick="bcmsPaginaRastreioNE(' + (RASTREIO_NE_PAGE + 1) + ')" ' + (RASTREIO_NE_PAGE >= totalPags ? 'disabled' : '') + '>Próxima ›</button>' +
    '</div>';
  document.getElementById('paginacaoRastreioNE').innerHTML = pagHtml;
}

function bcmsPaginaRastreioNE(p){
  if(p < 1) return;
  RASTREIO_NE_PAGE = p;
  bcmsRenderRastreioNE();
  var scr = document.getElementById('scroll-rastreio-ne');
  if(scr) scr.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function bcmsLimpaFiltrosRastreioNE(){
  if(document.getElementById('rBuscaNE')) document.getElementById('rBuscaNE').value = '';
  if(document.getElementById('rSelOMNE')) document.getElementById('rSelOMNE').value = '';
  if(document.getElementById('rSelNDNE')) document.getElementById('rSelNDNE').value = '';
  if(document.getElementById('rSelProvaNE')) document.getElementById('rSelProvaNE').value = '';
  bcmsFiltraRastreioNE();
}

/* ================= SUB-ABA 3: RAIO-X FORNECEDORES & PREGÕES ================= */
function bcmsFiltraForn(){
  var q = (document.getElementById('rBuscaForn') ? document.getElementById('rBuscaForn').value : '').toLowerCase().trim();
  var om = document.getElementById('rSelOMForn') ? document.getElementById('rSelOMForn').value : '';
  var nd = document.getElementById('rSelNDForn') ? document.getElementById('rSelNDForn').value : '';

  var topF = {};
  var topP = {};

  for(var i = 0; i < RASTREIO_NE_ITEMS.length; i++){
    var it = RASTREIO_NE_ITEMS[i];
    if(om && it.ug !== om) continue;
    if(nd && it.nd !== nd) continue;
    if(q){
      var match = (it.fav && it.fav.toLowerCase().indexOf(q) !== -1) ||
                  (it.doc && it.doc.indexOf(q) !== -1) ||
                  (it.proc && it.proc.toLowerCase().indexOf(q) !== -1);
      if(!match) continue;
    }

    if(it.fav && it.fav !== 'NAO SE APLICA'){
      topF[it.fav] = (topF[it.fav] || 0) + it.val;
    }
    if(it.proc){
      topP[it.proc] = (topP[it.proc] || 0) + it.val;
    }
  }

  // Ordena top fornecedores
  var listF = Object.keys(topF).map(function(k){ return { nome: k, val: topF[k] }; });
  listF.sort(function(a, b){ return b.val - a.val; });
  var top25F = listF.slice(0, 25);
  var maxF = top25F.length > 0 ? top25F[0].val : 1;

  var elF = document.getElementById('listaTopFornecedores');
  if(elF){
    if(top25F.length === 0){
      elF.innerHTML = '<p class="vazio" style="padding:20px;text-align:center;color:var(--ink-muted);">Nenhum fornecedor encontrado para os filtros selecionados.</p>';
    } else {
      var htmlF = '';
      for(var fIdx = 0; fIdx < top25F.length; fIdx++){
        var itemF = top25F[fIdx];
        var pctF = Math.min(100, Math.round(itemF.val / maxF * 100));
        htmlF += '<div class="rank-item" onclick="bcmsDetalheFornecedor(\'' + bcmsEsc(itemF.nome) + '\')" title="Clique para ver todos os empenhos deste fornecedor">' +
          '<div class="rank-info">' +
            '<span class="rank-pos">#' + (fIdx + 1) + '</span>' +
            '<span class="rank-nome">' + bcmsEsc(itemF.nome) + '</span>' +
            '<div class="kpi-bar-wrap"><div class="kpi-bar-fill" style="width:' + pctF + '%"></div></div>' +
          '</div>' +
          '<div class="rank-val">' + bcmsBRL(itemF.val) + '</div>' +
        '</div>';
      }
      elF.innerHTML = htmlF;
    }
  }

  // Ordena top processos
  var listP = Object.keys(topP).map(function(k){ return { proc: k, val: topP[k] }; });
  listP.sort(function(a, b){ return b.val - a.val; });
  var top20P = listP.slice(0, 20);
  var maxP = top20P.length > 0 ? top20P[0].val : 1;

  var elP = document.getElementById('listaTopProcessos');
  if(elP){
    if(top20P.length === 0){
      elP.innerHTML = '<p class="vazio" style="padding:20px;text-align:center;color:var(--ink-muted);">Nenhum processo/pregão encontrado para os filtros selecionados.</p>';
    } else {
      var htmlP = '';
      for(var pIdx = 0; pIdx < top20P.length; pIdx++){
        var itemP = top20P[pIdx];
        var pctP = Math.min(100, Math.round(itemP.val / maxP * 100));
        htmlP += '<div class="rank-item" onclick="bcmsDetalheProcesso(\'' + bcmsEsc(itemP.proc) + '\')" title="Clique para ver todos os empenhos deste pregão">' +
          '<div class="rank-info">' +
            '<span class="rank-pos">#' + (pIdx + 1) + '</span>' +
            '<span class="rank-nome">' + bcmsEsc(itemP.proc) + '</span>' +
            '<div class="kpi-bar-wrap"><div class="kpi-bar-fill" style="width:' + pctP + '%"></div></div>' +
          '</div>' +
          '<div class="rank-val">' + bcmsBRL(itemP.val) + '</div>' +
        '</div>';
      }
      elP.innerHTML = htmlP;
    }
  }

  var cForn = document.getElementById('rContadorForn');
  if(cForn){
    cForn.textContent = listF.length + ' Credores · ' + listP.length + ' Processos';
  }
}

function bcmsLimpaFiltrosForn(){
  if(document.getElementById('rBuscaForn')) document.getElementById('rBuscaForn').value = '';
  if(document.getElementById('rSelOMForn')) document.getElementById('rSelOMForn').value = '';
  if(document.getElementById('rSelNDForn')) document.getElementById('rSelNDForn').value = '';
  bcmsFiltraForn();
}

/* ================= SUB-ABA 4: PROVA REAL & AUDITORIA ================= */
function bcmsFiltraProva(){
  var q = (document.getElementById('rBuscaProva') ? document.getElementById('rBuscaProva').value : '').toLowerCase().trim();
  var om = document.getElementById('rSelOMProva') ? document.getElementById('rSelOMProva').value : '';
  var statusSel = document.getElementById('rSelStatusProva') ? document.getElementById('rSelStatusProva').value : '';

  var incs = (EMPENHODATA && EMPENHODATA.inconsistencias) ? EMPENHODATA.inconsistencias : [];
  var filtrados = incs.filter(function(it){
    if(om && it.ug !== om) return false;
    if(statusSel && it.status_slug !== statusSel) return false;
    if(q){
      var match = (it.ne && it.ne.toLowerCase().indexOf(q) !== -1) ||
                  (it.fav && it.fav.toLowerCase().indexOf(q) !== -1) ||
                  (it.nc_citada && it.nc_citada.toLowerCase().indexOf(q) !== -1) ||
                  (it.motivo && it.motivo.toLowerCase().indexOf(q) !== -1) ||
                  (it.pi_ne && it.pi_ne.toLowerCase().indexOf(q) !== -1);
      if(!match) return false;
    }
    return true;
  });

  var elInc = document.getElementById('corpoInconsistencias');
  var cont = document.getElementById('rContadorProva');
  if(cont){
    cont.textContent = 'Exibindo ' + filtrados.length + ' de ' + incs.length + ' Inconsistências';
  }

  if(elInc){
    if(filtrados.length === 0){
      elInc.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:24px;color:var(--ink-muted);">Nenhuma inconsistência encontrada para os filtros aplicados.</td></tr>';
      return;
    }
    var htmlInc = '';
    for(var i = 0; i < filtrados.length; i++){
      var inItem = filtrados[i];
      var sCls = inItem.status_slug === 'ok' ? 'status-disp' :
                (inItem.status_slug === 'info' ? 'status-parcial' :
                (inItem.status_slug === 'warn' ? 'status-canc' :
                (inItem.status_slug === 'danger' ? 'status-canc' : 'status-zerada')));

      htmlInc += '<tr class="cel-row">' +
        '<td class="mono2 font-bold"><a href="javascript:void(0)" onclick="bcmsDetalheNE(\'' + bcmsEsc(inItem.ne) + '\')" class="link-drill">' + bcmsEsc(inItem.ne) + '</a></td>' +
        '<td class="mono2">' + bcmsEsc(inItem.dia || '—') + '</td>' +
        '<td><span class="ug-pill fav">' + bcmsEsc(inItem.ug) + '</span></td>' +
        '<td><a href="javascript:void(0)" onclick="bcmsDetalheFornecedor(\'' + bcmsEsc(inItem.fav) + '\')" class="link-drill" style="color:var(--ink);">' + bcmsEsc(inItem.fav) + '</a></td>' +
        '<td class="num font-mono font-bold" style="color:var(--success-strong);">' + bcmsEsc(inItem.val_fmt) + '</td>' +
        '<td class="mono2">' + bcmsEsc(inItem.pi_ne) + '</td>' +
        '<td class="mono2">' + bcmsEsc(inItem.nd_ne) + '</td>' +
        '<td class="mono2"><a href="javascript:void(0)" onclick="bcmsDetalheNC(\'' + bcmsEsc(inItem.nc_citada) + '\')" class="link-drill" style="color:var(--danger);font-weight:bold;">' + bcmsEsc(inItem.nc_citada) + '</a></td>' +
        '<td><span class="pill-status ' + sCls + '">' + bcmsEsc(inItem.status_prova) + '</span><br><small style="color:var(--ink-muted);font-size:0.75rem;">' + bcmsEsc(inItem.motivo) + '</small></td>' +
        '<td style="text-align:center;"><button type="button" class="tbl-action-btn" onclick="bcmsDetalheNE(\'' + bcmsEsc(inItem.ne) + '\')">Inspecionar ↗</button></td>' +
      '</tr>';
    }
    elInc.innerHTML = htmlInc;
  }
}

function bcmsLimpaFiltrosProva(){
  if(document.getElementById('rBuscaProva')) document.getElementById('rBuscaProva').value = '';
  if(document.getElementById('rSelOMProva')) document.getElementById('rSelOMProva').value = '';
  if(document.getElementById('rSelStatusProva')) document.getElementById('rSelStatusProva').value = '';
  bcmsFiltraProva();
}

/* ================= SUB-ABA 5: CRÉDITOS PARADOS (RISCO) ================= */
function bcmsFiltraRisco(){
  var q = (document.getElementById('rBuscaRisco') ? document.getElementById('rBuscaRisco').value : '').toLowerCase().trim();
  var om = document.getElementById('rSelOMRisco') ? document.getElementById('rSelOMRisco').value : '';
  var faixa = document.getElementById('rSelFaixaRisco') ? document.getElementById('rSelFaixaRisco').value : '';

  var threshold = 15000;
  if(faixa === '50k') threshold = 50000;
  else if(faixa === '100k') threshold = 100000;
  else if(faixa === '500k') threshold = 500000;

  var parados = RASTREIO_NC_ITEMS.filter(function(x){
    if(x.cred < threshold || x.emp > 0) return false;
    if(om && x.uasg !== om) return false;
    if(q){
      var match = (x.nc && x.nc.toLowerCase().indexOf(q) !== -1) ||
                  (x.emit_nome && x.emit_nome.toLowerCase().indexOf(q) !== -1) ||
                  (x.pi && x.pi.toLowerCase().indexOf(q) !== -1) ||
                  (x.nd && x.nd.toLowerCase().indexOf(q) !== -1) ||
                  (x.obj && x.obj.toLowerCase().indexOf(q) !== -1);
      if(!match) return false;
    }
    return true;
  }).sort(function(a, b){ return b.cred - a.cred; });

  var elRisco = document.getElementById('corpoRastreioRisco');
  var cont = document.getElementById('rContadorRisco');
  if(cont){
    cont.textContent = parados.length + ' Créditos em Risco Crítico';
  }

  if(elRisco){
    if(parados.length === 0){
      elRisco.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:24px;color:var(--ink-muted);">Nenhum crédito com saldo parado em risco crítico identificado para os filtros selecionados.</td></tr>';
      return;
    }
    var htmlR = '';
    for(var rI = 0; rI < parados.length; rI++){
      var pItem = parados[rI];
      var emitCurto = pItem.emit_nome ? (pItem.emit_nome.length > 22 ? pItem.emit_nome.slice(0, 22) + '…' : pItem.emit_nome) : pItem.emit;

      htmlR += '<tr class="cel-row">' +
        '<td class="mono2"><a href="javascript:void(0)" onclick="bcmsDetalheNC(\'' + bcmsEsc(pItem.hid || pItem.nc) + '\')" class="link-drill">' + bcmsEsc(pItem.nc) + '</a></td>' +
        '<td class="mono2">' + bcmsEsc(pItem.dia || '—') + '</td>' +
        '<td><span class="pill-status status-canc">&gt; 45 dias</span></td>' +
        '<td title="' + bcmsEsc(pItem.emit_nome) + '"><span class="ug-pill emit">' + bcmsEsc(pItem.emit) + '</span> <small>' + bcmsEsc(emitCurto) + '</small></td>' +
        '<td><span class="ug-pill fav">' + bcmsEsc(pItem.uasg) + '</span> <b>' + bcmsEsc(pItem.om_sigla) + '</b></td>' +
        '<td class="mono2">' + bcmsEsc(pItem.pi) + '</td>' +
        '<td class="mono2">' + bcmsEsc(pItem.nd) + '</td>' +
        '<td class="num font-mono">' + bcmsBRL(pItem.prov) + '</td>' +
        '<td class="num font-mono anchor" style="color:var(--warning-ink);">' + bcmsBRL(pItem.cred) + '</td>' +
        '<td style="text-align:center;"><button type="button" class="tbl-action-btn" onclick="bcmsDetalheNC(\'' + bcmsEsc(pItem.hid || pItem.nc) + '\')">Inspecionar ↗</button></td>' +
      '</tr>';
    }
    elRisco.innerHTML = htmlR;
  }
}

function bcmsLimpaFiltrosRisco(){
  if(document.getElementById('rBuscaRisco')) document.getElementById('rBuscaRisco').value = '';
  if(document.getElementById('rSelOMRisco')) document.getElementById('rSelOMRisco').value = '';
  if(document.getElementById('rSelFaixaRisco')) document.getElementById('rSelFaixaRisco').value = '';
  bcmsFiltraRisco();
}

/* ================= MODAIS DE DRILL-DOWN PROFUNDO ================= */

function bcmsDetalheNE(neNum){
  if(!neNum) return;
  var neClean = String(neNum).trim();
  var neObj = null;
  if(typeof RASTREIO_NE_ITEMS !== 'undefined' && RASTREIO_NE_ITEMS){
    for(var i = 0; i < RASTREIO_NE_ITEMS.length; i++){
      var itemNe = RASTREIO_NE_ITEMS[i];
      if(itemNe.ne === neClean || itemNe.ne_full === neClean || itemNe.ne.slice(-12) === neClean.slice(-12)){
        neObj = itemNe; break;
      }
    }
  }
  if(!neObj){
    bcmsToast('Nota de Empenho ' + neClean + ' não localizada no extrato.');
    return;
  }

  bcmsModalPush(function(){ bcmsDetalheNE(neNum); });
  var hasBack = BCMS_MODAL_STACK.length > 1;

  var sCls = neObj.prova_slug === 'ok' ? 'status-ok' :
            (neObj.prova_slug === 'info' ? 'status-parcial' :
            (neObj.prova_slug === 'warn' ? 'status-canc' :
            (neObj.prova_slug === 'danger' ? 'status-canc' : 'status-zerada')));

  var aLiq = neObj.a_liquidar || 0;
  var celLiq = neObj.cel_liq || 0;
  var celPag = neObj.cel_pag || 0;
  var neLiq = neObj.ne_liq || 0;
  var liqExib = (celLiq > 0) ? celLiq : (neObj.pago || 0);
  var pagExib = (celPag > 0) ? celPag : liqExib;
  var pctCelLiq = neObj.pct_liq || 0;
  var pLiq = pctCelLiq > 0 ? pctCelLiq : (neObj.val > 0 ? Math.min(100, (liqExib / neObj.val) * 100) : 0);

  var h = '<div class="m-accent-bar" style="background:linear-gradient(90deg, #059669 0%, #10B981 50%, #2563EB 100%);"></div>';
  h += '<div class="m-content-wrap">';

  // Cabeçalho
  h += '  <div class="m-header-v2">';
  h += '    <div class="m-header-meta-row" style="padding-right:48px;">';
  h += '      <span class="m-badge-op" style="color:#059669;background:rgba(5,150,105,0.08);border-color:#05966966;">🏛️ UG ' + bcmsEsc(neObj.ug) + (neObj.ug_nome ? ' · ' + bcmsEsc(neObj.ug_nome) : '') + '</span>';
  h += '      <span class="m-badge-status-lg ' + sCls + '">🛡️ ' + bcmsEsc(neObj.prova) + '</span>';
  h += '    </div>';
  h += '    <div class="m-title-row" style="padding-right:48px;">';
  h += '      <div class="m-nc-code-block">';
  h += '        <h3 id="modal-title">Nota de Empenho ' + bcmsEsc(neObj.ne) + '</h3>';
  h += '      </div>';
  h += '      <button type="button" class="m-btn-pill" onclick="bcmsCopiarNC(\'' + bcmsEsc(neObj.ne) + '\', this)" title="Copiar número da NE">';
  h += '        📋 Copiar NE';
  h += '      </button>';
  h += '    </div>';
  h += '    <div class="m-sub-meta">';
  h += '      <span class="m-meta-chip">📅 Emissão: <b>' + bcmsEsc(neObj.dia || '—') + '</b></span>';
  h += '      <span class="m-meta-chip">🏛️ Exercício Financeiro 2026</span>';
  h += '    </div>';
  h += '  </div>';

  // Balanço Financeiro da NE
  h += '  <div class="m-fin-section">';
  h += '    <div class="m-fin-grid">';
  h += '      <div class="m-fin-card">';
  h += '        <span class="m-fin-label">Valor Total Empenhado</span>';
  h += '        <span class="m-fin-val col-emp" style="color:#059669;">' + bcmsBRL(neObj.val) + '</span>';
  h += '        <span class="m-fin-sub">Comprometimento formal da despesa</span>';
  h += '      </div>';
  h += '      <div class="m-fin-card">';
  h += '        <span class="m-fin-label">Crédito a Liquidar</span>';
  h += '        <span class="m-fin-val">' + bcmsBRL(aLiq) + '</span>';
  h += '        <span class="m-fin-sub">Conta contábil 622110000</span>';
  h += '      </div>';
  h += '      <div class="m-fin-card">';
  h += '        <span class="m-fin-label">Liquidado no SIAFI (Célula)</span>';
  h += '        <span class="m-fin-val col-prov" style="color:#059669;">' + bcmsBRL(liqExib) + '</span>';
  h += '        <span class="m-fin-sub">Pago: ' + bcmsBRL(pagExib) + '</span>';
  h += '      </div>';
  h += '      <div class="m-fin-card hero-saldo">';
  h += '        <span class="m-fin-label" style="color:#10B981;">Estágio da Dotação</span>';
  h += '        <span class="m-fin-val-hero">' + pLiq.toFixed(1) + '%</span>';
  h += '        <span class="m-fin-tag-hero">✓ ' + (pLiq >= 100 ? '100% Liquidado' : 'Liquidado no SIAFI') + '</span>';
  h += '      </div>';
  h += '    </div>';
  h += '  </div>';

  if(celLiq > 0){
    h += '  <div class="m-justif-card" style="border-left-color:#10B981;background:rgba(16,185,129,0.06);margin-bottom:16px;">';
    h += '    <div class="m-justif-header">';
    h += '      <span class="m-justif-title" style="color:#059669;">🏛️ Execução Orçamentária no SIAFI · PI ' + bcmsEsc(neObj.pi) + '</span>';
    h += '      <span class="m-badge-status-lg status-ok">✓ ' + (pctCelLiq >= 100 ? '100% LIQUIDADO' : pctCelLiq.toFixed(1) + '% LIQUIDADO') + '</span>';
    h += '    </div>';
    h += '    <div class="m-justif-body" style="font-size:0.875rem;line-height:1.5;color:var(--ink);">';
    h += '      A dotação orçamentária vinculada (<b>PI ' + bcmsEsc(neObj.pi) + ' · ND ' + bcmsEsc(neObj.nd) + '</b>) registra ';
    h += '      <b>' + bcmsBRL(celLiq) + '</b> de despesas liquidadas e <b>' + bcmsBRL(celPag) + '</b> pagas no SIAFI / Tesouro Gerencial.';
    if(neLiq > 0){
      h += ' Esta NE possui redução de saldo a liquidar correspondente a <b>' + bcmsBRL(neLiq) + '</b>.';
    }
    h += '    </div>';
    h += '  </div>';
  }

  // Cards Cadastrais
  h += '  <div class="m-class-grid" style="margin-bottom:16px;">';
  
  // Fornecedor
  h += '    <div class="m-class-card" style="cursor:pointer;" onclick="bcmsDetalheFornecedor(\'' + bcmsEsc(neObj.fav) + '\')" title="Clique para ver todos os empenhos deste fornecedor">';
  h += '      <span class="m-class-label">🏢 Favorecido / Credor (Clique ↗)</span>';
  h += '      <span class="m-class-code" style="font-size:0.875rem;color:var(--primary);">' + bcmsEsc(neObj.fav) + '</span>';
  h += '      <span class="m-class-desc">CNPJ/CPF: <b>' + bcmsEsc(neObj.doc || '—') + '</b></span>';
  h += '    </div>';

  // Processo / Pregão
  if(neObj.proc){
    h += '    <div class="m-class-card" style="cursor:pointer;" onclick="bcmsDetalheProcesso(\'' + bcmsEsc(neObj.proc) + '\')" title="Clique para ver todos os empenhos deste pregão">';
    h += '      <span class="m-class-label">📜 Processo / Pregão (Clique ↗)</span>';
    h += '      <span class="m-class-code" style="font-size:0.875rem;color:#059669;">' + bcmsEsc(neObj.proc) + '</span>';
    h += '      <span class="m-class-desc">Demanda licitatória SISLOG</span>';
    h += '    </div>';
  }

  // NC de Origem
  if(neObj.nc){
    h += '    <div class="m-class-card" style="cursor:pointer;" onclick="bcmsDetalheNC(\'' + bcmsEsc(neObj.nc) + '\')" title="Clique para ver a ficha completa da NC">';
    h += '      <span class="m-class-label">🏢 Nota de Crédito de Origem (Clique ↗)</span>';
    h += '      <span class="m-class-code" style="font-size:0.875rem;color:var(--primary);">' + bcmsEsc(neObj.nc) + '</span>';
    h += '      <span class="m-class-desc">' + (neObj.nc_citada ? 'Citada: ' + bcmsEsc(neObj.nc_citada) : 'Célula SIAFI vinculada') + '</span>';
    h += '    </div>';
  }

  // Célula SIAFI
  h += '    <div class="m-class-card">';
  h += '      <span class="m-class-label">⚙️ Célula Orçamentária SIAFI</span>';
  h += '      <span class="m-class-code" style="font-size:0.875rem;">ND ' + bcmsEsc(neObj.nd) + ' · PI ' + bcmsEsc(neObj.pi) + '</span>';
  h += '      <span class="m-class-desc">' + bcmsEsc(neObj.nd_desc || 'Despesa') + '</span>';
  h += '    </div>';

  h += '  </div>';

  // Laudo Prova Real
  var provaBorder = neObj.prova_slug === 'ok' ? '#10B981' : (neObj.prova_slug === 'info' ? '#3B82F6' : '#EF4444');
  var provaBg = neObj.prova_slug === 'ok' ? 'rgba(16,185,129,0.06)' : (neObj.prova_slug === 'info' ? 'rgba(59,130,246,0.06)' : 'rgba(239,68,68,0.08)');
  h += '  <div class="m-justif-card" style="border-left-color:' + provaBorder + ';background:' + provaBg + ';margin-bottom:16px;">';
  h += '    <div class="m-justif-header">';
  h += '      <span class="m-justif-title" style="color:' + provaBorder + ';">🛡️ Laudo de Auditoria: Prova Real SIAFI</span>';
  h += '    </div>';
  if(neObj.motivo){
    h += '    <p style="margin:0;font-size:0.8125rem;line-height:1.5;color:var(--danger);font-weight:600;">⚠️ Inconsistência Detectada: ' + bcmsEsc(neObj.motivo) + '</p>';
    h += '    <p style="margin:6px 0 0 0;font-size:0.75rem;color:var(--ink-muted);">A dotação orçamentária efetivamente debitada no SIAFI pertence à célula da unidade, divergindo do texto digitado pelo operador no campo Observação.</p>';
  } else {
    h += '    <p style="margin:0;font-size:0.8125rem;line-height:1.5;color:var(--ink);">✓ Lastro orçamentário integralmente auditado e confirmado entre a dotação SIAFI (UG + PI + ND) e a citação documental da despesa.</p>';
  }
  h += '  </div>';

  // Observação Completa
  h += '  <div class="m-justif-card" style="margin-bottom:20px;">';
  h += '    <div class="m-justif-header">';
  h += '      <span class="m-justif-title">📝 Histórico / Observação da NE (Texto Integral SIAFI)</span>';
  h += '      <button type="button" class="m-btn-pill" data-copy="' + bcmsEsc(neObj.desc || '') + '" onclick="bcmsCopiarTexto(this)">📋 Copiar Observação</button>';
  h += '    </div>';
  h += '    <div class="m-justif-body" style="font-size:0.8125rem;line-height:1.6;max-height:220px;overflow-y:auto;white-space:pre-wrap;">' + bcmsEsc(neObj.desc || 'Sem observação registrada.') + '</div>';
  h += '  </div>';

  // Rodapé
  h += '  <div class="m-footer-actions-v2">';
  if(hasBack){
    h += '    <button type="button" class="m-btn-pill" onclick="bcmsModalBack()">‹ Voltar</button>';
  }
  h += '    <button type="button" class="m-btn-pill" onclick="bcmsCopiarNC(\'' + bcmsEsc(neObj.ne) + '\', this)">📋 Copiar Nº da NE</button>';
  h += '    <button type="button" class="m-btn-pill primary" onclick="bcmsCelClose()">Fechar Janela ✕</button>';
  h += '  </div>';

  h += '</div>';

  document.getElementById('modal-body').innerHTML = h;
  var m = document.getElementById('modal');
  m.classList.add('open');
  m.setAttribute('aria-hidden', 'false');
  var x = document.querySelector('.modal-x');
  if(x) x.focus();
}

function bcmsDetalheFornecedor(favNome){
  if(!favNome) return;
  var fClean = String(favNome).trim();
  var nesDoForn = [];
  var totVal = 0;
  var docForn = '';
  var ugsSet = {};

  if(typeof RASTREIO_NE_ITEMS !== 'undefined' && RASTREIO_NE_ITEMS){
    for(var i = 0; i < RASTREIO_NE_ITEMS.length; i++){
      var it = RASTREIO_NE_ITEMS[i];
      if(it.fav === fClean){
        nesDoForn.push(it);
        totVal += it.val;
        if(it.doc && !docForn) docForn = it.doc;
        if(it.ug) ugsSet[it.ug] = (ugsSet[it.ug] || 0) + 1;
      }
    }
  }

  if(nesDoForn.length === 0){
    bcmsToast('Nenhum empenho encontrado para o credor ' + fClean);
    return;
  }

  bcmsModalPush(function(){ bcmsDetalheFornecedor(favNome); });
  var hasBack = BCMS_MODAL_STACK.length > 1;
  var ugsList = Object.keys(ugsSet).join(', ');

  var h = '<div class="m-accent-bar" style="background:linear-gradient(90deg, #8B5CF6 0%, #6366F1 50%, #059669 100%);"></div>';
  h += '<div class="m-content-wrap">';

  // Cabeçalho
  h += '  <div class="m-header-v2">';
  h += '    <div class="m-header-meta-row" style="padding-right:48px;">';
  h += '      <span class="m-badge-op" style="color:#8B5CF6;background:rgba(139,92,246,0.08);border-color:#8B5CF666;">🏢 CADASTRO DE FORNECEDOR / CREDOR</span>';
  h += '      <span class="m-badge-status-lg status-ok">' + nesDoForn.length + ' EMPENHOS EMITIDOS</span>';
  h += '    </div>';
  h += '    <div class="m-title-row" style="padding-right:48px;">';
  h += '      <div class="m-nc-code-block">';
  h += '        <h3 id="modal-title">' + bcmsEsc(fClean) + '</h3>';
  h += '      </div>';
  h += '      <button type="button" class="m-btn-pill" onclick="bcmsCopiarNC(\'' + bcmsEsc(docForn) + '\', this)" title="Copiar CNPJ/CPF">';
  h += '        📋 Copiar CNPJ/CPF';
  h += '      </button>';
  h += '    </div>';
  h += '    <div class="m-sub-meta">';
  h += '      <span class="m-meta-chip">📄 Documento: <b>' + bcmsEsc(docForn || 'Não Informado') + '</b></span>';
  h += '      <span class="m-meta-chip">🏛️ UGs Contratantes: <b>' + bcmsEsc(ugsList) + '</b></span>';
  h += '    </div>';
  h += '  </div>';

  // Balanço Financeiro
  h += '  <div class="m-fin-section">';
  h += '    <div class="m-fin-grid">';
  h += '      <div class="m-fin-card hero-saldo">';
  h += '        <span class="m-fin-label" style="color:#10B981;">Total Empenhado no Exercício</span>';
  h += '        <span class="m-fin-val-hero">' + bcmsBRL(totVal) + '</span>';
  h += '        <span class="m-fin-tag-hero">✓ Volume financeiro acumulado</span>';
  h += '      </div>';
  h += '      <div class="m-fin-card">';
  h += '        <span class="m-fin-label">Quantidade de NEs</span>';
  h += '        <span class="m-fin-val">' + nesDoForn.length + '</span>';
  h += '        <span class="m-fin-sub">Notas de empenho emitidas</span>';
  h += '      </div>';
  h += '      <div class="m-fin-card">';
  h += '        <span class="m-fin-label">Média por Empenho</span>';
  h += '        <span class="m-fin-val col-emp">' + bcmsBRL(totVal / nesDoForn.length) + '</span>';
  h += '        <span class="m-fin-sub">Ticket médio da contratação</span>';
  h += '      </div>';
  h += '    </div>';
  h += '  </div>';

  // Tabela de NEs emitidas para este fornecedor
  h += '  <div class="m-justif-card" style="border-left-color:#8B5CF6;margin-top:20px;">';
  h += '    <div class="m-justif-header">';
  h += '      <span class="m-justif-title">📦 Relação de Notas de Empenho Emitidas (' + nesDoForn.length + ')</span>';
  h += '    </div>';
  h += '    <div class="tbl-scroll" style="max-height:360px;"><table class="det det-compact"><thead><tr>';
  h += '      <th>Nota de Empenho</th><th>Emissão</th><th>UG</th><th>Processo / Pregão</th><th class="num">Valor da NE</th><th>ND</th><th>NC Origem</th><th style="text-align:center;">Ação</th>';
  h += '    </tr></thead><tbody>';

  for(var k = 0; k < nesDoForn.length; k++){
    var nIt = nesDoForn[k];
    h += '<tr class="cel-row">';
    h += '  <td class="font-mono font-bold"><a href="javascript:void(0)" onclick="bcmsDetalheNE(\'' + bcmsEsc(nIt.ne) + '\')" class="link-drill">' + bcmsEsc(nIt.ne) + '</a></td>';
    h += '  <td class="mono2">' + bcmsEsc(nIt.dia || '—') + '</td>';
    h += '  <td><span class="ug-pill fav">' + bcmsEsc(nIt.ug) + '</span></td>';
    h += '  <td>' + (nIt.proc ? '<a href="javascript:void(0)" onclick="bcmsDetalheProcesso(\'' + bcmsEsc(nIt.proc) + '\')" class="badge-pregao" style="cursor:pointer;">' + bcmsEsc(nIt.proc) + '</a>' : '—') + '</td>';
    h += '  <td class="num font-mono font-bold" style="color:var(--success-strong);">' + bcmsBRL(nIt.val) + '</td>';
    h += '  <td class="mono2">' + bcmsEsc(nIt.nd) + '</td>';
    h += '  <td class="mono2">' + (nIt.nc ? '<a href="javascript:void(0)" onclick="bcmsDetalheNC(\'' + bcmsEsc(nIt.nc) + '\')" class="link-drill">' + bcmsEsc(nIt.nc) + '</a>' : '—') + '</td>';
    h += '  <td style="text-align:center;"><button type="button" class="tbl-action-btn" onclick="bcmsDetalheNE(\'' + bcmsEsc(nIt.ne) + '\')">Detalhar ↗</button></td>';
    h += '</tr>';
  }

  h += '    </tbody></table></div>';
  h += '  </div>';

  // Rodapé
  h += '  <div class="m-footer-actions-v2">';
  if(hasBack){
    h += '    <button type="button" class="m-btn-pill" onclick="bcmsModalBack()">‹ Voltar</button>';
  }
  h += '    <button type="button" class="m-btn-pill" onclick="bcmsCopiarNC(\'' + bcmsEsc(docForn) + '\', this)">📋 Copiar CNPJ/CPF</button>';
  h += '    <button type="button" class="m-btn-pill primary" onclick="bcmsCelClose()">Fechar Janela ✕</button>';
  h += '  </div>';

  h += '</div>';

  document.getElementById('modal-body').innerHTML = h;
  var m = document.getElementById('modal');
  m.classList.add('open');
  m.setAttribute('aria-hidden', 'false');
  var x = document.querySelector('.modal-x');
  if(x) x.focus();
}

function bcmsDetalheProcesso(procStr){
  if(!procStr) return;
  var pClean = String(procStr).trim();
  var nesDoProc = [];
  var totVal = 0;
  var fornsSet = {};

  if(typeof RASTREIO_NE_ITEMS !== 'undefined' && RASTREIO_NE_ITEMS){
    for(var i = 0; i < RASTREIO_NE_ITEMS.length; i++){
      var it = RASTREIO_NE_ITEMS[i];
      if(it.proc === pClean){
        nesDoProc.push(it);
        totVal += it.val;
        if(it.fav && it.fav !== 'NAO SE APLICA') fornsSet[it.fav] = (fornsSet[it.fav] || 0) + it.val;
      }
    }
  }

  if(nesDoProc.length === 0){
    bcmsToast('Nenhum empenho encontrado para o processo ' + pClean);
    return;
  }

  bcmsModalPush(function(){ bcmsDetalheProcesso(procStr); });
  var hasBack = BCMS_MODAL_STACK.length > 1;

  var h = '<div class="m-accent-bar" style="background:linear-gradient(90deg, #F59E0B 0%, #059669 50%, #2563EB 100%);"></div>';
  h += '<div class="m-content-wrap">';

  // Cabeçalho
  h += '  <div class="m-header-v2">';
  h += '    <div class="m-header-meta-row" style="padding-right:48px;">';
  h += '      <span class="m-badge-op" style="color:#F59E0B;background:rgba(245,158,11,0.08);border-color:#F59E0B66;">📜 PROCESSO LICITATÓRIO / SISLOG</span>';
  h += '      <span class="m-badge-status-lg status-ok">' + nesDoProc.length + ' EMPENHOS VINCULADOS</span>';
  h += '    </div>';
  h += '    <div class="m-title-row" style="padding-right:48px;">';
  h += '      <div class="m-nc-code-block">';
  h += '        <h3 id="modal-title">' + bcmsEsc(pClean) + '</h3>';
  h += '      </div>';
  h += '    </div>';
  h += '    <div class="m-sub-meta">';
  h += '      <span class="m-meta-chip">🏢 Credores Atendidos: <b>' + Object.keys(fornsSet).length + ' fornecedores</b></span>';
  h += '    </div>';
  h += '  </div>';

  // Balanço Financeiro
  h += '  <div class="m-fin-section">';
  h += '    <div class="m-fin-grid">';
  h += '      <div class="m-fin-card hero-saldo">';
  h += '        <span class="m-fin-label" style="color:#10B981;">Total Empenhado no Processo</span>';
  h += '        <span class="m-fin-val-hero">' + bcmsBRL(totVal) + '</span>';
  h += '        <span class="m-fin-tag-hero">✓ Execução formal acumulada</span>';
  h += '      </div>';
  h += '      <div class="m-fin-card">';
  h += '        <span class="m-fin-label">NEs Emitidas</span>';
  h += '        <span class="m-fin-val">' + nesDoProc.length + '</span>';
  h += '        <span class="m-fin-sub">Contratações vinculadas</span>';
  h += '      </div>';
  h += '      <div class="m-fin-card">';
  h += '        <span class="m-fin-label">Fornecedores Distintos</span>';
  h += '        <span class="m-fin-val col-emp">' + Object.keys(fornsSet).length + '</span>';
  h += '        <span class="m-fin-sub">Pessoas jurídicas contratadas</span>';
  h += '      </div>';
  h += '    </div>';
  h += '  </div>';

  // Tabela de NEs
  h += '  <div class="m-justif-card" style="border-left-color:#F59E0B;margin-top:20px;">';
  h += '    <div class="m-justif-header">';
  h += '      <span class="m-justif-title">📦 Relação de Notas de Empenho Vinculadas (' + nesDoProc.length + ')</span>';
  h += '    </div>';
  h += '    <div class="tbl-scroll" style="max-height:360px;"><table class="det det-compact"><thead><tr>';
  h += '      <th>Nota de Empenho</th><th>Emissão</th><th>UG</th><th>Favorecido / Fornecedor</th><th class="num">Valor da NE</th><th>ND</th><th>NC Origem</th><th style="text-align:center;">Ação</th>';
  h += '    </tr></thead><tbody>';

  for(var k = 0; k < nesDoProc.length; k++){
    var nIt = nesDoProc[k];
    h += '<tr class="cel-row">';
    h += '  <td class="font-mono font-bold"><a href="javascript:void(0)" onclick="bcmsDetalheNE(\'' + bcmsEsc(nIt.ne) + '\')" class="link-drill">' + bcmsEsc(nIt.ne) + '</a></td>';
    h += '  <td class="mono2">' + bcmsEsc(nIt.dia || '—') + '</td>';
    h += '  <td><span class="ug-pill fav">' + bcmsEsc(nIt.ug) + '</span></td>';
    h += '  <td><a href="javascript:void(0)" onclick="bcmsDetalheFornecedor(\'' + bcmsEsc(nIt.fav) + '\')" class="link-drill" style="color:var(--ink);">' + bcmsEsc(nIt.fav) + '</a></td>';
    h += '  <td class="num font-mono font-bold" style="color:var(--success-strong);">' + bcmsBRL(nIt.val) + '</td>';
    h += '  <td class="mono2">' + bcmsEsc(nIt.nd) + '</td>';
    h += '  <td class="mono2">' + (nIt.nc ? '<a href="javascript:void(0)" onclick="bcmsDetalheNC(\'' + bcmsEsc(nIt.nc) + '\')" class="link-drill">' + bcmsEsc(nIt.nc) + '</a>' : '—') + '</td>';
    h += '  <td style="text-align:center;"><button type="button" class="tbl-action-btn" onclick="bcmsDetalheNE(\'' + bcmsEsc(nIt.ne) + '\')">Detalhar ↗</button></td>';
    h += '</tr>';
  }

  h += '    </tbody></table></div>';
  h += '  </div>';

  // Rodapé
  h += '  <div class="m-footer-actions-v2">';
  if(hasBack){
    h += '    <button type="button" class="m-btn-pill" onclick="bcmsModalBack()">‹ Voltar</button>';
  }
  h += '    <button type="button" class="m-btn-pill primary" onclick="bcmsCelClose()">Fechar Janela ✕</button>';
  h += '  </div>';

  h += '</div>';

  document.getElementById('modal-body').innerHTML = h;
  var m = document.getElementById('modal');
  m.classList.add('open');
  m.setAttribute('aria-hidden', 'false');
  var x = document.querySelector('.modal-x');
  if(x) x.focus();
}

/* ================= EXPORTAÇÃO EXCEL ================= */
function bcmsExportRastreioExcel(tipo){
  var list = (tipo === 'nc') ? RASTREIO_NC_FILTRADOS : RASTREIO_NE_FILTRADOS;
  if(!list || list.length === 0){
    bcmsToast('⚠ Nenhum registro para exportar.');
    return;
  }

  var xml = '<?xml version="1.0" encoding="UTF-8"?>\n' +
    '<?mso-application progid="Excel.Sheet"?>\n' +
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"\n' +
    ' xmlns:o="urn:schemas-microsoft-com:office:office"\n' +
    ' xmlns:x="urn:schemas-microsoft-com:office:excel"\n' +
    ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">\n' +
    '<Styles>\n' +
    '  <Style ss:ID="sHdr"><Font ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#059669" ss:Pattern="Solid"/></Style>\n' +
    '  <Style ss:ID="sCur"><NumberFormat ss:Format="R$ #,##0.00"/></Style>\n' +
    '</Styles>\n' +
    '<Worksheet ss:Name="' + (tipo === 'nc' ? 'Credito_Empenho' : 'Empenhos_Emitidos') + '">\n' +
    '<Table>\n';

  if(tipo === 'nc'){
    xml += '<Row ss:StyleID="sHdr">' +
      '<Cell><Data ss:Type="String">Nota de Crédito</Data></Cell>' +
      '<Cell><Data ss:Type="String">Emissão</Data></Cell>' +
      '<Cell><Data ss:Type="String">Órgão Concedente</Data></Cell>' +
      '<Cell><Data ss:Type="String">OMDS</Data></Cell>' +
      '<Cell><Data ss:Type="String">Plano Interno</Data></Cell>' +
      '<Cell><Data ss:Type="String">Natureza Despesa</Data></Cell>' +
      '<Cell><Data ss:Type="String">Crédito Recebido</Data></Cell>' +
      '<Cell><Data ss:Type="String">Total Empenhado</Data></Cell>' +
      '<Cell><Data ss:Type="String">Saldo Disponível</Data></Cell>' +
      '<Cell><Data ss:Type="String">Taxa Queima (%)</Data></Cell>' +
      '<Cell><Data ss:Type="String">Qtd NEs</Data></Cell>' +
    '</Row>\n';

    for(var i = 0; i < list.length; i++){
      var it = list[i];
      xml += '<Row>' +
        '<Cell><Data ss:Type="String">' + bcmsEscXml(it.nc) + '</Data></Cell>' +
        '<Cell><Data ss:Type="String">' + bcmsEscXml(it.dia) + '</Data></Cell>' +
        '<Cell><Data ss:Type="String">' + bcmsEscXml(it.emit_nome || it.emit) + '</Data></Cell>' +
        '<Cell><Data ss:Type="String">' + bcmsEscXml(it.om_sigla) + '</Data></Cell>' +
        '<Cell><Data ss:Type="String">' + bcmsEscXml(it.pi) + '</Data></Cell>' +
        '<Cell><Data ss:Type="String">' + bcmsEscXml(it.nd) + '</Data></Cell>' +
        '<Cell ss:StyleID="sCur"><Data ss:Type="Number">' + (it.prov || 0).toFixed(2) + '</Data></Cell>' +
        '<Cell ss:StyleID="sCur"><Data ss:Type="Number">' + (it.emp || 0).toFixed(2) + '</Data></Cell>' +
        '<Cell ss:StyleID="sCur"><Data ss:Type="Number">' + (it.cred || 0).toFixed(2) + '</Data></Cell>' +
        '<Cell><Data ss:Type="Number">' + (it.taxa_queima || 0).toFixed(1) + '</Data></Cell>' +
        '<Cell><Data ss:Type="Number">' + (it.nes ? it.nes.length : 0) + '</Data></Cell>' +
      '</Row>\n';
    }
  } else {
    xml += '<Row ss:StyleID="sHdr">' +
      '<Cell><Data ss:Type="String">Nota de Empenho</Data></Cell>' +
      '<Cell><Data ss:Type="String">Data Emissão</Data></Cell>' +
      '<Cell><Data ss:Type="String">UG Executora</Data></Cell>' +
      '<Cell><Data ss:Type="String">Favorecido / Fornecedor</Data></Cell>' +
      '<Cell><Data ss:Type="String">CNPJ / CPF</Data></Cell>' +
      '<Cell><Data ss:Type="String">Processo / Pregão</Data></Cell>' +
      '<Cell><Data ss:Type="String">Valor Empenhado</Data></Cell>' +
      '<Cell><Data ss:Type="String">Natureza Despesa</Data></Cell>' +
      '<Cell><Data ss:Type="String">Plano Interno</Data></Cell>' +
      '<Cell><Data ss:Type="String">NC de Origem</Data></Cell>' +
      '<Cell><Data ss:Type="String">Selo Prova Real</Data></Cell>' +
    '</Row>\n';

    for(var j = 0; j < list.length; j++){
      var ne = list[j];
      xml += '<Row>' +
        '<Cell><Data ss:Type="String">' + bcmsEscXml(ne.ne) + '</Data></Cell>' +
        '<Cell><Data ss:Type="String">' + bcmsEscXml(ne.dia) + '</Data></Cell>' +
        '<Cell><Data ss:Type="String">' + bcmsEscXml(ne.ug) + '</Data></Cell>' +
        '<Cell><Data ss:Type="String">' + bcmsEscXml(ne.fav) + '</Data></Cell>' +
        '<Cell><Data ss:Type="String">' + bcmsEscXml(ne.doc) + '</Data></Cell>' +
        '<Cell><Data ss:Type="String">' + bcmsEscXml(ne.proc) + '</Data></Cell>' +
        '<Cell ss:StyleID="sCur"><Data ss:Type="Number">' + (ne.val || 0).toFixed(2) + '</Data></Cell>' +
        '<Cell><Data ss:Type="String">' + bcmsEscXml(ne.nd) + '</Data></Cell>' +
        '<Cell><Data ss:Type="String">' + bcmsEscXml(ne.pi) + '</Data></Cell>' +
        '<Cell><Data ss:Type="String">' + bcmsEscXml(ne.nc) + '</Data></Cell>' +
        '<Cell><Data ss:Type="String">' + bcmsEscXml(ne.prova) + '</Data></Cell>' +
      '</Row>\n';
    }
  }

  xml += '</Table></Worksheet></Workbook>';

  var blob = new Blob([xml], { type: 'application/vnd.ms-excel;charset=utf-8;' });
  var url = URL.createObjectURL(blob);
  var link = document.createElement('a');
  link.setAttribute('href', url);
  link.setAttribute('download', 'SICOF_' + (tipo === 'nc' ? 'Credito_Empenho' : 'Empenhos') + '_BaApLog.xls');
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  bcmsToast('📊 Planilha exportada com sucesso (' + list.length + ' linhas)!');
}
"""
