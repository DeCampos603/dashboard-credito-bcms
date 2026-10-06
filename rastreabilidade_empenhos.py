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
    """Baixa o extrato de empenhos emitidos publicado no Google Sheets / Drive."""
    target_url = url or DEFAULT_EMPENHOS_URL
    req = urllib.request.Request(target_url, headers={"User-Agent": "Mozilla/5.0 (dashboard-bcms)"})
    tmp = os.path.join(tempfile.gettempdir(), "empenhos_baaplog_download.xlsx")
    with urllib.request.urlopen(req, timeout=120) as r, open(tmp, "wb") as f:
        f.write(r.read())
    if os.path.getsize(tmp) < 1000:
        raise SystemExit("Download de empenhos muito pequeno — verifique o link publicado.")
    return tmp

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

    # 1. Mapeamento do Acervo de Créditos para a PROVA DA VERDADE
    cred_audit_db = {} # nc_code -> list of dicts
    cel_to_ncs = {}    # (ug, pi, nd) -> list of nc_codes

    for cod, d in res_credito.items():
        for L in d.get("linhas", []):
            nc_raw = L.get("nc")
            if not nc_raw: continue
            m = nc_pat.search(nc_raw)
            if not m: continue
            nc_code = m.group(1).upper()
            cred_audit_db.setdefault(nc_code, []).append({
                "ug": cod,
                "emit": L.get("emit") or "",
                "emit_nome": L.get("emit_nome") or "",
                "pi": L.get("pi") or "",
                "nd": L.get("nd") or "",
                "prov": L.get("prov", 0.0) or L.get("cred", 0.0),
                "obj": L.get("obj") or ""
            })
            k_cel = (cod, L.get("pi"), L.get("nd"))
            cel_to_ncs.setdefault(k_cel, []).append(nc_code)
            # Elemento genérico (ex: 339000 para 339039)
            nd_elem = (L.get("nd") or "")[:4] + "00"
            k_elem = (cod, L.get("pi"), nd_elem)
            cel_to_ncs.setdefault(k_elem, []).append(nc_code)

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
        try: a_liquidar = float(str(r[15] or 0).replace("'", "").strip())
        except Exception: a_liquidar = 0.0
        try: liquidado_pago = float(str(r[21] or 0).replace("'", "").strip())
        except Exception: liquidado_pago = 0.0

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
                "dia": dia,
                "ug": ug_cur,
                "fav": fav_nome,
                "val": round(val, 2),
                "val_fmt": _fmt_brl(val),
                "pi_ne": pi,
                "nd_ne": nd,
                "nc_citada": nc_citada,
                "status_prova": status_prova,
                "status_slug": status_slug,
                "motivo": motivo_divergencia,
                "desc": desc[:100]
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
            "proc": proc_str,
            "desc": desc[:150],
            "a_liquidar": round(a_liquidar, 2),
            "pago": round(liquidado_pago, 2)
        }
        nes_list.append(ne_item)
        tot_val += val

        # Indexa por NC para drill-down e relacionamentos
        if nc_efetiva:
            nc_to_nes.setdefault(nc_efetiva, []).append({
                "ne": ne_num,
                "dia": dia,
                "fav": fav_nome,
                "doc": fav_doc,
                "val": round(val, 2),
                "proc": proc_str,
                "prova": status_prova,
                "prova_slug": status_slug,
                "desc": desc[:90]
            })

        # Agrega fornecedores
        if fav_nome and fav_nome != "NAO SE APLICA":
            top_forn[fav_nome] = top_forn.get(fav_nome, 0.0) + val

        # Agrega processos
        if proc_str:
            top_procs[proc_str] = top_procs.get(proc_str, 0.0) + val

    # Top fornecedores
    sorted_forn = sorted(top_forn.items(), key=lambda x: x[1], reverse=True)[:25]
    top_forn_list = [{"nome": f, "val": round(v, 2), "fmt": _fmt_brl(v)} for f, v in sorted_forn]

    # Top processos
    sorted_procs = sorted(top_procs.items(), key=lambda x: x[1], reverse=True)[:20]
    top_procs_list = [{"proc": p, "val": round(v, 2), "fmt": _fmt_brl(v)} for p, v in sorted_procs]

    return {
        "total_val": round(tot_val, 2),
        "total_qtd": len(nes_list),
        "qtd_fav": len(top_forn),
        "estat_prova": estat_prova,
        "inconsistencias": inconsistencias_citacao,
        "top_fornecedores": top_forn_list,
        "top_processos": top_procs_list,
        "nc_to_nes": nc_to_nes,
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
            <th style="width:44px;text-align:center;">NEs</th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNC('nc')" title="Ordenar por Nota de Crédito">Nota de Crédito <span class="sort" id="sort-nc-nc"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNC('dia')" title="Ordenar por Data">Recebido em <span class="sort" id="sort-nc-dia"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNC('emit_nome')" title="Ordenar por Órgão Concedente">Órgão Concedente <span class="sort" id="sort-nc-emit"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNC('om_sigla')" title="Ordenar por OM Favorecida">Unidade (OMDS) <span class="sort" id="sort-nc-uasg"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNC('pi')" title="Ordenar por Plano Interno">Plano Interno <span class="sort" id="sort-nc-pi"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNC('nd')" title="Ordenar por Natureza de Despesa">ND <span class="sort" id="sort-nc-nd"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNC('prov')" class="num" title="Ordenar por Provisão Recebida">Crédito Recebido <span class="sort" id="sort-nc-prov">▼</span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNC('emp')" class="num" title="Ordenar por Total Empenhado">Total Empenhado <span class="sort" id="sort-nc-emp"></span></th>
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
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNE('nd')" title="Ordenar por Natureza de Despesa">ND <span class="sort" id="sort-ne-nd"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNE('pi')" title="Ordenar por Plano Interno">Plano Interno <span class="sort" id="sort-ne-pi"></span></th>
            <th tabindex="0" role="button" onclick="bcmsSortRastreioNE('nc')" title="Ordenar por NC Citada">NC de Origem <span class="sort" id="sort-ne-nc"></span></th>
            <th style="text-align:center;">Selo Prova Real</th>
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
    <div class="grid2" style="margin-top:10px;">
      <div class="card" style="padding:24px;">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
          <h3 style="font-family:var(--serif);font-size:1.15rem;font-weight:700;color:var(--ink);margin:0;">🏢 Maiores Fornecedores por Volume Empenhado</h3>
          <span class="pill-fonte" style="font-weight:700;">Top 25 Credores</span>
        </div>
        <p style="font-size:0.8125rem;color:var(--ink-muted);margin:0 0 18px 0;line-height:1.5;">Concentração de despesa empenhada no âmbito da Base de Apoio Logístico do Exército em 2026.</p>
        <div class="ranking-lista" id="listaTopFornecedores"></div>
      </div>

      <div class="card" style="padding:24px;">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
          <h3 style="font-family:var(--serif);font-size:1.15rem;font-weight:700;color:var(--ink);margin:0;">📜 Principais Pregões & Processos Licitatórios</h3>
          <span class="pill-fonte" style="font-weight:700;">Demandas SISLOG</span>
        </div>
        <p style="font-size:0.8125rem;color:var(--ink-muted);margin:0 0 18px 0;line-height:1.5;">Processos administrativos e pregões eletrônicos mais referenciados nas emissões de empenho.</p>
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
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
        <h3 style="font-family:var(--serif);font-size:1.15rem;font-weight:700;color:var(--ink);margin:0;">📋 Registro de Inconsistências Detectadas na Observação ({qtd_alertas} ocorrências)</h3>
        <span class="pill-status status-canc" style="font-weight:700;">Auditoria Preventiva</span>
      </div>
      <p style="font-size:0.8125rem;color:var(--ink-muted);margin:0 0 16px 0;">Relação de Notas de Empenho onde a NC citada no texto diverge da célula orçamentária autêntica debitada no SIAFI. Recomendada retificação administrativa.</p>
      
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
  padding: 2px 7px;
  font-size: 0.6875rem;
  font-weight: 700;
  font-family: var(--mono);
  border-radius: 6px;
  background: var(--bg-subtle);
  border: 1px solid var(--border);
  color: var(--ink-muted);
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
  transition: transform .2s var(--ease-spring), border-color .2s ease;
}
.rank-item:hover {
  transform: translateX(4px);
  border-color: var(--primary-600);
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

function bcmsInitRastreio(){
  if(RASTREIO_INITIALIZED) return;
  if(typeof EMPENHODATA === 'undefined' || !EMPENHODATA) return;

  // 1. Constrói RASTREIO_NC_ITEMS enriquecendo as NCs do HISTDATA com as NEs
  var allNcs = [];
  if(typeof HISTDATA !== 'undefined' && HISTDATA && HISTDATA.items){
    allNcs = HISTDATA.items;
  }
  
  for(var i = 0; i < allNcs.length; i++){
    var it = allNcs[i];
    var nesVinculadas = (EMPENHODATA.nc_to_nes && EMPENHODATA.nc_to_nes[it.nc]) ? EMPENHODATA.nc_to_nes[it.nc] : [];
    var totEmpNE = 0;
    for(var j = 0; j < nesVinculadas.length; j++){
      totEmpNE += nesVinculadas[j].val;
    }
    
    var provVal = it.prov || 0;
    var taxaQueima = provVal > 0 ? (totEmpNE / provVal * 100) : 0;
    var faixa = 'zero';
    if(totEmpNE > 0){
      if(taxaQueima >= 80) faixa = 'alta';
      else if(taxaQueima >= 30) faixa = 'media';
      else faixa = 'baixa';
    }

    var ncObj = {
      hid: it.hid,
      nc: it.nc,
      dia: it.dia || '',
      emit: it.emit || '',
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
      emp: totEmpNE,
      cred: Math.max(0, provVal - totEmpNE),
      taxa_queima: taxaQueima,
      faixa_queima: faixa,
      nes: nesVinculadas
    };
    RASTREIO_NC_ITEMS.push(ncObj);
  }

  // 2. Base Geral de Empenhos
  RASTREIO_NE_ITEMS = EMPENHODATA.nes || [];

  // 3. Renderiza Top Fornecedores
  var elForn = document.getElementById('listaTopFornecedores');
  if(elForn && EMPENHODATA.top_fornecedores){
    var htmlF = '';
    var topF = EMPENHODATA.top_fornecedores.slice(0, 15);
    var maxValF = topF.length > 0 ? topF[0].val : 1;
    for(var fIdx = 0; fIdx < topF.length; fIdx++){
      var itemF = topF[fIdx];
      var pctF = Math.min(100, Math.round(itemF.val / maxValF * 100));
      htmlF += '<div class="rank-item">' +
        '<div class="rank-info">' +
          '<span class="rank-pos">#' + (fIdx + 1) + '</span>' +
          '<span class="rank-nome" title="' + bcmsEsc(itemF.nome) + '">' + bcmsEsc(itemF.nome) + '</span>' +
          '<div class="kpi-bar-wrap"><div class="kpi-bar-fill" style="width:' + pctF + '%"></div></div>' +
        '</div>' +
        '<div class="rank-val">' + bcmsEsc(itemF.fmt) + '</div>' +
      '</div>';
    }
    elForn.innerHTML = htmlF;
  }

  // 4. Renderiza Top Processos
  var elProc = document.getElementById('listaTopProcessos');
  if(elProc && EMPENHODATA.top_processos){
    var htmlP = '';
    var topP = EMPENHODATA.top_processos.slice(0, 15);
    var maxValP = topP.length > 0 ? topP[0].val : 1;
    for(var pIdx = 0; pIdx < topP.length; pIdx++){
      var itemP = topP[pIdx];
      var pctP = Math.min(100, Math.round(itemP.val / maxValP * 100));
      htmlP += '<div class="rank-item">' +
        '<div class="rank-info">' +
          '<span class="rank-pos">#' + (pIdx + 1) + '</span>' +
          '<span class="rank-nome">' + bcmsEsc(itemP.proc) + '</span>' +
          '<div class="kpi-bar-wrap"><div class="kpi-bar-fill" style="width:' + pctP + '%"></div></div>' +
        '</div>' +
        '<div class="rank-val">' + bcmsEsc(itemP.fmt) + '</div>' +
      '</div>';
    }
    elProc.innerHTML = htmlP;
  }

  // 5. Renderiza Inconsistências de Citação (Sub-aba Prova Real)
  var elInc = document.getElementById('corpoInconsistencias');
  if(elInc && EMPENHODATA.inconsistencias){
    var incs = EMPENHODATA.inconsistencias;
    var htmlInc = '';
    if(incs.length === 0){
      htmlInc = '<tr><td colspan="9" style="text-align:center;padding:24px;color:var(--ink-muted);">Nenhuma inconsistência de citação detectada. Todas as NEs possuem vínculo contábil comprovado.</td></tr>';
    } else {
      for(var iI = 0; iI < incs.length; iI++){
        var inItem = incs[iI];
        var sCls = inItem.status_slug === 'ok' ? 'status-disp' :
                  (inItem.status_slug === 'info' ? 'status-parcial' :
                  (inItem.status_slug === 'warn' ? 'status-canc' :
                  (inItem.status_slug === 'danger' ? 'status-canc' : 'status-zerada')));

        htmlInc += '<tr class="cel-row">' +
          '<td class="mono2 font-bold">' + bcmsEsc(inItem.ne) + '</td>' +
          '<td class="mono2">' + bcmsEsc(inItem.dia) + '</td>' +
          '<td><span class="ug-pill fav">' + bcmsEsc(inItem.ug) + '</span></td>' +
          '<td><b>' + bcmsEsc(inItem.fav) + '</b></td>' +
          '<td class="num font-mono font-bold" style="color:var(--success-strong);">' + bcmsEsc(inItem.val_fmt) + '</td>' +
          '<td class="mono2">' + bcmsEsc(inItem.pi_ne) + '</td>' +
          '<td class="mono2">' + bcmsEsc(inItem.nd_ne) + '</td>' +
          '<td class="mono2"><b style="color:var(--danger);">' + bcmsEsc(inItem.nc_citada) + '</b></td>' +
          '<td><span class="pill-status ' + sCls + '">' + bcmsEsc(inItem.status_prova) + '</span><br><small style="color:var(--ink-muted);font-size:0.75rem;">' + bcmsEsc(inItem.motivo) + '</small></td>' +
        '</tr>';
      }
    }
    elInc.innerHTML = htmlInc;
  }

  // 6. Renderiza Créditos Parados (Risco)
  var elRisco = document.getElementById('corpoRastreioRisco');
  if(elRisco){
    var parados = RASTREIO_NC_ITEMS.filter(function(x){
      return x.cred > 15000 && x.emp === 0;
    }).sort(function(a,b){ return b.cred - a.cred; }).slice(0, 20);

    var htmlR = '';
    if(parados.length === 0){
      htmlR = '<tr><td colspan="10" style="text-align:center;padding:24px;color:var(--ink-muted);">Nenhum crédito com saldo parado em risco crítico identificado.</td></tr>';
    } else {
      for(var rI = 0; rI < parados.length; rI++){
        var pItem = parados[rI];
        var emitCurto = pItem.emit_nome ? (pItem.emit_nome.length > 22 ? pItem.emit_nome.slice(0, 22) + '…' : pItem.emit_nome) : pItem.emit;

        htmlR += '<tr class="cel-row">' +
          '<td class="mono2"><span class="nc-num" onclick="bcmsDetalheNC(\'' + pItem.hid + '\')" style="cursor:pointer;color:var(--primary);">' + bcmsEsc(pItem.nc) + '</span></td>' +
          '<td class="mono2">' + bcmsEsc(pItem.dia) + '</td>' +
          '<td><span class="pill-status status-canc">&gt; 45 dias</span></td>' +
          '<td title="' + bcmsEsc(pItem.emit_nome) + '"><span class="ug-pill emit">' + bcmsEsc(pItem.emit) + '</span> <small>' + bcmsEsc(emitCurto) + '</small></td>' +
          '<td><span class="ug-pill fav">' + bcmsEsc(pItem.uasg) + '</span> <b>' + bcmsEsc(pItem.om_sigla) + '</b></td>' +
          '<td class="mono2">' + bcmsEsc(pItem.pi) + '</td>' +
          '<td class="mono2">' + bcmsEsc(pItem.nd) + '</td>' +
          '<td class="num font-mono">' + bcmsBRL(pItem.prov) + '</td>' +
          '<td class="num font-mono anchor" style="color:var(--warning-ink);">' + bcmsBRL(pItem.cred) + '</td>' +
          '<td style="text-align:center;"><button type="button" class="tbl-action-btn" onclick="bcmsDetalheNC(\'' + pItem.hid + '\')">Inspecionar ↗</button></td>' +
        '</tr>';
      }
    }
    elRisco.innerHTML = htmlR;
  }

  RASTREIO_INITIALIZED = true;
  bcmsFiltraRastreioNC();
  bcmsFiltraRastreioNE();
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

function bcmsFiltraRastreioNC(){
  var q = (document.getElementById('rBuscaNC') ? document.getElementById('rBuscaNC').value : '').toLowerCase().trim();
  var om = document.getElementById('rSelOM') ? document.getElementById('rSelOM').value : '';
  var qm = document.getElementById('rSelQueima') ? document.getElementById('rSelQueima').value : '';

  RASTREIO_NC_FILTRADOS = RASTREIO_NC_ITEMS.filter(function(it){
    if(om && it.uasg !== om) return false;
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

  // Atualiza indicadores visuais de ordenação
  var sortIds = ['nc', 'dia', 'emit', 'uasg', 'pi', 'nd', 'prov', 'emp', 'cred'];
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
    tbody.innerHTML = '<tr><td colspan="12" style="text-align:center;padding:24px;color:var(--ink-muted);">Nenhuma Nota de Crédito encontrada para os filtros aplicados.</td></tr>';
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
      '<td class="mono2"><span class="nc-num" onclick="bcmsDetalheNC(\'' + it.hid + '\')" style="cursor:pointer;color:var(--primary);" title="Abrir ficha cadastral da NC">' + bcmsEsc(it.nc) + '</span></td>' +
      '<td class="mono2">' + bcmsEsc(it.dia || '—') + '</td>' +
      '<td title="' + bcmsEsc(it.emit_nome) + '"><span class="ug-pill emit">' + bcmsEsc(it.emit) + '</span> <small>' + bcmsEsc(emitCurto) + '</small></td>' +
      '<td><span class="ug-pill fav">' + bcmsEsc(it.uasg) + '</span> <b>' + bcmsEsc(it.om_sigla) + '</b></td>' +
      '<td class="mono2">' + bcmsEsc(it.pi || '—') + '</td>' +
      '<td class="mono2">' + bcmsEsc(it.nd || '—') + '</td>' +
      '<td class="num font-mono">' + bcmsBRL(it.prov) + '</td>' +
      '<td class="num font-mono font-bold" style="color:var(--success-strong);">' + bcmsBRL(it.emp) + '</td>' +
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
      '<td colspan="12" style="padding:0;">' +
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
    '<div class="sub-ne-meta">Total Empenhado: <b style="color:var(--success-strong);">' + bcmsBRL(item.emp) + '</b> de <b>' + bcmsBRL(item.prov) + '</b> (' + item.taxa_queima.toFixed(1) + '% executado)</div>' +
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
    var pregaoTag = nItem.proc ? '<span class="tag-nd">' + bcmsEsc(nItem.proc) + '</span>' : '—';

    miniHtml += '<tr class="cel-row">' +
      '<td class="mono2 font-bold">' + bcmsEsc(nItem.ne) + '</td>' +
      '<td class="mono2">' + bcmsEsc(nItem.dia) + '</td>' +
      '<td><b>' + bcmsEsc(nItem.fav) + '</b></td>' +
      '<td class="mono2">' + bcmsEsc(nItem.doc) + '</td>' +
      '<td>' + pregaoTag + '</td>' +
      '<td class="num font-mono font-bold" style="color:var(--success-strong);">' + bcmsBRL(nItem.val) + '</td>' +
      '<td style="text-align:center;">' + badgeProva + '</td>' +
      '<td style="text-align:center;"><button type="button" class="tbl-action-btn" onclick="bcmsCopiarTexto(\'' + bcmsEsc(nItem.ne) + '\')">Copiar</button></td>' +
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
  if(document.getElementById('rSelQueima')) document.getElementById('rSelQueima').value = '';
  bcmsFiltraRastreioNC();
}

/* ================= AUDITORIA DE EMPENHOS (BOTTOM-UP) ================= */
function bcmsFiltraRastreioNE(){
  var q = (document.getElementById('rBuscaNE') ? document.getElementById('rBuscaNE').value : '').toLowerCase().trim();
  var om = document.getElementById('rSelOMNE') ? document.getElementById('rSelOMNE').value : '';
  var nd = document.getElementById('rSelNDNE') ? document.getElementById('rSelNDNE').value : '';

  RASTREIO_NE_FILTRADOS = RASTREIO_NE_ITEMS.filter(function(it){
    if(om && it.ug !== om) return false;
    if(nd && it.nd !== nd) return false;
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

  // Atualiza indicadores visuais de ordenação
  var sortIds = ['ne', 'dia', 'ug', 'fav', 'proc', 'val', 'nd', 'pi', 'nc'];
  for(var s = 0; s < sortIds.length; s++){
    var elS = document.getElementById('sort-ne-' + sortIds[s]);
    if(elS){
      elS.textContent = (sortIds[s] === col) ? (RASTREIO_SORT_NE.asc ? '▲' : '▼') : '';
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
    tbody.innerHTML = '<tr><td colspan="11" style="text-align:center;padding:24px;color:var(--ink-muted);">Nenhum empenho encontrado para os filtros aplicados.</td></tr>';
    document.getElementById('paginacaoRastreioNE').innerHTML = '';
    return;
  }

  var html = '';
  for(var i = 0; i < paginaItens.length; i++){
    var it = paginaItens[i];
    var temNc = it.nc && it.nc.indexOf('202') === 0;
    var sCls = it.prova_slug === 'ok' ? 'status-disp' :
              (it.prova_slug === 'info' ? 'status-parcial' :
              (it.prova_slug === 'warn' ? 'status-canc' :
              (it.prova_slug === 'danger' ? 'status-canc' : 'status-zerada')));
    var badgeProva = '<span class="pill-status ' + sCls + '">' + bcmsEsc(it.prova || 'Célula SIAFI') + '</span>';
    var pregaoTag = it.proc ? '<span class="tag-nd">' + bcmsEsc(it.proc) + '</span>' : '—';
    var ncLink = temNc ? '<a href="javascript:void(0)" onclick="bcmsDetalheNC(\'' + it.nc + '\')" class="nc-num" style="color:var(--primary);">' + bcmsEsc(it.nc) + '</a>' : '<span style="color:var(--ink-muted);">—</span>';

    html += '<tr class="cel-row">' +
      '<td class="mono2 font-bold">' + bcmsEsc(it.ne) + '</td>' +
      '<td class="mono2">' + bcmsEsc(it.dia) + '</td>' +
      '<td><span class="ug-pill fav">' + bcmsEsc(it.ug) + '</span></td>' +
      '<td title="' + bcmsEsc(it.fav) + '"><b>' + bcmsEsc(it.fav) + '</b></td>' +
      '<td class="mono2">' + bcmsEsc(it.doc) + '</td>' +
      '<td>' + pregaoTag + '</td>' +
      '<td class="num font-mono font-bold anchor" style="color:var(--success-strong);">' + bcmsBRL(it.val) + '</td>' +
      '<td class="mono2">' + bcmsEsc(it.nd) + '</td>' +
      '<td class="mono2">' + bcmsEsc(it.pi) + '</td>' +
      '<td class="mono2">' + ncLink + '</td>' +
      '<td style="text-align:center;">' + badgeProva + '</td>' +
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
  bcmsFiltraRastreioNE();
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
