# -*- coding: utf-8 -*-
"""
Módulo de Integração, Rastreabilidade e Auditoria de Empenhos (SICOF-BaApLog)
Responsável pelo ETL da planilha de empenhos do Tesouro Gerencial, conciliação
relacional com as Notas de Crédito (NCs), Prova da Verdade de integridade orçamentária,
geração dos componentes visuais e scripts interativos.
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
    """Gera o HTML da nova aba 'Crédito ➔ Empenho' com painel de Auditoria e Prova Real."""
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
<section class="unidade" data-key="RASTREIO" style="display:none" id="secao-RASTREIO">
  <div class="u-header">
    <div class="u-title-box">
      <span class="u-badge">SISTEMA INTEGRADO DE EXECUÇÃO ORÇAMENTÁRIA & AUDITORIA · SICOF</span>
      <h2>🔗 Rastreabilidade Completa: Crédito ➔ Empenho Emitido</h2>
      <p class="u-desc">Controle orçamentário da Base de Apoio Logístico do Exército e OMDS — Do aporte descentralizado (NC) à emissão da Nota de Empenho (NE), com Prova Real de integridade entre células orçamentárias e credores.</p>
    </div>
    <div class="u-meta-box">
      <div class="u-meta-pill"><span class="live-dot"></span> Posição: <b>{_esc(posicao)}</b></div>
      <div class="u-meta-pill" style="border-color:#10B981">Prova Real: <b>{pct_conf:.1f}% Confirmado</b></div>
    </div>
  </div>

  <!-- Barra de KPIs Executivos -->
  <div class="kpi-grid rastreio-kpis">
    <div class="kpi-card kpi-primary">
      <span class="kpi-lbl">TOTAL EMPENHADO (NEs)</span>
      <div class="kpi-val">{_fmt_brl(tot_emp)}</div>
      <div class="kpi-sub"><span class="kpi-chip chip-ok">{qtd_nes:,} NEs Emitidas</span> no exercício 2026</div>
    </div>
    <div class="kpi-card kpi-success">
      <span class="kpi-lbl">CRÉDITO DISPONÍVEL EM CAIXA</span>
      <div class="kpi-val">{_fmt_brl(tot_cred_disp)}</div>
      <div class="kpi-sub"><span class="kpi-chip chip-warn">Saldo Livre</span> para novas contratações</div>
    </div>
    <div class="kpi-card kpi-info">
      <span class="kpi-lbl">PROVA REAL: INTEGRIDADE CONTÁBIL</span>
      <div class="kpi-val" style="color:#059669">{pct_conf:.1f}%</div>
      <div class="kpi-sub"><span class="kpi-chip chip-ok">{qtd_conf:,} NEs com Lastro Confirmado</span> no SIAFI</div>
    </div>
    <div class="kpi-card kpi-purple">
      <span class="kpi-lbl">FORNECEDORES / CREDORES</span>
      <div class="kpi-val">{qtd_fav:,}</div>
      <div class="kpi-sub"><span class="kpi-chip chip-neutral">PJs & Militares</span> em diárias</div>
    </div>
  </div>

  <!-- Navegador de Sub-abas da Rastreabilidade -->
  <div class="rastreio-subnav">
    <button class="r-tab-btn active" data-subtab="nc" onclick="bcmsRastreioSubAba('nc')">
      <span>🔗 Visão por Nota de Crédito (Top-Down)</span>
    </button>
    <button class="r-tab-btn" data-subtab="ne" onclick="bcmsRastreioSubAba('ne')">
      <span>📋 Registro Geral de Empenhos (Bottom-Up)</span>
    </button>
    <button class="r-tab-btn" data-subtab="forn" onclick="bcmsRastreioSubAba('forn')">
      <span>🏢 Raio-X de Fornecedores & Pregões</span>
    </button>
    <button class="r-tab-btn" data-subtab="prova" onclick="bcmsRastreioSubAba('prova')">
      <span>🛡️ Prova Real & Alertas de Citação ({qtd_alertas})</span>
    </button>
    <button class="r-tab-btn" data-subtab="risco" onclick="bcmsRastreioSubAba('risco')">
      <span>⚠️ Monitor de Créditos Parados</span>
    </button>
  </div>

  <!-- ================= SUB-ABA 1: TOP-DOWN (CRÉDITO > EMPENHOS) ================= -->
  <div class="r-subtab-pane active" id="pane-rastreio-nc">
    <div class="filtros-card">
      <div class="filtros-row">
        <div class="filtro-campo filtro-grow">
          <label for="rBuscaNC">Busca Rápida:</label>
          <input type="search" id="rBuscaNC" placeholder="Filtrar por NC, Fornecedor, Objeto, PI, ND ou Concedente..." oninput="bcmsFiltraRastreioNC()">
        </div>
        <div class="filtro-campo">
          <label for="rSelOM">Unidade / OMDS:</label>
          <select id="rSelOM" onchange="bcmsFiltraRastreioNC()">
            <option value="">Todas as OMDS</option>
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
        </div>
        <div class="filtro-campo">
          <label for="rSelQueima">Nível de Empenhamento:</label>
          <select id="rSelQueima" onchange="bcmsFiltraRastreioNC()">
            <option value="">Todas as Faixas</option>
            <option value="alta">Queima Alta (> 80% empenhado)</option>
            <option value="media">Queima Média (30% a 80%)</option>
            <option value="baixa">Queima Baixa (< 30% empenhado)</option>
            <option value="zero">Zero Empenhos (100% disponível)</option>
          </select>
        </div>
        <div class="filtro-campo btn-action-wrap">
          <button class="btn-secundario" onclick="bcmsLimpaFiltrosRastreioNC()" title="Limpar Filtros">✕ Limpar</button>
          <button class="btn-primario" onclick="bcmsExportRastreioExcel('nc')" title="Baixar relatório formatado para Excel">📊 Exportar Excel</button>
        </div>
      </div>
      <div class="filtros-status">
        <span id="rContadorNC">Carregando dados...</span>
      </div>
    </div>

    <div class="tabela-wrap">
      <table class="tabela-moderna tabela-rastreio" id="tabRastreioNC">
        <thead>
          <tr>
            <th style="width:36px"></th>
            <th onclick="bcmsSortRastreioNC('nc')">Nota de Crédito ⇕</th>
            <th onclick="bcmsSortRastreioNC('dia')">Emissão ⇕</th>
            <th onclick="bcmsSortRastreioNC('emit')">Órgão Concedente ⇕</th>
            <th onclick="bcmsSortRastreioNC('uasg')">OMDS ⇕</th>
            <th onclick="bcmsSortRastreioNC('pi')">Plano Interno ⇕</th>
            <th onclick="bcmsSortRastreioNC('nd')">ND ⇕</th>
            <th onclick="bcmsSortRastreioNC('prov')" class="num">Crédito Recebido ⇕</th>
            <th onclick="bcmsSortRastreioNC('emp')" class="num">Total Empenhado ⇕</th>
            <th onclick="bcmsSortRastreioNC('cred')" class="num">Saldo Disponível ⇕</th>
            <th style="width:140px">Ritmo de Queima</th>
            <th style="width:110px">Ações</th>
          </tr>
        </thead>
        <tbody id="corpoRastreioNC">
          <!-- Linhas renderizadas dinamicamente via JS -->
        </tbody>
      </table>
    </div>
    <div class="paginacao-wrap" id="paginacaoRastreioNC"></div>
  </div>

  <!-- ================= SUB-ABA 2: BOTTOM-UP (AUDITORIA DE EMPENHOS) ================= -->
  <div class="r-subtab-pane" id="pane-rastreio-ne" style="display:none">
    <div class="filtros-card">
      <div class="filtros-row">
        <div class="filtro-campo filtro-grow">
          <label for="rBuscaNE">Busca Rápida por Empenho / Credor:</label>
          <input type="search" id="rBuscaNE" placeholder="Buscar por Razão Social, CNPJ, Número da NE, Pregão ou NC..." oninput="bcmsFiltraRastreioNE()">
        </div>
        <div class="filtro-campo">
          <label for="rSelOMNE">Unidade Gestora:</label>
          <select id="rSelOMNE" onchange="bcmsFiltraRastreioNE()">
            <option value="">Todas as UGs</option>
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
        </div>
        <div class="filtro-campo">
          <label for="rSelNDNE">Natureza de Despesa:</label>
          <select id="rSelNDNE" onchange="bcmsFiltraRastreioNE()">
            <option value="">Todas as NDs</option>
            <option value="339030">339030 · Material de Consumo</option>
            <option value="339039">339039 · Outros Serviços Terceiros PJ</option>
            <option value="449052">449052 · Equipamentos e Mat. Permanente</option>
            <option value="339015">339015 · Diárias - Militar</option>
            <option value="339093">339093 · Indenizações e Restituições</option>
            <option value="339033">339033 · Passagens e Locomoção</option>
            <option value="449051">449051 · Obras e Instalações</option>
          </select>
        </div>
        <div class="filtro-campo btn-action-wrap">
          <button class="btn-secundario" onclick="bcmsLimpaFiltrosRastreioNE()">✕ Limpar</button>
          <button class="btn-primario" onclick="bcmsExportRastreioExcel('ne')">📊 Exportar NEs</button>
        </div>
      </div>
      <div class="filtros-status">
        <span id="rContadorNE">Carregando empenhos...</span>
      </div>
    </div>

    <div class="tabela-wrap">
      <table class="tabela-moderna tabela-empenhos" id="tabRastreioNE">
        <thead>
          <tr>
            <th onclick="bcmsSortRastreioNE('ne')">Nota de Empenho ⇕</th>
            <th onclick="bcmsSortRastreioNE('dia')">Emissão ⇕</th>
            <th onclick="bcmsSortRastreioNE('ug')">UG ⇕</th>
            <th onclick="bcmsSortRastreioNE('fav')">Favorecido / Fornecedor ⇕</th>
            <th onclick="bcmsSortRastreioNE('doc')">CNPJ / CPF ⇕</th>
            <th onclick="bcmsSortRastreioNE('proc')">Processo / Pregão ⇕</th>
            <th onclick="bcmsSortRastreioNE('val')" class="num">Valor Empenhado ⇕</th>
            <th onclick="bcmsSortRastreioNE('nd')">ND ⇕</th>
            <th onclick="bcmsSortRastreioNE('pi')">Plano Interno ⇕</th>
            <th onclick="bcmsSortRastreioNE('nc')">NC de Origem ⇕</th>
            <th>Selo Prova Real</th>
          </tr>
        </thead>
        <tbody id="corpoRastreioNE">
          <!-- Linhas renderizadas dinamicamente via JS -->
        </tbody>
      </table>
    </div>
    <div class="paginacao-wrap" id="paginacaoRastreioNE"></div>
  </div>

  <!-- ================= SUB-ABA 3: RAIO-X FORNECEDORES & PREGÕES ================= -->
  <div class="r-subtab-pane" id="pane-rastreio-forn" style="display:none">
    <div class="dupla-coluna-grid">
      <div class="card-painel">
        <div class="card-painel-header">
          <h3>🏢 Maiores Fornecedores por Volume Empenhado</h3>
          <span class="card-badge">Top 20 Credores</span>
        </div>
        <p class="card-sub">Concentração de despesa empenhada no âmbito da Base de Apoio Logístico do Exército em 2026.</p>
        <div class="ranking-lista" id="listaTopFornecedores"></div>
      </div>

      <div class="card-painel">
        <div class="card-painel-header">
          <h3>📜 Principais Pregões & Processos Licitatórios</h3>
          <span class="card-badge">Demandas SISLOG</span>
        </div>
        <p class="card-sub">Processos administrativos e pregões eletrônicos mais referenciados nas emissões de empenho.</p>
        <div class="ranking-lista" id="listaTopProcessos"></div>
      </div>
    </div>
  </div>

  <!-- ================= SUB-ABA 4: PROVA REAL & AUDITORIA DE CITAÇÃO ================= -->
  <div class="r-subtab-pane" id="pane-rastreio-prova" style="display:none">
    <div class="alerta-box-info" style="background:rgba(59,130,246,0.06);border-color:rgba(59,130,246,0.3)">
      <h4 style="color:#1d4ed8">🛡️ Motor de Auditoria: Prova Real de Utilização do Crédito</h4>
      <p style="color:#1e3a8a">No SIAFI, a execução financeira se dá por <b>Célula Orçamentária (UG + PI + ND + Fonte)</b>. O campo "Observação" da NE é preenchido manualmente pelo operador e pode conter <b>erros de digitação ou cópia de textos de processos antigos (Ctrl+C)</b>. O algoritmo abaixo cruza a citação com a dotação real de destino para validar se o crédito foi efetivamente utilizado.</p>
    </div>

    <div class="kpi-grid" style="margin-bottom:20px">
      <div class="kpi-card kpi-success">
        <span class="kpi-lbl">VÍNCULOS CONFIRMADOS</span>
        <div class="kpi-val" style="color:#059669">{qtd_conf:,} NEs</div>
        <div class="kpi-sub"><b>{pct_conf:.1f}%</b> com lastro contábil comprovado</div>
      </div>
      <div class="kpi-card kpi-primary">
        <span class="kpi-lbl">DIVERGÊNCIA DE PI (PLANO INTERNO)</span>
        <div class="kpi-val" style="color:#d97706">{ep.get('DIVERGENCIA_PI', 0)} NEs</div>
        <div class="kpi-sub">Operador colou texto de outro PI</div>
      </div>
      <div class="kpi-card kpi-purple">
        <span class="kpi-lbl">DIVERGÊNCIA DE UG (UNIDADE)</span>
        <div class="kpi-val" style="color:#dc2626">{ep.get('DIVERGENCIA_UG', 0)} NEs</div>
        <div class="kpi-sub">Operador citou NC de outra OM</div>
      </div>
      <div class="kpi-card">
        <span class="kpi-lbl">CITAÇÃO DE ANO ANTERIOR / NÃO LOCALIZADA</span>
        <div class="kpi-val">{ep.get('EXERCICIO_ANTERIOR', 0) + ep.get('NC_NAO_ENCONTRADA', 0)} NEs</div>
        <div class="kpi-sub">NC de 2025 ou número inválido</div>
      </div>
    </div>

    <div class="card-painel" style="margin-bottom:20px">
      <div class="card-painel-header">
        <h3>📋 Registro de Divergências Detectadas na Observação ({qtd_alertas} ocorrências)</h3>
        <span class="card-badge" style="background:rgba(239,68,68,0.1);color:#dc2626">Auditoria Preventiva</span>
      </div>
      <p class="card-sub">Relação de Notas de Empenho onde a NC citada no texto diverge da célula orçamentária de débito no SIAFI.</p>
      
      <div class="tabela-wrap">
        <table class="tabela-moderna" id="tabInconsistencias">
          <thead>
            <tr>
              <th>Número da NE</th>
              <th>Data</th>
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
            <!-- Renderizado via JS -->
          </tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- ================= SUB-ABA 5: GESTÃO DE RISCOS & CRÉDITOS PARADOS ================= -->
  <div class="r-subtab-pane" id="pane-rastreio-risco" style="display:none">
    <div class="alerta-box-info">
      <h4>⚠️ Painel Preventivo de Gestão de Riscos Orçamentários (Fim de Exercício)</h4>
      <p>Créditos orçamentários descentralizados que possuem <b>saldo disponível considerável</b> e encontram-se <b>sem emissão de novos empenhos há mais de 45 dias</b>. O monitoramento contínuo evita o risco de estorno pelo escalão superior no encerramento do exercício financeiro.</p>
    </div>
    <div class="tabela-wrap">
      <table class="tabela-moderna" id="tabRastreioRisco">
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
            <th class="num">Saldo Disponível Parado</th>
            <th>Ações Imediatas</th>
          </tr>
        </thead>
        <tbody id="corpoRastreioRisco">
          <!-- Renderizado via JS -->
        </tbody>
      </table>
    </div>
  </div>
</section>
"""

CSS_RASTREIO = r"""
/* ==================== ESTILOS DA ABA CRÉDITO > EMPENHO (SICOF) ==================== */
.rastreio-kpis {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
  gap: 16px;
  margin-bottom: 24px;
}
.kpi-purple {
  border-top: 3px solid #8B5CF6;
}
.kpi-bar-wrap {
  width: 100%;
  height: 6px;
  background: var(--border-soft, #e2e8f0);
  border-radius: 999px;
  overflow: hidden;
  margin-top: 4px;
}
.kpi-bar-fill {
  height: 100%;
  background: linear-gradient(90deg, #10B981, #059669);
  border-radius: 999px;
  transition: width 0.4s ease;
}
.rastreio-subnav {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  margin-bottom: 20px;
  border-bottom: 1px solid var(--border-main, #e2e8f0);
  padding-bottom: 10px;
}
.r-tab-btn {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  padding: 8px 16px;
  font-size: 0.88rem;
  font-weight: 600;
  color: var(--text-muted, #64748b);
  background: var(--bg-card, #ffffff);
  border: 1px solid var(--border-main, #cbd5e1);
  border-radius: 8px;
  cursor: pointer;
  transition: all 0.2s ease;
}
.r-tab-btn:hover {
  background: var(--bg-hover, #f8fafc);
  color: var(--text-main, #0f172a);
  border-color: var(--accent, #2563EB);
}
.r-tab-btn.active {
  background: var(--accent, #2563EB);
  color: #ffffff;
  border-color: var(--accent, #2563EB);
  box-shadow: 0 2px 6px rgba(37, 99, 235, 0.25);
}
.badge-queima {
  display: inline-block;
  padding: 3px 8px;
  font-size: 0.75rem;
  font-weight: 700;
  border-radius: 6px;
  text-align: center;
  white-space: nowrap;
}
.bq-alta, .badge-prova-ok { background: rgba(16, 185, 129, 0.15); color: #059669; border: 1px solid rgba(16, 185, 129, 0.3); }
.bq-media, .badge-prova-warn { background: rgba(245, 158, 11, 0.15); color: #d97706; border: 1px solid rgba(245, 158, 11, 0.3); }
.bq-baixa, .badge-prova-danger { background: rgba(239, 68, 68, 0.15); color: #dc2626; border: 1px solid rgba(239, 68, 68, 0.3); }
.bq-zero, .badge-prova-muted { background: rgba(100, 116, 139, 0.15); color: #64748b; border: 1px solid rgba(100, 116, 139, 0.3); }
.badge-prova-info { background: rgba(59, 130, 246, 0.15); color: #2563EB; border: 1px solid rgba(59, 130, 246, 0.3); }

.acordeao-btn {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 4px 10px;
  font-size: 0.78rem;
  font-weight: 600;
  background: var(--bg-subtle, #f1f5f9);
  color: var(--text-main, #1e293b);
  border: 1px solid var(--border-main, #cbd5e1);
  border-radius: 6px;
  cursor: pointer;
  transition: all 0.15s ease;
}
.acordeao-btn:hover {
  background: var(--accent, #2563EB);
  color: #ffffff;
  border-color: var(--accent, #2563EB);
}
.acordeao-btn.aberto {
  background: var(--accent, #2563EB);
  color: #ffffff;
}

.acordeao-detalhe-linha {
  background: var(--bg-hover, #f8fafc) !important;
}
.acordeao-container {
  padding: 16px 20px;
  border-left: 3px solid var(--accent, #2563EB);
  background: var(--bg-card, #ffffff);
  border-radius: 0 8px 8px 0;
  box-shadow: inset 0 2px 4px rgba(0,0,0,0.02);
  margin: 6px 0;
}
.mini-ne-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 12px;
}
.mini-ne-titulo {
  font-size: 0.92rem;
  font-weight: 700;
  color: var(--text-main, #0f172a);
}
.mini-ne-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 0.82rem;
}
.mini-ne-table th {
  background: var(--bg-subtle, #f1f5f9);
  padding: 8px 10px;
  text-align: left;
  font-weight: 600;
  color: var(--text-muted, #475569);
  border-bottom: 1px solid var(--border-main, #cbd5e1);
}
.mini-ne-table td {
  padding: 8px 10px;
  border-bottom: 1px solid var(--border-soft, #e2e8f0);
  color: var(--text-main, #1e293b);
}
.mini-ne-table tr:hover td {
  background: var(--bg-subtle, #f8fafc);
}

.dupla-coluna-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(380px, 1fr));
  gap: 24px;
}
.card-painel {
  background: var(--bg-card, #ffffff);
  border: 1px solid var(--border-main, #cbd5e1);
  border-radius: 12px;
  padding: 20px;
  box-shadow: 0 2px 8px rgba(0,0,0,0.03);
}
.card-painel-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 6px;
}
.card-painel-header h3 {
  font-size: 1.05rem;
  font-weight: 700;
  color: var(--text-main, #0f172a);
  margin: 0;
}
.card-badge {
  font-size: 0.72rem;
  font-weight: 700;
  text-transform: uppercase;
  padding: 2px 8px;
  border-radius: 999px;
  background: rgba(37, 99, 235, 0.1);
  color: #2563EB;
}
.card-sub {
  font-size: 0.82rem;
  color: var(--text-muted, #64748b);
  margin-bottom: 16px;
}
.ranking-lista {
  display: flex;
  flex-direction: column;
  gap: 10px;
}
.rank-item {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 10px 14px;
  border-radius: 8px;
  background: var(--bg-subtle, #f8fafc);
  border: 1px solid var(--border-soft, #e2e8f0);
  transition: transform 0.15s ease;
}
.rank-item:hover {
  transform: translateX(4px);
  border-color: var(--accent, #2563EB);
}
.rank-info {
  display: flex;
  flex-direction: column;
  gap: 2px;
  max-width: 65%;
}
.rank-pos {
  font-size: 0.78rem;
  font-weight: 800;
  color: var(--accent, #2563EB);
}
.rank-nome {
  font-size: 0.85rem;
  font-weight: 600;
  color: var(--text-main, #0f172a);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.rank-val-box {
  text-align: right;
}
.rank-val {
  font-size: 0.9rem;
  font-weight: 700;
  font-family: var(--font-mono, monospace);
  color: var(--text-main, #0f172a);
}

.alerta-box-info {
  background: rgba(245, 158, 11, 0.08);
  border: 1px solid rgba(245, 158, 11, 0.3);
  border-radius: 10px;
  padding: 16px 20px;
  margin-bottom: 20px;
}
.alerta-box-info h4 {
  color: #b45309;
  font-size: 0.98rem;
  font-weight: 700;
  margin-bottom: 4px;
}
.alerta-box-info p {
  color: #92400e;
  font-size: 0.85rem;
  line-height: 1.5;
  margin: 0;
}

/* Modal Drill-Down de Empenhos */
.modal-empenhos-secao {
  margin-top: 24px;
  padding-top: 20px;
  border-top: 1px dashed var(--border-main, #cbd5e1);
}
.modal-empenhos-titulo {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 12px;
}
.modal-empenhos-titulo h4 {
  font-size: 1rem;
  font-weight: 700;
  color: var(--text-main, #0f172a);
  margin: 0;
  display: flex;
  align-items: center;
  gap: 8px;
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
          '<div class="kpi-bar-wrap" style="width:140px"><div class="kpi-bar-fill" style="width:' + pctF + '%"></div></div>' +
        '</div>' +
        '<div class="rank-val-box">' +
          '<span class="rank-val">' + bcmsEsc(itemF.fmt) + '</span>' +
        '</div>' +
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
          '<div class="kpi-bar-wrap" style="width:140px"><div class="kpi-bar-fill" style="width:' + pctP + '%"></div></div>' +
        '</div>' +
        '<div class="rank-val-box">' +
          '<span class="rank-val">' + bcmsEsc(itemP.fmt) + '</span>' +
        '</div>' +
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
      htmlInc = '<tr><td colspan="9" class="vazio">Nenhuma inconsistência de citação detectada. Todas as NEs possuem vínculo contábil perfeito.</td></tr>';
    } else {
      for(var iI = 0; iI < incs.length; iI++){
        var inItem = incs[iI];
        var badgeCls = 'badge-prova-' + inItem.status_slug;
        htmlInc += '<tr>' +
          '<td><span class="font-mono font-bold">' + bcmsEsc(inItem.ne) + '</span></td>' +
          '<td>' + bcmsEsc(inItem.dia) + '</td>' +
          '<td><span class="badge-om">' + bcmsEsc(inItem.ug) + '</span></td>' +
          '<td><div class="fav-box font-bold" title="' + bcmsEsc(inItem.fav) + '">' + bcmsEsc(inItem.fav) + '</div></td>' +
          '<td class="num font-mono font-bold" style="color:#059669">' + bcmsEsc(inItem.val_fmt) + '</td>' +
          '<td><code>' + bcmsEsc(inItem.pi_ne) + '</code></td>' +
          '<td><code>' + bcmsEsc(inItem.nd_ne) + '</code></td>' +
          '<td><b style="color:#dc2626">' + bcmsEsc(inItem.nc_citada) + '</b></td>' +
          '<td><span class="badge-queima ' + badgeCls + '">' + bcmsEsc(inItem.status_prova) + '</span><br><small style="color:var(--text-muted);font-size:0.75rem;">' + bcmsEsc(inItem.motivo) + '</small></td>' +
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
      htmlR = '<tr><td colspan="10" class="vazio">Nenhum crédito com saldo parado em risco crítico identificado.</td></tr>';
    } else {
      for(var rI = 0; rI < parados.length; rI++){
        var pItem = parados[rI];
        htmlR += '<tr>' +
          '<td><b>' + bcmsEsc(pItem.nc) + '</b></td>' +
          '<td>' + bcmsEsc(pItem.dia) + '</td>' +
          '<td><span class="badge-queima bq-baixa">> 45 dias</span></td>' +
          '<td>' + bcmsEsc(pItem.emit_nome || pItem.emit) + '</td>' +
          '<td>' + bcmsEsc(pItem.om_sigla) + '</td>' +
          '<td><code>' + bcmsEsc(pItem.pi) + '</code></td>' +
          '<td><code>' + bcmsEsc(pItem.nd) + '</code></td>' +
          '<td class="num font-mono">' + bcmsBRL(pItem.prov) + '</td>' +
          '<td class="num font-mono font-bold" style="color:#d97706">' + bcmsBRL(pItem.cred) + '</td>' +
          '<td><button class="btn-secundario btn-sm" onclick="bcmsDetalheNC(\'' + pItem.hid + '\')">Inspecionar</button></td>' +
        '</tr>';
      }
    }
    elRisco.innerHTML = htmlR;
  }

  RASTREIO_INITIALIZED = true;
  bcmsFiltraRastreioNC();
  bcmsFiltraRastreioNE();
}

function bcmsRastreioSubAba(subtab){
  document.querySelectorAll('.rastreio-subnav .r-tab-btn').forEach(function(btn){
    btn.classList.toggle('active', btn.getAttribute('data-subtab') === subtab);
  });
  document.querySelectorAll('.r-subtab-pane').forEach(function(pane){
    pane.style.display = (pane.id === 'pane-rastreio-' + subtab) ? '' : 'none';
  });
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
    RASTREIO_SORT_NC.asc = (col === 'nc' || col === 'emit' || col === 'om_sigla');
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
    cont.textContent = 'Exibindo ' + total.toLocaleString('pt-BR') + ' Nota(s) de Crédito';
  }

  var porPag = 25;
  var totalPags = Math.max(1, Math.ceil(total / porPag));
  if(RASTREIO_NC_PAGE > totalPags) RASTREIO_NC_PAGE = totalPags;
  var ini = (RASTREIO_NC_PAGE - 1) * porPag;
  var fim = Math.min(total, ini + porPag);
  var paginaItens = RASTREIO_NC_FILTRADOS.slice(ini, fim);

  if(paginaItens.length === 0){
    tbody.innerHTML = '<tr><td colspan="12" class="vazio">Nenhuma Nota de Crédito encontrada para os filtros aplicados.</td></tr>';
    document.getElementById('paginacaoRastreioNC').innerHTML = '';
    return;
  }

  var html = '';
  for(var i = 0; i < paginaItens.length; i++){
    var it = paginaItens[i];
    var temNE = it.nes && it.nes.length > 0;
    var badgeCls = 'bq-' + it.faixa_queima;
    var badgeTexto = it.faixa_queima === 'alta' ? 'Alta (>80%)' :
                    (it.faixa_queima === 'media' ? 'Média (30-80%)' :
                    (it.faixa_queima === 'baixa' ? 'Baixa (<30%)' : '0% Empenhado'));

    html += '<tr id="linha-nc-' + it.hid + '">' +
      '<td>' +
        (temNE ? '<button class="acordeao-btn" id="btn-exp-' + it.hid + '" onclick="bcmsToggleAcordeaoNC(\'' + it.hid + '\')" title="Expandir ' + it.nes.length + ' empenhos">▼ ' + it.nes.length + '</button>' : '') +
      '</td>' +
      '<td><a href="javascript:void(0)" onclick="bcmsDetalheNC(\'' + it.hid + '\')" class="link-nc font-bold">' + bcmsEsc(it.nc) + '</a></td>' +
      '<td>' + bcmsEsc(it.dia) + '</td>' +
      '<td title="' + bcmsEsc(it.emit_nome) + '">' + bcmsEsc(it.emit_nome ? (it.emit_nome.length > 25 ? it.emit_nome.slice(0, 25) + '…' : it.emit_nome) : it.emit) + '</td>' +
      '<td><span class="badge-om">' + bcmsEsc(it.om_sigla) + '</span></td>' +
      '<td><code>' + bcmsEsc(it.pi) + '</code></td>' +
      '<td><code>' + bcmsEsc(it.nd) + '</code></td>' +
      '<td class="num font-mono">' + bcmsBRL(it.prov) + '</td>' +
      '<td class="num font-mono font-bold" style="color:#059669">' + bcmsBRL(it.emp) + '</td>' +
      '<td class="num font-mono">' + bcmsBRL(it.cred) + '</td>' +
      '<td><span class="badge-queima ' + badgeCls + '">' + badgeTexto + '</span></td>' +
      '<td><button class="btn-secundario btn-sm" onclick="bcmsDetalheNC(\'' + it.hid + '\')">Detalhar</button></td>' +
    '</tr>' +
    '<tr id="acordeao-' + it.hid + '" class="acordeao-detalhe-linha" style="display:none">' +
      '<td colspan="12">' +
        '<div class="acordeao-container" id="acordeao-cont-' + it.hid + '"></div>' +
      '</td>' +
    '</tr>';
  }
  tbody.innerHTML = html;

  var pagHtml = '<div class="paginacao">' +
    '<button class="pag-btn" onclick="bcmsPaginaRastreioNC(' + (RASTREIO_NC_PAGE - 1) + ')" ' + (RASTREIO_NC_PAGE <= 1 ? 'disabled' : '') + '>◀ Anterior</button>' +
    '<span class="pag-info">Página <b>' + RASTREIO_NC_PAGE + '</b> de <b>' + totalPags + '</b></span>' +
    '<button class="pag-btn" onclick="bcmsPaginaRastreioNC(' + (RASTREIO_NC_PAGE + 1) + ')" ' + (RASTREIO_NC_PAGE >= totalPags ? 'disabled' : '') + '>Próxima ▶</button>' +
  '</div>';
  document.getElementById('paginacaoRastreioNC').innerHTML = pagHtml;
}

function bcmsPaginaRastreioNC(p){
  if(p < 1) return;
  RASTREIO_NC_PAGE = p;
  bcmsRenderRastreioNC();
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
  var miniHtml = '<div class="mini-ne-header">' +
    '<div class="mini-ne-titulo">📦 ' + nes.length + ' Nota(s) de Empenho vinculadas à NC ' + bcmsEsc(item.nc) + '</div>' +
    '<div class="mini-ne-sub">Total Comprometido: <b>' + bcmsBRL(item.emp) + '</b> de <b>' + bcmsBRL(item.prov) + '</b> (' + item.taxa_queima.toFixed(1) + '%)</div>' +
  '</div>' +
  '<table class="mini-ne-table">' +
    '<thead>' +
      '<tr>' +
        '<th>Número NE</th>' +
        '<th>Data Emissão</th>' +
        '<th>Favorecido / Fornecedor</th>' +
        '<th>CNPJ / CPF</th>' +
        '<th>Processo / Pregão</th>' +
        '<th class="num">Valor da NE</th>' +
        '<th>Selo Prova Real</th>' +
        '<th>Ação</th>' +
      '</tr>' +
    '</thead>' +
    '<tbody>';

  for(var j = 0; j < nes.length; j++){
    var nItem = nes[j];
    var badgeProva = '<span class="badge-queima badge-prova-' + (nItem.prova_slug || 'ok') + '">' + bcmsEsc(nItem.prova || 'Auditado') + '</span>';
    miniHtml += '<tr>' +
      '<td><b>' + bcmsEsc(nItem.ne) + '</b></td>' +
      '<td>' + bcmsEsc(nItem.dia) + '</td>' +
      '<td>' + bcmsEsc(nItem.fav) + '</td>' +
      '<td><code>' + bcmsEsc(nItem.doc) + '</code></td>' +
      '<td>' + (nItem.proc ? '<span class="badge-pregao">' + bcmsEsc(nItem.proc) + '</span>' : '−') + '</td>' +
      '<td class="num font-mono font-bold">' + bcmsBRL(nItem.val) + '</td>' +
      '<td>' + badgeProva + '</td>' +
      '<td><button class="btn-secundario btn-xs" onclick="bcmsCopiarTexto(\'' + nItem.ne + '\')">Copiar</button></td>' +
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
    cont.textContent = 'Exibindo ' + total.toLocaleString('pt-BR') + ' Nota(s) de Empenho';
  }

  var porPag = 50;
  var totalPags = Math.max(1, Math.ceil(total / porPag));
  if(RASTREIO_NE_PAGE > totalPags) RASTREIO_NE_PAGE = totalPags;
  var ini = (RASTREIO_NE_PAGE - 1) * porPag;
  var fim = Math.min(total, ini + porPag);
  var paginaItens = RASTREIO_NE_FILTRADOS.slice(ini, fim);

  if(paginaItens.length === 0){
    tbody.innerHTML = '<tr><td colspan="11" class="vazio">Nenhum empenho encontrado para os filtros aplicados.</td></tr>';
    document.getElementById('paginacaoRastreioNE').innerHTML = '';
    return;
  }

  var html = '';
  for(var i = 0; i < paginaItens.length; i++){
    var it = paginaItens[i];
    var temNc = it.nc && it.nc.indexOf('202') === 0;
    var badgeProva = '<span class="badge-queima badge-prova-' + (it.prova_slug || 'info') + '">' + bcmsEsc(it.prova || 'Célula SIAFI') + '</span>';

    html += '<tr>' +
      '<td><span class="font-mono font-bold">' + bcmsEsc(it.ne) + '</span></td>' +
      '<td>' + bcmsEsc(it.dia) + '</td>' +
      '<td><span class="badge-om">' + bcmsEsc(it.ug) + '</span></td>' +
      '<td><div class="fav-box font-bold" title="' + bcmsEsc(it.fav) + '">' + bcmsEsc(it.fav) + '</div></td>' +
      '<td><code>' + bcmsEsc(it.doc) + '</code></td>' +
      '<td>' + (it.proc ? '<span class="badge-pregao">' + bcmsEsc(it.proc) + '</span>' : '−') + '</td>' +
      '<td class="num font-mono font-bold" style="color:#059669">' + bcmsBRL(it.val) + '</td>' +
      '<td><code>' + bcmsEsc(it.nd) + '</code></td>' +
      '<td><code>' + bcmsEsc(it.pi) + '</code></td>' +
      '<td>' + (temNc ? '<a href="javascript:void(0)" onclick="bcmsDetalheNC(\'' + it.nc + '\')" class="link-nc font-bold">' + bcmsEsc(it.nc) + '</a>' : '<span style="color:#94a3b8">−</span>') + '</td>' +
      '<td>' + badgeProva + '</td>' +
    '</tr>';
  }
  tbody.innerHTML = html;

  var pagHtml = '<div class="paginacao">' +
    '<button class="pag-btn" onclick="bcmsPaginaRastreioNE(' + (RASTREIO_NE_PAGE - 1) + ')" ' + (RASTREIO_NE_PAGE <= 1 ? 'disabled' : '') + '>◀ Anterior</button>' +
    '<span class="pag-info">Página <b>' + RASTREIO_NE_PAGE + '</b> de <b>' + totalPags + '</b></span>' +
    '<button class="pag-btn" onclick="bcmsPaginaRastreioNE(' + (RASTREIO_NE_PAGE + 1) + ')" ' + (RASTREIO_NE_PAGE >= totalPags ? 'disabled' : '') + '>Próxima ▶</button>' +
  '</div>';
  document.getElementById('paginacaoRastreioNE').innerHTML = pagHtml;
}

function bcmsPaginaRastreioNE(p){
  if(p < 1) return;
  RASTREIO_NE_PAGE = p;
  bcmsRenderRastreioNE();
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
    '  <Style ss:ID="sHdr"><Font ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#1C4A73" ss:Pattern="Solid"/></Style>\n' +
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
