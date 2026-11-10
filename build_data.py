import json
import datetime
from pathlib import Path
from collections import defaultdict

def build(root_dir=None):
    if root_dir is None:
        p = Path(__file__).resolve().parent
        if p.name == "prototipo-v3":
            root_dir = p.parent
        elif (p / "data").exists():
            root_dir = p
        elif (p / "Dashboard-Credito-BCMS" / "data").exists():
            root_dir = p / "Dashboard-Credito-BCMS"
        else:
            root_dir = Path("Dashboard-Credito-BCMS").resolve() if Path("Dashboard-Credito-BCMS").exists() else p

    ROOT = Path(root_dir)
    data_json_path = ROOT / "data" / "data.json"
    resumos_path = ROOT / "data" / "resumos_nc.json"
    history_path = ROOT / "data" / "history.json"

    assert data_json_path.exists(), f"Erro: {data_json_path} não encontrado!"
    assert history_path.exists(), f"Erro: {history_path} não encontrado!"

    data_json = json.loads(data_json_path.read_text(encoding="utf-8"))
    resumos = json.loads(resumos_path.read_text(encoding="utf-8")) if resumos_path.exists() else {}
    history = json.loads(history_path.read_text(encoding="utf-8"))

    teladata = data_json.get("teladata", {})
    daydata = data_json.get("daydata", {})
    ncdata = data_json.get("ncdata", {})
    emp = data_json.get("empenhodata", {})
    siafi_totais = emp.get("om_totais_siafi", {})
    nes_raw = emp.get("nes", [])
    nc_to_nes = emp.get("nc_to_nes", {})

    UG_NAMES = {
        "160504": "CMAT - Centro de Material Bélico / COLOG",
        "160073": "SGS / DGO - Diretoria de Gestão Orçamentária",
        "160505": "DSup - Diretoria de Suprimento / COLOG",
        "167505": "D Sau - Diretoria de Saúde (FUSEx)",
        "167504": "D Sau - Diretoria de Saúde (Apoio Hospitalar)",
        "167503": "D Sau - Diretoria de Saúde (Descentralização)",
        "160298": "Comando Militar do Leste (CML)",
        "160539": "DMAvEx - Diretoria de Material de Aviação",
        "160502": "DEC - Departamento de Engenharia e Construção",
        "160035": "CCOMGEX - Cmdo Comunicações e Guerra Eletrônica",
        "167086": "D Sau - Fundo de Saúde do Exército",
        "160238": "Ba Ap Log - Base de Apoio Logístico",
        "167238": "Ba Ap Log (Fundo de Saúde)",
        "160329": "BCMS - Batalhão Central de Manutenção e Suprimento",
        "167329": "BCMS (Fundo de Saúde)",
        "160304": "BMSA - Batalhão de Manutenção e Suprimento de Armamento",
        "167304": "BMSA (Fundo de Saúde)",
        "160307": "1º D Sup - 1º Depósito de Suprimento",
        "167307": "1º D Sup (Fundo de Saúde)",
        "160321": "ECT - Estabelecimento Central de Transportes",
        "167321": "ECT (Fundo de Saúde)",
        "160246": "D C Mun - Depósito Central de Munições",
        "167246": "D C Mun (Fundo de Saúde)"
    }

    # 1. Mapa Completo de Notas de Crédito
    mapa_ncs = {}

    for k, v in ncdata.items():
        nc_code = v.get("nc")
        if not nc_code: continue
        short_nc = nc_code[-10:] if len(nc_code) >= 10 else nc_code
        om_key = k.split("_")[0]
        concedente_ug = nc_code[:6] if len(nc_code) >= 6 else ""
        concedente_nome = UG_NAMES.get(concedente_ug, f"UG {concedente_ug}" if concedente_ug else "Órgão Concedente")

        if nc_code not in mapa_ncs:
            prazo_val = resumos.get(nc_code, {}).get("prazo") or "30/11/2026"
            mapa_ncs[nc_code] = {
                "nc": nc_code,
                "short_nc": short_nc,
                "om": om_key,
                "ug": v.get("u", "160"),
                "concedente_ug": concedente_ug,
                "concedente_nome": concedente_nome,
                "acao": v.get("acao", ""),
                "pi": v.get("pi", ""),
                "nd": v.get("nd", ""),
                "ndn": v.get("ndn", ""),
                "op_principal": v.get("op", "DESCENTRALIZACAO DE CREDITO"),
                "dia": v.get("dia", ""),
                "val_alocado": 0.0,
                "val_reduzido": 0.0,
                "saldo_tela": 0.0,
                "dias_tela": 0,
                "prazo": prazo_val,
                "obj": v.get("obj", ""),
                "movimentacoes": [],
                "nes": []
            }
        entry = mapa_ncs[nc_code]
        val = float(v.get("val", 0.0))
        if val >= 0:
            entry["val_alocado"] += val
        else:
            entry["val_reduzido"] += abs(val)

        entry["movimentacoes"].append({
            "dia": v.get("dia", ""),
            "op": v.get("op", ""),
            "nd": v.get("nd", ""),
            "ndn": v.get("ndn", ""),
            "val": val,
            "obj": v.get("obj", "")
        })

    # Cruzar com teladata
    for k, v in teladata.items():
        nc_code = v.get("nc")
        if nc_code in mapa_ncs:
            mapa_ncs[nc_code]["saldo_tela"] = float(v.get("v", 0.0))
            mapa_ncs[nc_code]["dias_tela"] = int(v.get("dias", 0))
            if v.get("obj"):
                mapa_ncs[nc_code]["obj"] = v.get("obj")
            if v.get("nd"):
                mapa_ncs[nc_code]["nd"] = v.get("nd")
            if v.get("ndnome"):
                mapa_ncs[nc_code]["ndn"] = v.get("ndnome")
            if v.get("pi"):
                mapa_ncs[nc_code]["pi"] = v.get("pi")

    # Cruzar com nc_to_nes
    for nc_key, nes_list in nc_to_nes.items():
        matched = None
        if nc_key in mapa_ncs:
            matched = mapa_ncs[nc_key]
        else:
            for full_nc, item in mapa_ncs.items():
                if item["short_nc"] == nc_key or full_nc.endswith(nc_key):
                    matched = item
                    break
        if matched:
            matched["nes"] = [
                {
                    "ne": n.get("ne", ""),
                    "dia": n.get("dia", ""),
                    "fav": n.get("fav", ""),
                    "doc": n.get("doc", ""),
                    "val": float(n.get("val", 0.0)),
                    "proc": n.get("proc", ""),
                    "prova": n.get("prova", "Confirmado")
                }
                for n in nes_list
            ]

    # Adicionar ponteiros por código curto
    for full_nc, entry in list(mapa_ncs.items()):
        short = entry["short_nc"]
        if short and short not in mapa_ncs:
            mapa_ncs[short] = entry

    # 2. Processar Créditos em Tela
    creditos_em_tela = []
    for k, v in teladata.items():
        om_key = k.split("_")[0]
        nc_num = v.get("nc", "")
        dias = int(v.get("dias", 0))
        val_disp = float(v.get("v", 0.0))
        
        prazo = resumos.get(nc_num, {}).get("prazo")
        if not prazo:
            prazo = "15/11/2026" if dias > 45 else ("30/11/2026" if dias > 20 else "15/12/2026")
            
        info_mapa = mapa_ncs.get(nc_num, {})
        val_alocado = info_mapa.get("val_alocado", val_disp * 1.5)
        nes_nc = info_mapa.get("nes", [])
        val_emp = sum(n["val"] for n in nes_nc) if nes_nc else max(0.0, val_alocado - val_disp)
        if val_alocado <= 0:
            val_alocado = val_disp + val_emp
            
        taxa_exec = (val_emp / val_alocado * 100) if val_alocado > 0 else 0.0

        creditos_em_tela.append({
            "id": k,
            "om": om_key,
            "nc": nc_num,
            "uasg": v.get("uasg", ""),
            "u": str(v.get("u", "160")),
            "nd": v.get("nd", ""),
            "ndnome": v.get("ndnome", ""),
            "pi": v.get("pi", ""),
            "pinome": v.get("pinome", ""),
            "v": val_disp,
            "prov": round(val_alocado, 2),
            "emp": round(val_emp, 2),
            "saldo": val_disp,
            "taxa": round(taxa_exec, 2),
            "dias": dias,
            "dia": v.get("dia", ""),
            "obj": v.get("obj", ""),
            "prazo": prazo,
            "concedente": info_mapa.get("concedente_nome", "Órgão Concedente")
        })

    creditos_em_tela.sort(key=lambda x: x["v"], reverse=True)

    # 3. Processar Resumo Diário de Movimentações
    resumo_diario = {}
    for k, v in daydata.items():
        om_key = k.split("_")[0]
        if om_key not in resumo_diario:
            resumo_diario[om_key] = []
        resumo_diario[om_key].append(v)

    # 4. Unidades com Nomenclaturas Oficiais e Totais Dinâmicos
    om_definitions = [
        {
            "id": "BCMS",
            "sigla": "BCMS",
            "nome": "Batalhão Central de Manutenção e Suprimento",
            "descricao": "Manutenção de viaturas, blindados e suprimento classe IX",
            "logo": "assets/logos/BCMS.png",
            "accent": "#ef4444",
            "ug_ord": "160329",
            "ug_sau": "167329",
            "ugs": ["160329 (Ordinário)", "167329 (Saúde)"]
        },
        {
            "id": "BAAPLOG",
            "sigla": "Ba Ap Log",
            "nome": "Base de Apoio Logístico",
            "descricao": "Comando logístico regional e administração geral das operações",
            "logo": "assets/logos/BAAPLOG.png",
            "accent": "#3b82f6",
            "ug_ord": "160238",
            "ug_sau": "167238",
            "ugs": ["160238 (Ordinário)", "167238 (Saúde)"]
        },
        {
            "id": "DCMUN",
            "sigla": "D C Mun",
            "nome": "Depósito Central de Munições",
            "descricao": "Gestão, armazenagem e segurança estratégica de munições",
            "logo": "assets/logos/DCMUN.png",
            "accent": "#0ea5e9",
            "ug_ord": "160246",
            "ug_sau": "167246",
            "ugs": ["160246 (Ordinário)", "167246 (Saúde)"]
        },
        {
            "id": "BMSA",
            "sigla": "BMSA",
            "nome": "Batalhão de Manutenção e Suprimento de Armamento",
            "descricao": "Manutenção pesada de armamento leve e pesado do Exército",
            "logo": "assets/logos/BMSA.png",
            "accent": "#8b5cf6",
            "ug_ord": "160304",
            "ug_sau": "167304",
            "ugs": ["160304 (Ordinário)", "167304 (Saúde)"]
        },
        {
            "id": "DSUP1",
            "sigla": "1º D Sup",
            "nome": "1º Depósito de Suprimento",
            "descricao": "Distribuição e gestão de suprimentos logísticos regionais",
            "logo": "assets/logos/1DSUP.png",
            "accent": "#10b981",
            "ug_ord": "160307",
            "ug_sau": "167307",
            "ugs": ["160307 (Ordinário)", "167307 (Saúde)"]
        },
        {
            "id": "ECT",
            "sigla": "ECT",
            "nome": "Estabelecimento Central de Transportes",
            "descricao": "Gestão estratégica de frotas operacionais e transporte rodoviário",
            "logo": "assets/logos/Ect.png",
            "accent": "#f59e0b",
            "ug_ord": "160321",
            "ug_sau": "167321",
            "ugs": ["160321 (Ordinário)", "167321 (Saúde)"]
        }
    ]

    unidades = []
    tot_prov, tot_emp, tot_liq, tot_pag, tot_cred, tot_conc = 0.0, 0.0, 0.0, 0.0, 0.0, 0.0

    for om_def in om_definitions:
        sigla = om_def["sigla"]
        t_om = siafi_totais.get(sigla, {})
        t_ord = siafi_totais.get(om_def["ug_ord"], {})
        t_sau = siafi_totais.get(om_def["ug_sau"], {})

        prov = float(t_om.get("prov", 0.0))
        emp_val = float(t_om.get("emp", 0.0))
        liq = float(t_om.get("liq", 0.0))
        pag = float(t_om.get("pag", 0.0))
        cred = float(t_om.get("cred", 0.0))
        conc = float(t_om.get("conc", 0.0))

        tot_prov += prov
        tot_emp += emp_val
        tot_liq += liq
        tot_pag += pag
        tot_cred += cred
        tot_conc += conc

        taxa_emp = (emp_val / prov * 100) if prov > 0 else 0.0
        taxa_liq = (liq / emp_val * 100) if emp_val > 0 else 0.0
        taxa_pag = (pag / liq * 100) if liq > 0 else 0.0

        ord_prov = float(t_ord.get("prov", 0.0))
        ord_emp = float(t_ord.get("emp", 0.0))
        ord_liq = float(t_ord.get("liq", 0.0))
        ord_pag = float(t_ord.get("pag", 0.0))
        ord_cred = float(t_ord.get("cred", 0.0))

        sau_prov = float(t_sau.get("prov", 0.0))
        sau_emp = float(t_sau.get("emp", 0.0))
        sau_liq = float(t_sau.get("liq", 0.0))
        sau_pag = float(t_sau.get("pag", 0.0))
        sau_cred = float(t_sau.get("cred", 0.0))

        unidades.append({
            "id": om_def["id"],
            "sigla": sigla,
            "nome": om_def["nome"],
            "descricao": om_def["descricao"],
            "logo": om_def["logo"],
            "accent": om_def["accent"],
            "ugs": om_def["ugs"],
            "prov": round(prov, 2),
            "emp": round(emp_val, 2),
            "liq": round(liq, 2),
            "pag": round(pag, 2),
            "cred": round(cred, 2),
            "conc": round(conc, 2),
            "taxaEmpenho": round(taxa_emp, 2),
            "taxaLiquidacao": round(taxa_liq, 2),
            "taxaPagamento": round(taxa_pag, 2),
            "sub": {
                "ordinario": {
                    "ug": om_def["ug_ord"],
                    "prov": round(ord_prov, 2),
                    "emp": round(ord_emp, 2),
                    "liq": round(ord_liq, 2),
                    "pag": round(ord_pag, 2),
                    "cred": round(ord_cred, 2),
                    "taxaEmp": round((ord_emp / ord_prov * 100) if ord_prov > 0 else 0, 2)
                },
                "saude": {
                    "ug": om_def["ug_sau"],
                    "prov": round(sau_prov, 2),
                    "emp": round(sau_emp, 2),
                    "liq": round(sau_liq, 2),
                    "pag": round(sau_pag, 2),
                    "cred": round(sau_cred, 2),
                    "taxaEmp": round((sau_emp / sau_prov * 100) if sau_prov > 0 else 0, 2)
                }
            }
        })

    taxa_emp_total = (tot_emp / tot_prov * 100) if tot_prov > 0 else 0.0
    taxa_liq_total = (tot_liq / tot_emp * 100) if tot_emp > 0 else 0.0
    taxa_pag_total = (tot_pag / tot_liq * 100) if tot_liq > 0 else 0.0

    unidade_total = {
        "id": "total",
        "sigla": "CONSOLIDADO",
        "nome": "Todas as Unidades (Visão Global)",
        "descricao": "Base de Apoio Logístico e todas as 5 OMDS Subordinadas",
        "logo": None,
        "accent": "#38bdf8",
        "ugs": ["Consolidado Geral (12 UGs)"],
        "prov": round(tot_prov, 2),
        "emp": round(tot_emp, 2),
        "liq": round(tot_liq, 2),
        "pag": round(tot_pag, 2),
        "cred": round(tot_cred, 2),
        "conc": round(tot_conc, 2),
        "taxaEmpenho": round(taxa_emp_total, 2),
        "taxaLiquidacao": round(taxa_liq_total, 2),
        "taxaPagamento": round(taxa_pag_total, 2)
    }
    unidades.insert(0, unidade_total)

    # 5. Timeline dinâmica
    timeline = []
    if history:
        step = max(1, len(history) // 10)
        selected_history = [history[i] for i in range(0, len(history), step)]
        if history[-1] not in selected_history:
            selected_history.append(history[-1])

        for h in selected_history:
            d_str = h.get("data", "")
            t = h.get("total", {})
            try:
                parts = d_str.split("-")
                label = f"{parts[2]}/{parts[1]}"
            except Exception:
                label = d_str
            timeline.append({
                "data": label,
                "prov": round(float(t.get("prov", 0.0)), 2),
                "emp": round(float(t.get("emp", 0.0)), 2),
                "liq": round(float(t.get("liq", 0.0)), 2),
                "pag": round(float(t.get("pag", 0.0)), 2),
                "cred": round(float(t.get("cred", 0.0)), 2)
            })

    # 6. Top Fornecedores e Processos
    fav_map = defaultdict(lambda: {"val": 0.0, "qtd": 0, "cnpjs": set(), "oms": set()})
    proc_map = defaultdict(lambda: {"val": 0.0, "qtd": 0, "descs": [], "oms": set()})

    todas_nes = []
    for n in nes_raw:
        fav = (n.get("fav") or "").strip().upper()
        val = float(n.get("val") or 0.0)
        doc = (n.get("doc") or "").strip()
        om = n.get("om_sigla", "")
        proc = (n.get("proc") or "").strip()
        desc = (n.get("desc") or "").strip()
        ne_code = n.get("ne", "")
        ug = n.get("ug", "")
        dia = n.get("dia", "")
        nc = n.get("nc", "")

        if fav:
            fav_map[fav]["val"] += val
            fav_map[fav]["qtd"] += 1
            if doc: fav_map[fav]["cnpjs"].add(doc)
            if om: fav_map[fav]["oms"].add(om)

        if proc:
            proc_map[proc]["val"] += val
            proc_map[proc]["qtd"] += 1
            if desc and len(proc_map[proc]["descs"]) < 2:
                proc_map[proc]["descs"].append(desc[:100])
            if om: proc_map[proc]["oms"].add(om)

        todas_nes.append({
            "ne": ne_code,
            "ug": ug,
            "omSigla": om,
            "credor": fav,
            "cnpj": doc,
            "valor": round(val, 2),
            "data": dia,
            "processo": proc if proc else (n.get("acao", "") or "-"),
            "ncOrigem": nc,
            "status": "Liquidado" if (float(n.get("pago", 0)) > 0 or float(n.get("a_liquidar", 0)) == 0) else "Emitido",
            "aLiquidar": round(float(n.get("a_liquidar", 0.0)), 2),
            "pago": round(float(n.get("pago", 0.0)), 2)
        })

    todas_nes.sort(key=lambda x: (x["data"].split("/")[::-1] if "/" in x["data"] else [x["data"]]), reverse=True)

    top_fornecedores = []
    for name, info in sorted(fav_map.items(), key=lambda x: x[1]["val"], reverse=True)[:15]:
        cnpj_str = list(info["cnpjs"])[0] if info["cnpjs"] else "N/I"
        if len(cnpj_str) == 14 and cnpj_str.isdigit():
            cnpj_fmt = f"{cnpj_str[:2]}.{cnpj_str[2:5]}.{cnpj_str[5:8]}/{cnpj_str[8:12]}-{cnpj_str[12:]}"
        else:
            cnpj_fmt = cnpj_str
            
        top_fornecedores.append({
            "nome": name,
            "cnpj": cnpj_fmt,
            "qtdNEs": info["qtd"],
            "omds": ", ".join(sorted(info["oms"])),
            "val": round(info["val"], 2)
        })

    top_processos = []
    for proc_name, info in sorted(proc_map.items(), key=lambda x: x[1]["val"], reverse=True)[:15]:
        objeto = info["descs"][0] if info["descs"] else "Contratação / Aquisição de Material ou Serviço"
        top_processos.append({
            "proc": proc_name,
            "objeto": objeto,
            "omds": ", ".join(sorted(info["oms"])),
            "val": round(info["val"], 2)
        })

    # Data de posição e encerramento
    meta_meta = data_json.get("metadata", {})
    gerado_em = meta_meta.get("gerado_em", "")
    if "às" in gerado_em:
        posicao_data = gerado_em.split("às")[0].strip()
    else:
        posicao_data = datetime.date.today().strftime("%d/%m/%Y")

    hoje = datetime.date.today()
    fim_exercicio = datetime.date(2026, 12, 31)
    dias_restantes = max(0, (fim_exercicio - hoje).days)

    data_dict = {
        "meta": {
            "exercicio": "2026",
            "posicaoData": posicao_data,
            "fonte": "SIAFI / Tesouro Gerencial / SICOF",
            "orgaoSuperior": "Comando Logístico (COLOG) / Base de Apoio Logístico",
            "statusGeral": f"Execução Avançada ({round(taxa_emp_total, 2)}% Empenhado)",
            "diasAteEncerramento": dias_restantes,
            "totalNEs": len(todas_nes),
            "totalNEsAuditadas": len([n for n in todas_nes if n["valor"] > 0]),
            "valorTotalEmpenhado": round(tot_emp, 2)
        },
        "unidades": unidades,
        "creditosEmTela": creditos_em_tela,
        "resumoDiario": resumo_diario,
        "mapaNCs": mapa_ncs,
        "timeline": timeline,
        "topFornecedores": top_fornecedores,
        "topProcessos": top_processos,
        "todasNEs": todas_nes
    }

    js_output = "const DASHBOARD_DATA = " + json.dumps(data_dict, ensure_ascii=False, indent=2) + ";\n\n"
    js_output += """// Utilitários de formatação BRL oficial
function formatBRL(val) {
  if (val === null || val === undefined || isNaN(val)) return "R$ 0,00";
  return new Intl.NumberFormat("pt-BR", {
    style: "currency",
    currency: "BRL",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(val);
}

function formatPercent(val) {
  if (val === null || val === undefined || isNaN(val)) return "0,00%";
  return val.toFixed(2).replace(".", ",") + "%";
}
"""

    targets = [ROOT / "site" / "dashboard-data.js"]
    if (ROOT / "prototipo-v3").exists():
        targets.append(ROOT / "prototipo-v3" / "dashboard-data.js")

    for t in targets:
        t.parent.mkdir(parents=True, exist_ok=True)
        t.write_text(js_output, encoding="utf-8")
        print(f"✓ Sucesso! Gerado {t} ({t.stat().st_size} bytes)")

    print(f"Total NEs: {len(todas_nes)} | Total Créditos em Tela: {len(creditos_em_tela)} | Total NCs Indexadas: {len(mapa_ncs)}")
    return True

if __name__ == "__main__":
    build()
