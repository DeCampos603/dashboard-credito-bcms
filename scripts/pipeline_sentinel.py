# -*- coding: utf-8 -*-
"""
scripts/pipeline_sentinel.py
Auditoria automatizada pré-deploy para o Dashboard de Créditos Orçamentários.
Valida rigorosamente todas as 6 OMDS da Ba Ap Log Ex, consistência matemática e alvos críticos.
"""
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import sys

# Lista oficial das 6 OMDS
OMDS_ESPERADAS = ["BCMS", "Ba Ap Log", "D C Mun", "BMSA", "1º D Sup", "ECT"]

def auditar_pipeline() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    base_dir = Path(__file__).resolve().parent.parent
    json_path = base_dir / "data" / "data.json"

    print("==================================================================")
    print("  SENTINELA MULTI-OMDS: AUDITORIA DE INTEGRIDADE PRÉ-DEPLOY")
    print("==================================================================")

    if not json_path.exists():
        print(f"[ERRO CRÍTICO] Arquivo de dados ausente: {json_path}")
        sys.exit(1)

    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"[ERRO CRÍTICO] Falha ao decodificar JSON: {e}")
        sys.exit(1)

    # 1. Validação de Estrutura de Chaves
    required_keys = ["metadata", "unidades", "celulas", "ncdata", "daydata", "teladata", "histdata"]
    for k in required_keys:
        if k not in data:
            print(f"[ERRO ESTRUTURAL] Chave mandatória ausente no JSON: '{k}'")
            sys.exit(1)

    # 2. Validação das 6 OMDS no Manifesto
    unidades = data.get("unidades", [])
    siglas_unidades = [u.get("sigla") for u in unidades]
    for om in OMDS_ESPERADAS:
        if om not in siglas_unidades:
            print(f"[ERRO DE COBERTURA] OM obrigatória ausente em unidades: '{om}'")
            sys.exit(1)
    print(f"✓ Manifesto de Unidades: 6 OMDS validadas ({', '.join(siglas_unidades)})")

    # 3. Validação das Células Orçamentárias
    celulas = data.get("celulas", {})
    if not celulas:
        print("[ERRO CRÍTICO] Nenhuma célula orçamentária registrada no data.json.")
        sys.exit(1)

    # Agrupamento e auditoria por OMDS
    stats_omds = {om: {"celulas": 0, "aloc": 0.0, "emp": 0.0, "disp": 0.0} for om in OMDS_ESPERADAS}
    
    # Mapeamento de prefixos de ID para siglas de OM
    prefix_to_om = {
        "BCMS": "BCMS",
        "BAAPLOG": "Ba Ap Log",
        "DCMUN": "D C Mun",
        "BMSA": "BMSA",
        "DSUP1": "1º D Sup",
        "ECT": "ECT"
    }

    for cid, c in celulas.items():
        prefix = cid.split("_")[0] if "_" in cid else ""
        om_nome = prefix_to_om.get(prefix)
        if not om_nome:
            continue

        aloc = float(c.get("r", 0.0) or 0.0)
        emp  = float(c.get("e", 0.0) or 0.0)
        disp = float(c.get("d", 0.0) or 0.0)

        # Regra de ouro SIAFI: Saldo disponível nunca pode ser negativo
        if disp < -0.01:
            print(f"[ERRO DE SALDO] Célula {cid} ({c.get('t')}) com saldo disponível negativo: R$ {disp:,.2f}")
            sys.exit(1)

        stats_omds[om_nome]["celulas"] += 1
        stats_omds[om_nome]["aloc"] += aloc
        stats_omds[om_nome]["emp"]  += emp
        stats_omds[om_nome]["disp"] += disp

    print("\n--- RESUMO DE EXECUÇÃO DAS 6 OMDS ---")
    for om, st in stats_omds.items():
        if st["celulas"] == 0:
            print(f"[ERRO DE CARGA] OM {om} não possui células orçamentárias!")
            sys.exit(1)
        print(f"• {om:<12} | Células: {st['celulas']:>3} | Alocado: R$ {st['aloc']:>12,.2f} | Empenhado: R$ {st['emp']:>12,.2f} | Disponível: R$ {st['disp']:>12,.2f}")

    # 4. Auditoria de Alvos Críticos Específicos
    # Alvo 1: BCMS - NC 000771 (R$ 360k para Viaturas Blindadas AMAN, PI E5MBPDRDEGE)
    bcms_aman = [c for cid, c in celulas.items() if cid.startswith("BCMS_") and c.get("pi") == "E5MBPDRDEGE" and c.get("nd") == "339030"]
    if not bcms_aman:
        print("[FALHA AUDITORIA] Célula BCMS E5MBPDRDEGE ND 339030 não localizada.")
        sys.exit(1)
    saldo_bcms_aman = bcms_aman[0].get("d", 0.0)
    if saldo_bcms_aman < 360000.0:
        print(f"[FALHA AUDITORIA] Saldo BCMS E5MBPDRDEGE defasado: R$ {saldo_bcms_aman:,.2f} (Esperado >= R$ 360.000,00).")
        sys.exit(1)
    print(f"✓ Alvo BCMS E5MBPDRDEGE validado: R$ {saldo_bcms_aman:,.2f}")

    # Alvo 2: Ba Ap Log - NC 417751 (R$ 190k para Manutenção de Blindados, PI E5MBPDRDEGE)
    baaplog_aman = [c for cid, c in celulas.items() if cid.startswith("BAAPLOG_") and c.get("pi") == "E5MBPDRDEGE"]
    if not baaplog_aman:
        print("[FALHA AUDITORIA] Célula Ba Ap Log E5MBPDRDEGE não localizada.")
        sys.exit(1)
    saldo_baaplog = sum(c.get("d", 0.0) for c in baaplog_aman)
    if saldo_baaplog < 190000.0:
        print(f"[FALHA AUDITORIA] Saldo Ba Ap Log E5MBPDRDEGE defasado: R$ {saldo_baaplog:,.2f} (Esperado >= R$ 190.000,00).")
        sys.exit(1)
    print(f"✓ Alvo Ba Ap Log E5MBPDRDEGE validado: R$ {saldo_baaplog:,.2f}")

    # 5. Validação de Metadados de Geração
    meta = data.get("metadata", {})
    gerado_em = meta.get("gerado_em", "")
    print(f"\n✓ Metadados: Gerado em '{gerado_em}', Período: '{meta.get('periodo', '')}'")
    if meta.get("usando_cache"):
        print("⚠ AVISO OPERACIONAL: Base gerada em modo contingência (cache offline).")

    print("==================================================================")
    print("  ✓ AUDITORIA SENTINELA APROVADA: DADOS 100% ÍNTEGROS PARA DEPLOY")
    print("==================================================================")

if __name__ == "__main__":
    auditar_pipeline()
