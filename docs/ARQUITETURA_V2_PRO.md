# Arquitetura e Governança do Dashboard de Crédito Disponível V2.0 PRO
**Base de Apoio Logístico do Exército (Ba Ap Log Ex)**  
*SIAFI / Tesouro Gerencial — 6 OMDS Vinculadas*  
*Desenvolvido por: 3º Sgt De Campos · Corte Oficial 2026*

---

## 🏛️ 1. Visão Geral e Propósito

O **Dashboard de Crédito Disponível V2.0 PRO** centraliza e audita a execução orçamentária do Batalhão Central de Manutenção e Suprimento (BCMS) e das 5 organizações militares diretamente subordinadas à Base de Apoio Logístico do Exército:

1. **BCMS** (UASGs 160329 - OGU / 167329 - FEx)
2. **Ba Ap Log** (UASGs 160238 - OGU / 167238 - FEx)
3. **D C Mun** (UASGs 160246 - OGU / 167246 - FEx)
4. **BMSA** (UASGs 160304 - OGU / 167304 - FEx)
5. **1º D Sup** (UASGs 160307 - OGU / 167307 - FEx)
6. **ECT** (UASGs 160321 - OGU / 167321 - FEx)

A metodologia segue a equação canônica da contabilidade pública federal:
$$\text{Crédito Disponível Líquido} = \text{Provisão Recebida} - \text{Provisão Concedida} - \text{Despesas Empenhadas}$$

---

## 📊 2. Matriz Auditada das 6 OMDS (Corte Outubro/2026)

| Organização Militar | UASGs | Células Ativas | Provisão Líquida (R$) | Empenhado Real (R$) | Crédito Disponível (R$) | Destaques Orçamentários |
| :--- | :---: | :---: | :---: | :---: | :---: | :--- |
| **BCMS** | 160329 / 167329 | 90 | R$ 12.525.395,43 | R$ 10.883.773,68 | R$ 701.005,02 | NC 000771 (R$ 360k - Blindados AMAN) |
| **Ba Ap Log** | 160238 / 167238 | 121 | R$ 8.051.429,53 | R$ 4.284.215,53 | R$ 574.041,54 | NC 417751 (R$ 190k - Manutenção) |
| **D C Mun** | 160246 / 167246 | 97 | R$ 5.063.379,21 | R$ 4.081.024,64 | R$ 54.052,03 | Fechamento exato SIAFI |
| **BMSA** | 160304 / 167304 | 76 | R$ 5.825.388,48 | R$ 5.131.510,42 | R$ 36.619,90 | Fechamento exato SIAFI |
| **1º D Sup** | 160307 / 167307 | 75 | R$ 22.714.023,26 | R$ 17.344.867,57 | R$ 1.731.103,45 | Maior volume financeiro |
| **ECT** | 160321 / 167321 | 62 | R$ 6.453.012,93 | R$ 5.784.919,38 | R$ 186.114,43 | Fechamento exato SIAFI |
| **TOTAL** | **12 UASGs** | **521** | **R$ 60.632.628,84** | **R$ 47.509.311,22** | **R$ 3.282.936,37** | **100% Auditado e Consistente** |

---

## 🛠️ 3. Pipeline de Dados Unificado

O fluxo de processamento e deploy contínuo opera em quatro etapas atômicas via `python gerar_dashboard.py`:

```
┌─────────────────────────────────┐
│ 1. DOWNLOAD & VALIDAÇÃO BINÁRIA │
│    - Endpoints do Drive         │
│    - Magic Bytes (PK\x03\x04)    │
│    - Fallback: cache persistente│
└────────────────┬────────────────┘
                 │
                 ▼
┌─────────────────────────────────┐
│ 2. ETL SIAFI & ENRIQUECIMENTO   │
│    - Desduplicação de rótulos ND│
│    - Extração de Prazos de NC   │
│    - Isolamento de Centavos     │
└────────────────┬────────────────┘
                 │
                 ▼
┌─────────────────────────────────┐
│ 3. COMPILAÇÃO DASHBOARD V2 PRO  │
│    - Design Tokens (Figma One-X)│
│    - 7 Estados de UI / 60fps GPU│
│    - Setas da Barra de OMDS     │
└────────────────┬────────────────┘
                 │
                 ▼
┌─────────────────────────────────┐
│ 4. SENTINELA PRÉ-DEPLOY         │
│    - scripts/pipeline_sentinel  │
│    - Validação das 6 OMDS       │
│    - Vetos a saldos negativos   │
└─────────────────────────────────┘
```

---

## 🛡️ 4. Regras Críticas de Negócio e Defesa contra Regressões

1. **Anti-Colisão Semântica de Natureza de Despesa:**
   - Naturezas genéricas (ex: `339030` - Material de Consumo) devem sempre ser combinadas ao Plano Interno: `f"{nd_desc} · {pi_desc}"`. Isso impede que recursos distintos (ex: Armários vs Rancho) disputem a mesma linha.
2. **Busca Ativa sem Filtro Cego:**
   - Termos de pesquisa `>= 3 caracteres` desmarcam automaticamente a opção `Ocultar Zeradas`. Notas 100% liquidadas tornam-se visíveis na hora.
3. **Isolamento de Centavos:**
   - Detalhamentos negativos menores que R$ 1,00 (`-0,08`, `-0,51`, `-0,85`) são catalogados como `AJUSTE TÉCNICO DE CENTAVOS` para não poluir a percepção de cortes orçamentários.
4. **Banner Ativo de Defasagem Orçamentária:**
   - Se o diferencial entre o horário atual e `metadata.gerado_em_iso` exceder 24 horas, ou se a base tiver sido carregada em modo contingência offline, a interface aciona um alerta ostensivo no topo.
5. **Setas de Deslocamento das OMDS:**
   - As setas `#omdsScrollPrev` e `#omdsScrollNext` auxiliam a rolagem lateral em telas com espaço restrito e ocultam-se automaticamente em monitores widescreen.
