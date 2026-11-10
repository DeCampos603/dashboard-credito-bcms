# SICOF 3.0 — Sistema Integrado de Inteligência Orçamentária e Financeira
**Base de Apoio Logístico do Exército (Ba Ap Log Ex) & OMDS Subordinadas**

Painel executivo web modular de alta fidelidade para controle, rastreabilidade e governança orçamentária das Organizações Militares Diretamente Subordinadas (OMDS) — **BCMS** (160329/167329), **Ba Ap Log** (160238/167238), **D C Mun** (160246/167246), **BMSA** (160304/167304), **1º D Sup** (160307/167307) e **ECT** (160321/167321). 

O sistema integra o **Aporte do Recurso (Notas de Crédito - NC)** com o **Comprometimento da Despesa (Notas de Empenho - NE)**, credores (CNPJ/CPF) e processos administrativos/licitatórios. O painel é atualizado continuamente via **GitHub Actions** e publicado no **GitHub Pages**.

- **URL Pública Oficial**: [https://DeCampos603.github.io/dashboard-credito-bcms/](https://DeCampos603.github.io/dashboard-credito-bcms/)

---

## 🚀 Inovações da Interface SICOF 3.0

1. **🏆 Ranqueamento Executivo Exclusivo na Visão Consolidada**:
   - Aba dedicada `[ 🏆 Ranqueamento das Unidades ]` visível estritamente na visão global/consolidada.
   - Medalhas heráldicas circulares de alta precisão (🥇 1º, 🥈 2º, 🥉 3º), sem deformações e centralizadas via Flexbox.
   - Benchmark comparativo em 3 dimensões estratégicas: **Taxa de Empenho**, **Taxa de Liquidação** e **Taxa de Pagamento**.
   - Ao selecionar qualquer OMDS individual, o sistema redireciona com fluidez para os créditos específicos daquela unidade.

2. **📊 Exportação Excel (.xlsx) Nativa & Focada (ExcelJS)**:
   - Geração OpenXML (.xlsx) real em tela através da engine `ExcelJS` com fallback `SheetJS`.
   - Exportação dedicada aos **Créditos em Tela** respeitando a ordenação ativa do usuário (por Maior Valor ou Idade do Crédito).
   - Tipagem numérica estrita, AutoFilter nativo nas colunas, cabeçalho executivo Dark Navy (#0B132B), linhas zebradas suaves, painéis congelados e fórmulas analíticas `=SOMA()` e `=MÉDIA()`.

3. **✨ Animações Fluídas a 60 FPS (Aceleração por GPU)**:
   - Micro-interações táteis e feedback elástico em todos os cards de OMDS (`:active { transform: translateY(1px) scale(0.965); }`).
   - Efeito de onda e luz heráldica (*ripple effect*) disparado no ponto de clique do usuário.
   - Transições cinematográficas com *crossfade* nos brasões oficiais e animação elástica de entrada (*crest reveal*).
   - Fundo dinâmico com WebGL Shader GLSL suave a 60fps, com botão de liga/desliga para economia de energia.

4. **🌓 Modo Claro de Alto Contraste (WCAG 2.2 AAA)**:
   - Alternador de tema no topo com paleta executiva refinada: fundo branco puro/slate (#F8FAFC), cartões em relevo sutil (#FFFFFF) e bordas slate nítidas (#CBD5E1), garantindo contraste absoluto sob luz solar direta sem fadiga visual.

5. **🔍 Rastreabilidade Integral (Crédito ➔ Empenho)**:
   - Auditoria Bottom-Up de quase **3.000 Notas de Empenho** emitidas no ano, com busca preditiva por Razão Social, CNPJ, Número da NE ou processo.
   - Detalhamento de cada Nota de Crédito (NC) com histórico completo de movimentações, parcelas e empenhos emitidos contra o crédito.

6. **⚡ Arquitetura Modular e Desacoplada**:
   - `site/index.html` leve e semântico (~35 KB).
   - `site/styles.css` modular (~53 KB) com scrollbar personalizada.
   - `site/dashboard-app.js` desacoplado (~77 KB).
   - `site/dashboard-data.js` gerado dinamicamente via `build_data.py`.

---

## 🛠️ Arquitetura do Pipeline

```
[Google Drive / Tesouro Gerencial]
                │
                ▼
      gerar_dashboard.py
  (ETL, validação matemática e atualização de data/data.json)
                │
                ├─► scripts/pipeline_sentinel.py (Auditoria de 6 OMDS e integridade)
                ├─► build_data.py (Compilação do data bundle site/dashboard-data.js)
                │
                ▼
          Pasta site/ (index.html modular + assets + ExcelJS + shaders)
                │
                ▼
      Deploy GitHub Pages (GitHub Actions)
```

---

## 💻 Execução Local

```powershell
# Instalar dependências
py -3 -m pip install -r requirements.txt

# Executar pipeline de atualização completa
py -3 gerar_dashboard.py

# Iniciar servidor local
py -3 -m http.server 8090 --directory site
```
Acesse no navegador: `http://localhost:8090/`.

---

*Base de Apoio Logístico do Exército — SIAFI / Tesouro Gerencial / SICOF 3.0*
