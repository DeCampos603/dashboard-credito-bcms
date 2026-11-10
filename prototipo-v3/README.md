# SICOF 3.0 — Protótipo de Nova Geração (Dashboard de Crédito)

> **Modelo Independente e Desacoplado**  
> Interface orientada à identidade heráldica das Unidades (Logos), Transições Fluidas (60fps GPU), Bento Grid Executivo e Conformidade Integral com os Dados do SICOF.

---

## 🏛️ 1. Nomenclaturas Oficiais Atualizadas

Em total conformidade com a governança da Base de Apoio Logístico e OMDS:

| Sigla | Nomenclatura Oficial Canônica | Unidades Gestoras (UGs) |
| :--- | :--- | :--- |
| **Ba Ap Log** | **Base de Apoio Logístico** *(sem "do Exército")* | 160238 (Ordinário) / 167238 (Saúde) |
| **BCMS** | **Batalhão Central de Manutenção e Suprimento** | 160329 (Ordinário) / 167329 (Saúde) |
| **BMSA** | **Batalhão de Manutenção e Suprimento de Armamento** | 160304 (Ordinário) / 167304 (Saúde) |
| **D C Mun** | **Depósito Central de Munições** | 160246 (Ordinário) / 167246 (Saúde) |
| **1º D Sup** | **1º Depósito de Suprimento** | 160307 (Ordinário) / 167307 (Saúde) |
| **ECT** | **Estabelecimento Central de Transportes** | 160321 (Ordinário) / 167321 (Saúde) |

---

## 🎨 2. Nova Experiência do Usuário (UI/UX 3.0)

1. **Vitrine Inicial de Logos (Grid de Brasões):**
   - A página abre com foco visual nos brasões oficiais de todas as Unidades Subordinadas e na opção Consolidada.
   - Micro-interações táteis, efeito hover com elevação 3D (*Tilt*), brilho na cor de acento de cada OM e badge de execução (% de execução orçamentária).

2. **Carregamento Reativo Mediante Clique:**
   - Ao clicar na logo de qualquer OM, a interface transiciona suavemente sem recarregar a página.
   - O **Banner da Unidade** atualiza imediatamente abaixo das logos com o nome oficial completo, UGs ativas e seletor rápido de gestão (**Ordinário OGU** vs **Fundo de Saúde FUSEx**).

3. **Todas as Informações do SICOF Antigo Reorganizadas em Abas:**
   - **📊 Visão Geral & Velocidade:** Bento Grid de 6 KPIs monetários com contadores animados + Curva de aceleração orçamentária (Julho a Outubro/2026).
   - **🔗 Rastreabilidade Crédito ➔ Empenho (Top-Down):** Tabela de Notas de Crédito, finalidades, prazos e saldo, com abertura na **Gaveta Lateral (Drawer)** para detalhamento de empenhos vinculados.
   - **🔍 Auditoria de Empenhos (Bottom-Up):** Base consolidada de Notas de Empenho, credores, CNPJ, pregões e datas.
   - **🏢 Raio-X de Fornecedores & Pregões:** Top 10 fornecedores contratados em volume financeiro e principais Pregões Eletrônicos.
   - **⚠️ Monitor de Riscos Orçamentários:** Semáforo de alerta preventivo para o encerramento do exercício em 52 dias.

4. **Tecnologias Modernas Integradas:**
   - **Fundo Animado WebGL (GLSL Shaders):** Ruído *Simplex Noise* procedural dinâmico com amortecimento inercial (*Spring Lerp*) do cursor e auto-pausa inteligente.
   - **CSS Scroll-Driven Animations:** Barra de progresso de leitura nativa no topo executada 100% no compositor de GPU.
   - **Tokens de Movimento W3C DTCG:** Curva cinética Apple/Linear `cubic-bezier(0.16, 1.0, 0.3, 1.0)`.

---

## 📂 3. Estrutura de Arquivos do Protótipo

```
Dashboard-Credito-BCMS/prototipo-v3/
├── assets/logos/           # Brasões oficiais em alta resolução (BCMS, BAAPLOG, DCMUN, BMSA, 1DSUP, Ect)
├── index.html              # Estrutura semântica com Vitrine de Logos, Banner e Abas
├── styles.css              # Tokens W3C, Glassmorphism e estilos da vitrine
├── motion-background.js    # Shader WebGL GLSL com amortecimento inercial do cursor
├── dashboard-data.js       # Base de dados auditada com nomenclaturas oficiais corrigidas
├── dashboard-app.js        # Lógica de seleção de logos, abas e gaveta lateral
└── README.md               # Esta documentação
```

---

## ⚡ 4. Como Executar

Basta abrir o arquivo [`index.html`](index.html) em qualquer navegador web moderno ou iniciar um servidor local:
```powershell
cd "Dashboard-Credito-BCMS\prototipo-v3"
py -3 -m http.server 8080
```
Acesse: `http://localhost:8080`
