# SICOF — Sistema Integrado de Controle Orçamentário e Financeiro
**Base de Apoio Logístico do Exército (Ba Ap Log Ex) & OMDS Subordinadas**

Painel web (HTML estático autocontido) para controle integral da execução orçamentária das Organizações Militares Diretamente Subordinadas (OMDS) — **BCMS** (160329/167329), **Ba Ap Log** (160238/167238), **D C Mun** (160246/167246), **BMSA** (160304/167304), **1º D Sup** (160307/167307) e **ECT** (160321/167321) — integrando em tempo real o **Aporte do Recurso (Notas de Crédito - NC)** com o **Comprometimento da Despesa (Notas de Empenho - NE)**, credores (CNPJ) e processos licitatórios. Atualizado diariamente via **GitHub Actions** e hospedado no **GitHub Pages**.

- **URL pública**: `https://DeCampos603.github.io/dashboard-credito-bcms/`
- **NOVA ABA "🔗 Crédito ➔ Empenho" (Rastreabilidade Integral)**:
  - **Visão Top-Down (Crédito > Empenhos)**: Tabela em acordeão expansível vinculando cada Nota de Crédito aos seus respectivos empenhos emitidos, com cálculo automático da taxa de queima orçamentária e semáforo de execução.
  - **Visão Bottom-Up (Auditoria de Empenhos)**: Base consolidada das **2.646 Notas de Empenho** emitidas no ano (R$ 41,55 mi), com busca instantânea por Razão Social, CNPJ, Número da NE, Pregão ou NC de origem.
  - **Raio-X de Fornecedores & Pregões**: Ranking dos 846 fornecedores contratados e principais Pregões Eletrônicos (PE) que movimentam os recursos da Base.
  - **Monitor de Riscos Orçamentários**: Identificação preventiva de créditos parados há mais de 45 dias sem empenho para evitar perdas e estornos no encerramento do exercício financeiro.
- **Drill-down / Modal Enriquecido de NC**: Ao clicar em qualquer Nota de Crédito em qualquer aba do painel, a gaveta modal agora exibe a relação completa das **Notas de Empenho emitidas contra aquele crédito**, credores e valores consumidos.
- **Nova Subaba "Histórico Completo" por Unidade**: Cada OMDS conta com sua própria aba de histórico completo com 4 KPIs dedicados, filtros multicritério e exportação individual formatada para Excel.
- **Aba "Histórico de Notas de Crédito" (Consolidada)**: Visão geral de todas as 1.400+ NCs emitidas e recebidas na Base de Apoio Logístico no exercício corrente (2026).
- **Tema Escuro de Alta Legibilidade (WCAG AAA)**: Suporte dinâmico a tema claro e escuro a 60fps GPU, com tipografia fluida e contraste rigoroso.
- **Exportação em Excel (.xls formatado)**: Download instantâneo de relatórios formatados para todas as visões (em tela, histórico e empenhos).


---

## Como funciona

```
[Sua automação diária] ─► CRÉDITO DISP.xlsx (Google Drive, link público)
                                   │
      GitHub Actions (cron diário) baixa o xlsx ─► gerar_dashboard.py
                                   │  (limpa, filtra BCMS, calcula KPIs, atualiza histórico)
                                   ▼
             site/index.html  ─►  Deploy no GitHub Pages  ─►  URL pública
```

Arquivos:
- `gerar_dashboard.py` — baixa a planilha, gera `site/index.html` (autocontido) e atualiza `data/history.json`.
- `.github/workflows/dashboard.yml` — agenda diária + publicação no Pages.
- `requirements.txt` — dependência (`openpyxl`).

---

## Publicação (passo a passo, ~10 min)

**Pré-requisito:** uma conta no GitHub (grátis) em https://github.com.

### 1) Crie o repositório
- Em github.com → **New repository** → nome ex.: `dashboard-credito-bcms` → visibilidade **Public** → **Create**.

### 2) Envie os arquivos
No PowerShell, dentro desta pasta (`Dashboard-Credito-BCMS`):
```powershell
git init
git add .
git commit -m "Dashboard Crédito BCMS"
git branch -M main
git remote add origin https://github.com/SEU-USUARIO/dashboard-credito-bcms.git
git push -u origin main
```
(Se preferir, dá para **arrastar os arquivos** na interface do GitHub em "uploading an existing file".)

### 3) Habilite o GitHub Pages
- No repositório: **Settings → Pages**.
- Em **Build and deployment → Source**, selecione **GitHub Actions**.

### 4) (Opcional) Esconda o ID do arquivo
- **Settings → Secrets and variables → Actions → New repository secret**
- Nome: `DRIVE_FILE_ID` · Valor: `1Jv546wpWQSFAlep3oLRAg29hVy86iJxJ`
- (Se não fizer isso, o script usa o ID padrão embutido — funciona igual.)

### 5) Rode a primeira vez
- Aba **Actions** → workflow **"Atualizar Dashboard BCMS"** → **Run workflow**.
- Ao terminar (~1–2 min), a URL aparece no passo **"Publicar no GitHub Pages"** e também em **Settings → Pages**.

Pronto! A partir daí ele **atualiza sozinho todo dia** e você compartilha a URL.

---

## Agendamento

Definido em `.github/workflows/dashboard.yml`:
```yaml
schedule:
  - cron: "0 11 * * *"   # 11:00 UTC = 08:00 de Brasília
```
- Para mudar o horário, altere o cron (está em **UTC**; Brasília = UTC−3).
- Ex.: rodar 07:00 de Brasília → `0 10 * * *`. Vários horários? adicione mais linhas `- cron:`.
- O agendamento do GitHub é **melhor esforço** (pode atrasar alguns minutos). Rode manualmente por **Run workflow** quando quiser forçar.
- O commit diário do histórico mantém o repositório "ativo", então o agendamento **não** é desativado pela regra de 60 dias de inatividade do GitHub.

---

## Rodar localmente (teste)

```powershell
py -3 -m pip install -r requirements.txt
py -3 gerar_dashboard.py                       # baixa do Drive
# ou, para testar com um arquivo local:
py -3 gerar_dashboard.py --local "C:\caminho\CRÉDITO DISP.xlsx"
```
Abra `site\index.html` no navegador.

---

## Histórico / tendência

Cada execução grava um snapshot (data + KPIs) em `data/history.json`, versionado no repositório.
O gráfico **"Tendência — Crédito Disponível total"** aparece a partir do **2º dia** de execução
e vai ganhando pontos ao longo do tempo. (É algo que o Power BI, sozinho, não guarda.)

---

## Solução de problemas

| Sintoma | Correção |
|---|---|
| Workflow falha ao **baixar** / arquivo minúsculo | O compartilhamento saiu de "qualquer pessoa com o link". Reative no Drive (Compartilhar → Qualquer pessoa com o link → Leitor). |
| Erro **"Layout mudou: coluna ... não encontrada"** | O cabeçalho da planilha mudou de posição. Ajustar `gerar_dashboard.py` (nomes de coluna). |
| Pages não aparece / 404 | Confirme **Settings → Pages → Source = GitHub Actions** e que o workflow terminou com sucesso. |
| Falha no `git push` do histórico | Confirme que o repo não tem regra de proteção de branch bloqueando o `github-actions[bot]`. |
| Números diferentes da planilha | Confirme o **ID** do arquivo (secret `DRIVE_FILE_ID` ou o padrão no script). |

---

*Fonte: CRÉDITO DISP.xlsx (Tesouro Gerencial / SIAFI). Crédito Disponível = Provisão Recebida − Despesas Empenhadas.*
