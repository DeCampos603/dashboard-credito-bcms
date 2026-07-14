# Dashboard Crédito Disponível — BCMS

Painel web (HTML estático) do **Crédito Disponível** do BCMS — UASG **160329** (OGU)
e **167329** (Fundo do Exército) — gerado a partir do export do Tesouro Gerencial
`CRÉDITO DISP.xlsx` publicado no Google Drive e **atualizado automaticamente todos os dias**
via **GitHub Actions**, hospedado no **GitHub Pages**.

- **URL pública** (depois de publicar): `https://SEU-USUARIO.github.io/NOME-DO-REPO/`
- Qualquer pessoa com o link abre no navegador, **sem login**.
- Acumula um **histórico diário** → gráfico de tendência do Crédito Disponível.

> ⚠️ **O GitHub Pages é público.** Ao publicar, os números do BCMS (Crédito Disponível,
> empenhado, liquidado, pago, detalhe por NC) ficam **visíveis para qualquer pessoa com o link**
> e podem ser indexados por buscadores. Boa parte desse dado de execução já é público
> (Portal da Transparência / SIAFI), mas a decisão de expor é sua. Páginas com controle de
> acesso só existem em planos pagos (GitHub Enterprise).

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
