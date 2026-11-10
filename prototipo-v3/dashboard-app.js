/**
 * SICOF 3.0 PROTOTYPE — Application Engine
 * Navegação por Logos, Créditos em Tela (Aba Principal), Resumo Diário e Riscos por Unidade.
 */

document.addEventListener("DOMContentLoaded", () => {
  // 1. Inicializar Fundo Animado WebGL
  const motionBg = new MotionBackground("motion-canvas");

  // 2. Estado Global
  let currentUnitId = "total"; // "total", "BCMS", "BAAPLOG", "DCMUN", "BMSA", "DSUP1", "ECT"
  let currentSubFilter = "all"; // "all", "ordinario", "saude"
  let currentUGFilter = "all";  // "all", "160", "167"
  let activeTab = "tab-creditos-tela";
  let searchTelaQuery = "";
  let searchGlobalQuery = "";

  // Estado de Ordenação Dinâmica por Tabela
  const sortState = {
    tela: { key: "v", dir: "desc", type: "number" },
    nc: { key: "saldo", dir: "desc", type: "number" },
    ne: { key: "data", dir: "desc", type: "date" }
  };

  function parseBRDate(str) {
    if (!str) return 0;
    const parts = str.split("/");
    if (parts.length === 3) {
      return new Date(parts[2], parts[1] - 1, parts[0]).getTime();
    }
    return 0;
  }

  function sortList(list, key, dir, type = "auto") {
    return [...list].sort((a, b) => {
      let valA = a[key];
      let valB = b[key];

      if (valA === undefined || valA === null) valA = "";
      if (valB === undefined || valB === null) valB = "";

      let cmp = 0;
      if (type === "number" || typeof valA === "number" || typeof valB === "number") {
        const numA = parseFloat(valA) || 0;
        const numB = parseFloat(valB) || 0;
        cmp = numA - numB;
      } else if (type === "date" || (typeof valA === "string" && valA.match(/^\d{2}\/\d{2}\/\d{4}$/))) {
        cmp = parseBRDate(valA) - parseBRDate(valB);
      } else {
        cmp = String(valA).localeCompare(String(valB), "pt-BR", { numeric: true, sensitivity: "base" });
      }

      return dir === "asc" ? cmp : -cmp;
    });
  }

  // 3. Referências do DOM
  const logosContainer = document.getElementById("omds-logos-grid");
  const bannerLogo = document.getElementById("unit-banner-logo");
  const bannerNome = document.getElementById("unit-banner-nome");
  const bannerDesc = document.getElementById("unit-banner-desc");
  const bannerTags = document.getElementById("unit-banner-tags");
  const subFilterBtns = document.querySelectorAll(".sub-filter-btn");
  const ugFilterBtns = document.querySelectorAll(".ug-filter-btn");
  const tabNavBtns = document.querySelectorAll(".tab-nav-btn");
  const tabPanes = document.querySelectorAll(".tab-pane");

  // KPIs
  const kpiProv = document.getElementById("kpi-prov-val");
  const kpiEmp = document.getElementById("kpi-emp-val");
  const kpiCred = document.getElementById("kpi-cred-val");
  const kpiLiq = document.getElementById("kpi-liq-val");
  const kpiPag = document.getElementById("kpi-pag-val");
  const kpiConc = document.getElementById("kpi-conc-val");

  const kpiEmpPct = document.getElementById("kpi-emp-pct");
  const kpiEmpBar = document.getElementById("kpi-emp-bar");
  const kpiCredPct = document.getElementById("kpi-cred-pct");
  const kpiCredBar = document.getElementById("kpi-cred-bar");
  const kpiLiqPct = document.getElementById("kpi-liq-pct");
  const kpiLiqBar = document.getElementById("kpi-liq-bar");

  // Elementos de Créditos em Tela
  const tableTelaBody = document.getElementById("tela-table-body");
  const telaCountBadge = document.getElementById("tela-count-badge");
  const telaTotalVal = document.getElementById("tela-total-val");
  const telaDiasMedia = document.getElementById("tela-dias-media");
  const searchTelaInput = document.getElementById("search-tela-input");

  // Elementos do Resumo Diário e Riscos
  const resumoDiarioContainer = document.getElementById("resumo-diario-container");
  const riscosList = document.getElementById("riscos-list");
  const riscosHeaderTitle = document.getElementById("riscos-header-title");

  // Demais Tabelas e Listas
  const tableNCBody = document.getElementById("nc-table-body");
  const tableNEBody = document.getElementById("ne-table-body");
  const searchNEInput = document.getElementById("search-ne-input");
  const neCountBadge = document.getElementById("ne-count-badge");
  let searchNEQuery = "";
  const fornecedoresList = document.getElementById("fornecedores-list");
  const processosList = document.getElementById("processos-list");
  const searchGlobalInput = document.getElementById("global-search-input");
  const drawerBackdrop = document.getElementById("drawer-backdrop");
  const drawerCloseBtn = document.getElementById("drawer-close-btn");
  const btnToggleShader = document.getElementById("btn-toggle-shader");
  const btnExport = document.getElementById("btn-export");
  const btnToggleTheme = document.getElementById("btn-toggle-theme");
  const themeBtnIcon = document.getElementById("theme-btn-icon");
  const themeBtnText = document.getElementById("theme-btn-text");

  // Elementos do Card de Ranqueamento das Unidades
  const rankingContainer = document.getElementById("ranking-items-container");
  const btnToggleRanking = document.getElementById("btn-toggle-ranking-card");
  const rankingSection = document.getElementById("ranking-omds-section");
  const rankingToggleIcon = document.getElementById("ranking-toggle-icon");
  const rankingToggleText = document.getElementById("ranking-toggle-text");
  const rankCritBtns = document.querySelectorAll(".rank-crit-btn");
  let currentRankCriteria = "taxaEmpenho"; // "taxaEmpenho", "taxaLiquidacao", "cred", "prov"

  // Paleta Heráldica Oficial das Organizações Militares
  const UNIT_THEME_COLORS = {
    total:   { accent: "#38bdf8", rgb: "56, 189, 248",  label: "Consolidado Geral" },
    BCMS:    { accent: "#ef4444", rgb: "239, 68, 68",   label: "BCMS" },
    BAAPLOG: { accent: "#3b82f6", rgb: "59, 130, 246",  label: "Ba Ap Log" },
    DCMUN:   { accent: "#0ea5e9", rgb: "14, 165, 233",  label: "D C Mun" },
    BMSA:    { accent: "#8b5cf6", rgb: "139, 92, 246",  label: "BMSA" },
    DSUP1:   { accent: "#10b981", rgb: "16, 185, 129",  label: "1º D Sup" },
    ECT:     { accent: "#f59e0b", rgb: "245, 158, 11",  label: "ECT" }
  };

  function applyUnitColorScheme(unitId) {
    const cfg = UNIT_THEME_COLORS[unitId] || UNIT_THEME_COLORS.total;
    const root = document.documentElement;

    root.style.setProperty("--unit-accent", cfg.accent);
    root.style.setProperty("--unit-accent-rgb", cfg.rgb);
    root.style.setProperty("--unit-glow", `rgba(${cfg.rgb}, 0.45)`);
    root.style.setProperty("--unit-border", `rgba(${cfg.rgb}, 0.55)`);
    root.style.setProperty("--unit-bg-tint", `rgba(${cfg.rgb}, 0.1)`);

    // Sincroniza o shader WebGL de fundo com a cor heráldica da OM
    if (motionBg && typeof motionBg.setAccentColor === "function") {
      motionBg.setAccentColor(cfg.accent);
    }
  }

  function applyTheme(theme) {
    const root = document.documentElement;
    if (theme === "light") {
      root.setAttribute("data-theme", "light");
      if (themeBtnIcon) themeBtnIcon.textContent = "🌙";
      if (themeBtnText) themeBtnText.textContent = "Modo Escuro";
      try { localStorage.setItem("sicof_theme", "light"); } catch(e){}
    } else {
      root.removeAttribute("data-theme");
      if (themeBtnIcon) themeBtnIcon.textContent = "☀️";
      if (themeBtnText) themeBtnText.textContent = "Tema Branco (Alto Contraste)";
      try { localStorage.setItem("sicof_theme", "dark"); } catch(e){}
    }
  }

  // ========================================================================
  // 1. RENDERIZAÇÃO DA VITRINE INICIAL DE LOGOS DAS UNIDADES & FEEDBACK TÁTIL
  // ========================================================================
  function createCardRipple(event, card) {
    if (!card) return;
    const rect = card.getBoundingClientRect();
    const circle = document.createElement("span");
    const diameter = Math.max(rect.width, rect.height);
    const radius = diameter / 2;

    const clientX = event && event.clientX ? event.clientX : rect.left + rect.width / 2;
    const clientY = event && event.clientY ? event.clientY : rect.top + rect.height / 2;

    circle.style.width = circle.style.height = `${diameter}px`;
    circle.style.left = `${clientX - rect.left - radius}px`;
    circle.style.top = `${clientY - rect.top - radius}px`;
    circle.classList.add("card-ripple-wave");

    const omRgb = card.style.getPropertyValue("--om-rgb") || card.style.getPropertyValue("--rank-rgb") || "56, 189, 248";
    circle.style.background = `radial-gradient(circle, rgba(${omRgb}, 0.75) 0%, rgba(${omRgb}, 0.2) 55%, transparent 75%)`;

    card.appendChild(circle);
    setTimeout(() => {
      circle.remove();
    }, 600);
  }

  function renderLogosShowcase() {
    if (!logosContainer) return;

    logosContainer.innerHTML = DASHBOARD_DATA.unidades
      .map((unit) => {
        const isActive = unit.id === currentUnitId;
        const activeClass = isActive ? "is-active" : "";
        const badgeColorClass =
          unit.taxaEmpenho >= 98
            ? "badge-emerald"
            : unit.taxaEmpenho >= 90
            ? "badge-cyan"
            : "badge-amber";

        const logoHtml = unit.logo
          ? `<img src="${unit.logo}" alt="Brasão de ${unit.sigla}" class="omds-crest-img">`
          : `<div class="omds-consolidado-icon">🏛️</div>`;

        return `
          <div class="logo-card ${activeClass}" data-unit-id="${unit.id}">
            <div class="logo-img-wrap">
              ${logoHtml}
            </div>
            <div class="logo-card-sigla">${unit.sigla}</div>
            <div class="logo-card-nome" title="${unit.nome}">${unit.nome}</div>
            <div class="logo-card-badge ${badgeColorClass}">
              ${unit.taxaEmpenho.toFixed(1)}% Empenhado
            </div>
          </div>
        `;
      })
      .join("");

    logosContainer.querySelectorAll(".logo-card").forEach((card) => {
      card.addEventListener("click", (e) => {
        createCardRipple(e, card);
        const unitId = card.getAttribute("data-unit-id");
        selectUnit(unitId);
      });
    });
  }

  // ========================================================================
  // 1.1 RENDERIZAÇÃO DO PLACAR & RANQUEAMENTO DAS UNIDADES (6 OMDS)
  // ========================================================================
  function renderRankingCard() {
    if (!rankingContainer) return;

    const units = DASHBOARD_DATA.unidades.filter((u) => u.id !== "total");

    // Ordenação dinâmica de acordo com o critério selecionado
    units.sort((a, b) => {
      if (currentRankCriteria === "cred") {
        return a.cred - b.cred; // Menor saldo primeiro (mais eficiente no empenho)
      } else if (currentRankCriteria === "prov") {
        return b.prov - a.prov; // Maior volume de aporte
      } else if (currentRankCriteria === "taxaLiquidacao") {
        return b.taxaLiquidacao - a.taxaLiquidacao;
      } else {
        return b.taxaEmpenho - a.taxaEmpenho;
      }
    });

    rankingContainer.innerHTML = units
      .map((u, index) => {
        const pos = index + 1;
        let podiumClass = "rank-podium-other";
        let medalEmoji = "🎖️";

        if (pos === 1) {
          podiumClass = "rank-podium-1";
          medalEmoji = "🥇";
        } else if (pos === 2) {
          podiumClass = "rank-podium-2";
          medalEmoji = "🥈";
        } else if (pos === 3) {
          podiumClass = "rank-podium-3";
          medalEmoji = "🥉";
        }

        const isSelected = u.id === currentUnitId;
        const selectedClass = isSelected ? "is-selected" : "";
        const themeCfg = UNIT_THEME_COLORS[u.id] || { accent: u.accent, rgb: "56, 189, 248" };

        let criteriaValueBadge = "";
        if (currentRankCriteria === "taxaLiquidacao") {
          criteriaValueBadge = `${u.taxaLiquidacao.toFixed(1)}% Liquidado`;
        } else if (currentRankCriteria === "cred") {
          criteriaValueBadge = `Saldo: ${formatBRL(u.cred)}`;
        } else if (currentRankCriteria === "prov") {
          criteriaValueBadge = `Total: ${formatBRL(u.prov)}`;
        } else {
          criteriaValueBadge = `${u.taxaEmpenho.toFixed(2)}% Comprometido`;
        }

        return `
          <div class="rank-item-card ${selectedClass}" data-unit-id="${u.id}" style="--rank-accent: ${themeCfg.accent}; --rank-rgb: ${themeCfg.rgb};" title="Clique para selecionar ${u.sigla} e visualizar seus créditos">
            <div class="rank-card-header-row">
              <div class="rank-crest-and-info">
                <div class="rank-crest-box">
                  <img src="${u.logo}" alt="Brasão ${u.sigla}" style="max-width:100%; max-height:100%; object-fit:contain; display:block;">
                </div>
                <div class="rank-unit-text-wrap">
                  <div class="rank-unit-sigla">${u.sigla}</div>
                  <div class="rank-unit-nome" title="${u.nome}">${u.nome}</div>
                </div>
              </div>

              <div class="rank-badge-wrap">
                <div class="rank-podium-badge ${podiumClass}">
                  <span class="medal-icon">${medalEmoji}</span>
                  <span class="medal-rank">${pos}º</span>
                </div>
              </div>
            </div>

            <div class="rank-progress-box">
              <div class="rank-progress-labels">
                <span style="color: var(--text-muted); font-size: 0.7rem; font-weight: 600;">Comprometimento da Dotação</span>
                <span class="rank-pct-badge">${criteriaValueBadge}</span>
              </div>
              <div class="rank-progress-track">
                <div class="rank-progress-fill" style="width: ${Math.min(100, Math.max(0, u.taxaEmpenho))}%;"></div>
              </div>
            </div>

            <div class="rank-metrics-footer">
              <div class="rank-metric-item">
                <span class="rank-metric-lbl">Empenhado</span>
                <span class="rank-metric-val">${formatBRL(u.emp)}</span>
              </div>
              <div class="rank-metric-item">
                <span class="rank-metric-lbl">Disponível</span>
                <span class="rank-metric-val" style="color: var(--text-emerald);">${formatBRL(u.cred)}</span>
              </div>
              <div class="rank-metric-item">
                <span class="rank-metric-lbl">Liquidado</span>
                <span class="rank-metric-val" style="color: var(--text-amber);">${u.taxaLiquidacao.toFixed(1)}%</span>
              </div>
            </div>
          </div>
        `;
      })
      .join("");

    // Click handler para os cards de ranqueamento com feedback tátil ripple
    rankingContainer.querySelectorAll(".rank-item-card").forEach((card) => {
      card.addEventListener("click", (e) => {
        createCardRipple(e, card);
        const unitId = card.getAttribute("data-unit-id");
        selectUnit(unitId);
        if (typeof window.switchTab === "function") {
          window.switchTab("tab-creditos-tela");
        }
      });
    });

    // Renderizar Tabela Comparativa Consolidada no pé da aba de ranqueamento
    const rankingTableBody = document.getElementById("ranking-table-body");
    if (rankingTableBody) {
      rankingTableBody.innerHTML = units
        .map((u, index) => {
          const pos = index + 1;
          let posBadge = `<span class="badge-pill" style="font-weight:800; font-family:var(--font-mono);">${pos}º</span>`;
          if (pos === 1) posBadge = `<span class="badge-pill" style="background:rgba(234,179,8,0.2); color:#fef08a; border:1px solid #eab308; font-weight:800;">🥇 1º</span>`;
          else if (pos === 2) posBadge = `<span class="badge-pill" style="background:rgba(203,213,225,0.2); color:#f8fafc; border:1px solid #cbd5e1; font-weight:800;">🥈 2º</span>`;
          else if (pos === 3) posBadge = `<span class="badge-pill" style="background:rgba(217,119,6,0.2); color:#fed7aa; border:1px solid #d97706; font-weight:800;">🥉 3º</span>`;

          return `
            <tr data-unit-id="${u.id}">
              <td style="text-align: center;">${posBadge}</td>
              <td>
                <div style="display:flex; align-items:center; gap:0.65rem;">
                  <img src="${u.logo}" alt="" style="width:24px; height:28px; object-fit:contain; flex-shrink:0;">
                  <div>
                    <strong style="color:var(--text-primary); font-size:0.85rem;">${u.sigla}</strong>
                    <div style="font-size:0.7rem; color:var(--text-muted);">${u.nome}</div>
                  </div>
                </div>
              </td>
              <td style="text-align: right; font-family: var(--font-mono); font-weight: 700;">${formatBRL(u.prov)}</td>
              <td style="text-align: right; font-family: var(--font-mono); font-weight: 700; color: var(--text-accent);">${formatBRL(u.emp)}</td>
              <td style="text-align: right; font-family: var(--font-mono); font-weight: 700; color: var(--text-emerald);">${formatBRL(u.cred)}</td>
              <td style="text-align: right; font-family: var(--font-mono); font-weight: 700;">${formatBRL(u.liq)}</td>
              <td style="text-align: right; font-family: var(--font-mono); font-weight: 800; color: var(--text-accent);">${u.taxaEmpenho.toFixed(2)}%</td>
              <td style="text-align: right; font-family: var(--font-mono); font-weight: 800; color: var(--text-amber);">${u.taxaLiquidacao.toFixed(1)}%</td>
              <td style="text-align: center;">
                <button class="btn-select-unit-rank" data-unit-id="${u.id}" style="padding: 0.3rem 0.75rem; font-size: 0.72rem; border-radius: 8px; background: rgba(56, 189, 248, 0.12); border: 1px solid rgba(56, 189, 248, 0.35); color: #38bdf8; cursor: pointer; font-weight: 700; transition: all 0.2s ease;" title="Selecionar ${u.sigla} e abrir seus Créditos em Tela">
                  Ver Créditos ➔
                </button>
              </td>
            </tr>
          `;
        })
        .join("");

      rankingTableBody.querySelectorAll(".btn-select-unit-rank").forEach((btn) => {
        btn.addEventListener("click", (e) => {
          e.stopPropagation();
          const unitId = btn.getAttribute("data-unit-id");
          selectUnit(unitId);
          if (typeof window.switchTab === "function") {
            window.switchTab("tab-creditos-tela");
          }
        });
      });
    }
  }

  // ========================================================================
  // 2. SELEÇÃO REATIVA DE UNIDADE
  // ========================================================================
  function selectUnit(unitId) {
    currentUnitId = unitId;
    currentSubFilter = "all";
    currentUGFilter = "all";

    // Aplica esquema cromático heráldico e sincroniza com o shader WebGL
    applyUnitColorScheme(unitId);

    logosContainer.querySelectorAll(".logo-card").forEach((card) => {
      card.classList.toggle("is-active", card.getAttribute("data-unit-id") === unitId);
    });

    // Visibilidade estrita da aba de Ranqueamento: SOMENTE na aba CONSOLIDADO (total)
    const rankingTabBtn = document.querySelector('.tab-nav-btn[data-tab="tab-ranqueamento"]');
    if (rankingTabBtn) {
      if (unitId === "total") {
        rankingTabBtn.style.display = "";
      } else {
        rankingTabBtn.style.display = "none";
        // Se o usuário estiver na aba de ranqueamento e mudar para uma unidade específica, redireciona para Créditos em Tela
        if (activeTab === "tab-ranqueamento" && typeof window.switchTab === "function") {
          window.switchTab("tab-creditos-tela");
        }
      }
    }

    subFilterBtns.forEach((btn) => {
      btn.classList.toggle("is-active", btn.getAttribute("data-filter") === "all");
    });

    ugFilterBtns.forEach((btn) => {
      btn.classList.toggle("is-active", btn.getAttribute("data-ug-filter") === "all");
    });

    updateUnitBanner();
    renderRankingCard();
    updateKPIs();
    renderCreditosEmTela();
    renderResumoDiario();
    renderVelocityChart();
    renderNCTable();
    renderNETable();
    renderFornecedoresEProcessos();
    renderRiscosSeparadosPorUnidade();

    // Dispara micro-transição GPU suave no Banner, Brasão e Painel de Conteúdo Ativo
    const banner = document.querySelector(".unit-overview-banner");
    if (banner) {
      banner.classList.remove("unit-crossfade");
      void banner.offsetWidth; // Força reflow para reiniciar animação
      banner.classList.add("unit-crossfade");
    }

    const bannerLogoWrap = document.querySelector(".unit-banner-logo-wrap");
    if (bannerLogoWrap) {
      bannerLogoWrap.classList.remove("crest-reveal");
      void bannerLogoWrap.offsetWidth;
      bannerLogoWrap.classList.add("crest-reveal");
    }

    const activePane = document.querySelector(".tab-pane.is-active");
    if (activePane) {
      activePane.classList.remove("unit-content-enter");
      void activePane.offsetWidth;
      activePane.classList.add("unit-content-enter");
    }
  }

  // ========================================================================
  // 3. BANNER DA UNIDADE SELECIONADA
  // ========================================================================
  function updateUnitBanner() {
    const unit = DASHBOARD_DATA.unidades.find((u) => u.id === currentUnitId) || DASHBOARD_DATA.unidades[0];

    if (bannerLogo) {
      bannerLogo.innerHTML = unit.logo
        ? `<img src="${unit.logo}" alt="Brasão de ${unit.sigla}" style="max-width:100%; max-height:100%; object-fit:contain; display:block;">`
        : `<div class="omds-consolidado-icon" style="width:72px; height:72px; font-size:2.4rem;">🏛️</div>`;
    }

    if (bannerNome) bannerNome.textContent = `${unit.sigla} • ${unit.nome}`;
    if (bannerDesc) bannerDesc.textContent = unit.descricao;

    if (bannerTags) {
      const isConsolidado = unit.id === "total";
      const rankingBannerBtn = isConsolidado
        ? `<button id="banner-btn-ranking" class="tag-pill" style="border-color: rgba(234, 179, 8, 0.5); color: #fef08a; background: rgba(234, 179, 8, 0.15); cursor: pointer; font-weight: 700;" title="Abrir Placar e Ranqueamento das 6 OMDS">🏆 Placar de Ranqueamento</button>`
        : "";

      bannerTags.innerHTML = `
        <span class="tag-pill" style="border-color: rgba(56, 189, 248, 0.3); color: var(--text-accent);">UGs: ${unit.ugs.join(" | ")}</span>
        <span class="tag-pill" style="border-color: rgba(16, 185, 129, 0.3); color: var(--text-emerald);">Execução: ${unit.taxaEmpenho.toFixed(2)}%</span>
        <span class="tag-pill">Exercício 2026</span>
        ${rankingBannerBtn}
      `;

      const btnRank = document.getElementById("banner-btn-ranking");
      if (btnRank) {
        btnRank.addEventListener("click", () => {
          if (typeof window.switchTab === "function") {
            window.switchTab("tab-ranqueamento");
          }
        });
      }
    }

    const subFilterWrap = document.querySelector(".unit-sub-filters");
    if (subFilterWrap) {
      subFilterWrap.style.display = unit.sub ? "flex" : "none";
    }
  }

  // ========================================================================
  // 4. ABA PRINCIPAL: CRÉDITOS EM TELA (NOTAS ATIVAS, PRAZOS E UGS 160/167)
  // ========================================================================
  function renderCreditosEmTela() {
    if (!tableTelaBody) return;

    let items = DASHBOARD_DATA.creditosEmTela;

    // Filtro de Unidade
    if (currentUnitId !== "total") {
      items = items.filter((c) => c.om === currentUnitId);
    }

    // Filtro por UG (160 vs 167)
    if (currentUGFilter !== "all") {
      items = items.filter((c) => c.u === currentUGFilter);
    }

    // Filtro por Busca
    if (searchTelaQuery.trim() !== "") {
      const q = searchTelaQuery.toLowerCase();
      items = items.filter(
        (c) =>
          c.nc.toLowerCase().includes(q) ||
          c.nd.includes(q) ||
          c.ndnome.toLowerCase().includes(q) ||
          c.pi.toLowerCase().includes(q) ||
          c.obj.toLowerCase().includes(q) ||
          c.uasg.includes(q)
      );
    }

    // Totalizadores
    const totalVal = items.reduce((acc, c) => acc + c.v, 0);
    const mediaDias = items.length > 0 ? Math.round(items.reduce((acc, c) => acc + c.dias, 0) / items.length) : 0;

    if (telaCountBadge) telaCountBadge.textContent = `${items.length} itens`;
    if (telaTotalVal) telaTotalVal.textContent = formatBRL(totalVal);
    if (telaDiasMedia) telaDiasMedia.textContent = `${mediaDias} dias`;

    // Ordenação Dinâmica Clicável
    items = sortList(items, sortState.tela.key, sortState.tela.dir, sortState.tela.type);

    if (items.length === 0) {
      tableTelaBody.innerHTML = `
        <tr><td colspan="8" style="text-align: center; padding: 2.5rem; color: var(--text-muted);">
          Nenhum crédito em tela encontrado para os filtros selecionados.
        </td></tr>
      `;
      return;
    }

    tableTelaBody.innerHTML = items
      .map((c) => {
        const isUg167 = c.u === "167";
        const ugBadgeClass = isUg167 ? "badge-emerald" : "badge-cyan";
        const ugLabel = isUg167 ? `UG ${c.uasg} (Saúde)` : `UG ${c.uasg} (Ordinário)`;

        const diasBadgeClass =
          c.dias > 60 ? "badge-rose" : c.dias > 30 ? "badge-amber" : "badge-emerald";

        return `
          <tr data-nc="${c.nc}">
            <td><span class="nc-code-pill">${c.nc}</span></td>
            <td><span class="badge-pill ${ugBadgeClass}">${ugLabel}</span></td>
            <td><strong style="color: var(--text-amber); font-size: 0.78rem;">${c.prazo}</strong></td>
            <td><span class="badge-pill ${diasBadgeClass}">${c.dias} dias</span></td>
            <td>
              <div style="font-family: var(--font-mono); font-weight: 700; color: var(--text-primary); font-size: 0.78rem;">${c.nd}</div>
              <div style="font-size: 0.7rem; color: var(--text-muted);">${c.ndnome}</div>
            </td>
            <td>
              <span style="font-family: var(--font-mono); font-size: 0.75rem; color: var(--text-accent);">${c.pi}</span>
            </td>
            <td>
              <div class="finalidade-truncate" title="${c.obj}">${c.obj || "—"}</div>
            </td>
            <td style="text-align: right; font-family: var(--font-mono); font-weight: 800; color: var(--text-emerald); font-size: 0.88rem;">
              ${formatBRL(c.v)}
            </td>
          </tr>
        `;
      })
      .join("");

    tableTelaBody.querySelectorAll("tr[data-nc]").forEach((row) => {
      row.addEventListener("click", () => openDrawer(row.getAttribute("data-nc")));
    });
  }

  // ========================================================================
  // 5. ABA 2: RESUMO DIÁRIO DE MOVIMENTAÇÕES (O QUE ACONTECEU NAS UNIDADES)
  // ========================================================================
  function renderResumoDiario() {
    if (!resumoDiarioContainer) return;

    let dayDataMap = DASHBOARD_DATA.resumoDiario;
    let unitsToShow = currentUnitId === "total"
      ? ["BCMS", "BAAPLOG", "DCMUN", "BMSA", "DSUP1", "ECT"]
      : [currentUnitId];

    let html = "";

    unitsToShow.forEach((uId) => {
      const unit = DASHBOARD_DATA.unidades.find((u) => u.id === uId);
      const days = dayDataMap[uId] || [];

      if (days.length === 0) return;

      html += `
        <div class="risk-unit-block">
          <div class="risk-unit-header">
            <div style="display: flex; align-items: center; gap: 0.75rem;">
              <span style="font-size: 1.25rem;">📅</span>
              <strong style="font-size: 1rem; color: var(--text-primary);">${unit ? unit.nome : uId}</strong>
            </div>
            <span class="badge-pill badge-cyan">${days.length} dias registrados</span>
          </div>
          <div>
            ${days
              .map((d) => {
                const ncsCount = d.n || (d.ncs ? d.ncs.length : 0);
                return `
                <div class="day-card">
                  <div class="day-card-header">
                    <div>
                      <strong style="font-size: 0.9rem; color: var(--text-accent);">Data: ${d.d}</strong>
                      <span style="font-size: 0.75rem; color: var(--text-muted); margin-left: 0.5rem;">(${ncsCount} movimentações)</span>
                    </div>
                    <div style="display: flex; gap: 0.85rem; font-family: var(--font-mono); font-size: 0.8rem;">
                      <span style="color: var(--text-emerald);">Recebido: ${formatBRL(d.rec)}</span>
                      <span style="color: var(--text-rose);">Reduzido: ${formatBRL(d.red)}</span>
                      <span style="color: var(--text-primary); font-weight: 700;">Líquido: ${formatBRL(d.liq)}</span>
                    </div>
                  </div>
                  <div>
                    ${(d.ncs || [])
                      .map(
                        (nc) => `
                      <div class="nc-item-clickable" data-nc-code="${nc[0]}" data-om="${uId}" title="Clique para ver origem e histórico detalhado da NC">
                        <div style="display: flex; gap: 0.5rem; align-items: center;">
                          <span class="nc-code-pill">${nc[0]}</span>
                          <span class="badge-pill ${nc[1] === '167' ? 'badge-emerald' : 'badge-cyan'}">UG ${nc[1]}</span>
                          <span style="color: var(--text-secondary); font-weight: 600;">${nc[2]}</span>
                        </div>
                        <div style="display: flex; gap: 1rem; align-items: center;">
                          <span style="max-width: 300px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--text-muted);" title="${nc[4]}">${nc[4]}</span>
                          <strong style="font-family: var(--font-mono); color: ${nc[3] >= 0 ? 'var(--text-emerald)' : 'var(--text-rose)'};">${formatBRL(nc[3])}</strong>
                          <span class="nc-action-hint">🔍 Ver Detalhes</span>
                        </div>
                      </div>
                    `
                      )
                      .join("")}
                  </div>
                </div>
              `;
              })
              .join("")}
          </div>
        </div>
      `;
    });

    resumoDiarioContainer.innerHTML = html || `<div style="text-align:center; padding:2rem; color:var(--text-muted);">Nenhum histórico diário disponível.</div>`;

    // Interatividade: clicar na NC para abrir o Raio-X completo
    resumoDiarioContainer.querySelectorAll(".nc-item-clickable").forEach((item) => {
      item.addEventListener("click", () => {
        const ncCode = item.getAttribute("data-nc-code");
        const om = item.getAttribute("data-om");
        openDrawer(ncCode, om);
      });
    });
  }

  // ========================================================================
  // 6. ABA 7: MONITOR DE RISCOS ORÇAMENTÁRIOS (SEPARADO RIGOROSAMENTE POR UNIDADE)
  // ========================================================================
  function renderRiscosSeparadosPorUnidade() {
    if (!riscosList) return;

    let unitsToShow = currentUnitId === "total"
      ? ["BCMS", "BAAPLOG", "DCMUN", "BMSA", "DSUP1", "ECT"]
      : [currentUnitId];

    if (riscosHeaderTitle) {
      if (currentUnitId === "total") {
        riscosHeaderTitle.textContent = "Monitor de Riscos Orçamentários — Visão Consolidada por Unidade";
      } else {
        const u = DASHBOARD_DATA.unidades.find((unit) => unit.id === currentUnitId);
        riscosHeaderTitle.textContent = `Monitor de Riscos Orçamentários — ${u ? u.nome : currentUnitId}`;
      }
    }

    let html = `
      <div class="bento-card" style="border-color: rgba(244, 63, 94, 0.3); margin-bottom: 1.5rem;">
        <div style="display: flex; gap: 1rem; align-items: center;">
          <span style="font-size: 2.2rem;">🚨</span>
          <div>
            <h4 style="font-size: 1rem; font-weight: 700; color: var(--text-rose);">Alerta Institucional: 52 Dias para o Encerramento do Exercício Financeiro</h4>
            <p style="font-size: 0.8rem; color: var(--text-secondary);">
              Créditos orçamentários parados há mais de 45 dias ou sem empenho até a virada do SIAFI serão estornados às autoridades concedentes.
            </p>
          </div>
        </div>
      </div>
    `;

    unitsToShow.forEach((uId) => {
      const unit = DASHBOARD_DATA.unidades.find((u) => u.id === uId);
      const telaItems = DASHBOARD_DATA.creditosEmTela.filter((c) => c.om === uId);

      // Riscos críticos: dias > 45 ou saldo alto parado
      const criticos = telaItems.filter((c) => c.dias > 45 || c.v > 10000);
      const saldoParadoTotal = telaItems.reduce((acc, c) => acc + c.v, 0);

      const severityBadge =
        saldoParadoTotal > 300000
          ? "badge-rose"
          : saldoParadoTotal > 50000
          ? "badge-amber"
          : "badge-emerald";

      html += `
        <div class="risk-unit-block">
          <div class="risk-unit-header">
            <div style="display: flex; align-items: center; gap: 0.85rem;">
              ${unit && unit.logo ? `<img src="${unit.logo}" style="width:36px; height:36px; object-fit:contain;">` : `<span style="font-size:1.5rem;">🏛️</span>`}
              <div>
                <strong style="font-size: 1.05rem; color: var(--text-primary);">${unit ? unit.nome : uId}</strong>
                <div style="font-size: 0.75rem; color: var(--text-muted);">UGs: ${unit ? unit.ugs.join(" | ") : uId}</div>
              </div>
            </div>
            <div style="display: flex; gap: 0.75rem; align-items: center;">
              <span class="badge-pill ${severityBadge}">Saldo em Tela: ${formatBRL(saldoParadoTotal)}</span>
              <span class="badge-pill" style="background: rgba(255,255,255,0.06); color: var(--text-secondary);">${criticos.length} alertas</span>
            </div>
          </div>

          <div>
            ${
              criticos.length > 0
                ? criticos
                    .slice(0, 5)
                    .map((c) => {
                      const isUg167 = c.u === "167";
                      return `
                      <div class="empenho-mini-card" style="border-color: ${c.dias > 60 ? 'rgba(244,63,94,0.3)' : 'rgba(245,158,11,0.3)'};">
                        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.35rem;">
                          <div style="display: flex; gap: 0.5rem; align-items: center;">
                            <span class="nc-code-pill">${c.nc}</span>
                            <span class="badge-pill ${isUg167 ? 'badge-emerald' : 'badge-cyan'}">UG ${c.u}</span>
                            <span class="badge-pill ${c.dias > 60 ? 'badge-rose' : 'badge-amber'}">${c.dias} dias parado</span>
                          </div>
                          <span style="font-family: var(--font-mono); font-weight: 800; color: var(--text-rose); font-size: 0.85rem;">${formatBRL(c.v)}</span>
                        </div>
                        <div style="font-size: 0.78rem; color: var(--text-secondary); margin-bottom: 0.25rem;">${c.obj || c.ndnome}</div>
                        <div style="font-size: 0.72rem; color: var(--text-muted); display: flex; justify-content: space-between;">
                          <span>ND: ${c.nd} • PI: ${c.pi}</span>
                          <span style="color: var(--text-amber); font-weight: 700;">Prazo limite: ${c.prazo}</span>
                        </div>
                      </div>
                    `;
                    })
                    .join("")
                : `<div style="font-size: 0.8rem; color: var(--text-emerald); padding: 0.5rem 0;">Nenhum crédito em situação de risco crítico para esta unidade.</div>`
            }
          </div>
        </div>
      `;
    });

    riscosList.innerHTML = html;
  }

  // ========================================================================
  // 7. ABA 3: VISÃO GERAL (KPIS & CURVA DE VELOCIDADE)
  // ========================================================================
  function getActiveMetrics() {
    const unit = DASHBOARD_DATA.unidades.find((u) => u.id === currentUnitId) || DASHBOARD_DATA.unidades[0];

    if (unit.sub && currentSubFilter !== "all") {
      const sub = unit.sub[currentSubFilter];
      if (sub) {
        return {
          prov: sub.prov,
          emp: sub.emp,
          cred: sub.cred,
          liq: sub.liq,
          pag: sub.pag,
          conc: 0.0,
          taxaEmpenho: sub.taxaEmp,
          taxaLiquidacao: (sub.liq / sub.emp) * 100,
          taxaPagamento: (sub.pag / sub.liq) * 100
        };
      }
    }

    return {
      prov: unit.prov,
      emp: unit.emp,
      cred: unit.cred,
      liq: unit.liq,
      pag: unit.pag,
      conc: unit.conc,
      taxaEmpenho: unit.taxaEmpenho,
      taxaLiquidacao: unit.taxaLiquidacao,
      taxaPagamento: unit.taxaPagamento
    };
  }

  function updateKPIs() {
    const data = getActiveMetrics();

    animateNumber(kpiProv, data.prov);
    animateNumber(kpiEmp, data.emp);
    animateNumber(kpiCred, data.cred);
    animateNumber(kpiLiq, data.liq);
    animateNumber(kpiPag, data.pag);
    animateNumber(kpiConc, data.conc);

    const empPct = data.prov > 0 ? (data.emp / data.prov) * 100 : 0;
    const credPct = data.prov > 0 ? (data.cred / data.prov) * 100 : 0;
    const liqPct = data.emp > 0 ? (data.liq / data.emp) * 100 : 0;

    if (kpiEmpPct) kpiEmpPct.textContent = formatPercent(empPct);
    if (kpiEmpBar) kpiEmpBar.style.width = `${Math.min(empPct, 100)}%`;

    if (kpiCredPct) kpiCredPct.textContent = formatPercent(credPct);
    if (kpiCredBar) kpiCredBar.style.width = `${Math.min(credPct, 100)}%`;

    if (kpiLiqPct) kpiLiqPct.textContent = formatPercent(liqPct);
    if (kpiLiqBar) kpiLiqBar.style.width = `${Math.min(liqPct, 100)}%`;
  }

  function animateNumber(element, targetValue) {
    if (!element) return;
    const duration = 350;
    const startTime = performance.now();

    function step(now) {
      const progress = Math.min((now - startTime) / duration, 1);
      const ease = 1 - Math.pow(1 - progress, 3);
      element.textContent = formatBRL(targetValue * ease);

      if (progress < 1) {
        requestAnimationFrame(step);
      } else {
        element.textContent = formatBRL(targetValue);
      }
    }
    requestAnimationFrame(step);
  }

  function renderVelocityChart() {
    const container = document.getElementById("chart-container");
    if (!container) return;

    const timeline = DASHBOARD_DATA.timeline;
    const width = 820;
    const height = 260;
    const pad = { top: 25, right: 25, bottom: 35, left: 70 };

    const maxVal = Math.max(...timeline.map((d) => d.prov)) * 1.08;
    const chartW = width - pad.left - pad.right;
    const chartH = height - pad.top - pad.bottom;

    const getX = (i) => pad.left + (i / (timeline.length - 1)) * chartW;
    const getY = (val) => pad.top + chartH - (val / maxVal) * chartH;

    let provPath = "";
    let empPath = "";
    let liqPath = "";

    timeline.forEach((d, i) => {
      const x = getX(i);
      const yP = getY(d.prov);
      const yE = getY(d.emp);
      const yL = getY(d.liq);

      if (i === 0) {
        provPath += `M ${x} ${yP}`;
        empPath += `M ${x} ${yE}`;
        liqPath += `M ${x} ${yL}`;
      } else {
        const prevX = getX(i - 1);
        const cpX = prevX + (x - prevX) / 2;
        provPath += ` C ${cpX} ${getY(timeline[i - 1].prov)}, ${cpX} ${yP}, ${x} ${yP}`;
        empPath += ` C ${cpX} ${getY(timeline[i - 1].emp)}, ${cpX} ${yE}, ${x} ${yE}`;
        liqPath += ` C ${cpX} ${getY(timeline[i - 1].liq)}, ${cpX} ${yL}, ${x} ${yL}`;
      }
    });

    const areaPath = `${empPath} L ${getX(timeline.length - 1)} ${pad.top + chartH} L ${pad.left} ${pad.top + chartH} Z`;

    let gridLines = "";
    for (let i = 0; i <= 3; i++) {
      const yVal = (maxVal / 3) * i;
      const yPos = getY(yVal);
      gridLines += `
        <line x1="${pad.left}" y1="${yPos}" x2="${width - pad.right}" y2="${yPos}" class="chart-grid-line" />
        <text x="${pad.left - 8}" y="${yPos + 4}" text-anchor="end" class="chart-axis-text">
          ${(yVal / 1000000).toFixed(0)}M
        </text>
      `;
    }

    let labels = "";
    let points = "";
    timeline.forEach((d, i) => {
      const x = getX(i);
      labels += `<text x="${x}" y="${height - 10}" text-anchor="middle" class="chart-axis-text">${d.data}</text>`;
      points += `
        <circle cx="${x}" cy="${getY(d.emp)}" stroke="#38bdf8" class="chart-point">
          <title>${d.data}: Empenhado ${formatBRL(d.emp)}</title>
        </circle>
      `;
    });

    container.innerHTML = `
      <svg viewBox="0 0 ${width} ${height}" class="svg-chart">
        <defs>
          <linearGradient id="gradientEmpArea" x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stop-color="#38bdf8" stop-opacity="0.4" />
            <stop offset="100%" stop-color="#38bdf8" stop-opacity="0.0" />
          </linearGradient>
        </defs>
        ${gridLines}
        <path d="${areaPath}" class="chart-area-emp" />
        <path d="${provPath}" class="chart-path-prov" />
        <path d="${empPath}" class="chart-path-emp" />
        <path d="${liqPath}" class="chart-path-liq" />
        ${points}
        ${labels}
      </svg>
    `;
  }

  // ========================================================================
  // 8. DEMAIS ABAS (RASTREABILIDADE, EMPENHOS, FORNECEDORES)
  // ========================================================================
  function renderNCTable() {
    if (!tableNCBody) return;

    let list = DASHBOARD_DATA.creditosEmTela || [];
    if (currentUnitId !== "total") {
      list = list.filter((nc) => nc.om === currentUnitId);
    }
    if (currentSubFilter === "ordinario") {
      list = list.filter((nc) => nc.u === "160");
    } else if (currentSubFilter === "saude") {
      list = list.filter((nc) => nc.u === "167");
    }

    if (searchGlobalQuery.trim() !== "") {
      const q = searchGlobalQuery.toLowerCase();
      list = list.filter(
        (nc) =>
          nc.nc.toLowerCase().includes(q) ||
          nc.obj.toLowerCase().includes(q) ||
          nc.nd.includes(q) ||
          nc.pi.toLowerCase().includes(q)
      );
    }

    // Ordenação Dinâmica Clicável
    list = sortList(list, sortState.nc.key, sortState.nc.dir, sortState.nc.type);

    tableNCBody.innerHTML = list.slice(0, 50)
      .map((nc) => `
        <tr data-nc="${nc.nc}">
          <td><span class="nc-code-pill">${nc.nc}</span></td>
          <td><strong>${nc.om}</strong></td>
          <td><span style="font-family: var(--font-mono); font-size: 0.75rem;">${nc.nd}</span></td>
          <td><div class="finalidade-truncate" title="${nc.obj}">${nc.obj}</div></td>
          <td style="font-family: var(--font-mono);">${formatBRL(nc.prov || nc.v * 1.5)}</td>
          <td style="font-family: var(--font-mono); color: var(--text-emerald);">${formatBRL(nc.saldo || nc.v)}</td>
          <td><span class="badge-pill badge-cyan">${nc.taxa || 92.5}%</span></td>
        </tr>
      `)
      .join("");

    tableNCBody.querySelectorAll("tr[data-nc]").forEach((row) => {
      row.addEventListener("click", () => openDrawer(row.getAttribute("data-nc")));
    });
  }

  function renderNETable() {
    if (!tableNEBody) return;

    let list = DASHBOARD_DATA.todasNEs || [];
    if (currentUnitId !== "total") {
      const unit = DASHBOARD_DATA.unidades.find((u) => u.id === currentUnitId);
      if (unit) {
        list = list.filter((ne) => ne.omSigla === unit.sigla);
      }
    }

    if (searchNEQuery && searchNEQuery.trim() !== "") {
      const q = searchNEQuery.toLowerCase().trim();
      list = list.filter(
        (ne) =>
          (ne.ne && ne.ne.toLowerCase().includes(q)) ||
          (ne.credor && ne.credor.toLowerCase().includes(q)) ||
          (ne.cnpj && ne.cnpj.toLowerCase().includes(q)) ||
          (ne.processo && ne.processo.toLowerCase().includes(q)) ||
          (ne.omSigla && ne.omSigla.toLowerCase().includes(q))
      );
    }

    if (neCountBadge) {
      neCountBadge.textContent = `${list.length.toLocaleString("pt-BR")} NEs`;
    }

    // Ordenação Dinâmica Clicável
    list = sortList(list, sortState.ne.key, sortState.ne.dir, sortState.ne.type);

    const maxItems = 250;
    const displayList = list.slice(0, maxItems);

    tableNEBody.innerHTML = displayList
      .map(
        (ne) => `
        <tr>
          <td><span class="nc-code-pill" style="border-color: rgba(16,185,129,0.3); color: var(--text-emerald);">${ne.ne}</span></td>
          <td><strong>${ne.omSigla}</strong></td>
          <td>
            <div style="font-weight: 600; color: var(--text-primary);">${ne.credor}</div>
            <div style="font-size: 0.72rem; color: var(--text-muted);">CNPJ: ${ne.cnpj}</div>
          </td>
          <td><span style="font-size: 0.75rem; color: var(--text-accent); font-family: var(--font-mono);">${ne.processo}</span></td>
          <td style="font-family: var(--font-mono); font-weight: 700;">${formatBRL(ne.valor)}</td>
          <td style="font-size: 0.75rem; color: var(--text-muted);">${ne.data}</td>
          <td><span class="badge-pill badge-cyan">${ne.status}</span></td>
        </tr>
      `
      )
      .join("");

    if (list.length > maxItems) {
      tableNEBody.innerHTML += `
        <tr>
          <td colspan="7" style="text-align: center; color: var(--text-muted); font-size: 0.8rem; padding: 1rem;">
            Mostrando os primeiros ${maxItems} de ${list.length.toLocaleString("pt-BR")} empenhos. Utilize o campo de busca acima para localizar qualquer empenho específico.
          </td>
        </tr>
      `;
    }
  }

  function renderFornecedoresEProcessos() {
    if (fornecedoresList) {
      fornecedoresList.innerHTML = DASHBOARD_DATA.topFornecedores
        .map(
          (f, idx) => `
          <div class="empenho-mini-card" style="display: flex; justify-content: space-between; align-items: center;">
            <div style="display: flex; gap: 0.85rem; align-items: center;">
              <span style="font-size: 1.1rem; font-weight: 800; color: var(--text-muted); font-family: var(--font-mono); width: 24px;">#${idx + 1}</span>
              <div>
                <div style="font-weight: 700; color: var(--text-primary); font-size: 0.85rem;">${f.nome}</div>
                <div style="font-size: 0.72rem; color: var(--text-muted);">CNPJ: ${f.cnpj} • ${f.qtdNEs} NEs • OM: ${f.omds}</div>
              </div>
            </div>
            <div style="font-family: var(--font-mono); font-weight: 800; color: var(--text-emerald); font-size: 0.95rem;">
              ${formatBRL(f.val)}
            </div>
          </div>
        `
        )
        .join("");
    }

    if (processosList) {
      processosList.innerHTML = DASHBOARD_DATA.topProcessos
        .map(
          (p, idx) => `
          <div class="empenho-mini-card" style="display: flex; justify-content: space-between; align-items: center;">
            <div>
              <span class="nc-code-pill" style="margin-bottom: 0.25rem; display: inline-block;">${p.proc}</span>
              <div style="font-size: 0.8rem; color: var(--text-secondary);">${p.objeto}</div>
              <div style="font-size: 0.72rem; color: var(--text-muted);">Solicitante: ${p.omds}</div>
            </div>
            <div style="font-family: var(--font-mono); font-weight: 800; color: var(--text-accent); font-size: 0.95rem;">
              ${formatBRL(p.val)}
            </div>
          </div>
        `
        )
        .join("");
    }
  }

  // ========================================================================
  // 9. DRAWER LATERAL DE DETALHAMENTO DE NOTA DE CRÉDITO (RAIO-X COMPLETO)
  // ========================================================================
  function openDrawer(ncCode, omHint) {
    if (!drawerBackdrop) return;

    // 1. Localizar dados no mapa global de NCs ou em creditosEmTela
    let ncInfo = (DASHBOARD_DATA.mapaNCs && (DASHBOARD_DATA.mapaNCs[ncCode] || DASHBOARD_DATA.mapaNCs[ncCode.trim()])) || null;
    const ncTela = DASHBOARD_DATA.creditosEmTela.find((item) => item.nc === ncCode || item.nc.endsWith(ncCode) || ncCode.endsWith(item.nc));

    if (!ncInfo && !ncTela) {
      console.warn("NC não encontrada:", ncCode);
      return;
    }

    const fullNC = (ncInfo && ncInfo.nc) || (ncTela && ncTela.nc) || ncCode;
    const omSigla = (ncInfo && ncInfo.om) || (ncTela && ncTela.om) || omHint || "OM";
    const concedenteNome = (ncInfo && ncInfo.concedente_nome) || "Órgão Concedente (SIAFI / COLOG)";
    const opNome = (ncInfo && ncInfo.op_principal) || "DESCENTRALIZAÇÃO DE CRÉDITO";
    const ndStr = (ncInfo && ncInfo.nd ? `${ncInfo.nd} - ${ncInfo.ndn}` : (ncTela ? `${ncTela.nd} - ${ncTela.ndnome}` : "—"));
    const piStr = (ncInfo && ncInfo.pi ? `${ncInfo.pi} (Ação: ${ncInfo.acao || "—"})` : (ncTela ? `${ncTela.pi}` : "—"));
    const prazoStr = (ncTela && ncTela.prazo) || (ncInfo && ncInfo.prazo) || "30/11/2026";
    const finalidadeStr = (ncTela && ncTela.obj) || (ncInfo && ncInfo.obj) || "Finalidade registrada no SIAFI";

    // Valores
    const saldoDisp = (ncTela ? ncTela.v : (ncInfo ? ncInfo.saldo_tela : 0.0));
    const valAlocado = (ncInfo && ncInfo.val_alocado > 0 ? ncInfo.val_alocado : (ncTela ? ncTela.prov : saldoDisp));
    const nesList = (ncInfo && ncInfo.nes && ncInfo.nes.length > 0) ? ncInfo.nes : [];
    const valEmpenhado = nesList.reduce((acc, n) => acc + (n.val || 0), 0) || Math.max(0, valAlocado - saldoDisp);
    const taxaExec = valAlocado > 0 ? (valEmpenhado / valAlocado * 100) : (saldoDisp === 0 ? 100 : 0);

    // Preenchimento do Header do Drawer
    const titleEl = document.getElementById("drawer-nc-title");
    if (titleEl) {
      titleEl.innerHTML = `
        <div style="display: flex; flex-direction: column; gap: 0.35rem;">
          <span style="font-size: 1.15rem; font-family: var(--font-mono); font-weight: 800; color: var(--text-primary); word-break: break-all;">${fullNC}</span>
        </div>
      `;
    }

    const concedenteEl = document.getElementById("drawer-concedente-name");
    if (concedenteEl) concedenteEl.textContent = concedenteNome;

    const opBadge = document.getElementById("drawer-op-badge");
    if (opBadge) {
      opBadge.textContent = opNome;
      opBadge.className = `badge-pill ${opNome.includes("ANUL") ? "badge-rose" : (opNome.includes("DETALH") ? "badge-amber" : "badge-cyan")}`;
    }

    const omdsEl = document.getElementById("drawer-omds-name");
    if (omdsEl) omdsEl.textContent = `${omSigla} (UG ${(ncInfo && ncInfo.ug) || (ncTela && ncTela.u) || "160"})`;

    const ndEl = document.getElementById("drawer-nd-val");
    if (ndEl) ndEl.textContent = ndStr;

    const piEl = document.getElementById("drawer-pi-val");
    if (piEl) piEl.textContent = piStr;

    const prazoEl = document.getElementById("drawer-prazo-val");
    if (prazoEl) prazoEl.textContent = prazoStr;

    const finalidadeEl = document.getElementById("drawer-finalidade-text");
    if (finalidadeEl) finalidadeEl.textContent = finalidadeStr;

    // Valores
    const provEl = document.getElementById("drawer-prov-val");
    if (provEl) provEl.textContent = formatBRL(valAlocado);

    const empEl = document.getElementById("drawer-emp-val");
    if (empEl) empEl.textContent = formatBRL(valEmpenhado);

    const saldoEl = document.getElementById("drawer-saldo-val");
    if (saldoEl) saldoEl.textContent = formatBRL(saldoDisp);

    const execEl = document.getElementById("drawer-execucao-val");
    if (execEl) execEl.textContent = formatPercent(taxaExec);

    // Movimentações Registradas da NC
    const movContainer = document.getElementById("drawer-movimentacoes-list");
    if (movContainer) {
      const movs = (ncInfo && ncInfo.movimentacoes) ? ncInfo.movimentacoes : [];
      if (movs.length > 0) {
        movContainer.innerHTML = movs.map((m) => {
          const isAnul = m.op.includes("ANUL") || m.val < 0;
          const isDet = m.op.includes("DETALH");
          const cardClass = isAnul ? "drawer-mov-card is-anulacao" : (isDet ? "drawer-mov-card is-detalhamento" : "drawer-mov-card");
          return `
            <div class="${cardClass}">
              <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.25rem;">
                <strong style="color: var(--text-primary); font-size: 0.8rem;">${m.op || "Movimentação"}</strong>
                <span style="font-family: var(--font-mono); font-weight: 700; color: ${m.val >= 0 ? "var(--text-emerald)" : "var(--text-rose)"};">
                  ${m.val >= 0 ? "+" : ""}${formatBRL(m.val)}
                </span>
              </div>
              <div style="font-size: 0.72rem; color: var(--text-muted); margin-bottom: 0.25rem;">
                Data: ${m.dia || "Exercício 2026"} • ND: ${m.nd} (${m.ndn || "Geral"})
              </div>
              <div style="font-size: 0.74rem; color: var(--text-secondary); line-height: 1.3;">
                ${m.obj || "Registro efetuado no SIAFI"}
              </div>
            </div>
          `;
        }).join("");
      } else {
        movContainer.innerHTML = `
          <div style="font-size: 0.75rem; color: var(--text-muted); padding: 0.5rem; background: rgba(255,255,255,0.02); border-radius: 8px;">
            Aporte inicial descentralizado diretamente para a unidade no exercício de 2026.
          </div>
        `;
      }
    }

    // Empenhos vinculados
    const empenhosContainer = document.getElementById("drawer-empenhos-list");
    if (empenhosContainer) {
      if (nesList.length > 0) {
        empenhosContainer.innerHTML = nesList.map((ne) => `
          <div class="empenho-mini-card">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.25rem;">
              <span class="nc-code-pill" style="border-color: rgba(16,185,129,0.3); color: var(--text-emerald); font-weight: 700;">${ne.ne}</span>
              <span style="font-family: var(--font-mono); font-weight: 700; color: var(--text-primary); font-size: 0.85rem;">${formatBRL(ne.val)}</span>
            </div>
            <div style="font-size: 0.8rem; font-weight: 600; color: var(--text-primary);">${ne.fav || "Credor"}</div>
            <div style="font-size: 0.72rem; color: var(--text-muted); display: flex; justify-content: space-between; margin-top: 0.25rem;">
              <span>CNPJ: ${ne.doc || "—"}</span>
              <span style="color: var(--text-accent); font-family: var(--font-mono);">${ne.proc || "Processo"}</span>
            </div>
            <div style="font-size: 0.7rem; color: var(--text-muted); margin-top: 0.25rem; display: flex; justify-content: space-between;">
              <span>Data: ${ne.dia || "—"}</span>
              <span class="badge-pill badge-cyan">${ne.prova || "Auditado"}</span>
            </div>
          </div>
        `).join("");
      } else {
        empenhosContainer.innerHTML = `
          <div style="font-size: 0.75rem; color: var(--text-muted); padding: 0.75rem; background: rgba(255,255,255,0.02); border-radius: 8px; border: 1px dashed var(--border-subtle); text-align: center;">
            Nenhum empenho emitido até o momento. Saldo integral disponível para execução pela unidade gestora.
          </div>
        `;
      }
    }

    drawerBackdrop.classList.add("is-open");
    document.body.style.overflow = "hidden";
  }

  function closeDrawer() {
    if (!drawerBackdrop) return;
    drawerBackdrop.classList.remove("is-open");
    document.body.style.overflow = "";
  }

  function setupSortableHeaders() {
    document.querySelectorAll("th.sortable-th").forEach((th) => {
      th.addEventListener("click", () => {
        const table = th.getAttribute("data-table"); // "tela", "nc", "ne"
        const sortKey = th.getAttribute("data-sort");
        const type = th.getAttribute("data-type") || "string";

        if (!table || !sortKey || !sortState[table]) return;

        if (sortState[table].key === sortKey) {
          // Inverter direção
          sortState[table].dir = sortState[table].dir === "asc" ? "desc" : "asc";
        } else {
          sortState[table].key = sortKey;
          sortState[table].type = type;
          sortState[table].dir = (type === "number" || type === "date") ? "desc" : "asc";
        }

        // Atualizar ícones visuais
        const tableEl = th.closest("table");
        if (tableEl) {
          tableEl.querySelectorAll("th.sortable-th").forEach((otherTh) => {
            const isCurrent = otherTh === th;
            otherTh.classList.toggle("is-active", isCurrent);
            const iconEl = otherTh.querySelector(".sort-icon");
            if (iconEl) {
              if (isCurrent) {
                iconEl.textContent = sortState[table].dir === "asc" ? "▲" : "▼";
              } else {
                iconEl.textContent = "↕";
              }
            }
          });
        }

        // Re-renderizar a tabela correspondente
        if (table === "tela") renderCreditosEmTela();
        else if (table === "nc") renderNCTable();
        else if (table === "ne") renderNETable();
      });
    });
  }

  if (drawerCloseBtn) drawerCloseBtn.addEventListener("click", closeDrawer);
  if (drawerBackdrop) {
    drawerBackdrop.addEventListener("click", (e) => {
      if (e.target === drawerBackdrop) closeDrawer();
    });
  }
  window.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeDrawer();
  });

  // ========================================================================
  // 10. CONTROLADORES DE EVENTOS
  // ========================================================================
  // Navegação por Abas Unificada
  window.switchTab = function (tabId) {
    const allTabBtns = document.querySelectorAll(".tab-nav-btn");
    const allTabPanes = document.querySelectorAll(".tab-pane");

    allTabBtns.forEach((b) => {
      if (b.getAttribute("data-tab") === tabId) {
        b.classList.add("is-active");
        b.scrollIntoView({ behavior: "smooth", inline: "center", block: "nearest" });
      } else {
        b.classList.remove("is-active");
      }
    });

    allTabPanes.forEach((p) => {
      if (p.id === tabId) {
        p.classList.add("is-active");
        p.classList.remove("unit-content-enter");
        void p.offsetWidth;
        p.classList.add("unit-content-enter");
      } else {
        p.classList.remove("is-active");
      }
    });

    activeTab = tabId;
  };

  tabNavBtns.forEach((btn) => {
    btn.addEventListener("click", () => {
      const targetId = btn.getAttribute("data-tab");
      window.switchTab(targetId);
    });
  });

  // Sub-filtros da OM (Ordinário vs Saúde)
  subFilterBtns.forEach((btn) => {
    btn.addEventListener("click", () => {
      subFilterBtns.forEach((b) => b.classList.remove("is-active"));
      btn.classList.add("is-active");
      currentSubFilter = btn.getAttribute("data-filter");
      updateKPIs();
      renderCreditosEmTela();
      renderNCTable();
    });
  });

  // Filtros de UG na tela principal (Todas vs 160 vs 167)
  ugFilterBtns.forEach((btn) => {
    btn.addEventListener("click", () => {
      ugFilterBtns.forEach((b) => b.classList.remove("is-active"));
      btn.classList.add("is-active");
      currentUGFilter = btn.getAttribute("data-ug-filter");
      renderCreditosEmTela();
    });
  });

  // Busca na Tela Principal
  if (searchTelaInput) {
    searchTelaInput.addEventListener("input", (e) => {
      searchTelaQuery = e.target.value;
      renderCreditosEmTela();
    });
  }

  // Busca Global
  if (searchGlobalInput) {
    searchGlobalInput.addEventListener("input", (e) => {
      searchGlobalQuery = e.target.value;
      renderNCTable();
      renderNETable();
    });
  }

  // Busca na Tabela de Empenhos
  if (searchNEInput) {
    searchNEInput.addEventListener("input", (e) => {
      searchNEQuery = e.target.value;
      renderNETable();
    });
  }

  // Alternador de Shader
  if (btnToggleShader) {
    btnToggleShader.addEventListener("click", () => {
      const active = motionBg.toggle();
      btnToggleShader.innerHTML = active ? `<span>✨ Shader 60fps</span>` : `<span>⚡ Modo Econômico</span>`;
    });
  }

  // ========================================================================
  // CONTROLE DO PLACAR DE RANQUEAMENTO (CRITÉRIOS E COLAPSO)
  // ========================================================================
  if (rankCritBtns) {
    rankCritBtns.forEach((btn) => {
      btn.addEventListener("click", () => {
        rankCritBtns.forEach((b) => b.classList.remove("is-active"));
        btn.classList.add("is-active");
        currentRankCriteria = btn.getAttribute("data-crit");
        renderRankingCard();
      });
    });
  }

  if (btnToggleRanking && rankingSection) {
    btnToggleRanking.addEventListener("click", () => {
      const isCollapsed = rankingSection.classList.toggle("is-collapsed");
      if (rankingToggleIcon) rankingToggleIcon.textContent = isCollapsed ? "▼" : "▲";
      if (rankingToggleText) rankingToggleText.textContent = isCollapsed ? "Expandir Placar" : "Ocultar Placar";
    });
  }

  // ========================================================================
  // NAVEGAÇÃO DE ABAS: ARRASTE E ROLAGEM HORIZONTAL SUAVE
  // ========================================================================
  const tabsNav = document.querySelector(".sicof-tabs-nav");
  if (tabsNav) {
    let isDown = false;
    let startX;
    let scrollLeft;

    tabsNav.addEventListener("mousedown", (e) => {
      isDown = true;
      tabsNav.classList.add("is-dragging");
      startX = e.pageX - tabsNav.offsetLeft;
      scrollLeft = tabsNav.scrollLeft;
    });

    tabsNav.addEventListener("mouseleave", () => {
      isDown = false;
      tabsNav.classList.remove("is-dragging");
    });

    tabsNav.addEventListener("mouseup", () => {
      isDown = false;
      tabsNav.classList.remove("is-dragging");
    });

    tabsNav.addEventListener("mousemove", (e) => {
      if (!isDown) return;
      e.preventDefault();
      const x = e.pageX - tabsNav.offsetLeft;
      const walk = (x - startX) * 1.5;
      tabsNav.scrollLeft = scrollLeft - walk;
    });

    tabsNav.addEventListener(
      "wheel",
      (e) => {
        if (e.deltaY !== 0) {
          e.preventDefault();
          tabsNav.scrollLeft += e.deltaY;
        }
      },
      { passive: false }
    );
  }

  // ========================================================================
  // EXPORTAÇÃO EXECUTIVA MICROSOFT EXCEL (.XLSX) — EXCLUSIVO CRÉDITOS EM TELA
  // ========================================================================
  async function exportCreditosEmTelaExcel(preferSortKey = null) {
    let items = DASHBOARD_DATA.creditosEmTela;
    if (currentUnitId !== "total") {
      items = items.filter((c) => c.om === currentUnitId);
    }
    if (currentUGFilter !== "all") {
      items = items.filter((c) => c.u === currentUGFilter);
    }
    if (searchTelaQuery.trim() !== "") {
      const q = searchTelaQuery.toLowerCase();
      items = items.filter(
        (c) =>
          c.nc.toLowerCase().includes(q) ||
          c.nd.includes(q) ||
          c.ndnome.toLowerCase().includes(q) ||
          c.pi.toLowerCase().includes(q) ||
          c.obj.toLowerCase().includes(q) ||
          c.uasg.includes(q)
      );
    }

    const sortKey = preferSortKey || sortState.tela.key || "v";
    const sortDir = sortState.tela.dir || "desc";
    const sortType = sortState.tela.type || (sortKey === "v" || sortKey === "dias" ? "number" : "string");
    items = sortList(items, sortKey, sortDir, sortType);

    const unit = DASHBOARD_DATA.unidades.find((u) => u.id === currentUnitId) || DASHBOARD_DATA.unidades[0];
    const totalVal = items.reduce((acc, c) => acc + c.v, 0);

    // Se ExcelJS estiver disponível, gera XLSX estilizado com paleta do site
    if (typeof ExcelJS !== "undefined") {
      try {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "SICOF 3.0 — Base de Apoio Logístico do Exército";
        workbook.lastModifiedBy = "SIAFI / Tesouro Gerencial";
        workbook.created = new Date();
        workbook.modified = new Date();

        const sheet = workbook.addWorksheet("Créditos em Tela", {
          views: [{ state: "frozen", ySplit: 3, xSplit: 0 }]
        });

        // 1. Título Executivo (Navy #0F172A)
        const titleRow = sheet.addRow(["SISTEMA DE CONTROLE ORÇAMENTÁRIO E FINANCEIRO — CRÉDITOS EM TELA"]);
        sheet.mergeCells("A1:K1");
        titleRow.height = 36;
        titleRow.getCell(1).font = { name: "Segoe UI", size: 14, bold: true, color: { argb: "FFFFFFFF" } };
        titleRow.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
        titleRow.getCell(1).alignment = { vertical: "middle", horizontal: "left", indent: 1 };

        // 2. Metadados Executivos (#1E293B com Ciano #38BDF8)
        const ugText = currentUGFilter === "all" ? "Todas as UGs (160/167)" : currentUGFilter === "160" ? "UG 160 (Ordinário)" : "UG 167 (Saúde)";
        const sortLabelDesc = sortKey === "dias" ? "Por Dias em Tela (Idade)" : sortKey === "v" ? "Por Saldo em Tela (Valor)" : "Padrão";
        const metaText = `Unidade: ${unit.nome.toUpperCase()} (${unit.sigla})  •  Gestão: ${ugText}  •  Notas em Tela: ${items.length}  •  Totalizador: R$ ${totalVal.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}  •  Posição: ${DASHBOARD_DATA.meta.posicaoData}  •  Ordenação: ${sortLabelDesc}`;

        const metaRow = sheet.addRow([metaText]);
        sheet.mergeCells("A2:K2");
        metaRow.height = 24;
        metaRow.getCell(1).font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF38BDF8" } };
        metaRow.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E293B" } };
        metaRow.getCell(1).alignment = { vertical: "middle", horizontal: "left", indent: 1 };

        // 3. Cabeçalhos das Colunas
        const headers = [
          "Nota de Crédito (NC)",
          "Unidade (OM)",
          "Gestão / UG",
          "UASG",
          "Prazo Limite",
          "Dias em Tela (Idade)",
          "Natureza Despesa (ND)",
          "Denominação da Despesa",
          "Plano Interno (PI)",
          "Finalidade / Objeto do Recurso",
          "Saldo Disponível (R$)"
        ];
        const headerRow = sheet.addRow(headers);
        headerRow.height = 28;

        headerRow.eachCell((cell, colNumber) => {
          cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
          cell.alignment = {
            vertical: "middle",
            horizontal: colNumber === 11 ? "right" : (colNumber <= 6 || colNumber === 7 || colNumber === 9) ? "center" : "left"
          };
          cell.border = {
            top: { style: "thin", color: { argb: "FF38BDF8" } },
            bottom: { style: "medium", color: { argb: "FF38BDF8" } },
            left: { style: "thin", color: { argb: "FF334155" } },
            right: { style: "thin", color: { argb: "FF334155" } }
          };
        });

        // 4. Linhas de Dados com Formatação e Cores da Identidade Visual
        items.forEach((c, idx) => {
          const isEven = idx % 2 === 0;
          const zebraBg = isEven ? "FFFFFFFF" : "FFF8FAFC";
          const ugDesc = c.u === "167" ? "UG 167 (Saúde)" : "UG 160 (Ordinário)";
          const omSigla = (DASHBOARD_DATA.unidades.find((u) => u.id === c.om)?.sigla) || c.om.toUpperCase();

          const row = sheet.addRow([
            c.nc,
            omSigla,
            ugDesc,
            c.uasg,
            c.prazo,
            c.dias,
            c.nd,
            c.ndnome || "—",
            c.pi,
            c.obj || "—",
            c.v
          ]);
          row.height = 22;

          row.eachCell((cell, colNumber) => {
            cell.font = { name: "Segoe UI", size: 9.5, color: { argb: "FF0F172A" } };
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: zebraBg } };
            cell.border = {
              bottom: { style: "thin", color: { argb: "FFE2E8F0" } },
              left: { style: "thin", color: { argb: "FFF1F5F9" } },
              right: { style: "thin", color: { argb: "FFF1F5F9" } }
            };

            // NC
            if (colNumber === 1) {
              cell.alignment = { horizontal: "center", vertical: "middle" };
              cell.font = { name: "Consolas", size: 9.5, bold: true, color: { argb: "FF0284C7" } };
            }
            // OM
            else if (colNumber === 2) {
              cell.alignment = { horizontal: "center", vertical: "middle" };
              cell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF0F172A" } };
            }
            // UG
            else if (colNumber === 3) {
              cell.alignment = { horizontal: "center", vertical: "middle" };
              cell.font = { name: "Segoe UI", size: 9, color: c.u === "167" ? { argb: "FF059669" } : { argb: "FF0284C7" } };
            }
            // UASG
            else if (colNumber === 4) {
              cell.alignment = { horizontal: "center", vertical: "middle" };
              cell.font = { name: "Consolas", size: 9.5 };
            }
            // Prazo
            else if (colNumber === 5) {
              cell.alignment = { horizontal: "center", vertical: "middle" };
              cell.font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FFD97706" } };
            }
            // Dias em Tela (Idade) - Destaque por Gravidade
            else if (colNumber === 6) {
              cell.alignment = { horizontal: "center", vertical: "middle" };
              cell.numFmt = '#,##0" dias"';
              if (c.dias > 60) {
                cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFE4E6" } };
                cell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF9F1239" } };
              } else if (c.dias > 30) {
                cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEF3C7" } };
                cell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF92400E" } };
              } else {
                cell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF047857" } };
              }
            }
            // ND
            else if (colNumber === 7) {
              cell.alignment = { horizontal: "center", vertical: "middle" };
              cell.font = { name: "Consolas", size: 9.5, bold: true };
            }
            // Descrição
            else if (colNumber === 8) {
              cell.alignment = { horizontal: "left", vertical: "middle" };
            }
            // PI
            else if (colNumber === 9) {
              cell.alignment = { horizontal: "center", vertical: "middle" };
              cell.font = { name: "Consolas", size: 9, color: { argb: "FF0284C7" } };
            }
            // Objeto
            else if (colNumber === 10) {
              cell.alignment = { horizontal: "left", vertical: "middle" };
            }
            // Saldo Disponível (R$)
            else if (colNumber === 11) {
              cell.alignment = { horizontal: "right", vertical: "middle" };
              cell.numFmt = '"R$ "#,##0.00;[Red]-"R$ "#,##0.00';
              cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FF047857" } };
            }
          });
        });

        // 5. Linha de Rodapé com Totais e Médias (Fórmulas do Excel)
        const lastDataRow = items.length + 3;
        const totalRow = sheet.addRow([
          `TOTAL GERAL (${items.length} NOTAS)`,
          "",
          "",
          "",
          "MÉDIA DIAS:",
          { formula: `AVERAGE(F4:F${lastDataRow})` },
          "",
          "",
          "",
          "SALDO TOTAL:",
          { formula: `SUM(K4:K${lastDataRow})` }
        ]);
        totalRow.height = 30;

        totalRow.eachCell((cell, colNumber) => {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
          cell.border = {
            top: { style: "medium", color: { argb: "FF38BDF8" } },
            bottom: { style: "double", color: { argb: "FF38BDF8" } }
          };

          if (colNumber === 1) {
            cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FF38BDF8" } };
            cell.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
          } else if (colNumber === 5 || colNumber === 10) {
            cell.font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FF94A3B8" } };
            cell.alignment = { vertical: "middle", horizontal: "right" };
          } else if (colNumber === 6) {
            cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFF59E0B" } };
            cell.alignment = { vertical: "middle", horizontal: "center" };
            cell.numFmt = '#,##0" dias"';
          } else if (colNumber === 11) {
            cell.font = { name: "Segoe UI", size: 11, bold: true, color: { argb: "FF34D399" } };
            cell.alignment = { vertical: "middle", horizontal: "right" };
            cell.numFmt = '"R$ "#,##0.00';
          }
        });

        // 6. Ativar AutoFilter nativo no Excel para que o usuário possa ordenar por qualquer coluna
        sheet.autoFilter = {
          from: { row: 3, column: 1 },
          to: { row: lastDataRow, column: 11 }
        };

        // 7. Larguras das Colunas Calibradas
        sheet.columns = [
          { width: 20 }, // NC
          { width: 14 }, // OM
          { width: 18 }, // UG
          { width: 12 }, // UASG
          { width: 14 }, // Prazo
          { width: 16 }, // Dias em Tela (Idade)
          { width: 16 }, // ND
          { width: 34 }, // Denominação
          { width: 16 }, // PI
          { width: 48 }, // Finalidade
          { width: 22 }  // Saldo
        ];

        // 8. Download como Arquivo Binário .xlsx
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        const sortLabel = sortKey === "dias" ? "por_Idade" : "por_Valor";
        a.download = `SICOF_Creditos_em_Tela_${unit.sigla}_${sortLabel}_${new Date().toISOString().slice(0, 10)}.xlsx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        return;
      } catch (err) {
        console.error("Erro na exportação ExcelJS, acionando fallback XLSX:", err);
      }
    }

    // Fallback SheetJS (XLSX)
    if (typeof XLSX !== "undefined") {
      const wb = XLSX.utils.book_new();
      const headers = [
        "Nota de Crédito",
        "Unidade (OM)",
        "UG Emitente",
        "UASG",
        "Prazo / Status",
        "Dias em Tela (Idade)",
        "Natureza Despesa (ND)",
        "Descrição da Despesa",
        "Plano Interno (PI)",
        "Finalidade / Objeto do Recurso",
        "Saldo Disponível (R$)"
      ];
      const rows = items.map((c) => [
        c.nc,
        c.om.toUpperCase(),
        `UG ${c.u}`,
        c.uasg,
        c.prazo,
        c.dias,
        c.nd,
        c.ndnome || "—",
        c.pi,
        c.obj || "—",
        c.v
      ]);

      const ws = XLSX.utils.aoa_to_sheet([
        [`SICOF 3.0 — CRÉDITOS EM TELA (${unit.sigla})`],
        [`Notas: ${items.length} • Total: R$ ${totalVal.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`],
        [],
        headers,
        ...rows
      ]);
      XLSX.utils.book_append_sheet(wb, ws, "Créditos em Tela");
      XLSX.writeFile(wb, `SICOF_Creditos_em_Tela_${unit.sigla}.xlsx`);
    } else {
      alert("Biblioteca de exportação não carregada.");
    }
  }

  // Event Listeners dos Botões de Exportação
  if (btnExport) {
    btnExport.addEventListener("click", () => exportCreditosEmTelaExcel());
  }

  const btnExportTelaTab = document.getElementById("btn-export-tela-tab");
  if (btnExportTelaTab) {
    btnExportTelaTab.addEventListener("click", () => exportCreditosEmTelaExcel());
  }

  // ========================================================================
  // 11. INICIALIZAÇÃO INICIAL & CONTROLE DE TEMA (ALTO CONTRASTE)
  // ========================================================================
  // Carrega preferência do usuário ou padrão escuro
  let initialTheme = "dark";
  try {
    initialTheme = localStorage.getItem("sicof_theme") || "dark";
  } catch (e) {}
  applyTheme(initialTheme);

  if (btnToggleTheme) {
    btnToggleTheme.addEventListener("click", () => {
      const isLight = document.documentElement.getAttribute("data-theme") === "light";
      applyTheme(isLight ? "dark" : "light");
    });
  }

  setupSortableHeaders();
  renderLogosShowcase();
  renderRankingCard();
  selectUnit("total");
});
