"""
=============================================================================
PAINEL EXECUTIVO DE CRÉDITO DISPONÍVEL & EMPENHOS 2.0 (DASHBOARD V2 PRO)
Base de Apoio Logístico do Exército · SIAFI / Tesouro Gerencial
Autor: Principal Data Visualization Engineer & SIAFI Auditor
Desenvolvido por: 3º Sgt De Campos · Corte Oficial 2026
Arquitetura: WCAG 2.2 AAA + 60fps GPU + IBCS & Tufte + Mobile-First (360px-1860px)
=============================================================================
"""

import json
import os
import re
import sys
import pathlib

sys.stdout.reconfigure(encoding='utf-8')

# Caminhos canônicos do workspace
ROOT_DIR = pathlib.Path(__file__).parent
V1_HTML_PATH = ROOT_DIR / "site" / "index.html"
DATA_JSON_PATH = ROOT_DIR / "data" / "data.json"
HIST_JSON_PATH = ROOT_DIR / "data" / "history.json"
OUT_HTML_PATH = ROOT_DIR / "site" / "index.html"

def build():
    print("==================================================================")
    print("  BUILD DASHBOARD V2 PRO: UI/UX, 7 ESTADOS & MOBILE-FIRST")
    print("==================================================================")

    assert DATA_JSON_PATH.exists(), f"Erro: {DATA_JSON_PATH} não encontrado!"
    assert V1_HTML_PATH.exists(), f"Erro: {V1_HTML_PATH} não encontrado!"

    print("1. Carregando bancos de dados consolidados...")
    raw_data = json.loads(DATA_JSON_PATH.read_text(encoding="utf-8"))
    hist_data = json.loads(HIST_JSON_PATH.read_text(encoding="utf-8")) if HIST_JSON_PATH.exists() else []

    print("2. Lendo matriz base e templates...")
    v1_html = V1_HTML_PATH.read_text(encoding="utf-8")

    # Extração precisa do DOM e scripts
    body_match = re.search(r'<body[^>]*>(.*?)<script>\s*var\s+CELDATA\s*=', v1_html, re.DOTALL)
    if not body_match:
        body_match = re.search(r'<body[^>]*>(.*?)<script', v1_html, re.DOTALL)
    if not body_match:
        raise ValueError("Falha ao extrair body da matriz.")
    body_inner = body_match.group(1)

    script_m = re.search(r'<script>\s*var\s+CELDATA\s*=', v1_html)
    script_start = script_m.start() if script_m else v1_html.find("<script")
    script_text = v1_html[script_start:]
    syn_pos = script_text.find("var BCMS_SYNONYMS =")
    if syn_pos == -1:
        syn_pos = script_text.find("BCMS_SYNONYMS =")
    js_logic = script_text[syn_pos:] if syn_pos != -1 else ""

    # Extração de CSS base
    style_match = re.search(r'<style>(.*?)</style>', v1_html, re.DOTALL)
    v1_css = style_match.group(1) if style_match else ""

    # Carregamento dos Design Tokens oficiais do Figma (figmacraft / One-X Design System)
    tokens_path = ROOT_DIR / "design-tokens.css"
    if tokens_path.exists():
        (ROOT_DIR / "site").mkdir(exist_ok=True)
        import shutil
        shutil.copy2(tokens_path, ROOT_DIR / "site" / "design-tokens.css")
        figma_tokens_css = tokens_path.read_text(encoding="utf-8")
    else:
        figma_tokens_css = ""

    # Camada V2: Design Tokens Figma, Cognitive UI/UX, IBCS, Tufte, 7 Estados de UI, Touch Targets WCAG e Mobile-First
    v2_css_enhancements = """
/* ==========================================================================
   V2.0 PRO MASTER STYLES - FIGMA ONE-X DESIGN SYSTEM, 7 ESTADOS & 60FPS GPU
   ========================================================================== */
*, *::before, *::after {
  box-sizing: border-box;
}

html, body {
  overflow-x: hidden !important;
  max-width: 100vw !important;
  margin: 0 !important;
  padding: 0 !important;
  font-family: var(--font-sans, 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif);
  background-color: var(--figma-bg-base, var(--bg, #0B0F19)) !important;
  color: var(--figma-text-primary, var(--ink, #F8FAFC)) !important;
}

body.modal-open {
  overflow: hidden !important;
  touch-action: none !important;
}

:root {
  --tr-fast: 150ms cubic-bezier(0.4, 0, 0.2, 1);
  --tr-smooth: 250ms cubic-bezier(0.16, 1, 0.3, 1);
  --ease-spring: cubic-bezier(0.175, 0.885, 0.32, 1.275);
  --font-mono: 'JetBrains Mono', monospace;
}

/* 60fps GPU Hardware Acceleration */
.kpi, .card, .painel, .ranking-podium, .btn, .btn-v2, .btn-excel, .omds, .toptab, .tab, .modal-panel {
  will-change: transform, box-shadow;
  transform: translateZ(0);
  backface-visibility: hidden;
}

/* Iconografia SVG Oficial Figma */
.figma-icon {
  display: inline-flex !important;
  align-items: center !important;
  justify-content: center !important;
  flex-shrink: 0 !important;
  vertical-align: middle !important;
}

/* Banner de Defasagem de Dados & Contingência */
.staleness-banner {
  display: none !important;
  align-items: center !important;
  gap: 0.75rem !important;
  padding: 0.85rem 1.25rem !important;
  margin-bottom: 1.25rem !important;
  border-radius: var(--radius-lg, 12px) !important;
  font-size: 0.875rem !important;
  font-weight: 600 !important;
  line-height: 1.45 !important;
  background: rgba(239, 68, 68, 0.14) !important;
  border: 1px solid rgba(239, 68, 68, 0.45) !important;
  color: #FCA5A5 !important;
  box-shadow: 0 4px 14px rgba(239, 68, 68, 0.15) !important;
}
.staleness-banner.show {
  display: flex !important;
}

/* Chips de Contagem Regressiva e Prazos de Empenho (OAA/2026) */
.chip-prazo {
  display: inline-flex !important;
  align-items: center !important;
  gap: 3px !important;
  padding: 1px 7px !important;
  border-radius: 999px !important;
  font-size: 0.68rem !important;
  font-weight: 700 !important;
  letter-spacing: 0.02em !important;
  margin-left: 6px !important;
  vertical-align: middle !important;
  line-height: 1.3 !important;
  white-space: nowrap !important;
}
.chip-prazo-critica {
  background: rgba(239, 68, 68, 0.18) !important;
  color: #F87171 !important;
  border: 1px solid rgba(239, 68, 68, 0.45) !important;
  box-shadow: 0 0 8px rgba(239, 68, 68, 0.25) !important;
}
.chip-prazo-atencao {
  background: rgba(245, 158, 11, 0.18) !important;
  color: #FBBF24 !important;
  border: 1px solid rgba(245, 158, 11, 0.45) !important;
}
.chip-prazo-normal {
  background: rgba(16, 185, 129, 0.15) !important;
  color: #34D399 !important;
  border: 1px solid rgba(16, 185, 129, 0.35) !important;
}
.chip-prazo-vencido {
  background: rgba(148, 163, 184, 0.15) !important;
  color: #94A3B8 !important;
  border: 1px solid rgba(148, 163, 184, 0.3) !important;
}

/* Botões de navegação lateral discretos para OMDS */
.omds-nav-wrap {
  position: relative !important;
  max-width: min(98vw, 1860px) !important;
  margin: 0 auto !important;
  display: flex !important;
  align-items: center !important;
}
.omds-nav-in {
  width: 100% !important;
  padding: 10px 42px !important;
}
.omds-scroll-btn {
  position: absolute !important;
  top: 50% !important;
  transform: translateY(-50%) !important;
  z-index: 40 !important;
  width: 32px !important;
  height: 32px !important;
  min-height: 32px !important;
  padding: 0 !important;
  border-radius: 50% !important;
  background: var(--figma-bg-card, #1E293B) !important;
  border: 1px solid var(--figma-border-strong, rgba(255, 255, 255, 0.2)) !important;
  color: var(--figma-text-primary, #F8FAFC) !important;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.4) !important;
  display: flex !important;
  align-items: center !important;
  justify-content: center !important;
  cursor: pointer !important;
  transition: opacity .18s ease, transform .18s ease, background-color .18s ease !important;
  opacity: 0.9 !important;
}
.omds-scroll-btn:hover {
  opacity: 1 !important;
  background: var(--surface, #334155) !important;
  transform: translateY(-50%) scale(1.08) !important;
}
.omds-scroll-btn:active {
  transform: translateY(-50%) scale(0.94) !important;
}
.omds-scroll-btn.hidden {
  opacity: 0 !important;
  pointer-events: none !important;
}
.omds-scroll-prev { left: 8px !important; }
.omds-scroll-next { right: 8px !important; }

/* Ocultação global e elegante de barras de rolagem nativas no Desktop */
.omds-nav-in, .toptabs, .tabs, .tbl-wrap {
  scrollbar-width: none !important;
  -ms-overflow-style: none !important;
}
.omds-nav-in::-webkit-scrollbar,
.toptabs::-webkit-scrollbar,
.tabs::-webkit-scrollbar,
.tbl-wrap::-webkit-scrollbar {
  display: none !important;
  width: 0 !important;
  height: 0 !important;
}
.staleness-banner.staleness-cache {
  background: rgba(245, 158, 11, 0.14) !important;
  border: 1px solid rgba(245, 158, 11, 0.45) !important;
  color: #FCD34D !important;
  box-shadow: 0 4px 14px rgba(245, 158, 11, 0.15) !important;
}
.badge-ajuste-centavos {
  background: rgba(148, 163, 184, 0.15) !important;
  color: #94A3B8 !important;
  padding: 2px 6px !important;
  border-radius: 4px !important;
  font-size: 0.7rem !important;
  font-weight: 600 !important;
  margin-left: 6px !important;
  border: 1px solid rgba(148, 163, 184, 0.3) !important;
}

/* CARDS, PAINÉIS & CONTAINERS COM DESIGN TOKENS FIGMA */
.card, .kpi, .painel, .ranking-podium, .tab-pane {
  background: var(--figma-bg-card, #111827) !important;
  border: 1px solid var(--figma-border, #1E293B) !important;
  border-radius: var(--radius-lg, 14px) !important;
  box-shadow: var(--figma-shadow-card, 0 4px 16px rgba(0, 0, 0, 0.35)) !important;
  transition: transform var(--tr-fast), box-shadow var(--tr-fast), border-color var(--tr-fast) !important;
}

.kpi:hover, .card:hover {
  box-shadow: var(--figma-shadow-hover, 0 8px 24px rgba(0, 0, 0, 0.45)) !important;
  border-color: var(--figma-border-strong, rgba(255, 255, 255, 0.16)) !important;
}

.kpi-lbl {
  color: var(--figma-text-secondary, #94A3B8) !important;
  font-size: var(--font-size-xs, 0.75rem) !important;
  font-weight: var(--font-weight-semibold, 600) !important;
  text-transform: uppercase !important;
  letter-spacing: 0.04em !important;
}

.kpi-val {
  color: var(--figma-text-primary, #F8FAFC) !important;
  font-family: var(--font-mono, monospace) !important;
  font-weight: var(--font-weight-black, 800) !important;
  font-variant-numeric: tabular-nums !important;
}

/* CONTAINER WIDESCREEN ERGONÔMICO (min(98vw, 1860px)) */
.wrap, .app-container {
  max-width: min(98vw, 1860px) !important;
  margin: 0 auto !important;
  padding: 1rem 1.25rem 3.5rem !important;
}

/* ==========================================================================
   SISTEMA DE BOTÕES & CONTROLES: GEOMETRIA, ALINHAMENTO & 7 ESTADOS
   ========================================================================== */
button, .btn, .btn-v2, .btn-excel, .flt-limpa, .m-btn-pill, .omds, .toptab, .tab, .pag-btn, .modal-x {
  display: inline-flex !important;
  align-items: center !important;
  justify-content: center !important;
  gap: 0.5rem !important;
  white-space: nowrap !important;
  line-height: 1 !important;
  cursor: pointer !important;
  user-select: none !important;
  text-decoration: none !important;
  transition: transform var(--tr-fast), background-color var(--tr-fast), border-color var(--tr-fast), box-shadow var(--tr-fast), color var(--tr-fast), filter var(--tr-fast) !important;
}

/* Alinhamento de ícones e SVGs internos */
button > svg, .btn > svg, .btn-v2 > svg, .m-btn-pill > svg, .btn-excel > svg {
  flex-shrink: 0 !important;
  vertical-align: middle !important;
}
button > span, .btn > span, .btn-v2 > span {
  display: inline-flex !important;
  align-items: center !important;
}

/* 7 ESTADOS DE UI */
/* 1. Idle: padrão de cada variante */

/* 2. Hover */
button:hover, .btn:hover, .btn-v2:hover, .m-btn-pill:hover, .omds:hover, .toptab:hover, .tab:hover {
  filter: brightness(1.06) !important;
  transform: translateY(-1px) !important;
}
.btn-excel:hover {
  filter: brightness(1.1) !important;
  box-shadow: 0 4px 12px rgba(16, 124, 65, 0.35) !important;
  transform: translateY(-1px) !important;
}

/* 3. Active (Fitts / Feedback Tátil) */
button:active, .btn:active, .btn-v2:active, .btn-excel:active, .flt-limpa:active, .m-btn-pill:active, .omds:active, .toptab:active, .tab:active {
  filter: brightness(0.94) !important;
  transform: translateY(1px) !important;
}

/* 4. Focus-Visible (WCAG 2.2 AAA) */
button:focus-visible, .btn:focus-visible, .btn-v2:focus-visible, .btn-excel:focus-visible, .flt-limpa:focus-visible, .m-btn-pill:focus-visible, .omds:focus-visible, .toptab:focus-visible, .tab:focus-visible {
  outline: 2px solid var(--primary-600, #2563EB) !important;
  outline-offset: 2px !important;
  box-shadow: 0 0 0 4px rgba(37, 99, 235, 0.18) !important;
}

/* 5. Disabled */
button:disabled, .btn:disabled, .btn-v2:disabled, .btn-excel:disabled, .flt-limpa:disabled, [aria-disabled="true"] {
  opacity: 0.45 !important;
  cursor: not-allowed !important;
  pointer-events: none !important;
  transform: none !important;
  filter: grayscale(0.6) !important;
}

/* 6. Loading (aria-busy="true") */
button[aria-busy="true"], .btn[aria-busy="true"] {
  cursor: wait !important;
  opacity: 0.75 !important;
  pointer-events: none !important;
}

/* 7. Success & Error */
.btn-success { background: var(--success-bg) !important; color: var(--success-strong) !important; border-color: var(--success-border) !important; }
.btn-danger, .btn-error { background: var(--danger-bg) !important; color: var(--danger) !important; border-color: var(--danger-border) !important; }

/* HIERARQUIA VISUAL (EFEITO VON RESTORFF) */
.btn-v2-primary {
  background: linear-gradient(135deg, #1C4A73, #2563EB) !important;
  color: #FFFFFF !important;
  border: 1px solid rgba(255, 255, 255, 0.22) !important;
  box-shadow: 0 2px 8px rgba(37, 99, 235, 0.25) !important;
  font-weight: 700 !important;
}

.btn-v2 {
  font-size: 0.825rem !important;
  font-weight: 600 !important;
  padding: 0.45rem 0.95rem !important;
  border-radius: 0.6rem !important;
  border: 1px solid var(--border) !important;
  background: var(--bg-subtle) !important;
  color: var(--ink) !important;
  min-height: 38px !important;
}

/* TOPBAR EXECUTIVA MODERNIZADA */
.topbar {
  display: flex !important;
  align-items: center !important;
  justify-content: space-between !important;
  gap: 1.5rem !important;
  padding: 1rem 1.5rem !important;
  background: var(--surface) !important;
  border: 1px solid var(--border) !important;
  border-radius: 1rem !important;
  box-shadow: 0 4px 16px rgba(0,0,0,0.06) !important;
  margin-bottom: 1.25rem !important;
  transition: background-color var(--tr-smooth), border-color var(--tr-smooth) !important;
}

.topbar .brand {
  display: flex !important;
  align-items: center !important;
  gap: 1rem !important;
}

.topbar .brasao {
  width: 52px !important;
  height: 52px !important;
  border-radius: 0.75rem !important;
  padding: 4px !important;
  background: var(--bg-subtle) !important;
  border: 1px solid var(--border-strong) !important;
  box-shadow: 0 2px 8px rgba(0,0,0,0.15) !important;
  object-fit: contain !important;
  flex-shrink: 0 !important;
}

.topbar h1 {
  font-size: 1.35rem !important;
  font-weight: 800 !important;
  letter-spacing: -0.02em !important;
  display: flex !important;
  align-items: center !important;
  gap: 0.5rem !important;
}

.topbar .badge-v2 {
  font-size: 0.72rem !important;
  font-weight: 800 !important;
  font-family: var(--font-mono) !important;
  padding: 0.2rem 0.55rem !important;
  background: linear-gradient(135deg, #059669, #10B981) !important;
  color: #FFFFFF !important;
  border-radius: 999px !important;
  text-transform: uppercase !important;
  letter-spacing: 0.05em !important;
  vertical-align: middle !important;
  display: inline-block !important;
}

.topbar-actions {
  display: flex !important;
  align-items: center !important;
  gap: 0.65rem !important;
  flex-wrap: nowrap !important;
  flex-shrink: 0 !important;
}
@media (max-width: 1024px) {
  .topbar-actions {
    flex-wrap: wrap !important;
  }
}

.status-live-pill {
  display: inline-flex !important;
  align-items: center !important;
  gap: 0.45rem !important;
  font-size: 0.75rem !important;
  font-family: var(--font-mono) !important;
  font-weight: 700 !important;
  padding: 0.4rem 0.85rem !important;
  background: var(--success-bg) !important;
  color: var(--success-strong) !important;
  border: 1px solid var(--success-border) !important;
  border-radius: 999px !important;
  white-space: nowrap !important;
}

.status-live-dot {
  width: 8px !important;
  height: 8px !important;
  border-radius: 50% !important;
  background: var(--success) !important;
  box-shadow: 0 0 8px var(--success) !important;
  animation: pulse-dot 2s infinite !important;
}

@keyframes pulse-dot {
  0%, 100% { opacity: 1; transform: scale(1); }
  50% { opacity: 0.4; transform: scale(0.85); }
}

/* ==========================================================================
   PADRÃO IBCS & TUFTE EM TABELAS: PROIBIÇÃO DE DIVISORES VERTICAIS
   ========================================================================== */
table.det {
  border-collapse: collapse !important;
  width: 100% !important;
  font-size: 0.825rem !important;
  border: none !important;
}

table.det th, table.det td {
  border-left: none !important;
  border-right: none !important;
  border-top: none !important;
  border-bottom: 1px solid var(--border) !important;
  padding: 0.7rem 0.85rem !important;
  vertical-align: middle !important;
}

table.det thead th {
  position: sticky !important;
  top: 0 !important;
  z-index: 10 !important;
  background: var(--surface) !important;
  color: var(--ink-muted) !important;
  font-weight: 700 !important;
  font-size: 0.75rem !important;
  text-transform: uppercase !important;
  letter-spacing: 0.04em !important;
  border-bottom: 1px solid var(--border-strong) !important;
  user-select: none !important;
}

/* ALINHAMENTOS IBCS */
table.det .num, table.det td.num {
  text-align: right !important;
  font-variant-numeric: tabular-nums !important;
  font-family: var(--font-mono) !important;
  font-weight: 600 !important;
}

/* LINHA DE TOTAIS EM DESTAQUE (BORDA DUPLA SUPERIOR) */
table.det tfoot tr td, table.det tr.tot td {
  border-top: 3px double var(--border-strong) !important;
  border-bottom: none !important;
  font-weight: 700 !important;
  background: var(--surface) !important;
  color: var(--ink) !important;
  padding: 0.85rem !important;
}

/* LINHAS TOTALMENTE CLICÁVEIS COM CHEVRON E HOVER */
table.det tbody tr.cel-row {
  cursor: pointer !important;
  transition: background-color var(--tr-fast) !important;
}
table.det tbody tr.cel-row:hover {
  background-color: var(--bg-subtle) !important;
}
table.det tbody tr.cel-row:focus-visible {
  outline: 2px solid var(--primary-600) !important;
  outline-offset: -2px !important;
}

/* ==========================================================================
   ADAPTAÇÃO MOBILE-FIRST INTEGRAL (BREAKPOINTS: 1024px, 768px, 480px, 360px)
   ========================================================================== */

/* TABLETS & TELAS MÉDIAS (<= 1024px) */
@media (max-width: 1024px) {
  .kpis {
    grid-template-columns: repeat(2, 1fr) !important;
    gap: 12px !important;
  }
  .wrap, .app-container {
    padding: 0.85rem 1rem 3rem !important;
  }
}

/* DISPOSITIVOS MÓVEIS (<= 768px) - WCAG 2.2 TOUCH TARGET >= 48px */
@media (max-width: 768px) {
  /* Touch target mínimo 48px com espaçamento >= 8px */
  button, .btn, .btn-v2, .btn-excel, .flt-limpa, .m-btn-pill, .omds, .toptab, .tab, .pag-btn {
    min-height: 48px !important;
    padding: 0.65rem 1rem !important;
  }

  /* Topbar Mobile */
  .topbar {
    flex-direction: column !important;
    align-items: stretch !important;
    gap: 0.85rem !important;
    padding: 0.85rem 1rem !important;
    margin-bottom: 0.85rem !important;
  }
  .topbar .brand {
    gap: 0.75rem !important;
  }
  .topbar .brasao {
    width: 44px !important;
    height: 44px !important;
  }
  .topbar h1 {
    font-size: 1.15rem !important;
    flex-wrap: wrap !important;
  }
  .topbar .subtitle {
    font-size: 0.75rem !important;
    line-height: 1.3 !important;
  }
  .topbar-actions {
    display: flex !important;
    flex-wrap: wrap !important;
    gap: 0.5rem !important;
    width: 100% !important;
  }
  .topbar-actions .status-live-pill {
    width: 100% !important;
    justify-content: center !important;
    margin-bottom: 0.25rem !important;
  }
  .topbar-actions .btn-v2 {
    flex: 1 1 calc(50% - 0.25rem) !important;
    min-width: 130px !important;
    justify-content: center !important;
  }

  /* Barra de Navegação OMDS Mobile com Fade Lateral */
  .omds-nav {
    position: sticky !important;
    top: 0 !important;
    z-index: 35 !important;
  }
  .omds-nav-in {
    padding: 8px 12px !important;
    gap: 8px !important;
    overflow-x: auto !important;
    scroll-snap-type: x mandatory !important;
    -webkit-overflow-scrolling: touch !important;
    scrollbar-width: none !important;
  }
  .omds-nav-in::-webkit-scrollbar {
    display: none !important;
  }
  .omds {
    scroll-snap-align: start !important;
    flex: 0 0 auto !important;
    min-height: 48px !important;
    font-size: 0.8rem !important;
    padding: 6px 14px 6px 10px !important;
  }

  /* Sub-abas Toptabs Mobile (Scroll Horizontal Suave) */
  .toptabs {
    display: flex !important;
    overflow-x: auto !important;
    scroll-snap-type: x mandatory !important;
    -webkit-overflow-scrolling: touch !important;
    scrollbar-width: none !important;
    padding: 4px !important;
    gap: 6px !important;
    margin: 16px 0 12px !important;
  }
  .toptabs::-webkit-scrollbar {
    display: none !important;
  }
  .toptab {
    flex: 0 0 auto !important;
    white-space: nowrap !important;
    scroll-snap-align: start !important;
    min-height: 48px !important;
    padding: 8px 16px !important;
    font-size: 0.8125rem !important;
  }

  /* Abas UASG / Consolidado (.tabs) */
  .tabs {
    overflow-x: auto !important;
    scroll-snap-type: x mandatory !important;
    -webkit-overflow-scrolling: touch !important;
    scrollbar-width: none !important;
    gap: 6px !important;
    margin-bottom: 12px !important;
  }
  .tabs::-webkit-scrollbar {
    display: none !important;
  }
  .tab {
    flex: 0 0 auto !important;
    scroll-snap-align: start !important;
    min-height: 48px !important;
    padding: 8px 14px !important;
  }

  /* KPI Cards no Mobile */
  .kpis {
    grid-template-columns: 1fr !important;
    gap: 10px !important;
  }
  .kpi {
    padding: 12px 14px !important;
  }
  .kpi-val {
    font-size: 1.35rem !important;
  }

  /* Barras de Ferramentas e Filtros (.tbl-tools) */
  .tbl-tools {
    flex-direction: column !important;
    align-items: stretch !important;
    gap: 8px !important;
  }
  .tbl-search {
    width: 100% !important;
    min-width: 0 !important;
    min-height: 48px !important;
  }
  .btn-excel, .btn-excel-lg {
    width: 100% !important;
    min-height: 48px !important;
    justify-content: center !important;
  }
  .flt-limpa {
    width: 100% !important;
    min-height: 44px !important;
    justify-content: center !important;
  }

  /* Pódio de Ranking no Mobile */
  .ranking-podium {
    flex-direction: column !important;
    align-items: stretch !important;
    gap: 12px !important;
  }
  .podium-step {
    width: 100% !important;
    min-height: 100px !important;
  }

  /* ==========================================================================
     MODAL DRILL-DOWN EM FORMATO BOTTOM SHEET NO MOBILE
     ========================================================================== */
  #modal {
    align-items: flex-end !important;
    padding: 0 !important;
  }
  .modal-panel {
    max-width: 100% !important;
    width: 100% !important;
    max-height: 92dvh !important;
    max-height: 92vh !important;
    border-radius: 20px 20px 0 0 !important;
    margin: 0 !important;
    border-bottom: none !important;
    animation: sheetSlideUp 0.28s cubic-bezier(0.16, 1, 0.3, 1) !important;
  }
  @keyframes sheetSlideUp {
    from { transform: translateY(100%); }
    to { transform: translateY(0); }
  }
  .modal-x {
    top: 12px !important;
    right: 12px !important;
    width: 48px !important;
    height: 48px !important;
    font-size: 1.25rem !important;
  }
  .m-btn-pill {
    min-height: 44px !important;
    padding: 8px 14px !important;
  }

  /* Rodapé Mobile */
  .rodape-v2 {
    flex-direction: column !important;
    align-items: flex-start !important;
    gap: 1rem !important;
    padding: 1.25rem 1rem !important;
  }
}

/* TELAS COMPACTAS (<= 480px / 360px a 480px) */
@media (max-width: 480px) {
  .app-container, .wrap {
    padding: 0.5rem 0.65rem 3rem !important;
  }
  .topbar-actions .btn-v2 {
    flex: 1 1 100% !important;
  }
  .topbar h1 {
    font-size: 1.05rem !important;
  }
}

@media (prefers-reduced-motion: reduce) {
  *, *::before, *::after {
    animation-duration: 0.01ms !important;
    animation-iteration-count: 1 !important;
    transition-duration: 0.01ms !important;
    scroll-behavior: auto !important;
  }
}
"""

    # Enhanced Header Replacement with Figma Design System SVGs (preserving all IDs: emblema, uTitulo, uNome, uUasg, themeBtn)
    v2_header_replacement = """
<header class="topbar">
  <div class="brand">
    <img class="brasao" id="emblema" src="assets/logos/BCMS.png" alt="Brasão BCMS" loading="eager" width="52" height="52">
    <div>
      <h1 id="uTitulo">
        Crédito Disponível — BCMS
        <span class="badge-v2">v2.0 PRO</span>
      </h1>
      <p class="subtitle"><span id="uNome">Batalhão Central de Manutenção e Suprimento</span> · Tesouro Gerencial / SIAFI · <span id="uUasg">UASGs 160329 (OGU) e 167329 (FEx)</span></p>
    </div>
  </div>
  <div class="topbar-actions">
    <div class="status-live-pill">
      <span class="status-live-dot"></span>
      <svg class="figma-icon" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><polyline points="9 12 11 14 15 10"/></svg>
      <span>SIAFI SINCRONIZADO</span>
    </div>
    <button class="btn-v2" onclick="bcmsExportHistoricoExcel()" title="Exportar Acervo Geral de Notas de Crédito">
      <svg class="figma-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
      <span>Exportar NCs</span>
    </button>
    <button class="btn-v2" onclick="bcmsExportRastreioExcel('ne')" title="Exportar Notas de Empenho Emitidas">
      <svg class="figma-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><path d="M8 13h8"/><path d="M8 17h8"/><path d="M10 9h1"/></svg>
      <span>Exportar NEs</span>
    </button>
    <button class="btn-v2 btn-v2-primary" onclick="window.print()" title="Imprimir Relatório Executivo em formato A4">
      <svg class="figma-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>
      <span>A4</span>
    </button>
    <button class="theme btn-v2" id="themeBtn" aria-pressed="false" aria-label="Alternar tema claro/escuro" onclick="bcmsTheme()" title="Alternar tema">
      <svg viewBox="0 0 24 24" class="ic-sun" style="width:16px;height:16px;" aria-hidden="true"><circle cx="12" cy="12" r="4.5" style="fill:currentColor"/><g style="stroke:currentColor;stroke-width:1.8;stroke-linecap:round"><path d="M12 2v2.5M12 19.5v2.5M2 12h2.5M19.5 12h2.5M4.93 4.93l1.77 1.77M17.3 17.3l1.77 1.77M19.07 4.93l-1.77 1.77M6.7 17.3l-1.77 1.77"/></g></svg>
      <svg viewBox="0 0 24 24" class="ic-moon" style="width:16px;height:16px;" aria-hidden="true"><path d="M20 14.5A8 8 0 019.5 4 8 8 0 1020 14.5z" style="fill:currentColor"/></svg>
      <span>Tema</span>
    </button>
  </div>
</header>
"""

    # Enhanced Footer Replacement
    v2_footer_replacement = """
<footer class="rodape-v2">
  <div class="rodape-v2-left">
    <div style="font-size: 1.8rem;">🏛️</div>
    <div class="rodape-v2-credits">
      <div><strong>Base de Apoio Logístico do Exército (Ba Ap Log Ex)</strong> · Diretoria de Gestão Orçamentária e Financeira</div>
      <div style="font-size: 0.8rem; margin-top: 0.25rem;">
        💻 <b>Desenvolvido por:</b> 3º Sgt De Campos · Corte Oficial SIAFI / Tesouro Gerencial 2026
      </div>
      <div style="font-size: 0.75rem; color: var(--ink-soft); margin-top: 0.15rem;">
        <b>Metodologia:</b> Crédito Disponível = Provisão Recebida − Provisão Concedida − Despesas Empenhadas (saldo líquido não empenhado no Tesouro Gerencial / SIAFI).
      </div>
    </div>
  </div>
  <div>
    <span class="status-live-pill" style="font-size: 0.75rem;">Versão 2.0.0 PRO · WCAG 2.2 AAA · 100% Offline / Nativo</span>
  </div>
</footer>
"""

    # Replace header and footer in body
    body_inner = re.sub(r'<header class="topbar">.*?</header>', v2_header_replacement, body_inner, flags=re.DOTALL)
    body_inner = re.sub(r'<footer class="rodape">.*?</footer>', v2_footer_replacement, body_inner, flags=re.DOTALL)

    # Serialize datasets
    print("3. Serializando variáveis de dados para o bundle...")
    embedded_metadata = json.dumps(raw_data.get("metadata", {}), ensure_ascii=False)
    embedded_celdata = json.dumps(raw_data["celulas"], ensure_ascii=False)
    embedded_ncdata = json.dumps(raw_data["ncdata"], ensure_ascii=False)
    embedded_daydata = json.dumps(raw_data["daydata"], ensure_ascii=False)
    embedded_teladata = json.dumps(raw_data["teladata"], ensure_ascii=False)
    embedded_unidades = json.dumps(raw_data["unidades"], ensure_ascii=False)
    embedded_histdata = json.dumps(raw_data["histdata"], ensure_ascii=False)
    embedded_empenhodata = json.dumps(raw_data["empenhodata"], ensure_ascii=False)
    embedded_history_series = json.dumps(hist_data, ensure_ascii=False)

    # Injeção de melhorias de runtime no JS: Scroll Lock, Trava Anti-Zero e Redimensionamento Responsivo
    v2_runtime_js_patch = """
    // 1. Trava de precisão financeira anti-zero negativo
    (function(){
      var origFmt = window.fmtMoeda;
      if (typeof origFmt === 'function') {
        window.fmtMoeda = function(val) {
          if (val === null || val === undefined || isNaN(val)) return "R$ 0,00";
          if (Math.abs(val) < 0.005) return "R$ 0,00";
          return origFmt(val);
        };
      }
    })();

    // 2. Trava de scroll de modal e observador de mutação (WCAG Focus & Scroll Lock)
    (function(){
      var m = document.getElementById('modal');
      if(!m) return;
      var obs = new MutationObserver(function(muts){
        muts.forEach(function(mut){
          if(mut.attributeName === 'class'){
            if(m.classList.contains('open')){
              document.body.classList.add('modal-open');
              document.body.style.overflow = 'hidden';
            } else {
              document.body.classList.remove('modal-open');
              document.body.style.overflow = '';
            }
          }
        });
      });
      obs.observe(m, { attributes: true });
    })();

    // 3. Suporte universal a teclado (Enter/Space) em linhas clicáveis de tabelas
    document.addEventListener('keydown', function(e){
      if(e.key === 'Enter' || e.key === ' '){
        var row = e.target.closest('tr.cel-row');
        if(row && document.activeElement === row){
          e.preventDefault();
          row.click();
        }
      }
    });

    // 5. Navegação lateral com setas discretas da barra OMDS
    window.bcmsScrollOMDS = function(direction) {
      var el = document.getElementById('omdsNavIn') || document.querySelector('.omds-nav-in');
      if (!el) return;
      var step = Math.max(220, Math.floor(el.clientWidth * 0.65));
      el.scrollBy({ left: direction * step, behavior: 'smooth' });
    };

    window.bcmsUpdateOMDSScrollButtons = function() {
      var el = document.getElementById('omdsNavIn') || document.querySelector('.omds-nav-in');
      var prev = document.getElementById('omdsScrollPrev');
      var next = document.getElementById('omdsScrollNext');
      if (!el || !prev || !next) return;
      var maxScroll = el.scrollWidth - el.clientWidth;
      if (maxScroll <= 8) {
        prev.classList.add('hidden');
        next.classList.add('hidden');
        return;
      }
      if (el.scrollLeft <= 8) {
        prev.classList.add('hidden');
      } else {
        prev.classList.remove('hidden');
      }
      if (el.scrollLeft >= maxScroll - 8) {
        next.classList.add('hidden');
      } else {
        next.classList.remove('hidden');
      }
    };

    window.addEventListener('resize', window.bcmsUpdateOMDSScrollButtons);
    setTimeout(window.bcmsUpdateOMDSScrollButtons, 120);
    setTimeout(window.bcmsUpdateOMDSScrollButtons, 600);

    // 4. Verificação de Defasagem de Dados e Alerta de Validade
    (function(){
      var meta = (typeof RAW_METADATA !== 'undefined' ? RAW_METADATA : null);
      var b = document.getElementById('stalenessBanner');
      var txt = document.getElementById('stalenessText');
      if (!b || !txt) return;
      if (!meta) {
        b.classList.remove('show');
        b.style.display = 'none';
        return;
      }
      var iso = meta.gerado_em_iso;
      var usandoCache = meta.usando_cache;
      var agora = new Date();
      var dt = iso ? new Date(iso) : null;
      var diffHoras = dt ? (agora - dt) / (1000 * 3600) : 0;
      if (usandoCache) {
        b.className = 'staleness-banner staleness-cache show';
        txt.innerHTML = '<strong>ATENÇÃO (MODO DE CONTINGÊNCIA / CACHE OFFLINE):</strong> O download do Google Drive não respondeu. Exibindo dados locais offline gerados em <b>' + (meta.gerado_em || 'data anterior') + '</b>. Novas notas emitidas recentemente podem não constar.';
      } else if (diffHoras > 24) {
        b.className = 'staleness-banner show';
        txt.innerHTML = '<strong>ALERTA DE DEFASAGEM ORÇAMENTÁRIA:</strong> Base de dados gerada há mais de ' + Math.floor(diffHoras) + ' horas (' + (meta.gerado_em || '') + '). Verifique se o Google Drive foi abastecido com o último relatório do SIAFI / Tesouro Gerencial.';
      } else {
        b.classList.remove('show');
        b.style.display = 'none';
      }
    })();
"""

    print("4. Montando HTML final unificado...")
    full_v2_html = f"""<!doctype html>
<html lang="pt-BR" style="--accent:#DC2626">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
  <meta name="color-scheme" content="light dark">
  <meta name="description" content="Painel Executivo de Crédito Disponível & Empenhos 2.0 - Base de Apoio Logístico do Exército e OMDS Vinculadas">
  <title>Crédito Disponível 2.0 · OMDS Ba Ap Log Ex</title>
  <link rel="stylesheet" href="design-tokens.css">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600;700&family=Newsreader:ital,opsz,wght@0,6..72,500;0,6..72,600;0,6..72,700;1,6..72,400&display=swap" rel="stylesheet">
  <style>
{figma_tokens_css}

{v1_css}

{v2_css_enhancements}
  </style>
</head>
<body>
  <div class="bcms-bar" aria-hidden="true"></div>
  <div class="app-container wrap">
    <div id="stalenessBanner" class="staleness-banner" style="display:none;" role="alert">
      <svg class="figma-icon" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
      <div id="stalenessText"></div>
    </div>
{body_inner}
  </div>

  <script>
    var RAW_METADATA = {embedded_metadata};
    var CELDATA = {embedded_celdata};
    var NCDATA = {embedded_ncdata};
    var DAYDATA = {embedded_daydata};
    var TELADATA = {embedded_teladata};
    var UNIDADES = {embedded_unidades};
    var HISTDATA = {embedded_histdata};
    var EMPENHODATA = {embedded_empenhodata};
    var MASTER_HISTORY = {embedded_history_series};

{v2_runtime_js_patch}

{js_logic}
  </script>
</body>
</html>
"""

    print(f"5. Gravando arquivo em {OUT_HTML_PATH}...")
    OUT_HTML_PATH.write_text(full_v2_html, encoding="utf-8")
    print(f"✓ Sucesso! Gerado {OUT_HTML_PATH} ({OUT_HTML_PATH.stat().st_size:,} bytes)")

if __name__ == "__main__":
    build()
