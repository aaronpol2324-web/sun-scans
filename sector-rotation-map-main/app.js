/* ================================================================
   Sector Rotation Map — Complete Live Dashboard Script
   ================================================================ */

const API_BASE = "";

const state = {
  benchmark: "SPY",
  tailLength: 8,
  customSymbols: [],
  sectors: {},
  holdings: {},
  themeStructure: {},
  highlighted: null,
  loading: false,
  currentStatsGroup: "sectors", // Tracks the active universe for stats table & RRG chart
  currentMegaTheme: null,       // Set when drilled into a specific mega theme
  currentSubTheme: null,        // Set when drilled into a specific sub-theme
  navHistory: [],               // Stack of previous {group, megaTheme, subTheme} snapshots, for Back
  viewMode: "rrg",              // "rrg" or "charts"
};

let themeRankingsMap = {}; // Maps "Mega::Sub" to its ranking metrics

const BENCHMARKS = [
  { id: "SPY", label: "SPY" },
  { id: "QQQ", label: "QQQ" },
  { id: "IWM", label: "IWM" },
];

const TAIL_LENGTHS = [6, 8, 12, 16, 20, 52];

const BENCHMARK_NAMES = {
  SPY: "SPY",
  QQQ: "QQQ",
  IWM: "IWM",
  XLK: "Technology",
  XLF: "Financials",
  XLV: "Health Care",
  XLE: "Energy",
  XLP: "Consumer Staples",
  XLB: "Materials",
  XLRE: "Real Estate",
  XLI: "Industrials",
  XLU: "Utilities",
  XLY: "Discretionary",
  XLC: "Comm Services",
};

async function fetchRRGData() {
  // Sync the RRG chart with whatever universe the stats table is currently showing -
  // sectors, mega themes, sub-themes, or a specific drill-down - exactly like group-stats.
  const params = new URLSearchParams({
    benchmark: state.benchmark,
    tail: state.tailLength,
  });

  if (state.currentSubTheme) {
    params.set("sub_theme", state.currentSubTheme);
  } else if (state.currentMegaTheme) {
    params.set("mega_theme", state.currentMegaTheme);
  } else {
    params.set("group", state.currentStatsGroup);
  }

  // customSymbols here means individually-added extra tickers (e.g. clicked from the
  // theme sidebar), layered on top of whatever the group/drill-down above resolves to -
  // NOT a replacement for it.
  const extra = state.customSymbols.filter(s => s !== state.benchmark).join(",");
  if (extra) params.set("extra", extra);

  const url = `${API_BASE}/api/rrg?${params.toString()}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`API error: ${res.status}`);
  return res.json();
}

async function fetchHoldings() {
  const res = await fetch(`${API_BASE}/api/holdings`);
  if (!res.ok) throw new Error("Failed to load holdings");
  return res.json();
}

async function fetchThemeStructure() {
  const res = await fetch(`${API_BASE}/api/theme-structure`);
  if (!res.ok) throw new Error("Failed to load theme structure");
  return res.json();
}

/* ── Load Theme Rankings and Index them for the Sidebar ── */
async function loadThemeRankingsForSidebar() {
    try {
        const response = await fetch(`${API_BASE}/api/theme-rankings?benchmark=${state.benchmark}`);
        const data = await response.json();

        let items = [];
        if (Array.isArray(data)) {
            items = data;
        } else if (data && Array.isArray(data.rankings)) {
            items = data.rankings;
        }

        themeRankingsMap = {};
        items.forEach(item => {
            const key = `${item.mega_theme}::${item.sub_theme}`;
            themeRankingsMap[key] = item;
        });

        renderThemePanel();
    } catch (err) {
        console.error('Error fetching theme rankings for sidebar:', err);
    }
}

/* ── Render Theme Panel with Embedded RS Scores (Column 3) ── */
function renderThemePanel() {
  const panel = document.getElementById("themeTreeContainer");
  if (!panel) return;

  panel.innerHTML = "";
  const structure = state.themeStructure;

  if (!structure || Object.keys(structure).length === 0) {
    panel.innerHTML = `<div style="color: #8b949e; font-size: 11px; padding: 8px;">No themes loaded</div>`;
    return;
  }

  for (const [megaThemeName, subThemes] of Object.entries(structure)) {
    const megaGroup = document.createElement("div");
    megaGroup.style.cssText = "margin-bottom: 12px; cursor: pointer;";

    const megaTitle = document.createElement("div");
    megaTitle.style.cssText = "color: #58a6ff; font-weight: bold; font-size: 11px; margin-bottom: 4px; text-transform: uppercase; letter-spacing: 0.5px; display: flex; justify-content: space-between; align-items: center;";

    megaTitle.innerHTML = `<span>📁 ${megaThemeName}</span>`;
    megaGroup.appendChild(megaTitle);

    const subGroupContainer = document.createElement("div");
    subGroupContainer.style.cssText = "margin-left: 8px; display: none;";

    let allMegaTickers = [];

    if (subThemes && typeof subThemes === "object") {
      for (const [subThemeName, tickers] of Object.entries(subThemes)) {
        if (Array.isArray(tickers)) {
          allMegaTickers = allMegaTickers.concat(tickers);
        }

        const rankKey = `${megaThemeName}::${subThemeName}`;
        const rankData = themeRankingsMap[rankKey];
        const compRs = rankData ? rankData.composite_rs : null;
        const rsColor = compRs === null ? "#8b949e" : (compRs >= 0 ? "#4ade80" : "#f87171");
        const rsText = compRs !== null ? `${compRs > 0 ? '+' : ''}${compRs}%` : '...';

        const subDiv = document.createElement("div");
        subDiv.style.cssText = "margin-bottom: 6px; border-left: 2px solid #30363d; padding-left: 6px; cursor: pointer; display: flex; justify-content: space-between; align-items: center;";

        const subTitle = document.createElement("div");
        subTitle.style.cssText = "color: #e6edf3; font-size: 11px; font-weight: 600;";
        subTitle.textContent = `↳ ${subThemeName}`;

        const rsBadge = document.createElement("span");
        rsBadge.style.cssText = `font-size: 10px; font-family: 'JetBrains Mono', monospace; font-weight: bold; color: ${rsColor}; background: #21262d; padding: 1px 5px; border-radius: 3px; border: 1px solid #30363d;`;
        rsBadge.textContent = rsText;

        subDiv.appendChild(subTitle);
        subDiv.appendChild(rsBadge);

        subDiv.onclick = (e) => {
          e.stopPropagation();
          updateTickerSidebar(subThemeName, tickers, rankData);
        };

        subGroupContainer.appendChild(subDiv);
      }
    }

    megaGroup.onclick = () => {
      const isVisible = subGroupContainer.style.display === "block";
      subGroupContainer.style.display = isVisible ? "none" : "block";
      updateTickerSidebar(megaThemeName, [...new Set(allMegaTickers)], null);
    };

    megaGroup.appendChild(subGroupContainer);
    panel.appendChild(megaGroup);
  }
}

/* ── Helper: Update Column 4 Ticker Metrics Sidebar with Live Performance ── */
async function updateTickerSidebar(title, tickers, rankData) {
  const titleEl = document.getElementById("tickerMetricsTitle");
  const listEl = document.getElementById("tickerMetricsList");

  if (titleEl) {
      if (rankData) {
          titleEl.innerHTML = `${title} <span style="font-size:10px; font-family:'JetBrains Mono',monospace; color:${rankData.composite_rs >= 0 ? '#4ade80':'#f87171'}">(${rankData.composite_rs}%)</span>`;
      } else {
          titleEl.textContent = title;
      }
  }

  if (!listEl) return;
  listEl.innerHTML = `<div style="color: #8b949e; font-size: 11px; padding: 8px;">Loading ticker metrics...</div>`;

  if (!tickers || tickers.length === 0) {
    listEl.innerHTML = `<div style="color: #8b949e; font-size: 11px; padding: 8px; font-style: italic;">No tickers found</div>`;
    return;
  }

  try {
    const res = await fetch(`${API_BASE}/api/theme-ticker-stats`);
    const statsData = await res.ok ? await res.json() : [];

    const statsMap = {};
    statsData.forEach(item => {
        statsMap[item.symbol] = item;
    });

    listEl.innerHTML = "";

    tickers.forEach(ticker => {
      const badge = document.createElement("div");
      badge.style.cssText = "display: flex; justify-content: space-between; align-items: center; background: #21262d; color: #c9d1d9; border: 1px solid #30363d; padding: 6px 10px; margin-bottom: 4px; border-radius: 4px; font-size: 11px; font-family: 'JetBrains Mono', monospace; cursor: pointer;";

      let perfHtml = '<span style="font-size: 9px; color: #8b949e;">No stats</span>';

      const tickerStats = statsMap[ticker];
      if (tickerStats && tickerStats.chg_1m !== undefined) {
          const chg1M = tickerStats.chg_1m * 100;
          const color = chg1M >= 0 ? "#4ade80" : "#f87171";
          perfHtml = `<span style="font-size: 10px; font-weight: bold; color: ${color};">${chg1M > 0 ? '+' : ''}${chg1M.toFixed(2)}%</span>`;
      }

      badge.innerHTML = `
        <span style="font-weight: bold; color: #e6edf3;">${ticker}</span>
        ${perfHtml}
      `;

      badge.onclick = () => {
        if (!state.customSymbols.includes(ticker)) {
          state.customSymbols.push(ticker);
          loadData();
        }
      };

      listEl.appendChild(badge);
    });

  } catch (err) {
    console.error("Failed to load ticker sidebar stats:", err);
    listEl.innerHTML = `<div style="color: #f87171; font-size: 11px; padding: 8px;">Error loading metrics</div>`;
  }
}

/* ── Render D3 Chart ── */
/* ── View Mode Toggle: RRG chart vs. Candlestick Grid ── */
function setViewMode(mode) {
  if (state.viewMode === mode) return;
  state.viewMode = mode;

  const rrgEl = document.getElementById("chartWrapper");
  const gridEl = document.getElementById("chartGridContainer");
  if (rrgEl) rrgEl.style.display = mode === "rrg" ? "" : "none";
  if (gridEl) gridEl.style.display = mode === "charts" ? "" : "none";

  renderSidebar(); // refresh button active-states

  if (mode === "charts") {
    loadChartGrid();
  }
}

/* Render one small SVG OHLC bar chart into a container element */
function renderMiniCandleChart(container, candles, width, height) {
  if (!candles || candles.length === 0) {
    container.innerHTML = '<div style="color:#8b949e; font-size:11px; text-align:center; padding-top:40px;">No data</div>';
    return;
  }

  const margin = { top: 16, right: 4, bottom: 16, left: 4 };
  const innerW = width - margin.left - margin.right;
  const innerH = height - margin.top - margin.bottom;

  const highs = candles.map(d => d.h);
  const lows = candles.map(d => d.l);
  const maxV = Math.max(...highs);
  const minV = Math.min(...lows);
  const pad = (maxV - minV) * 0.05 || 1;
  const yMax = maxV + pad;
  const yMin = minV - pad;

  const n = candles.length;
  const slot = innerW / n;
  const tickLen = Math.max(2, Math.min(8, slot * 0.4));

  const yScale = (v) => margin.top + innerH - ((v - yMin) / (yMax - yMin)) * innerH;

  const GRID_COLOR = "rgba(255,255,255,0.08)";
  const DIM_BAR = "rgba(255,255,255,0.55)";
  const BRIGHT_BAR = "#ffffff";
  const gradId = `grad-${Math.random().toString(36).slice(2, 9)}`;

  let svg = `<svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" style="background:#000;">`;
  svg += `<defs><linearGradient id="${gradId}" x1="0" y1="0" x2="0" y2="1">
    <stop offset="0%" stop-color="#ffffff" stop-opacity="0.16"/>
    <stop offset="100%" stop-color="#ffffff" stop-opacity="0"/>
  </linearGradient></defs>`;

  // Subtle horizontal gridlines
  const gridLines = 3;
  for (let i = 1; i <= gridLines; i++) {
    const y = margin.top + (innerH / (gridLines + 1)) * i;
    svg += `<line x1="${margin.left}" y1="${y}" x2="${width - margin.right}" y2="${y}" stroke="${GRID_COLOR}" stroke-width="1" stroke-dasharray="2,3"/>`;
  }

  // Area fill under the close-price line (drawn behind the bars)
  const closePts = candles.map((d, i) => [margin.left + i * slot + slot / 2, yScale(d.c)]);
  const areaBottom = margin.top + innerH;
  let areaPath = `M ${closePts[0][0]},${areaBottom} `;
  closePts.forEach(p => { areaPath += `L ${p[0]},${p[1]} `; });
  areaPath += `L ${closePts[closePts.length - 1][0]},${areaBottom} Z`;
  svg += `<path d="${areaPath}" fill="url(#${gradId})" stroke="none"/>`;

  // OHLC bars - last bar brighter/thicker so "now" stands out
  candles.forEach((d, i) => {
    const isLast = i === candles.length - 1;
    const color = isLast ? BRIGHT_BAR : DIM_BAR;
    const sw = isLast ? 1.4 : 1.2;
    const x = margin.left + i * slot + slot / 2;
    svg += `<line x1="${x}" y1="${yScale(d.h)}" x2="${x}" y2="${yScale(d.l)}" stroke="${color}" stroke-width="${sw}"/>`;
    svg += `<line x1="${x - tickLen}" y1="${yScale(d.o)}" x2="${x}" y2="${yScale(d.o)}" stroke="${color}" stroke-width="${sw}"/>`;
    svg += `<line x1="${x}" y1="${yScale(d.c)}" x2="${x + tickLen}" y2="${yScale(d.c)}" stroke="${color}" stroke-width="${sw}"/>`;
  });

  // High / low labels
  svg += `<text x="${margin.left}" y="${margin.top - 5}" fill="#8b949e" font-size="9" font-family="'JetBrains Mono', monospace">H ${maxV.toFixed(2)}</text>`;
  svg += `<text x="${margin.left}" y="${height - 5}" fill="#8b949e" font-size="9" font-family="'JetBrains Mono', monospace">L ${minV.toFixed(2)}</text>`;

  svg += `</svg>`;
  container.innerHTML = svg;
}

/* Fetch and render the candlestick grid for whatever's currently in the stats table */
const CHART_BATCH_SIZE = 30;

/* Fetch and render the candlestick grid for whatever's currently in the stats table,
   in bounded batches rather than one giant request - a filtered Sub-Themes view with
   no state filter can mean 150+ theme rows, each needing dozens of underlying tickers'
   OHLC pulled, which is a very different scale of request than the 11 sector ETFs. */
async function loadChartGrid() {
  const gridEl = document.getElementById("chartGridContainer");
  if (!gridEl) return;

  const fullData = state.lastFullStatsData || [];
  const rows = state.activeStateFilter
    ? fullData.filter(d => d.state === state.activeStateFilter)
    : fullData;

  state.chartGridRows = rows;
  state.chartGridLoaded = 0;

  if (rows.length === 0) {
    gridEl.innerHTML = '<div style="color:#8b949e; padding:20px; text-align:center;">Nothing to show for this filter.</div>';
    return;
  }

  gridEl.innerHTML = "";
  gridEl.style.display = "grid";
  gridEl.style.gridTemplateColumns = "repeat(auto-fill, minmax(420px, 1fr))";
  gridEl.style.gridAutoRows = "320px";
  gridEl.style.gap = "14px";

  await loadMoreCharts();
}

async function loadMoreCharts() {
  const gridEl = document.getElementById("chartGridContainer");
  if (!gridEl || !state.chartGridRows) return;

  const rows = state.chartGridRows;
  const start = state.chartGridLoaded;
  const batch = rows.slice(start, start + CHART_BATCH_SIZE);
  if (batch.length === 0) return;

  const existingFooter = document.getElementById("chartGridFooter");
  if (existingFooter) existingFooter.remove();

  const loadingMsg = document.createElement("div");
  loadingMsg.id = "chartGridLoadingMsg";
  loadingMsg.style.cssText = "grid-column: 1 / -1; text-align:center; padding:10px; color:#8b949e; font-size:12px;";
  loadingMsg.textContent = `Loading ${batch.length} more...`;
  gridEl.appendChild(loadingMsg);

  // JSON-encoded, not comma-joined: some theme names contain literal commas
  // (e.g. "Pipelines, LNG & Refining"), which would corrupt a plain comma-split list.
  const symbols = JSON.stringify(batch.map(r => r.symbol));
  const params = new URLSearchParams({ symbols, days: 90 });
  if (state.currentSubTheme) {
    params.set("sub_theme", state.currentSubTheme);
  } else if (state.currentMegaTheme) {
    params.set("mega_theme", state.currentMegaTheme);
  } else {
    params.set("group", state.currentStatsGroup);
  }

  let candleData = {};
  let failed = false;
  try {
    const res = await fetch(`${API_BASE}/api/candles?${params.toString()}`);
    if (res.ok) {
      candleData = await res.json();
    } else {
      failed = true;
      console.error(`Candle grid batch fetch failed: HTTP ${res.status}`);
    }
  } catch (err) {
    failed = true;
    console.error("Failed to load candle grid batch:", err);
  }

  // Bail out quietly if the user has since switched away from Charts mode or changed
  // the underlying view while this fetch was in flight.
  if (state.viewMode !== "charts") return;

  document.getElementById("chartGridLoadingMsg")?.remove();

  if (failed) {
    const errMsg = document.createElement("div");
    errMsg.style.cssText = "grid-column: 1 / -1; text-align:center; padding:10px; color:#f87171; font-size:12px;";
    errMsg.textContent = "Failed to load this batch of charts.";
    gridEl.appendChild(errMsg);
    return;
  }

  batch.forEach(row => {
    const entry = candleData[row.symbol];
    const card = document.createElement("div");
    card.style.cssText = "background:#161b22; border:1px solid #30363d; border-radius:6px; padding:6px; display:flex; flex-direction:column; transition: border-color 0.15s ease;";
    card.onmouseenter = () => { card.style.borderColor = "#58a6ff"; };
    card.onmouseleave = () => { card.style.borderColor = "#30363d"; };

    const header = document.createElement("div");
    header.style.cssText = "display:flex; justify-content:space-between; align-items:baseline; font-family:'JetBrains Mono',monospace; font-size:12px; color:#e6edf3; padding:0 4px 4px 4px;";

    const priceStr = row.price !== undefined && row.price !== null ? `$${row.price.toFixed(2)}` : "";
    let chgHtml = "";
    if (row.chg_1d !== undefined && row.chg_1d !== null) {
      const chgPct = row.chg_1d * 100;
      const chgColor = chgPct >= 0 ? "#4ade80" : "#f87171";
      chgHtml = `<span style="color:${chgColor}; font-weight:bold; margin-left:6px;">${chgPct >= 0 ? '+' : ''}${chgPct.toFixed(2)}%</span>`;
    }
    header.innerHTML = `<strong>${row.symbol}</strong><span style="color:#8b949e;">${priceStr}${chgHtml}</span>`;
    card.appendChild(header);

    const chartDiv = document.createElement("div");
    chartDiv.style.cssText = "flex:1; min-height:0;";
    card.appendChild(chartDiv);

    gridEl.appendChild(card);

    // Measure after insertion so the mini chart fills its actual card size
    requestAnimationFrame(() => {
      const w = chartDiv.clientWidth || 400;
      const h = chartDiv.clientHeight || 280;
      renderMiniCandleChart(chartDiv, entry ? entry.candles : [], w, h);
    });
  });

  state.chartGridLoaded += batch.length;

  const footer = document.createElement("div");
  footer.id = "chartGridFooter";
  footer.style.cssText = "grid-column: 1 / -1; text-align:center; padding:10px;";
  if (state.chartGridLoaded < rows.length) {
    const remaining = rows.length - state.chartGridLoaded;
    footer.innerHTML = `<button id="loadMoreChartsBtn" style="padding:6px 16px; background:#21262d; color:#e6edf3; border:1px solid #30363d; border-radius:4px; cursor:pointer; font-size:12px;">Load ${Math.min(CHART_BATCH_SIZE, remaining)} more (${state.chartGridLoaded}/${rows.length})</button>`;
  } else {
    footer.innerHTML = `<span style="color:#8b949e; font-size:12px;">Showing all ${rows.length}.</span>`;
  }
  gridEl.appendChild(footer);

  const loadMoreBtn = document.getElementById("loadMoreChartsBtn");
  if (loadMoreBtn) loadMoreBtn.onclick = () => loadMoreCharts();
}

function renderChart() {
  const container = document.getElementById("chartWrapper");
  if (!container) return;
  container.innerHTML = "";

  const width = container.clientWidth || 850;
  const height = container.clientHeight || 340;

  const svg = d3.select(container)
    .append("svg")
    .attr("width", width)
    .attr("height", height);

  svg.append("rect")
     .attr("width", width)
     .attr("height", height)
     .attr("fill", "#1b2a32");

  const sectors = Object.values(state.sectors);
  if (sectors.length === 0) return;

  let allX = [100], allY = [100];
  sectors.forEach(s => {
    if (s.tail) {
      s.tail.forEach(d => {
        allX.push(d.rs_ratio);
        allY.push(d.rs_momentum);
      });
    }
  });

  const rawXMin = d3.min(allX), rawXMax = d3.max(allX);
  const rawYMin = d3.min(allY), rawYMax = d3.max(allY);

  const xMin = Math.min(98, rawXMin - 0.5);
  const xMax = Math.max(102, rawXMax + 0.5);
  const yMin = Math.min(98, rawYMin - 0.5);
  const yMax = Math.max(102, rawYMax + 0.5);

  const padding = 50;
  const xScale = d3.scaleLinear().domain([xMin, xMax]).range([padding, width - padding]);
  const yScale = d3.scaleLinear().domain([yMin, yMax]).range([height - padding, padding]);

  const cx = xScale(100);
  const cy = yScale(100);

  svg.append("rect").attr("x", cx).attr("y", 0).attr("width", width - cx).attr("height", cy).attr("fill", "rgba(0, 151, 167, 0.08)");
  svg.append("rect").attr("x", 0).attr("y", 0).attr("width", cx).attr("height", cy).attr("fill", "rgba(250, 204, 21, 0.05)");
  svg.append("rect").attr("x", 0).attr("y", cy).attr("width", cx).attr("height", height - cy).attr("fill", "rgba(234, 88, 12, 0.06)");
  svg.append("rect").attr("x", cx).attr("y", cy).attr("width", width - cx).attr("height", height - cy).attr("fill", "rgba(192, 132, 252, 0.05)");

  const labelStyle = "font-family: Inter, sans-serif; font-size: 10px; font-weight: 600; letter-spacing: 1px; fill: #8b949e; opacity: 0.5;";
  svg.append("text").attr("x", cx + 15).attr("y", 24).text("LEADING").attr("style", labelStyle);
  svg.append("text").attr("x", padding + 10).attr("y", 24).text("IMPROVING").attr("style", labelStyle);
  svg.append("text").attr("x", padding + 10).attr("y", height - padding - 6).text("LAGGING").attr("style", labelStyle);
  svg.append("text").attr("x", cx + 15).attr("y", height - padding - 6).text("WEAKENING").attr("style", labelStyle);

  svg.append("line").attr("x1", cx).attr("y1", 0).attr("x2", cx).attr("y2", height).attr("stroke", "#4a6270").attr("stroke-dasharray", "4").attr("stroke-width", 1.2);
  svg.append("line").attr("x1", 0).attr("y1", cy).attr("x2", width).attr("y2", cy).attr("stroke", "#4a6270").attr("stroke-dasharray", "4").attr("stroke-width", 1.2);

  const tickCount = 5;
  const xTicks = xScale.ticks(tickCount);
  const yTicks = yScale.ticks(tickCount);

  svg.append("line").attr("x1", padding).attr("y1", height - padding).attr("x2", width - padding).attr("y2", height - padding).attr("stroke", "#30363d").attr("stroke-width", 1.5);
  xTicks.forEach(val => {
    const tx = xScale(val);
    svg.append("line").attr("x1", tx).attr("y1", height - padding).attr("x2", tx).attr("y2", height - padding + 4).attr("stroke", "#30363d");
    svg.append("text").attr("x", tx).attr("y", height - padding + 14).attr("fill", "#8b949e").attr("text-anchor", "middle").style("font-size", "9px").text(val.toFixed(1));
  });

  svg.append("line").attr("x1", padding).attr("y1", padding).attr("x2", padding).attr("y2", height - padding).attr("stroke", "#30363d").attr("stroke-width", 1.5);
  yTicks.forEach(val => {
    const ty = yScale(val);
    svg.append("line").attr("x1", padding - 4).attr("y1", ty).attr("x2", padding).attr("y2", ty).attr("stroke", "#30363d");
    svg.append("text").attr("x", padding - 8).attr("y", ty + 3).attr("fill", "#8b949e").attr("text-anchor", "end").style("font-size", "9px").text(val.toFixed(1));
  });

  const lineGen = d3.line()
    .x(d => xScale(d.rs_ratio))
    .y(d => yScale(d.rs_momentum));

  sectors.forEach(s => {
    if (!s.tail || s.tail.length === 0) return;
    const color = s.color || "#457b9d";
    const isIsolated = state.highlighted && state.highlighted !== s.symbol;
    const opacityVal = isIsolated ? 0.15 : 0.9;

    svg.append("path")
       .datum(s.tail)
       .attr("fill", "none")
       .attr("stroke", color)
       .attr("stroke-width", state.highlighted === s.symbol ? 2.5 : 1.5)
       .attr("opacity", opacityVal * 0.7)
       .attr("d", lineGen);

    s.tail.forEach((d, i) => {
      const isCurrent = (i === s.tail.length - 1);
      const progress = s.tail.length > 1 ? i / (s.tail.length - 1) : 1;

      const radius = isCurrent ? 5.5 : (1.5 + progress * 3);
      const dotOpacity = isCurrent ? 1 : (opacityVal * (0.4 + progress * 0.5));

      svg.append("circle")
         .attr("cx", xScale(d.rs_ratio))
         .attr("cy", yScale(d.rs_momentum))
         .attr("r", radius)
         .attr("fill", color)
         .attr("opacity", dotOpacity);

      if (isCurrent) {
        svg.append("circle")
           .attr("cx", xScale(d.rs_ratio))
           .attr("cy", yScale(d.rs_momentum))
           .attr("r", 2)
           .attr("fill", "#ffffff")
           .attr("opacity", 0.95);
      }
    });

    const current = s.tail[s.tail.length - 1];
    if (current) {
      const hx = xScale(current.rs_ratio);
      const hy = yScale(current.rs_momentum);
      svg.append("text")
         .attr("x", hx + 10)
         .attr("y", hy + 3)
         .text(s.symbol)
         .attr("fill", "#ffffff")
         .attr("opacity", opacityVal)
         .style("font-size", "11px")
         .style("font-weight", "bold");
    }
  });
}

/* ── Render Sidebar ── */
function renderSidebar() {
  const benchContainer = document.getElementById("benchmarkButtons");
  if (benchContainer) {
    benchContainer.innerHTML = "";
    BENCHMARKS.forEach(b => {
      const btn = document.createElement("button");
      btn.className = `btn ${state.benchmark === b.id ? "active" : ""}`;
      btn.style.cssText = `padding: 3px 8px; background: ${state.benchmark === b.id ? "#00838f" : "#21262d"}; color: #e6edf3; border: 1px solid #30363d; border-radius: 4px; cursor: pointer; font-size: 11px;`;
      btn.textContent = b.label;
      btn.onclick = () => {
        state.benchmark = b.id;
        loadData();
        loadThemeRankingsForSidebar();
      };
      benchContainer.appendChild(btn);
    });
  }

  const tailContainer = document.getElementById("tailButtons");
  if (tailContainer) {
    tailContainer.innerHTML = "";
    TAIL_LENGTHS.forEach(t => {
      const btn = document.createElement("button");
      btn.className = `btn ${state.tailLength === t ? "active" : ""}`;
      btn.style.cssText = `padding: 3px 8px; background: ${state.tailLength === t ? "#00838f" : "#21262d"}; color: #e6edf3; border: 1px solid #30363d; border-radius: 4px; cursor: pointer; font-size: 11px;`;
      btn.textContent = t;
      btn.onclick = () => {
        state.tailLength = t;
        loadData();
      };
      tailContainer.appendChild(btn);
    });
  }

  const viewModeContainer = document.getElementById("viewModeButtons");
  if (viewModeContainer) {
    viewModeContainer.innerHTML = "";
    [["rrg", "RRG"], ["charts", "Charts"]].forEach(([id, label]) => {
      const btn = document.createElement("button");
      btn.className = `btn ${state.viewMode === id ? "active" : ""}`;
      btn.style.cssText = `padding: 3px 8px; background: ${state.viewMode === id ? "#00838f" : "#21262d"}; color: #e6edf3; border: 1px solid #30363d; border-radius: 4px; cursor: pointer; font-size: 11px;`;
      btn.textContent = label;
      btn.onclick = () => setViewMode(id);
      viewModeContainer.appendChild(btn);
    });
  }

  const container = document.getElementById("sectorList");
  if (!container) return;
  container.innerHTML = "";

  const sectors = Object.values(state.sectors);
  if (sectors.length === 0) {
    container.innerHTML = `<div style="padding: 10px; color: #8b949e; text-align: center; font-size: 11px;">No sectors loaded</div>`;
    return;
  }

  const getQuadrantColor = (quadrant) => {
    if (!quadrant) return "#8b949e";
    const q = quadrant.toUpperCase();
    if (q.includes("LEADING")) return "#00bcd4";
    if (q.includes("IMPROVING")) return "#facc15";
    if (q.includes("LAGGING")) return "#ea580c";
    if (q.includes("WEAKENING")) return "#c084fc";
    return "#8b949e";
  };

  sectors.forEach(s => {
    const current = s.tail && s.tail.length > 0 ? s.tail[s.tail.length - 1] : null;
    const rsRatio = current ? current.rs_ratio : (s.rs_ratio || 0);
    const rsMom = current ? current.rs_momentum : (s.rs_momentum || 0);
    const fullName = BENCHMARK_NAMES[s.symbol] || s.description || s.symbol;
    const quadColor = getQuadrantColor(s.quadrant);

    const item = document.createElement("div");
    item.className = "sector-item";
    if (state.highlighted === s.symbol) {
      item.style.backgroundColor = "rgba(0, 131, 143, 0.25)";
    }

    item.onclick = () => {
      state.highlighted = (state.highlighted === s.symbol) ? null : s.symbol;
      renderChart();
      renderSidebar();
    };

    item.innerHTML = `
      <div style="display: flex; align-items: center; justify-content: space-between; padding: 7px 10px; border-bottom: 1px solid #30363d; cursor: pointer;">
        <div style="display: flex; align-items: center; gap: 6px; overflow: hidden;">
          <span style="width: 7px; height: 7px; border-radius: 50%; background-color: ${s.color || '#457b9d'}; display: inline-block; flex-shrink: 0;"></span>
          <div style="display: flex; flex-direction: column; overflow: hidden;">
            <strong style="color: #e6edf3; font-size: 11px; line-height: 1.2;">${s.symbol}</strong>
            <span style="color: #8b949e; font-size: 9px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 90px;">${fullName}</span>
          </div>
        </div>
        <div style="display: flex; flex-direction: column; align-items: flex-end; flex-shrink: 0; font-size: 9px;">
          <span style="color: ${quadColor}; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px;">${s.quadrant || ''}</span>
          <span style="color: #8b949e; font-family: 'JetBrains Mono', monospace; margin-top: 1px;">${rsRatio.toFixed(1)} / ${rsMom.toFixed(1)}</span>
        </div>
      </div>
    `;
    container.appendChild(item);
  });
}

/* ── Handle Stats Table Universe Switching & RRG Sync ── */
async function onStatsGroupChanged(groupVal) {
  pushNavHistory();

  state.currentStatsGroup = groupVal;
  state.currentMegaTheme = null;
  state.currentSubTheme = null;

  try {
    await loadData();
  } catch (err) {
    console.error("Failed to update stats group:", err);
  }
}

/* ── Load Dashboard Stats Table ── */
/* ── State Filter Helpers (Leading/Emerging/Neutral/Fading/Lagging/Breaking Down) ── */
function getStatsUrlParams() {
  if (state.currentSubTheme) return `sub_theme=${encodeURIComponent(state.currentSubTheme)}`;
  if (state.currentMegaTheme) return `mega_theme=${encodeURIComponent(state.currentMegaTheme)}`;
  return `group=${state.currentStatsGroup}`;
}

/* ── Navigation History (Back button for the stats table / RRG universe) ── */
function pushNavHistory() {
  state.navHistory.push({
    group: state.currentStatsGroup,
    megaTheme: state.currentMegaTheme,
    subTheme: state.currentSubTheme,
  });
  updateBackButton();
}

function updateBackButton() {
  const btn = document.getElementById("statsBackBtn");
  if (btn) btn.style.display = state.navHistory.length > 0 ? "" : "none";
}

async function goBackStats() {
  const prev = state.navHistory.pop();
  if (!prev) return;

  state.currentStatsGroup = prev.group;
  state.currentMegaTheme = prev.megaTheme;
  state.currentSubTheme = prev.subTheme;

  const sel = document.getElementById("statsGroupSelect");
  if (sel) sel.value = state.currentStatsGroup;

  updateBackButton();
  await loadData();
}

function applyStateFilterToSectors(sectorsObj, fullStatsData) {
  if (!state.activeStateFilter || !fullStatsData) return sectorsObj;
  const stateMap = Object.fromEntries(fullStatsData.map(d => [d.symbol, d.state]));
  const filtered = {};
  for (const [sym, val] of Object.entries(sectorsObj || {})) {
    if (stateMap[sym] === state.activeStateFilter) filtered[sym] = val;
  }
  return filtered;
}

async function setStateFilter(filterVal) {
  state.activeStateFilter = (state.activeStateFilter === filterVal) ? null : filterVal;
  await loadData();
}

const STATE_FILTER_ORDER = ["Leading", "Emerging", "Neutral", "Fading", "Lagging", "Breaking Down"];
const STATE_FILTER_COLORS = {
  "Leading": "#1b5e20",
  "Emerging": "#4caf50",
  "Neutral": "#546e7a",
  "Fading": "#e65100",
  "Lagging": "#b71c1c",
  "Breaking Down": "#7f0000",
};

function renderStateFilterBar(fullData) {
  const bar = document.getElementById("stateFilterBar");
  if (!bar) return;

  const counts = {};
  STATE_FILTER_ORDER.forEach(s => counts[s] = 0);
  fullData.forEach(d => { if (d.state && counts[d.state] !== undefined) counts[d.state]++; });

  const chip = (label, count, colorKey, isActive) => {
    const bg = colorKey ? STATE_FILTER_COLORS[colorKey] : "#30363d";
    const border = isActive ? "2px solid #e6edf3" : "2px solid transparent";
    return `<span data-filter="${colorKey || ''}" style="cursor:pointer; user-select:none; background-color:${bg}; color:#fff; padding:4px 10px; border-radius:12px; font-size:11px; font-family:'Inter',sans-serif; border:${border};">${label} ${count}</span>`;
  };

  const totalCount = fullData.length;
  let html = chip("All", totalCount, null, !state.activeStateFilter);
  STATE_FILTER_ORDER.forEach(s => {
    html += chip(s, counts[s], s, state.activeStateFilter === s);
  });
  bar.innerHTML = html;

  bar.querySelectorAll("[data-filter]").forEach(el => {
    el.onclick = () => setStateFilter(el.dataset.filter || null);
  });
}

/* ── Load Dashboard Stats Table with Drill-Down Support ── */
/* ── Load Dashboard Stats Table with Drill-Down Support ── */
async function loadStatsTable(urlParams = getStatsUrlParams()) {
    try {
        const response = await fetch(`${API_BASE}/api/group-stats?${urlParams}`);
        if (!response.ok) return;
        const rawData = await response.json();
        state.lastFullStatsData = rawData;

        const data = state.activeStateFilter
            ? rawData.filter(d => d.state === state.activeStateFilter)
            : rawData;

        const tbody = document.querySelector("#statsTable tbody");
        if (!tbody) return;

        renderStateFilterBar(rawData);

        let currentSortCol = null;
        let sortAscending = false;

        const getColorblindTint = (val) => {
          if (val === undefined || val === null || isNaN(val)) return "transparent";
          if (val === 0) return "transparent";
          if (val > 0.05) return "#004b5b";
          if (val > 0.01) return "#009688";
          if (val > 0.00) return "#4db6ac";
          if (val < -0.05) return "#9f3600";
          if (val < -0.01) return "#d86b00";
          if (val < -0.00) return "#ffa000";
          return "transparent";
        };

        const STATE_COLORS = {
          "Leading": "#1b5e20",
          "Emerging": "#4caf50",
          "Neutral": "#546e7a",
          "Fading": "#e65100",
          "Lagging": "#b71c1c",
          "Breaking Down": "#7f0000",
        };
        const getStateBadge = (state) => {
          if (!state) return '<span style="color:#8b949e;">&mdash;</span>';
          const bg = STATE_COLORS[state] || "#546e7a";
          return `<span style="background-color:${bg}; color:#fff; padding:2px 8px; border-radius:10px; font-size:11px; white-space:nowrap;">${state}</span>`;
        };

        const renderTableRows = (items) => {
            tbody.innerHTML = "";
            items.forEach(item => {
                const row = document.createElement("tr");

                if (item.is_group) {
                    row.style.cursor = "pointer";
                    row.title = "Click to drill down";
                }

                // Sectors are ranked 1-N now (not rated 0-100), so skip the color tint that
                // assumes a 50-midpoint scale and would render every sector as false-red.
                const isSectorView = state.currentStatsGroup === "sectors" && !state.currentMegaTheme && !state.currentSubTheme;
                const rsCellStyle = isSectorView
                    ? "color: #e6edf3; font-weight: bold;"
                    : `background-color: ${getColorblindTint((item.rs_score - 50) / 50)}; color: #ffffff; font-weight: bold;`;
                const rsCellText = item.rs_score !== undefined && item.rs_score !== null
                    ? (isSectorView ? `#${item.rs_score}` : item.rs_score)
                    : '...';

                // Composite RS is rendered right after description
                row.innerHTML = `
                    <td><strong>${item.symbol}</strong></td>
                    <td style="text-align: left;">${item.description}</td>
                    <td style="text-align: center;">${getStateBadge(item.state)}</td>
                    <td style="${rsCellStyle}">
                        ${rsCellText}
                    </td>
                    <td style="background-color: ${getColorblindTint(item.rs_roc_8w / 100)}; color: #e6edf3;">
                        ${item.rs_roc_8w !== undefined && item.rs_roc_8w !== null ? (item.rs_roc_8w > 0 ? `+${item.rs_roc_8w.toFixed(2)}pp` : `${item.rs_roc_8w.toFixed(2)}pp`) : '...'}
                    </td>
                    <td title="${(item.top_tickers && item.top_tickers.length) ? item.top_tickers.map(t => `${t.symbol} ${t.score}`).join(', ') : ''}">
                        ${item.breadth_total ? `${item.breadth_count}/${item.breadth_total}` : '—'}
                    </td>
                    <td>$${item.price.toFixed(2)}</td>
                    <td style="background-color: ${getColorblindTint(item.chg_1d)}; color: #e6edf3;">${(item.chg_1d * 100).toFixed(2)}%</td>
                    <td style="background-color: ${getColorblindTint(item.chg_1w)}; color: #e6edf3;">${(item.chg_1w * 100).toFixed(2)}%</td>
                    <td style="background-color: ${getColorblindTint(item.chg_2w)}; color: #e6edf3;">${(item.chg_2w * 100).toFixed(2)}%</td>
                    <td style="background-color: ${getColorblindTint(item.chg_1m)}; color: #e6edf3;">${(item.chg_1m * 100).toFixed(2)}%</td>
                    <td style="background-color: ${getColorblindTint(item.chg_2m)}; color: #e6edf3;">${(item.chg_2m * 100).toFixed(2)}%</td>
                    <td style="background-color: ${getColorblindTint(item.chg_3m)}; color: #e6edf3;">${(item.chg_3m * 100).toFixed(2)}%</td>
                    <td style="background-color: ${getColorblindTint(item.chg_6m)}; color: #e6edf3;">${(item.chg_6m * 100).toFixed(2)}%</td>
                    <td style="background-color: ${getColorblindTint(item.chg_1y)}; color: #e6edf3;">${(item.chg_1y * 100).toFixed(2)}%</td>
                    <td>${item.volume.toLocaleString()}</td>
                    <td>${item.avg_vol.toLocaleString()}</td>
                    <td style="background-color: ${getColorblindTint(item.vs_20sma)}; color: #e6edf3;">${(item.vs_20sma * 100).toFixed(2)}%</td>
                    <td style="background-color: ${getColorblindTint(item.vs_50sma)}; color: #e6edf3;">${(item.vs_50sma * 100).toFixed(2)}%</td>
                    <td style="background-color: ${getColorblindTint(item.vs_200sma)}; color: #e6edf3;">${(item.vs_200sma * 100).toFixed(2)}%</td>
                `;

                if (item.is_group) {
                    row.onclick = async () => {
                        if (item.next_level !== "mega" && item.next_level !== "sub") return;

                        pushNavHistory();

                        if (item.next_level === "mega") {
                            state.currentMegaTheme = item.symbol;
                            state.currentSubTheme = null;
                        } else {
                            state.currentSubTheme = item.symbol;
                            state.currentMegaTheme = null;
                        }
                        await loadData();
                    };
                }

                tbody.appendChild(row);
            });
        };

        renderTableRows(data);

        // Header sorting logic matching the exact column keys
        const headers = document.querySelectorAll("#statsTable th");
        headers.forEach((th, index) => {
            th.style.cursor = "pointer";
            th.title = "Click to sort";
            th.onclick = () => {
                if (currentSortCol === index) {
                    sortAscending = !sortAscending;
                } else {
                    currentSortCol = index;
                    sortAscending = false;
                }

                const keys = [
                    "symbol", "description", "state", "rs_score", "rs_roc_8w", "breadth_count", "price",
                    "chg_1d", "chg_1w", "chg_2w", "chg_1m", "chg_2m", "chg_3m", "chg_6m", "chg_1y",
                    "volume", "avg_vol", "vs_20sma", "vs_50sma", "vs_200sma"
                ];
                const key = keys[index];

                // "State" has a meaningful order (RRG-style leadership ranking), not
                // alphabetical - sort by rank instead of localeCompare for this column.
                const STATE_RANK = {
                    "Leading": 5, "Emerging": 4, "Fading": 3,
                    "Neutral": 2, "Lagging": 1, "Breaking Down": 0
                };

                data.sort((a, b) => {
                    let valA = a[key];
                    let valB = b[key];

                    if (key === "state") {
                        const rankA = STATE_RANK[valA] ?? -1;
                        const rankB = STATE_RANK[valB] ?? -1;
                        return sortAscending ? rankA - rankB : rankB - rankA;
                    }

                    if (key === "rs_score" && state.currentStatsGroup === "sectors" && !state.currentMegaTheme && !state.currentSubTheme) {
                        // Sectors' rs_score is a rank (1 = best), the opposite of a 0-100
                        // score where higher = better - invert so "descending" still means
                        // "best first" like every other column.
                        if (valA === null || valA === undefined) return 1;
                        if (valB === null || valB === undefined) return -1;
                        return sortAscending ? valB - valA : valA - valB;
                    }

                    if (typeof valA === "string") {
                        return sortAscending ? valA.localeCompare(valB) : valB.localeCompare(valA);
                    }
                    // Treat missing values (e.g. rs_roc_8w with <1.5y of history) as
                    // always sorting to the bottom, regardless of sort direction.
                    if (valA === null || valA === undefined) return 1;
                    if (valB === null || valB === undefined) return -1;
                    return sortAscending ? valA - valB : valB - valA;
                });

                renderTableRows(data);
            };
        });

    } catch (err) {
        console.error("Failed to load dashboard stats table:", err);
    }
}

/* ── Main Load Function ── */
async function loadData() {
  try {
    // Fetch stats table first so we have per-symbol state (Leading/Emerging/etc) available
    // to filter the RRG chart's tails by, if a state filter is active.
    await loadStatsTable();

    const rrgData = await fetchRRGData();
    state.allSectors = rrgData.sectors;
    state.sectors = applyStateFilterToSectors(state.allSectors, state.lastFullStatsData);

    renderChart();
    renderSidebar();

    if (state.viewMode === "charts") {
      loadChartGrid();
    }
  } catch (err) {
    console.error("Failed to load data:", err);
  }
}

document.addEventListener("DOMContentLoaded", async () => {
    try {
        state.holdings = await fetchHoldings();
    } catch (e) {
        console.error("Failed to load holdings:", e);
    }
    
    try {
        state.themeStructure = await fetchThemeStructure();
        renderThemePanel();
    } catch (e) {
        console.error("Failed to load theme structure:", e);
    }

    await loadData();
    await loadThemeRankingsForSidebar();
});