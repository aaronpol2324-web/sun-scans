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
  // Only use explicitly added custom symbols or default to sectors
  const extra = state.customSymbols.filter(s => s !== state.benchmark).join(",");
  const url = `${API_BASE}/api/rrg?benchmark=${state.benchmark}&tail=${state.tailLength}${extra ? "&extra=" + extra : ""}`;
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
function renderChart() {
  const container = document.getElementById("chartContainer");
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
  state.currentStatsGroup = groupVal;

  try {
    const res = await fetch(`${API_BASE}/api/group-stats?group=${groupVal}`);
    if (res.ok) {
      const data = await res.json();
      state.customSymbols = data.map(item => item.symbol).filter(sym => sym !== state.benchmark);
      await loadData();
    }
  } catch (err) {
    console.error("Failed to update stats group:", err);
  }
}

/* ── Load Dashboard Stats Table ── */
/* ── Load Dashboard Stats Table with Drill-Down Support ── */
/* ── Load Dashboard Stats Table with Drill-Down Support ── */
async function loadStatsTable(urlParams = `group=${state.currentStatsGroup}`) {
    try {
        const response = await fetch(`${API_BASE}/api/group-stats?${urlParams}`);
        if (!response.ok) return;
        const data = await response.json();

        const tbody = document.querySelector("#statsTable tbody");
        if (!tbody) return;

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

        const renderTableRows = (items) => {
            tbody.innerHTML = "";
            items.forEach(item => {
                const row = document.createElement("tr");

                if (item.is_group) {
                    row.style.cursor = "pointer";
                    row.title = "Click to drill down";
                }

                // Composite RS is rendered right after description
                row.innerHTML = `
                    <td><strong>${item.symbol}</strong></td>
                    <td style="text-align: left;">${item.description}</td>
                    <td style="background-color: ${getColorblindTint(item.composite_rs / 100)}; color: #ffffff; font-weight: bold;">
                        ${item.composite_rs !== undefined && item.composite_rs !== null ? (item.composite_rs > 0 ? `+${item.composite_rs.toFixed(2)}%` : `${item.composite_rs.toFixed(2)}%`) : '...'}
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
                        let drillParam = "";
                        if (item.next_level === "mega") {
                            drillParam = `mega_theme=${encodeURIComponent(item.symbol)}`;
                        } else if (item.next_level === "sub") {
                            drillParam = `sub_theme=${encodeURIComponent(item.symbol)}`;
                        }

                        if (drillParam) {
                            state.customSymbols = [];
                            await loadStatsTable(drillParam);

                            const visibleTickers = data.map(d => d.symbol).filter(sym => sym !== state.benchmark);
                            state.customSymbols = visibleTickers;
                            const rrgData = await fetchRRGData();
                            state.sectors = rrgData.sectors;
                            renderChart();
                        }
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
                    "symbol", "description", "composite_rs", "price",
                    "chg_1d", "chg_1w", "chg_2w", "chg_1m", "chg_2m", "chg_3m", "chg_6m", "chg_1y",
                    "volume", "avg_vol", "vs_20sma", "vs_50sma", "vs_200sma"
                ];
                const key = keys[index];

                data.sort((a, b) => {
                    let valA = a[key];
                    let valB = b[key];
                    if (typeof valA === "string") {
                        return sortAscending ? valA.localeCompare(valB) : valB.localeCompare(valA);
                    }
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
    const data = await fetchRRGData();
    state.sectors = data.sectors;
    renderChart();
    renderSidebar();
    loadStatsTable();
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