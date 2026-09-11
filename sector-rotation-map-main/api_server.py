import os
import json
import time
import hashlib
from datetime import datetime, timedelta
import pandas as pd
import yfinance as yf
from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import logging
from fastapi.responses import Response
import os
import time

# The name of the file that will live on your hard drive
LOCAL_CACHE_FILE = "market_data_cache.json"


def get_local_cache(symbols):
    """Loads whatever historical price data we already have saved on disk."""
    if os.path.exists(LOCAL_CACHE_FILE):
        if time.time() - os.path.getmtime(LOCAL_CACHE_FILE) < 86400:
            try:
                df = pd.read_json(LOCAL_CACHE_FILE, orient="split")
                if not isinstance(df.index, pd.DatetimeIndex):
                    df.index = pd.to_datetime(df.index)
                return df
            except Exception:
                pass
    return None


def save_local_cache(new_df):
    """Merges newly downloaded stocks into our persistent local cache file."""
    try:
        if os.path.exists(LOCAL_CACHE_FILE):
            existing_df = pd.read_json(LOCAL_CACHE_FILE, orient="split")
            if not isinstance(existing_df.index, pd.DatetimeIndex):
                existing_df.index = pd.to_datetime(existing_df.index)
            # Combine old and new stocks together
            combined_df = pd.concat([existing_df, new_df], axis=1)
            # Remove duplicate columns if a ticker was updated
            combined_df = combined_df.loc[:, ~combined_df.columns.duplicated()]
            combined_df.to_json(LOCAL_CACHE_FILE, orient="split", date_format="iso")
            return
    except Exception:
        pass

    # Fallback: if no cache exists yet, save the new one directly
    new_df.to_json(LOCAL_CACHE_FILE, orient="split", date_format="iso")


def save_local_cache(df):
    """Saves the DataFrame directly using Pandas built-in JSON handler."""
    print("--> Saving fresh data to local computer cache...")
    try:
        df.to_json(LOCAL_CACHE_FILE, orient="split", date_format="iso")
    except Exception as e:
        print(f"Could not save local cache: {e}")

# Suppress yfinance internal warning prints completely
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

app = FastAPI()
@app.get("/favicon.ico", include_in_schema=False)
async def favicon():
    return Response(status_code=204)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

SECTOR_ETFS = {
    "XLK": {"name": "Technology", "color": "#00bcd4"},
    "XLF": {"name": "Financials", "color": "#3f51b5"},
    "XLV": {"name": "Health Care", "color": "#4caf50"},
    "XLE": {"name": "Energy", "color": "#ff9800"},
    "XLP": {"name": "Consumer Staples", "color": "#e91e63"},
    "XLB": {"name": "Materials", "color": "#9c27b0"},
    "XLRE": {"name": "Real Estate", "color": "#009688"},
    "XLI": {"name": "Industrials", "color": "#795548"},
    "XLU": {"name": "Utilities", "color": "#607d8b"},
    "XLY": {"name": "Discretionary", "color": "#cddc39"},
    "XLC": {"name": "Comm Services", "color": "#ffeb3b"},
}

import time
import hashlib
import requests
import pandas as pd
from datetime import datetime, timedelta

_price_cache = {}
CACHE_TTL = 300

def _cache_key(symbols, start_date, end_date, timeframe):
    raw = f"{sorted(symbols)}:{start_date}:{end_date}:{timeframe}"
    return hashlib.md5(raw.encode()).hexdigest()


def fetch_prices(symbols, start_date=None, end_date=None, timeframe="1D"):
    if end_date is None:
        end_date = (datetime.now() + timedelta(days=1)).strftime('%Y-%m-%d')
    if start_date is None:
        start_date = (datetime.now() - timedelta(days=730)).strftime('%Y-%m-%d')

    symbols = list(symbols)

    # 1. Load what we already have from disk cache
    saved_dataframe = get_local_cache(symbols)
    missing_symbols = []

    if saved_dataframe is not None:
        missing_symbols = [s for s in symbols if s not in saved_dataframe.columns]
        if not missing_symbols:
            print("--> Loading 100% from local computer cache (Instant)")
            return saved_dataframe
    else:
        missing_symbols = symbols

    # If everything is already cached, we wouldn't reach here.
    # Now we only fetch data for the actual missing symbols instead of the whole portfolio!
    symbols_to_download = missing_symbols if saved_dataframe is not None else symbols

    def _extract_close(df, sym):
        if df is None or df.empty:
            return None

        close_col = None
        if isinstance(df.columns, pd.MultiIndex):
            if ('Adj Close', sym) in df.columns:
                close_col = df[('Adj Close', sym)]
            elif ('Close', sym) in df.columns:
                close_col = df[('Close', sym)]
            else:
                return None
        else:
            if 'Adj Close' in df.columns:
                close_col = df['Adj Close']
            elif 'Close' in df.columns:
                close_col = df['Close']
            else:
                return None

        if not isinstance(close_col.index, pd.DatetimeIndex):
            close_col.index = pd.to_datetime(close_col.index)
        if close_col.index.tz is not None:
            close_col.index = close_col.index.tz_localize(None)

        return close_col if not close_col.dropna().empty else None

    data_dict = {}
    CHUNK_SIZE = 150

    for i in range(0, len(symbols_to_download), CHUNK_SIZE):
        chunk = symbols_to_download[i:i + CHUNK_SIZE]
        try:
            downloaded_df = yf.download(chunk, start=start_date, end=end_date, progress=False, auto_adjust=True)
        except Exception:
            continue

        if downloaded_df is None or downloaded_df.empty:
            continue

        if len(chunk) == 1:
            close_col = _extract_close(downloaded_df, chunk[0])
            if close_col is not None:
                data_dict[chunk[0]] = close_col
        else:
            for sym in chunk:
                close_col = _extract_close(downloaded_df, sym)
                if close_col is not None:
                    data_dict[sym] = close_col

    new_resampled = None
    if data_dict:
        all_close = pd.DataFrame(data_dict)
        if timeframe == "1D":
            new_resampled = all_close.dropna(how="all")
        elif timeframe == "1W":
            new_resampled = all_close.resample("W-FRI").last().dropna(how="all")
        elif timeframe == "1M":
            new_resampled = all_close.resample("ME").last().dropna(how="all")
        else:
            new_resampled = all_close.resample("1D").last().dropna(how="all")

    # 2. Merge newly fetched missing symbols with existing disk cache
    if saved_dataframe is not None:
        if new_resampled is not None and not new_resampled.empty:
            combined_df = pd.concat([saved_dataframe, new_resampled], axis=1)
            combined_df = combined_df.loc[:, ~combined_df.columns.duplicated()]
            save_local_cache(combined_df)
            return combined_df
        return saved_dataframe
    else:
        if new_resampled is not None:
            save_local_cache(new_resampled)
            return new_resampled
        else:
            raise HTTPException(status_code=404, detail="No data returned for given symbols")

def load_taxonomy():
    """Load the complete mega-theme, theme, and ticker taxonomy from the single JSON structure."""
    json_path = os.path.join(os.path.dirname(__file__), "mega_theme_structure.json")
    with open(json_path, "r") as f:
        return json.load(f)

@app.get("/api/theme-structure")
def get_theme_structure():
    return load_taxonomy()


@app.get("/api/theme-rankings")
def get_theme_rankings(benchmark: str = Query("SPY")):
    benchmark = benchmark.upper().strip()
    structure = load_taxonomy()  # <-- Updated from load_excel_taxonomy() to load_taxonomy()

    all_tickers = {benchmark}
    subtheme_map = []

    for mega_name, sub_dict in structure.items():
        for sub_name, tickers in sub_dict.items():
            if tickers:
                clean_tickers = [t.upper().strip() for t in tickers if t]
                if clean_tickers:
                    subtheme_map.append({
                        "mega": mega_name,
                        "sub": sub_name,
                        "tickers": clean_tickers
                    })
                    all_tickers.update(clean_tickers)

    prices_df = fetch_prices(list(all_tickers), timeframe="1D")
    if benchmark not in prices_df.columns:
        raise HTTPException(status_code=404, detail=f"Benchmark {benchmark} not found")

    bench_closes = prices_df[benchmark]
    rankings = []

    for item in subtheme_map:
        valid_tickers = [t for t in item["tickers"] if t in prices_df.columns]
        if not valid_tickers:
            continue

        sub_closes = prices_df[valid_tickers].mean(axis=1).dropna()
        combined = pd.concat([sub_closes, bench_closes], axis=1).dropna()
        if len(combined) < 50:
            continue

        sub_c = combined.iloc[:, 0]
        bench_c = combined.iloc[:, 1]

        def get_rs_return(bars):
            if len(combined) > bars:
                sub_perf = (sub_c.iloc[-1] / sub_c.iloc[-1 - bars]) - 1
                bench_perf = (bench_c.iloc[-1] / bench_c.iloc[-1 - bars]) - 1
                return float(sub_perf - bench_perf)
            return 0.0

        rs_1m = get_rs_return(21)
        rs_3m = get_rs_return(63)
        rs_6m = get_rs_return(126)
        rs_1y = get_rs_return(252)
        composite_rs = (rs_1m * 0.1) + (rs_3m * 0.2) + (rs_6m * 0.3) + (rs_1y * 0.4)

        high_52w = sub_c.tail(min(252, len(sub_c))).max()
        current_price = sub_c.iloc[-1]
        dist_from_high = float((current_price / high_52w) - 1) if high_52w > 0 else 0.0

        rankings.append({
            "mega_theme": item["mega"],
            "sub_theme": item["sub"],
            "tickers": valid_tickers,
            "current_price": round(float(current_price), 2),
            "rs_1m": round(rs_1m * 100, 2),
            "rs_3m": round(rs_3m * 100, 2),
            "rs_6m": round(rs_6m * 100, 2),
            "rs_1y": round(rs_1y * 100, 2),
            "composite_rs": round(composite_rs * 100, 2),
            "dist_from_high": round(dist_from_high * 100, 2),
        })

    rankings.sort(key=lambda x: x["composite_rs"], reverse=True)
    return {"benchmark": benchmark, "rankings": rankings}


@app.get("/api/dashboard-stats")
def get_dashboard_stats():
    """Restricted strictly to SPY and core sector ETFs for the bottom stats table."""
    tracked_tickers = ["SPY"] + list(SECTOR_ETFS.keys())

    prices_df = fetch_prices(tracked_tickers, timeframe="1D")

    volumes_df = pd.DataFrame()
    try:
        raw_vol = yf.download(tracked_tickers, period="3y", progress=False)
        if isinstance(raw_vol.columns, pd.MultiIndex) and "Volume" in raw_vol.columns.levels[0]:
            volumes_df = raw_vol["Volume"]
        elif "Volume" in raw_vol.columns:
            volumes_df = raw_vol[["Volume"]]
    except Exception:
        pass

    stats_list = []

    for symbol in tracked_tickers:
        if symbol not in prices_df.columns:
            continue
        series = prices_df[symbol].dropna()

        if len(series) < 20:
            continue

        current_price = float(series.iloc[-1])
        vol = 0
        avg_vol = 0
        if not volumes_df.empty:
            if symbol in volumes_df.columns:
                v_ser = volumes_df[symbol].dropna()
                if not v_ser.empty:
                    vol = int(v_ser.iloc[-1])
                    avg_vol = int(v_ser.tail(20).mean())

        def safe_chg(bars):
            """
            Pure trading-session lookback: compares the current close to the close
            exactly `bars` trading sessions ago. Matches standard platform convention
            and is consistent with the other stats endpoints in this app.
            """
            if series is None or len(series) <= bars:
                return 0.0

            base_price = series.iloc[-1 - bars]
            curr_price = series.iloc[-1]

            if not base_price:
                return 0.0

            return float((curr_price / base_price) - 1.0)

        sma_20 = series.tail(20).mean() if len(series) >= 20 else current_price
        sma_50 = series.tail(50).mean() if len(series) >= 50 else current_price
        sma_200 = series.tail(200).mean() if len(series) >= 200 else current_price

        if symbol == "SPY":
            desc = "SPDR S&P 500 ETF Trust"
        elif symbol in SECTOR_ETFS:
            desc = f"{SECTOR_ETFS.get(symbol, {}).get('name', symbol)} Sector ETF"
        else:
            desc = f"Equity / Asset: {symbol}"

        stats_list.append({
            "symbol": symbol,
            "description": SECTOR_ETFS.get(symbol, {}).get("name", "Benchmark"),
            "price": current_price,
            "chg_1d": round(safe_chg(1), 4),
            "chg_1w": round(safe_chg(5), 4),
            "chg_2w": round(safe_chg(10), 4),
            "chg_1m": round(safe_chg(21), 4),
            "chg_2m": round(safe_chg(42), 4),
            "chg_3m": round(safe_chg(63), 4),
            "chg_6m": round(safe_chg(126), 4),
            "chg_1y": round(safe_chg(252), 4),
            "volume": vol,
            "avg_volume": avg_vol,
            # ... other indicators ...
        })

    return {"stats": stats_list}


_STATS_CACHE = []
_CACHE_LOADED = False


@app.get("/api/theme-ticker-stats")
def get_theme_ticker_stats():
    """Cached stats endpoint exclusively for sidebar theme tickers."""
    global _STATS_CACHE, _CACHE_LOADED
    if _CACHE_LOADED and _STATS_CACHE:
        return _STATS_CACHE

    try:
        structure = load_taxonomy()
        theme_tickers = set()
        if structure and isinstance(structure, dict):
            for mega_name, sub_dict in structure.items():
                if isinstance(sub_dict, dict):
                    for sub_name, tickers in sub_dict.items():
                        if isinstance(tickers, list):
                            for t in tickers:
                                if t and isinstance(t, str):
                                    theme_tickers.add(t.upper().strip())

        tracked_tickers = list(theme_tickers)
        if not tracked_tickers:
            return []

        # Uses your existing fetch_prices helper
        prices_df = fetch_prices(tracked_tickers, timeframe="1D")
        stats_list = []

        for symbol in tracked_tickers:
            if symbol not in prices_df.columns:
                continue
            series = prices_df[symbol].dropna()
            if len(series) < 5:
                continue

            current_price = float(series.iloc[-1])

            def safe_chg(bars):
                if series is None or len(series) < 2:
                    return 0.0

                # For 1W (bars=5), anchor strictly to the previous Friday's close
                if bars == 5:
                    # Find all Friday bars in the index
                    fridays = series[series.index.weekday == 4]
                    if not fridays.empty:
                        # Get the last Friday that is strictly before the current week/day
                        current_date = series.index[-1]
                        past_fridays = fridays[fridays.index < current_date]
                        if not past_fridays.empty:
                            base_price = past_fridays.iloc[-1]
                            return float((series.iloc[-1] / base_price) - 1.0)

                # Fallback to standard session lookback for other timeframes
                if len(series) > bars:
                    return float((series.iloc[-1] / series.iloc[-1 - bars]) - 1.0)
                elif len(series) > 1:
                    return float((series.iloc[-1] / series.iloc[0]) - 1.0)
                return 0.0
    except Exception as e:
        print(f"Error in theme-ticker-stats: {e}")
        return []

def compute_single_rrg(sector_prices, benchmark_prices, tail_length=8):
    """Calculates RRG ratio and momentum with strict date alignment."""
    # 1. Align the theme prices and benchmark prices by date first
    combined = pd.concat([sector_prices, benchmark_prices], axis=1).dropna()
    if len(combined) < 52:
        return None

    sec_c = combined.iloc[:, 0]
    bench_c = combined.iloc[:, 1]

    # 2. Run the math on cleanly aligned series
    raw_rs = (sec_c / bench_c) * 100
    rs_smoothed = raw_rs.ewm(span=10, adjust=False).mean()

    rolling_mean = rs_smoothed.rolling(window=52, min_periods=20).mean()
    rolling_std = rs_smoothed.rolling(window=52, min_periods=20).std()

    rolling_std = rolling_std.replace(0, float('nan'))
    rs_ratio = 100 + ((rs_smoothed - rolling_mean) / rolling_std) * 2

    rs_momentum_raw = rs_ratio - rs_ratio.shift(1)
    mom_smoothed = rs_momentum_raw.ewm(span=5, adjust=False).mean()
    mom_mean = mom_smoothed.rolling(window=52, min_periods=20).mean()
    mom_std = mom_smoothed.rolling(window=52, min_periods=20).std()
    mom_std = mom_std.replace(0, float('nan'))
    rs_momentum = 100 + ((mom_smoothed - mom_mean) / mom_std) * 2

    valid = rs_ratio.notna() & rs_momentum.notna()
    rs_r = rs_ratio[valid]
    rs_m = rs_momentum[valid]

    if len(rs_r) == 0:
        return None

    n = min(tail_length, len(rs_r))
    tail = []
    for i in range(-n, 0):
        tail.append({
            "rs_ratio": round(float(rs_r.iloc[i]), 2),
            "rs_momentum": round(float(rs_m.iloc[i]), 2),
        })

    current = tail[-1] if tail else None
    if current:
        r, m = current["rs_ratio"], current["rs_momentum"]
        if r >= 100 and m >= 100:
            quadrant = "Leading"
        elif r >= 100 and m < 100:
            quadrant = "Weakening"
        elif r < 100 and m >= 100:
            quadrant = "Improving"
        else:
            quadrant = "Lagging"
    else:
        quadrant = "Unknown"

    return {
        "tail": tail,
        "quadrant": quadrant,
        "rs_ratio": current["rs_ratio"] if current else 100.0,
        "rs_momentum": current["rs_momentum"] if current else 100.0
    }


@app.get("/api/rrg")
def get_rrg(
    benchmark: str = "SPY",
    tail: int = Query(8, ge=4, le=60),
    timeframe: str = "1W",
    group: str = Query("sectors"),
    mega_theme: str = Query(None),
    sub_theme: str = Query(None)
):
    try:
        benchmark = benchmark.upper().strip()
        timeframe = timeframe.upper().strip()
        group = group.lower().strip()
        structure = load_taxonomy() or {}

        tickers_to_fetch = [benchmark]
        target_mapping = []

        # 1. Resolve symbols dynamically based on table context (Sectors, Mega, or Sub-themes)
        if sub_theme:
            sub_theme = sub_theme.strip()
            for m_name, sub_dict in structure.items():
                if sub_theme in sub_dict:
                    tickers = sub_dict[sub_theme]
                    for t in tickers:
                        if t:
                            clean_t = t.upper().strip()
                            target_mapping.append({"symbol": clean_t, "description": f"Component of {sub_theme}"})
                            tickers_to_fetch.append(clean_t)
        elif mega_theme:
            mega_theme = mega_theme.strip()
            if mega_theme in structure:
                sub_dict = structure[mega_theme]
                for s_name, t_list in sub_dict.items():
                    valid_t = [t.upper().strip() for t in t_list if t]
                    if valid_t:
                        # Represent sub-themes as group nodes on the RRG chart
                        target_mapping.append({"symbol": s_name, "description": f"Sub-Theme under {mega_theme}", "tickers": valid_t})
                        tickers_to_fetch.extend(valid_t)
        elif group == "sectors":
            for sym, info in SECTOR_ETFS.items():
                target_mapping.append({"symbol": sym, "description": info["name"], "color": info["color"], "tickers": [sym]})
                tickers_to_fetch.append(sym)
        elif group == "mega_themes":
            for m_name, sub_dict in structure.items():
                all_m_tickers = []
                for s_name, t_list in sub_dict.items():
                    if t_list:
                        all_m_tickers.extend([t.upper().strip() for t in t_list if t])
                if all_m_tickers:
                    target_mapping.append({"symbol": m_name, "description": "Mega-Theme Group", "tickers": list(set(all_m_tickers))})
                    tickers_to_fetch.extend(all_m_tickers)
        elif group == "sub_themes":
            for m_name, sub_dict in structure.items():
                for s_name, t_list in sub_dict.items():
                    valid_t = [t.upper().strip() for t in t_list if t]
                    if valid_t:
                        target_mapping.append({"symbol": s_name, "description": f"Sub-Theme ({m_name})", "tickers": valid_t})
                        tickers_to_fetch.extend(valid_t)

        # 2. Fetch prices for the resolved universe using your local cache
        prices_df = fetch_prices(list(set(tickers_to_fetch)), timeframe=timeframe)
        if prices_df is None or benchmark not in prices_df.columns:
            return {"benchmark": benchmark, "sectors": {}}

        bench_closes = prices_df[benchmark]
        sectors_data = {}

        # 3. Compute RRG tails matching the current table view
        for item in target_mapping:
            sym_key = item["symbol"]
            if sym_key == benchmark:
                continue

            if "tickers" in item:
                comp_ts = [t for t in item["tickers"] if t in prices_df.columns]
                if not comp_ts:
                    continue
                series_to_use = prices_df[comp_ts].mean(axis=1).dropna() if len(comp_ts) > 1 else prices_df[comp_ts[0]].dropna()
            else:
                if sym_key not in prices_df.columns:
                    continue
                series_to_use = prices_df[sym_key].dropna()

            if len(series_to_use) < 5:
                continue

            rrg_res = compute_single_rrg(series_to_use, bench_closes, tail_length=tail)
            if rrg_res:
                color = SECTOR_ETFS.get(sym_key, {}).get("color", "#00bcd4")
                sectors_data[sym_key] = {
                    "symbol": sym_key,
                    "description": item["description"],
                    "color": color,
                    "quadrant": rrg_res["quadrant"],
                    "rs_ratio": rrg_res["rs_ratio"],
                    "rs_momentum": rrg_res["rs_momentum"],
                    "tail": rrg_res["tail"]
                }

        return {"benchmark": benchmark, "sectors": sectors_data}

    except Exception as e:
        print(f"CRITICAL ERROR in get_rrg: {e}")
        return {"benchmark": benchmark, "sectors": {}}



@app.get("/api/group-stats")
@app.get("/api/group-stats")
def get_group_stats(
        group: str = Query("sectors"),
        mega_theme: str = Query(None),
        sub_theme: str = Query(None)
):
    group = group.lower().strip()
    structure = load_taxonomy()

    tickers_to_fetch = ["SPY"]  # Always ensure SPY is present for benchmark comparison
    target_mapping = []
    group_mode = "sectors"

    if sub_theme:
        sub_theme = sub_theme.strip()
        for m_name, sub_dict in structure.items():
            if sub_theme in sub_dict:
                tickers = sub_dict[sub_theme]
                for t in tickers:
                    if t:
                        clean_t = t.upper().strip()
                        target_mapping.append(
                            {"symbol": clean_t, "description": f"Component of {sub_theme}", "is_group": False})
                        tickers_to_fetch.append(clean_t)
        group_mode = "tickers"

    elif mega_theme:
        mega_theme = mega_theme.strip()
        if mega_theme in structure:
            sub_dict = structure[mega_theme]
            for s_name, t_list in sub_dict.items():
                valid_t = [t.upper().strip() for t in t_list if t]
                if valid_t:
                    target_mapping.append({
                        "symbol": s_name,
                        "description": f"Sub-Theme under {mega_theme}",
                        "tickers": valid_t,
                        "is_group": True,
                        "next_level": "sub"
                    })
                    tickers_to_fetch.extend(valid_t)
        group_mode = "sub_themes_filtered"

    elif group == "sectors":
        tickers_to_fetch.extend(list(SECTOR_ETFS.keys()))
        group_mode = "sectors"

    elif group == "mega_themes":
        for m_name, sub_dict in structure.items():
            all_m_tickers = []
            for s_name, t_list in sub_dict.items():
                if t_list:
                    all_m_tickers.extend([t.upper().strip() for t in t_list if t])
            if all_m_tickers:
                target_mapping.append({
                    "symbol": m_name,
                    "description": "Mega-Theme Group",
                    "tickers": list(set(all_m_tickers)),
                    "is_group": True,
                    "next_level": "mega"
                })
                tickers_to_fetch.extend(all_m_tickers)
        group_mode = "mega_themes"

    elif group == "sub_themes":
        for m_name, sub_dict in structure.items():
            for s_name, t_list in sub_dict.items():
                valid_t = [t.upper().strip() for t in t_list if t]
                if valid_t:
                    target_mapping.append({
                        "symbol": s_name,
                        "description": f"Sub-Theme ({m_name})",
                        "tickers": valid_t,
                        "is_group": True,
                        "next_level": "sub"
                    })
                    tickers_to_fetch.extend(valid_t)
        group_mode = "sub_themes"

    prices_df = fetch_prices(list(set(tickers_to_fetch)), timeframe="1D") if tickers_to_fetch else pd.DataFrame()
    bench_series = prices_df["SPY"].dropna() if "SPY" in prices_df.columns else None

    stats_list = []

    def calculate_metrics(series, symbol, desc, is_grp=False, n_lvl="none"):
        current_price = float(series.iloc[-1])

        def get_rs(bars):
            if bench_series is not None and len(series) > bars and len(bench_series) > bars:
                try:
                    s_perf = (series.iloc[-1] / series.iloc[-1 - bars]) - 1
                    b_perf = (bench_series.iloc[-1] / bench_series.iloc[-1 - bars]) - 1
                    return float(s_perf - b_perf)
                except Exception:
                    return 0.0
            return 0.0

        def safe_chg(bars):
            # Keep fallback for short-term bar-based lookups if needed,
            # or handle months/years via calendar offset if bars >= 126
            if series is None or len(series) < 2:
                return 0.0

            end_date = series.index[-1]

            # Map long-term bar checks to calendar offsets
            if bars == 126:  # 6 Months
                target_date = end_date - pd.DateOffset(months=6)
            elif bars == 252:  # 1 Year
                target_date = end_date - pd.DateOffset(years=1)
            else:
                # Standard bar-based check for short terms (1D, 1W, etc.)
                if len(series) <= bars:
                    return 0.0
                past_val = series.iloc[-1 - bars]
                curr_val = series.iloc[-1]
                if past_val == 0 or pd.isna(past_val) or pd.isna(curr_val):
                    return 0.0
                return (curr_val - past_val) / past_val

            # Calendar offset lookup for 6M and 1Y
            valid_subset = series[series.index <= target_date]
            past_val = valid_subset.iloc[-1] if not valid_subset.empty else series.iloc[0]
            curr_val = series.iloc[-1]

            if past_val == 0 or pd.isna(past_val) or pd.isna(curr_val):
                return 0.0

            return (curr_val - past_val) / past_val

        rs_1m = get_rs(21)
        rs_3m = get_rs(63)
        rs_6m = get_rs(126)
        rs_1y = get_rs(252)
        composite_rs = (rs_1m * 0.1) + (rs_3m * 0.2) + (rs_6m * 0.3) + (rs_1y * 0.4)

        sma_20 = series.tail(20).mean() if len(series) >= 20 else current_price
        sma_50 = series.tail(50).mean() if len(series) >= 50 else current_price
        sma_200 = series.tail(200).mean() if len(series) >= 200 else current_price

        return {
            "symbol": symbol,
            "description": desc,
            "price": round(current_price, 2),
            "chg_1d": round(safe_chg(1), 4),
            "chg_1w": round(safe_chg(5), 4),
            "chg_2w": round(safe_chg(10), 4),
            "chg_1m": round(safe_chg(21), 4),
            "chg_2m": round(safe_chg(42), 4),
            "chg_3m": round(safe_chg(63), 4),
            "chg_6m": round(safe_chg(126), 4),
            "chg_1y": round(safe_chg(252), 4),
            "rs_1m": round(rs_1m * 100, 2),
            "rs_3m": round(rs_3m * 100, 2),
            "rs_6m": round(rs_6m * 100, 2),
            "rs_1y": round(rs_1y * 100, 2),
            "composite_rs": round(composite_rs * 100, 2),
            "volume": 0,
            "avg_vol": 0,
            "vs_20sma": round(float((current_price / sma_20) - 1), 4) if sma_20 else 0.0,
            "vs_50sma": round(float((current_price / sma_50) - 1), 4) if sma_50 else 0.0,
            "vs_200sma": round(float((current_price / sma_200) - 1), 4) if sma_200 else 0.0,
            "is_group": is_grp,
            "next_level": n_lvl
        }

    if group_mode in ["sectors", "tickers"]:
        symbols_to_loop = [s for s in (tickers_to_fetch if group_mode == "sectors" else [m["symbol"] for m in target_mapping]) if s != "SPY"]
        for sym in symbols_to_loop:
            if sym not in prices_df.columns:
                continue
            series = prices_df[sym].dropna()
            if len(series) < 5:
                continue
            desc = SECTOR_ETFS.get(sym, {}).get("name", f"Asset: {sym}") if group_mode == "sectors" else f"Ticker: {sym}"
            stats_list.append(calculate_metrics(series, sym, desc, is_grp=False))
    else:
        for item in target_mapping:
            comp_tickers = [t for t in item["tickers"] if t in prices_df.columns]
            if not comp_tickers:
                continue
            sub_closes = prices_df[comp_tickers].mean(axis=1).dropna()
            if len(sub_closes) < 5:
                continue
            stats_list.append(calculate_metrics(
                sub_closes,
                item["symbol"],
                item["description"],
                is_grp=True,
                n_lvl=item.get("next_level", "sub")
            ))

    return stats_list

@app.get("/api/holdings")
def get_holdings():
    return {}

from fastapi.responses import FileResponse

@app.get("/")
def serve_index():
    return FileResponse("index.html")

@app.get("/app.js")
def serve_js():
    return FileResponse("app.js", media_type="application/javascript")

@app.get("/style.css")
def serve_css():
    return FileResponse("style.css", media_type="text/css")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("api_server:app", host="127.0.0.1", port=8000, reload=True)