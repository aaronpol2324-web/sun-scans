import concurrent.futures
import re
from tradingview_screener import Query, col
import pandas as pd
import yfinance as yf
from finvizfinance.screener.custom import Custom
import logging
from datetime import datetime

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# =====================================================================
# STANDARDIZED METRIC COLUMNS
# ---------------------------------------------------------------------
# Every scan tab (TradingView-based, Finviz-based, and the leveraged-ETF
# tab) is written out with this exact same set of columns, in this exact
# order. Each source pipes its own raw fields onto these same internal
# names below, so they merge cleanly on Momentum/All_Scans too.
# =====================================================================

# TradingView fields every individual/momentum scan selects, on top of
# whatever extra fields that specific scan needs purely for its own
# filter logic (SMA10, EMA5, etc. below). 'Perf.Y', 'sector', 'industry'
# and 'average_volume_30d_calc' are newer additions here and a bit less
# battle-tested than the rest of this list -- if TradingView ever
# rejects one of them, run_individual/run_momentum fall back to
# SAFE_FALLBACK_FIELDS below rather than losing the whole scan.
# 'Volatility.M' was already used in this script's filter conditions
# before any of this, so it's treated as proven, not speculative.
STANDARD_DISPLAY_FIELDS = [
    'name', 'close', 'volume', 'market_cap_basic', 'float_shares_outstanding',
    'average_volume_30d_calc', 'price_52_week_low', 'change',
    'Perf.W', 'Perf.1M', 'Perf.3M', 'Perf.6M', 'Perf.Y',
    'sector', 'industry',
    'earnings_per_share_diluted_yoy_growth_fq', 'total_revenue_yoy_growth_fq',
    'Volatility.M',
]

# A conservative fallback used only if the full field list above causes
# TradingView to reject the request (e.g. one of the newer field names
# above turns out wrong). Every field here was already used successfully
# in the original version of this script.
SAFE_FALLBACK_FIELDS = [
    'name', 'close', 'volume', 'market_cap_basic', 'float_shares_outstanding',
    'price_52_week_low', 'change', 'Perf.W', 'Perf.1M', 'Perf.3M', 'Perf.6M',
    'earnings_per_share_diluted_yoy_growth_fq', 'total_revenue_yoy_growth_fq',
    'Volatility.M',
]

# The final, fixed column set (raw/internal names) written to every
# single tab, in this exact order. Anything not in this list gets
# dropped from the output; anything in this list that a given scan's
# source can't supply is written out blank rather than omitted, so
# every tab has identical headers.
ORDERED_METRIC_COLUMNS = [
    'name',                                        # Ticker
    'close',                                        # Price
    'volume',                                       # Volume
    'Rel_Volume',                                    # Relative Volume (ratio, 30d avg based)
    'Dollar_Volume',                                 # $ Volume (abbreviated)
    'market_cap_basic',                              # Market Cap (abbreviated)
    'float_shares_outstanding',                      # Float
    'change',                                        # 1D %
    'Perf.W',                                        # 1W %
    'Perf.1M',                                       # 1M %
    'Perf.3M',                                       # 3M %
    'Perf.6M',                                       # 6M %
    'Perf.Y',                                        # 1Y %
    'sector',                                        # Sector
    'industry',                                      # Industry
    'earnings_per_share_diluted_yoy_growth_fq',      # EPS YoY Growth % (Qtr)
    'total_revenue_yoy_growth_fq',                   # Revenue YoY Growth % (Qtr)
    'Pct_Above_52W_Low',                             # % Above 52W Low
    'Volatility.M',                                  # 1M Volatility %
    'Status',                                        # Match/No Match (leveraged ETFs only; blank elsewhere)
]

def _dedup(fields):
    seen = []
    for f in fields:
        if f not in seen:
            seen.append(f)
    return seen


def _build_query(fields, where_conditions, order_by, ascending, limit):
    return (
        Query().set_markets('america')
        .select(*_dedup(fields))
        .where(*where_conditions)
        .order_by(order_by, ascending=ascending)
        .limit(limit)
    )


# 1. Individual standalone operational & tightness scan tabs
# Each entry is a spec (not a pre-built Query) so run_individual can
# retry with SAFE_FALLBACK_FIELDS if the full field list ever errors.
individual_scans = {
    "1_Fundamental_Growth": {
        "extra_fields": [],
        "where": lambda: [
            col('type') == 'stock',
            col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']),
            col('market_cap_basic') > 300_000_000,
            col('average_volume_60d_calc') > 300_000,
            col('float_shares_outstanding') < 100_000_000,
            col('earnings_per_share_diluted_yoy_growth_fq') > 25,
            col('free_cash_flow_yoy_growth_ttm') > 25,
            col('total_revenue_yoy_growth_fq') > 25
        ],
        "order_by": "change", "ascending": False, "limit": 300,
    },
    "3_Post_Earnings_Cont_Base": {
        "extra_fields": ['SMA20'],
        "where": lambda: [
            col('type') == 'stock',
            col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']),
            col('close') > col('SMA20'),
            col('market_cap_basic') > 50_000_000,
            col('average_volume_60d_calc') > 250_000,
            col('relative_volume_10d_calc') >= 2,
            col('float_shares_outstanding') < 50_000_000,
            col('gap') > 5
        ],
        "order_by": "change", "ascending": False, "limit": 300,
    },
    "4_Strongest_Stock_JK": {
        "extra_fields": ['SMA10', 'SMA50'],
        "where": lambda: [
            col('type') == 'stock',
            col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']),
            col('market_cap_basic').between(300_000_000, 10_000_000_000),
            col('average_volume_60d_calc') > 500_000,
            col('Volatility.M') > 3,
            col('float_shares_outstanding') < 50_000_000,
            col('earnings_per_share_diluted_yoy_growth_fq') > 25,
            col('total_revenue_yoy_growth_fq') > 25,
            col('close') > col('SMA50')
        ],
        "order_by": "change", "ascending": False, "limit": 300,
    },
    "5_Strongest_Stock_10B_Rev_30_JK": {
        "extra_fields": ['SMA10', 'SMA50'],
        "where": lambda: [
            col('type') == 'stock',
            col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']),
            col('market_cap_basic') > 10_000_000_000,
            col('average_volume_60d_calc') > 500_000,
            col('Volatility.M') > 2,
            col('float_shares_outstanding') < 150_000_000,
            col('earnings_per_share_diluted_yoy_growth_fq') > 25,
            col('total_revenue_yoy_growth_fq') > 25,
            col('close') > col('SMA50')
        ],
        "order_by": "change", "ascending": False, "limit": 300,
    },
    "Daily_Tightness_Swing": {
        "extra_fields": ['EMA5', 'SMA10', 'SMA20'],
        "where": lambda: [
            col('type') == 'stock',
            col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']),
            col('market_cap_basic') > 300_000_000,
            col('average_volume_60d_calc') > 300_000,
            col('volume') > 100_000,
            col('float_shares_outstanding') < 50_000_000,
            col('Volatility.M') > 3.5,
            col('Perf.W') < 5
        ],
        "order_by": "change", "ascending": False, "limit": 300,
    },
}

# 2. All 8 Momentum Scans
momentum_scans = {
    "Mom_1W_Small": {
        "mcap_group": "$300M - $10B", "timeframe": "1 Week", "is_large": False,
        "extra_fields": ['SMA10'],
        "where": lambda: [col('type') == 'stock', col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']), col('market_cap_basic').between(300_000_000, 10_000_000_000), col('average_volume_60d_calc') > 300_000, col('volume') > 100_000, col('float_shares_outstanding') < 50_000_000, col('Volatility.M') > 3, col('Perf.W') > 20],
    },
    "Mom_1M_Small": {
        "mcap_group": "$300M - $10B", "timeframe": "1 Month", "is_large": False,
        "extra_fields": ['SMA10'],
        "where": lambda: [col('type') == 'stock', col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']), col('market_cap_basic').between(300_000_000, 10_000_000_000), col('average_volume_60d_calc') > 300_000, col('volume') > 100_000, col('float_shares_outstanding') < 50_000_000, col('Volatility.M') > 3, col('Perf.1M') > 30],
    },
    "Mom_3M_Small": {
        "mcap_group": "$300M - $10B", "timeframe": "3 Months", "is_large": False,
        "extra_fields": ['SMA10'],
        "where": lambda: [col('type') == 'stock', col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']), col('market_cap_basic').between(300_000_000, 10_000_000_000), col('average_volume_60d_calc') > 300_000, col('volume') > 100_000, col('float_shares_outstanding') < 50_000_000, col('Volatility.M') > 3, col('Perf.3M') > 70],
    },
    "Mom_6M_Small": {
        "mcap_group": "$300M - $10B", "timeframe": "6 Months", "is_large": False,
        "extra_fields": ['SMA10'],
        "where": lambda: [col('type') == 'stock', col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']), col('market_cap_basic').between(300_000_000, 10_000_000_000), col('average_volume_60d_calc') > 300_000, col('volume') > 100_000, col('float_shares_outstanding') < 50_000_000, col('Volatility.M') > 3, col('Perf.6M') > 100],
    },
    "Mom_1W_Large": {
        "mcap_group": "> $10B", "timeframe": "1 Week", "is_large": True,
        "extra_fields": ['SMA10'],
        "where": lambda: [col('type') == 'stock', col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']), col('market_cap_basic') > 10_000_000_000, col('average_volume_60d_calc') > 300_000, col('volume') > 100_000, col('float_shares_outstanding') < 150_000_000, col('Perf.W') > 20],
    },
    "Mom_1M_Large": {
        "mcap_group": "> $10B", "timeframe": "1 Month", "is_large": True,
        "extra_fields": ['SMA10'],
        "where": lambda: [col('type') == 'stock', col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']), col('market_cap_basic') > 10_000_000_000, col('average_volume_60d_calc') > 300_000, col('volume') > 100_000, col('float_shares_outstanding') < 150_000_000, col('Perf.1M') > 30],
    },
    "Mom_3M_Large": {
        "mcap_group": "> $10B", "timeframe": "3 Months", "is_large": True,
        "extra_fields": ['SMA10'],
        "where": lambda: [col('type') == 'stock', col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']), col('market_cap_basic') > 10_000_000_000, col('average_volume_60d_calc') > 300_000, col('volume') > 100_000, col('float_shares_outstanding') < 150_000_000, col('Perf.3M') > 70],
    },
    "Mom_6M_Large": {
        "mcap_group": "> $10B", "timeframe": "6 Months", "is_large": True,
        "extra_fields": ['SMA10'],
        "where": lambda: [col('type') == 'stock', col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']), col('market_cap_basic') > 10_000_000_000, col('average_volume_60d_calc') > 300_000, col('volume') > 100_000, col('float_shares_outstanding') < 150_000_000, col('Perf.6M') > 100],
    },
}

# 3. Finviz Scan Configuration (filters only -- columns are fixed for
# every Finviz scan via FINVIZ_CUSTOM_COLUMNS below)
finviz_scans = {
    "Finviz_High_Short_Float": {
        'Market Cap.': '+Small (over $300mln)',
        'Average Volume': 'Over 1M',
        'Float': 'Under 100M',
        'Float Short': 'Over 30%'
    },
    "Finviz_IPO_Weekly": {
        'Market Cap.': '+Small (over $300mln)',
        'EPS growthnext year': 'Positive (>0%)',
        'Average Volume': 'Over 1M',
        'IPO Date': 'In the last 3 years',
        '50-Day Simple Moving Average': 'Price above SMA50'
    },
    "Finviz_Steve_Jacobs_RS": {
        'Market Cap.': '+Mid (over $2bln)',
        'EPS growthnext year': 'Positive (>0%)',
        'Average Volume': 'Over 500K',
        'Price': 'Over $10',
        'Performance': 'Month +20%',
        'Performance 2': 'Quarter +30%',
        '200-Day Simple Moving Average': 'Price above SMA200',
        '50-Day Simple Moving Average': 'Price above SMA50'
    }
}

leveraged_tickers = [
    "TQQQ", "SOXX", "QLD", "SSO", "SPXL", "FNGU", "TECL", "UPRO", "NVDL", "MUU",
    "TSLL", "BULZ", "USD", "FAS", "TMF", "SNXX", "AGQ", "KORU", "GDXU", "ROM",
    "NUGT", "TNA", "GGLL", "AMDL", "UDOW", "UGL", "UYG", "FNGO", "YINN", "MSFU",
    "MULL", "LABU", "CONL", "DDM", "TSMX", "NAIL", "SPYU", "NVDU", "JNUG", "METU",
    "PLTU", "NVDX", "UCO", "BOIL", "BMNU", "MVLL", "PTIR", "DPST", "DFEN",
    "INTW", "SPCH", "URTY", "AMZU", "SPUU", "SNDU", "ERX", "NOWL", "MSTX", "ASTX",
    "UWM", "IRE", "DGP", "GUSH", "AVL", "MQQQ", "AAOX", "ORCX", "LITX", "UVIX",
    "CWEB", "NBIL", "FBL", "CURE", "EDC", "CRCG", "SKUU", "AAPU", "BEX", "NEBX",
    "TSLT", "SMCX", "NFXL", "MVV", "RKLX", "CRCA", "BABX", "ORCU", "DLLL", "IONX",
    "FNGG", "SHNY", "SNDG", "OKLL", "CRWG", "NVII", "MRVU", "GDXD", "BRZU", "BIB",
    "CBRG", "ROBN", "WEBL", "COHX", "APPX", "WDCX", "CHAU", "RXL", "SKHX", "NRGU",
    "OILU", "DIG", "CWVX", "HIBL", "QQQU", "CRDU", "MSFL", "NBIG", "TSLR", "CRMG",
    "AMUU", "MSOX", "MIDU", "HOOG", "LOFF", "MAGX", "UBT", "LINT", "QBTX", "LABX",
    "TDAX", "IREX", "IONL", "CRWL", "URSP", "INDL", "QCML", "URE", "SOFX", "RDTL",
    "SPCU", "ARMG", "XUSP", "TSLG", "FNGD", "HIMZ", "CCUP", "GOOX", "LVLN", "SMU",
    "EURL", "TSMU", "DUSL", "AVGG", "BMNG", "APLX", "VRTL", "BNKU", "UTSL", "DRN",
    "AMZZ", "NVDG", "IBX", "BEG", "UYM", "BRKU", "SMCL", "RGTX", "NVTX", "EET",
    "MRAL", "LRCU", "TEMT", "ADBG", "MRNX", "PALU", "SPAL", "URAA", "NVOX", "QPUX",
    "NFLU", "UMDD", "MVPL", "TSMG", "RDWU", "AVGU", "TSII", "ASMI", "ASMG", "ONDL",
    "AMA", "SAA", "CRWU", "QQUP", "EFO", "UXI", "GEVX", "GEVD", "TYD", "PLTG",
    "LNOK", "IREG"
]


def run_leveraged_etfs():
    """Pull ~1 year of daily bars for every leveraged ETF in one batched
    Yahoo Finance request, then compute the exact same standardized
    metrics (relative volume, $ volume, 1D/1W/1M/3M/6M/1Y %, % above
    52-week low, 1-month volatility %) used on every other tab, plus
    the MATCH/NO MATCH Status column. ETFs don't have a market cap,
    float, sector/industry, or EPS/revenue growth in the way stocks
    do, so those columns are simply left blank here."""
    name = "Leveraged_Setups"

    if not leveraged_tickers:
        return name, pd.DataFrame()

    logger = yf.utils.get_yf_logger()
    original_level = logger.level
    logger.setLevel(logging.CRITICAL)

    results = []
    try:
        # ~1 year of history in a single batched request, so 1Y performance
        # and a true 52-week low can be computed for every ticker at once.
        data = yf.download(leveraged_tickers, period="1y", progress=False, group_by='ticker')

        for t in leveraged_tickers:
            try:
                if len(leveraged_tickers) == 1:
                    df = data
                else:
                    df = data[t] if t in data.columns.levels[0] else pd.DataFrame()

                df = df.dropna(subset=['Close', 'Volume'])
                n = len(df)
                if df.empty or n < 30:
                    continue

                close = df['Close']
                volume = df['Volume']
                latest_close = float(close.iloc[-1])
                latest_volume = float(volume.iloc[-1])

                def pct_change_over(bars):
                    if n > bars:
                        base = float(close.iloc[-(bars + 1)])
                        if base:
                            return round(((latest_close - base) / base) * 100, 2)
                    return None

                # 52-week (or shorter, if less history exists) low
                window_52w = close.tail(min(n, 252))
                week52_low = float(window_52w.min())
                pct_above_52w_low = round(((latest_close - week52_low) / week52_low) * 100, 2) if week52_low else None

                # 1-month volatility %, ~21 trading days (average daily
                # high-low range as a % of close), so ETFs get a volatility
                # figure using the same underlying data as everything else
                # instead of leaving it blank.
                vol_window = df.tail(min(n, 21))
                volatility_1m = round((((vol_window['High'] - vol_window['Low']) / vol_window['Close']) * 100).mean(), 2)

                # Relative volume vs. its own 30-day average (matches the
                # 30-day window used for the TradingView-sourced scans)
                avg_volume_30 = float(volume.tail(min(n, 30)).mean())
                rel_volume = round(latest_volume / avg_volume_30, 2) if avg_volume_30 else None

                dollar_volume = round(latest_close * latest_volume, 2)

                # MATCH / NO MATCH still uses a 60-day-average-based
                # liquidity + range check, same as the original criteria
                # (independent of which volatility metric gets displayed).
                match_window = df.tail(min(n, 60))
                avg_dollar_vol_60 = float((match_window['Volume'] * match_window['Close']).mean())
                range_pct_60 = float((((match_window['High'] - match_window['Low']) / match_window['Close']) * 100).mean())
                status = "MATCH" if (avg_dollar_vol_60 >= 100_000_000 and range_pct_60 > 1.5) else "NO MATCH"

                results.append({
                    'name': t,
                    'close': round(latest_close, 2),
                    'volume': latest_volume,
                    'Rel_Volume': rel_volume,
                    'Dollar_Volume': dollar_volume,
                    'change': pct_change_over(1),
                    'Perf.W': pct_change_over(5),
                    'Perf.1M': pct_change_over(21),
                    'Perf.3M': pct_change_over(63),
                    'Perf.6M': pct_change_over(126),
                    'Perf.Y': pct_change_over(252),
                    'Pct_Above_52W_Low': pct_above_52w_low,
                    'Volatility.M': volatility_1m,
                    'Status': status,
                })
            except Exception:
                continue
    except Exception:
        pass
    finally:
        logger.setLevel(original_level)

    out_df = pd.DataFrame(results)
    if not out_df.empty:
        out_df.insert(0, 'Source_Scan', name)
        out_df = out_df.sort_values(by=['Dollar_Volume'], ascending=[False], na_position='last')

    return name, out_df


def run_individual(name, spec):
    fields = spec.get('extra_fields', []) + STANDARD_DISPLAY_FIELDS
    try:
        query = _build_query(fields, spec['where'](), spec['order_by'], spec['ascending'], spec['limit'])
        _, df = query.get_scanner_data()
    except Exception as e:
        print(f"Warning: full field set failed for {name} ({e}); retrying with a reduced field set.")
        try:
            fallback_fields = spec.get('extra_fields', []) + SAFE_FALLBACK_FIELDS
            query = _build_query(fallback_fields, spec['where'](), spec['order_by'], spec['ascending'], spec['limit'])
            _, df = query.get_scanner_data()
        except Exception as e2:
            print(f"Error in {name}: {e2}")
            return name, pd.DataFrame()

    try:
        if not df.empty:
            if name == "4_Strongest_Stock_JK":
                df = df[(df['close'] >= df['price_52_week_low'] * 1.70) & (df['SMA10'] <= df['close']) & (df['SMA10'] >= df['close'] * 0.90)].copy()
            elif name == "5_Strongest_Stock_10B_Rev_30_JK":
                df = df[(df['close'] >= df['price_52_week_low'] * 1.70) & (df['SMA10'] <= df['close']) & (df['SMA10'] >= df['close'] * 0.97)].copy()
            elif name == "Daily_Tightness_Swing":
                df = df[
                    (df['close'] >= df['price_52_week_low'] * 1.50) &
                    (df['EMA5'] <= df['close']) &
                    (df['EMA5'] >= df['close'] * 0.97) &
                    (df['SMA10'] > df['SMA20'])
                ].copy()

            df.insert(0, 'Source_Scan', name)
        return name, df
    except Exception as e:
        print(f"Error in {name}: {e}")
        return name, pd.DataFrame()

def run_momentum(key, info):
    fields = info.get('extra_fields', []) + STANDARD_DISPLAY_FIELDS
    try:
        query = _build_query(fields, info['where'](), 'change', False, 300)
        _, df = query.get_scanner_data()
    except Exception as e:
        print(f"Warning: full field set failed for momentum {key} ({e}); retrying with a reduced field set.")
        try:
            fallback_fields = info.get('extra_fields', []) + SAFE_FALLBACK_FIELDS
            query = _build_query(fallback_fields, info['where'](), 'change', False, 300)
            _, df = query.get_scanner_data()
        except Exception as e2:
            print(f"Error in momentum {key}: {e2}")
            return key, pd.DataFrame()

    try:
        if not df.empty:
            low_mult = 0.80 if not info['is_large'] else 0.90
            df = df[(df['close'] >= df['price_52_week_low'] * 1.50) & (df['SMA10'] <= df['close']) & (df['SMA10'] >= df['close'] * low_mult)].copy()
            df.insert(0, 'Source_Scan', key)
            df.insert(1, 'Market_Cap_Group', info['mcap_group'])
            df.insert(2, 'Timeframe', info['timeframe'])
            return key, df
    except Exception as e:
        print(f"Error in momentum {key}: {e}")
    return key, pd.DataFrame()

# Finviz's plain Overview screener only returns a fixed set of columns
# (Ticker/Company/Sector/Industry/Country/Market Cap/P/E/Price/Change/
# Volume) with no way to customize them. The Custom screener class lets
# us request exactly the columns we need by index (see
# finvizfinance.constants.CUSTOM_SCREENER_COLUMNS) so Finviz-sourced
# rows can carry the same standardized metrics as everything else.
FINVIZ_CUSTOM_COLUMNS = [1, 3, 4, 6, 22, 23, 25, 42, 43, 44, 45, 46, 51, 58, 64, 65, 66, 67, 81]
#                        Ticker Sector Industry MarketCap. EPSQ/Q SalesQ/Q Float
#                        PerfWeek PerfMonth PerfQuart PerfHalf PerfYear VolatilityMonth
#                        52WLow RelVolume Price Change Volume PreviousClose

FINVIZ_COLUMN_MAP = {
    'Ticker': 'name',
    'Sector': 'sector',
    'Industry': 'industry',
    'Market Cap': 'market_cap_basic',
    'EPS Q/Q': 'earnings_per_share_diluted_yoy_growth_fq',
    'Sales Q/Q': 'total_revenue_yoy_growth_fq',
    'Float': 'float_shares_outstanding',
    'Perf Week': 'Perf.W',
    'Perf Month': 'Perf.1M',
    'Perf Quart': 'Perf.3M',
    'Perf Half': 'Perf.6M',
    'Perf Year': 'Perf.Y',
    'Volatility M': 'Volatility.M',
    '52W Low': 'price_52_week_low',
    'Rel Volume': 'Rel_Volume',
    'Price': 'close',
    'Change': 'change',      # covers either spelling Finviz's page might use
    'Change %': 'change',    # confirmed actual header text as of Sep 2026
    'Volume': 'volume',
    # Finviz's live 'Change %' is intraday-only -- it prints as "-" (blank)
    # whenever nothing has traded yet today (pre-market, weekends,
    # holidays), unlike TradingView's 'change' which always carries the
    # last completed session's move. Previous Close is a static reference
    # value that's populated regardless, so it's pulled in purely as a
    # fallback: add_derived_metrics() below fills any blank 'change' with
    # (Price - Previous Close) / Previous Close, matching what
    # TradingView shows on the same non-trading day.
    'Previous Close': 'Prev_Close',
    'Prev Close': 'Prev_Close',
}

# finvizfinance parses any "%"-suffixed cell as a fraction (e.g. 5.23% ->
# 0.0523). Scale these back up to the same percent-unit convention
# (5.23) used everywhere else in this workbook. Rel_Volume, Price,
# Volume, Market Cap, and 52W Low are NOT percent-suffixed on Finviz's
# site, so they don't need this.
FINVIZ_PERCENT_FIELDS = [
    'change', 'Perf.W', 'Perf.1M', 'Perf.3M', 'Perf.6M', 'Perf.Y', 'Volatility.M',
    'earnings_per_share_diluted_yoy_growth_fq', 'total_revenue_yoy_growth_fq',
]

def _normalize_header(h) -> str:
    """Collapse whitespace and lowercase a column header, so a rename
    lookup isn't defeated by stray/odd whitespace or case differences
    in what Finviz's page actually returns."""
    return re.sub(r'\s+', ' ', str(h)).strip().lower()


# Case/whitespace-insensitive version of FINVIZ_COLUMN_MAP, built once.
_FINVIZ_COLUMN_MAP_NORMALIZED = {_normalize_header(k): v for k, v in FINVIZ_COLUMN_MAP.items()}

# Fields every Finviz scan should always come back with. If one goes
# missing after the rename below, it's either a genuinely-closed-market
# situation (e.g. 'change' is Finviz's live intraday move, which prints
# as "-" -- and therefore blank -- on weekends/holidays when nothing
# has traded yet today) or a real header-mapping mismatch; the warning
# printed below shows the raw columns Finviz actually returned so the
# two cases are easy to tell apart from the console output.
FINVIZ_EXPECTED_FIELDS = ['name', 'close', 'volume', 'change']


def run_finviz_scan(name, filters_dict):
    try:
        fcustom = Custom()
        fcustom.set_filter(filters_dict=filters_dict)
        df = fcustom.screener_view(columns=list(FINVIZ_CUSTOM_COLUMNS))

        if df is not None and not df.empty:
            # Match Finviz's returned headers to FINVIZ_COLUMN_MAP
            # case/whitespace-insensitively rather than with an exact
            # rename(), so a stray space or capitalization difference in
            # what the page sends back doesn't silently drop a column
            # (leaving it blank downstream) instead of just renaming it.
            rename_map = {}
            for c in df.columns:
                target = _FINVIZ_COLUMN_MAP_NORMALIZED.get(_normalize_header(c))
                if target:
                    rename_map[c] = target
            df = df.rename(columns=rename_map)

            missing = [f for f in FINVIZ_EXPECTED_FIELDS if f not in df.columns]
            if missing:
                print(f"Warning: Finviz scan {name} is missing expected field(s) {missing} "
                      f"after column mapping. Raw columns received: {list(df.columns)}")

            for field in FINVIZ_PERCENT_FIELDS:
                if field in df.columns:
                    df[field] = pd.to_numeric(df[field], errors='coerce') * 100

            # Previous Close isn't %-suffixed, so the base scraper leaves it
            # as a plain numeric string -- convert it here rather than via
            # FINVIZ_PERCENT_FIELDS (which would incorrectly x100 it).
            if 'Prev_Close' in df.columns:
                df['Prev_Close'] = pd.to_numeric(df['Prev_Close'], errors='coerce')

            if 'change' in df.columns and df['change'].isna().all() and 'Prev_Close' not in df.columns:
                print(f"Note: Finviz scan {name}'s 'Change %' came back entirely blank "
                      f"(likely no live trading today) and no Previous Close column was "
                      f"found to fall back on. Raw columns received: {list(df.columns)}")

            df.insert(0, 'Source_Scan', name)
            return name, df
    except Exception as e:
        print(f"Error in Finviz scan {name}: {e}")
    return name, pd.DataFrame()


# =====================================================================
# EXCEL OUTPUT STYLING
# ---------------------------------------------------------------------
# Everything below controls ONLY how results_dict gets written to the
# .xlsx file at the end of the script. None of it touches the scan
# logic above or the console/print output — that all stays identical.
# =====================================================================

# ---- Color palette (colorblind-friendly: slate/cyan vs. orange, no red/green) ----
COLOR_HEADER_BG = "2F3B44"       # dark slate header background
COLOR_HEADER_FONT = "FFFFFF"    # white header text
COLOR_ACCENT_CYAN = "4FD1C5"    # cyan accent (borders, hyperlinks)
COLOR_SCALE_LOW = "F2994A"      # orange = low/negative end of a metric
COLOR_SCALE_MID = "F2F2F0"      # neutral light grey = midpoint
COLOR_SCALE_HIGH = "5B7C99"     # slate-blue = high/positive end of a metric
COLOR_ZEBRA = "F5F6F7"          # subtle zebra-stripe fill
COLOR_MATCH_FILL = "E7ECEF"     # soft slate tint for Status == MATCH (good)
COLOR_NOMATCH_FILL = "FCE8D6"   # soft orange tint for Status == NO MATCH (bad)

# ---- Friendly column names shown in Excel (raw field -> readable label) ----
FRIENDLY_NAMES = {
    'name': 'Ticker',
    'close': 'Price',
    'volume': 'Volume',
    'Rel_Volume': 'Relative Volume',
    'Dollar_Volume': '$ Volume',
    'market_cap_basic': 'Market Cap',
    'float_shares_outstanding': 'Float',
    'change': '1D %',
    'Perf.W': '1W %',
    'Perf.1M': '1M %',
    'Perf.3M': '3M %',
    'Perf.6M': '6M %',
    'Perf.Y': '1Y %',
    'sector': 'Sector',
    'industry': 'Industry',
    'earnings_per_share_diluted_yoy_growth_fq': 'EPS YoY Growth % (Qtr)',
    'total_revenue_yoy_growth_fq': 'Revenue YoY Growth % (Qtr)',
    'Pct_Above_52W_Low': '% Above 52W Low',
    'Volatility.M': '1M Volatility %',
    'Status': 'Status',
    # Structural columns (combined sheets only)
    'Source_Scan': 'Scan',
    'Market_Cap_Group': 'Market Cap Group',
    'Timeframe': 'Timeframe',
}

# Sheets that stack multiple scans together and therefore need the
# 'Scan' (and grouping) columns kept and pinned to the front
COMBINED_SHEETS = {'All_Scans', 'Momentum'}

# Friendly-name column format buckets
PERCENT_COLUMNS = {
    '1D %', '1W %', '1M %', '3M %', '6M %', '1Y %',
    'EPS YoY Growth % (Qtr)', 'Revenue YoY Growth % (Qtr)',
    '% Above 52W Low', '1M Volatility %',
}
# Abbreviated (K/M/B) dollar formats, per the user's request
ABBREVIATED_CURRENCY_COLUMNS = {'Market Cap', '$ Volume'}
INTEGER_COLUMNS = {'Volume', 'Float'}
PRICE_COLUMNS = {'Price'}
RATIO_COLUMNS = {'Relative Volume'}

ABBREVIATED_CURRENCY_FORMAT = '[>=1000000000]$#,##0.0,,,"B";[>=1000000]$#,##0.0,,"M";$#,##0.0,"K"'

# Order in which individual sheets should appear (matches the ticker
# print-out order already used further down in the script)
SHEET_DISPLAY_ORDER = [
    "Mom_1W_Small", "Mom_1M_Small", "Mom_3M_Small", "Mom_6M_Small",
    "Mom_1W_Large", "Mom_1M_Large", "Mom_3M_Large", "Mom_6M_Large",
    "1_Fundamental_Growth", "3_Post_Earnings_Cont_Base",
    "4_Strongest_Stock_JK", "5_Strongest_Stock_10B_Rev_30_JK",
    "Daily_Tightness_Swing", "Leveraged_Setups", "Finviz_High_Short_Float",
    "Finviz_IPO_Weekly", "Finviz_Steve_Jacobs_RS"
]


def _safe_ratio(numerator, denominator):
    ratio = numerator / denominator.replace(0, pd.NA)
    return ratio.replace([float('inf'), float('-inf')], pd.NA)


def add_derived_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """Compute Rel_Volume / Dollar_Volume / Pct_Above_52W_Low wherever
    the source columns needed for them are available, without
    overwriting a value a source (Finviz, the leveraged-ETF section)
    already computed and provided directly."""
    df = df.copy()

    if 'volume' in df.columns and 'average_volume_30d_calc' in df.columns:
        computed = _safe_ratio(df['volume'], df['average_volume_30d_calc']).round(2)
        df['Rel_Volume'] = df['Rel_Volume'].fillna(computed) if 'Rel_Volume' in df.columns else computed

    if 'close' in df.columns and 'volume' in df.columns:
        computed = (df['close'] * df['volume']).round(2)
        df['Dollar_Volume'] = df['Dollar_Volume'].fillna(computed) if 'Dollar_Volume' in df.columns else computed

    if 'close' in df.columns and 'price_52_week_low' in df.columns:
        computed = (_safe_ratio(df['close'] - df['price_52_week_low'], df['price_52_week_low']) * 100).round(2)
        df['Pct_Above_52W_Low'] = df['Pct_Above_52W_Low'].fillna(computed) if 'Pct_Above_52W_Low' in df.columns else computed

    # Finviz-only fallback: its live 'change' (1D %) is intraday-only and
    # comes back blank whenever nothing has traded yet today (weekends,
    # holidays, pre-market). 'Prev_Close' is a static reference value that's
    # populated regardless, so wherever 'change' is missing, derive it from
    # Price vs. Previous Close -- the same last-completed-session move
    # TradingView's 'change' already carries forward on those days. This
    # never overwrites a real intraday value Finviz did provide.
    if 'change' in df.columns and 'close' in df.columns and 'Prev_Close' in df.columns:
        computed = (_safe_ratio(df['close'] - df['Prev_Close'], df['Prev_Close']) * 100).round(2)
        df['change'] = df['change'].fillna(computed)

    return df


def finalize_columns(df: pd.DataFrame, is_combined: bool) -> pd.DataFrame:
    """Build the final, fixed-column output: the 'Scan' column up front
    on combined sheets, then every column in ORDERED_METRIC_COLUMNS in
    that exact order -- created blank if the source didn't provide it --
    so every tab has identical headers. Market_Cap_Group/Timeframe are
    still computed internally (Momentum is sorted by them), but are
    dropped here rather than shown as their own columns."""
    out = pd.DataFrame(index=df.index)
    if is_combined and 'Source_Scan' in df.columns:
        out['Source_Scan'] = df['Source_Scan']
    for c in ORDERED_METRIC_COLUMNS:
        out[c] = df[c] if c in df.columns else pd.NA
    return out


def prepare_sheet_df(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    """Full pipeline applied to a scan's dataframe right before it is
    written to its Excel tab: derive metrics -> fix columns -> rename."""
    if df is None or df.empty:
        return df
    is_combined = sheet_name in COMBINED_SHEETS
    df = add_derived_metrics(df)
    df = finalize_columns(df, is_combined)
    df = df.rename(columns=FRIENDLY_NAMES)
    return df


def _column_format(header: str):
    """Return the Excel number format string for a given friendly
    column header, or None if it should stay general/text."""
    if header in PERCENT_COLUMNS:
        return '0.00"%"'   # literal % suffix; values are already in percent units
    if header in ABBREVIATED_CURRENCY_COLUMNS:
        return ABBREVIATED_CURRENCY_FORMAT
    if header in INTEGER_COLUMNS:
        return '#,##0'
    if header in PRICE_COLUMNS:
        return '0.00'
    if header in RATIO_COLUMNS:
        return '0.00"x"'
    return None


def _preview_text(value, header: str) -> str:
    """Approximate how `value` will actually be DISPLAYED in Excel for
    this column, so column widths can be fit to that rather than to
    Python's full-precision str() of a raw float (which is often much
    longer than what the number_format will ever show)."""
    if value is None:
        return ""
    try:
        if isinstance(value, float) and value != value:  # NaN
            return ""
        if header in PERCENT_COLUMNS:
            return f"{float(value):.2f}%"
        if header in ABBREVIATED_CURRENCY_COLUMNS:
            v = float(value)
            av = abs(v)
            if av >= 1_000_000_000:
                return f"${v / 1_000_000_000:.1f}B"
            if av >= 1_000_000:
                return f"${v / 1_000_000:.1f}M"
            return f"${v / 1_000:.1f}K"
        if header in INTEGER_COLUMNS:
            return f"{float(value):,.0f}"
        if header in PRICE_COLUMNS:
            return f"{float(value):.2f}"
        if header in RATIO_COLUMNS:
            return f"{float(value):.2f}x"
    except (TypeError, ValueError):
        pass
    return str(value)


def style_worksheet(ws: Worksheet, n_rows: int, n_cols: int):
    """Apply header styling, freeze panes, autofilter, column widths,
    number formats, zebra striping, and color-scale / status
    conditional formatting to one already-populated worksheet.

    Column widths are fit to the DATA only (not the header text), so
    columns stay tight even under long headers like "Revenue YoY
    Growth % (Qtr)". The header row wraps and grows tall enough to
    stay legible at that width instead."""
    if n_rows == 0 or n_cols == 0:
        return

    header_fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type='solid')
    header_font = Font(bold=True, color=COLOR_HEADER_FONT, size=11)
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    header_border = Border(bottom=Side(style='medium', color=COLOR_ACCENT_CYAN))
    zebra_fill = PatternFill(start_color=COLOR_ZEBRA, end_color=COLOR_ZEBRA, fill_type='solid')

    headers = []
    for col_idx in range(1, n_cols + 1):
        cell = ws.cell(row=1, column=col_idx)
        headers.append(cell.value)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align
        cell.border = header_border

    # Zebra striping on data rows
    for row_idx in range(2, n_rows + 1):
        if row_idx % 2 == 0:
            for col_idx in range(1, n_cols + 1):
                ws.cell(row=row_idx, column=col_idx).fill = zebra_fill

    MIN_COL_WIDTH = 8    # a floor so very-short data (e.g. "MATCH") isn't cramped, and so
                         # single long header words (e.g. "Relative") don't split mid-word
    MAX_COL_WIDTH = 45   # a safety ceiling for a truly extreme outlier value; long but normal
                         # text (e.g. a GICS industry name) should still get its real width
    max_header_lines = 1

    # Per-column number format + data-only width
    for col_idx, header in enumerate(headers, start=1):
        col_letter = get_column_letter(col_idx)
        fmt = _column_format(header)
        data_max_len = 0
        for row_idx in range(2, n_rows + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if fmt:
                cell.number_format = fmt
            val_len = len(_preview_text(cell.value, header))
            if val_len > data_max_len:
                data_max_len = val_len
        width = max(MIN_COL_WIDTH, min(data_max_len + 2, MAX_COL_WIDTH))
        ws.column_dimensions[col_letter].width = width

        header_len = len(str(header)) if header else 0
        lines_needed = max(1, -(-header_len // max(int(width) - 1, 1)))  # ceil division
        max_header_lines = max(max_header_lines, lines_needed)

        # Color-scale conditional formatting on percent-style metric columns
        if header in PERCENT_COLUMNS and n_rows > 1:
            rng = f"{col_letter}2:{col_letter}{n_rows}"
            ws.conditional_formatting.add(
                rng,
                ColorScaleRule(
                    start_type='min', start_color=COLOR_SCALE_LOW,
                    mid_type='percentile', mid_value=50, mid_color=COLOR_SCALE_MID,
                    end_type='max', end_color=COLOR_SCALE_HIGH,
                )
            )

        # Highlight MATCH / NO MATCH text status (Leveraged_Setups sheet,
        # and any combined sheet that includes leveraged-ETF rows)
        if header == 'Status' and n_rows > 1:
            rng = f"{col_letter}2:{col_letter}{n_rows}"
            ws.conditional_formatting.add(
                rng, CellIsRule(operator='equal', formula=['"MATCH"'],
                                 fill=PatternFill(start_color=COLOR_MATCH_FILL, end_color=COLOR_MATCH_FILL, fill_type='solid'))
            )
            ws.conditional_formatting.add(
                rng, CellIsRule(operator='equal', formula=['"NO MATCH"'],
                                 fill=PatternFill(start_color=COLOR_NOMATCH_FILL, end_color=COLOR_NOMATCH_FILL, fill_type='solid'))
            )

    # Grow the header row to fit however many lines its longest wrapped
    # header needs at the (now data-driven, often narrow) column widths.
    ws.row_dimensions[1].height = 15 * max_header_lines + 8

    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions


def write_summary_sheet(writer, sheet_counts: dict, sheet_order: list):
    """Write a first 'Summary' tab: a title, generation timestamp, and
    one row per scan tab with its match count and a clickable link
    that jumps straight to that tab."""
    book = writer.book
    ws = book.create_sheet('Summary', 0)

    title_font = Font(bold=True, size=16, color=COLOR_HEADER_BG)
    subtitle_font = Font(italic=True, size=10, color="666666")
    link_font = Font(color=COLOR_ACCENT_CYAN, underline='single', bold=True)
    header_fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type='solid')
    header_font = Font(bold=True, color=COLOR_HEADER_FONT)

    ws['A1'] = "Trading Scans — Summary"
    ws['A1'].font = title_font
    ws['A2'] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws['A2'].font = subtitle_font
    ws['A4'] = "See the 'All_Scans' tab for every match on a single page (use the Scan filter dropdown)."
    ws['A4'].font = subtitle_font

    header_row = 6
    for col_idx, label in enumerate(['Scan', 'Matches', 'Open Tab'], start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=label)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center')

    row = header_row + 1
    for name in sheet_order:
        if name not in sheet_counts:
            continue
        ws.cell(row=row, column=1, value=name)
        ws.cell(row=row, column=2, value=sheet_counts[name]).alignment = Alignment(horizontal='center')
        link_cell = ws.cell(row=row, column=3, value="Open →")
        link_cell.hyperlink = f"#'{name}'!A1"
        link_cell.font = link_font
        link_cell.alignment = Alignment(horizontal='center')
        if row % 2 == 0:
            for c in range(1, 4):
                ws.cell(row=row, column=c).fill = PatternFill(start_color=COLOR_ZEBRA, end_color=COLOR_ZEBRA, fill_type='solid')
        row += 1

    ws.column_dimensions['A'].width = 34
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 14


def write_styled_workbook(results_dict: dict, path: str):
    """Write results_dict to an .xlsx file with all the formatting
    above applied: styled headers, freeze/filter, number formats,
    color-scale + status highlighting, the same fixed metric columns
    on every tab, a Summary tab, and 'All_Scans' as the one-page
    combined view."""
    sheet_counts = {name: (0 if df is None or df.empty else len(df)) for name, df in results_dict.items()}

    # Sheet order: All_Scans (one-page view) and Momentum first, then
    # everything else in the logical scan order, then any leftovers.
    ordered_names = [n for n in ('All_Scans', 'Momentum') if n in results_dict]
    ordered_names += [n for n in SHEET_DISPLAY_ORDER if n in results_dict and n not in ordered_names]
    ordered_names += [n for n in results_dict if n not in ordered_names]

    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        for sheet_name in ordered_names:
            df = results_dict[sheet_name]
            if df is None or df.empty:
                continue
            styled_df = prepare_sheet_df(df, sheet_name)
            styled_df.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]
            style_worksheet(ws, n_rows=len(styled_df) + 1, n_cols=len(styled_df.columns))

        write_summary_sheet(writer, sheet_counts, ordered_names)


if __name__ == "__main__":
    print("Executing all scans concurrently...")
    results_dict = {}
    momentum_dfs = []
    all_collected_dfs = []

    with concurrent.futures.ThreadPoolExecutor(max_workers=14) as executor:
        ind_futures = {executor.submit(run_individual, name, spec): name for name, spec in individual_scans.items()}
        mom_futures = {executor.submit(run_momentum, key, info): key for key, info in momentum_scans.items()}
        finviz_futures = {executor.submit(run_finviz_scan, name, filters): name for name, filters in finviz_scans.items()}
        lev_future = executor.submit(run_leveraged_etfs)

        for future in concurrent.futures.as_completed(ind_futures):
            name, df = future.result()
            results_dict[name] = df
            if not df.empty:
                all_collected_dfs.append(df)
            print(f"Finished {name}: {len(df)} matches.")

        for future in concurrent.futures.as_completed(mom_futures):
            key, df = future.result()
            if not df.empty:
                momentum_dfs.append(df)
                all_collected_dfs.append(df)

        for future in concurrent.futures.as_completed(finviz_futures):
            name, df = future.result()
            results_dict[name] = df
            if not df.empty:
                all_collected_dfs.append(df)
            print(f"Finished {name}: {len(df)} matches.")

        lev_name, lev_df = lev_future.result()
        results_dict[lev_name] = lev_df

        if not lev_df.empty:
            matched_lev_df = lev_df[lev_df['Status'] == 'MATCH'].copy()
            if not matched_lev_df.empty:
                all_collected_dfs.append(matched_lev_df)
            print(f"Finished {lev_name}: {len(matched_lev_df)} matches found.")

    if momentum_dfs:
        master_momentum = pd.concat(momentum_dfs, ignore_index=True)
        master_momentum.sort_values(by=['Market_Cap_Group', 'Timeframe', 'change'], ascending=[True, True, False], inplace=True)
        results_dict['Momentum'] = master_momentum

    if all_collected_dfs:
        master_all_scans = pd.concat(all_collected_dfs, ignore_index=True)
        results_dict['All_Scans'] = master_all_scans

        print("\n" + "="*60)
        print(" COPYABLE COMMA-SEPARATED TICKER LISTS (LOGICAL ORDER)")
        print("="*60)

        logical_order = [
            "Mom_1W_Small", "Mom_1M_Small", "Mom_3M_Small", "Mom_6M_Small",
            "Mom_1W_Large", "Mom_1M_Large", "Mom_3M_Large", "Mom_6M_Large",
            "1_Fundamental_Growth", "3_Post_Earnings_Cont_Base",
            "4_Strongest_Stock_JK", "5_Strongest_Stock_10B_Rev_30_JK",
            "Daily_Tightness_Swing", "Leveraged_Setups", "Finviz_High_Short_Float",
            "Finviz_IPO_Weekly", "Finviz_Steve_Jacobs_RS"
        ]

        if 'Source_Scan' in master_all_scans.columns and 'name' in master_all_scans.columns:
            for scan_name in logical_order:
                group_df = master_all_scans[master_all_scans['Source_Scan'] == scan_name]
                if not group_df.empty:
                    scan_tickers = sorted(group_df['name'].dropna().unique())
                    scan_string = ", ".join(scan_tickers)
                    print(f"\n[{scan_name}] ({len(scan_tickers)} tickers):")
                    print(scan_string)

        if 'name' in master_all_scans.columns:
            unique_tickers = sorted(master_all_scans['name'].dropna().unique())
            ticker_string = ", ".join(unique_tickers)
            print("\n" + "="*60)
            print(f" MASTER COMMA-SEPARATED LIST (ALL SCANS - {len(unique_tickers)} tickers):")
            print("="*60)
            print(ticker_string)
            print("="*60 + "\n")


    excel_path = r"C:\TradingScans\My_Scans.xlsx"
    try:
        write_styled_workbook(results_dict, excel_path)
        print(f"Done! Saved cleanly to: {excel_path}")
    except PermissionError:
        alt_path = r"C:\TradingScans\My_Scans_NEW.xlsx"
        write_styled_workbook(results_dict, alt_path)
        print(f"\n[WARNING] Excel file was locked. Saved master file to: {alt_path}")