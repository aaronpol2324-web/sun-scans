import concurrent.futures
import numbers
import re
from tradingview_screener import Query, col
import pandas as pd
import yfinance as yf
from finvizfinance.screener.custom import Custom
from finvizfinance.constants import filter_dict as _finviz_filter_dict
import logging
from datetime import datetime

import xlsxwriter

# Excel output now goes through xlsxwriter directly (not openpyxl, and not
# pandas' ExcelWriter wrapper) -- the one thing openpyxl categorically
# cannot do, at any version, is WRITE native Excel sparklines (it can't
# even fully read them back -- see the "Sparkline Group extension is not
# supported" warning openpyxl itself prints on a file that has them). The
# RS Dashboard tabs' in-cell trend charts need that, so every tab in this
# workbook is now built with direct write() calls into an xlsxwriter
# Workbook rather than styling a DataFrame after pandas dumps it in.

# finvizfinance (as installed, v1.5.0) is missing "Sales growth ttm" from
# its filter table -- confirmed this is a real, live Finviz filter (its
# own screener UI has it, and the URL it generates uses the code
# "fa_salesyoyttm_<code>"), the library itself just never added it. Patch
# it into the SAME dict object the library's screener code reads from
# (mutating in place, not reassigning, so the library's own imported
# reference picks it up too), using the identical option codes as its
# "EPS growth ttm" entry -- Finviz uses this same neg/pos/poslow/high/
# u5-u30/o5-o30 convention across all of its growth-rate filters.
if 'Sales growth ttm' not in _finviz_filter_dict:
    _finviz_filter_dict['Sales growth ttm'] = {
        'prefix': 'fa_salesyoyttm',
        'option': {
            'Any': '',
            'Negative (<0%)': 'neg',
            'Positive (>0%)': 'pos',
            'Positive Low (0-10%)': 'poslow',
            'High (>25%)': 'high',
            'Under 5%': 'u5', 'Under 10%': 'u10', 'Under 15%': 'u15',
            'Under 20%': 'u20', 'Under 25%': 'u25', 'Under 30%': 'u30',
            'Over 5%': 'o5', 'Over 10%': 'o10', 'Over 15%': 'o15',
            'Over 20%': 'o20', 'Over 25%': 'o25', 'Over 30%': 'o30',
        },
    }

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
    # "PEG Delayed Reaction" (Jeff Sun / @jfsrev, TradingView screener) --
    # names that reported earnings last week, are still liquid/volatile
    # enough to trade, and haven't necessarily finished reacting to the
    # print yet. Unlike the 8 scans above, this one's full criteria is
    # expressed in its own `where` (no separate 52-week-low/SMA10-band
    # filter), and one condition can't be pushed to TradingView's API at
    # all: "Price x 60D avg volume > $50M" needs two columns multiplied
    # together, which this library's Column class has no operator for --
    # so it's applied locally in run_momentum() instead, after the query.
    # 'earnings_release_trading_date_fq' (last reported earnings date) is
    # this project's best inference from the confirmed, symmetric
    # 'earnings_release_next_trading_date_fq' field already documented in
    # the tradingview_screener library -- if TradingView rejects it,
    # run_momentum() automatically retries this one scan without the
    # earnings-date condition and prints a warning rather than losing it.
    "Mom_PEG_Delayed_Reaction": {
        "mcap_group": "> $1B", "timeframe": "Post-Earnings (PEG Delayed)",
        "extra_fields": ['ADR', 'average_volume_60d_calc'],
        "where": lambda: [
            col('type') == 'stock',
            col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']),
            col('market_cap_basic') > 1_000_000_000,
            col('ADR') > 4,
            col('average_volume_60d_calc') > 2_000_000,
            col('earnings_release_trading_date_fq').in_week_range(-1, -1),
        ],
        "where_fallback": lambda: [
            col('type') == 'stock',
            col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']),
            col('market_cap_basic') > 1_000_000_000,
            col('ADR') > 4,
            col('average_volume_60d_calc') > 2_000_000,
        ],
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
    },
    # "High-Velocity Quality Screener" (Jeff Sun / @jfsrev) -- large-cap
    # quality-growth names: reasonable valuation, positive/growing
    # earnings and sales, strong balance sheet and margins, in a
    # confirmed long-term uptrend. 'Sales growth ttm' is patched into
    # finvizfinance's filter table above since the library doesn't ship
    # it natively even though it's a real, live Finviz filter.
    "Finviz_High_Velocity_Quality": {
        'Market Cap.': '+Large (over $10bln)',
        'P/E': 'Under 40',
        'Forward P/E': 'Under 25',
        'PEG': 'Under 2',
        'EPS growthnext year': 'Positive (>0%)',
        'EPS growthpast 5 years': 'Positive (>0%)',
        'Sales growthqtr over qtr': 'Positive (>0%)',
        'Sales growth ttm': 'Over 10%',
        'Return on Investment': 'Over +10%',
        'Current Ratio': 'Over 1',
        'LT Debt/Equity': 'Under 1',
        'Gross Margin': 'Over 15%',
        'Operating Margin': 'Positive (>0%)',
        'Net Profit Margin': 'Positive (>0%)',
        '200-Day Simple Moving Average': 'Price above SMA200',
        'Current Volume': 'Over 0',
    },
}

leveraged_tickers = [
    "TQQQ", "SOXL", "QLD", "SSO", "FNGU", "SPXL", "TECL", "UPRO", "TSLL", "NVDL",
    "MUU", "BULZ", "USD", "TMF", "FAS", "SNXX", "AGQ", "ROM", "GDXU", "TNA",
    "NUGT", "KORU", "GGLL", "AMDL", "UGL", "UDOW", "UYG", "FNGO", "MULL", "YINN",
    "TSMX", "SPYU", "NAIL", "MSFU", "MSTU", "DDM", "LABU", "METU", "NVDU", "CONL",
    "JNUG", "UCO", "NVDX", "PLTU", "BOIL", "BMNU", "MVLL", "SNDU", "DPST", "SPCH",
    "PTIR", "IRE", "DFEN", "URTY", "SPUU", "INTW", "AMZU", "MQQQ", "MSTX", "ERX",
    "AVL", "AAOX", "GUSH", "ASTX", "ORCX", "NOWL", "UWM", "DGP", "BEX", "FBL",
    "SKUU", "DLLL", "AAPU", "RKLX", "NBIL", "CWEB", "UVIX", "CURE", "EDC", "ORCU",
    "LITX", "TSLT", "MVV", "NFXL", "IONX", "CRCA", "NEBX", "GDXD", "CRCG", "ROBN",
    "FNGG", "SMCX", "NVII", "BRZU", "MRVU", "BABX", "SKHX", "CRWG", "SHNY", "SNDG",
    "OKLL", "CRDU", "WEBL", "RXL", "DIG", "BIB", "CHAU", "NRGU", "COHX", "CRWL",
    "NBIG", "OILU", "WDCX", "CBRG", "APPX", "LOFF", "TSLR", "QQQU", "AMUU", "HIBL",
    "QCML", "GOOX", "LINT", "MSOX", "QBTX", "MSFL", "TDAX", "SPCU", "MAGX", "HOOG",
    "MIDU", "IONL", "CWVX", "RDTL", "CRMG", "UBT", "HIMZ", "IREX", "SOFX", "FNGD",
    "ARMG", "URE", "INDL", "XUSP", "BMNG", "LVLN", "SMU", "BEG", "LABX", "EET",
    "AVGG", "UTSL", "MRAL", "EURL", "TSLG", "APLX", "PALU", "BNKU", "DUSL", "TSMU",
    "BRKU", "URSP", "RGTX", "SPAL", "DRN", "VRTL", "SMCL", "CCUP", "NVDG", "UYM",
    "AMZZ", "TEMT", "ADBG", "IBX", "AVGU", "GEVX", "NVOX", "MVPL", "QPUX", "RDWU",
    "TSMG", "TSII", "ONDL", "LNOK", "URAA", "NVTX", "LRCU", "IREG", "EFO", "UMDD",
    "ASMG", "OILD", "QCMU", "TYD", "SAA", "QQUP", "NFLU", "UXI", "AMA", "SMHU",
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
    where_conditions = info['where']()
    df = None
    try:
        query = _build_query(fields, where_conditions, 'change', False, 300)
        _, df = query.get_scanner_data()
    except Exception as e:
        # Scans with a 'where_fallback' (currently just Mom_PEG_Delayed_Reaction)
        # get one extra retry tier: drop the condition this project is least
        # sure of (an inferred TradingView field name) before falling back
        # to a reduced field set like every other scan does.
        if 'where_fallback' in info:
            print(f"Warning: momentum {key}'s where clause failed ({e}); "
                  f"retrying without its least-certain condition -- results "
                  f"will be broader than intended until that's confirmed.")
            try:
                where_conditions = info['where_fallback']()
                query = _build_query(fields, where_conditions, 'change', False, 300)
                _, df = query.get_scanner_data()
            except Exception as e2:
                e = e2

        if df is None:
            print(f"Warning: full field set failed for momentum {key} ({e}); retrying with a reduced field set.")
            try:
                fallback_fields = info.get('extra_fields', []) + SAFE_FALLBACK_FIELDS
                query = _build_query(fallback_fields, where_conditions, 'change', False, 300)
                _, df = query.get_scanner_data()
            except Exception as e3:
                print(f"Error in momentum {key}: {e3}")
                return key, pd.DataFrame()

    try:
        if not df.empty:
            if key == "Mom_PEG_Delayed_Reaction":
                # This scan's own where() already expresses its full
                # criteria -- the only piece that can't be pushed to
                # TradingView's API is "Price x 60D avg volume > $50M"
                # (no arithmetic-between-columns filter exists there), so
                # it's applied locally here instead.
                if 'average_volume_60d_calc' in df.columns:
                    df = df[(df['close'] * df['average_volume_60d_calc']) > 50_000_000].copy()
            else:
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
# RS DASHBOARD (RS_Groups / RS_Indices_Sectors tabs)
# ---------------------------------------------------------------------
# This used to chase a facsimile of a specific paid dashboard (Jeff
# Sun's, @jfsrev) whose exact scoring formula was never published.
# After months of reverse-engineering guesses that kept missing, we
# dropped that entirely in favor of Real Relative Strength (RRS): a
# fully public, well-documented indicator (the ThinkOrSwim community's
# "RealRelativeStrength" study -- the user supplied the actual script)
# with no unknowns to guess at. Every number below is computed exactly
# as that script defines it; nothing here is an approximation of
# somebody else's private formula.
#
# THE IDEA: a ratio like "close / SPY's close" treats a 1% move the
# same whether the ticker is a sleepy utility or a wild small-cap --
# but a 1% day is nothing for the small-cap and huge for the utility.
# RRS fixes that by asking "given how much SPY moved, and how volatile
# this ticker normally is, how big a move would we EXPECT? Was the
# ACTUAL move bigger or smaller than that, and by how much (in units
# of this ticker's own typical daily range)?" A positive RRS means the
# ticker outperformed what its own volatility would predict from SPY's
# move; negative means it underperformed. It's a volatility-adjusted
# excess return, not a percentage -- so despite the column names below
# keeping "RS" in them, these are NOT percentiles or percentages, and
# aren't formatted as one.
#
# THE FORMULA (window = RS_THRUST_WINDOW for "RS Thrust", or
# RS_ONE_MONTH_WINDOW for "1-Mth RRS" -- see _rs_rrs):
#   symbol_move   = ticker's close now minus its close `window` bars ago
#   bench_move    = SPY's close now minus its close `window` bars ago
#   symbol_atr    = Wilder's ATR(window) of the ticker
#   bench_atr     = Wilder's ATR(window) of SPY
#   power_index   = bench_move / bench_atr        -- SPY's move, in units of SPY's own ATR
#   expected_move = power_index * symbol_atr       -- what the ticker "should" have done, scaled to ITS OWN volatility
#   RRS           = (symbol_move - expected_move) / symbol_atr   -- the excess, in units of the ticker's own ATR
#
# SPY's own row can't be compared to itself this way -- the algebra
# above cancels to exactly 0 for every single day when ticker == SPY,
# a flat, uninformative line. So for that one row (or when SPY history
# is unavailable at all) we skip the "expected move" subtraction and
# just report symbol_move / symbol_atr -- "how many ATRs has this
# moved," a self-referential momentum reading rather than an RS one.
#
# "RS Thrust" uses the short (RS_THRUST_WINDOW, "the past week")
# window -- fast-moving, for who's accelerating right now. "1-Mth
# RRS" uses the long (RS_ONE_MONTH_WINDOW) window -- slower, for the
# underlying trend. The "1-Mth RS" histogram plots the trailing
# RS_SPARKLINE_WINDOW days of that same slow RRS reading, so the
# number and the chart are the same series -- and because RRS is a
# genuinely signed quantity (unlike the stochastic/percentile designs
# tried before this), the histogram is colored two-tone: one color
# above zero, another below, matching what the original script's own
# chart does.
# =====================================================================

RS_THRUST_WINDOW = 5      # "the past week"
RS_ONE_MONTH_WINDOW = 21

# Just the watchlist of (ticker, name) pairs -- the row order actually
# written to the sheet is a default sort by 1-Mth RRS (see
# write_rs_groups_sheet), so the order here doesn't matter.
#
# Every entry below is one we've directly confirmed, by ticker AND
# name, in his most recent COMPLETE "Groups" table screenshot
# (x.com/jfsrev/status/2101207434270552238, posted 9/19/2026 -- its own
# right-hand text panel independently lists the same tickers/names as
# the table itself, confirming the table wasn't cut off). That post is
# now the single source of truth for this list -- not a running union
# of every screenshot we've ever been shown. A prior round of this
# list had accumulated 20 tickers (CLOU, XHS, KIE, BETZ, IHF, PBJ,
# USO, SOCL, IYZ, CNBS, WCLD, IGV, XOP, EWY, GNR, EWZ, ICLN, UFO, XSW,
# GUNR) that were confirmed at some earlier point but are NOT in this
# latest complete table, so they're dropped rather than held onto on
# the assumption they're still current.
RS_GROUPS = [
    ("BUG", "Pure Cybersecurity Software"),
    ("CIBR", "Cybersecurity Software & Infra"),
    ("BOAT", "Global Shipping"),
    ("MAGS", "Magnificent 7"),
    ("WOOD", "Timber & Lumber"),
    ("TLT", "20+ Year Treasury Bonds"),
    ("ARKG", "ARK Genomics"),
    ("FDN", "US Internet Giants"),
    ("UNG", "Natural Gas"),
    ("ESPO", "E-Sports"),
    ("AIQ", "AI Software & Data Processing"),
    ("HYDR", "Hydrogen Energy & Fuel Cells"),
    ("IDGT", "Digital Infrastructure & Data Centers"),
    ("XTL", "Telecom"),
    ("ARKX", "ARK Space Exploration"),
    ("SVIX", "Short VIX Futures (Volatility)"),
    ("ASHR", "China A-Share"),
    ("BUZZ", "Social Media"),
    ("NASA", "Space Economy"),
    ("ARKK", "ARK Innovation"),
    ("SPMO", "S&P 500 Momentum"),
    ("XSD", "Semiconductors (Equal)"),
    ("DRAM", "Memory"),
    ("ETHA", "Ether Spot"),
    ("AMLP", "Energy MLPs"),
    ("BOTZ", "AI & Robotics"),
    ("ARKQ", "ARK Robotics"),
    ("ARKW", "ARK Internet & Next-Gen Tech"),
    ("QTUM", "Quantum/AI"),
    ("BAI", "AI & Tech Active"),
    ("MEME", "Meme"),
    ("SOXX", "Broad Semiconductor"),
    ("WGMI", "Internet Services & Infrastructure"),
    ("SMH", "Semiconductors Giants"),
    ("SOLZ", "Solana Spot"),
    ("IBIT", "Bitcoin Spot"),
    ("BLOK", "Blockchain"),
    ("ROBO", "Robotics & Automation"),
    ("PPLT", "Platinum"),
    ("GXC", "China ETF"),
    ("GRID", "Infrastructure & Electrical Equipment"),
    ("KURE", "China Health"),
    ("ARKF", "ARK Fintech"),
    ("SLV", "Silver"),
    ("PBE", "Dynamic Biotech"),
]

# (category, ticker, name) -- category becomes its own filterable
# column (via the auto-filter every tab already gets) rather than a
# merged banner row, so this can go through the exact same generic
# tab-writing pipeline as everything else.
RS_INDICES_SECTORS = [
    ("Index", "RSP", "S&P 500 Equal Weight"),
    ("Index", "SPY", "S&P 500"),
    ("Index", "QQQ", "Nasdaq-100"),
    ("Index", "QQQE", "Nasdaq-100 Equal Weight"),
    ("Index", "IWM", "Russell 2000"),
    ("Index", "DIA", "Dow 30"),
    ("Index", "SPMO", "S&P 500 Momentum"),
    ("Index", "TLT", "20+ Year Treasury Bonds"),
    ("Segment", "IJS", "Small-Cap 600 Value"),
    ("Segment", "IJR", "Small-Cap 600"),
    ("Segment", "IJT", "Small-Cap 600 Growth"),
    ("Segment", "IJJ", "MidCap 400 Value"),
    ("Segment", "IJH", "MidCap 400"),
    ("Segment", "IJK", "MidCap 400 Growth"),
    ("Segment", "IVE", "Large-Cap 500 Value"),
    ("Segment", "IVV", "S&P 500"),
    ("Segment", "IVW", "Large-Cap 500 Growth"),
    ("EW Sector", "RSPH", "Equal Weight Health Care"),
    ("EW Sector", "RSPT", "Equal Weight Technology"),
    ("EW Sector", "SPY", "S&P 500"),
    ("EW Sector", "RSPG", "Equal Weight Energy"),
    ("EW Sector", "RSPC", "Equal Weight Communication"),
    ("EW Sector", "RSPU", "Equal Weight Utilities"),
    ("EW Sector", "RSPM", "Equal Weight Material"),
    ("EW Sector", "RSPR", "Equal Weight Real Estate"),
    ("EW Sector", "RSPS", "Equal Weight Staples"),
    ("EW Sector", "RSPN", "Equal Weight Industrial"),
    ("EW Sector", "RSPF", "Equal Weight Financials"),
    ("EW Sector", "RSPD", "Equal Weight Discretionary"),
    ("SPDR Sector", "SPY", "S&P 500"),
    ("SPDR Sector", "XLV", "Health Care"),
    ("SPDR Sector", "XLK", "Technology"),
    ("SPDR Sector", "XLC", "Communication Services"),
    ("SPDR Sector", "XLP", "Consumer Staples"),
    ("SPDR Sector", "XLE", "Energy"),
    ("SPDR Sector", "XLU", "Utilities"),
    ("SPDR Sector", "XLRE", "Real Estate"),
    ("SPDR Sector", "XLB", "Materials"),
    ("SPDR Sector", "XLF", "Financials"),
    ("SPDR Sector", "XLI", "Industrials"),
    ("SPDR Sector", "XLY", "Consumer Discretionary"),
]


def _fetch_rs_history(tickers):
    """Same batched-yfinance-download pattern as run_leveraged_etfs:
    ~1 year of daily bars in one request, with a per-ticker
    dropna/min-length guard so one bad symbol can't take down the
    whole tab. Returns {ticker: DataFrame} for every ticker that came
    back with enough history; a ticker simply absent from the dict
    shows up as a blank row later rather than a fabricated one."""
    logger = yf.utils.get_yf_logger()
    original_level = logger.level
    logger.setLevel(logging.CRITICAL)
    history = {}
    try:
        data = yf.download(tickers, period="1y", progress=False, group_by='ticker')
        for t in tickers:
            try:
                df = data if len(tickers) == 1 else (data[t] if t in data.columns.levels[0] else pd.DataFrame())
                df = df.dropna(subset=['Close'])
                if len(df) >= RS_ONE_MONTH_WINDOW + 1:
                    history[t] = df
            except Exception:
                continue
    except Exception as e:
        print(f"Warning: RS dashboard's batched download failed ({e}); its tabs will be mostly blank this run.")
    finally:
        logger.setLevel(original_level)
    return history


RS_SPARKLINE_WINDOW = 21


def _rs_true_range(high: pd.Series, low: pd.Series, close: pd.Series) -> pd.Series:
    """Classic True Range: the largest of today's high-low range, and
    the gap between today's high/low and YESTERDAY's close -- so a gap
    up or down on the open still counts as "range," not just the
    intraday bar. The building block for Wilder's ATR below."""
    prev_close = close.shift(1)
    return pd.concat([
        high - low,
        (high - prev_close).abs(),
        (low - prev_close).abs(),
    ], axis=1).max(axis=1)


def _rs_wilders_atr(high: pd.Series, low: pd.Series, close: pd.Series, window: int) -> pd.Series:
    """Wilder's Average True Range: True Range smoothed with an EMA of
    alpha = 1/window (the classic Wilder smoothing constant), which is
    what `ewm(alpha=1/window, adjust=False)` computes. min_periods
    keeps the first `window`-1 bars NaN (not enough history yet)
    instead of quietly averaging over a too-short warmup window."""
    tr = _rs_true_range(high, low, close)
    return tr.ewm(alpha=1.0 / window, adjust=False, min_periods=window).mean()


def _rs_rrs(symbol_high, symbol_low, symbol_close,
            bench_high, bench_low, bench_close, window: int) -> pd.Series:
    """Real Relative Strength (RRS), vectorized across `symbol_close`'s
    own date index -- see the RS DASHBOARD header comment for the full
    derivation. `bench_*` must already be reindexed to that same date
    index by the caller.

    In words: take SPY's `window`-session move, express it in units of
    SPY's OWN typical volatility (its Wilder's ATR) -- that's the
    "power index," how hard the market pushed relative to its own
    normal noise. Scale that push by this ticker's own volatility to
    get an "expected move" -- what this ticker "should" have done if
    it simply moved with the market, proportionally. RRS is how far
    the ACTUAL move over/undershot that expectation, again measured in
    units of the ticker's own ATR: positive means it outperformed what
    its own volatility would predict from SPY's move, negative means
    it underperformed. Not a percentage -- a multiple of the ticker's
    own typical daily range."""
    symbol_atr = _rs_wilders_atr(symbol_high, symbol_low, symbol_close, window)
    bench_atr = _rs_wilders_atr(bench_high, bench_low, bench_close, window)
    symbol_move = symbol_close - symbol_close.shift(window)
    bench_move = bench_close - bench_close.shift(window)
    power_index = bench_move / bench_atr.where(bench_atr != 0)
    expected_move = power_index * symbol_atr
    return (symbol_move - expected_move) / symbol_atr.where(symbol_atr != 0)


def _rs_self_momentum(symbol_high, symbol_low, symbol_close, window: int) -> pd.Series:
    """Fallback for SPY's own row (or any ticker with no benchmark
    history to compare against): _rs_rrs's algebra cancels to exactly
    0 for every single day when the ticker IS the benchmark (a flat,
    uninformative line), so instead this skips the "expected move"
    subtraction entirely and reports the ticker's own `window`-session
    move in units of its own ATR -- "how many ATRs has this moved," a
    self-referential momentum reading rather than a relative-strength
    one."""
    symbol_atr = _rs_wilders_atr(symbol_high, symbol_low, symbol_close, window)
    symbol_move = symbol_close - symbol_close.shift(window)
    return symbol_move / symbol_atr.where(symbol_atr != 0)


def _rs_metrics_for(df: pd.DataFrame, spy_df: pd.DataFrame = None, ticker: str = None):
    """Fields for one ticker: latest price, 1D/1-Month % change, % off
    the trailing 52-week high, the trailing-21-session price series
    the price sparkline draws from, and the Real Relative Strength
    (RRS) readings ("RS Thrust" / "1-Mth RRS") that _rs_rrs computes --
    see that function and the RS DASHBOARD header comment above for
    what these are and why.

    spy_df: SPY's own OHLC DataFrame (same columns as `df`, i.e. High/
    Low/Close) -- the benchmark RRS is measured against. Optional only
    so this function still works standalone/in tests without it.

    ticker: the symbol this row IS, so we can tell when it's SPY
    itself. RS_Indices_Sectors carries SPY as an actual tracked row
    (under Index/EW Sector/SPDR Sector) in addition to using it as
    the external benchmark -- see _rs_self_momentum for why that row
    needs a different formula, not just a ratio against itself."""
    close = df['Close']
    high = df['High']
    low = df['Low']
    last = float(close.iloc[-1])
    pct_1d = round((last / float(close.iloc[-2]) - 1) * 100, 2) if len(close) >= 2 else None
    if len(close) > RS_ONE_MONTH_WINDOW:
        pct_1m = round((last / float(close.iloc[-1 - RS_ONE_MONTH_WINDOW]) - 1) * 100, 2)
    else:
        pct_1m = None

    week52_high = float(close.tail(min(len(close), 252)).max())
    pct_off_high = round((last / week52_high - 1) * 100, 2) if week52_high else None
    price_spark = close.tail(RS_SPARKLINE_WINDOW).astype(float).tolist()

    if ticker == "SPY" or spy_df is None:
        rrs_fast_series = _rs_self_momentum(high, low, close, RS_THRUST_WINDOW)
        rrs_slow_series = _rs_self_momentum(high, low, close, RS_ONE_MONTH_WINDOW)
    else:
        # Reindex SPY's OHLC onto THIS ticker's own date index rather
        # than assuming the two series line up positionally -- a
        # ticker that's missing a session SPY has (or vice versa)
        # would otherwise silently pair up the wrong days.
        bench_high = spy_df['High'].reindex(df.index)
        bench_low = spy_df['Low'].reindex(df.index)
        bench_close = spy_df['Close'].reindex(df.index)
        rrs_fast_series = _rs_rrs(high, low, close, bench_high, bench_low, bench_close, RS_THRUST_WINDOW)
        rrs_slow_series = _rs_rrs(high, low, close, bench_high, bench_low, bench_close, RS_ONE_MONTH_WINDOW)

    def _last_valid(s: pd.Series):
        s = s.dropna()
        return float(s.iloc[-1]) if len(s) else None

    thrust = _last_valid(rrs_fast_series)     # "RS Thrust" -- the fast/velocity reading
    onem_rrs = _last_valid(rrs_slow_series)   # "1-Mth RRS" -- the slow/structure reading

    # The histogram plots the trailing window of the SLOW RRS series
    # itself -- the exact same series "1-Mth RRS" is read from, not a
    # derived change or level of something else. RRS is already a
    # genuinely signed "how far above/below expectation" reading at
    # every date, so the number and the chart are, again, the same
    # series -- and because real-valued RRS essentially never ties two
    # days exactly, the single-peak-bar guarantee holds in practice
    # without needing a banded/bounded reading to force it.
    hist_series = rrs_slow_series.tail(RS_SPARKLINE_WINDOW).dropna()
    rs_spark = hist_series.round(4).tolist() if len(hist_series) else []

    return {"price": round(last, 2), "pct_1d": pct_1d, "pct_1m": pct_1m,
            "pct_off_high": pct_off_high, "thrust": thrust, "onem_rrs": onem_rrs,
            "price_spark": price_spark, "rs_spark": rs_spark}


def _rs_row(extra_fields, rec):
    """Builds one output row. "RS Thrust" and "1-Mth RRS" are both Real
    Relative Strength (RRS) readings straight off _rs_metrics_for --
    see _rs_rrs and the RS DASHBOARD header comment -- at two different
    lookback windows (RS_THRUST_WINDOW / RS_ONE_MONTH_WINDOW). Both are
    self-referential (or self-vs-SPY) per ticker, so a row's own
    numbers don't depend on which other tickers happen to be on the
    same tab or in the same category block. Unlike the old percentile/
    blend design this replaced, there's no clamping, banding, or
    blending here -- these are the formula's own raw output, in units
    of the ticker's own ATR, rounded to 2 decimal places."""
    base = dict(extra_fields)
    if rec is None:
        base.update({"Price": None, "RS Thrust": None, "1-Mth RRS": None,
                     "1D %": None, "1M %": None, "Off 52W High %": None,
                     "_PriceSpark": [], "_RSSpark": []})
    else:
        thrust, onem_rrs = rec["thrust"], rec["onem_rrs"]
        base.update({
            "Price": rec["price"],
            "RS Thrust": round(thrust, 2) if thrust is not None else None,
            "1-Mth RRS": round(onem_rrs, 2) if onem_rrs is not None else None,
            "1D %": rec["pct_1d"], "1M %": rec["pct_1m"], "Off 52W High %": rec["pct_off_high"],
            "_PriceSpark": rec["price_spark"], "_RSSpark": rec["rs_spark"],
        })
    return base


def run_rs_dashboard():
    """Builds the RS_Groups and RS_Indices_Sectors tabs. Returns
    (groups_df, indices_df); either can come back with some/all rows
    blank (price data unavailable) rather than raising, matching this
    project's usual degrade-visibly-not-silently philosophy.

    Both "RS Thrust" and "1-Mth RRS" are Real Relative Strength (RRS)
    readings, self-referential (or self-vs-SPY) per ticker (see
    _rs_rrs) -- a row's numbers don't depend on any other row, so
    RS_Groups and RS_Indices_Sectors don't need separate whole-tab vs.
    per-category ranking universes here; each ticker is just computed
    on its own."""
    all_tickers = sorted(set([t for t, _ in RS_GROUPS] + [t for _, t, _ in RS_INDICES_SECTORS] + ["SPY"]))
    history = _fetch_rs_history(all_tickers)
    spy_df = history.get("SPY")  # full OHLC DataFrame -- RRS needs SPY's High/Low too, not just Close

    g_records = [(_rs_metrics_for(history[ticker], spy_df, ticker) if ticker in history else None) for ticker, _ in RS_GROUPS]
    groups_df = pd.DataFrame([
        _rs_row({"Ticker": ticker, "Name": name}, rec)
        for (ticker, name), rec in zip(RS_GROUPS, g_records)
    ])

    indices_rows = []
    current_category = None
    cat_rows = []  # [(ticker, name, rec), ...] for the category currently being accumulated

    def flush_category():
        for ticker, name, rec in cat_rows:
            indices_rows.append(_rs_row({"Category": current_category, "Ticker": ticker, "Name": name}, rec))

    for category, ticker, name in RS_INDICES_SECTORS:
        if category != current_category:
            flush_category()
            cat_rows = []
            current_category = category
        rec = _rs_metrics_for(history[ticker], spy_df, ticker) if ticker in history else None
        cat_rows.append((ticker, name, rec))
    flush_category()
    indices_df = pd.DataFrame(indices_rows)

    n_priced = sum(1 for r in g_records if r is not None) + sum(1 for row in indices_rows if row.get("Price") is not None)
    n_total = len(g_records) + len(indices_rows)
    print(f"Finished RS Dashboard: {n_priced}/{n_total} tickers priced.")
    return groups_df, indices_df


# =====================================================================
# TREND TEMPLATE SCAN (N-Value / Bucket scoring)
# ---------------------------------------------------------------------
# An 11-rule stock screener plus 2 index-level "Market Direction" rules,
# reproduced from a 12-rule set the user sourced from a third-party
# website (see the uploaded sample CSV this was reverse-engineered
# from) -- minus that source's "Institutional Ownership >= 5%" rule
# (n=12 originally), which is dropped here entirely: TradingView's
# screener silently returns null for an unrecognized field instead of
# erroring, so a first attempt at this rule (a guessed field name that
# almost certainly doesn't exist) came back reading as a false FAIL on
# every single stock rather than a detectable error -- worse than not
# having the rule at all. Every stock is scored two ways rather than
# simple pass/fail:
#
#   - N-Value Rating: each rule is worth 2**n (n = the source site's
#     own original numbering, 1-12 minus the dropped 4 -- see
#     TT_RULE_NVALUES). A stock's N-Value Rating is the sum of 2**n
#     over every rule it PASSES. Range: 0 (fails everything) to
#     TT_MAX_NVALUE (passes everything) -- lower than the source
#     site's own 8190 max since one rule's weight (2**4 = 16) is
#     permanently missing.
#   - Bucket Rating: for each rule, how far past (or short of) its
#     passing threshold the stock's actual value is, as a fraction of
#     that threshold, multiplied by the rule's own 2**n weight -- so
#     barely passing contributes little and passing by a wide margin
#     contributes a lot, even between two stocks that pass the exact
#     same set of rules. Every rule uses this SAME percent-over-
#     threshold formula (see _tt_percent_over/_tt_rule_score) --
#     including Liquidity and Previous-Close, where the source site's
#     own sample data didn't fit that formula (its Liquidity score was
#     always 0 and its Previous-Close score was always a flat max
#     value). That formula is capped per-rule at TT_MAX_PERCENT_OVER
#     (see _tt_rule_score) -- without a cap, Liquidity's $20M floor and
#     Previous-Close's $10 floor are both so low that a normal stock
#     clears them by 10-50x, not the 15-20% a moving-average rule
#     typically clears its own threshold by, so those two rules alone
#     were dominating two-thirds or more of the total score. The one
#     rule that keeps a fixed 0 Bucket contribution on purpose,
#     uncapped or not, is 200d-MA-Rising: it's a pass/fail streak check
#     with no underlying "percent over threshold" to measure at all
#     (see below).
#
# Two rules need a full year of daily price history rather than a
# single TradingView snapshot value -- Relative Strength (an EMA of a
# daily-return ratio vs SPY) and whether the 200-day MA has been
# rising for 21 straight sessions -- the same constraint the RS
# Dashboard above has, and for the same reason: pulling that much
# history for the whole market is slow. So this scan runs in two
# stages: one TradingView query applies every OTHER rule (plus a
# broad-but-real universe filter) first, and only the tickers that
# query returns get a batched yfinance history pull for those two
# history-dependent rules.
#
# The two "Market Direction" rules (21d MA > 50d MA, 50d MA rising for
# 21 sessions) are evaluated against SPY by run_market_direction()
# below, separately from the per-stock scan -- per the user's choice,
# these are shown as an informational readout on the Summary tab and
# do NOT gate or filter the stock screener.
# =====================================================================

# (rule key -> n-value), using the source site's own original numbering
# (1-12) with 4 (Institutional Ownership) permanently absent -- see the
# section header for why. 2**n is both the N-Value Rating a passing
# rule contributes AND the Bucket Rating weight (before the
# TT_MAX_PERCENT_OVER cap) that rule's percent-over-threshold gets
# multiplied by.
TT_RULE_NVALUES = {
    'sma50_gt_sma150': 12,
    'sma150_gt_sma200': 11,
    'week52_span': 10,
    'relative_strength': 9,
    'liquidity': 8,
    'close_above_52w_high_25pct': 7,
    'prev_close_above_10': 6,
    'sma200_rising_21d': 5,
    'close_above_sma50': 3,
    'sales_qoq_yoy': 2,
    'eps_qoq_yoy': 1,
}

# Friendly labels for each rule's PASS/FAIL column on the tab.
TT_RULE_LABELS = {
    'sma50_gt_sma150': 'SMA50 > SMA150',
    'sma150_gt_sma200': 'SMA150 > SMA200',
    'week52_span': '52W High/Low Span',
    'relative_strength': 'Relative Strength > 1.0',
    'liquidity': 'Liquidity >= $20M',
    'close_above_52w_high_25pct': 'Within 25% of 52W High',
    'prev_close_above_10': 'Prev Close > $10',
    'sma200_rising_21d': '200d MA Rising (21d)',
    'close_above_sma50': 'Close > SMA50',
    'sales_qoq_yoy': 'Sales QoQ YoY > 25%',
    'eps_qoq_yoy': 'EPS QoQ YoY > 18%',
}

# Fixed walk order for every place that needs to iterate all 11 rules
# (column layout, total-score calcs, etc) -- n=12 down to n=1, skipping
# the dropped n=4.
TT_RULE_ORDER = list(TT_RULE_NVALUES.keys())

# Max possible N-Value Rating (every rule passed) -- lower than the
# source site's own 8190 since n=4 (Institutional Ownership) is gone.
TT_MAX_NVALUE = sum(2 ** n for n in TT_RULE_NVALUES.values())

# Cap on how far past its own threshold a single rule's Bucket Rating
# contribution can count, as a multiple of the threshold itself (2.0 =
# 200% over) -- see the section header for why this exists. Applied
# uniformly to every rule in _tt_rule_score, not just Liquidity/
# Previous-Close, so there's still only one formula, just a bounded one.
TT_MAX_PERCENT_OVER = 2.0

# Universe scoping for the initial TradingView query -- broad enough to
# rank a wide set of stocks, bounded enough that the stage-2 history
# pull (one yfinance call per survivor) stays fast. Tune freely.
TT_UNIVERSE_LIMIT = 500

# The source site's Liquidity rule wants "SMA50_volume" -- a 50-day
# average volume. That specific field isn't a confirmed TradingView
# name in this project; 'average_volume_60d_calc' is (used throughout
# the rest of this file already), so it stands in here as the closest
# proven equivalent: SMA50 (price) x 60-day average volume, rather
# than x 50-day.
TT_LIQUIDITY_VOLUME_FIELD = 'average_volume_60d_calc'


def _tt_percent_over(actual, threshold):
    """(actual - threshold) / threshold -- the Bucket Rating's "percent
    difference between the actual value and the passing threshold,"
    shared by every rule below (see the section header for why this
    one formula is used for all twelve). `actual` is always a Series;
    `threshold` may be one too (comparing two moving averages) or a
    plain number (e.g. the $10 Previous-Close floor). A zero threshold
    is treated as undefined (NA) rather than raising a divide-by-zero,
    and +-inf results (a near-zero threshold with a nonzero actual)
    collapse to NA the same way."""
    if isinstance(threshold, pd.Series):
        denom = threshold.mask(threshold == 0)
    else:
        denom = pd.NA if threshold == 0 else threshold
    ratio = (actual - threshold) / denom
    if isinstance(ratio, pd.Series):
        ratio = ratio.replace([float('inf'), float('-inf')], pd.NA)
    return ratio


def _tt_safe_round(series: pd.Series, ndigits: int) -> pd.Series:
    """Round a Series to ndigits, tolerating an all-NA (object-dtype)
    placeholder Series -- pd.NA has no __round__, so a plain .round()
    raises on the "field wasn't available this run" columns
    run_trend_template builds. pd.to_numeric coerces those to a proper
    all-NaN float64 Series first, which .round() handles fine; a
    Series that's already numeric passes through unaffected."""
    return pd.to_numeric(series, errors='coerce').round(ndigits)


def _tt_rule_score(passed: pd.Series, percent_over: pd.Series, n_value: int) -> pd.Series:
    """A rule's Bucket Rating contribution: percent_over (capped at
    TT_MAX_PERCENT_OVER -- see the section header) x 2**n_value if it
    passed, 0 otherwise -- including wherever the metric needed for it
    wasn't available at all, and never negative for a failing rule (a
    stock that fails still keeps whatever it earned from the rules it
    DOES pass)."""
    weight = 2 ** n_value
    passed_bool = passed.fillna(False).astype(bool)
    capped = percent_over.fillna(0.0).clip(upper=TT_MAX_PERCENT_OVER)
    contribution = capped * weight
    return contribution.where(passed_bool, 0.0)


def run_market_direction(index_ticker: str = "SPY"):
    """The two index-level "Market Direction" rules, evaluated against
    `index_ticker` (SPY, matching every other cross-market benchmark
    use in this file). Returns {'index_ticker', 'ma21_gt_ma50',
    'ma50_rising_21d'}; either bool comes back None if there wasn't
    enough history to compute it, rather than a guessed pass/fail.

    INFORMATIONAL ONLY: per the user's choice, this does not filter or
    hide anything on the Trend_Template tab -- write_summary_sheet
    just shows it as a readout at the top of the Summary tab."""
    result = {'index_ticker': index_ticker, 'ma21_gt_ma50': None, 'ma50_rising_21d': None}
    history = _fetch_rs_history([index_ticker])
    df = history.get(index_ticker)
    if df is None or len(df) < 71:  # 50d MA needs 50 bars, plus 21 more of the MA's own history
        print(f"Warning: not enough {index_ticker} history to compute Market Direction rules.")
        return result

    close = df['Close']
    sma21 = close.rolling(21).mean()
    sma50 = close.rolling(50).mean()
    if pd.notna(sma21.iloc[-1]) and pd.notna(sma50.iloc[-1]):
        result['ma21_gt_ma50'] = bool(sma21.iloc[-1] > sma50.iloc[-1])

    sma50_diff_tail = sma50.diff().tail(RS_ONE_MONTH_WINDOW)
    if len(sma50_diff_tail) == RS_ONE_MONTH_WINDOW and sma50_diff_tail.notna().all():
        result['ma50_rising_21d'] = bool((sma50_diff_tail > 0).all())

    print(f"Finished Market Direction ({index_ticker}): "
          f"21d>50d MA = {result['ma21_gt_ma50']}, 50d MA rising = {result['ma50_rising_21d']}.")
    return result


def run_trend_template():
    """The 11-rule Trend Template N-Value/Bucket scoring scan (see the
    section header above) against a broad TradingView stock universe.
    Relative Strength and 200d-MA-Rising are computed in a second pass
    over just the tickers the initial query returns (see the section
    header for why)."""
    name = "Trend_Template"
    risky_fields = ['SMA50', 'SMA150', 'SMA200', 'price_52_week_high', TT_LIQUIDITY_VOLUME_FIELD]
    where_conditions = [
        col('type') == 'stock',
        col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']),
        col('market_cap_basic') > 300_000_000,
        col('average_volume_60d_calc') > 200_000,
    ]

    df = None
    try:
        fields = STANDARD_DISPLAY_FIELDS + risky_fields
        query = _build_query(fields, where_conditions, 'change', False, TT_UNIVERSE_LIMIT)
        _, df = query.get_scanner_data()
    except Exception as e:
        print(f"Warning: {name}'s full field set failed ({e}); retrying with a minimal field "
              f"set -- SMA150/SMA200/52-week-high will all be left blank/not evaluated this run.")
        try:
            fields = SAFE_FALLBACK_FIELDS + ['SMA50', TT_LIQUIDITY_VOLUME_FIELD]
            query = _build_query(fields, where_conditions, 'change', False, TT_UNIVERSE_LIMIT)
            _, df = query.get_scanner_data()
        except Exception as e2:
            print(f"Error in {name}: {e2}")
            return name, pd.DataFrame()

    if df is None or df.empty:
        return name, pd.DataFrame()

    df = df.copy()
    has_sma150 = 'SMA150' in df.columns
    has_sma200 = 'SMA200' in df.columns
    has_52w_high = 'price_52_week_high' in df.columns

    passed = pd.DataFrame(index=df.index)
    scores = pd.DataFrame(index=df.index)

    if has_sma150:
        passed['sma50_gt_sma150'] = df['SMA50'] > df['SMA150']
        scores['sma50_gt_sma150'] = _tt_rule_score(
            passed['sma50_gt_sma150'], _tt_percent_over(df['SMA50'], df['SMA150']), TT_RULE_NVALUES['sma50_gt_sma150'])
    else:
        passed['sma50_gt_sma150'] = pd.NA
        scores['sma50_gt_sma150'] = 0.0

    if has_sma150 and has_sma200:
        passed['sma150_gt_sma200'] = df['SMA150'] > df['SMA200']
        scores['sma150_gt_sma200'] = _tt_rule_score(
            passed['sma150_gt_sma200'], _tt_percent_over(df['SMA150'], df['SMA200']), TT_RULE_NVALUES['sma150_gt_sma200'])
    else:
        passed['sma150_gt_sma200'] = pd.NA
        scores['sma150_gt_sma200'] = 0.0

    if has_52w_high:
        span_actual = 0.75 * df['price_52_week_high']
        span_threshold = 1.25 * df['price_52_week_low']
        passed['week52_span'] = span_actual > span_threshold
        scores['week52_span'] = _tt_rule_score(
            passed['week52_span'], _tt_percent_over(span_actual, span_threshold), TT_RULE_NVALUES['week52_span'])
    else:
        passed['week52_span'] = pd.NA
        scores['week52_span'] = 0.0

    if 'SMA50' in df.columns and TT_LIQUIDITY_VOLUME_FIELD in df.columns:
        liquidity_value = df['SMA50'] * df[TT_LIQUIDITY_VOLUME_FIELD]
        passed['liquidity'] = liquidity_value >= 20_000_000
        scores['liquidity'] = _tt_rule_score(
            passed['liquidity'], _tt_percent_over(liquidity_value, 20_000_000), TT_RULE_NVALUES['liquidity'])
    else:
        liquidity_value = pd.Series(pd.NA, index=df.index)
        passed['liquidity'] = pd.NA
        scores['liquidity'] = 0.0

    if has_52w_high:
        threshold = 0.75 * df['price_52_week_high']
        passed['close_above_52w_high_25pct'] = df['close'] > threshold
        scores['close_above_52w_high_25pct'] = _tt_rule_score(
            passed['close_above_52w_high_25pct'], _tt_percent_over(df['close'], threshold), TT_RULE_NVALUES['close_above_52w_high_25pct'])
    else:
        passed['close_above_52w_high_25pct'] = pd.NA
        scores['close_above_52w_high_25pct'] = 0.0

    passed['prev_close_above_10'] = df['close'] > 10
    scores['prev_close_above_10'] = _tt_rule_score(
        passed['prev_close_above_10'], _tt_percent_over(df['close'], 10), TT_RULE_NVALUES['prev_close_above_10'])

    if 'SMA50' in df.columns:
        passed['close_above_sma50'] = df['close'] > df['SMA50']
        scores['close_above_sma50'] = _tt_rule_score(
            passed['close_above_sma50'], _tt_percent_over(df['close'], df['SMA50']), TT_RULE_NVALUES['close_above_sma50'])
    else:
        passed['close_above_sma50'] = pd.NA
        scores['close_above_sma50'] = 0.0

    if 'total_revenue_yoy_growth_fq' in df.columns:
        passed['sales_qoq_yoy'] = df['total_revenue_yoy_growth_fq'] > 25
        scores['sales_qoq_yoy'] = _tt_rule_score(
            passed['sales_qoq_yoy'], _tt_percent_over(df['total_revenue_yoy_growth_fq'], 25), TT_RULE_NVALUES['sales_qoq_yoy'])
    else:
        passed['sales_qoq_yoy'] = pd.NA
        scores['sales_qoq_yoy'] = 0.0

    if 'earnings_per_share_diluted_yoy_growth_fq' in df.columns:
        passed['eps_qoq_yoy'] = df['earnings_per_share_diluted_yoy_growth_fq'] > 18
        scores['eps_qoq_yoy'] = _tt_rule_score(
            passed['eps_qoq_yoy'], _tt_percent_over(df['earnings_per_share_diluted_yoy_growth_fq'], 18), TT_RULE_NVALUES['eps_qoq_yoy'])
    else:
        passed['eps_qoq_yoy'] = pd.NA
        scores['eps_qoq_yoy'] = 0.0

    # -- Stage 2: Relative Strength and 200d-MA-Rising, both of which
    # need a year of daily history rather than a single snapshot value.
    # Only pulled for tickers this query already returned -- SPY comes
    # along as the RS benchmark.
    tickers = df['name'].dropna().unique().tolist()
    history = _fetch_rs_history(tickers + ['SPY']) if tickers else {}
    spy_df = history.get('SPY')

    rs_values, sma200_rising = [], []
    for ticker in df['name']:
        rec_rs, rec_slope = None, None
        t_df = history.get(ticker)
        if t_df is not None and spy_df is not None:
            try:
                close = t_df['Close']
                ret = close.pct_change()
                spy_close = spy_df['Close'].reindex(close.index)
                spy_ret = spy_close.pct_change()
                # NOTE: this is a ratio of two daily returns, exactly as
                # the source rule defines it -- inherently noisy on any
                # day SPY's own move is near zero (a big/undefined
                # ratio), which _safe_ratio only guards against for an
                # EXACT zero denominator, not a merely tiny one. That
                # noise is a property of the rule itself, not a bug
                # here.
                ratio = _safe_ratio(ret, spy_ret)
                ema60 = ratio.ewm(span=60, adjust=False, min_periods=60).mean()
                if pd.notna(ema60.iloc[-1]):
                    rec_rs = float(ema60.iloc[-1])
            except Exception:
                rec_rs = None
        if t_df is not None:
            try:
                sma200 = t_df['Close'].rolling(200).mean()
                diffs = sma200.diff().tail(RS_ONE_MONTH_WINDOW)
                if len(diffs) == RS_ONE_MONTH_WINDOW and diffs.notna().all():
                    rec_slope = bool((diffs > 0).all())
            except Exception:
                rec_slope = None
        rs_values.append(rec_rs)
        sma200_rising.append(rec_slope)

    df['Relative_Strength_Value'] = rs_values
    passed['relative_strength'] = df['Relative_Strength_Value'] > 1.0
    scores['relative_strength'] = _tt_rule_score(
        passed['relative_strength'], _tt_percent_over(df['Relative_Strength_Value'], 1.0), TT_RULE_NVALUES['relative_strength'])

    df['SMA200_Rising_21d'] = pd.Series(sma200_rising, index=df.index).astype('boolean')
    passed['sma200_rising_21d'] = df['SMA200_Rising_21d']
    # No natural "percent over threshold" for a streak-consistency check
    # (see the section header) -- contributes to the N-Value Rating only.
    scores['sma200_rising_21d'] = 0.0

    n_value_rating = sum(
        passed[rule].fillna(False).astype(int) * (2 ** TT_RULE_NVALUES[rule])
        for rule in TT_RULE_ORDER
    )
    bucket_rating = scores[TT_RULE_ORDER].sum(axis=1)
    rules_passed = passed[TT_RULE_ORDER].apply(lambda c: c.fillna(False)).sum(axis=1)

    # Rounded before it ever reaches the DataFrame -- not just cosmetic:
    # unrounded floats from these calculations (e.g. SMA50 x volume, or
    # 0.75 x price_52_week_high) often carry tiny binary-float noise
    # (928.5500000000001, not 928.55), and _fit_column sizes a column's
    # width off str(value) whenever its header isn't one it specially
    # recognizes -- so that noise was making these columns dramatically
    # wider than the 2-decimal number actually shown in each cell.
    na_col = pd.Series(pd.NA, index=df.index)
    out = pd.DataFrame({
        'name': df['name'],
        'N_Value_Rating': n_value_rating,
        'Bucket_Rating': bucket_rating.round(2),
        'Rules_Passed': rules_passed,
        'close': _tt_safe_round(df['close'], 2),
        'sector': df.get('sector'),
        'industry': df.get('industry'),
        'market_cap_basic': df.get('market_cap_basic'),
        'SMA50': _tt_safe_round(df['SMA50'] if 'SMA50' in df.columns else na_col, 2),
        'SMA150': _tt_safe_round(df['SMA150'] if has_sma150 else na_col, 2),
        'SMA200': _tt_safe_round(df['SMA200'] if has_sma200 else na_col, 2),
        'price_52_week_high': _tt_safe_round(df['price_52_week_high'] if has_52w_high else na_col, 2),
        'price_52_week_low': _tt_safe_round(df['price_52_week_low'] if 'price_52_week_low' in df.columns else na_col, 2),
        'Relative_Strength_Value': _tt_safe_round(df['Relative_Strength_Value'], 3),
        'Liquidity_Value': _tt_safe_round(liquidity_value, 2),
        'Sales_QoQ_YoY': _tt_safe_round(df['total_revenue_yoy_growth_fq'] if 'total_revenue_yoy_growth_fq' in df.columns else na_col, 2),
        'EPS_QoQ_YoY': _tt_safe_round(df['earnings_per_share_diluted_yoy_growth_fq'] if 'earnings_per_share_diluted_yoy_growth_fq' in df.columns else na_col, 2),
    })
    for rule in TT_RULE_ORDER:
        out[f'Pass_{rule}'] = passed[rule]

    out = out.sort_values(by='Bucket_Rating', ascending=False, na_position='last', kind='stable').reset_index(drop=True)
    print(f"Finished {name}: {len(out)} tickers scored.")
    return name, out


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

# The RS Dashboard tabs (see run_rs_dashboard) -- not a "scan" in the
# match/no-match sense, so they skip the universal fixed-column pipeline
# every scan tab shares (see write_styled_workbook).
RS_DASHBOARD_SHEETS = {'RS_Groups', 'RS_Indices_Sectors'}

# Friendly-name column format buckets
PERCENT_COLUMNS = {
    '1D %', '1W %', '1M %', '3M %', '6M %', '1Y %',
    'EPS YoY Growth % (Qtr)', 'Revenue YoY Growth % (Qtr)',
    '% Above 52W Low', '1M Volatility %',
    # RS Dashboard tabs (RS_Groups / RS_Indices_Sectors) -- these reuse
    # '1D %'/'1M %' above directly since they mean the same thing, and
    # add this one of their own. "RS Thrust" and "1-Mth RRS" are NOT
    # percentages -- see RATIO_COLUMNS below.
    'Off 52W High %',
}
# Abbreviated (K/M/B) dollar formats, per the user's request
ABBREVIATED_CURRENCY_COLUMNS = {'Market Cap', '$ Volume'}
INTEGER_COLUMNS = {'Volume', 'Float'}
PRICE_COLUMNS = {'Price'}
# "RS Thrust" / "1-Mth RRS" are Real Relative Strength readings (see the
# RS DASHBOARD header comment) -- a volatility-adjusted excess move,
# expressed in units of the ticker's own ATR, not a percentage. The
# "x" suffix reads naturally as "2.35 ATRs of excess move."
RATIO_COLUMNS = {'Relative Volume', 'RS Thrust', '1-Mth RRS'}

ABBREVIATED_CURRENCY_FORMAT = '[>=1000000000]$#,##0.0,,,"B";[>=1000000]$#,##0.0,,"M";$#,##0.0,"K"'

# Order in which individual sheets should appear (matches the ticker
# print-out order already used further down in the script)
SHEET_DISPLAY_ORDER = [
    "Mom_1W_Small", "Mom_1M_Small", "Mom_3M_Small", "Mom_6M_Small",
    "Mom_1W_Large", "Mom_1M_Large", "Mom_3M_Large", "Mom_6M_Large",
    "1_Fundamental_Growth", "3_Post_Earnings_Cont_Base",
    "4_Strongest_Stock_JK", "5_Strongest_Stock_10B_Rev_30_JK",
    "Daily_Tightness_Swing", "Leveraged_Setups", "Finviz_High_Short_Float",
    "Finviz_IPO_Weekly", "Finviz_Steve_Jacobs_RS", "Finviz_High_Velocity_Quality"
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


def _hex(color: str) -> str:
    return f"#{color}"


def _get_format(workbook, fmt_cache: dict, key, props: dict):
    """Format objects are meant to be created once and reused in
    xlsxwriter (unlike openpyxl, where a Font/PatternFill is just a
    cheap plain object) -- this cache is what makes that possible
    across the many small format variants (one per number-format x
    zebra/plain combination) every sheet needs."""
    if key not in fmt_cache:
        fmt_cache[key] = workbook.add_format(props)
    return fmt_cache[key]


MIN_COL_WIDTH = 8    # a floor so very-short data (e.g. "MATCH") isn't cramped, and so
                     # single long header words (e.g. "Relative") don't split mid-word
MAX_COL_WIDTH = 45   # a safety ceiling for a truly extreme outlier value; long but normal
                     # text (e.g. a GICS industry name) should still get its real width


def _header_format(workbook, fmt_cache):
    return _get_format(workbook, fmt_cache, ('header',), {
        'bold': True, 'bg_color': _hex(COLOR_HEADER_BG), 'font_color': _hex(COLOR_HEADER_FONT),
        'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        'bottom': 2, 'bottom_color': _hex(COLOR_ACCENT_CYAN),
    })


def _data_format(workbook, fmt_cache, numfmt, zebra: bool):
    props = {'num_format': numfmt} if numfmt else {}
    if zebra:
        props = dict(props, bg_color=_hex(COLOR_ZEBRA))
    return _get_format(workbook, fmt_cache, ('data', numfmt, zebra), props)


def _write_cell(ws, row, col, value, fmt):
    if pd.isna(value):
        ws.write_blank(row, col, None, fmt)
    elif isinstance(value, numbers.Number):
        ws.write_number(row, col, float(value), fmt)
    else:
        ws.write_string(row, col, str(value), fmt)


def _fit_column(ws, workbook, fmt_cache, col_idx, header, values, n_rows, apply_conditional=True):
    """Shared by every tab: data-only column width, the header's wrapped
    line count (for row-height calc), and -- when apply_conditional is
    True -- the color-scale / MATCH-status conditional formatting. Used
    both by the generic per-scan writer and, per-column, by the RS
    Dashboard's custom block writer so the two stay visually identical."""
    data_max_len = max((len(_preview_text(v, header)) for v in values), default=0)
    width = max(MIN_COL_WIDTH, min(data_max_len + 2, MAX_COL_WIDTH))
    ws.set_column(col_idx, col_idx, width)
    header_len = len(str(header)) if header else 0
    lines_needed = max(1, -(-header_len // max(int(width) - 1, 1)))  # ceil division

    if apply_conditional and n_rows > 0:
        if header in PERCENT_COLUMNS:
            ws.conditional_format(1, col_idx, n_rows, col_idx, {
                'type': '3_color_scale',
                'min_color': _hex(COLOR_SCALE_LOW), 'mid_color': _hex(COLOR_SCALE_MID), 'max_color': _hex(COLOR_SCALE_HIGH),
                'min_type': 'min', 'mid_type': 'percentile', 'mid_value': 50, 'max_type': 'max',
            })
        if header == 'Status':
            match_fmt = _get_format(workbook, fmt_cache, ('status', 'match'), {'bg_color': _hex(COLOR_MATCH_FILL)})
            nomatch_fmt = _get_format(workbook, fmt_cache, ('status', 'nomatch'), {'bg_color': _hex(COLOR_NOMATCH_FILL)})
            ws.conditional_format(1, col_idx, n_rows, col_idx, {'type': 'cell', 'criteria': 'equal to', 'value': '"MATCH"', 'format': match_fmt})
            ws.conditional_format(1, col_idx, n_rows, col_idx, {'type': 'cell', 'criteria': 'equal to', 'value': '"NO MATCH"', 'format': nomatch_fmt})

    return lines_needed


def write_metric_sheet(workbook, fmt_cache: dict, sheet_name: str, df: pd.DataFrame):
    """Write one already-finalized (prepare_sheet_df'd) DataFrame as a
    fully styled tab: header styling, freeze panes, autofilter, column
    widths, number formats, zebra striping, and color-scale / status
    conditional formatting. This is the xlsxwriter equivalent of the
    old openpyxl combo of `df.to_excel()` + `style_worksheet()` --
    combined into one pass here because xlsxwriter, unlike openpyxl,
    can't restyle a cell after it's written, so the value and its
    format have to be written together."""
    n_rows, n_cols = len(df), len(df.columns)
    if n_rows == 0 or n_cols == 0:
        return None
    ws = workbook.add_worksheet(sheet_name)
    headers = list(df.columns)
    header_fmt = _header_format(workbook, fmt_cache)
    for c, header in enumerate(headers):
        ws.write(0, c, header, header_fmt)

    max_header_lines = 1
    for c, header in enumerate(headers):
        numfmt = _column_format(header)
        plain_fmt = _data_format(workbook, fmt_cache, numfmt, zebra=False)
        zebra_fmt = _data_format(workbook, fmt_cache, numfmt, zebra=True)
        col_values = df.iloc[:, c].tolist()
        for r, value in enumerate(col_values):
            fmt = zebra_fmt if r % 2 == 0 else plain_fmt
            _write_cell(ws, r + 1, c, value, fmt)
        lines_needed = _fit_column(ws, workbook, fmt_cache, c, header, col_values, n_rows)
        max_header_lines = max(max_header_lines, lines_needed)

    ws.set_row(0, 15 * max_header_lines + 8)
    ws.freeze_panes(1, 1)
    ws.autofilter(0, 0, n_rows, n_cols - 1)
    return ws


# ---------------------------------------------------------------------
# RS Dashboard tabs (RS_Groups / RS_Indices_Sectors -- see
# run_rs_dashboard). These reuse _hex/_get_format/_data_format/
# _write_cell/_column_format so their look matches every other tab, but
# need their own writer for two things write_metric_sheet doesn't do:
# native per-row sparklines, and (Indices_Sectors only) repeating the
# header row once per category block instead of once for the whole
# sheet, so each block can be selected and sorted independently.
# ---------------------------------------------------------------------

RS_TAB_HEADERS = ["Ticker", "Name", "RS Thrust", "1-Mth RRS", "1-Mth Chart",
                   "1-Mth RS", "1D %", "1M %", "Off 52W High %", "Price"]
(RS_COL_TICKER, RS_COL_NAME, RS_COL_THRUST, RS_COL_1MRS, RS_COL_SPARK_PRICE,
 RS_COL_SPARK_RS, RS_COL_1D, RS_COL_1M, RS_COL_OFFHIGH, RS_COL_PRICE) = range(len(RS_TAB_HEADERS))

# add_sparkline's 'range' option must point at real cells on the sheet --
# it can't be fed a literal list of values -- so every row also carries
# two blocks of hidden helper columns (trailing daily close, then
# trailing daily % change) that the visible "1-Mth Chart"/"1-Mth RS"
# cells' sparklines are drawn from.
RS_PRICE_DATA_START_COL = len(RS_TAB_HEADERS)
RS_RS_DATA_START_COL = RS_PRICE_DATA_START_COL + RS_SPARKLINE_WINDOW
RS_TOTAL_COLS = RS_RS_DATA_START_COL + RS_SPARKLINE_WINDOW


def _rs_add_table(ws, workbook, fmt_cache, header_row: int, last_data_row: int,
                   col_offset: int = 0, extra_headers: list = None):
    """Turn this block's range into a native Excel Table so its header
    row gets Excel's own sort/filter dropdown arrows, scoped to just
    this block. A plain autofilter (ws.autofilter(), used by every other
    tab) only supports ONE range per worksheet, which is exactly why
    RS_Indices_Sectors' category sections weren't independently sortable
    before -- a Table object has no such limit, so each category (and
    RS_Groups' one block) gets its own.

    The table's range covers ALL of this block's columns -- the visible
    ones AND the hidden sparkline-source helper columns out at
    RS_PRICE_DATA_START_COL/RS_RS_DATA_START_COL -- not just the
    visible ones. This matters because Excel's own Table sort (the
    dropdown arrows this function adds) only reorders columns that are
    actually part of the table; anything outside its range is left
    exactly where it was. The hidden helper columns are what each row's
    sparklines read from (see _rs_add_sparklines) -- leaving them out of
    the table was the bug: sorting by any visible column reordered the
    visible cells (ticker, RS Thrust, etc.) while every sparkline stayed
    anchored to its original row and kept reading the OLD data still
    sitting there, so charts didn't travel with their own ticker after a
    sort. Including the hidden columns in the same table means an Excel
    sort moves the whole logical row -- visible values and hidden
    sparkline data together -- so whichever ticker lands in a row brings
    its own chart data with it.

    style/banded_* are all turned off so the table only adds the
    header's filter buttons and doesn't paint over the zebra striping
    and color scales already written. Excel requires every table column
    to have its own non-blank, unique header string even when the
    column itself is hidden (see _rs_setup_worksheet) and never shown,
    so the hidden columns get plain placeholder names.

    col_offset/extra_headers: for write_rs_flat_sheet's leading
    identifying columns (e.g. "Category"/"Sub-Theme") that sit in front
    of the usual RS_TAB_HEADERS block -- see that function. Every other
    caller leaves these at their defaults, which reproduces the exact
    table this function always built."""
    extra_headers = extra_headers or []
    header_fmt = _header_format(workbook, fmt_cache)
    columns = [{'header': h, 'header_format': header_fmt} for h in extra_headers]
    columns += [{'header': h, 'header_format': header_fmt} for h in RS_TAB_HEADERS]
    columns += [{'header': f"_PriceData{i + 1}", 'header_format': header_fmt} for i in range(RS_SPARKLINE_WINDOW)]
    columns += [{'header': f"_RSData{i + 1}", 'header_format': header_fmt} for i in range(RS_SPARKLINE_WINDOW)]
    ws.add_table(header_row, 0, last_data_row, col_offset + RS_TOTAL_COLS - 1, {
        'columns': columns,
        'style': None,
        'banded_rows': False,
        'banded_columns': False,
        'first_column': False,
        'last_column': False,
        'autofilter': True,
    })


def _rs_write_row(ws, workbook, fmt_cache, row, row_dict: dict, zebra: bool, col_offset: int = 0):
    """Write one ticker's visible cells (formats matched to the rest of
    the workbook via _column_format, so e.g. 'Price' and 'Off 52W High %'
    look identical to their counterparts on a regular scan tab) plus its
    hidden sparkline-source cells. The sparkline itself is a separate,
    worksheet-level call (_rs_add_sparklines) -- add_sparkline isn't a
    per-cell write.

    col_offset shifts every column this writes by that many columns --
    see write_rs_flat_sheet, whose leading identifying columns (written
    separately, by _rs_write_group_cells) occupy columns [0, col_offset)."""
    zebra_bg = {'bg_color': _hex(COLOR_ZEBRA)} if zebra else {}
    for i, header in enumerate(RS_TAB_HEADERS):
        c = col_offset + i
        if header in ("1-Mth Chart", "1-Mth RS"):
            fmt = _data_format(workbook, fmt_cache, None, zebra)
            ws.write_blank(row, c, None, fmt)  # the sparkline draws over this blank cell
            continue
        value = row_dict.get(header)
        is_missing = value is None or (isinstance(value, float) and pd.isna(value))
        if header == "Ticker":
            fmt = _get_format(workbook, fmt_cache, ('rs', 'ticker', zebra), dict(zebra_bg, bold=True))
            ws.write_string(row, c, str(value), fmt)
        elif is_missing:
            fmt = _get_format(workbook, fmt_cache, ('rs', 'na', zebra),
                               dict(zebra_bg, align='center', font_color='#999999', italic=True))
            ws.write_string(row, c, "n/a", fmt)
        else:
            numfmt = _column_format(header)
            fmt = _data_format(workbook, fmt_cache, numfmt, zebra)
            _write_cell(ws, row, c, value, fmt)

    hidden_fmt = _get_format(workbook, fmt_cache, ('rs', 'hidden'), {'num_format': '0.0000'})
    price_start = col_offset + RS_PRICE_DATA_START_COL
    rs_start = col_offset + RS_RS_DATA_START_COL
    for i, v in enumerate(row_dict.get("_PriceSpark") or []):
        ws.write_number(row, price_start + i, float(v), hidden_fmt)
    for i, v in enumerate(row_dict.get("_RSSpark") or []):
        ws.write_number(row, rs_start + i, float(v), hidden_fmt)


def _rs_write_group_cells(ws, workbook, fmt_cache, row, row_dict: dict, group_cols: list, zebra: bool):
    """write_rs_flat_sheet's leading identifying columns (e.g. "Category"
    for a mega theme, "Sub-Theme" for a sub-theme) -- plain left-aligned
    text cells, styled like the rest of the row (same zebra shading).
    Every one of these is also a real Excel Table column, so it gets its
    own header filter dropdown -- checking just one value there is how a
    caller drills down (e.g. filter Sub_Themes to one mega theme's name,
    or Tickers to one sub-theme's name)."""
    zebra_bg = {'bg_color': _hex(COLOR_ZEBRA)} if zebra else {}
    fmt = _get_format(workbook, fmt_cache, ('rs', 'group', zebra),
                       dict(zebra_bg, align='left', valign='vcenter', indent=1))
    for c, header in enumerate(group_cols):
        value = row_dict.get(header)
        ws.write_string(row, c, "" if value is None else str(value), fmt)


def _rs_add_sparklines(ws, row, row_dict: dict, col_offset: int = 0):
    """Line sparkline for the trailing month of price (a cyan trendline
    with a dark-slate marker dot at the period's highest close) and a
    column sparkline of the trailing month's "1-Mth RRS" reading
    itself, day by day -- see _rs_metrics_for's hist_series comment.
    Unlike every earlier design tried here, RRS is a genuinely signed
    quantity (positive = outperforming what its own volatility would
    predict from SPY's move, negative = underperforming), so the
    histogram is two-tone: slate-blue bars above zero, orange bars
    below -- the same two colors this workbook already uses as the
    high/low ends of every percent column's color scale -- with the
    single tallest (most positive, or least negative) bar in the
    window additionally darkened to dark slate so "today vs. its own
    recent RRS range" still reads at a glance."""
    # show_hidden=True is not optional here: Excel does not plot sparkline
    # source data that lives in hidden rows/columns unless told to (the
    # "Show data in hidden rows and columns" sparkline setting, off by
    # default) -- and the whole point of RS_PRICE_DATA_START_COL/
    # RS_RS_DATA_START_COL is that they're hidden helper columns. Without
    # this, every sparkline on the sheet renders as a blank cell in real
    # Excel even though the XML and the data are otherwise correct.
    price_spark = row_dict.get("_PriceSpark") or []
    rs_spark = row_dict.get("_RSSpark") or []
    price_start = col_offset + RS_PRICE_DATA_START_COL
    rs_start = col_offset + RS_RS_DATA_START_COL
    if len(price_spark) >= 2:
        ws.add_sparkline(row, col_offset + RS_COL_SPARK_PRICE, {
            'range': xlsxwriter.utility.xl_range(row, price_start, row, price_start + len(price_spark) - 1),
            'type': 'line', 'weight': 1.25,
            'series_color': _hex(COLOR_ACCENT_CYAN),
            'markers': False, 'high_point': True,
            'high_color': _hex(COLOR_HEADER_BG),  # same dark slate as the RS histogram's peak bar
            'show_hidden': True,
        })
    if len(rs_spark) >= 2:
        ws.add_sparkline(row, col_offset + RS_COL_SPARK_RS, {
            'range': xlsxwriter.utility.xl_range(row, rs_start, row, rs_start + len(rs_spark) - 1),
            'type': 'column',
            'series_color': _hex(COLOR_SCALE_HIGH),      # positive RRS bars
            'negative_points': True,
            'negative_color': _hex(COLOR_SCALE_LOW),     # negative RRS bars
            'high_point': True, 'high_color': _hex(COLOR_HEADER_BG),
            'show_hidden': True,
        })


def _rs_conditional_format(ws, first_row, last_row, col_offset: int = 0):
    """The same 3-color scale every other tab's percent columns get (see
    _fit_column), applied to just this row range -- the whole sheet for
    RS_Groups, or one category block at a time for RS_Indices_Sectors,
    so each block's coloring is scaled against only its own rows
    rather than the whole tab (a strong sector shouldn't all read as
    "average" just because a stronger one sits elsewhere on the tab)."""
    if last_row < first_row:
        return
    for col_idx in (RS_COL_THRUST, RS_COL_1MRS, RS_COL_1D, RS_COL_1M, RS_COL_OFFHIGH):
        ws.conditional_format(first_row, col_offset + col_idx, last_row, col_offset + col_idx, {
            'type': '3_color_scale',
            'min_color': _hex(COLOR_SCALE_LOW), 'mid_color': _hex(COLOR_SCALE_MID), 'max_color': _hex(COLOR_SCALE_HIGH),
            'min_type': 'min', 'mid_type': 'percentile', 'mid_value': 50, 'max_type': 'max',
        })


def _rs_setup_worksheet(ws, col_offset: int = 0):
    ws.set_column(col_offset + RS_COL_TICKER, col_offset + RS_COL_TICKER, 8)
    ws.set_column(col_offset + RS_COL_NAME, col_offset + RS_COL_NAME, 30)
    ws.set_column(col_offset + RS_COL_THRUST, col_offset + RS_COL_1MRS, 13)
    ws.set_column(col_offset + RS_COL_SPARK_PRICE, col_offset + RS_COL_SPARK_RS, 14)
    ws.set_column(col_offset + RS_COL_1D, col_offset + RS_COL_OFFHIGH, 12)
    ws.set_column(col_offset + RS_COL_PRICE, col_offset + RS_COL_PRICE, 10)
    ws.set_column(col_offset + RS_PRICE_DATA_START_COL, col_offset + RS_TOTAL_COLS - 1, None, None, {'hidden': True})


def _rs_write_block(ws, workbook, fmt_cache, df: pd.DataFrame, start_row: int,
                     col_offset: int = 0, group_cols: list = None) -> int:
    """Write one self-contained mini-table -- its own header row, then
    one data row per ticker with sparklines -- and return the row just
    after the last data row. Writing each category as its own Excel
    Table (its own repeated header row included) is what makes it
    independently sortable: click that header's own filter arrow and
    Excel sorts/filters just this block, without disturbing any other
    category.

    col_offset/group_cols: write_rs_flat_sheet's leading identifying
    columns (see _rs_write_group_cells) -- every other caller leaves
    these at their defaults and gets exactly the block this always
    wrote."""
    group_cols = group_cols or []
    row = start_row + 1  # start_row is reserved for the header; _rs_add_table writes it
    first_data_row = row
    for i, (_, series) in enumerate(df.iterrows()):
        row_dict = series.to_dict()
        zebra = (i % 2 == 0)
        if group_cols:
            _rs_write_group_cells(ws, workbook, fmt_cache, row, row_dict, group_cols, zebra)
        _rs_write_row(ws, workbook, fmt_cache, row, row_dict, zebra=zebra, col_offset=col_offset)
        _rs_add_sparklines(ws, row, row_dict, col_offset=col_offset)
        row += 1
    _rs_conditional_format(ws, first_data_row, row - 1, col_offset=col_offset)
    _rs_add_table(ws, workbook, fmt_cache, start_row, row - 1, col_offset=col_offset, extra_headers=group_cols)
    return row


def write_rs_flat_sheet(workbook, fmt_cache: dict, df: pd.DataFrame, sheet_name: str, group_cols: list = None):
    """A single, flat, fully sortable/filterable Excel Table -- every row
    under one shared header, no per-category banner blocks -- for a
    caller that wants ALL rows visible and sortable together rather than
    split into independent per-category mini-tables (compare
    write_rs_indices_sheet, which deliberately does the opposite).

    group_cols are extra identifying columns (e.g. "Category" for a mega
    theme, "Sub-Theme") rendered as real, sortable, FILTERABLE leading
    data columns -- every column in an Excel Table gets its own header
    dropdown, so checking/unchecking values in a group column's dropdown
    is how a caller drills down. theme_rotation.py uses this for all
    three of its tabs: Mega_Themes (no group_cols -- nothing above a
    mega theme), Sub_Themes (group_cols=["Category"] -- filter to one
    mega theme), and Tickers (group_cols=["Category", "Sub-Theme"] --
    filter to one sub-theme's actual constituents).

    Reuses every bit of scans.py's RS-tab rendering (row formatting, the
    two-tone RRS histogram sparklines, per-block Table, 3-color
    conditional formatting) via the col_offset/group_cols plumbing those
    functions accept -- this function only adds the leading group_cols
    block and default sorts the whole tab (same 1-Mth RRS / RS Thrust
    order as every other RS tab) before writing it as one block."""
    if df is None or df.empty:
        return None
    group_cols = group_cols or []
    col_offset = len(group_cols)
    df = df.sort_values(by=["1-Mth RRS", "RS Thrust"], ascending=False,
                         na_position="last", kind="stable")
    ws = workbook.add_worksheet(sheet_name)
    _rs_setup_worksheet(ws, col_offset=col_offset)
    for i, _ in enumerate(group_cols):
        ws.set_column(i, i, 24)
    ws.freeze_panes(1, col_offset + 2)
    ws.set_row(0, 30)
    _rs_write_block(ws, workbook, fmt_cache, df, start_row=0, col_offset=col_offset, group_cols=group_cols)
    return ws


def write_rs_groups_sheet(workbook, fmt_cache: dict, df: pd.DataFrame):
    """RS_Groups: one flat, self-sortable table (all ~45 tickers, each
    scored independently -- see run_rs_dashboard).

    Default row order: sorted by 1-Mth RRS (then RS Thrust as a
    tiebreaker), descending -- strongest trend-vs-expectation first.
    This is only the file's DEFAULT order -- the sortable header (see
    _rs_add_table) lets it be re-sorted by hand at any time."""
    if df is None or df.empty:
        return None
    df = df.sort_values(by=["1-Mth RRS", "RS Thrust"], ascending=False,
                         na_position="last", kind="stable")
    ws = workbook.add_worksheet('RS_Groups')
    _rs_setup_worksheet(ws)
    ws.freeze_panes(1, 2)
    ws.set_row(0, 30)
    _rs_write_block(ws, workbook, fmt_cache, df, start_row=0)
    return ws


# Categories whose row order is left exactly as configured in
# RS_INDICES_SECTORS rather than default-sorted. Segment is
# deliberately kept in small-to-large-cap order (IJS/IJR/IJT ->
# IJJ/IJH/IJK -> IVE/IVV/IVW) so the progression across cap sizes stays
# visually intact -- sorting it by score would scramble that. Every
# other category is sorted (see write_rs_indices_sheet) by 1-Mth RRS
# first, RS Thrust as the tiebreaker. This is only the file's DEFAULT
# row order -- the sortable header on every block (see _rs_add_table)
# lets it be re-sorted by hand at any time.
RS_UNSORTED_CATEGORIES = {"Segment"}


def write_rs_indices_sheet(workbook, fmt_cache: dict, df: pd.DataFrame, sheet_name: str = 'RS_Indices_Sectors'):
    """RS_Indices_Sectors: one mini-table per category (Index / Segment /
    EW Sector / SPDR Sector), each with its own repeated header row --
    every ticker is scored independently (see run_rs_dashboard), so
    there's no shared ranking universe a category block needs to stay
    separate for; the split is purely for the layout -- separated by a
    merged category banner row instead of a single flat table with a
    'Category' column. Row freezing is column-only (not row-only)
    because the header row's position shifts from block to block, so
    freezing a fixed row wouldn't keep every block's own header in
    view the way it does on RS_Groups.

    sheet_name defaults to this tab's own name for scans.py's own
    three-tab workbook, but this whole function -- layout, sorting,
    sparklines, per-block Excel Table, everything -- is completely
    generic over any DataFrame with this same Category/Ticker/Name/...
    schema. theme_rotation.py reuses it as-is for its own, differently-
    named standalone workbook (its rows are sub-themes, not tickers)
    rather than duplicating this rendering pipeline a second time."""
    if df is None or df.empty:
        return None
    ws = workbook.add_worksheet(sheet_name)
    _rs_setup_worksheet(ws)
    ws.freeze_panes(0, 2)
    category_fmt = _get_format(workbook, fmt_cache, ('rs', 'category'), {
        'bold': True, 'bg_color': _hex(COLOR_HEADER_BG), 'font_color': _hex(COLOR_HEADER_FONT),
        'font_size': 11, 'align': 'left', 'valign': 'vcenter', 'indent': 1,
        'top': 2, 'top_color': _hex(COLOR_ACCENT_CYAN),
    })
    row = 0
    for category in df['Category'].drop_duplicates():
        block = df[df['Category'] == category]
        if category not in RS_UNSORTED_CATEGORIES:
            block = block.sort_values(by=["1-Mth RRS", "RS Thrust"], ascending=False,
                                       na_position="last", kind="stable")
        if row > 0:
            row += 1  # blank separator row between category blocks
        ws.merge_range(row, 0, row, len(RS_TAB_HEADERS) - 1, str(category), category_fmt)
        row += 1
        row = _rs_write_block(ws, workbook, fmt_cache, block, start_row=row)
    return ws


def write_trend_template_sheet(workbook, fmt_cache: dict, df: pd.DataFrame):
    """Trend_Template: the 11-rule N-Value/Bucket-score screener (see
    the TREND TEMPLATE SCAN section header for the full rules and
    scoring, and run_trend_template for how the DataFrame is built).
    Bespoke writer (not write_metric_sheet) because this tab's columns --
    two composite scores, a passed-rule count, eleven PASS/FAIL
    columns, then the raw supporting values -- don't match the fixed
    schema every other scan tab shares; same reasoning as
    RS_Groups/RS_Indices_Sectors having their own writers.

    Default row order: sorted by Bucket Rating descending (see
    run_trend_template) -- the sortable header (autofilter) lets it be
    re-sorted by hand at any time, same as every other tab."""
    if df is None or df.empty:
        return None

    display = pd.DataFrame({
        'Ticker': df['name'],
        'N-Value Score': df['N_Value_Rating'],
        'Bucket Score': df['Bucket_Rating'],
        'Rules Passed': df['Rules_Passed'],
        'Price': df['close'],
        'Sector': df['sector'],
        'Industry': df['industry'],
        'Market Cap': df['market_cap_basic'],
    })
    for r in TT_RULE_ORDER:
        display[TT_RULE_LABELS[r]] = df[f'Pass_{r}'].map(lambda v: 'N/A' if pd.isna(v) else ('PASS' if v else 'FAIL'))
    display['SMA50'] = df['SMA50']
    display['SMA150'] = df['SMA150']
    display['SMA200'] = df['SMA200']
    display['52W High'] = df['price_52_week_high']
    display['52W Low'] = df['price_52_week_low']
    display['Relative Strength (EMA60)'] = df['Relative_Strength_Value']
    display['Liquidity ($)'] = df['Liquidity_Value']
    display['Sales QoQ YoY %'] = df['Sales_QoQ_YoY']
    display['EPS QoQ YoY %'] = df['EPS_QoQ_YoY']

    n_rows, n_cols = len(display), len(display.columns)
    if n_rows == 0 or n_cols == 0:
        return None
    ws = workbook.add_worksheet('Trend_Template')
    headers = list(display.columns)
    header_fmt = _header_format(workbook, fmt_cache)
    for c, header in enumerate(headers):
        ws.write(0, c, header, header_fmt)

    pass_fail_headers = {TT_RULE_LABELS[r] for r in TT_RULE_ORDER}
    percent_headers = {'Sales QoQ YoY %', 'EPS QoQ YoY %'}
    price_headers = {'Price', 'SMA50', 'SMA150', 'SMA200', '52W High', '52W Low'}
    currency_headers = {'Market Cap', 'Liquidity ($)'}
    ratio_headers = {'Relative Strength (EMA60)'}
    integer_headers = {'Rules Passed', 'N-Value Score'}

    def numfmt_for(header):
        if header in pass_fail_headers or header in ('Ticker', 'Sector', 'Industry'):
            return None
        if header == 'Bucket Score':
            return '#,##0.00'
        if header in integer_headers:
            return '#,##0'
        if header in percent_headers:
            return '0.00"%"'
        if header in currency_headers:
            return ABBREVIATED_CURRENCY_FORMAT
        if header in price_headers or header in ratio_headers:
            return '0.00'
        return None

    max_header_lines = 1
    for c, header in enumerate(headers):
        numfmt = numfmt_for(header)
        plain_fmt = _data_format(workbook, fmt_cache, numfmt, zebra=False)
        zebra_fmt = _data_format(workbook, fmt_cache, numfmt, zebra=True)
        col_values = display.iloc[:, c].tolist()
        for r, value in enumerate(col_values):
            fmt = zebra_fmt if r % 2 == 0 else plain_fmt
            _write_cell(ws, r + 1, c, value, fmt)
        lines_needed = _fit_column(ws, workbook, fmt_cache, c, header, col_values, n_rows, apply_conditional=False)
        max_header_lines = max(max_header_lines, lines_needed)

        if header in pass_fail_headers:
            match_fmt = _get_format(workbook, fmt_cache, ('status', 'match'), {'bg_color': _hex(COLOR_MATCH_FILL)})
            nomatch_fmt = _get_format(workbook, fmt_cache, ('status', 'nomatch'), {'bg_color': _hex(COLOR_NOMATCH_FILL)})
            ws.conditional_format(1, c, n_rows, c, {'type': 'cell', 'criteria': 'equal to', 'value': '"PASS"', 'format': match_fmt})
            ws.conditional_format(1, c, n_rows, c, {'type': 'cell', 'criteria': 'equal to', 'value': '"FAIL"', 'format': nomatch_fmt})
        elif header in ('N-Value Score', 'Bucket Score'):
            ws.conditional_format(1, c, n_rows, c, {
                'type': '3_color_scale',
                'min_color': _hex(COLOR_SCALE_LOW), 'mid_color': _hex(COLOR_SCALE_MID), 'max_color': _hex(COLOR_SCALE_HIGH),
                'min_type': 'min', 'mid_type': 'percentile', 'mid_value': 50, 'max_type': 'max',
            })

    ws.set_row(0, 15 * max_header_lines + 8)
    ws.freeze_panes(1, 1)
    ws.autofilter(0, 0, n_rows, n_cols - 1)
    return ws


def write_summary_sheet(workbook, fmt_cache: dict, sheet_counts: dict, sheet_order: list, market_direction: dict = None):
    """Write a first 'Summary' tab: a title, generation timestamp,
    optionally the Market Direction readout (see below), and one row
    per scan tab with its match count and a clickable link that jumps
    straight to that tab. Must be called BEFORE any other sheet is
    added to `workbook` -- unlike openpyxl's `book.create_sheet(
    'Summary', 0)`, xlsxwriter has no way to reorder sheets after the
    fact, so "Summary is tab 1" only works if it's written first.

    market_direction (optional): the dict run_market_direction()
    returns -- {'index_ticker', 'ma21_gt_ma50', 'ma50_rising_21d'}.
    When given, two PASS/FAIL/N-A status lines are shown here,
    purely informational (see the TREND TEMPLATE SCAN section header
    for why this doesn't filter the Trend_Template tab itself)."""
    ws = workbook.add_worksheet('Summary')

    title_fmt = _get_format(workbook, fmt_cache, ('summary', 'title'), {'bold': True, 'font_size': 16, 'font_color': _hex(COLOR_HEADER_BG)})
    subtitle_fmt = _get_format(workbook, fmt_cache, ('summary', 'subtitle'), {'italic': True, 'font_size': 10, 'font_color': '#666666'})
    header_fmt = _get_format(workbook, fmt_cache, ('summary', 'header'), {
        'bold': True, 'bg_color': _hex(COLOR_HEADER_BG), 'font_color': _hex(COLOR_HEADER_FONT), 'align': 'center',
    })
    link_fmt = _get_format(workbook, fmt_cache, ('summary', 'link'), {'font_color': _hex(COLOR_ACCENT_CYAN), 'underline': True, 'bold': True, 'align': 'center'})
    count_fmt = _get_format(workbook, fmt_cache, ('summary', 'count'), {'align': 'center'})
    count_zebra_fmt = _get_format(workbook, fmt_cache, ('summary', 'count_zebra'), {'align': 'center', 'bg_color': _hex(COLOR_ZEBRA)})
    name_zebra_fmt = _get_format(workbook, fmt_cache, ('summary', 'name_zebra'), {'bg_color': _hex(COLOR_ZEBRA)})
    link_zebra_fmt = _get_format(workbook, fmt_cache, ('summary', 'link_zebra'), {
        'font_color': _hex(COLOR_ACCENT_CYAN), 'underline': True, 'bold': True, 'align': 'center', 'bg_color': _hex(COLOR_ZEBRA),
    })
    md_label_fmt = _get_format(workbook, fmt_cache, ('summary', 'md_label'), {'bold': True})
    md_pass_fmt = _get_format(workbook, fmt_cache, ('summary', 'md_pass'), {'bold': True, 'bg_color': _hex(COLOR_MATCH_FILL)})
    md_fail_fmt = _get_format(workbook, fmt_cache, ('summary', 'md_fail'), {'bold': True, 'bg_color': _hex(COLOR_NOMATCH_FILL)})

    ws.write(0, 0, "Trading Scans — Summary", title_fmt)
    ws.write(1, 0, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", subtitle_fmt)

    next_row = 3
    if market_direction:
        idx = market_direction.get('index_ticker', 'SPY')

        def md_status(value):
            if value is None:
                return "N/A", md_label_fmt
            return ("PASS", md_pass_fmt) if value else ("FAIL", md_fail_fmt)

        ws.write(next_row, 0, f"Market Direction ({idx}): 21d MA > 50d MA", md_label_fmt)
        text, fmt = md_status(market_direction.get('ma21_gt_ma50'))
        ws.write(next_row, 1, text, fmt)
        next_row += 1
        ws.write(next_row, 0, f"Market Direction ({idx}): 50d MA Rising (21 sessions)", md_label_fmt)
        text, fmt = md_status(market_direction.get('ma50_rising_21d'))
        ws.write(next_row, 1, text, fmt)
        next_row += 2

    ws.write(next_row, 0, "See the 'All_Scans' tab for every match on a single page (use the Scan filter dropdown).", subtitle_fmt)
    next_row += 2

    header_row = next_row
    for col_idx, label in enumerate(['Scan', 'Matches', 'Open Tab']):
        ws.write(header_row, col_idx, label, header_fmt)

    row = header_row + 1
    for name in sheet_order:
        if name not in sheet_counts:
            continue
        zebra = (row % 2 == 0)
        ws.write(row, 0, name, name_zebra_fmt if zebra else None)
        ws.write_number(row, 1, sheet_counts[name], count_zebra_fmt if zebra else count_fmt)
        ws.write_url(row, 2, f"internal:'{name}'!A1", link_zebra_fmt if zebra else link_fmt, string="Open →")
        row += 1

    ws.set_column(0, 0, 34)
    ws.set_column(1, 1, 12)
    ws.set_column(2, 2, 14)


# Sheets whose Summary-tab "Matches" count should reflect only the rows
# that satisfy the scan's own MATCH/NO MATCH criteria (its 'Status'
# column) rather than every row on the tab. The tab itself is untouched --
# this only changes what number the Summary page reports for it.
SHEETS_COUNT_MATCHES_ONLY = {'Leveraged_Setups'}


def _summary_count(name: str, df: pd.DataFrame) -> int:
    if name in SHEETS_COUNT_MATCHES_ONLY and 'Status' in df.columns:
        return int((df['Status'] == 'MATCH').sum())
    return len(df)


def write_styled_workbook(results_dict: dict, path: str, market_direction: dict = None):
    """Write results_dict to an .xlsx file with all the formatting
    above applied: styled headers, freeze/filter, number formats,
    color-scale + status highlighting, the same fixed metric columns
    on every tab, a Summary tab, and 'All_Scans' as the one-page
    combined view.

    market_direction (optional): passed straight through to
    write_summary_sheet -- see run_market_direction/that function's
    docstring."""
    sheet_counts = {name: (0 if df is None or df.empty else _summary_count(name, df)) for name, df in results_dict.items()}

    # Sheet order: All_Scans (one-page view) and Momentum first, then the
    # RS Dashboard tabs and Trend_Template (right after Summary/All_Scans/
    # Momentum, ahead of the individual scans -- all three are dashboard-
    # style tabs with their own bespoke schema/writer, not "scan matches"),
    # then everything else in the logical scan order, then any leftovers.
    ordered_names = [n for n in ('All_Scans', 'Momentum', 'RS_Groups', 'RS_Indices_Sectors', 'Trend_Template') if n in results_dict]
    ordered_names += [n for n in SHEET_DISPLAY_ORDER if n in results_dict and n not in ordered_names]
    ordered_names += [n for n in results_dict if n not in ordered_names]

    workbook = xlsxwriter.Workbook(path)
    fmt_cache = {}
    try:
        # Summary must be written FIRST -- xlsxwriter has no sheet-reorder
        # capability, so "Summary is tab 1" only holds if nothing else is
        # added to the workbook before it.
        write_summary_sheet(workbook, fmt_cache, sheet_counts, ordered_names, market_direction=market_direction)

        for sheet_name in ordered_names:
            df = results_dict[sheet_name]
            if df is None or df.empty:
                continue
            if sheet_name == 'RS_Groups':
                write_rs_groups_sheet(workbook, fmt_cache, df)
            elif sheet_name == 'RS_Indices_Sectors':
                write_rs_indices_sheet(workbook, fmt_cache, df)
            elif sheet_name == 'Trend_Template':
                write_trend_template_sheet(workbook, fmt_cache, df)
            else:
                styled_df = prepare_sheet_df(df, sheet_name)
                write_metric_sheet(workbook, fmt_cache, sheet_name, styled_df)
    finally:
        workbook.close()


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
        rs_future = executor.submit(run_rs_dashboard)
        tt_future = executor.submit(run_trend_template)
        md_future = executor.submit(run_market_direction)

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

        # RS Dashboard tabs stay out of all_collected_dfs/All_Scans -- they
        # aren't "matches" from a scan and don't share the other tabs'
        # column schema, so folding them in would just produce a mess of
        # mismatched columns on the combined view.
        rs_groups_df, rs_indices_df = rs_future.result()
        results_dict['RS_Groups'] = rs_groups_df
        results_dict['RS_Indices_Sectors'] = rs_indices_df

        # Trend_Template likewise stays out of all_collected_dfs/All_Scans --
        # its own bespoke N-Value/Bucket-score schema doesn't match the
        # other tabs' columns either. Market Direction isn't a tab at all,
        # just two status lines write_summary_sheet shows on the Summary tab.
        tt_name, tt_df = tt_future.result()
        results_dict[tt_name] = tt_df
        market_direction = md_future.result()

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
            "Finviz_IPO_Weekly", "Finviz_Steve_Jacobs_RS", "Finviz_High_Velocity_Quality"
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
        write_styled_workbook(results_dict, excel_path, market_direction)
        print(f"Done! Saved cleanly to: {excel_path}")
    except PermissionError:
        alt_path = r"C:\TradingScans\My_Scans_NEW.xlsx"
        write_styled_workbook(results_dict, alt_path, market_direction)
        print(f"\n[WARNING] Excel file was locked. Saved master file to: {alt_path}")