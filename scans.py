import concurrent.futures
from tradingview_screener import Query, col
import pandas as pd
import yfinance as yf
from finvizfinance.screener.overview import Overview
import logging

# 1. Individual standalone operational & tightness scan tabs
individual_scans = {
    "1_Fundamental_Growth": (
        Query().set_markets('america')
        .select('name', 'close', 'change', 'volume', 'market_cap_basic')
        .where(
            col('type') == 'stock',
            col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']),
            col('market_cap_basic') > 300_000_000,
            col('average_volume_60d_calc') > 300_000,
            col('float_shares_outstanding') < 100_000_000,
            col('earnings_per_share_diluted_yoy_growth_fq') > 25,
            col('free_cash_flow_yoy_growth_ttm') > 25,
            col('total_revenue_yoy_growth_fq') > 25
        ).order_by('change', ascending=False).limit(300)
    ),
    "3_Post_Earnings_Cont_Base": (
        Query().set_markets('america')
        .select('name', 'close', 'change', 'volume', 'market_cap_basic')
        .where(
            col('type') == 'stock',
            col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']),
            col('close') > col('SMA20'),
            col('market_cap_basic') > 50_000_000,
            col('average_volume_60d_calc') > 250_000,
            col('relative_volume_10d_calc') >= 2,
            col('float_shares_outstanding') < 50_000_000,
            col('gap') > 5
        ).order_by('change', ascending=False).limit(300)
    ),
    "4_Strongest_Stock_JK": (
        Query().set_markets('america')
        .select('name', 'close', 'change', 'volume', 'market_cap_basic', 'price_52_week_low', 'SMA10', 'SMA50', 'earnings_per_share_diluted_yoy_growth_fq', 'total_revenue_yoy_growth_fq')
        .where(
            col('type') == 'stock',
            col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']),
            col('market_cap_basic').between(300_000_000, 10_000_000_000),
            col('average_volume_60d_calc') > 500_000,
            col('Volatility.M') > 3,
            col('float_shares_outstanding') < 50_000_000,
            col('earnings_per_share_diluted_yoy_growth_fq') > 25,
            col('total_revenue_yoy_growth_fq') > 25,
            col('close') > col('SMA50')
        ).order_by('change', ascending=False).limit(300)
    ),
    "5_Strongest_Stock_10B_Rev_30_JK": (
        Query().set_markets('america')
        .select('name', 'close', 'change', 'volume', 'market_cap_basic', 'price_52_week_low', 'SMA10', 'SMA50', 'earnings_per_share_diluted_yoy_growth_fq', 'total_revenue_yoy_growth_fq')
        .where(
            col('type') == 'stock',
            col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']),
            col('market_cap_basic') > 10_000_000_000,
            col('average_volume_60d_calc') > 500_000,
            col('Volatility.M') > 2,
            col('float_shares_outstanding') < 150_000_000,
            col('earnings_per_share_diluted_yoy_growth_fq') > 25,
            col('total_revenue_yoy_growth_fq') > 25,
            col('close') > col('SMA50')
        ).order_by('change', ascending=False).limit(300)
    ),
    "Daily_Tightness_Swing": (
        Query().set_markets('america')
        .select('name', 'close', 'change', 'volume', 'market_cap_basic', 'price_52_week_low', 'EMA5', 'SMA10', 'SMA20', 'Volatility.M', 'Perf.W')
        .where(
            col('type') == 'stock',
            col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']),
            col('market_cap_basic') > 300_000_000,
            col('average_volume_60d_calc') > 300_000,
            col('volume') > 100_000,
            col('float_shares_outstanding') < 50_000_000,
            col('Volatility.M') > 3.5,
            col('Perf.W') < 5
        ).order_by('change', ascending=False).limit(300)
    )
}

# 2. All 8 Momentum Scans
momentum_scans = {
    "Mom_1W_Small": {
        "mcap_group": "$300M - $10B", "timeframe": "1 Week", "is_large": False,
        "query": Query().set_markets('america').select('name', 'close', 'change', 'volume', 'market_cap_basic', 'price_52_week_low', 'SMA10', 'Perf.W', 'Volatility.M')
        .where(col('type') == 'stock', col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']), col('market_cap_basic').between(300_000_000, 10_000_000_000), col('average_volume_60d_calc') > 300_000, col('volume') > 100_000, col('float_shares_outstanding') < 50_000_000, col('Volatility.M') > 3, col('Perf.W') > 20).order_by('change', ascending=False).limit(300)
    },
    "Mom_1M_Small": {
        "mcap_group": "$300M - $10B", "timeframe": "1 Month", "is_large": False,
        "query": Query().set_markets('america').select('name', 'close', 'change', 'volume', 'market_cap_basic', 'price_52_week_low', 'SMA10', 'Perf.1M', 'Volatility.M')
        .where(col('type') == 'stock', col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']), col('market_cap_basic').between(300_000_000, 10_000_000_000), col('average_volume_60d_calc') > 300_000, col('volume') > 100_000, col('float_shares_outstanding') < 50_000_000, col('Volatility.M') > 3, col('Perf.1M') > 30).order_by('change', ascending=False).limit(300)
    },
    "Mom_3M_Small": {
        "mcap_group": "$300M - $10B", "timeframe": "3 Months", "is_large": False,
        "query": Query().set_markets('america').select('name', 'close', 'change', 'volume', 'market_cap_basic', 'price_52_week_low', 'SMA10', 'Perf.3M', 'Volatility.M')
        .where(col('type') == 'stock', col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']), col('market_cap_basic').between(300_000_000, 10_000_000_000), col('average_volume_60d_calc') > 300_000, col('volume') > 100_000, col('float_shares_outstanding') < 50_000_000, col('Volatility.M') > 3, col('Perf.3M') > 70).order_by('change', ascending=False).limit(300)
    },
    "Mom_6M_Small": {
        "mcap_group": "$300M - $10B", "timeframe": "6 Months", "is_large": False,
        "query": Query().set_markets('america').select('name', 'close', 'change', 'volume', 'market_cap_basic', 'price_52_week_low', 'SMA10', 'Perf.6M', 'Volatility.M')
        .where(col('type') == 'stock', col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']), col('market_cap_basic').between(300_000_000, 10_000_000_000), col('average_volume_60d_calc') > 300_000, col('volume') > 100_000, col('float_shares_outstanding') < 50_000_000, col('Volatility.M') > 3, col('Perf.6M') > 100).order_by('change', ascending=False).limit(300)
    },
    "Mom_1W_Large": {
        "mcap_group": "> $10B", "timeframe": "1 Week", "is_large": True,
        "query": Query().set_markets('america').select('name', 'close', 'change', 'volume', 'market_cap_basic', 'price_52_week_low', 'SMA10', 'Perf.W')
        .where(col('type') == 'stock', col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']), col('market_cap_basic') > 10_000_000_000, col('average_volume_60d_calc') > 300_000, col('volume') > 100_000, col('float_shares_outstanding') < 150_000_000, col('Perf.W') > 20).order_by('change', ascending=False).limit(300)
    },
    "Mom_1M_Large": {
        "mcap_group": "> $10B", "timeframe": "1 Month", "is_large": True,
        "query": Query().set_markets('america').select('name', 'close', 'change', 'volume', 'market_cap_basic', 'price_52_week_low', 'SMA10', 'Perf.1M')
        .where(col('type') == 'stock', col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']), col('market_cap_basic') > 10_000_000_000, col('average_volume_60d_calc') > 300_000, col('volume') > 100_000, col('float_shares_outstanding') < 150_000_000, col('Perf.1M') > 30).order_by('change', ascending=False).limit(300)
    },
    "Mom_3M_Large": {
        "mcap_group": "> $10B", "timeframe": "3 Months", "is_large": True,
        "query": Query().set_markets('america').select('name', 'close', 'change', 'volume', 'market_cap_basic', 'price_52_week_low', 'SMA10', 'Perf.3M')
        .where(col('type') == 'stock', col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']), col('market_cap_basic') > 10_000_000_000, col('average_volume_60d_calc') > 300_000, col('volume') > 100_000, col('float_shares_outstanding') < 150_000_000, col('Perf.3M') > 70).order_by('change', ascending=False).limit(300)
    },
    "Mom_6M_Large": {
        "mcap_group": "> $10B", "timeframe": "6 Months", "is_large": True,
        "query": Query().set_markets('america').select('name', 'close', 'change', 'volume', 'market_cap_basic', 'price_52_week_low', 'SMA10', 'Perf.6M')
        .where(col('type') == 'stock', col('exchange').isin(['NASDAQ', 'NYSE', 'AMEX']), col('market_cap_basic') > 10_000_000_000, col('average_volume_60d_calc') > 300_000, col('volume') > 100_000, col('float_shares_outstanding') < 150_000_000, col('Perf.6M') > 100).order_by('change', ascending=False).limit(300)
    }
}

# 3. Finviz Scan Configuration
finviz_scans = {
    "Finviz_High_Short_Float": {
        'Market Cap.': '+Small (over $300mln)',
        'Average Volume': 'Over 1M',
        'Float': 'Under 100M',
        'Float Short': 'Over 30%'
    },
    "Finviz_IPO_Weekly": {
        'Market Cap.': '+Mid (over $2bln)',
        'EPS growthnext year': 'Positive (>0%)',
        'Average Volume': 'Over 1M',
        'IPO Date': 'In the last year'
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
    name = "Leveraged_Setups"
    results = []

    if not leveraged_tickers:
        return results

    logger = yf.utils.get_yf_logger()
    original_level = logger.level
    logger.setLevel(logging.CRITICAL)

    try:
        # Download ALL tickers at once in a single network request
        data = yf.download(leveraged_tickers, period="3mo", progress=False, group_by='ticker')

        # Loop through the downloaded data per ticker
        for t in leveraged_tickers:
            try:
                if len(leveraged_tickers) == 1:
                    df = data
                else:
                    df = data[t] if t in data.columns.levels[0] else pd.DataFrame()

                df = df.dropna(subset=['Close', 'Volume'])

                if not df.empty and len(df) >= 60:
                    recent_60 = df.tail(60)
                    close = recent_60['Close'].squeeze()
                    volume = recent_60['Volume'].squeeze()
                    high = recent_60['High'].squeeze()
                    low = recent_60['Low'].squeeze()

                    avg_dollar_vol = (volume * close).mean()
                    adr_pct = ((high - low) / close * 100).mean()
                    latest_close = float(close.iloc[-1])
                    latest_change = float(((latest_close - float(close.iloc[-2])) / float(close.iloc[-2])) * 100)

                    meets_criteria = (avg_dollar_vol >= 100_000_000) and (adr_pct > 1.5)
                    status = "MATCH" if meets_criteria else "NO MATCH"

                    results.append({
                        'name': t,
                        'Status': status,
                        'close': round(latest_close, 2),
                        'change': round(latest_change, 2),
                        'avg_dollar_volume': round(avg_dollar_vol, 2),
                        'adr_pct': round(adr_pct, 2)
                    })
            except Exception:
                continue

    except Exception:
        pass
    finally:
        # Always restore the logger level no matter what happens
        logger.setLevel(original_level)

    out_df = pd.DataFrame(results)
    if not out_df.empty:
        out_df.insert(0, 'Source_Scan', name)
        out_df = out_df.sort_values(by=['avg_dollar_volume'], ascending=[False], na_position='last')

    return name, out_df

def run_individual(name, query):
    try:
        _, df = query.get_scanner_data()
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
    try:
        _, df = info['query'].get_scanner_data()
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

def run_finviz_scan(name, filters_dict):
    try:
        foverview = Overview()
        foverview.set_filter(filters_dict=filters_dict)
        df = foverview.screener_view()
        
        if df is not None and not df.empty:
            if 'Ticker' in df.columns:
                df.rename(columns={'Ticker': 'name'}, inplace=True)
            df.insert(0, 'Source_Scan', name)
            return name, df
    except Exception as e:
        print(f"Error in Finviz scan {name}: {e}")
    return name, pd.DataFrame()

if __name__ == "__main__":
    print("Executing all scans concurrently...")
    results_dict = {}
    momentum_dfs = []
    all_collected_dfs = []

    with concurrent.futures.ThreadPoolExecutor(max_workers=14) as executor:
        ind_futures = {executor.submit(run_individual, name, q): name for name, q in individual_scans.items()}
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
        with pd.ExcelWriter(excel_path) as writer:
            for sheet_name, df in results_dict.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Done! Saved cleanly to: {excel_path}")
    except PermissionError:
        alt_path = r"C:\TradingScans\My_Scans_NEW.xlsx"
        with pd.ExcelWriter(alt_path) as writer:
            for sheet_name, df in results_dict.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"\n[WARNING] Excel file was locked. Saved master file to: {alt_path}")