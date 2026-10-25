#!/usr/bin/env python3
"""Fetch market data, three fiscal years of financial statements and peer data from Yahoo Finance
and write them into an Excel file that the Equity Valuation Platform imports directly
(Import & export -> Import file).

Usage (on your own computer, with internet):
    pip3 install yfinance openpyxl
    python3 fetch_yahoo.py BIRD --exchange IDX --peers WEHA LRNA BPTR JAYA
    python3 fetch_yahoo.py AAPL --exchange NASDAQ --peers MSFT GOOGL

Output: <TICKER>_yahoo_import_<date>.xlsx with the sheets Market, Financials, Peers (+ Notes).
Amounts are converted to million units of the reporting currency; percentages are decimals.
Fields Yahoo does not provide (risk-free rate, equity risk premium, unlevered beta, debt spread,
lease liabilities if not reported, forecast assumptions) are NOT written, so the values already in
the platform stay unchanged. Yahoo data are not audited: check key lines against the annual report.
"""
import argparse, datetime as dt, math, sys

SUFFIX = {"IDX": ".JK", "BEI": ".JK", "NYSE": "", "NASDAQ": "", "AMEX": "", "LSE": ".L", "XETRA": ".DE", "FWB": ".F",
          "SGX": ".SI", "ASX": ".AX", "HKEX": ".HK", "TSE": ".T", "BIST": ".IS", "SET": ".BK", "KLSE": ".KL",
          "BURSA": ".KL", "NSE": ".NS", "BSE": ".BO", "SIX": ".SW", "TSX": ".TO", "EURONEXT": ".PA"}
MN = 1e6

def num(x):
    try:
        v = float(x)
        return None if math.isnan(v) or math.isinf(v) else v
    except (TypeError, ValueError):
        return None

def row(df, *labels):
    """Return {year: value} for the first label present in a yfinance statement (columns = period end dates)."""
    if df is None or getattr(df, "empty", True):
        return {}
    for lab in labels:
        if lab in df.index:
            out = {}
            for col, v in df.loc[lab].items():
                y = str(getattr(col, "year", str(col)[:4]))
                if num(v) is not None:
                    out[y] = num(v)
            return out
    return {}

def years_of(*dfs, n=3):
    ys = set()
    for df in dfs:
        if df is not None and not getattr(df, "empty", True):
            ys |= {str(getattr(c, "year", str(c)[:4])) for c in df.columns}
    return sorted(ys)[-n:]

def price_on(tk, date):
    """Last close on or before the valuation date."""
    end = dt.date.fromisoformat(date) + dt.timedelta(days=1)
    h = tk.history(start=end - dt.timedelta(days=14), end=end, auto_adjust=False)
    if h is None or h.empty:
        return None, None
    last = h.iloc[-1]
    return num(last["Close"]), str(h.index[-1])[:10]

def safe_info(tk):
    try: return tk.info or {}
    except Exception: return {}

def fi_shares(tk):
    try: return fi_shares(tk)
    except Exception: return None

def stmt_shares(bal, inc):
    """Shares outstanding from the statements (works when Yahoo's quote-summary 'info' is refused)."""
    for df, labels in ((bal, ("Ordinary Shares Number", "Share Issued")), (inc, ("Basic Average Shares", "Diluted Average Shares"))):
        d = row(df, *labels)
        if d: return d[max(d)]
    return None

def statements(tk):
    return tk.income_stmt, tk.balance_sheet, tk.cashflow

def fin_block(tk, years, notes, tkname):
    inc, bal, cf = statements(tk)
    def m(d, sign=1):  # to million, optional sign flip (capex / dividends are negative on Yahoo)
        return [None if d.get(y) is None else sign * d[y] / MN for y in years]
    def raw(d):
        return [d.get(y) for y in years]
    da = row(cf, "Depreciation And Amortization", "Depreciation Amortization Depletion") or row(inc, "Reconciled Depreciation")
    debt = row(bal, "Total Debt"); leases = row(bal, "Capital Lease Obligations")
    bank = {y: debt[y] - leases.get(y, 0.0) for y in debt}
    ca = row(bal, "Current Assets"); cl = row(bal, "Current Liabilities")
    cash_sti = row(bal, "Cash Cash Equivalents And Short Term Investments", "Cash And Cash Equivalents")
    cdebt = row(bal, "Current Debt And Capital Lease Obligation", "Current Debt")
    nwc = {y: (ca[y] - cash_sti.get(y, 0.0)) - (cl[y] - cdebt.get(y, 0.0)) for y in ca if y in cl}
    items = {
        "revenue": m(row(inc, "Total Revenue", "Operating Revenue")),
        "ebit": m(row(inc, "Operating Income", "EBIT")),
        "da": m(da),
        "otherIncome": m(row(inc, "Other Non Operating Income Expenses", "Other Income Expense")),
        "financeCharges": m({k: abs(v) for k, v in row(inc, "Interest Expense", "Interest Expense Non Operating").items()}),
        "pbt": m(row(inc, "Pretax Income")),
        "taxExpense": m(row(inc, "Tax Provision")),
        "netIncome": m(row(inc, "Net Income Including Noncontrolling Interests", "Net Income")),
        "niParent": m(row(inc, "Net Income Common Stockholders", "Net Income")),
        "eps": raw(row(inc, "Basic EPS", "Diluted EPS")),
        "cfo": m(row(cf, "Operating Cash Flow")),
        "capex": m(row(cf, "Capital Expenditure", "Purchase Of PPE"), -1),
        "disposalProceeds": m(row(cf, "Sale Of PPE", "Sale Of Property Plant And Equipment")),
        "dividends": m(row(cf, "Cash Dividends Paid", "Common Stock Dividend Paid"), -1),
        "totalAssets": m(row(bal, "Total Assets")),
        "totalLiabilities": m(row(bal, "Total Liabilities Net Minority Interest")),
        "equityTotal": m(row(bal, "Total Equity Gross Minority Interest")),
        "equityParent": m(row(bal, "Stockholders Equity", "Common Stock Equity")),
        "nci": m(row(bal, "Minority Interest")),
        "cash": m(row(bal, "Cash And Cash Equivalents", "Cash Cash Equivalents And Short Term Investments")),
        "bankLoans": m(bank),
        "leases": m(leases),
        "opNWC": m(nwc),
    }
    # Only write rows that Yahoo actually delivered; empty rows would erase what is in the platform.
    keep = {k: v for k, v in items.items() if any(x is not None for x in v)}
    if "nci" not in keep and "equityTotal" in keep and "equityParent" in keep:
        keep["nci"] = [None if a is None or b is None else a - b for a, b in zip(keep["equityTotal"], keep["equityParent"])]
    missing = [k for k in items if k not in keep]
    if missing:
        notes.append(f"{tkname}: Yahoo has no data for {', '.join(missing)} - fill these from the annual report (the platform keeps its current values).")
    return keep

def growth(d, years):
    if len(years) < 2: return None
    a, b = d.get(years[-2]), d.get(years[-1])
    return None if a in (None, 0) or b is None else (b - a) / abs(a)

PEER_REQUIRED = ("price", "sharesMn", "eps", "bvps", "debt", "cash")

def peer_row(sym, ticker, date, notes):
    """One peer row; every field falls back to a second Yahoo source so that small IDX companies are covered."""
    import yfinance as yf
    tk = yf.Ticker(sym)
    info = safe_info(tk)
    price, pdate = price_on(tk, date)
    if price is None: price = num(info.get("currentPrice")) or num(info.get("regularMarketPrice")) or num(info.get("previousClose"))
    inc, bal, _ = statements(tk)
    ys = years_of(inc, bal, n=2)
    def latest(d):   # value of the latest year that has one
        if not d: return None
        if ys and ys[-1] in d: return d[ys[-1]]
        return d[max(d)]
    shares = num(info.get("sharesOutstanding")) or fi_shares(tk) or num(info.get("impliedSharesOutstanding")) or stmt_shares(bal, inc)
    eq = latest(row(bal, "Stockholders Equity", "Common Stock Equity", "Total Equity Gross Minority Interest"))
    rev, ni = row(inc, "Total Revenue", "Operating Revenue"), row(inc, "Net Income Common Stockholders", "Net Income", "Net Income From Continuing Operation Net Minority Interest")
    eps = latest(row(inc, "Basic EPS", "Diluted EPS"))
    if eps is None and shares and latest(ni) is not None: eps = latest(ni) / shares
    if eps is None: eps = num(info.get("trailingEps")) or num(info.get("epsTrailingTwelveMonths"))
    debt = latest(row(bal, "Total Debt"))
    if debt is None:
        parts = [latest(row(bal, k)) for k in ("Current Debt And Capital Lease Obligation", "Long Term Debt And Capital Lease Obligation")]
        if all(p is None for p in parts): parts = [latest(row(bal, k)) for k in ("Current Debt", "Long Term Debt", "Current Capital Lease Obligation", "Long Term Capital Lease Obligation")]
        if any(p is not None for p in parts): debt = sum(p or 0.0 for p in parts)
    if debt is None and num(info.get("totalDebt")) is not None: debt = num(info.get("totalDebt"))
    if debt is None and latest(row(bal, "Total Liabilities Net Minority Interest")) is not None:
        debt = 0.0; notes.append(f"Peer {ticker}: no borrowings reported on Yahoo's balance sheet - debt set to 0; check the annual report.")
    cash = latest(row(bal, "Cash And Cash Equivalents", "Cash Cash Equivalents And Short Term Investments", "Cash Financial"))
    if cash is None: cash = num(info.get("totalCash"))
    bvps = eq / shares if eq and shares else num(info.get("bookValue"))
    r = {"ticker": ticker, "name": info.get("longName") or info.get("shortName") or ticker, "price": price, "priceDate": pdate or date,
         "sharesMn": shares / MN if shares else None, "eps": eps, "bvps": bvps,
         "revGrowth": growth(rev, ys), "niGrowth": growth(ni, ys), "debt": debt / MN if debt is not None else None,
         "cash": cash / MN if cash is not None else None, "evEbitda": num(info.get("enterpriseToEbitda")),
         "audited": False, "closeModel": True, "source": f"Yahoo Finance via yfinance, FY{ys[-1] if ys else '?'}, retrieved {dt.date.today()}"}
    if r["evEbitda"] is None:   # EV / EBITDA from the statements when Yahoo's ratio is not available
        ebitda = latest(row(inc, "EBITDA", "Normalized EBITDA"))
        if ebitda and ebitda > 0 and price and shares and debt is not None and cash is not None:
            r["evEbitda"] = round((price * shares + debt - cash) / ebitda, 2)
    gaps = [k for k in PEER_REQUIRED if r.get(k) is None]
    r["complete"] = not gaps
    if gaps: notes.append(f"Peer {ticker}: Yahoo has no {', '.join(gaps)} - not usable as a peer.")
    fc, pc = info.get("financialCurrency"), info.get("currency")
    if fc and pc and fc != pc: notes.append(f"Peer {ticker}: statements in {fc} but price in {pc} - convert EPS, BVPS, debt and cash before use.")
    return r

PEER_COLS = ["ticker", "name", "price", "priceDate", "sharesMn", "eps", "bvps", "revGrowth", "niGrowth", "debt", "cash",
             "evEbitda", "audited", "closeModel", "source"]

def collect(ticker, exchange, peers, date):
    """Retrieve everything from Yahoo and return a payload dict (also used by data_connector.py)."""
    import yfinance as yf
    ex = exchange.upper(); sfx = SUFFIX.get(ex, "")
    sym = ticker.upper() + sfx
    notes = []
    tk = yf.Ticker(sym)
    info = safe_info(tk)
    price, pdate = price_on(tk, date)
    if not info and price is None:
        raise RuntimeError(f"Yahoo returned nothing for {sym}. Check the ticker and the exchange.")
    inc, bal, cf = statements(tk)
    years = years_of(inc, bal, cf, n=3)
    if len(years) < 3: notes.append(f"Yahoo has only {len(years)} fiscal years for {sym}; the platform needs 3.")
    fin = fin_block(tk, years, notes, sym)
    shares = num(info.get("sharesOutstanding")) or fi_shares(tk) or stmt_shares(bal, inc)
    tax = None
    if "taxExpense" in fin and "pbt" in fin and fin["pbt"][-1]:
        tax = fin["taxExpense"][-1] / fin["pbt"][-1]
    fc, pc = info.get("financialCurrency"), info.get("currency")
    if fc and pc and fc != pc:
        notes.append(f"WARNING: statements are in {fc} but the share price is in {pc}. Convert one of them before valuing.")
    src = f"Yahoo Finance ({sym}) via yfinance, retrieved {dt.datetime.now():%Y-%m-%d %H:%M}"
    meta = {"company": info.get("longName") or info.get("shortName") or ticker.upper(), "ticker": ticker.upper(), "exchange": ex, "currency": fc or pc,
            "country": info.get("country"), "yahooIndustry": info.get("industry"), "yahooSector": info.get("sector")}
    market = {"price": price, "priceDate": pdate, "sharesMn": shares / MN if shares else None, "betaReg": num(info.get("beta")), "tax": tax}
    for k, v in list(market.items()):
        if v is None:
            notes.append(f"{sym}: no Yahoo value for {k} - keep or enter it manually."); del market[k]
    notes.append("Not on Yahoo, enter yourself: risk-free rate, equity risk premium, unlevered beta (Damodaran), debt spread. "
                 "Tax = effective rate of the last year; replace it by the statutory rate if that is more representative.")
    prow = []
    for p in peers:
        try:
            prow.append(peer_row(p.upper() + sfx, p.upper(), date, notes))
        except Exception as e:  # keep going if one peer fails
            notes.append(f"Peer {p}: could not be retrieved ({e}).")
    if peers:
        notes.append("Peers: 'closeModel' is set to TRUE and 'audited' to FALSE for new peers; adjust them on the Peer data tab.")
    return {"meta": meta, "market": market, "marketSource": src, "valuationDate": date,
            "financials": {"years": years, "items": fin, "source": src, "lineSource": {k: "Yahoo" for k in fin}},
            "peers": prow, "notes": notes}

def write_xlsx(P, out):
    from openpyxl import Workbook
    wb = Workbook(); ws = wb.active; ws.title = "Market"
    ws.append(["key", "value", "source"])
    for k, v in P["meta"].items():
        if v is not None and k in ("company", "ticker", "exchange", "currency"): ws.append([k, v, P["marketSource"]])
    ws.append(["valuationDate", P["valuationDate"], ""])
    for k, v in P["market"].items(): ws.append([k, v, P["marketSource"]])
    wf = wb.create_sheet("Financials"); wf.append(["item"] + P["financials"]["years"])
    for k, v in P["financials"]["items"].items(): wf.append([k] + v)
    if P["peers"]:
        wp = wb.create_sheet("Peers"); wp.append(PEER_COLS)
        for r in P["peers"]: wp.append([r[c] for c in PEER_COLS])
    wn = wb.create_sheet("Notes"); wn.append(["Notes"])
    for n in P["notes"]: wn.append([n])
    wb.save(out)

def build(ticker, exchange, peers, date, out):
    P = collect(ticker, exchange, peers, date)
    write_xlsx(P, out)
    return P["notes"]

def main():
    ap = argparse.ArgumentParser(description="Fetch Yahoo Finance data into an import file for the Equity Valuation Platform.")
    ap.add_argument("ticker", help="ticker without suffix, e.g. BIRD or AAPL")
    ap.add_argument("--exchange", default="IDX", help="exchange code as used in the platform (IDX, NYSE, NASDAQ, LSE, XETRA, ...)")
    ap.add_argument("--peers", nargs="*", default=[], help="peer tickers, e.g. WEHA LRNA BPTR JAYA")
    ap.add_argument("--date", default=str(dt.date.today()), help="valuation date YYYY-MM-DD (price = last close on or before)")
    ap.add_argument("--out", default=None, help="output file name")
    a = ap.parse_args()
    out = a.out or f"{a.ticker.upper()}_yahoo_import_{a.date}.xlsx"
    try:
        notes = build(a.ticker, a.exchange, a.peers, a.date, out)
    except SystemExit:
        raise
    except Exception as e:
        sys.exit(f"Could not retrieve data from Yahoo Finance ({type(e).__name__}: {e}).\n"
                 "Check the internet connection, the ticker and the exchange, update yfinance (pip3 install -U yfinance) and try again in a few minutes.")
    print(f"Written {out}. Import it in the platform: Import & export -> Import file.")
    for n in notes: print(" -", n)

if __name__ == "__main__":
    main()
