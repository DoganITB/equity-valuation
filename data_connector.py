#!/usr/bin/env python3
"""Local data connector for the Equity Valuation Platform.

A browser page may not download data from Yahoo Finance or idx.co.id itself (the sites block it). This small
server runs on your own computer, fetches the data in Python and hands it to the platform, which then fills
Market & company, Financial history and Peer data with one click ("Fetch data").

Sources
  * Yahoo Finance via yfinance: share price on the valuation date, shares outstanding, beta, peer data and,
    if chosen, the financial statements.
  * IDX official filings (audited annual XBRL) via the open-source package IDX_GetFinancialStatement
    (https://github.com/celeebi/IDX_GetFinancialStatement, MIT licence, by celeebi). Optional: used when the package
    is found; lines it does not provide are taken from Yahoo. The company search also uses the IDX company list.

Start (once: pip3 install yfinance openpyxl pandas; for IDX filings also the requirements of that package):
    python3 data_connector.py                                  # Yahoo only
    python3 data_connector.py --idx-repo ./IDX_GetFinancialStatement   # with IDX filings
Then open valuation_platform.html in the browser -> Market & company -> Fetch data.
Listens on http://127.0.0.1:8765 only (not reachable from other computers). Stop with Ctrl+C.
"""
import argparse, datetime as dt, json, os, sys, tempfile, threading, traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import fetch_yahoo as FY

VERSION = "1.0"
STATE = {"idx": None, "idx_error": None, "emitens": None, "cache_dir": os.path.join(tempfile.gettempdir(), "evp_idx_cache")}
LOCK = threading.Lock()
FETCH_LOCK = threading.Lock()

# ---------------------------------------------------------------- IDX package (optional)
def load_idx(repo):
    cands = [repo] if repo else [os.path.join(HERE, "IDX_GetFinancialStatement"), os.path.join(os.getcwd(), "IDX_GetFinancialStatement")]
    for c in cands:
        if c and os.path.isdir(os.path.join(c, "idx_scraper")) and c not in sys.path:
            sys.path.insert(0, c)
    try:
        from idx_scraper.api import IDXApi
        from idx_scraper.scraper import IDXFinancialScraper
        STATE["idx"] = {"api": IDXApi, "scraper": IDXFinancialScraper}
    except Exception as e:
        STATE["idx"] = None
        STATE["idx_error"] = f"{type(e).__name__}: {e}"

def idx_api():
    if not STATE["idx"]: return None
    with LOCK:
        if "api_obj" not in STATE["idx"]:
            STATE["idx"]["api_obj"] = STATE["idx"]["api"]()
        return STATE["idx"]["api_obj"]

def _pick(d, *names):
    low = {str(k).lower().replace("_", ""): v for k, v in d.items()}
    for n in names:
        v = low.get(n.lower().replace("_", ""))
        if v not in (None, ""): return str(v).strip()
    return ""

def idx_blocked():
    """True for 6 h after IDX refused this server (Cloudflare often blocks cloud IPs) - avoids slow retries on every click."""
    t = STATE.get("idx_blocked_at")
    return bool(t and (dt.datetime.now() - t).total_seconds() < 21600)

def idx_companies():
    """All IDX-listed shares with sector/industry where IDX provides them (cached 12 h)."""
    c = STATE.get("companies")
    if c and (dt.datetime.now() - c["at"]).total_seconds() < 43200: return c
    rows, src, errs = [], "", []
    api = idx_api()
    try:   # company profiles carry sector and industry
        if not api: raise RuntimeError(f"IDX package not available ({STATE['idx_error']})")
        r = api.client.get("/primary/ListedCompany/GetCompanyProfiles", params={"emitenType": "s", "start": 0, "length": 9999})
        j = r.json(); data = j.get("data") if isinstance(j, dict) else j
        for d in data or []:
            t = _pick(d, "KodeEmiten", "Kode_Emiten", "Code")
            if t: rows.append({"ticker": t, "name": _pick(d, "NamaEmiten", "Nama_Emiten", "Name"), "sector": _pick(d, "Sektor", "Sector"),
                               "industry": _pick(d, "Industri", "SubSektor", "Sub_Sektor", "Industry", "SubIndustri"), "board": _pick(d, "PapanPencatatan", "Papan_Pencatatan", "ListingBoard"),
                               "listed": _pick(d, "TanggalPencatatan", "Tanggal_Pencatatan", "ListingDate")[:10]})
        src = "IDX listed company profiles"
    except Exception as e:
        errs.append(f"IDX profiles: {type(e).__name__}: {e}"); sys.stderr.write(errs[-1] + "\n")
    if not rows and api and STATE["idx"]: STATE["idx_blocked_at"] = dt.datetime.now()
    if not rows and api and not idx_blocked():   # plain issuer list (code and name only)
        try:
            rows = [{"ticker": e.get("KodeEmiten", "").strip(), "name": e.get("NamaEmiten", "").strip(), "sector": "", "industry": "", "board": "", "listed": ""}
                    for e in api.get_emitens() if e.get("KodeEmiten")]
            src = "IDX issuer list"
        except Exception as e:
            errs.append(f"IDX issuer list: {type(e).__name__}: {e}"); STATE["idx_blocked_at"] = dt.datetime.now()
    if not rows:   # IDX blocks many cloud servers: Yahoo's equity screener for the Jakarta exchange
        try:
            rows = yahoo_companies(); src = "Yahoo Finance equity screener, exchange JKT (IDX not reachable from this server)"
        except Exception as e:
            errs.append(f"Yahoo screener: {type(e).__name__}: {e}")
    if not rows: raise RuntimeError("; ".join(errs) or "no source available")
    rows.sort(key=lambda r: r["ticker"])
    STATE["companies"] = {"at": dt.datetime.now(), "rows": rows, "source": src}
    return STATE["companies"]

YAHOO_SECTORS = ["Basic Materials", "Communication Services", "Consumer Cyclical", "Consumer Defensive", "Energy", "Financial Services",
                 "Healthcare", "Industrials", "Real Estate", "Technology", "Utilities"]

def yahoo_companies():
    import yfinance as yf
    Q = yf.EquityQuery; rows = {}
    for sec in YAHOO_SECTORS:
        off = 0
        while off < 2000:
            res = yf.screen(Q("and", [Q("eq", ["exchange", "JKT"]), Q("eq", ["sector", sec])]), offset=off, size=250) or {}
            qs = res.get("quotes", []) or []
            for q in qs:
                sym = q.get("symbol", "")
                if not sym.endswith(".JK"): continue
                t = sym[:-3]
                rows[t] = {"ticker": t, "name": q.get("longName") or q.get("shortName") or t, "sector": sec, "industry": q.get("industry", "") or "",
                           "board": "", "listed": "", "mcap": q.get("marketCap")}
            if len(qs) < 250: break
            off += 250
    return sorted(rows.values(), key=lambda r: r["ticker"])

def yahoo_industry_peers(industry):
    import yfinance as yf
    Q = yf.EquityQuery
    res = yf.screen(Q("and", [Q("eq", ["exchange", "JKT"]), Q("eq", ["industry", industry])]), size=100) or {}
    return [(q["symbol"][:-3], q.get("longName") or q.get("shortName") or q["symbol"], q.get("marketCap")) for q in res.get("quotes", []) if q.get("symbol", "").endswith(".JK")]

def peer_candidates(ticker, cap=12, P=None):
    """Same IDX industry first, then same sector; returns [(ticker, name, level)] (level: 'industry' or 'sector')."""
    rows = idx_companies()["rows"]; me = next((r for r in rows if r["ticker"] == ticker.upper()), None)
    yind = ((P or {}).get("meta") or {}).get("yahooIndustry")
    if (not me or not me.get("industry")) and yind:   # Yahoo classification: same Yahoo industry on the Jakarta exchange, then same sector by size
        import math
        out = []
        try: out = [(t, n, "industry") for t, n, _ in yahoo_industry_peers(yind) if t != ticker.upper()]
        except Exception as e: sys.stderr.write(f"Yahoo industry screen failed: {e}\n")
        sec = (me or {}).get("sector") or ((P or {}).get("meta") or {}).get("yahooSector")
        if len(out) < 6 and sec:
            m = (P or {}).get("market") or {}; tm = (m.get("price") or 0) * (m.get("sharesMn") or 0) * 1e6
            pool = [r for r in rows if r["sector"] == sec and r["ticker"] != ticker.upper() and r["ticker"] not in {o[0] for o in out}]
            if tm: pool.sort(key=lambda r: abs(math.log((r.get("mcap") or 1) / tm)) if r.get("mcap") else 99)
            out += [(r["ticker"], r["name"], "sector") for r in pool]
        return out[:cap], me, f"{len([o for o in out if o[2] == 'industry'])} in the same Yahoo industry ({yind}), sector {sec or '-'}"
    if not me or not (me["industry"] or me["sector"]): return [], None, "no sector data for this company in the IDX list"
    out = [(r["ticker"], r["name"], "industry") for r in rows if me["industry"] and r["industry"] == me["industry"] and r["ticker"] != me["ticker"]]
    if len(out) < 6 and me["sector"]:
        out += [(r["ticker"], r["name"], "sector") for r in rows if r["sector"] == me["sector"] and r["ticker"] != me["ticker"] and r["ticker"] not in {o[0] for o in out}]
    return out[:cap], me, f"{len([o for o in out if o[2] == 'industry'])} in the same industry ({me['industry'] or '-'}), sector {me['sector'] or '-'}"

def _target_metrics(P):
    import math
    F = P["financials"]; it = F.get("items", {}); ys = F.get("years", [])
    last = lambda k: (it.get(k) or [None])[-1] if it.get(k) else None
    rev = it.get("revenue") or []
    g = (rev[-1] / rev[-2] - 1) if len(rev) >= 2 and rev[-1] and rev[-2] else None
    eq = last("equityTotal"); debt = (last("bankLoans") or 0) + (last("leases") or 0)
    m = P["market"]
    return {"mcap": (m.get("price") or 0) * (m.get("sharesMn") or 0) or None, "growth": g, "de": debt / eq if eq else None,
            "eps": last("eps"), "profitable": (last("niParent") or 0) > 0}

def score_peer(t, r, level):
    """0-100 points with the reasons shown to the user."""
    import math
    pts, why = 0.0, []
    if level == "industry": pts += 40; why.append("same industry (+40)")
    else: pts += 20; why.append("same sector only (+20)")
    mc = (r.get("price") or 0) * (r.get("sharesMn") or 0)
    if mc and t["mcap"]:
        ratio = mc / t["mcap"]; s = 25 * max(0.0, 1 - abs(math.log(ratio)) / math.log(10)); pts += s
        why.append(f"size {ratio:.2f}x of the company (+{s:.0f})")
    else: why.append("size unknown (+0)")
    if r.get("eps") is not None and r["eps"] > 0: pts += 15; why.append("profitable, P/E usable (+15)")
    else: why.append("loss or no EPS, P/E not usable (+0)")
    if r.get("revGrowth") is not None and t["growth"] is not None:
        s = 10 * max(0.0, 1 - abs(r["revGrowth"] - t["growth"]) / 0.30); pts += s
        why.append(f"revenue growth {r['revGrowth']*100:.0f}% vs {t['growth']*100:.0f}% (+{s:.0f})")
    else: why.append("growth unknown (+0)")
    eq = (r.get("bvps") or 0) * (r.get("sharesMn") or 0)
    if eq > 0 and r.get("debt") is not None and t["de"] is not None:
        de = r["debt"] / eq; s = 10 * max(0.0, 1 - abs(de - t["de"]) / 2.0); pts += s
        why.append(f"debt/equity {de:.2f} vs {t['de']:.2f} (+{s:.0f})")
    else: why.append("leverage unknown (+0)")
    return round(pts, 1), why, (mc / t["mcap"] if mc and t["mcap"] else None)

def select_peers(P, ticker, exchange, date, n=5):
    cands, me, info = peer_candidates(ticker, P=P)
    if not cands: return [], [], f"Peers could not be chosen automatically ({info}); add them by hand."
    sfx = FY.SUFFIX.get(exchange.upper(), ""); t = _target_metrics(P); evald = []; notes = []
    for tk, name, level in cands:
        try:
            r = FY.peer_row(tk + sfx, tk, date, notes)
        except Exception as e:
            evald.append({"ticker": tk, "name": name, "score": 0, "reasons": [f"no data ({type(e).__name__})"], "selected": False}); continue
        if not r.get("complete", True):
            miss = [k for k in FY.PEER_REQUIRED if r.get(k) is None]
            evald.append({"ticker": tk, "name": r.get("name") or name, "score": 0, "reasons": [f"not usable: Yahoo has no {', '.join(miss)}"], "selected": False}); continue
        sc, why, ratio = score_peer(t, r, level)
        r["closeModel"] = level == "industry" and ratio is not None and 0.2 <= ratio <= 5
        evald.append({"ticker": tk, "name": r.get("name") or name, "score": sc, "reasons": why, "selected": False, "_row": r})
    evald.sort(key=lambda x: -x["score"])
    chosen = [e for e in evald if e.get("_row") and e["score"] >= 40][:n]
    for e in chosen: e["selected"] = True
    rows = [e["_row"] for e in chosen]
    for e in evald: e.pop("_row", None)
    msg = (f"Peers chosen by score from {len(evald)} candidates ({info}): " + ", ".join(f"{e['ticker']} {e['score']:.0f}" for e in chosen) + "."
           if chosen else f"No candidate reached 40 points ({info}); add peers by hand.")
    return rows, evald, msg

def idx_search(q):
    api = idx_api()
    if not api: return None
    return [{"ticker": e.get("KodeEmiten", "").strip(), "name": e.get("NamaEmiten", "").strip(), "exchange": "IDX", "source": "IDX company list"}
            for e in api.search_emiten(q)][:30]

def yahoo_search(q):
    import yfinance as yf
    try:
        quotes = yf.Search(q, max_results=15).quotes
    except Exception as e:
        STATE["last_search_error"] = f"Yahoo search: {type(e).__name__}: {e}"
        return []
    inv = {v: k for k, v in FY.SUFFIX.items() if v}
    inv[".JK"] = "IDX"
    out = []
    for x in quotes:
        sym = x.get("symbol", "")
        if x.get("quoteType") not in (None, "EQUITY"): continue
        base, sfx = (sym.rsplit(".", 1)[0], "." + sym.rsplit(".", 1)[1]) if "." in sym else (sym, "")
        ex = inv.get(sfx) or (x.get("exchange") or "").upper()
        if not sfx and ex in ("NMS", "NGM", "NCM", "NAS"): ex = "NASDAQ"
        if not sfx and ex in ("NYQ", "NYS"): ex = "NYSE"
        out.append({"ticker": base, "name": x.get("longname") or x.get("shortname") or sym, "exchange": ex, "source": "Yahoo search"})
    return out

def g(obj, *names):
    for n in names:
        v = getattr(obj, n, None)
        if v is not None:
            try:
                v = float(v)
            except (TypeError, ValueError):
                continue
            return v
    return None

def idx_financials(ticker, years):
    """Run the IDX package for the given fiscal years and map its normalised statements to platform fields."""
    S = STATE["idx"]["scraper"](output_dir=STATE["cache_dir"], client=idx_api().client)
    res = S.build_ready_to_forecast(ticker=ticker, years=[int(y) for y in years])
    if not res or res.get("status") != "success":
        raise RuntimeError(res.get("message", "IDX retrieval failed") if res else "IDX retrieval failed")
    bank = res.get("archetype") == "banking"
    out = {}
    for y in res.get("years", []):
        i, b, c = res["is_dict"].get(y), res["bs_dict"].get(y), res["cf_dict"].get(y)
        if i is None or b is None or c is None: continue
        if bank:
            m = {"revenue": g(i, "total_operating_income"), "ebit": g(i, "operating_profit"), "pbt": g(i, "ebt"), "taxExpense": g(i, "tax_expense"),
                 "netIncome": g(i, "net_income"), "niParent": g(i, "net_income_parent", "net_income"), "totalAssets": g(b, "total_assets"),
                 "totalLiabilities": g(b, "total_liabilities"), "equityTotal": g(b, "total_equity"), "cash": g(b, "cash_and_equivalents"),
                 "cfo": g(c, "operating_cash_flow"), "dividends": g(c, "dividends_paid")}
        else:
            st, lt = g(b, "short_term_debt") or 0.0, g(b, "long_term_debt") or 0.0
            ca, cl, cash = g(b, "total_current_assets"), g(b, "total_current_liabilities"), g(b, "cash_and_equivalents")
            m = {"revenue": g(i, "revenue"), "ebit": g(i, "operating_income_ebit"), "da": g(i, "depreciation_amortization"),
                 "otherIncome": g(i, "other_non_operating"), "financeCharges": g(i, "finance_costs"), "pbt": g(i, "ebt"),
                 "taxExpense": g(i, "tax_expense"), "netIncome": g(i, "net_income"), "niParent": g(i, "net_income_parent"),
                 "cfo": g(c, "operating_cash_flow"), "capex": g(c, "capex"), "dividends": g(c, "dividends_paid"),
                 "totalAssets": g(b, "total_assets"), "totalLiabilities": g(b, "total_liabilities"), "equityTotal": g(b, "total_equity"),
                 "cash": cash, "bankLoans": st + lt,
                 "opNWC": None if not ca or not cl else (ca - (cash or 0.0)) - (cl - st)}
        for k in ("capex", "dividends", "financeCharges"):
            if m.get(k) is not None: m[k] = abs(m[k])
        # the package reports a missing line as 0.0: treat zeros as "not found" except where 0 is a normal value
        out[str(y)] = {k: v for k, v in m.items() if v is not None and (v != 0.0 or k == "otherIncome")}
    return out, bank

# ---------------------------------------------------------------- Damodaran (public data sets, NYU Stern)
DAMO = "https://pages.stern.nyu.edu/~adamodar/"
BETA_FILE = {"IDX": "betaemerg.xls", "BEI": "betaemerg.xls", "SET": "betaemerg.xls", "KLSE": "betaemerg.xls", "BURSA": "betaemerg.xls", "BIST": "betaemerg.xls",
             "HKEX": "betaChina.xls", "NSE": "betaIndia.xls", "BSE": "betaIndia.xls", "NYSE": "betas.xls", "NASDAQ": "betas.xls", "AMEX": "betas.xls",
             "LSE": "betaEurope.xls", "XETRA": "betaEurope.xls", "FWB": "betaEurope.xls", "SIX": "betaEurope.xls", "EURONEXT": "betaEurope.xls",
             "TSE": "betaJapan.xls", "ASX": "betaRest.xls", "TSX": "betaRest.xls"}
EX_COUNTRY = {"IDX": "Indonesia", "BEI": "Indonesia", "NYSE": "United States", "NASDAQ": "United States", "AMEX": "United States", "LSE": "United Kingdom",
              "XETRA": "Germany", "FWB": "Germany", "SGX": "Singapore", "ASX": "Australia", "HKEX": "Hong Kong", "TSE": "Japan", "BIST": "Turkey",
              "SET": "Thailand", "KLSE": "Malaysia", "BURSA": "Malaysia", "NSE": "India", "BSE": "India", "SIX": "Switzerland", "TSX": "Canada", "EURONEXT": "France"}
DCACHE = {}

def http_get(url):
    if url in DCACHE: return DCACHE[url]
    from curl_cffi import requests as cr   # same HTTP stack as the IDX package; plain requests also works for Damodaran
    r = cr.get(url, timeout=40, impersonate="chrome131")
    if r.status_code != 200: raise RuntimeError(f"HTTP {r.status_code} for {url}")
    DCACHE[url] = r.content
    return r.content

def pct(x):
    try:
        t = str(x).strip().replace(",", "")
        if t in ("", "NA", "nan", "None"): return None
        return float(t[:-1]) / 100 if t.endswith("%") else float(t)
    except ValueError:
        return None

def html_rows(content):
    import pandas as pd, io
    rows = []
    for t in pd.read_html(io.BytesIO(content), header=None):
        rows += [[("" if (isinstance(c, float) and c != c) else " ".join(str(c).split())) for c in r] for r in t.values.tolist()]
    return rows

def damo_country(country):
    import difflib
    rows = html_rows(http_get(DAMO + "New_Home_Page/datafile/ctryprem.html"))
    h = next(r for r in rows if r and r[0] == "Country")
    col = lambda *keys: next(i for i, c in enumerate(h) if all(k in c.lower() for k in keys))
    ie, it, idef, icrp = col("equity risk"), col("tax"), col("default"), col("country risk")
    data = {r[0]: r for r in rows if r and r[0] and r[0] != "Country" and len(r) > ie}
    names = list(data)
    hit = country if country in data else (difflib.get_close_matches(country or "", names, 1, 0.6) or [None])[0]
    if not hit: raise RuntimeError(f"country '{country}' not in Damodaran's table")
    r = data[hit]
    return {"country": hit, "erp": pct(r[ie]), "tax": pct(r[it]), "defaultSpread": pct(r[idef]), "crp": pct(r[icrp])}

def damo_betas(fname):
    import pandas as pd, io
    sheets = pd.read_excel(io.BytesIO(http_get(DAMO + "pc/datasets/" + fname)), sheet_name=None, header=None)
    for df in sheets.values():
        vals = [["" if (isinstance(c, float) and c != c) or c is None else str(c) for c in row] for row in df.values.tolist()]
        for i, r in enumerate(vals):
            r = [" ".join(c.split()) for c in r]
            if "Industry Name" in r:
                ni = r.index("Industry Name")
                jb = next((j for j, c in enumerate(r) if "corrected for cash" in c.lower()), None)
                if jb is None: jb = next(j for j, c in enumerate(r) if c.lower().startswith("unlevered beta"))
                out = []
                for rr in vals[i + 1:]:
                    name = " ".join(str(rr[ni]).split()); b = pct(rr[jb])
                    if name and name != "nan" and b is not None: out.append({"name": name, "betaU": round(b, 4)})
                if out: return out
    raise RuntimeError(f"no beta table found in {fname}")

def damo_ratings():
    rows = html_rows(http_get(DAMO + "New_Home_Page/datafile/ratings.html"))
    out = []
    for r in rows:
        c = [x for x in r if x != ""][:4]
        if len(c) == 4 and pct(c[0]) is not None and pct(c[1]) is not None and str(c[3]).endswith("%"):
            out.append({"lo": float(str(c[0]).replace(",", "")), "hi": float(str(c[1]).replace(",", "")), "rating": c[2], "spread": pct(c[3])})
    if not out: raise RuntimeError("rating table not found")
    return out

def suggest_industry(yahoo_industry, yahoo_sector, names):
    import difflib
    for q in (yahoo_industry, yahoo_sector):
        if q:
            m = difflib.get_close_matches(q, names, 1, 0.45)
            if m: return m[0]
            ql = q.lower().split()[0]
            m = [n for n in names if ql in n.lower()]
            if m: return m[0]
    return "Total Market" if "Total Market" in names else names[-1]

def us_riskfree():
    import yfinance as yf
    h = yf.Ticker("^TNX").history(period="10d")
    return None if h is None or h.empty else round(float(h["Close"].iloc[-1]) / 100, 5)

def damodaran(exchange, country, y_ind, y_sec, currency):
    ex = (exchange or "").upper(); out = {"notes": []}
    when = f"{dt.date.today()}"
    try:
        c = damo_country(country or EX_COUNTRY.get(ex))
        out.update(c); out["countrySource"] = f"Damodaran, Country Risk Premiums (ctryprem), retrieved {when}: {c['country']}"
    except Exception as e: out["notes"].append(f"Damodaran country data not available ({type(e).__name__}: {e}).")
    f = BETA_FILE.get(ex, "betaGlobal.xls")
    try:
        inds = damo_betas(f); out["industries"] = inds; out["betaFile"] = f
        out["suggestedIndustry"] = suggest_industry(y_ind, y_sec, [i["name"] for i in inds])
        out["betaSource"] = f"Damodaran, Betas by Sector ({f}), unlevered beta corrected for cash, retrieved {when}"
    except Exception as e: out["notes"].append(f"Damodaran betas not available ({type(e).__name__}: {e}).")
    try:
        out["ratings"] = damo_ratings(); out["ratingSource"] = f"Damodaran, synthetic rating table for large non-financial firms, retrieved {when}"
    except Exception as e: out["notes"].append(f"Damodaran rating table not available ({type(e).__name__}: {e}).")
    if (currency or "").upper() == "USD":
        try:
            rf = us_riskfree()
            if rf: out["rf"] = rf; out["rfSource"] = f"US Treasury 10-year yield (Yahoo ^TNX), retrieved {when}"
        except Exception as e: out["notes"].append(f"US 10-year yield not available ({e}).")
    else:
        out["notes"].append(f"Risk-free rate: enter the 10-year government bond yield in {currency or 'the local currency'} (no free reliable source for all countries).")
    return out

# ---------------------------------------------------------------- fetch
def fetch(ticker, exchange, peers, date, fin_source):
    try:
        P = FY.collect(ticker, exchange, peers, date)
    except Exception as e:
        raise RuntimeError(f"Yahoo Finance could not be reached or does not know {ticker} on {exchange} ({type(e).__name__}: {e}). "
                           "Check the internet connection and the ticker; if it persists run: pip3 install -U yfinance")
    P["connector"] = {"version": VERSION, "financialSource": "Yahoo"}
    try:
        P["damodaran"] = damodaran(exchange, P["meta"].get("country"), P["meta"].get("yahooIndustry"), P["meta"].get("yahooSector"), P["meta"].get("currency"))
    except Exception as e:
        P["notes"].append(f"Damodaran data not available ({type(e).__name__}: {e}).")
    if P.get("damodaran"):
        P["notes"] = [n for n in P["notes"] if not n.startswith("Not on Yahoo")] + ["Cost-of-capital inputs: see 'Suggested market inputs' (Damodaran); nothing is changed until you tick and apply."]
    if fin_source == "idx":
        if (exchange or "").upper() not in ("IDX", "BEI"):
            P["notes"].append("IDX filings are only available for IDX-listed companies; Yahoo statements used.")
        elif idx_blocked():
            P["notes"].append("IDX (idx.co.id) does not answer this server at the moment; statements taken from Yahoo Finance.")
        elif not STATE["idx"]:
            P["notes"].append(f"IDX package not available ({STATE['idx_error']}); Yahoo statements used. Start the connector with --idx-repo <folder>.")
        else:
            fy = int(date[:4]) - 1
            years = P["financials"]["years"] or [str(fy - 2), str(fy - 1), str(fy)]
            try:
                idx, bank = idx_financials(ticker.upper(), years)
            except Exception as e:
                idx, bank = {}, False; STATE["idx_blocked_at"] = dt.datetime.now()
                P["notes"].append(f"IDX filings could not be retrieved ({type(e).__name__}: {e}); Yahoo statements used.")
            if idx:
                yrs = sorted(set(years) | set(idx))[-3:]
                ya = {k: dict(zip(P["financials"]["years"], v)) for k, v in P["financials"]["items"].items()}
                keys = sorted(set(ya) | {k for d in idx.values() for k in d})
                items, line = {}, {}
                for k in keys:
                    vals, srcs = [], set()
                    for y in yrs:
                        v = idx.get(y, {}).get(k)
                        if v is not None: srcs.add("IDX")
                        else:
                            v = ya.get(k, {}).get(y)
                            if v is not None: srcs.add("Yahoo")
                        vals.append(v)
                    if any(v is not None for v in vals):
                        items[k] = vals; line[k] = "+".join(sorted(srcs))
                # parent equity and NCI: IDX package gives total equity only; split with Yahoo's NCI when available
                if "equityTotal" in items and "nci" in items and line.get("equityTotal", "").startswith("IDX"):
                    items["equityParent"] = [None if a is None or n is None else a - n for a, n in zip(items["equityTotal"], items["nci"])]
                    line["equityParent"] = "IDX total equity − Yahoo NCI"
                P["financials"] = {"years": yrs, "items": items, "lineSource": line,
                                   "source": f"IDX audited annual filings (XBRL) via IDX_GetFinancialStatement, gaps filled from Yahoo Finance; retrieved {dt.datetime.now():%Y-%m-%d %H:%M}"}
                P["connector"]["financialSource"] = "IDX"
                P["profile"] = {"isFinancial": bank}
                yo = [k for k, v in line.items() if v == "Yahoo"]
                if yo: P["notes"].append("From Yahoo because the IDX filing had no value: " + ", ".join(yo) + ".")
                if P["meta"].get("currency") and P["meta"]["currency"] != "IDR":
                    P["notes"].append("IDX package reports in IDR million; check the currency.")
    return P

# ---------------------------------------------------------------- HTTP
class H(BaseHTTPRequestHandler):
    def _send(self, code, obj):
        body = json.dumps(obj, default=str).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Private-Network", "true")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers(); self.wfile.write(body)
    def do_HEAD(self):
        self.send_response(200); self.send_header("Content-Type", "text/html"); self.end_headers()
    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "*")
        self.send_header("Access-Control-Allow-Private-Network", "true")
        self.end_headers()
    def log_message(self, fmt, *a):
        sys.stderr.write("  " + (fmt % a) + "\n")
    def do_GET(self):
        u = urlparse(self.path); q = {k: v[0] for k, v in parse_qs(u.query).items()}
        try:
            if u.path in ("/", "/index.html"):
                page = os.path.join(HERE, "valuation_platform.html")
                if not os.path.exists(page): return self._send(404, {"error": "valuation_platform.html not found next to data_connector.py"})
                body = open(page, "rb").read()
                self.send_response(200); self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Cache-Control", "no-cache"); self.send_header("Content-Length", str(len(body))); self.end_headers()
                return self.wfile.write(body)
            if u.path == "/health":
                return self._send(200, {"ok": True, "version": VERSION, "yahoo": True, "idx": bool(STATE["idx"]) and not idx_blocked(), "idxBlocked": idx_blocked(), "idxError": STATE["idx_error"]})
            if u.path == "/companies":
                try:
                    c = idx_companies()
                    return self._send(200, {"source": c["source"], "retrieved": c["at"].strftime("%Y-%m-%d %H:%M"), "companies": c["rows"]})
                except Exception as e:
                    return self._send(502, {"error": f"IDX company list not available ({type(e).__name__}: {e})"})
            if u.path == "/search":
                text = q.get("q", "").strip()
                if len(text) < 2: return self._send(200, {"results": []})
                res = []; STATE["last_search_error"] = None; idx_err = None
                if q.get("exchange", "IDX").upper() in ("IDX", "BEI", ""):
                    try: res = idx_search(text) or []
                    except Exception as e: idx_err = f"IDX list: {type(e).__name__}: {e}"; sys.stderr.write(idx_err + "\n")
                seen = {r["ticker"] for r in res}
                res += [r for r in yahoo_search(text) if r["ticker"] not in seen]
                if not res and (idx_err or STATE.get("last_search_error")):
                    return self._send(502, {"error": "search sources not reachable - check the internet connection (" + "; ".join(x for x in (idx_err, STATE.get("last_search_error")) if x) + ")"})
                return self._send(200, {"results": res[:30]})
            if u.path == "/fetch":
                t = q.get("ticker", "").strip()
                if not t: return self._send(400, {"error": "ticker missing"})
                peers = [p for p in q.get("peers", "").replace(",", " ").split() if p]
                date = q.get("date") or str(dt.date.today())
                if len(peers) > 12: return self._send(400, {"error": "at most 12 peers per request"})
                auto = peers == ["auto"]
                if auto: peers = []
                with FETCH_LOCK:
                    P = fetch(t, q.get("exchange", "IDX"), peers, date, q.get("fin", "yahoo"))
                    if auto:
                        try: rows, sel, msg = select_peers(P, t, q.get("exchange", "IDX"), date)
                        except Exception as e: rows, sel, msg = [], [], f"Peers could not be chosen automatically ({type(e).__name__}: {e})."
                        P["peers"] = rows; P["peerSelection"] = sel; P["autoPeers"] = [r["ticker"] for r in rows]; P["notes"].insert(0, msg)
                return self._send(200, P)
            return self._send(404, {"error": "unknown path"})
        except Exception as e:
            traceback.print_exc()
            return self._send(500, {"error": f"{type(e).__name__}: {e}"})

def serve(port=8765, repo=None, quiet=False, host="127.0.0.1"):
    load_idx(repo)
    srv = ThreadingHTTPServer((host, port), H)
    if not quiet:
        print(f"Equity Valuation Platform data connector {VERSION} on http://{host}:{port}  (platform page: http://127.0.0.1:{port}/)")
        print("  Yahoo Finance: on" + ("  |  IDX filings: on" if STATE["idx"] else f"  |  IDX filings: off ({STATE['idx_error']})"))
        print("  Open valuation_platform.html -> Market & company -> Fetch data. Stop with Ctrl+C.")
    return srv

if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--port", type=int, default=int(os.environ.get("PORT", 8765)))
    ap.add_argument("--host", default=os.environ.get("HOST", "127.0.0.1"), help="0.0.0.0 to serve the platform publicly (hosting)")
    ap.add_argument("--idx-repo", default=None, help="folder of IDX_GetFinancialStatement (contains idx_scraper/)")
    a = ap.parse_args()
    srv = serve(a.port, a.idx_repo, host=a.host)
    try: srv.serve_forever()
    except KeyboardInterrupt: print("\nstopped")
