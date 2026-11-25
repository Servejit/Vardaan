import io
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from streamlit_autorefresh import st_autorefresh


def _sym(value):
    value = re.sub(r"[^A-Z0-9&_-]", "", re.sub(r"\.NS$", "", str(value or "").strip().upper()))
    return "" if value in {"", "SYMBOL", "STOCK", "STOCKNAME", "STOCK_NAME", "NAME", "NSE", "TICKER", "COMPANY", "SCRIP"} else value


def _num(value):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    raw = str(value).strip().replace(",", "").replace("−", "-").replace("%", "")
    for part in reversed(re.findall(r"\(([^()]*)\)", raw)):
        match = re.search(r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)", part)
        if match:
            return float(match.group(0))
    match = re.search(r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)", raw)
    return float(match.group(0)) if match else None


def _read_master(master_bytes):
    if not master_bytes:
        raise ValueError("Master Excel is not available.")
    wb = load_workbook(io.BytesIO(master_bytes), data_only=True, read_only=True)
    try:
        sheets = sorted(wb.worksheets, key=lambda ws: (ws.title.casefold() != "summary", ws.title.casefold()))
        required = ("16< D-B / Avg.4", "16> C-B / Avg.4")
        clean = lambda x: " ".join(str(x or "").strip().casefold().split())
        for ws in sheets:
            for hrow, values in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row or 1, 30), values_only=True), 1):
                cols = {clean(v): i for i, v in enumerate(values) if v is not None}
                if all(clean(h) in cols for h in required):
                    symbols, caps = [], {}
                    for row in ws.iter_rows(min_row=hrow + 1, values_only=True):
                        if not row:
                            continue
                        symbol = _sym(row[0] if row else None)
                        if symbol:
                            symbols.append(symbol)
                            caps[symbol] = {
                                "o2l4": _num(row[cols[clean(required[0])]] if cols[clean(required[0])] < len(row) else None),
                                "o2h4": _num(row[cols[clean(required[1])]] if cols[clean(required[1])] < len(row) else None),
                            }
                    if symbols:
                        return list(dict.fromkeys(symbols)), caps, ws.title
        raise ValueError("Could not find both Master headings: 16< D-B / Avg.4 and 16> C-B / Avg.4.")
    finally:
        wb.close()


@st.cache_data(ttl=25, max_entries=2, show_spinner=False)
def _market_data(symbols):
    import yfinance as yf

    def fetch(symbol):
        try:
            ticker = yf.Ticker(f"{symbol}.NS")
            daily = ticker.history(period="2mo", interval="1d", auto_adjust=False, actions=False, prepost=False, timeout=18)
            intra = ticker.history(period="1d", interval="1m", auto_adjust=False, actions=False, prepost=False, timeout=18)
            if intra is None or intra.empty:
                intra = ticker.history(period="5d", interval="5m", auto_adjust=False, actions=False, prepost=False, timeout=18)
            if daily is None or daily.empty or intra is None or intra.empty:
                return symbol, None
            daily = daily.dropna(subset=["Low", "High", "Close"])
            intra = intra.dropna(subset=["Close"]).sort_index()
            if daily.empty or intra.empty:
                return symbol, None
            session_date = intra.index[-1].date()
            completed = daily[[idx.date() < session_date for idx in daily.index]].tail(20)
            if len(completed) < 20:
                return symbol, None
            lows = sorted(float(v) for v in completed["Low"].dropna())[:5]
            highs = sorted((float(v) for v in completed["High"].dropna()), reverse=True)[:5]
            if len(lows) != 5 or len(highs) != 5:
                return symbol, None
            session = intra[[idx.date() == session_date for idx in intra.index]]
            bars = []
            for stamp, row in session.iterrows():
                close = _num(row.get("Close"))
                if close is not None:
                    bars.append({"time": stamp.isoformat(), "open": _num(row.get("Open")),
                                 "high": _num(row.get("High")), "low": _num(row.get("Low")), "close": close})
            if not bars:
                return symbol, None
            return symbol, {
                "session_date": session_date.isoformat(),
                "previous_close": float(completed["Close"].iloc[-1]),
                "open": bars[0]["open"],
                "high": max((b["high"] for b in bars if b["high"] is not None), default=None),
                "low": min((b["low"] for b in bars if b["low"] is not None), default=None),
                "lows5": lows, "highs5": highs, "avg_low": sum(lows) / 5,
                "avg_high": sum(highs) / 5, "bars": bars,
            }
        except Exception:
            return symbol, None

    results = {}
    symbols = list(dict.fromkeys(symbols))
    with ThreadPoolExecutor(max_workers=min(8, max(1, len(symbols)))) as pool:
        futures = [pool.submit(fetch, symbol) for symbol in symbols]
        for future in as_completed(futures):
            symbol, data = future.result()
            if data:
                results[symbol] = data
    return results


def _events(symbol, data):
    """Return at most one currently-active signal per stock and side."""
    bars = data["bars"]
    if not bars:
        return []

    low_avg = data["avg_low"]
    high_avg = data["avg_high"]
    previous_close = data["previous_close"]
    current_price = bars[-1]["close"]
    last_time = bars[-1]["time"]
    day_low = data.get("low")
    day_high = data.get("high")
    active = []

    # BUY remains active only while the latest traded price is back above
    # Avg.Low, after today's Low has reached at least 1% below Avg.Low.
    if (
        previous_close > low_avg
        and day_low is not None
        and day_low <= low_avg * 0.99
        and current_price > low_avg
    ):
        active.append({
            "symbol": symbol,
            "side": "BUY",
            "time": last_time,
            "entry": current_price,
        })

    # SELL remains active only while the latest traded price is back below
    # Avg.High, after today's High has reached at least 1% above Avg.High.
    if (
        previous_close < high_avg
        and day_high is not None
        and day_high >= high_avg * 1.01
        and current_price < high_avg
    ):
        active.append({
            "symbol": symbol,
            "side": "SELL",
            "time": last_time,
            "entry": current_price,
        })

    return active


def _price(value):
    return "—" if value is None else f"₹{value:,.2f}"


def render_avg_high_low_dashboard(master_bytes, app_name="6THSENSE"):
    if app_name.upper() == "VARDAAN":
        st.session_state["_vardaan_active_theme"] = {
            "deep": "#5B1748", "primary": "#EC64D5", "accent": "#EC64D5",
            "ink": "#3F1738", "accent_text": "#5B1748", "tint": "#F8DDF2",
            "canvas": "#FFF5FC", "border": "#E7A5D6", "soft_text": "#5B1748",
            "left_accent": "#EC64D5",
        }
    st.markdown("""
    <style>


    [data-testid="stExpander"]{border:1.5px solid #EC64D5!important;border-radius:12px!important;overflow:hidden!important;background:#FFF5FC!important;box-shadow:0 4px 14px rgba(91,23,72,.10)!important;margin:8px 0 18px!important}
    [data-testid="stExpander"] details>summary{background:linear-gradient(110deg,#8A246D 0%,#B43B91 58%,#EC64D5 100%)!important;color:#FFF8FD!important;padding:12px 16px!important;font-weight:900!important;letter-spacing:.5px!important}
    [data-testid="stExpander"] details>summary *{color:#FFF8FD!important;-webkit-text-fill-color:#FFF8FD!important}
    [data-testid="stExpander"] details>summary:hover{background:linear-gradient(110deg,#5B1748 0%,#B43B91 58%,#EC64D5 100%)!important}
    [data-testid="stExpander"] [data-testid="stExpanderDetails"]{background:#FFF5FC!important;padding:10px!important}
    .ahl-table-wrap{width:100%;overflow-x:auto;border:1px solid #EC64D5;border-radius:12px;background:#FFF5FC;box-shadow:0 4px 14px rgba(91,23,72,.10);margin:8px 0 18px}
    .ahl-table-wrap table.ahl-data-table{width:100%;min-width:850px;border-collapse:separate!important;border-spacing:0!important;background:#FFF5FC!important;color:#3F1738!important;font-size:13px!important}
    .ahl-table-wrap table.ahl-data-table thead th{position:sticky;top:0;background:#B43B91!important;color:#FFF8FD!important;-webkit-text-fill-color:#FFF8FD!important;font-weight:800!important;text-align:left!important;padding:11px 12px!important;border:0!important;border-right:1px solid #E9B7DE!important;border-bottom:2px solid #8A246D!important;white-space:nowrap!important}
    .ahl-table-wrap table.ahl-data-table tbody td{background:#FFF5FC!important;color:#3F1738!important;-webkit-text-fill-color:#3F1738!important;padding:10px 12px!important;border:0!important;border-right:1px solid #EBD0E4!important;border-bottom:1px solid #EBD0E4!important;white-space:nowrap!important}
    .ahl-table-wrap table.ahl-data-table tbody tr:nth-child(even) td{background:#FCE8F7!important}
    .ahl-table-wrap table.ahl-data-table tbody tr:hover td{background:#F4C4E8!important;color:#5B1748!important;-webkit-text-fill-color:#5B1748!important}
    .ahl-table-wrap table.ahl-data-table tbody tr:last-child td{border-bottom:0!important}
    .ahl-table-wrap table.ahl-data-table th:last-child,.ahl-table-wrap table.ahl-data-table td:last-child{border-right:0!important}
    .ahl-table-wrap table.ahl-data-table tbody td:first-child{font-weight:800!important;color:#8A246D!important;-webkit-text-fill-color:#8A246D!important}

    .stApp{background:radial-gradient(ellipse at 8% 0%,#F7DDF1 0%,transparent 43%),linear-gradient(135deg,#FFF5FC 0%,#FCE8F7 52%,#FFFAFE 100%)!important;color:#3F1738!important;--brand-flower-color:#EC64D5!important;--signal-nav-border:#E7A5D6;--signal-nav-accent:#EC64D5;--signal-nav-soft:#F8DDF2;--signal-nav-canvas:#FFF5FC;--signal-nav-text:#5B1748}
    [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label{
      background:linear-gradient(135deg,#FFF8FD 0%,#FCE8F7 100%)!important;
      border:1px solid #E7A5D6!important;border-radius:10px!important;
      color:#5B1748!important;margin:3px 0!important;padding:8px 10px!important;
    }
    [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:hover{
      background:linear-gradient(110deg,#F8DDF2 0%,#FFF5FC 100%)!important;
      border-color:#EC64D5!important;
    }
    [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked){
      background:linear-gradient(110deg,#F4C4E8 0%,#FCE8F7 100%)!important;
      border:1.5px solid #EC64D5!important;box-shadow:inset 4px 0 0 #EC64D5!important;
    }
    [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label,
    [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label *{
      color:#5B1748!important;-webkit-text-fill-color:#5B1748!important;
    }
    [data-testid="stSidebar"] .stButton > button,
    [data-testid="stSidebar"] [data-testid="stFormSubmitButton"] button{
      border:1px solid #E7A5D6!important;border-radius:10px!important;
      color:#5B1748!important;background:linear-gradient(180deg,#FBEAF7 0%,#FFF5FC 55%,#FFF8FD 100%)!important;
      box-shadow:0 3px 10px rgba(91,23,72,.08)!important;
    }
    [data-testid="stSidebar"] .stButton > button:hover,
    [data-testid="stSidebar"] [data-testid="stFormSubmitButton"] button:hover{
      border-color:#EC64D5!important;background:linear-gradient(180deg,#F8DDF2 0%,#FCE8F7 100%)!important;
      color:#5B1748!important;
    }
    [data-testid="stDataFrame"],[data-testid="stTable"]{
      background:#FFF5FC!important;border:1px solid #EC64D5!important;
      border-radius:12px!important;overflow:hidden!important;
    }
    [data-testid="stHeader"]{background:rgba(255,245,252,.96)!important;border-bottom:1px solid #E7A5D6!important}
    [data-testid="stSidebar"]{background:linear-gradient(180deg,#FBEAF7 0%,#FFF5FC 55%,#FFF8FD 100%)!important;border-right:1.5px solid #E7A5D6!important}
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,[data-testid="stSidebar"] label,[data-testid="stSidebar"] [data-testid="stWidgetLabel"]{color:#5B1748!important}
    [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label,[data-testid="stSidebar"] [data-baseweb="select"]>div,[data-testid="stSidebar"] input{background:#FFF8FD!important;border-color:#E7A5D6!important;color:#3F1738!important}
    h1,h1 *{color:#EC64D5!important;-webkit-text-fill-color:#EC64D5!important}h2,h3,p,[data-testid="stMarkdownContainer"]{color:#3F1738}
    [data-testid="stMainBlockContainer"]{padding-top:3.5rem!important}.block-container{padding-top:3.5rem!important}h1{position:relative!important;text-align:center!important;white-space:nowrap!important;display:block!important;width:100%!important;font-size:2.5rem!important;line-height:1.15!important;margin-top:48px!important;padding-top:0!important}
    h1::before{content:"";position:absolute;left:50%;top:-28px;width:76px;height:52px;transform:translate(-50%,-50%);z-index:2;pointer-events:none;background:linear-gradient(110deg,var(--brand-flower-color,#EC64D5) 0%,var(--brand-flower-color,#EC64D5) 30%,rgba(255,255,255,.96) 48%,var(--brand-flower-color,#EC64D5) 66%,var(--brand-flower-color,#EC64D5) 100%);background-size:300% 300%;animation:dependentSignalShimmer 4.2s linear infinite;-webkit-mask:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 72' fill='none' stroke='white' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M60 58 C43 43 43 24 60 8 C77 24 77 43 60 58Z'/%3E%3Cpath d='M55 58 C34 51 22 35 27 20 C43 26 54 40 55 58Z'/%3E%3Cpath d='M65 58 C86 51 98 35 93 20 C77 26 66 40 65 58Z'/%3E%3Cpath d='M49 59 C30 60 16 52 10 40 C28 38 43 45 49 59Z'/%3E%3Cpath d='M71 59 C90 60 104 52 110 40 C92 38 77 45 71 59Z'/%3E%3Cpath d='M34 64 Q60 69 86 64'/%3E%3C/svg%3E") center/contain no-repeat;mask:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 72' fill='none' stroke='white' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M60 58 C43 43 43 24 60 8 C77 24 77 43 60 58Z'/%3E%3Cpath d='M55 58 C34 51 22 35 27 20 C43 26 54 40 55 58Z'/%3E%3Cpath d='M65 58 C86 51 98 35 93 20 C77 26 66 40 65 58Z'/%3E%3Cpath d='M49 59 C30 60 16 52 10 40 C28 38 43 45 49 59Z'/%3E%3Cpath d='M71 59 C90 60 104 52 110 40 C92 38 77 45 71 59Z'/%3E%3Cpath d='M34 64 Q60 69 86 64'/%3E%3C/svg%3E") center/contain no-repeat}
    @keyframes dependentSignalShimmer{0%{background-position:0% 100%;filter:brightness(1)}50%{background-position:100% 0%;filter:brightness(1.12)}100%{background-position:0% 100%;filter:brightness(1)}}
    [data-testid="stButton"] button,[data-testid="stDownloadButton"] button{border:1px solid #EC64D5!important;color:#FFF8FD!important;background:linear-gradient(110deg,#8A246D 0%,#EC64D5 58%,#EC64D5 100%)!important;border-radius:10px!important;font-weight:800!important}
    
    [data-testid="stDataFrame"],[data-testid="stTable"]{border:1px solid #E7A5D6!important;border-radius:12px!important}
    .ahl-hero{padding:20px 22px;margin:4px 0 18px;border-radius:18px;color:#FFF8FD;background:radial-gradient(circle at 90% 10%,rgba(255,255,255,.22),transparent 28%),linear-gradient(115deg,#5B1748 0%,#B43B91 48%,#EC64D5 100%);border:1px solid #E7A5D6;box-shadow:0 12px 28px rgba(91,23,72,.18)}
    .ahl-kicker{font-size:10px;font-weight:900;letter-spacing:2px;color:#F5C7E9}.ahl-title{font-size:clamp(27px,4vw,36px);font-weight:950;line-height:1.12;margin:7px 0;color:#FFF}.ahl-sub{font-size:12px;line-height:1.55;color:#F8E3F3}
    .ahl-card{box-sizing:border-box;display:flex;flex-direction:column;justify-content:space-between;width:100%;min-height:240px;margin:12px 0 18px;padding:23px 23px 22px;border-radius:21px;color:#FFF8FD;border:3px solid transparent;background:linear-gradient(125deg,#32102B 0%,#8A246D 52%,#5B1748 100%) padding-box,linear-gradient(110deg,#E7A5D6 0%,#FFF8FD 34%,#EC64D5 55%,#E7A5D6 100%) border-box;box-shadow:0 14px 30px rgba(91,23,72,.24)}
    .ahl-card-sell{color:#F7FBFF;background:linear-gradient(125deg,#123B70 0%,#287BC4 52%,#65A4EC 100%) padding-box,linear-gradient(110deg,#B8D9FF 0%,#FFFFFF 34%,#65A4EC 55%,#B8D9FF 100%) border-box;box-shadow:0 14px 30px rgba(28,91,157,.28)}
    .ahl-card-sell .ahl-brand{color:#D9EBFF;text-shadow:0 0 12px rgba(101,164,236,.65)}
    .ahl-card-sell .ahl-side{background:#EAF4FF;color:#174F85;border-color:#B8D9FF}
    .ahl-card-sell .ahl-label{color:#EAF4FF!important;text-shadow:0 1px 1px rgba(0,20,45,.45)}
    .ahl-card-sell .ahl-value,.ahl-card-sell .ahl-target{color:#EAF4FF!important;font-size:1.25rem;font-weight:950;text-shadow:0 1px 2px rgba(0,20,45,.55)}
    .ahl-card-sell .ahl-divider{background:linear-gradient(90deg,transparent,#65A4EC,rgba(234,244,255,.98),#65A4EC,transparent)}
    .ahl-brand{position:relative;z-index:1;text-align:center;font-size:11px;font-weight:900;letter-spacing:3px;color:#F5C7E9;margin:0 0 15px;text-shadow:0 0 12px rgba(236,100,213,.55)}.ahl-card-head{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;min-height:36px}.ahl-symbol{font-size:1.78rem;line-height:1.15;font-weight:950;color:#FFF}.ahl-side{font-size:.72rem;font-weight:950;letter-spacing:1.15px;padding:8px 12px;border-radius:99px;background:#FFF8FD;color:#8A246D;border:1px solid #E7A5D6}
    .ahl-divider{height:1px;width:100%;background:linear-gradient(90deg,transparent,#E7A5D6,rgba(255,247,237,.95),#E7A5D6,transparent);margin:19px 0 17px;flex:0 0 auto}
    .ahl-prices{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;margin-top:0}.ahl-label{font-size:.66rem;font-weight:900;letter-spacing:1.1px;color:#F5C7E9}.ahl-value{font-size:1.25rem;line-height:1.25;font-weight:950;color:#FFF;margin-top:7px;white-space:nowrap}.ahl-target{color:#FDE68A}.ahl-section{font-size:13px;font-weight:950;letter-spacing:1.2px;color:#8A246D;margin:18px 0 9px}
    @media(max-width:560px){.ahl-card{min-height:240px;padding:23px 20px 22px;border-radius:21px}.ahl-value{font-size:1.15rem}.ahl-symbol{font-size:1.55rem}}
    </style>""", unsafe_allow_html=True)
    st.title("6THSENSE-VARDAAN")


    now = datetime.now(ZoneInfo("Asia/Kolkata"))
    is_open = now.weekday() < 5 and 555 <= now.hour * 60 + now.minute <= 930
    st_autorefresh(interval=15000 if is_open else 60000, limit=None, key=f"ahl_refresh_{app_name.lower()}")

    try:
        symbols, caps, sheet_name = _read_master(master_bytes)
    except Exception as exc:
        st.error(f"Master Excel could not be read: {exc}")
        return

    with st.spinner(f"Downloading 20-day history and intraday prices for {len(symbols)} stocks..."):
        data_map = _market_data(tuple(symbols))
    if not data_map:
        st.warning("Yahoo Finance did not return usable data. The dashboard will retry automatically.")
        return

    rows, buy_events, sell_events = [], [], []
    for symbol in symbols:
        data = data_map.get(symbol)
        if not data:
            continue
        master = caps.get(symbol, {})
        buy_cap_pct = 1.50 if master.get("o2l4") is not None and master["o2l4"] < -3.00 else 1.00
        sell_cap_pct = 1.50 if master.get("o2h4") is not None and master["o2h4"] > 3.00 else 1.00
        day_low, day_high = data.get("low"), data.get("high")
        avg_low, avg_high = data["avg_low"], data["avg_high"]
        rows.append({
            "Stock": symbol, "Open": data.get("open"), "Today's High": day_high,
            "Today's Low": day_low, "Avg.Low": avg_low, "Avg.High": avg_high,
            "% Down Avg": ((day_low - avg_low) / avg_low * 100) if day_low is not None and avg_low else None,
            "% Above Avg": ((day_high - avg_high) / avg_high * 100) if day_high is not None and avg_high else None,
            "5 Lowest Lows": " · ".join(f"{v:.2f}" for v in data["lows5"]),
            "5 Highest Highs": " · ".join(f"{v:.2f}" for v in data["highs5"]),
        })
        for event in _events(symbol, data):
            entry, open_price = event["entry"], data.get("open")
            if event["side"] == "BUY":
                cap = open_price * (1 + buy_cap_pct / 100) if open_price else None
                event["target"] = min(entry * 1.0035, cap) if cap is not None else entry * 1.0035
                buy_events.append(event)
            else:
                cap = open_price * (1 - sell_cap_pct / 100) if open_price else None
                event["target"] = max(entry * .9965, cap) if cap is not None else entry * .9965
                sell_events.append(event)

    # When there are no BUY signals, the first active signal card is SELL.
    # Match the whole dashboard theme to that topmost Arctic 825 card.
    if not buy_events and sell_events:
        if app_name.upper() == "VARDAAN":
            st.session_state["_vardaan_active_theme"] = {
                "deep": "#123B70", "primary": "#65A4EC", "accent": "#65A4EC",
                "ink": "#123B70", "accent_text": "#174F85", "tint": "#DCEEFF",
                "canvas": "#F2F8FF", "border": "#B8D9FF", "soft_text": "#174F85",
                "left_accent": "#65A4EC",
            }
        st.markdown("""
        <style>
        .stApp{background:radial-gradient(ellipse at 8% 0%,#DCEEFF 0%,transparent 43%),linear-gradient(135deg,#F2F8FF 0%,#E5F2FF 52%,#FAFCFF 100%)!important;color:#123B70!important;--brand-flower-color:#65A4EC!important;--signal-nav-border:#B8D9FF;--signal-nav-accent:#65A4EC;--signal-nav-soft:#DCEEFF;--signal-nav-canvas:#F2F8FF;--signal-nav-text:#174F85}
        [data-testid="stHeader"]{background:rgba(242,248,255,.96)!important;border-bottom:1px solid #B8D9FF!important}
        [data-testid="stSidebar"]{background:linear-gradient(180deg,#EAF4FF 0%,#F2F8FF 55%,#FAFCFF 100%)!important;border-right:1.5px solid #B8D9FF!important}
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,[data-testid="stSidebar"] label,[data-testid="stSidebar"] [data-testid="stWidgetLabel"]{color:#174F85!important}
        [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label,[data-testid="stSidebar"] [data-baseweb="select"]>div,[data-testid="stSidebar"] input{background:#FAFCFF!important;border-color:#B8D9FF!important;color:#123B70!important}
        h1,h1 *{color:#2789CC!important;-webkit-text-fill-color:#2789CC!important}h2,h3,p,[data-testid="stMarkdownContainer"]{color:#123B70}

        [data-testid="stButton"] button,[data-testid="stDownloadButton"] button{border:1px solid #65A4EC!important;color:#FFFFFF!important;background:linear-gradient(110deg,#174F85 0%,#2789CC 58%,#65A4EC 100%)!important}
        /* Topmost SELL card selects the Arctic 825 palette for the entire sidebar, as in the other dashboards. */
        html body .stApp [data-testid="stSidebar"],
        html body .stApp section[data-testid="stSidebar"]{
          background:linear-gradient(180deg,#DCEEFF 0%,#EAF4FF 42%,#F2F8FF 100%)!important;
          border-right:1.5px solid #65A4EC!important;
          --sidebar-theme-deep:#123B70;--sidebar-theme-mid:#287BC4;--sidebar-theme-main:#65A4EC;
          --sidebar-theme-border:#65A4EC;--sidebar-theme-text:#174F85;
        }
        html body .stApp [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
        html body .stApp [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
        html body .stApp [data-testid="stSidebar"] label,
        html body .stApp [data-testid="stSidebar"] [data-testid="stWidgetLabel"]{
          color:#174F85!important;-webkit-text-fill-color:#174F85!important;
        }
        /* Arctic 825 selector/radio choices: keep them light to match the sidebar surface. */
        html body .stApp [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label,
        html body .stApp [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label *{
          color:#174F85!important;-webkit-text-fill-color:#174F85!important;
        }
        html body .stApp [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label{
          background:linear-gradient(135deg,#DCEEFF 0%,#EAF4FF 55%,#F2F8FF 100%)!important;
          border:1.5px solid #B8D9FF!important;border-radius:10px!important;
        }
        html body .stApp [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked),
        html body .stApp [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:has([aria-checked="true"]){
          background:linear-gradient(135deg,#D6EAFF 0%,#E5F2FF 55%,#F2F8FF 100%)!important;
          border:1.5px solid #65A4EC!important;
          box-shadow:inset 4px 0 0 #65A4EC,0 2px 7px rgba(18,59,112,.08)!important;
          color:#174F85!important;-webkit-text-fill-color:#174F85!important;
        }
        html body .stApp [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) *,
        html body .stApp [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:has([aria-checked="true"]) *{
          color:#174F85!important;-webkit-text-fill-color:#174F85!important;
        }
        html body .stApp [data-testid="stSidebar"] [data-baseweb="select"]>div,
        html body .stApp [data-testid="stSidebar"] [data-baseweb="select"] [role="combobox"]{
          background:linear-gradient(135deg,#DCEEFF 0%,#EAF4FF 55%,#F2F8FF 100%)!important;
          border:1.5px solid #B8D9FF!important;
          color:#174F85!important;-webkit-text-fill-color:#174F85!important;
        }
        html body .stApp [data-testid="stSidebar"] [data-baseweb="select"] *{
          color:#174F85!important;-webkit-text-fill-color:#174F85!important;
        }
        /* Unlock Master Control and Log out use the same light Arctic sidebar styling as selectors. */
        html body .stApp [data-testid="stSidebar"] .stButton > button,
        html body .stApp [data-testid="stSidebar"] [data-testid="stFormSubmitButton"] button,
        html body .stApp [data-testid="stSidebar"] [data-testid^="stBaseButton-"],
        html body .stApp [data-testid="stSidebar"] button[kind]{
          border:1.5px solid #B8D9FF!important;border-radius:10px!important;
          color:#174F85!important;-webkit-text-fill-color:#174F85!important;
          background:linear-gradient(135deg,#DCEEFF 0%,#EAF4FF 55%,#F2F8FF 100%)!important;
          box-shadow:0 2px 7px rgba(18,59,112,.08)!important;
          font-weight:850!important;
        }
        html body .stApp [data-testid="stSidebar"] .stButton > button *,
        html body .stApp [data-testid="stSidebar"] [data-testid="stFormSubmitButton"] button *,
        html body .stApp [data-testid="stSidebar"] [data-testid^="stBaseButton-"] *{
          color:#174F85!important;-webkit-text-fill-color:#174F85!important;
        }
        html body .stApp [data-testid="stSidebar"] .stButton > button:hover,
        html body .stApp [data-testid="stSidebar"] [data-testid="stFormSubmitButton"] button:hover,
        html body .stApp [data-testid="stSidebar"] [data-testid^="stBaseButton-"]:hover,
        html body .stApp [data-testid="stSidebar"] button[kind]:hover{
          border-color:#65A4EC!important;
          background:linear-gradient(135deg,#D6EAFF 0%,#E5F2FF 55%,#F2F8FF 100%)!important;
          color:#174F85!important;-webkit-text-fill-color:#174F85!important;
          box-shadow:inset 4px 0 0 #65A4EC,0 3px 9px rgba(18,59,112,.12)!important;
        }
        /* SELL flower: soft flowing diagonal white shimmer with a gentle glow. */
        @keyframes dependentSignalSellSweep{
          0%{background-position:0% 100%;filter:brightness(1) drop-shadow(0 0 0 rgba(255,255,255,0))}
          50%{filter:brightness(1.06) drop-shadow(0 0 3px rgba(255,255,255,.28))}
          100%{background-position:100% 0%;filter:brightness(1) drop-shadow(0 0 0 rgba(255,255,255,0))}
        }
        @-webkit-keyframes dependentSignalSellSweep{
          0%{background-position:0% 100%;-webkit-filter:brightness(1) drop-shadow(0 0 0 rgba(255,255,255,0))}
          50%{-webkit-filter:brightness(1.06) drop-shadow(0 0 3px rgba(255,255,255,.28))}
          100%{background-position:100% 0%;-webkit-filter:brightness(1) drop-shadow(0 0 0 rgba(255,255,255,0))}
        }
        h1::before{
          background:linear-gradient(110deg,var(--brand-flower-color,#65A4EC) 0%,var(--brand-flower-color,#65A4EC) 36%,rgba(255,255,255,.28) 45%,rgba(255,255,255,.88) 52%,rgba(255,255,255,.28) 59%,var(--brand-flower-color,#65A4EC) 68%,var(--brand-flower-color,#65A4EC) 100%)!important;
          background-size:350% 350%!important;
          animation:dependentSignalSellSweep 5.2s ease-in-out infinite!important;
          -webkit-animation:dependentSignalSellSweep 5.2s ease-in-out infinite!important;
        }
        [data-testid="stExpander"]{border:1.5px solid #65A4EC!important;background:#F2F8FF!important}
        [data-testid="stExpander"] details>summary{background:linear-gradient(110deg,#174F85 0%,#2789CC 58%,#65A4EC 100%)!important;color:#FFFFFF!important}
        [data-testid="stExpander"] details>summary:hover{background:linear-gradient(110deg,#123B70 0%,#2789CC 58%,#65A4EC 100%)!important}
        [data-testid="stExpander"] [data-testid="stExpanderDetails"]{background:#F2F8FF!important}
        .ahl-table-wrap{border-color:#65A4EC;background:#F2F8FF}
        .ahl-table-wrap table.ahl-data-table{background:#F2F8FF!important;color:#123B70!important}
        .ahl-table-wrap table.ahl-data-table thead th{background:#287BC4!important;color:#FFFFFF!important;-webkit-text-fill-color:#FFFFFF!important;border-right-color:#B8D9FF!important;border-bottom-color:#174F85!important}
        .ahl-table-wrap table.ahl-data-table tbody td{background:#F2F8FF!important;color:#123B70!important;-webkit-text-fill-color:#123B70!important;border-right-color:#D5E8FA!important;border-bottom-color:#D5E8FA!important}
        .ahl-table-wrap table.ahl-data-table tbody tr:nth-child(even) td{background:#EAF4FF!important}
        .ahl-table-wrap table.ahl-data-table tbody tr:hover td{background:#D6EAFF!important;color:#123B70!important;-webkit-text-fill-color:#123B70!important}
        .ahl-table-wrap table.ahl-data-table tbody td:first-child{color:#174F85!important;-webkit-text-fill-color:#174F85!important}
        [data-testid="stDataFrame"],[data-testid="stTable"]{border-color:#B8D9FF!important}
        .ahl-section{color:#174F85!important}
        </style>
        """, unsafe_allow_html=True)

    def render_cards(events, side):
        events.sort(key=lambda item: item["time"], reverse=True)
        if not events:
            st.info(f"No {side} crossing signal yet.")
            return
        for event in events:
            st.markdown(
                f'<div class="ahl-card{" ahl-card-sell" if side == "SELL" else ""}"><div class="ahl-brand">6THSENSE · VARDAAN</div><div class="ahl-card-head"><div class="ahl-symbol">{event["symbol"]}</div><div class="ahl-side">{side} SIGNAL</div></div>'
                f'<div class="ahl-divider"></div>'
                f'<div class="ahl-prices"><div><div class="ahl-label">ENTRY PRICE</div><div class="ahl-value">{_price(event["entry"])}</div></div>'
                f'<div><div class="ahl-label">TARGET PRICE</div><div class="ahl-value ahl-target">{_price(event["target"])}</div></div></div></div>',
                unsafe_allow_html=True,
            )

    st.markdown('<div class="ahl-section">BUY SIGNALS</div>', unsafe_allow_html=True)
    render_cards(buy_events, "BUY")
    st.markdown('<div class="ahl-section">SELL SIGNALS</div>', unsafe_allow_html=True)
    render_cards(sell_events, "SELL")

    with st.expander("TODAY'S MARKET DATA", expanded=False):
        if rows:
            frame = pd.DataFrame(rows)
            for col in ("Open", "Today's High", "Today's Low", "Avg.Low", "Avg.High"):
                frame[col] = frame[col].map(_price)
            for col in ("% Down Avg", "% Above Avg"):
                frame[col] = frame[col].map(lambda value: "—" if pd.isna(value) else f"{value:+.2f}%")
            # Keep the five lowest Lows and five highest Highs as the final columns.
            order = ["Stock", "Open", "Today's High", "Today's Low", "Avg.Low", "Avg.High",
                     "% Down Avg", "% Above Avg", "5 Lowest Lows", "5 Highest Highs"]
            table_html = frame[order].to_html(index=False, escape=True, classes="ahl-data-table", border=0)
            st.markdown(f'<div class="ahl-table-wrap">{table_html}</div>', unsafe_allow_html=True)
        else:
            st.warning("No stocks have complete 20-session history and intraday data yet.")
