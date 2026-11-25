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
            "deep": "#7C2D12", "primary": "#EA580C", "accent": "#F97316",
            "ink": "#54240D", "accent_text": "#9A3412", "tint": "#FFE2C2",
            "canvas": "#FFF4E8", "border": "#FDBA74", "soft_text": "#7C2D12",
            "left_accent": "#EA580C",
        }
    st.markdown("""
    <style>
    .stApp{background:radial-gradient(ellipse at 8% 0%,#FFE0BE 0%,transparent 43%),linear-gradient(135deg,#FFF4E8 0%,#FFE8D1 52%,#FFFAF5 100%)!important;color:#54240D!important;--brand-flower-color:#F97316!important;--signal-nav-border:#FDBA74;--signal-nav-accent:#EA580C;--signal-nav-soft:#FFE2C2;--signal-nav-canvas:#FFF4E8;--signal-nav-text:#7C2D12}
    [data-testid="stHeader"]{background:rgba(255,244,232,.96)!important;border-bottom:1px solid #FDBA74!important}
    [data-testid="stSidebar"]{background:linear-gradient(180deg,#FFEDD9 0%,#FFF4E8 55%,#FFF9F2 100%)!important;border-right:1.5px solid #FDBA74!important}
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,[data-testid="stSidebar"] label,[data-testid="stSidebar"] [data-testid="stWidgetLabel"]{color:#7C2D12!important}
    [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label,[data-testid="stSidebar"] [data-baseweb="select"]>div,[data-testid="stSidebar"] input{background:#FFF9F2!important;border-color:#FDBA74!important;color:#54240D!important}
    h1,h2,h3,p,[data-testid="stMarkdownContainer"]{color:#54240D}
    [data-testid="stButton"] button,[data-testid="stDownloadButton"] button{border:1px solid #EA580C!important;color:#FFF7ED!important;background:linear-gradient(110deg,#9A3412 0%,#EA580C 58%,#F97316 100%)!important;border-radius:10px!important;font-weight:800!important}
    [data-testid="stDataFrame"],[data-testid="stTable"]{border:1px solid #FDBA74!important;border-radius:12px!important}
    .ahl-hero{padding:20px 22px;margin:4px 0 18px;border-radius:18px;color:#FFF7ED;background:radial-gradient(circle at 90% 10%,rgba(255,255,255,.22),transparent 28%),linear-gradient(115deg,#7C2D12 0%,#C2410C 48%,#EA580C 100%);border:1px solid #FDBA74;box-shadow:0 12px 28px rgba(124,45,18,.18)}
    .ahl-kicker{font-size:10px;font-weight:900;letter-spacing:2px;color:#FED7AA}.ahl-title{font-size:clamp(27px,4vw,36px);font-weight:950;line-height:1.12;margin:7px 0;color:#FFF}.ahl-sub{font-size:12px;line-height:1.55;color:#FFEDD5}
    .ahl-card{margin:8px 0 14px;padding:18px 20px;border-radius:18px;color:#FFF7ED;border:2px solid transparent;background:linear-gradient(125deg,#431407 0%,#9A3412 52%,#7C2D12 100%) padding-box,linear-gradient(110deg,#FDBA74 0%,#FFF7ED 34%,#F97316 55%,#FDBA74 100%) border-box;box-shadow:0 10px 26px rgba(124,45,18,.20)}
    .ahl-card-head{display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap}.ahl-symbol{font-size:1.5rem;font-weight:950;color:#FFF}.ahl-side{font-size:.68rem;font-weight:950;padding:7px 11px;border-radius:99px;background:#FFF7ED;color:#9A3412;border:1px solid #FDBA74}
    .ahl-prices{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;margin-top:16px}.ahl-label{font-size:.65rem;font-weight:900;letter-spacing:1.1px;color:#FED7AA}.ahl-value{font-size:1.35rem;font-weight:950;color:#FFF;margin-top:5px;white-space:nowrap}.ahl-target{color:#FDE68A}.ahl-section{font-size:13px;font-weight:950;letter-spacing:1.2px;color:#9A3412;margin:18px 0 9px}
    @media(max-width:560px){.ahl-card{padding:15px}.ahl-value{font-size:1.12rem}.ahl-symbol{font-size:1.3rem}}
    </style>""", unsafe_allow_html=True)
    st.title("Avg.High Low")
    st.markdown(f'<div class="ahl-hero"><div class="ahl-kicker">{app_name} · INTRADAY REVERSAL SCANNER</div><div class="ahl-title">Avg.High Low</div><div class="ahl-sub">20 completed trading sessions · Five lowest Lows and five highest Highs per stock · Live recovery crossings</div></div>', unsafe_allow_html=True)

    now = datetime.now(ZoneInfo("Asia/Kolkata"))
    is_open = now.weekday() < 5 and 555 <= now.hour * 60 + now.minute <= 930
    st_autorefresh(interval=15000 if is_open else 60000, limit=None, key=f"ahl_refresh_{app_name.lower()}")

    try:
        symbols, caps, sheet_name = _read_master(master_bytes)
    except Exception as exc:
        st.error(f"Master Excel could not be read: {exc}")
        return
    st.caption(f"Universe: {len(symbols)} Master Excel stocks · Sheet: {sheet_name} · Historical source: Yahoo Finance")
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

    def render_cards(events, side):
        events.sort(key=lambda item: item["time"], reverse=True)
        if not events:
            st.info(f"No {side} crossing signal yet.")
            return
        for event in events:
            st.markdown(
                f'<div class="ahl-card"><div class="ahl-card-head"><div class="ahl-symbol">{event["symbol"]}</div><div class="ahl-side">{side} SIGNAL</div></div>'
                f'<div class="ahl-prices"><div><div class="ahl-label">ENTRY PRICE</div><div class="ahl-value">{_price(event["entry"])}</div></div>'
                f'<div><div class="ahl-label">TARGET PRICE</div><div class="ahl-value ahl-target">{_price(event["target"])}</div></div></div></div>',
                unsafe_allow_html=True,
            )

    st.markdown('<div class="ahl-section">BUY SIGNALS</div>', unsafe_allow_html=True)
    render_cards(buy_events, "BUY")
    st.markdown('<div class="ahl-section">SELL SIGNALS</div>', unsafe_allow_html=True)
    render_cards(sell_events, "SELL")

    st.markdown('<div class="ahl-section">TODAY\'S MARKET DATA</div>', unsafe_allow_html=True)
    if rows:
        frame = pd.DataFrame(rows)
        for col in ("Open", "Today's High", "Today's Low", "Avg.Low", "Avg.High"):
            frame[col] = frame[col].map(_price)
        for col in ("% Down Avg", "% Above Avg"):
            frame[col] = frame[col].map(lambda value: "—" if pd.isna(value) else f"{value:+.2f}%")
        # Keep the five lowest Lows and five highest Highs as the final columns.
        order = ["Stock", "Open", "Today's High", "Today's Low", "Avg.Low", "Avg.High",
                 "% Down Avg", "% Above Avg", "5 Lowest Lows", "5 Highest Highs"]
        st.dataframe(frame[order], hide_index=True, use_container_width=True)
    else:
        st.warning("No stocks have complete 20-session history and intraday data yet.")
