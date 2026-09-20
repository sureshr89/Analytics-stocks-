import streamlit as st
import pandas as pd
import numpy as np
import sqlite3, hashlib, io, re, os
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Trading Journal", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

DB = "trading_journal_clean.db"

# ---------- Mobile-first visual design ----------
st.markdown("""
<style>
:root { --bg:#0b1020; --card:#151c2f; --muted:#94a3b8; --text:#f8fafc; }
.block-container { max-width: 1500px; padding: 2rem 1.1rem 3rem; }
h1 { font-size: clamp(1.65rem, 4vw, 2.5rem) !important; margin-bottom:.15rem !important; }
h2 { font-size: 1.35rem !important; }
h3 { font-size: 1.05rem !important; }
p, label, .stCaption { font-size: .86rem !important; }
[data-testid="stMetric"] { background:linear-gradient(145deg,#151c2f,#10172a); border:1px solid #27324d; border-radius:14px; padding:.75rem .85rem; min-height:92px; }
[data-testid="stMetricLabel"] { font-size:.72rem !important; color:#94a3b8 !important; }
[data-testid="stMetricValue"] { font-size:1.25rem !important; font-weight:750 !important; }
[data-testid="stMetricDelta"] { font-size:.7rem !important; }
.stTabs [data-baseweb="tab"] { font-size:.78rem; padding:.55rem .65rem; }
.stButton button, .stDownloadButton button { border-radius:10px; font-weight:700; }
div[data-testid="stDataFrame"] { font-size:.72rem; }
[data-testid="stFileUploader"] { border-radius:12px; }
@media (max-width: 700px) {
  .block-container { padding:1.65rem .55rem 2rem; }
  [data-testid="stMetric"] { min-height:82px; padding:.6rem .65rem; }
  [data-testid="stMetricValue"] { font-size:1.05rem !important; }
  .stTabs [data-baseweb="tab"] { font-size:.7rem; padding:.45rem .5rem; }
  [data-testid="stHorizontalBlock"] { gap:.45rem; }
  /* Force only metric/card rows into a true 2-column mobile grid */
  [data-testid="stHorizontalBlock"]:has([data-testid="stMetric"]) {
    display:grid !important;
    grid-template-columns:repeat(2,minmax(0,1fr)) !important;
    grid-auto-flow:row !important;
    width:100% !important;
    gap:.45rem !important;
  }
  [data-testid="stHorizontalBlock"]:has([data-testid="stMetric"]) > div[data-testid="column"] {
    box-sizing:border-box !important;
    width:100% !important;
    max-width:none !important;
    min-width:0 !important;
    flex:none !important;
  }
  [data-testid="stHorizontalBlock"]:has([data-testid="stMetric"]) [data-testid="stMetric"] {
    width:100% !important;
    box-sizing:border-box !important;
  }
  .stPlotlyChart, .js-plotly-plot { width:100% !important; max-width:100% !important; max-height:none !important; }
  .plot-container, .svg-container { width:100% !important; }
  .js-plotly-plot .plotly { width:100% !important; }
  .stPlotlyChart, .stPlotlyChart * { pointer-events: none !important; }
  /* Prevent browser/Plotly touch gestures from turning a chart into a zoom surface */
  .stPlotlyChart,
  .stPlotlyChart *,
  .js-plotly-plot,
  .js-plotly-plot *,
  .plotly,
  .plot-container,
  .svg-container {
    touch-action: none !important;
    overscroll-behavior: none !important;
  }
}
</style>
""", unsafe_allow_html=True)

def conn():
    c=sqlite3.connect(DB)
    c.execute("""CREATE TABLE IF NOT EXISTS trades(
      trade_hash TEXT PRIMARY KEY, source_hash TEXT, asset_class TEXT, symbol TEXT,
      instrument TEXT, option_type TEXT, strike REAL, expiry TEXT, qty REAL,
      buy_date TEXT, buy_price REAL, buy_value REAL, sell_date TEXT, sell_price REAL,
      sell_value REAL, pnl REAL, remark TEXT, uploaded_at TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS sources(
      source_hash TEXT PRIMARY KEY, filename TEXT, asset_class TEXT,
      period_start TEXT, period_end TEXT, uploaded_at TEXT, trade_count INTEGER,
      gross_pnl REAL, report_gross_pnl REAL, report_charges REAL)""")
    scols={r[1] for r in c.execute("pragma table_info(sources)").fetchall()}
    if "report_gross_pnl" not in scols: c.execute("alter table sources add column report_gross_pnl REAL")
    if "report_charges" not in scols: c.execute("alter table sources add column report_charges REAL")
    if "report_period_start" not in scols: c.execute("alter table sources add column report_period_start TEXT")
    if "report_period_end" not in scols: c.execute("alter table sources add column report_period_end TEXT")
    c.execute("""CREATE TABLE IF NOT EXISTS charges(
      source_hash TEXT, asset_class TEXT, period_start TEXT, period_end TEXT,
      charge_name TEXT, amount REAL, PRIMARY KEY(source_hash,charge_name))""")
    # Keep two different concepts separate:
    #   * sources.period_* = actual trade coverage in the uploaded file
    #   * sources.report_period_* = date range printed by the broker report
    # Charges belong to the broker report period, not to the min/max trade
    # dates. Mixing those periods can make a full-period charge total look
    # like a one-day charge and incorrectly produce a Last Traded Day net P&L.
    c.execute("""
        update charges
           set period_start=(
               select report_period_start from sources s
                where s.source_hash=charges.source_hash
           ),
               period_end=(
               select report_period_end from sources s
                where s.source_hash=charges.source_hash
           )
         where exists(
               select 1 from sources s
                where s.source_hash=charges.source_hash
                  and s.report_period_start is not null
                  and s.report_period_end is not null
           )
    """)
    c.commit(); return c

def parse_date(v):
    x=pd.to_datetime(v,errors="coerce",dayfirst=True)
    return None if pd.isna(x) else x.strftime("%Y-%m-%d %H:%M:%S")

def period_from(raw):
    text=" ".join(raw.astype(str).fillna("").head(20).values.flatten())
    dates=[]
    for d in re.findall(r'(\d{1,2}[- /]\w{3,9}[- /]\d{2,4}|\d{1,2}[-/]\d{1,2}[-/]\d{4})',text):
        x=pd.to_datetime(d,errors="coerce",dayfirst=True)
        if not pd.isna(x): dates.append(x)
    return (min(dates).strftime("%Y-%m-%d"),max(dates).strftime("%Y-%m-%d")) if len(dates)>=2 else (None,None)

def option_fields(s):
    s=str(s)
    opt="CE" if re.search(r"\bCall\b",s,re.I) else ("PE" if re.search(r"\bPut\b",s,re.I) else None)
    m=re.search(r'\s(\d+(?:\.\d+)?)\s+(?:Call|Put)\s*$',s,re.I)
    strike=float(m.group(1)) if m else None
    m=re.search(r'\b(\d{1,2}\s+[A-Z]{3}\s+\d{2})\b',s,re.I)
    expiry=None
    if m:
        x=pd.to_datetime(m.group(1),format="%d %b %y",errors="coerce")
        if not pd.isna(x): expiry=x.strftime("%Y-%m-%d")
    return opt,strike,expiry

def extract(uploaded,filename):
    b=uploaded.getvalue(); sh=hashlib.sha256(b).hexdigest()
    xls=pd.ExcelFile(io.BytesIO(b)); raw0=pd.read_excel(io.BytesIO(b),sheet_name=xls.sheet_names[0],header=None)
    report_ps,report_pe=period_from(raw0)
    ps,pe=report_ps,report_pe
    lowname=filename.lower()
    asset="Commodities" if "commodit" in lowname else ("Stocks" if "stocks" in lowname else "F&O")
    trade=None
    for shn in xls.sheet_names:
        raw=pd.read_excel(io.BytesIO(b),sheet_name=shn,header=None)
        header=None
        for i,row in raw.iterrows():
            vals=[str(x).strip().lower() for x in row.tolist()]
            if "buy date" in vals and "sell date" in vals: header=i; break
        if header is not None:
            d=raw.iloc[header+1:].copy()
            d.columns=[str(x).strip() if not pd.isna(x) else f"col_{j}" for j,x in enumerate(raw.iloc[header])]
            first=d.columns[0]; d=d[d[first].notna()]
            d=d[~d[first].astype(str).str.contains("total|unrealised trades|disclaimer|groww",case=False,na=False)]
            trade=d; break
    rows=[]
    if trade is not None:
        cols={str(c).lower():c for c in trade.columns}
        def col(keys):
            for k in keys:
                for lc,c in cols.items():
                    if k in lc:return c
        cs=col(["stock name","scrip name"]); cq=col(["quantity"]); cbd=col(["buy date"]); cbp=col(["buy price"])
        cbv=col(["buy value"]); csd=col(["sell date"]); csp=col(["sell price"]); csv=col(["sell value"])
        cp=col(["realized p&l","realised p&l"]); cr=col(["remark"])
        for _,r in trade.iterrows():
            symbol=str(r[cs]).strip() if cs else ""; pnl=pd.to_numeric(r[cp],errors="coerce") if cp else np.nan
            if not symbol or pd.isna(pnl): continue
            opt,strike,expiry=option_fields(symbol)
            instrument="Options" if opt else ("Futures" if re.search(r"\bFut\b",symbol,re.I) else "Stocks")
            vals=[asset,symbol,r[cq],r[cbd],r[cbp],r[csd],r[csp],pnl]
            rows.append(dict(trade_hash=hashlib.sha256("|".join(map(str,vals)).encode()).hexdigest(),
              source_hash=sh,asset_class=asset,symbol=symbol,instrument=instrument,option_type=opt,
              strike=strike,expiry=expiry,qty=pd.to_numeric(r[cq],errors="coerce"),
              buy_date=parse_date(r[cbd]),buy_price=pd.to_numeric(r[cbp],errors="coerce"),
              buy_value=pd.to_numeric(r[cbv],errors="coerce"),sell_date=parse_date(r[csd]),
              sell_price=pd.to_numeric(r[csp],errors="coerce"),sell_value=pd.to_numeric(r[csv],errors="coerce"),
              pnl=float(pnl),remark=str(r[cr]) if cr and not pd.isna(r[cr]) else "",
              uploaded_at=datetime.now().isoformat(timespec="seconds")))
    report_gross=None; report_charges=None; in_realized=False
    for _,row in raw0.iterrows():
        label=str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else ""; ll=label.lower()
        if ll=="realised p&l":
            in_realized=True
            v=pd.to_numeric(row.iloc[1],errors="coerce") if len(row)>1 else np.nan
            if not pd.isna(v): report_gross=float(v); in_realized=False
        elif in_realized and ll=="total":
            v=pd.to_numeric(row.iloc[1],errors="coerce") if len(row)>1 else np.nan
            if not pd.isna(v): report_gross=float(v)
            in_realized=False
    charges=[]; in_ch=False
    for _,row in raw0.iterrows():
        label=str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else ""; ll=label.lower()
        if ll=="charges": in_ch=True; continue
        if in_ch and label=="": break
        if in_ch:
            amt=pd.to_numeric(row.iloc[1],errors="coerce") if len(row)>1 else np.nan
            if not pd.isna(amt) and (any(k in ll for k in ["exchange transaction","sebi","stt","ctt","stamp duty","ipft","brokerage","gst","dp charges","mis charges"]) or ll=="total"):
                charges.append((label,float(amt)))
            if ll=="total" and not pd.isna(amt): report_charges=float(amt)
    # Prefer the actual completed-trade period over dates found in
    # report headers. This avoids treating a Last Trading Day report as a
    # financial-year report simply because the header mentions 01 Apr.
    trade_dates=[]
    for rr in rows:
        for key in ("buy_date","sell_date"):
            if rr.get(key):
                trade_dates.append(pd.to_datetime(rr[key],errors="coerce"))
    trade_dates=[d for d in trade_dates if not pd.isna(d)]
    if trade_dates:
        ps=min(trade_dates).strftime("%Y-%m-%d")
        pe=max(trade_dates).strftime("%Y-%m-%d")

    return sh,asset,ps,pe,rows,charges,report_gross,report_charges,report_ps,report_pe

def save(uploaded,filename):
    sh,asset,ps,pe,rows,charges,report_gross,report_charges,report_ps,report_pe=extract(uploaded,filename); c=conn()
    existing=c.execute(
        'select source_hash,period_start,period_end,report_period_start,report_period_end '
        'from sources where asset_class=?',(asset,)
    ).fetchall()
    replace_hashes=[]

    # A newly uploaded broker report is authoritative for its printed report
    # period. This is important when the report contains no trades on some
    # later dates: stale rows from an older upload must not survive outside
    # the new file's actual trade rows.
    if any(old_hash==sh for old_hash,*_ in existing):
        replace_hashes.append(sh)

    if report_ps and report_pe:
        rps=pd.to_datetime(report_ps).normalize()
        rpe=pd.to_datetime(report_pe).normalize()
        for old_hash,ops,ope,orps,orpe in existing:
            if old_hash in replace_hashes:
                continue
            if orps and orpe:
                ors=pd.to_datetime(orps).normalize()
                ore=pd.to_datetime(orpe).normalize()
            elif ops and ope:
                ors=pd.to_datetime(ops).normalize()
                ore=pd.to_datetime(ope).normalize()
            else:
                continue
            # Replace any older report fully contained in the new report,
            # or an older report with the same report range. This removes
            # stale trades that are absent from the newly supplied file.
            if rps<=ors and ore<=rpe:
                replace_hashes.append(old_hash)
            elif ors==rps and ore==rpe:
                replace_hashes.append(old_hash)

    # Do not use "already covered" here: coverage based on min/max trade
    # dates can incorrectly preserve stale rows when a broker report's
    # printed period is wider than its actual trade rows.
    if replace_hashes:
        q=','.join('?'*len(replace_hashes))
        c.execute(f'delete from trades where source_hash in ({q})',replace_hashes)
        c.execute(f'delete from charges where source_hash in ({q})',replace_hashes)
        c.execute(f'delete from sources where source_hash in ({q})',replace_hashes)
    # Clear all old trades inside the broker report's stated period before
    # inserting the current file. This prevents an older upload from making
    # Last Trading Day appear later than the latest trade in the new file.
    if report_ps and report_pe:
        c.execute(
            'delete from trades where asset_class=? and sell_date>=? and sell_date<=?',
            (asset,report_ps,report_pe)
        )
    elif ps and pe:
        c.execute(
            'delete from trades where asset_class=? and sell_date>=? and sell_date<=?',
            (asset,ps,pe)
        )
    for r in rows: c.execute('insert or ignore into trades values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',tuple(r[k] for k in ['trade_hash','source_hash','asset_class','symbol','instrument','option_type','strike','expiry','qty','buy_date','buy_price','buy_value','sell_date','sell_price','sell_value','pnl','remark','uploaded_at']))
    # Charges are report-level figures. Store the broker's printed report
    # period, never the min/max dates of the individual trade rows.
    charge_ps,charge_pe=report_ps,report_pe
    for label,amt in charges:
        c.execute(
            'insert or replace into charges values(?,?,?,?,?,?)',
            (sh,asset,charge_ps,charge_pe,label,amt)
        )
    c.execute('insert into sources(source_hash,filename,asset_class,period_start,period_end,uploaded_at,trade_count,gross_pnl,report_gross_pnl,report_charges,report_period_start,report_period_end) values(?,?,?,?,?,?,?,?,?,?,?,?)',(sh,filename,asset,ps,pe,datetime.now().isoformat(timespec='seconds'),len(rows),sum(r['pnl'] for r in rows),report_gross,report_charges,report_ps,report_pe))
    c.commit(); return len(rows),'Reconciled & imported',asset

def money(x): return f"₹{float(x):,.2f}"
def pct(x): return f"{float(x):.2f}%"

def chart_layout(fig, height=300):
    fig.update_layout(
        height=height, margin=dict(l=10,r=10,t=45,b=10),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(size=11), hovermode="x unified",
        legend=dict(orientation="h",y=1.08,x=0)
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="rgba(148,163,184,.12)", zerolinecolor="rgba(148,163,184,.25)")
    return fig

st.title("📊 Trading Journal")
st.caption("Current-year trading analysis • Realised P&L • Charges • Net P&L • cumulative profit/loss")

with st.sidebar:
    st.header("📤 Upload EOD")
    uploads=st.file_uploader("Stocks / Equity F&O / Commodities",type=["xlsx","xls"],accept_multiple_files=True)
    if uploads and st.button("⚡ Process & Update",type="primary",use_container_width=True):
        for u in uploads:
            n,msg,a=save(u,u.name); st.success(f"{a}: {msg} • {n} trades")
        st.rerun()
    st.markdown("---")
    st.caption("Only the current calendar year is shown.")

df=pd.read_sql_query("select * from trades",conn())
ch=pd.read_sql_query("select * from charges",conn())
if df.empty:
    st.info("📤 Upload your first EOD Excel from the left sidebar.")
    st.stop()

df["sell_date"]=pd.to_datetime(df.sell_date,errors="coerce").astype("datetime64[ns]")
df["buy_date"]=pd.to_datetime(df.buy_date,errors="coerce").astype("datetime64[ns]")
df["win"]=df.pnl>0
df["loss"]=df.pnl<0

current_year=datetime.now().year
df=df[df.sell_date.dt.year==current_year].copy()
if df.empty:
    st.info(f"No completed trades are loaded for {current_year}. Upload this year's EOD files.")
    st.stop()

ch["period_end"]=pd.to_datetime(ch["period_end"],errors="coerce").astype("datetime64[ns]")
ch_year=ch[(ch.period_end.dt.year==current_year) & (ch.charge_name.str.lower()=="total")].copy()

with st.sidebar:
    st.header("🔎 Filters")
    asset=st.selectbox("Asset",["All"]+sorted(df.asset_class.dropna().unique().tolist()))
    inst=st.selectbox("Instrument",["All"]+sorted(df.instrument.dropna().unique().tolist()))
    syms=st.multiselect("Symbols",sorted(df.symbol.dropna().unique().tolist()))

f=df.copy()
if asset!="All": f=f[f.asset_class==asset]
if inst!="All": f=f[f.instrument==inst]
if syms: f=f[f.symbol.isin(syms)]

def money(x):
    return f"₹{float(x):,.2f}"

def pct(x):
    return f"{float(x):.2f}%"

def chart_layout(fig,height=300):
    fig.update_layout(height=height,margin=dict(l=10,r=10,t=50,b=10),
                      paper_bgcolor="rgba(0,0,0,0)",plot_bgcolor="rgba(0,0,0,0)",
                      font=dict(size=11),legend=dict(orientation="h",y=1.08,x=0))
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="rgba(148,163,184,.12)",zerolinecolor="rgba(148,163,184,.25)")
    return fig

def chart(fig,height=300):
    fig.update_layout(dragmode="pan")
    st.plotly_chart(
        chart_layout(fig,height),
        use_container_width=True,
        config={
            "staticPlot":True,
            "displayModeBar":False,
            "responsive":True,
            "scrollZoom":False,
            "doubleClick":False,
            "displaylogo":False
        }
    )

def section_data(name):
    x=f[f.asset_class.eq(name)].copy()
    charges=float(ch_year[ch_year.asset_class.eq(name)].amount.sum()) if not ch_year.empty else 0.0

    # Prefer broker report-level realised P&L for section/year totals.
    # Trade-row reconstruction can differ from Groww's report total.
    src=pd.read_sql_query(
        "select source_hash,period_start,period_end,report_gross_pnl from sources where asset_class=?",
        conn(), params=(name,)
    )
    if not src.empty:
        src["period_start"]=pd.to_datetime(src["period_start"],errors="coerce")
        src["period_end"]=pd.to_datetime(src["period_end"],errors="coerce")
        src=src[src["period_end"].dt.year.eq(current_year) & src["report_gross_pnl"].notna()].copy()

    gross=None
    if not src.empty:
        # A broader report replaces contained reports in save(). Use the
        # widest stored broker report when it covers the stored period.
        if len(src)==1:
            gross=float(src["report_gross_pnl"].iloc[0])
        else:
            src["span_days"]=(src["period_end"]-src["period_start"]).dt.days.fillna(-1)
            widest=src.sort_values(["span_days","period_end"],ascending=[False,False]).iloc[0]
            covered=src[
                (src["period_start"]>=widest["period_start"]) &
                (src["period_end"]<=widest["period_end"])
            ]
            if len(covered)==len(src):
                gross=float(widest["report_gross_pnl"])
            else:
                gross=float(src["report_gross_pnl"].sum())

    if gross is None:
        gross=float(x.pnl.sum())
    return x,charges,gross,gross-charges

def stocks_timing_view(x):
    st.subheader("⏰ Entry, exit & weekday analysis")
    buy_dt=pd.to_datetime(x.buy_date,errors="coerce")
    sell_dt=pd.to_datetime(x.sell_date,errors="coerce")
    has_buy_time=((buy_dt.dt.hour.fillna(0)!=0)|(buy_dt.dt.minute.fillna(0)!=0)|(buy_dt.dt.second.fillna(0)!=0))
    has_sell_time=((sell_dt.dt.hour.fillna(0)!=0)|(sell_dt.dt.minute.fillna(0)!=0)|(sell_dt.dt.second.fillna(0)!=0))
    has_time=bool(has_buy_time.any() or has_sell_time.any())

    t=x.assign(BuyDay=buy_dt.dt.day_name(),HoldHours=(sell_dt-buy_dt).dt.total_seconds()/3600)
    weekdays=["Monday","Tuesday","Wednesday","Thursday","Friday"]
    day=t.groupby("BuyDay",dropna=True).agg(PnL=("pnl","sum"),Trades=("pnl","size"),WinRate=("win","mean")).reset_index()
    day["WinRate"]=day.WinRate*100
    day["Order"]=pd.Categorical(day.BuyDay,categories=weekdays,ordered=True)
    day=day.sort_values("Order")

    if has_time:
        t=t.assign(BuyTime=buy_dt.dt.hour+buy_dt.dt.minute/60+buy_dt.dt.second/3600,
                   SellTime=sell_dt.dt.hour+sell_dt.dt.minute/60+sell_dt.dt.second/3600,
                   BuyHour=buy_dt.dt.hour,SellHour=sell_dt.dt.hour)
        t["Entry Slot"]=pd.cut(t.BuyTime,[-0.01,10,12,14,16,24],labels=["Before 10:00","10:00–12:00","12:00–14:00","14:00–16:00","After 16:00"])
        t["Exit Slot"]=pd.cut(t.SellTime,[-0.01,10,12,14,16,24],labels=["Before 10:00","10:00–12:00","12:00–14:00","14:00–16:00","After 16:00"])

        def slot_stats(col):
            z=t.dropna(subset=[col]).groupby(col,observed=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),WinRate=("win","mean")).reset_index()
            z["WinRate"]=z.WinRate*100
            return z[z.Trades>0]

        entry=slot_stats("Entry Slot")
        exit_=slot_stats("Exit Slot")
        a,b=st.columns(2)
        with a:
            fig=px.bar(entry,x="Entry Slot",y="PnL",color="PnL",color_continuous_scale="RdYlGn",title="Stocks — P&L by entry time")
            fig.update_yaxes(tickformat=",.2f"); chart(fig,310)
            if not entry.empty:
                good=entry.loc[entry.PnL.idxmax()]; bad=entry.loc[entry.PnL.idxmin()]
                if good.PnL>0: st.success(f"🔎 What went good: {good['Entry Slot']} produced {money(good.PnL)} across {int(good.Trades)} trades ({pct(good.WinRate)} win rate).")
                if bad.PnL<0: st.error(f"🔎 What went bad: {bad['Entry Slot']} lost {money(bad.PnL)} across {int(bad.Trades)} trades.")
        with b:
            fig=px.bar(exit_,x="Exit Slot",y="PnL",color="PnL",color_continuous_scale="RdYlGn",title="Stocks — P&L by exit time")
            fig.update_yaxes(tickformat=",.2f"); chart(fig,310)
            if not exit_.empty:
                good=exit_.loc[exit_.PnL.idxmax()]; bad=exit_.loc[exit_.PnL.idxmin()]
                if good.PnL>0: st.success(f"🔎 What went good: {good['Exit Slot']} exits produced {money(good.PnL)} across {int(good.Trades)} trades ({pct(good.WinRate)} win rate).")
                if bad.PnL<0: st.error(f"🔎 What went bad: {bad['Exit Slot']} exits lost {money(bad.PnL)} across {int(bad.Trades)} trades.")

        st.subheader("🕐 Entry → exit time combinations")
        hour=t.dropna(subset=["BuyHour","SellHour"]).groupby(["BuyHour","SellHour"]).agg(PnL=("pnl","sum"),Trades=("pnl","size"),WinRate=("win","mean")).reset_index()
        hour["WinRate"]=hour.WinRate*100
        hour["Buy time"]=hour.BuyHour.map(lambda v:f"{int(v):02d}:00")
        hour["Exit time"]=hour.SellHour.map(lambda v:f"{int(v):02d}:00")
        hour["Time Pair"]=hour["Buy time"]+" → "+hour["Exit time"]
        fig=px.bar(hour.sort_values("PnL"),x="PnL",y="Time Pair",orientation="h",color="PnL",color_continuous_scale="RdYlGn",title="Stocks — P&L by entry → exit time")
        fig.update_xaxes(tickformat=",.2f"); chart(fig,340)
        if not hour.empty:
            good=hour.loc[hour.PnL.idxmax()]; bad=hour.loc[hour.PnL.idxmin()]
            if good.PnL>0: st.success(f"🔎 What went good: {good['Buy time']} → {good['Exit time']} produced {money(good.PnL)} across {int(good.Trades)} trades.")
            if bad.PnL<0: st.error(f"🔎 What went bad: {bad['Buy time']} → {bad['Exit time']} lost {money(bad.PnL)} across {int(bad.Trades)} trades.")
        hold=t.dropna(subset=["HoldHours"])
        if not hold.empty and hold.HoldHours.notna().any():
            st.caption(f"Average holding time: {hold.HoldHours.mean():.2f} hours • Median: {hold.HoldHours.median():.2f} hours.")
    else:
        st.info("⏰ Entry/exit time analysis is waiting for EOD files that contain actual timestamps. The currently stored trade dates contain dates only, so no artificial time performance is shown.")

    fig=px.bar(day,x="BuyDay",y="PnL",color="PnL",color_continuous_scale="RdYlGn",title="Stocks — P&L by entry weekday")
    fig.update_yaxes(tickformat=",.2f"); chart(fig,310)

    weekday_trades=day[["BuyDay","Trades"]].copy()
    weekday_trades=weekday_trades[weekday_trades["Trades"]>0]
    if not weekday_trades.empty:
        fig=px.pie(
            weekday_trades,names="BuyDay",values="Trades",hole=0.45,
            title="Stocks — trade distribution by entry weekday"
        )
        chart(fig,300)

    if not day.empty:
        good=day.loc[day.PnL.idxmax()]; bad=day.loc[day.PnL.idxmin()]
        if good.PnL>0: st.success(f"🔎 What went good: {good.BuyDay} produced {money(good.PnL)} across {int(good.Trades)} trades ({pct(good.WinRate)} win rate).")
        if bad.PnL<0: st.error(f"🔎 What went bad: {bad.BuyDay} lost {money(bad.PnL)} across {int(bad.Trades)} trades ({pct(bad.WinRate)} win rate).")

def weekday_pnl_view(x, title):
    st.subheader("📅 Entry weekday analysis")
    buy_dt=pd.to_datetime(x.buy_date,errors="coerce")
    t=x.assign(BuyDay=buy_dt.dt.day_name())
    weekdays=["Monday","Tuesday","Wednesday","Thursday","Friday"]
    day=t.groupby("BuyDay",dropna=True).agg(
        PnL=("pnl","sum"),Trades=("pnl","size"),WinRate=("win","mean")
    ).reset_index()
    day["WinRate"]=day.WinRate*100
    day["Order"]=pd.Categorical(day.BuyDay,categories=weekdays,ordered=True)
    day=day.sort_values("Order")

    fig=px.bar(
        day,x="BuyDay",y="PnL",color="PnL",
        color_continuous_scale="RdYlGn",
        title=f"{title} — P&L by entry weekday"
    )
    fig.update_yaxes(tickformat=",.2f")
    chart(fig,310)

    weekday_trades=day[["BuyDay","Trades"]].copy()
    weekday_trades=weekday_trades[weekday_trades["Trades"]>0]
    if not weekday_trades.empty:
        fig=px.pie(
            weekday_trades,names="BuyDay",values="Trades",hole=0.45,
            title=f"{title} — trade distribution by entry weekday"
        )
        chart(fig,300)

    if not day.empty:
        good=day.loc[day.PnL.idxmax()]
        bad=day.loc[day.PnL.idxmin()]
        if good.PnL>0:
            st.success(
                f"🔎 What went good: {good.BuyDay} produced {money(good.PnL)} "
                f"across {int(good.Trades)} trades ({pct(good.WinRate)} win rate)."
            )
        if bad.PnL<0:
            st.error(
                f"🔎 What went bad: {bad.BuyDay} lost {money(bad.PnL)} "
                f"across {int(bad.Trades)} trades ({pct(bad.WinRate)} win rate)."
            )

def section_view(title,emoji,asset_name):
    x,charges,gross,net=section_data(asset_name)
    st.header(f"{emoji} {title}")
    st.caption(f"{current_year} only • Realised P&L uses broker-reported EOD realised P&L when available • Net realised P&L = Realised P&L − reported charges")

    if x.empty:
        st.info(f"No {title} trades loaded for {current_year}.")
        return

    wins=x.loc[x.pnl>0,"pnl"].sum()
    losses=abs(x.loc[x.pnl<0,"pnl"].sum())
    win_rate=(x.pnl>0).mean()*100
    pf=wins/losses if losses else np.inf

    a,b,c,d,e=st.columns(5)
    a.metric("Net realised P&L",money(net))
    b.metric("Realised P&L",money(gross))
    c.metric("Charges",money(charges))
    d.metric("Trades",f"{len(x):,}")
    e.metric("Win rate",pct(win_rate))

    if net>0:
        st.success(f"🟢 {title}: {current_year} is NET PROFITABLE after reported charges by {money(net)}.")
    elif net<0:
        st.error(f"🔴 {title}: {current_year} is NET LOSS after reported charges by {money(net)}.")
    else:
        st.info(f"🔵 {title}: {current_year} is approximately break-even after reported charges.")

    sym=x.groupby("symbol",as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"),Losses=("pnl",lambda s:(s<0).sum()))
    sym["WinRate"]=sym.Wins/sym.Trades*100
    sym["FailureRate"]=sym.Losses/sym.Trades*100

    daily=x.groupby("sell_date",as_index=False).agg(PnL=("pnl","sum"))
    daily["Cumulative"]=daily.PnL.cumsum()
    a,b=st.columns(2)
    with a:
        fig=px.line(daily,x="sell_date",y="Cumulative",markers=True,
                    title=f"{title} — cumulative realised P&L ({current_year})")
        fig.update_yaxes(tickformat=",.2f")
        chart(fig,320)
    with b:
        sv=sym.sort_values("PnL")
        show=pd.concat([sv.head(7),sv.tail(7)]).drop_duplicates()
        fig=px.bar(show,x="PnL",y="symbol",orientation="h",color="PnL",
                   color_continuous_scale="RdYlGn",
                   title=f"{title} — cumulative P&L by symbol")
        fig.update_xaxes(tickformat=",.2f")
        chart(fig,320)

    monthly=x.assign(MonthNum=x.sell_date.dt.month,Month=x.sell_date.dt.strftime("%b")) \
             .groupby(["MonthNum","Month"],as_index=False) \
             .agg(PnL=("pnl","sum"),Trades=("pnl","size")) \
             .sort_values("MonthNum")
    fig=px.bar(monthly,x="Month",y="PnL",color="PnL",
               color_continuous_scale="RdYlGn",
               title=f"{title} — monthly realised P&L ({current_year})")
    fig.update_yaxes(tickformat=",.2f")
    chart(fig,300)

    observation_text=f"Realised P&L {money(gross)} − reported charges {money(charges)} = net {money(net)}. "
    observation_text += f"Profit factor is {pf:.2f}." if np.isfinite(pf) else "There are no losing trades, so profit factor is undefined/infinite."
    st.info("🔎 Analysis — "+observation_text)

    # The 100% Success vs 100% Failure view is useful for the
    # Stocks section only. Keep F&O and Commodities focused on day-level
    # reconciliation and P&L/charges analytics.
    if asset_name=="Stocks":
        # Build the success/failure views directly from symbol aggregates.
        # The Symbol analysis table was intentionally removed.
        perfect=sym[sym["WinRate"]==100].copy()
        failed=sym[sym["FailureRate"]==100].copy()

        st.subheader("🎯 100% Success vs 100% Failure stocks")
        st.caption(
            "100% Success = every completed trade for that stock was profitable. "
            "100% Failure = every completed trade for that stock was a loss."
        )

        st.markdown("### 🟢 100% Success")
        if perfect.empty:
            st.info("No stock has a 100% success record in the current data.")
        else:
            p=perfect.sort_values(["Trades","PnL"],ascending=[False,False]).copy()
            p["Label"]=p["symbol"]+" ("+p["Trades"].astype(int).astype(str)+" trades)"
            fig=px.bar(
                p.sort_values("PnL"),
                x="PnL",
                y="Label",
                orientation="h",
                text="PnL",
                title="Stocks with 100% winning trades"
            )
            fig.update_traces(texttemplate="%{text:,.2f}",textposition="outside")
            fig.update_xaxes(tickformat=",.2f")
            chart(fig,max(300,min(650,230+len(p)*32)))
            st.success(
                "All trades were profitable for: "
                + ", ".join(f"{r.symbol} ({int(r.Trades)} trades)" for _,r in p.iterrows())
            )

        st.markdown("### 🔴 100% Failure")
        if failed.empty:
            st.info("No stock has a 100% failure record in the current data.")
        else:
            q=failed.sort_values(["Trades","PnL"],ascending=[False,True]).copy()
            q["Label"]=q["symbol"]+" ("+q["Trades"].astype(int).astype(str)+" trades)"
            fig=px.bar(
                q.sort_values("PnL"),
                x="PnL",
                y="Label",
                orientation="h",
                text="PnL",
                title="Stocks with 100% losing trades"
            )
            fig.update_traces(texttemplate="%{text:,.2f}",textposition="outside")
            fig.update_xaxes(tickformat=",.2f")
            chart(fig,max(300,min(650,230+len(q)*32)))
            st.error(
                "All trades were losing for: "
                + ", ".join(f"{r.symbol} ({int(r.Trades)} trades)" for _,r in q.iterrows())
            )

    if asset_name=="Stocks":
        stocks_timing_view(x)
    else:
        weekday_pnl_view(x, title)
    st.caption(
        f"Note: {asset_name} broker charges are reconciled at report level. "
        "A report-level charge total is not distributed across individual symbols "
        "or trading days unless the source itself contains day-specific charges."
    )

def overall_summary_view():
    """Current-year cumulative summary directly below the page title."""
    total_gross=0.0
    total_charges=0.0
    for section_name in ["Stocks","F&O","Commodities"]:
        _,section_charges,section_gross,_=section_data(section_name)
        total_gross += section_gross
        total_charges += section_charges
    total_net=total_gross-total_charges
    total_trades=len(df)
    total_win_rate=float((df.pnl>0).mean()*100) if not df.empty else 0.0

    st.markdown("### 📌 Cumulative details")
    a,b,c,d,e=st.columns(5)
    a.metric("Net realised P&L",money(total_net))
    b.metric("Realised P&L",money(total_gross))
    c.metric("Charges",money(total_charges))
    d.metric("Trades",f"{total_trades:,}")
    e.metric("Win rate",pct(total_win_rate))
    st.caption(
        f"{current_year} cumulative across Stocks + Equity F&O + Commodities • "
        "Realised P&L • broker-reported charges • Net P&L • cumulative performance"
    )

overall_summary_view()

tabs=st.tabs(["📈 Stocks","🎯 Equity F&O","⛽ Commodities"])
with tabs[0]:
    section_view("Stocks","📈","Stocks")
with tabs[1]:
    section_view("Equity F&O","🎯","F&O")
with tabs[2]:
    section_view("Commodities","⛽","Commodities")

st.markdown("---")
st.caption(f"📅 Dashboard scope: calendar year {current_year} only • three sections • realised P&L • reported charges • net P&L • cumulative profit/loss")