import streamlit as st
import pandas as pd
import numpy as np
import sqlite3, hashlib, io, re
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Trading Journal", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

DB = "trading_journal.db"

# ---------- Mobile-first visual design ----------
st.markdown("""
<style>
:root { --bg:#0b1020; --card:#151c2f; --muted:#94a3b8; --text:#f8fafc; }
.block-container { max-width: 1500px; padding: 1rem 1.1rem 3rem; }
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
  .block-container { padding:.55rem .55rem 2rem; }
  [data-testid="stMetric"] { min-height:82px; padding:.6rem .65rem; }
  [data-testid="stMetricValue"] { font-size:1.05rem !important; }
  .stTabs [data-baseweb="tab"] { font-size:.7rem; padding:.45rem .5rem; }
  [data-testid="stHorizontalBlock"] { gap:.45rem; }
  .js-plotly-plot { max-height:320px; }
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
    c.execute("""CREATE TABLE IF NOT EXISTS charges(
      source_hash TEXT, asset_class TEXT, period_start TEXT, period_end TEXT,
      charge_name TEXT, amount REAL, PRIMARY KEY(source_hash,charge_name))""")
    c.commit(); return c

def parse_date(v):
    x=pd.to_datetime(v,errors="coerce",dayfirst=True)
    return None if pd.isna(x) else x.strftime("%Y-%m-%d")

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
    ps,pe=period_from(raw0)
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
    return sh,asset,ps,pe,rows,charges,report_gross,report_charges

def save(uploaded,filename):
    sh,asset,ps,pe,rows,charges,report_gross,report_charges=extract(uploaded,filename); c=conn()
    existing=c.execute('select source_hash,period_start,period_end from sources where asset_class=?',(asset,)).fetchall()
    replace_hashes=[]; covered=False
    if ps and pe:
        nps=pd.to_datetime(ps); npe=pd.to_datetime(pe)
        for old_hash,ops,ope in existing:
            if not ops or not ope: continue
            ods=pd.to_datetime(ops); ode=pd.to_datetime(ope)
            if ods==nps and ode==npe: replace_hashes.append(old_hash)
            elif ods>=nps and ode<=npe: replace_hashes.append(old_hash)
            elif ods<=nps and ode>=npe: covered=True
    if covered and not replace_hashes: return 0,'Already covered by existing report',asset
    if replace_hashes:
        q=','.join('?'*len(replace_hashes))
        c.execute(f'delete from trades where source_hash in ({q})',replace_hashes)
        c.execute(f'delete from charges where source_hash in ({q})',replace_hashes)
        c.execute(f'delete from sources where source_hash in ({q})',replace_hashes)
    if ps and pe: c.execute('delete from trades where asset_class=? and sell_date>=? and sell_date<=?',(asset,ps,pe))
    for r in rows: c.execute('insert or ignore into trades values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',tuple(r[k] for k in ['trade_hash','source_hash','asset_class','symbol','instrument','option_type','strike','expiry','qty','buy_date','buy_price','buy_value','sell_date','sell_price','sell_value','pnl','remark','uploaded_at']))
    for label,amt in charges: c.execute('insert or ignore into charges values(?,?,?,?,?,?)',(sh,asset,ps,pe,label,amt))
    c.execute('insert into sources(source_hash,filename,asset_class,period_start,period_end,uploaded_at,trade_count,gross_pnl,report_gross_pnl,report_charges) values(?,?,?,?,?,?,?,?,?,?)',(sh,filename,asset,ps,pe,datetime.now().isoformat(timespec='seconds'),len(rows),sum(r['pnl'] for r in rows),report_gross,report_charges))
    c.commit(); return len(rows),'Reconciled & imported',asset

def money(x): return f"₹{x:,.0f}"
def pct(x): return f"{x:.1%}"

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
st.caption("Upload EOD files anytime • automatic history • duplicate protection • data-driven review")

with st.sidebar:
    st.header("📤 Upload EOD")
    uploads=st.file_uploader("Stocks / F&O / Commodities",type=["xlsx","xls"],accept_multiple_files=True)
    if uploads and st.button("⚡ Process & Update",type="primary",use_container_width=True):
        for u in uploads:
            n,msg,a=save(u,u.name); st.success(f"{a}: {msg} • {n} trades")
        st.rerun()

df=pd.read_sql_query("select * from trades",conn()); ch=pd.read_sql_query("select * from charges",conn())
if df.empty:
    st.info("📤 Upload your first EOD Excel from the left sidebar.")
    st.stop()

df["sell_date"]=pd.to_datetime(df.sell_date,errors="coerce"); df["buy_date"]=pd.to_datetime(df.buy_date,errors="coerce")
df["month"]=df.sell_date.dt.to_period("M").astype(str); df["win"]=df.pnl>0; df["loss"]=df.pnl<0

with st.sidebar:
    st.header("🔎 Filters")
    asset=st.selectbox("Asset",["All"]+sorted(df.asset_class.dropna().unique().tolist()))
    inst=st.selectbox("Instrument",["All"]+sorted(df.instrument.dropna().unique().tolist()))
    syms=st.multiselect("Symbols",sorted(df.symbol.dropna().unique().tolist()))
f=df.copy()
if asset!="All": f=f[f.asset_class==asset]
if inst!="All": f=f[f.instrument==inst]
if syms: f=f[f.symbol.isin(syms)]

src=pd.read_sql_query("select * from sources",conn())
src["period_end"]=pd.to_datetime(src["period_end"],errors="coerce")
if not src.empty and src["report_gross_pnl"].notna().any() and inst=="All" and not syms:
    scoped=src[src.asset_class.isin(f.asset_class.dropna().unique())]
    gross=float(scoped["report_gross_pnl"].fillna(0).sum())
    charge=float(scoped["report_charges"].fillna(0).sum())
else:
    gross=float(f.pnl.sum())
    charge=0.0
    if not ch.empty:
        ch["period_end"]=pd.to_datetime(ch.period_end,errors="coerce")
        for aa in f.asset_class.dropna().unique():
            z=ch[ch.asset_class==aa]
            if len(z):
                last=z.period_end.max()
                charge += z[(z.period_end==last)&(z.charge_name.str.lower()=="total")].amount.sum()
wins=f.loc[f.win,"pnl"].sum(); losses=f.loc[f.loss,"pnl"].sum(); n=len(f)
net=gross-charge

# Compact KPI cards
cols=st.columns(5)
for c,label,val in zip(cols,["Gross P&L","Charges","Net P&L","Trades","Win rate"],
                       [money(gross),money(charge),money(net),f"{n:,}",pct(f.win.mean()) if n else "0%"]):
    c.metric(label,val)


tabs=st.tabs(["🏠 Decision Dashboard","📅 Daily Signals","🎯 F&O Edge","🧠 Pre-Trade Review"])

def safe_money(v):
    return money(float(v)) if pd.notna(v) else "—"

def insight_card(title, body, kind="info"):
    msg = "**" + str(title) + "**  \n" + str(body)
    if kind=="good": st.success(msg)
    elif kind=="bad": st.error(msg)
    elif kind=="warn": st.warning(msg)
    else: st.info(msg)

with tabs[0]:
    st.subheader("🎯 Decision Dashboard")
    st.caption("Not a trade diary. This page turns your realised history into clear evidence for your next trade: what has worked, what has hurt, and where risk needs extra checking.")

    day=f.groupby("sell_date",as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"))
    day["Win_Rate"]=day.Wins/day.Trades
    day["Type"]=np.where(day.PnL>0,"Profit day","Loss day")
    day["Cumulative"]=day.PnL.cumsum()
    day["Peak"]=day.Cumulative.cummax()
    day["Drawdown"]=day.Cumulative-day.Peak

    sym=f.groupby(["asset_class","symbol"],as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"),Avg=("pnl","mean"))
    sym["Win_Rate"]=sym.Wins/sym.Trades
    ins=f.groupby(["asset_class","instrument"],as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"))
    ins["Win_Rate"]=ins.Wins/ins.Trades

    good_days=int((day.PnL>0).sum()); bad_days=int((day.PnL<0).sum())
    best_day=day.PnL.max() if len(day) else 0; worst_day=day.PnL.min() if len(day) else 0
    avg_win=f.loc[f.win,"pnl"].mean() if f.win.any() else 0
    avg_loss=f.loc[f.loss,"pnl"].mean() if f.loss.any() else 0

    k1,k2,k3,k4,k5=st.columns(5)
    k1.metric("Net P&L",money(net))
    k2.metric("Profit days",str(good_days))
    k3.metric("Loss days",str(bad_days))
    k4.metric("Best day",money(best_day))
    k5.metric("Worst day",money(worst_day))

    # One-glance composition and outcome
    a,b=st.columns(2)
    with a:
        split=pd.DataFrame({"Result":["Profitable","Losing"],"Amount":[max(wins,0),abs(min(losses,0))]})
        fig=px.pie(split,names="Result",values="Amount",hole=.58,title="Where trade P&L came from")
        st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})
    with b:
        asset_view=f.groupby("asset_class",as_index=False).agg(PnL=("pnl","sum"))
        fig=px.bar(asset_view,x="asset_class",y="PnL",color="PnL",color_continuous_scale="RdYlGn",title="Asset-level contribution")
        st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})

    # The most useful visual: recent-to-history equity path
    fig=px.line(day,x="sell_date",y="Cumulative",markers=True,title="Cumulative P&L — turning points")
    fig.update_traces(line_width=3)
    st.plotly_chart(chart_layout(fig,315),use_container_width=True,config={"displayModeBar":False})

    a,b=st.columns(2)
    with a:
        fig=px.bar(day,x="sell_date",y="PnL",color="Type",title="Daily outcome — green profit / red loss")
        st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})
    with b:
        fig=px.area(day,x="sell_date",y="Drawdown",title="Drawdown — capital given back")
        st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})

    # Clear success / failure evidence
    st.subheader("🟢 What is working")
    good=sym[sym.PnL>0].sort_values("PnL",ascending=False).head(7)
    bad=sym[sym.PnL<0].sort_values("PnL").head(7)
    a,b=st.columns(2)
    with a:
        if len(good):
            fig=px.bar(good.sort_values("PnL"),x="PnL",y="symbol",orientation="h",color="asset_class",title="Symbols with positive realised P&L")
            st.plotly_chart(chart_layout(fig,290),use_container_width=True,config={"displayModeBar":False})
        else: st.info("No profitable symbols in the current filter.")
    with b:
        if len(bad):
            fig=px.bar(bad.sort_values("PnL"),x="PnL",y="symbol",orientation="h",color="asset_class",title="Symbols with negative realised P&L")
            st.plotly_chart(chart_layout(fig,290),use_container_width=True,config={"displayModeBar":False})
        else: st.success("No losing symbols in the current filter.")

    # Pre-trade evidence, not predictions
    if len(sym):
        best=sym.loc[sym.PnL.idxmax()]
        worst=sym.loc[sym.PnL.idxmin()]
        if best.PnL>0:
            insight_card("Evidence to carry forward",f"{best.symbol} is the strongest symbol in this selected history at {money(best.PnL)}. Review the setups, sizing and conditions behind those trades before repeating the process.","good")
        if worst.PnL<0:
            insight_card("Risk check",f"{worst.symbol} is the weakest symbol at {money(worst.PnL)}. Treat a new trade there as a review trigger: verify setup quality, size and exit plan before entry.","bad")
    if avg_loss<0 and avg_win>0:
        insight_card("Loss-size check",f"Average winner: {money(avg_win)} • average loser: {money(avg_loss)}. Before entry, define the maximum acceptable loss so one trade does not dominate several winners.","warn")

with tabs[1]:
    st.subheader("📅 Daily Signals")
    st.caption("Find repeatable good/bad days and spikes. The goal is to improve the next decision, not to scroll through every historical trade.")

    day=f.groupby("sell_date",as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"))
    day["Win_Rate"]=day.Wins/day.Trades
    day["Type"]=np.where(day.PnL>0,"Profit day","Loss day")
    day["abs_pnl"]=day.PnL.abs()
    threshold=day.abs_pnl.mean()+2*day.abs_pnl.std() if len(day)>2 else np.inf
    day["Spike"]=day.abs_pnl>threshold

    monthly=f.groupby("month",as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"))
    monthly["Win_Rate"]=monthly.Wins/monthly.Trades
    monthly["Type"]=np.where(monthly.PnL>0,"Profitable month","Loss month")
    monthly["Cumulative"]=monthly.PnL.cumsum()

    a,b=st.columns(2)
    with a:
        fig=px.bar(monthly,x="month",y="PnL",color="Type",title="Monthly result")
        st.plotly_chart(chart_layout(fig,290),use_container_width=True,config={"displayModeBar":False})
    with b:
        fig=px.line(monthly,x="month",y="Cumulative",markers=True,title="Cumulative monthly path")
        fig.update_traces(line_width=3)
        st.plotly_chart(chart_layout(fig,290),use_container_width=True,config={"displayModeBar":False})

    st.subheader("🟢🔴 Good days vs bad days")
    a,b=st.columns(2)
    with a:
        fig=px.bar(day,x="sell_date",y="PnL",color="Type",title="Every trading day")
        st.plotly_chart(chart_layout(fig,310),use_container_width=True,config={"displayModeBar":False})
    with b:
        split=day.groupby("Type",as_index=False).size().rename(columns={"size":"Days"})
        fig=px.pie(split,names="Type",values="Days",hole=.55,title="Share of profit vs loss days")
        st.plotly_chart(chart_layout(fig,290),use_container_width=True,config={"displayModeBar":False})

    a,b=st.columns(2)
    with a:
        best=day.nlargest(7,"PnL").sort_values("PnL")
        fig=px.bar(best,x="PnL",y="sell_date",orientation="h",color="PnL",color_continuous_scale="RdYlGn",title="Best days")
        st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})
    with b:
        worst=day.nsmallest(7,"PnL").sort_values("PnL")
        fig=px.bar(worst,x="PnL",y="sell_date",orientation="h",color="PnL",color_continuous_scale="RdYlGn",title="Worst days")
        st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})

    st.subheader("⚡ Spikes worth investigating")
    spikes=day[day.Spike]
    if len(spikes):
        fig=px.bar(spikes,x="sell_date",y="PnL",color="PnL",color_continuous_scale="RdYlGn",title="Unusually large daily P&L moves")
        st.plotly_chart(chart_layout(fig,280),use_container_width=True,config={"displayModeBar":False})
        st.warning("Spike = daily P&L unusually large versus your own selected history. Review what changed that day: size, instrument, setup, number of trades and exit behaviour.")
    else:
        st.success("No unusual daily P&L spike detected.")

    fw=f.copy(); fw["Weekday"]=fw.sell_date.dt.day_name()
    order=["Monday","Tuesday","Wednesday","Thursday","Friday"]
    wd=fw.groupby("Weekday",as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"))
    wd["Win_Rate"]=wd.Wins/wd.Trades
    wd["Weekday"]=pd.Categorical(wd.Weekday,categories=order,ordered=True)
    wd=wd.sort_values("Weekday")
    fig=px.bar(wd,x="Weekday",y="PnL",color="PnL",color_continuous_scale="RdYlGn",title="P&L by weekday")
    st.plotly_chart(chart_layout(fig,275),use_container_width=True,config={"displayModeBar":False})

    st.subheader("🗓️ Month quality")
    splitm=monthly.groupby("Type",as_index=False).size().rename(columns={"size":"Months"})
    fig=px.pie(splitm,names="Type",values="Months",hole=.55,title="Profitable vs loss months")
    st.plotly_chart(chart_layout(fig,270),use_container_width=True,config={"displayModeBar":False})

with tabs[2]:
    st.subheader("🎯 F&O Edge")
    st.caption("Use this to compare the parts of your F&O activity that have actually worked: futures/options, CE/PE and expiry vs non-expiry days.")

    ins=f.groupby(["asset_class","instrument"],as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"))
    ins["Win_Rate"]=ins.Wins/ins.Trades
    fig=px.bar(ins,x="instrument",y="PnL",color="asset_class",barmode="group",title="Futures vs Options vs Stocks")
    st.plotly_chart(chart_layout(fig,285),use_container_width=True,config={"displayModeBar":False})

    op=f[f.option_type.notna()].copy()
    if len(op):
        a,b=st.columns(2)
        with a:
            cepe=op.groupby("option_type",as_index=False).agg(PnL=("pnl","sum"))
            fig=px.pie(cepe,names="option_type",values="PnL",hole=.55,title="CE vs PE P&L contribution")
            st.plotly_chart(chart_layout(fig,290),use_container_width=True,config={"displayModeBar":False})
        with b:
            exp=op[op.expiry.notna()].groupby("expiry",as_index=False).agg(PnL=("pnl","sum")).sort_values("PnL")
            if len(exp):
                fig=px.bar(exp,x="expiry",y="PnL",color="PnL",color_continuous_scale="RdYlGn",title="P&L by option expiry")
                st.plotly_chart(chart_layout(fig,290),use_container_width=True,config={"displayModeBar":False})

        op["is_expiry_day"]=op.apply(lambda r: pd.notna(r.sell_date) and pd.notna(r.expiry) and r.sell_date.strftime("%Y-%m-%d")==str(r.expiry),axis=1)
        ex=op.groupby("is_expiry_day",as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"))
        ex["Day_Type"]=ex.is_expiry_day.map({True:"Expiry day",False:"Non-expiry day"})
        fig=px.bar(ex,x="Day_Type",y="PnL",color="Day_Type",title="Expiry day vs non-expiry day")
        st.plotly_chart(chart_layout(fig,285),use_container_width=True,config={"displayModeBar":False})
    else:
        st.info("No option trades in the current filter.")

with tabs[3]:
    st.subheader("🧠 Pre-Trade Review")
    st.caption("A short evidence-based checklist generated from your own history. It is deliberately focused on what to check before entering the next trade.")

    sym=f.groupby(["asset_class","symbol"],as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"),Avg=("pnl","mean"))
    sym["Win_Rate"]=sym.Wins/sym.Trades
    ins=f.groupby(["asset_class","instrument"],as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"))
    ins["Win_Rate"]=ins.Wins/ins.Trades

    if len(sym):
        best=sym.nlargest(1,"PnL").iloc[0]
        worst=sym.nsmallest(1,"PnL").iloc[0]
        if best.PnL>0:
            insight_card("1. Reuse the process, not the symbol blindly",f"Your strongest selected symbol is {best.symbol}: {money(best.PnL)} across {int(best.Trades)} trades. Before another trade, identify what setup/conditions were common in its profitable trades.","good")
        if worst.PnL<0:
            insight_card("2. Slow down on repeated weak areas",f"{worst.symbol} is at {money(worst.PnL)} across {int(worst.Trades)} trades. Before another entry, require a clear setup and predefined exit instead of relying on the historical symbol name.","bad")

    if len(ins):
        weak=ins.nsmallest(1,"PnL").iloc[0]
        strong=ins.nlargest(1,"PnL").iloc[0]
        if strong.PnL>0:
            insight_card("3. Know your stronger instrument bucket",f"{strong.asset_class} {strong.instrument} contributes {money(strong.PnL)} in the selected history. Check whether the next setup matches the conditions behind those results.","good")
        if weak.PnL<0:
            insight_card("4. Risk-check your weaker instrument bucket",f"{weak.asset_class} {weak.instrument} contributes {money(weak.PnL)}. Before entry, check size, stop/exit plan and whether the setup is one you have historically handled well.","warn")

    if f.loss.any():
        largest=f.loc[f.loss].nsmallest(1,"pnl").iloc[0]
        insight_card("5. Protect against a repeat of the largest loss",f"Largest recorded loss: {money(largest.pnl)} in {largest.symbol}. Define the maximum loss and position size before sending the order.","warn")

    if charge>0:
        insight_card("6. Check cost before high-turnover trades",f"Selected-period charges are {money(charge)}. If the setup has small expected movement, compare the planned payoff with your historical cost burden before entering.","warn")

    # Compact decision checklist
    st.subheader("✅ Before you press Buy / Sell")
    checks=[
        "Is this setup similar to a setup that has actually made money in my history?",
        "Is position size consistent with my historical risk, not with how confident I feel today?",
        "Where is the invalidation / maximum acceptable loss?",
        "Is the expected move large enough to justify charges and slippage?",
        "Am I entering because of a defined setup, or because I am trying to recover a recent loss?"
    ]
    for i,c in enumerate(checks,1):
        st.checkbox(c,key=f"pretrade_{i}")
    st.caption("The dashboard uses historical evidence to frame questions. It does not predict the next trade or tell you what position to take.")

