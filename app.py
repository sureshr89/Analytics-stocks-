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

tabs=st.tabs(["🏠 Overview","📅 Daily / Monthly","🏆 Success Analysis","⚠️ Review"])

with tabs[0]:
    st.subheader("🎯 Performance Command Center")
    st.caption("Use this page to see where your money is being made, where it is being lost, and what deserves attention before the next trade.")

    # Clear success/failure split
    result_split=pd.DataFrame({"Result":["Profitable","Losing"],"PnL":[wins,abs(losses)]})
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Net after charges",money(net))
    c2.metric("Profitable trades",f"{int(f.win.sum()):,}")
    c3.metric("Losing trades",f"{int(f.loss.sum()):,}")
    c4.metric("Best trade",money(f.pnl.max()) if n else "—")

    a,b=st.columns(2)
    with a:
        fig=px.pie(result_split,names="Result",values="PnL",hole=.55,title="Profit vs loss contribution")
        st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})
    with b:
        asset_view=f.groupby("asset_class",as_index=False).agg(PnL=("pnl","sum"))
        fig=px.bar(asset_view,x="asset_class",y="PnL",color="PnL",color_continuous_scale="RdYlGn",title="Which asset is helping or hurting?")
        st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})

    daily=f.groupby("sell_date",as_index=False).pnl.sum().sort_values("sell_date")
    daily["cumulative"]=daily.pnl.cumsum()
    daily["peak"]=daily.cumulative.cummax()
    daily["drawdown"]=daily.cumulative-daily.peak
    daily["spike"]=daily.pnl.abs()>daily.pnl.abs().mean()+2*daily.pnl.abs().std() if len(daily)>2 else False

    st.subheader("📈 Your trading journey")
    fig=px.line(daily,x="sell_date",y="cumulative",markers=True,title="Cumulative P&L — growth and turning points")
    fig.update_traces(line_width=3)
    st.plotly_chart(chart_layout(fig,310),use_container_width=True,config={"displayModeBar":False})

    a,b=st.columns(2)
    with a:
        fig=px.bar(daily,x="sell_date",y="pnl",color="pnl",color_continuous_scale="RdYlGn",title="Every trading day — green = profit, red = loss")
        st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})
    with b:
        fig=px.area(daily,x="sell_date",y="drawdown",title="Drawdown — where the account gave back gains")
        st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})

    st.subheader("🔥 Trading spikes")
    spikes=daily[daily.spike].copy()
    if len(spikes):
        fig=px.bar(spikes,x="sell_date",y="pnl",color="pnl",color_continuous_scale="RdYlGn",title="Unusually large profit/loss days")
        st.plotly_chart(chart_layout(fig,270),use_container_width=True,config={"displayModeBar":False})
        st.info("Spike days are statistical outliers in daily P&L. Open those dates in your broker record and review size, setup and market conditions.")
    else:
        st.success("No unusually large daily P&L spikes detected in the selected data.")

    seg=f.groupby(["asset_class","instrument"],as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Win_Rate=("win","mean"))
    fig=px.bar(seg,x="PnL",y="instrument",color="asset_class",orientation="h",title="Success / failure by instrument",text_auto=".2s")
    st.plotly_chart(chart_layout(fig,310),use_container_width=True,config={"displayModeBar":False})

    st.subheader("🏆 Your strongest areas")
    sym=f.groupby(["asset_class","symbol"],as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"))
    sym["Win_Rate"]=sym.Wins/sym.Trades
    a,b=st.columns(2)
    with a:
        top=sym.nlargest(8,"PnL")
        fig=px.bar(top,x="PnL",y="symbol",orientation="h",color="asset_class",title="Symbols contributing most profit")
        st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})
    with b:
        bottom=sym.nsmallest(8,"PnL")
        fig=px.bar(bottom,x="PnL",y="symbol",orientation="h",color="asset_class",title="Symbols causing most loss")
        st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})

with tabs[1]:
    st.subheader("📅 Daily & Monthly Analysis")
    st.caption("This page is for finding repeatable good days, bad days, strong months, weak months and unusual changes — not for reading a raw trade list.")

    m=f.groupby("month").agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"),Losses=("loss","sum")).reset_index()
    m["Win_Rate"]=m.Wins/m.Trades
    m["Cumulative"]=m.PnL.cumsum()

    a,b=st.columns(2)
    with a:
        fig=px.bar(m,x="month",y="PnL",color="PnL",color_continuous_scale="RdYlGn",title="Month-by-month P&L")
        st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})
    with b:
        fig=px.line(m,x="month",y="Cumulative",markers=True,title="Cumulative monthly result")
        fig.update_traces(line_width=3)
        st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})

    day=f.groupby("sell_date").agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"),Losses=("loss","sum")).reset_index()
    day["Win_Rate"]=day.Wins/day.Trades
    day["Day_Type"]=np.where(day.PnL>0,"Profit day","Loss day")
    avg_abs=day.PnL.abs().mean()
    sd_abs=day.PnL.abs().std()
    day["Spike"]=day.PnL.abs()>avg_abs+2*sd_abs if len(day)>2 else False

    st.subheader("🟢🔴 Good days vs bad days")
    a,b=st.columns(2)
    with a:
        fig=px.bar(day,x="sell_date",y="PnL",color="Day_Type",title="Daily profit / loss")
        st.plotly_chart(chart_layout(fig,320),use_container_width=True,config={"displayModeBar":False})
    with b:
        split=day.groupby("Day_Type",as_index=False).size().rename(columns={"size":"Days"})
        fig=px.pie(split,names="Day_Type",values="Days",hole=.5,title="Share of profitable vs losing days")
        st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})

    best=day.nlargest(7,"PnL")
    worst=day.nsmallest(7,"PnL")
    a,b=st.columns(2)
    with a:
        fig=px.bar(best.sort_values("PnL"),x="PnL",y="sell_date",orientation="h",color="PnL",color_continuous_scale="RdYlGn",title="Best trading days")
        st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})
    with b:
        fig=px.bar(worst.sort_values("PnL"),x="PnL",y="sell_date",orientation="h",color="PnL",color_continuous_scale="RdYlGn",title="Worst trading days")
        st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})

    st.subheader("📊 Day behaviour")
    weekday_order=["Monday","Tuesday","Wednesday","Thursday","Friday"]
    fw=f.copy(); fw["Weekday"]=fw.sell_date.dt.day_name()
    wd=fw.groupby("Weekday",as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"))
    wd["Win_Rate"]=wd.Wins/wd.Trades
    wd["Weekday"]=pd.Categorical(wd.Weekday,categories=weekday_order,ordered=True)
    wd=wd.sort_values("Weekday")
    fig=px.bar(wd,x="Weekday",y="PnL",color="PnL",color_continuous_scale="RdYlGn",title="P&L by weekday")
    st.plotly_chart(chart_layout(fig,270),use_container_width=True,config={"displayModeBar":False})

    st.subheader("🗓️ Good / bad months at a glance")
    month_split=m.assign(Type=np.where(m.PnL>0,"Profitable month","Loss month")).groupby("Type",as_index=False).size().rename(columns={"size":"Months"})
    fig=px.pie(month_split,names="Type",values="Months",hole=.5,title="Profitable vs loss months")
    st.plotly_chart(chart_layout(fig,280),use_container_width=True,config={"displayModeBar":False})

    if day.Spike.any():
        st.subheader("🚨 P&L spikes to investigate")
        sp=day[day.Spike]
        fig=px.bar(sp,x="sell_date",y="PnL",color="PnL",color_continuous_scale="RdYlGn",title="Unusual daily moves")
        st.plotly_chart(chart_layout(fig,280),use_container_width=True,config={"displayModeBar":False})

    st.subheader("🎯 Before your next trade")
    st.info("Use the strongest symbols/instruments and profitable-day patterns as evidence to review your process — and use the worst days, largest losses and spikes as a checklist of what to avoid repeating. This is a data review, not a prediction of the next trade.")

with tabs[2]:
    st.subheader("🏆 Where are you most successful?")
    st.caption("Success is measured from your realised trade history: profitable days, profitable symbols, win rate and average P&L.")

    # Best trading days
    day = f.groupby("sell_date").agg(
        PnL=("pnl","sum"), Trades=("pnl","size"), Wins=("win","sum")
    ).reset_index()
    day["Win_Rate"] = day.Wins / day.Trades
    day["Day"] = day.sell_date.dt.strftime("%a, %d %b")
    day["Weekday"] = day.sell_date.dt.day_name()
    day["Profitable_Day"] = day.PnL > 0

    best_day = day.loc[day.PnL.idxmax()] if len(day) else None
    worst_day = day.loc[day.PnL.idxmin()] if len(day) else None
    profitable_days = int(day.Profitable_Day.sum())
    losing_days = int((day.PnL < 0).sum())

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Profitable days", f"{profitable_days}")
    c2.metric("Losing days", f"{losing_days}")
    c3.metric("Best day", money(best_day.PnL) if best_day is not None else "—")
    c4.metric("Worst day", money(worst_day.PnL) if worst_day is not None else "—")

    st.subheader("📅 Trading-day P&L pattern")
    if len(day):
        fig=px.bar(day.sort_values("sell_date"),x="Day",y="PnL",color="PnL",color_continuous_scale="RdYlGn",title="Every trading day")
        st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})

    # Weekday consistency
    weekday_order=["Monday","Tuesday","Wednesday","Thursday","Friday"]
    wd=f.groupby("Weekday") if "Weekday" in f.columns else None
    fw=f.copy()
    fw["Weekday"]=fw.sell_date.dt.day_name()
    wd=fw.groupby("Weekday").agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum")).reindex(weekday_order).dropna(how="all").reset_index()
    wd["Win_Rate"]=wd.Wins/wd.Trades
    st.subheader("🗓️ Which weekday works best?")
    if len(wd):
        fig=px.bar(wd,x="Weekday",y="PnL",title="P&L by weekday",color="PnL",color_continuous_scale="RdYlGn")
        st.plotly_chart(chart_layout(fig,260),use_container_width=True,config={"displayModeBar":False})

    # Symbol success
    sym=f.groupby(["asset_class","symbol"]).agg(
        PnL=("pnl","sum"), Trades=("pnl","size"), Wins=("win","sum"), Avg_Trade=("pnl","mean")
    ).reset_index()
    sym["Win_Rate"]=sym.Wins/sym.Trades
    sym["Profitable"]=sym.PnL>0
    st.subheader("📈 Symbol success map")
    x,y=st.columns(2)
    with x:
        top=sym.nlargest(10,"PnL")
        fig=px.bar(top,x="PnL",y="symbol",orientation="h",color="asset_class",title="Top 10 by P&L")
        st.plotly_chart(chart_layout(fig,310),use_container_width=True,config={"displayModeBar":False})
    with y:
        topw=sym.sort_values("Win_Rate",ascending=False).head(10)
        fig=px.bar(topw,x="Win_Rate",y="symbol",orientation="h",color="asset_class",title="Highest win rate")
        fig.update_xaxes(tickformat=".0%")
        st.plotly_chart(chart_layout(fig,310),use_container_width=True,config={"displayModeBar":False})

    # Month consistency
    mm=f.groupby("month").agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum")).reset_index()
    mm["Win_Rate"]=mm.Wins/mm.Trades
    mm["Profitable"]=mm.PnL>0
    st.subheader("📆 Monthly consistency")
    fig=px.bar(mm,x="month",y="PnL",color="PnL",color_continuous_scale="RdYlGn",title="Monthly consistency")
    st.plotly_chart(chart_layout(fig,270),use_container_width=True,config={"displayModeBar":False})

    # Instrument / option analysis
    st.subheader("🔎 Instrument P&L")
    q=f.groupby(["asset_class","instrument"]).agg(PnL=("pnl","sum"),Trades=("pnl","size")).reset_index()
    fig=px.bar(q,x="instrument",y="PnL",color="asset_class",barmode="group",title="Stocks vs Futures vs Options")
    st.plotly_chart(chart_layout(fig,280),use_container_width=True,config={"displayModeBar":False})
    op=f[f.option_type.notna()]
    if len(op):
        x,y=st.columns(2)
        with x:
            ce=op.groupby("option_type",as_index=False).agg(PnL=("pnl","sum"))
            fig=px.bar(ce,x="option_type",y="PnL",color="option_type",title="CE vs PE")
            st.plotly_chart(chart_layout(fig,260),use_container_width=True,config={"displayModeBar":False})
        with y:
            exp=op.groupby("expiry",as_index=False).agg(PnL=("pnl","sum")).sort_values("PnL")
            fig=px.bar(exp,x="expiry",y="PnL",color="PnL",color_continuous_scale="RdYlGn",title="P&L by expiry")
            st.plotly_chart(chart_layout(fig,260),use_container_width=True,config={"displayModeBar":False})
        op["is_expiry_day"]=op.apply(lambda r: pd.notna(r.sell_date) and pd.notna(r.expiry) and r.sell_date.strftime("%Y-%m-%d")==str(r.expiry),axis=1)
        exd=op.groupby("is_expiry_day",as_index=False).agg(PnL=("pnl","sum"))
        exd["Day_Type"]=exd.is_expiry_day.map({True:"Expiry day",False:"Non-expiry day"})
        fig=px.bar(exd,x="Day_Type",y="PnL",color="Day_Type",title="Expiry day vs non-expiry day")
        st.plotly_chart(chart_layout(fig,270),use_container_width=True,config={"displayModeBar":False})

with tabs[3]:
    st.subheader("🧭 Success & Failure Review")
    st.caption("Clear evidence from your history: what repeatedly works, what repeatedly hurts, and which behaviours deserve a pre-trade check.")

    byins=f.groupby(["asset_class","instrument"],as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"))
    byins["Win_Rate"]=byins.Wins/byins.Trades

    a,b=st.columns(2)
    with a:
        good=byins.nlargest(8,"PnL")
        fig=px.bar(good,x="PnL",y="instrument",orientation="h",color="asset_class",title="Where results are strongest")
        st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})
    with b:
        bad=byins.nsmallest(8,"PnL")
        fig=px.bar(bad,x="PnL",y="instrument",orientation="h",color="asset_class",title="Where results are weakest")
        st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})

    sym=f.groupby(["asset_class","symbol"],as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"),Avg_Trade=("pnl","mean"))
    sym["Win_Rate"]=sym.Wins/sym.Trades
    a,b=st.columns(2)
    with a:
        fig=px.bar(sym.nlargest(10,"PnL"),x="PnL",y="symbol",orientation="h",color="asset_class",title="Symbols with strongest results")
        st.plotly_chart(chart_layout(fig,320),use_container_width=True,config={"displayModeBar":False})
    with b:
        fig=px.bar(sym.nsmallest(10,"PnL"),x="PnL",y="symbol",orientation="h",color="asset_class",title="Symbols with weakest results")
        st.plotly_chart(chart_layout(fig,320),use_container_width=True,config={"displayModeBar":False})

    op=f[f.option_type.notna()].copy()
    if len(op):
        st.subheader("⚙️ F&O / Options behaviour")
        a,b=st.columns(2)
        with a:
            cepe=op.groupby("option_type",as_index=False).agg(PnL=("pnl","sum"))
            fig=px.pie(cepe,names="option_type",values="PnL",hole=.5,title="CE / PE contribution")
            st.plotly_chart(chart_layout(fig,280),use_container_width=True,config={"displayModeBar":False})
        with b:
            if op.expiry.notna().any():
                exp=op.groupby("expiry",as_index=False).agg(PnL=("pnl","sum")).sort_values("PnL")
                fig=px.bar(exp,x="expiry",y="PnL",color="PnL",color_continuous_scale="RdYlGn",title="Options P&L by expiry")
                st.plotly_chart(chart_layout(fig,280),use_container_width=True,config={"displayModeBar":False})

    st.subheader("🔴 Largest failures to learn from")
    ll=f.nsmallest(10,"pnl").copy()
    ll["label"]=ll.symbol.astype(str)+" • "+ll.sell_date.dt.strftime("%d %b")
    fig=px.bar(ll.sort_values("pnl"),x="pnl",y="label",orientation="h",color="asset_class",title="Largest losing trades")
    st.plotly_chart(chart_layout(fig,330),use_container_width=True,config={"displayModeBar":False})

    # Data-driven pre-trade checklist
    checklist=[]
    if len(sym):
        bad_sym=sym.nsmallest(1,"PnL").iloc[0]
        good_sym=sym.nlargest(1,"PnL").iloc[0]
        checklist.append(f"Review why {good_sym.symbol} produced {money(good_sym.PnL)} before assuming another setup will behave the same way.")
        checklist.append(f"Put extra scrutiny on {bad_sym.symbol}: historical P&L is {money(bad_sym.PnL)} in the selected data.")
    if len(byins):
        bad_ins=byins.nsmallest(1,"PnL").iloc[0]
        checklist.append(f"Check the setup and sizing rules for {bad_ins.asset_class} {bad_ins.instrument}; its selected-history P&L is {money(bad_ins.PnL)}.")
    if f.loss.any():
        maxloss=f.loc[f.loss].nsmallest(1,"pnl").iloc[0]
        checklist.append(f"Before entry, define the maximum acceptable loss. Your largest recorded loss was {money(maxloss.pnl)} on {maxloss.symbol}.")
    if charge>0:
        checklist.append(f"Review transaction cost before high-turnover trades: selected-period charges are {money(charge)}.")
    for item in checklist:
        st.warning("• "+item)
    st.caption("These prompts describe historical patterns. They do not predict the next trade or guarantee that a past pattern will repeat.")

