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
st.caption("Upload EOD files anytime • automatic history • duplicate protection • gross/net reconciliation • data-driven review")

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

df["sell_date"]=pd.to_datetime(df.sell_date,errors="coerce").astype("datetime64[ns]"); df["buy_date"]=pd.to_datetime(df.buy_date,errors="coerce").astype("datetime64[ns]")
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
src["period_end"]=pd.to_datetime(src["period_end"],errors="coerce").astype("datetime64[ns]")
# Reconstructed trade rows are the source of truth for gross P&L.
# Report-level gross totals can double-count overlapping/corrected EOD files.
gross=float(f.pnl.sum())
charge=0.0
    if not ch.empty:
        ch["period_end"]=pd.to_datetime(ch["period_end"],errors="coerce").astype("datetime64[ns]")
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



def safe_money(v):
    return money(float(v)) if pd.notna(v) else "—"

def observation(title, text, kind="info"):
    icon={"good":"🟢","bad":"🔴","warn":"🟠","info":"🔵"}.get(kind,"🔵")
    st.markdown(f"**{icon} {title}** — {text}")

def chart(fig, height=300):
    return st.plotly_chart(chart_layout(fig,height), use_container_width=True, config={"displayModeBar":False,"responsive":True})

def period_frame(base, mode):
    x=base.copy()
    if x.empty: return x
    latest=x["sell_date"].max()
    if pd.isna(latest): return x
    if mode=="Latest trading day":
        d=x[x.sell_date==latest]
        return d
    today=pd.Timestamp(datetime.now().date())
    if mode=="Current month":
        return x[(x.sell_date.dt.year==today.year)&(x.sell_date.dt.month==today.month)]
    if mode=="Past 3 months":
        return x[x.sell_date>=today-pd.DateOffset(months=3)]
    if mode=="Current year":
        return x[x.sell_date.dt.year==today.year]
    return x

def period_label(mode, x):
    if x.empty: return mode
    return f"{mode} • {x.sell_date.min().strftime('%d %b %Y')} → {x.sell_date.max().strftime('%d %b %Y')}"

def add_underlying(x):
    x=x.copy()
    def u(v):
        s=str(v).upper()
        for k in ["BANKNIFTY","FINNIFTY","MIDCPNIFTY","NIFTY","SENSEX","BANKEX"]:
            if k in s: return k
        return str(v).split()[0] if str(v).strip() else "Unknown"
    x["Underlying"]=x.symbol.map(u)
    return x

tabs=st.tabs(["📊 Overview","🗓️ Periods & Spikes","🎯 NIFTY / F&O / Commodities","🚦 Next-Trade Panel"])

# Global decision horizon
st.markdown("### ⏱️ Decision horizon")
horizon=st.radio("Use the same horizon across the dashboard",["Latest trading day","Current month","Past 3 months","Current year","All history"],horizontal=True,label_visibility="collapsed")
hf=period_frame(f,horizon)
if hf.empty:
    st.info("No trades are available in this period. Select another horizon or upload the missing EOD file.")
st.caption(period_label(horizon,hf) if not hf.empty else f"{horizon} • no trades loaded")

with tabs[0]:
    st.subheader("📊 Performance Map")
    day=f.groupby("sell_date",as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"))
    day["Win_Rate"]=day.Wins/day.Trades
    day["Cumulative"]=day.PnL.cumsum()
    day["Peak"]=day.Cumulative.cummax()
    day["Drawdown"]=day.Cumulative-day.Peak

    # KPI row stays compact; charts carry the analysis.
    gross_h=float(hf.pnl.sum()); trades_h=len(hf); wins_h=int(hf.win.sum())
    avg_h=float(hf.pnl.mean()) if trades_h else 0
    pf=(hf.loc[hf.pnl>0,"pnl"].sum()/abs(hf.loc[hf.pnl<0,"pnl"].sum())) if (hf.pnl<0).any() else np.inf
    c1,c2,c3,c4,c5=st.columns(5)
    # Trade P&L is gross realised P&L; charges are tracked separately.\n    horizon_charge=0.0\n    if not ch.empty and not hf.empty:\n        ch_tmp=ch.copy(); ch_tmp["period_end"]=pd.to_datetime(ch_tmp["period_end"],errors="coerce").astype("datetime64[ns]")\n        horizon_charge=float(ch_tmp[(ch_tmp.asset_class.isin(hf.asset_class.unique())) & (ch_tmp.period_end>=hf.sell_date.min()) & (ch_tmp.period_end<=hf.sell_date.max()) & (ch_tmp.charge_name.str.lower()=="total")].amount.sum())\n    horizon_net=gross_h-horizon_charge\n    c1.metric("Net P&L (estimated)",money(horizon_net))
    c2.metric("Trades",f"{trades_h:,}")
    c3.metric("Win rate",pct(wins_h/trades_h) if trades_h else "0%")
    c4.metric("Avg trade",money(avg_h))
    c5.metric("Profit factor",f"{pf:.2f}" if np.isfinite(pf) else "∞")

    a,b=st.columns(2)
    with a:
        res=pd.DataFrame({"Result":["Wins","Losses"],"Amount":[hf.loc[hf.pnl>0,"pnl"].sum(),abs(hf.loc[hf.pnl<0,"pnl"].sum())]})
        fig=px.pie(res,names="Result",values="Amount",hole=.60,title="Win vs loss contribution")
        chart(fig,290)
        observation("What went right",f"Winners contributed {money(res.Amount.iloc[0])} versus {money(res.Amount.iloc[1])} of losing P&L in this horizon." if res.Amount.sum() else "No realised P&L in this horizon.","good" if res.Amount.iloc[0]>=res.Amount.iloc[1] else "warn")
    with b:
        av=hf.groupby("asset_class",as_index=False).agg(PnL=("pnl","sum")).sort_values("PnL")
        fig=px.bar(av,x="PnL",y="asset_class",orientation="h",color="PnL",color_continuous_scale="RdYlGn",title="Asset contribution")
        chart(fig,290)
        top=av.iloc[-1] if len(av) else None
        observation("What went right / wrong",f"{top.asset_class} is the largest positive contributor at {money(top.PnL)}." if top is not None and top.PnL>0 else "No positive asset contribution in this horizon.","good" if top is not None and top.PnL>0 else "warn")

    # Full cumulative curve, plus recent horizon overlay.
    fig=px.line(day,x="sell_date",y="Cumulative",markers=True,title="Overall cumulative P&L")
    fig.update_traces(line_width=3)
    chart(fig,330)
    if len(day):
        last=day.iloc[-1]
        observation("Equity read",f"Latest day closed at {money(last.PnL)}; cumulative P&L is {money(last.Cumulative)}. The distance below the prior peak is {money(last.Drawdown)}.","good" if last.Drawdown==0 else "warn")

    a,b=st.columns(2)
    with a:
        recent=day.tail(min(20,len(day))).copy()
        recent["Type"]=np.where(recent.PnL>=0,"Profit","Loss")
        fig=px.bar(recent,x="sell_date",y="PnL",color="Type",title="Last 20 trading days")
        chart(fig,300)
        if len(recent):
            avg_recent=recent.PnL.mean()
            observation("Recent rhythm",f"Average daily result is {money(avg_recent)} across {len(recent)} trading days.","good" if avg_recent>0 else "bad")
    with b:
        fig=px.area(day,x="sell_date",y="Drawdown",title="Drawdown / capital given back")
        chart(fig,300)
        dd=day.Drawdown.min() if len(day) else 0
        observation("Risk signal",f"Maximum recorded drawdown in loaded history is {money(dd)}.","warn" if dd<0 else "good")

    # Concentration and cost visual
    a,b=st.columns(2)
    with a:
        sym=hf.groupby("symbol",as_index=False).agg(PnL=("pnl","sum")).sort_values("PnL")
        show=pd.concat([sym.head(5),sym.tail(5)]).drop_duplicates().sort_values("PnL") if len(sym)>10 else sym
        fig=px.bar(show,x="PnL",y="symbol",orientation="h",color="PnL",color_continuous_scale="RdYlGn",title="Top / bottom symbol contribution")
        chart(fig,330)
        if len(sym):
            observation("Concentration",f"Best symbol contributed {money(sym.PnL.max())}; weakest contributed {money(sym.PnL.min())}.", "good" if sym.PnL.max()>abs(sym.PnL.min()) else "warn")
    with b:
        charge_by=hf.groupby("asset_class",as_index=False).agg(Gross=("pnl","sum"))
        if not ch.empty:
            cc=ch[(ch.period_end>=hf.sell_date.min()) & (ch.period_end<=hf.sell_date.max()) & (ch.charge_name.str.lower()=="total")].groupby("asset_class",as_index=False).agg(Charges=("amount","sum"))
            charge_by=charge_by.merge(cc,on="asset_class",how="left").fillna(0)
        else: charge_by["Charges"]=0
        charge_by["Charge_%"]=np.where(charge_by.Gross.abs()>0,charge_by.Charges/charge_by.Gross.abs()*100,0)
        fig=px.bar(charge_by,x="asset_class",y=["Gross","Charges"],barmode="group",title="Gross P&L vs reported charges")
        chart(fig,330)
        observation("Cost check",f"Total tracked charges in loaded data are {money(charge_by.Charges.sum())}. Watch high-turnover periods where charges consume a large share of gross P&L.","warn" if charge_by.Charges.sum()>max(0,charge_by.Gross.sum())*.15 else "info")

with tabs[1]:
    st.subheader("🗓️ Timeframe + Spike Engine")
    day=f.groupby("sell_date",as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"))
    day["Win_Rate"]=day.Wins/day.Trades
    day["Cumulative"]=day.PnL.cumsum()
    day["AbsPnL"]=day.PnL.abs()
    if len(day)>=4:
        roll=day.AbsPnL.rolling(10,min_periods=3).median()
        day["SpikeScore"]=day.AbsPnL/roll.replace(0,np.nan)
        day["Spike"]=day.SpikeScore>=2
    else:
        day["SpikeScore"]=1.0; day["Spike"]=False

    monthly=f.groupby("month",as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"))
    monthly["Win_Rate"]=monthly.Wins/monthly.Trades
    monthly["Cumulative"]=monthly.PnL.cumsum()
    monthly["Type"]=np.where(monthly.PnL>=0,"Profit","Loss")

    # Requested period views
    period_rows=[]
    latest=f.sell_date.max()
    today=pd.Timestamp(datetime.now().date())
    if pd.notna(latest):
        prev=f[f.sell_date==latest]
        cm=f[(f.sell_date.dt.year==today.year)&(f.sell_date.dt.month==today.month)]
        y=f[f.sell_date.dt.year==today.year]
        period_rows=[
            {"Period":"Latest trading day","PnL":prev.pnl.sum(),"Trades":len(prev),"WinRate":prev.win.mean()*100 if len(prev) else 0},
            {"Period":"Current month","PnL":cm.pnl.sum(),"Trades":len(cm),"WinRate":cm.win.mean()*100 if len(cm) else 0},
            {"Period":"Current year","PnL":y.pnl.sum(),"Trades":len(y),"WinRate":y.win.mean()*100 if len(y) else 0},
            {"Period":"Overall","PnL":f.pnl.sum(),"Trades":len(f),"WinRate":f.win.mean()*100 if len(f) else 0}
        ]
    pv=pd.DataFrame(period_rows)
    if len(pv):
        a,b=st.columns(2)
        with a:
            fig=px.bar(pv,x="Period",y="PnL",color="PnL",color_continuous_scale="RdYlGn",title="Latest trading day → current month → current year → overall")
            chart(fig,300)
            observation("Period comparison",f"Current month: {money(pv.loc[pv.Period=='Current month','PnL'].iloc[0])}; current year: {money(pv.loc[pv.Period=='Current year','PnL'].iloc[0])}. Use the comparison to see whether recent results are aligned with the longer record.","info")
        with b:
            fig=px.bar(pv,x="Period",y="WinRate",color="WinRate",color_continuous_scale="RdYlGn",title="Win rate by period")
            fig.update_yaxes(ticksuffix="%")
            chart(fig,300)
            observation("Consistency",f"Overall win rate is {pv.loc[pv.Period=='Overall','WinRate'].iloc[0]:.1f}%; current month is {pv.loc[pv.Period=='Current month','WinRate'].iloc[0]:.1f}%.","good" if pv.loc[pv.Period=='Current month','WinRate'].iloc[0]>=pv.loc[pv.Period=='Overall','WinRate'].iloc[0] else "warn")

    a,b=st.columns(2)
    with a:
        fig=px.bar(monthly,x="month",y="PnL",color="Type",title="Monthly P&L")
        chart(fig,300)
        observation("Monthly read",f"{int((monthly.PnL>0).sum())} profitable months and {int((monthly.PnL<0).sum())} loss months are visible in the loaded data.","good" if monthly.PnL.sum()>0 else "warn")
    with b:
        fig=px.line(monthly,x="month",y="Cumulative",markers=True,title="Monthly cumulative path")
        chart(fig,300)
        observation("Trend",f"Ending monthly cumulative P&L is {money(monthly.Cumulative.iloc[-1]) if len(monthly) else '—'}.","good" if len(monthly) and monthly.Cumulative.iloc[-1]>0 else "warn")

    spikes=day[day.Spike].copy()
    fig=px.bar(day,x="sell_date",y="PnL",color="Spike",title="Daily P&L + abnormal spikes")
    chart(fig,330)
    if len(spikes):
        worst=spikes.loc[spikes.PnL.abs().idxmax()]
        observation("Spike alert",f"{len(spikes)} unusual daily move(s). Largest spike was {money(worst.PnL)} on {worst.sell_date.strftime('%d %b %Y')}. Review size, number of trades, instrument and exit discipline on those days.","warn")
    else:
        observation("Spike alert","No abnormal daily P&L spike detected versus the rolling baseline.","good")

    # Weekday and trade-size concentration
    fw=f.copy(); fw["Weekday"]=fw.sell_date.dt.day_name()
    order=["Monday","Tuesday","Wednesday","Thursday","Friday"]
    wd=fw.groupby("Weekday",as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"))
    wd["Weekday"]=pd.Categorical(wd.Weekday,categories=order,ordered=True); wd=wd.sort_values("Weekday")
    a,b=st.columns(2)
    with a:
        fig=px.bar(wd,x="Weekday",y="PnL",color="PnL",color_continuous_scale="RdYlGn",title="P&L by weekday")
        chart(fig,290)
        if len(wd):
            z=wd.loc[wd.PnL.abs().idxmax()]
            observation("Weekday pattern",f"Largest absolute weekday contribution is {z.Weekday}: {money(z.PnL)}.","info")
    with b:
        if len(f)>1:
            vals=f.pnl.abs()
            q=vals.quantile(.9)
            size_spikes=f[vals>=q]
            fig=px.bar(size_spikes.sort_values("pnl"),x="pnl",y="symbol",orientation="h",color="pnl",color_continuous_scale="RdYlGn",title="Largest trade P&L spikes")
            chart(fig,290)
            observation("Trade spike",f"Top 10% of trade P&L magnitude contains {len(size_spikes)} trade(s). Check whether these are repeatable setups or oversized exceptions.","warn")
        else:
            st.info("More trades are needed for trade-spike analysis.")

with tabs[2]:
    st.subheader("🎯 NIFTY • F&O • Commodities")
    st.caption("The dashboard separates the buckets so the next setup can be checked against the same market/instrument family.")
    fo=add_underlying(f)

    # Asset mix
    av=fo.groupby("asset_class",as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"))
    av["WinRate"]=av.Wins/av.Trades*100
    a,b=st.columns(2)
    with a:
        fig=px.bar(av,x="asset_class",y="PnL",color="PnL",color_continuous_scale="RdYlGn",title="Stocks vs F&O vs Commodities")
        chart(fig,300)
        observation("Asset read","Compare contribution and win rate before assuming the next trade belongs in the same bucket.","info")
    with b:
        fig=px.pie(av,names="asset_class",values="Trades",hole=.58,title="Trade mix by asset")
        chart(fig,300)
        if len(av):
            observation("Concentration",f"{av.Trades.max()/av.Trades.sum():.0%} of trades sit in {av.loc[av.Trades.idxmax(),'asset_class']}.","warn" if av.Trades.max()/av.Trades.sum()>.7 else "info")

    # NIFTY family and F&O instrument breakdown
    fo_mask=fo.asset_class.eq("F&O") | fo.Underlying.isin(["NIFTY","BANKNIFTY","FINNIFTY","MIDCPNIFTY","SENSEX","BANKEX"])
    nf=fo[fo_mask].copy()
    if len(nf):
        uv=nf.groupby("Underlying",as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"))
        uv["WinRate"]=uv.Wins/uv.Trades*100
        a,b=st.columns(2)
        with a:
            fig=px.bar(uv.sort_values("PnL"),x="PnL",y="Underlying",orientation="h",color="PnL",color_continuous_scale="RdYlGn",title="Index / F&O underlying contribution")
            chart(fig,320)
            best=uv.loc[uv.PnL.idxmax()]
            observation("Underlying read",f"Highest realised contribution: {best.Underlying} at {money(best.PnL)}. Check the setup and sizing behind that result rather than copying the symbol alone.","good" if best.PnL>0 else "warn")
        with b:
            fig=px.bar(uv,x="Underlying",y="WinRate",color="WinRate",color_continuous_scale="RdYlGn",title="Win rate by underlying")
            fig.update_yaxes(ticksuffix="%")
            chart(fig,320)
            observation("Consistency",f"Best observed underlying win rate: {uv.WinRate.max():.1f}%.","info")
    else:
        st.info("No NIFTY/index F&O rows detected in the loaded data yet. Upload an F&O EOD file and the dashboard will populate this section.")

    op=fo[(fo.option_type.notna())].copy()
    if len(op):
        a,b=st.columns(2)
        with a:
            cepe=op.groupby("option_type",as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"))
            fig=px.bar(cepe,x="option_type",y="PnL",color="PnL",color_continuous_scale="RdYlGn",title="CE vs PE realised P&L")
            chart(fig,290)
            z=cepe.loc[cepe.PnL.idxmax()]
            observation("Option-side read",f"{z.option_type} has the larger realised contribution at {money(z.PnL)}.","info")
        with b:
            ex=op[op.expiry.notna()].groupby("expiry",as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"))
            fig=px.bar(ex,x="expiry",y="PnL",color="PnL",color_continuous_scale="RdYlGn",title="P&L by expiry")
            chart(fig,290)
            observation("Expiry read",f"{len(ex)} expiry bucket(s) are represented. Check whether losses cluster around a specific expiry.","info")

        if "strike" in op.columns and op.strike.notna().any():
            strike=op.groupby(["Underlying","option_type"],as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"))
            fig=px.bar(strike,x="Underlying",y="PnL",color="option_type",barmode="group",title="Underlying × CE/PE contribution")
            chart(fig,300)

    com=fo[fo.asset_class.eq("Commodities")]
    if len(com):
        cv=com.groupby("symbol",as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"))
        cv["WinRate"]=cv.Wins/cv.Trades*100
        fig=px.bar(cv.sort_values("PnL"),x="PnL",y="symbol",orientation="h",color="PnL",color_continuous_scale="RdYlGn",title="Commodity contribution")
        chart(fig,310)
        z=cv.loc[cv.PnL.idxmax()]
        observation("Commodity read",f"Highest commodity contribution is {z.symbol} at {money(z.PnL)}.","info")
    else:
        st.info("No commodity trades loaded yet.")

with tabs[3]:
    st.subheader("🚦 Next-Trade Panel")
    st.caption("This is a decision checklist built from your own realised history. It does not predict price direction or issue a Buy/Sell signal.")

    base=f.copy()
    if len(base):
        # Evidence blocks: strongest/weakest buckets and recent regime.
        by_inst=base.groupby(["asset_class","instrument"],as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"))
        by_inst["WinRate"]=by_inst.Wins/by_inst.Trades*100
        by_sym=base.groupby("symbol",as_index=False).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"))
        by_sym["WinRate"]=by_sym.Wins/by_sym.Trades*100

        c1,c2,c3=st.columns(3)
        strongest=by_inst.loc[by_inst.PnL.idxmax()] if len(by_inst) else None
        weakest=by_inst.loc[by_inst.PnL.idxmin()] if len(by_inst) else None
        recent20=base.sort_values("sell_date").tail(20)
        c1.metric("Recent 20 P&L",money(recent20.pnl.sum()))
        c2.metric("Recent win rate",pct(recent20.win.mean()) if len(recent20) else "0%")
        c3.metric("Max single loss",money(base.pnl.min()))

        a,b=st.columns(2)
        with a:
            instplot=by_inst.sort_values("PnL")
            fig=px.bar(instplot,x="PnL",y="instrument",color="PnL",color_continuous_scale="RdYlGn",title="Historical P&L by instrument")
            chart(fig,310)
            if strongest is not None:
                observation("Before entering",f"Compare the proposed trade's instrument with your historical record: {strongest.asset_class} {strongest.instrument} has the largest contribution at {money(strongest.PnL)}.","good" if strongest.PnL>0 else "warn")
        with b:
            if len(by_sym):
                top=by_sym.sort_values("PnL").tail(8)
                fig=px.bar(top,x="PnL",y="symbol",orientation="h",color="PnL",color_continuous_scale="RdYlGn",title="Highest cumulative P&L symbols")
                chart(fig,310)
                observation("Do not chase",f"Historical symbol performance is evidence, not a reason by itself to enter. Verify today's setup independently.","warn")
        
        # Risk budget visual from actual losses
        losses=base.loc[base.pnl<0,"pnl"].abs()
        wins=base.loc[base.pnl>0,"pnl"]
        risk=pd.DataFrame({"Bucket":["Median loss","90th pct loss","Median win","90th pct win"],
                           "Amount":[losses.median() if len(losses) else 0,losses.quantile(.9) if len(losses) else 0,wins.median() if len(wins) else 0,wins.quantile(.9) if len(wins) else 0]})
        fig=px.bar(risk,x="Bucket",y="Amount",color="Amount",color_continuous_scale="RdYlGn",title="Historical trade outcome size — risk reference")
        chart(fig,300)
        observation("Risk sizing prompt",f"Historical median loss is {money(losses.median()) if len(losses) else '—'} and 90th-percentile loss is {money(losses.quantile(.9)) if len(losses) else '—'}. Set your maximum loss before entry.","warn")

    st.markdown("### ✅ 30-second entry gate")
    checks=[
        "Setup matches a defined playbook / pattern I have tested.",
        "Instrument bucket is intentional (Stocks / NIFTY F&O / other F&O / Commodities).",
        "Entry, invalidation and exit are defined before order placement.",
        "Position size fits my pre-defined maximum loss.",
        "Expected payoff is large enough to justify charges/slippage.",
        "I am not entering to recover a recent loss or because of FOMO."
    ]
    for i,q in enumerate(checks,1):
        st.checkbox(q,key=f"decision_gate_{i}")
    done=sum(st.session_state.get(f"decision_gate_{i}",False) for i in range(1,len(checks)+1))
    st.progress(done/len(checks),text=f"Entry gate completed: {done}/{len(checks)}")
    if done==len(checks):
        st.success("All checklist items completed. The dashboard is ready for you to make your own trade decision.")
    else:
        st.info("Finish the checks that apply before entering. The dashboard deliberately does not turn historical performance into a Buy/Sell prediction.")

st.markdown("---")
st.caption("📱 Mobile-first • charts before tables • every chart has an observation • historical evidence only • Stocks + NIFTY/F&O + Commodities • latest trading day / current month / current year / overall views")
