import streamlit as st
import pandas as pd
import numpy as np
import sqlite3, hashlib, io, re, os
from datetime import datetime

st.set_page_config(page_title="Trading Journal & Analytics", page_icon="📊", layout="wide")
DB = "trading_journal.db"

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
      gross_pnl REAL)""")
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
    s=str(s); opt="CE" if re.search(r"\bCall\b",s,re.I) else ("PE" if re.search(r"\bPut\b",s,re.I) else None)
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
    charges=[]; in_ch=False
    for _,row in raw0.iterrows():
        label=str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else ""; ll=label.lower()
        if ll=="charges": in_ch=True; continue
        if in_ch and label=="": break
        if in_ch:
            amt=pd.to_numeric(row.iloc[1],errors="coerce") if len(row)>1 else np.nan
            if not pd.isna(amt) and (any(k in ll for k in ["exchange transaction","sebi","stt","ctt","stamp duty","ipft","brokerage","gst","dp charges","mis charges"]) or ll=="total"):
                charges.append((label,float(amt)))
    return sh,asset,ps,pe,rows,charges

def save(uploaded,filename):
    sh,asset,ps,pe,rows,charges=extract(uploaded,filename); c=conn()
    if c.execute("select 1 from sources where source_hash=?",(sh,)).fetchone(): return 0,"Already uploaded",asset
    for r in rows:
        c.execute("insert or ignore into trades values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
          tuple(r[k] for k in ["trade_hash","source_hash","asset_class","symbol","instrument","option_type","strike","expiry","qty","buy_date","buy_price","buy_value","sell_date","sell_price","sell_value","pnl","remark","uploaded_at"]))
    for label,amt in charges:c.execute("insert or ignore into charges values(?,?,?,?,?,?)",(sh,asset,ps,pe,label,amt))
    c.execute("insert into sources values(?,?,?,?,?,?,?,?)",(sh,filename,asset,ps,pe,datetime.now().isoformat(timespec="seconds"),len(rows),sum(r["pnl"] for r in rows)))
    c.commit(); return len(rows),"Imported",asset

def money(x): return f"₹{x:,.0f}"
def pct(x): return f"{x:.1%}"

st.title("📊 Trading Journal & Analytics")
st.caption("Upload any EOD report whenever you want. Exact duplicate files/trades are ignored and the dashboard rebuilds automatically.")

with st.sidebar:
    st.header("1. Upload EOD")
    uploads=st.file_uploader("Stocks / F&O / Commodities Excel",type=["xlsx","xls"],accept_multiple_files=True)
    if uploads and st.button("Process & Update",type="primary"):
        for u in uploads:
            n,msg,a=save(u,u.name); st.success(f"{u.name}: {msg} — {n} trades")
        st.rerun()

df=pd.read_sql_query("select * from trades",conn()); ch=pd.read_sql_query("select * from charges",conn())
if df.empty:
    st.info("Upload your first EOD Excel from the sidebar."); st.stop()
df["sell_date"]=pd.to_datetime(df.sell_date,errors="coerce"); df["buy_date"]=pd.to_datetime(df.buy_date,errors="coerce")
df["month"]=df.sell_date.dt.to_period("M").astype(str); df["win"]=df.pnl>0; df["loss"]=df.pnl<0
with st.sidebar:
    st.header("2. Filters")
    asset=st.selectbox("Asset class",["All"]+sorted(df.asset_class.dropna().unique().tolist()))
    inst=st.selectbox("Instrument",["All"]+sorted(df.instrument.dropna().unique().tolist()))
    syms=st.multiselect("Symbols",sorted(df.symbol.dropna().unique().tolist()))
f=df.copy()
if asset!="All":f=f[f.asset_class==asset]
if inst!="All":f=f[f.instrument==inst]
if syms:f=f[f.symbol.isin(syms)]

gross=f.pnl.sum(); wins=f.loc[f.win,"pnl"].sum(); losses=f.loc[f.loss,"pnl"].sum(); n=len(f)
# Use latest broker charge snapshot for each selected asset, avoiding repeated cumulative-report charges.
charge=0
if not ch.empty:
    ch["period_end"]=pd.to_datetime(ch.period_end,errors="coerce")
    for a in f.asset_class.dropna().unique():
        z=ch[ch.asset_class==a]
        if len(z):
            last=z.period_end.max(); charge+=z[(z.period_end==last)&(z.charge_name.str.lower()=="total")].amount.sum()
net=gross-charge
a,b,c,d,e,fm=st.columns(6)
a.metric("Gross P&L",money(gross)); b.metric("Charges",money(charge)); c.metric("Net P&L",money(net))
d.metric("Trades",f"{n:,}"); e.metric("Win rate",pct(f.win.mean()) if n else "0%"); fm.metric("Profit factor",f"{wins/abs(losses):.2f}" if losses<0 else "∞")

tabs=st.tabs(["📈 Dashboard","📅 Daily & Monthly","🔎 Deep Analysis","⚠️ Mistake Finder","🧾 Trades"])
with tabs[0]:
    daily=f.groupby("sell_date",as_index=False).pnl.sum().sort_values("sell_date")
    daily["cumulative"]=daily.pnl.cumsum(); daily["peak"]=daily.cumulative.cummax(); daily["drawdown"]=daily.cumulative-daily.peak
    x,y=st.columns(2)
    with x: st.subheader("Cumulative P&L"); st.line_chart(daily.set_index("sell_date")["cumulative"])
    with y: st.subheader("Daily P&L"); st.bar_chart(daily.set_index("sell_date").pnl)
    st.subheader("Drawdown"); st.area_chart(daily.set_index("sell_date").drawdown)
    st.subheader("Segment performance")
    st.dataframe(f.groupby(["asset_class","instrument"]).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Win_Rate=("win","mean"),Avg_Trade=("pnl","mean")).reset_index().sort_values("PnL"),hide_index=True,use_container_width=True)

with tabs[1]:
    m=f.groupby("month").agg(Gross_PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"),Losses=("loss","sum")).reset_index()
    m["Win_Rate"]=m.Wins/m.Trades; m["Cumulative_Gross"]=m.Gross_PnL.cumsum()
    st.subheader("Month-wise performance"); st.dataframe(m,hide_index=True,use_container_width=True)
    x,y=st.columns(2)
    with x: st.bar_chart(m.set_index("month").Gross_PnL)
    with y: st.line_chart(m.set_index("month").Cumulative_Gross)
    st.subheader("Daily performance"); st.line_chart(f.groupby("sell_date").pnl.sum())

with tabs[2]:
    st.subheader("Instrument")
    st.dataframe(f.groupby(["asset_class","instrument"]).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Win_Rate=("win","mean"),Avg_Trade=("pnl","mean")).reset_index(),hide_index=True,use_container_width=True)
    op=f[f.option_type.notna()]
    if len(op):
        st.subheader("Options: CE vs PE"); st.dataframe(op.groupby(["asset_class","option_type"]).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Win_Rate=("win","mean"),Avg_Trade=("pnl","mean")).reset_index(),hide_index=True,use_container_width=True)
        st.subheader("Options by expiry"); st.dataframe(op.groupby(["expiry"]).agg(PnL=("pnl","sum"),Trades=("pnl","size")).reset_index().sort_values("PnL"),hide_index=True,use_container_width=True)
    sym=f.groupby(["asset_class","symbol"]).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum")).reset_index()
    x,y=st.columns(2); x.write("Top winners"); x.dataframe(sym.nlargest(10,"PnL"),hide_index=True,use_container_width=True)
    y.write("Biggest losers"); y.dataframe(sym.nsmallest(10,"PnL"),hide_index=True,use_container_width=True)

with tabs[3]:
    st.subheader("Automatic data-driven mistake finder")
    warnings=[]
    byins=f.groupby(["asset_class","instrument"]).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Win_Rate=("win","mean")).reset_index()
    for _,r in byins.iterrows():
        if r.PnL<0 and r.Trades>=5:warnings.append(("Loss concentration",f"{r.asset_class} {r.instrument}: {money(r.PnL)} over {int(r.Trades)} trades. Review this segment before increasing size."))
    if gross and charge>abs(gross)*0.10:warnings.append(("Cost drag",f"Charges of {money(charge)} are large relative to gross P&L. Review turnover and low-edge trades."))
    if n:
        avg=f.pnl.mean(); medwin=f.loc[f.win,"pnl"].median() if f.win.any() else np.nan; maxloss=abs(f.loc[f.loss,"pnl"].min()) if f.loss.any() else 0
        if avg<0:warnings.append(("Negative expectancy",f"Average realised trade is {money(avg)}. Identify the losing setup before increasing size."))
        if medwin and maxloss>3*medwin:warnings.append(("Loss sizing",f"Largest loss {money(maxloss)} is >3× median winning trade {money(medwin)}. Review stop-loss/position sizing."))
    op=f[f.instrument=="Options"]
    if len(op) and op.pnl.sum()<0:warnings.append(("Options review",f"Options are down {money(op.pnl.sum())}. Compare CE/PE, expiry and strike buckets."))
    if not warnings:st.success("No major rule-based warnings for the selected filters.")
    for title,msg in warnings:st.warning(f"**{title}:** {msg}")
    st.caption("Alerts are statistical patterns in the uploaded data, not a prediction or financial advice.")
    st.subheader("Largest losses to review")
    st.dataframe(f.nsmallest(15,"pnl")[["sell_date","asset_class","instrument","symbol","qty","buy_price","sell_price","pnl","remark"]],hide_index=True,use_container_width=True)

with tabs[4]:
    st.dataframe(f.sort_values("sell_date",ascending=False),hide_index=True,use_container_width=True)
    st.download_button("Download filtered CSV",f.to_csv(index=False).encode(),"trading_journal.csv","text/csv")
