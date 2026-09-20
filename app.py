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

gross=f.pnl.sum(); wins=f.loc[f.win,"pnl"].sum(); losses=f.loc[f.loss,"pnl"].sum(); n=len(f)
charge=0
if not ch.empty:
    ch["period_end"]=pd.to_datetime(ch.period_end,errors="coerce")
    for a in f.asset_class.dropna().unique():
        z=ch[ch.asset_class==a]
        if len(z):
            last=z.period_end.max()
            charge += z[(z.period_end==last)&(z.charge_name.str.lower()=="total")].amount.sum()
net=gross-charge

# Compact KPI cards
cols=st.columns(6)
for c,label,val in zip(cols,["Gross P&L","Charges","Net P&L","Trades","Win rate","Profit factor"],
                       [money(gross),money(charge),money(net),f"{n:,}",pct(f.win.mean()) if n else "0%",f"{wins/abs(losses):.2f}" if losses<0 else "∞"]):
    c.metric(label,val)

tabs=st.tabs(["🏠 Overview","📅 Daily / Monthly","🏆 Success Analysis","⚠️ Review"])

with tabs[0]:
    daily=f.groupby("sell_date",as_index=False).pnl.sum().sort_values("sell_date")
    daily["cumulative"]=daily.pnl.cumsum(); daily["peak"]=daily.cumulative.cummax(); daily["drawdown"]=daily.cumulative-daily.peak
    x,y=st.columns(2)
    with x:
        fig=px.line(daily,x="sell_date",y="cumulative",markers=True,title="Cumulative P&L")
        fig.update_traces(line_width=3)
        st.plotly_chart(chart_layout(fig,290),use_container_width=True,config={"displayModeBar":False})
    with y:
        fig=px.bar(daily,x="sell_date",y="pnl",title="Daily P&L")
        fig.update_traces(marker_line_width=0)
        st.plotly_chart(chart_layout(fig,290),use_container_width=True,config={"displayModeBar":False})
    fig=px.area(daily,x="sell_date",y="drawdown",title="Drawdown")
    st.plotly_chart(chart_layout(fig,250),use_container_width=True,config={"displayModeBar":False})

    seg=f.groupby(["asset_class","instrument"]).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Win_Rate=("win","mean"),Avg_Trade=("pnl","mean")).reset_index().sort_values("PnL")
    fig=px.bar(seg,x="PnL",y="instrument",color="asset_class",orientation="h",title="P&L by segment",text_auto=".2s")
    st.plotly_chart(chart_layout(fig,300),use_container_width=True,config={"displayModeBar":False})

with tabs[1]:
    m=f.groupby("month").agg(Gross_PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum"),Losses=("loss","sum")).reset_index()
    m["Win_Rate"]=m.Wins/m.Trades; m["Cumulative_Gross"]=m.Gross_PnL.cumsum()
    st.subheader("Month-wise")
    st.dataframe(m.style.format({"Gross_PnL":"₹{:,.0f}","Cumulative_Gross":"₹{:,.0f}","Win_Rate":"{:.1%}"}),hide_index=True,use_container_width=True)
    x,y=st.columns(2)
    with x:
        fig=px.bar(m,x="month",y="Gross_PnL",title="Monthly P&L",color="Gross_PnL",color_continuous_scale="RdYlGn")
        st.plotly_chart(chart_layout(fig,270),use_container_width=True,config={"displayModeBar":False})
    with y:
        fig=px.line(m,x="month",y="Cumulative_Gross",markers=True,title="Cumulative monthly P&L")
        fig.update_traces(line_width=3)
        st.plotly_chart(chart_layout(fig,270),use_container_width=True,config={"displayModeBar":False})

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

    st.subheader("📅 Best and worst trading days")
    if len(day):
        st.dataframe(
            day.sort_values("PnL",ascending=False)[["Day","PnL","Trades","Win_Rate"]]
            .style.format({"PnL":"₹{:,.0f}","Win_Rate":"{:.1%}"}),
            hide_index=True,use_container_width=True
        )

    # Weekday consistency
    weekday_order=["Monday","Tuesday","Wednesday","Thursday","Friday"]
    wd=f.groupby("Weekday") if "Weekday" in f.columns else None
    fw=f.copy()
    fw["Weekday"]=fw.sell_date.dt.day_name()
    wd=fw.groupby("Weekday").agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum")).reindex(weekday_order).dropna(how="all").reset_index()
    wd["Win_Rate"]=wd.Wins/wd.Trades
    st.subheader("🗓️ Which weekday works best?")
    st.dataframe(wd.style.format({"PnL":"₹{:,.0f}","Win_Rate":"{:.1%}"}),hide_index=True,use_container_width=True)
    if len(wd):
        fig=px.bar(wd,x="Weekday",y="PnL",title="P&L by weekday",color="PnL",color_continuous_scale="RdYlGn")
        st.plotly_chart(chart_layout(fig,260),use_container_width=True,config={"displayModeBar":False})

    # Symbol success
    sym=f.groupby(["asset_class","symbol"]).agg(
        PnL=("pnl","sum"), Trades=("pnl","size"), Wins=("win","sum"), Avg_Trade=("pnl","mean")
    ).reset_index()
    sym["Win_Rate"]=sym.Wins/sym.Trades
    sym["Profitable"]=sym.PnL>0
    st.subheader("📈 Which stocks/contracts are working?")
    st.dataframe(
        sym.sort_values(["PnL","Win_Rate"],ascending=[False,False])
        .style.format({"PnL":"₹{:,.0f}","Avg_Trade":"₹{:,.0f}","Win_Rate":"{:.1%}"}),
        hide_index=True,use_container_width=True
    )

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
    st.dataframe(mm.style.format({"PnL":"₹{:,.0f}","Win_Rate":"{:.1%}"}),hide_index=True,use_container_width=True)

    # Instrument / option analysis
    st.subheader("🔎 Instrument success")
    q=f.groupby(["asset_class","instrument"]).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum")).reset_index()
    q["Win_Rate"]=q.Wins/q.Trades
    st.dataframe(q.style.format({"PnL":"₹{:,.0f}","Win_Rate":"{:.1%}"}),hide_index=True,use_container_width=True)
    op=f[f.option_type.notna()]
    if len(op):
        st.subheader("Options: CE vs PE")
        ce=op.groupby("option_type").agg(PnL=("pnl","sum"),Trades=("pnl","size"),Wins=("win","sum")).reset_index()
        ce["Win_Rate"]=ce.Wins/ce.Trades
        st.dataframe(ce.style.format({"PnL":"₹{:,.0f}","Win_Rate":"{:.1%}"}),hide_index=True,use_container_width=True)
        exp=op.groupby("expiry").agg(PnL=("pnl","sum"),Trades=("pnl","size")).reset_index().sort_values("PnL",ascending=False)
        st.subheader("Options by expiry")
        st.dataframe(exp.style.format({"PnL":"₹{:,.0f}"}),hide_index=True,use_container_width=True)

with tabs[3]:
    st.subheader("⚠️ What to review")
    warnings=[]
    byins=f.groupby(["asset_class","instrument"]).agg(PnL=("pnl","sum"),Trades=("pnl","size"),Win_Rate=("win","mean")).reset_index()
    for _,r in byins.iterrows():
        if r.PnL<0 and r.Trades>=5:
            warnings.append(("🔴 Loss concentration",f"{r.asset_class} {r.instrument}: {money(r.PnL)} over {int(r.Trades)} trades. Review this segment before increasing size."))
    if gross and charge>abs(gross)*0.10:
        warnings.append(("🟠 Cost drag",f"Charges {money(charge)} are large relative to gross P&L. Review turnover and low-edge trades."))
    if n:
        avg=f.pnl.mean(); medwin=f.loc[f.win,"pnl"].median() if f.win.any() else np.nan
        maxloss=abs(f.loc[f.loss,"pnl"].min()) if f.loss.any() else 0
        if avg<0:
            warnings.append(("🔴 Negative expectancy",f"Average trade is {money(avg)}. Review recurring losing patterns before increasing size."))
        if medwin and maxloss>3*medwin:
            warnings.append(("🟡 Loss-size imbalance",f"Largest loss {money(maxloss)} is >3× median winner {money(medwin)}. Review position sizing and exits."))
    op=f[f.instrument=="Options"]
    if len(op) and op.pnl.sum()<0:
        warnings.append(("🟣 Options review",f"Options are down {money(op.pnl.sum())}. Compare CE/PE, expiry and strike buckets."))
    if not warnings:
        st.success("✅ No major rule-based warnings for the selected filters.")
    for title,msg in warnings:
        st.warning(f"**{title}**  \\n{msg}")
    st.caption("These are statistical review prompts from your trade history, not predictions or financial advice.")
    st.subheader("Largest losses to review")
    st.dataframe(f.nsmallest(15,"pnl")[["sell_date","asset_class","instrument","symbol","qty","buy_price","sell_price","pnl","remark"]].style.format({"pnl":"₹{:,.0f}","buy_price":"₹{:,.2f}","sell_price":"₹{:,.2f}"}),hide_index=True,use_container_width=True)

