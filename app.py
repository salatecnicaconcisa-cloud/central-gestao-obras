import streamlit as st
import pandas as pd
import numpy as np
import requests, re, unicodedata
from pathlib import Path
from datetime import datetime, timedelta

st.set_page_config(page_title="Central de Gestão de Obras", page_icon="🏗️", layout="wide", initial_sidebar_state="expanded")

# ---------- CONFIG ----------
DATA_DIR = Path("dados_online")
DATA_DIR.mkdir(exist_ok=True)
RDO_FILE = DATA_DIR / "rdo_atual.xlsx"
MAQ_FILE = DATA_DIR / "maquinas_atual.xlsx"

def secret(name, fallback):
    try: return st.secrets.get(name, fallback)
    except: return fallback

ADMIN_PASSWORD = secret("ADMIN_PASSWORD", "trocar-admin-123")
VIEWER_PASSWORD = secret("VIEWER_PASSWORD", "engenharia-123")

# ---------- VISUAL ----------
st.markdown("""
<style>
:root{color-scheme:light}
.stApp{background:#f5f9fd;color:#102d5a}
[data-testid="stHeader"]{background:#f5f9fd}
.block-container{max-width:1600px;padding:1.15rem 1.6rem 3rem}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#eef6ff,#fff);border-right:1px solid #d9e6f3}
[data-testid="stSidebar"] *{color:#173b6c}
h1{color:#0b3475!important;font-size:2.6rem!important;font-weight:900!important;letter-spacing:-.035em}
h2,h3{color:#123c75!important;font-weight:850!important}
p,span,label{color:#284c76}
[data-testid="stWidgetLabel"],[data-testid="stWidgetLabel"] *,label,label *{color:#173b6c!important;font-weight:800!important;opacity:1!important}
[data-baseweb="select"]>div,[data-testid="stDateInput"]>div>div,[data-testid="stTextInput"] input{
 background:#fff!important;border:1px solid #d5e3f1!important;border-radius:11px!important;color:#173b6c!important}
[data-testid="stMetric"]{background:#fff!important;border:1px solid #e0e9f3!important;border-radius:17px!important;
 box-shadow:0 6px 22px rgba(35,74,120,.08)!important;padding:17px 19px!important;min-height:116px}
[data-testid="stMetricLabel"],[data-testid="stMetricLabel"] *{color:#315d8c!important;font-weight:900!important;opacity:1!important}
[data-testid="stMetricValue"],[data-testid="stMetricValue"] *{color:#082f6f!important;font-weight:950!important;opacity:1!important}
[data-testid="stMetricDelta"],[data-testid="stMetricDelta"] *{font-weight:800!important;opacity:1!important}
[data-testid="stPlotlyChart"],[data-testid="stDataFrame"]{background:#fff;border:1px solid #e0e9f3;border-radius:17px;
 box-shadow:0 5px 18px rgba(35,74,120,.07);padding:8px}
.stButton>button,.stDownloadButton>button{border-radius:10px!important;font-weight:850!important}
.stButton>button[kind="primary"]{background:#1677f2!important;color:white!important;border:0!important}
.stButton>button[kind="primary"] *{color:white!important}
[data-testid="stAlert"]{border-radius:13px}
.hero{background:linear-gradient(105deg,#edf6ff,#fff);border:1px solid #dce8f4;border-radius:20px;
 padding:18px 23px;margin-bottom:14px;box-shadow:0 6px 20px rgba(35,74,120,.06)}
.hero-title{font-size:2.15rem;font-weight:950;color:#0b3475}
.hero-sub{color:#607b99;font-weight:650}
.kpi{background:#fff;border-radius:17px;padding:16px 17px;border:1px solid #e0e9f3;box-shadow:0 6px 22px rgba(35,74,120,.08);min-height:112px}
.kpi-label{font-size:.82rem;font-weight:900;text-transform:uppercase}
.kpi-value{font-size:2rem;font-weight:950;color:#082f6f;margin-top:8px}
.blue{border-left:9px solid #2384ff}.orange{border-left:9px solid #ff982f}.green{border-left:9px solid #31c46d}.purple{border-left:9px solid #9658ee}
.red{border-left:9px solid #ff4545}.yellow{border-left:9px solid #ffc83d}.cyan{border-left:9px solid #27c6d8}.sky{border-left:9px solid #4ba4ff}
.blue .kpi-label{color:#1673df}.orange .kpi-label{color:#d76d00}.green .kpi-label{color:#0b9650}.purple .kpi-label{color:#6f2bc6}
.red .kpi-label{color:#e12c2c}.yellow .kpi-label{color:#bd8300}.cyan .kpi-label{color:#008ca4}.sky .kpi-label{color:#1673df}
.section-card{background:#fff;border:1px solid #e0e9f3;border-radius:17px;padding:14px 16px;box-shadow:0 5px 18px rgba(35,74,120,.07)}
</style>
""", unsafe_allow_html=True)

# ---------- HELPERS ----------
def norm(x):
    x="" if pd.isna(x) else str(x)
    return "".join(c for c in unicodedata.normalize("NFKD",x) if not unicodedata.combining(c)).strip().upper()

def hours(v):
    if pd.isna(v): return 0.0
    if isinstance(v, timedelta): return v.total_seconds()/3600
    if hasattr(v,"hour") and hasattr(v,"minute"): return v.hour+v.minute/60+getattr(v,"second",0)/3600
    if isinstance(v,(int,float,np.number)):
        return float(v)*24 if 0 <= float(v) < 1.5 else float(v)
    t=str(v).strip()
    m=re.match(r"^(\d+):(\d+)(?::(\d+))?$",t)
    if m:return int(m.group(1))+int(m.group(2))/60+(int(m.group(3) or 0))/3600
    try:return float(t.replace(",","."))
    except:return 0.0

@st.cache_data(show_spinner=False)
def load_rdo(path):
    x=pd.read_excel(path,sheet_name="Base de Dados Ajustada ",engine="openpyxl")
    x.columns=[str(c).strip() for c in x.columns]
    x["Data"]=pd.to_datetime(x.get("Inicio"),errors="coerce",dayfirst=True)
    x["Obra"]=x.get("Obra","").astype(str).str.strip()
    x["Nome"]=x.get("Nome do Colaborador","").astype(str).str.strip()
    x["HH"]=x.get("Total Horas do Dia",0).apply(hours)
    return x.dropna(subset=["Data"])

@st.cache_data(show_spinner=False)
def load_maq(path):
    x=pd.read_excel(path,sheet_name="Lançamentos",engine="openpyxl")
    x.columns=[str(c).strip() for c in x.columns]
    x["Data"]=pd.to_datetime(x.get("Data"),errors="coerce",dayfirst=True)
    def _col(candidatos, padrao=""):
        mapa={norm(c):c for c in x.columns}
        for cand in candidatos:
            if norm(cand) in mapa:
                return x[mapa[norm(cand)]]
        return pd.Series([padrao]*len(x), index=x.index)

    x["Obra"]=_col(["Obra / local","Obra/local","Obra","Local"]).astype(str).str.strip()
    x["Prefixo"]=_col(["Prefixo","Frota","Equipamento"]).astype(str).str.strip()
    x["HM"]=_col(["Trabalhando / Deslocamento","Trabalhando/Deslocamento","Total ligada","Horas Máquina","Horas Maquina"],0).apply(hours)
    x["Paradas"]=_col(["Parada ligada","Parada Ligada","Horas paradas","Horas Paradas"],0).apply(hours)
    x["KM"]=pd.to_numeric(_col(["Distância (km)","Distancia (km)","Distância","Distancia","KM","Km rodados"],0),errors="coerce").fillna(0)
    # combustível total informado, quando houver
    fuel_cols=[c for c in x.columns if "diesel" in norm(c).lower() or "combust" in norm(c).lower()]
    x["Combustivel"]=0.0
    for c in fuel_cols:
        vals=pd.to_numeric(x[c],errors="coerce").fillna(0)
        if vals.sum()>0: x["Combustivel"]+=vals
    return x.dropna(subset=["Data"])

@st.cache_data(show_spinner=False)
def load_clima(path):
    try:
        x=pd.read_excel(path,sheet_name="Lançamento Diário",header=2,engine="openpyxl")
        x.columns=[str(c).strip() for c in x.columns]
        x["Data"]=pd.to_datetime(x.get("Data"),errors="coerce",dayfirst=True)
        x["Obra"]=x.get("Nome da Obra","").astype(str).str.strip()
        x["Condicao"]=x.get("Condição Climática","").astype(str).str.strip()
        x["Praticavel"]=x.get("Obra Praticável?","").astype(str).str.strip()
        return x.dropna(subset=["Data"])
    except:return pd.DataFrame()

@st.cache_data(ttl=900,show_spinner=False)
def weather(city):
    q=re.sub(r"\([^)]*\)"," ",city)
    q=re.sub(r"\b(OPUB|PRIV|AT|\d{2}\.\d{2})\b"," ",q,flags=re.I)
    q=re.sub(r"\d+[-/]\d+"," ",q)
    q=re.sub(r"\s+"," ",q).strip(" -")
    # favor final textual segment
    parts=[p.strip() for p in re.split(r" - ",q) if p.strip()]
    q=parts[-1] if parts else q
    g=requests.get("https://geocoding-api.open-meteo.com/v1/search",params={"name":q,"count":1,"language":"pt","format":"json","countryCode":"BR"},timeout=8)
    rr=g.json().get("results",[])
    if not rr:return None
    loc=rr[0]
    w=requests.get("https://api.open-meteo.com/v1/forecast",params={
        "latitude":loc["latitude"],"longitude":loc["longitude"],
        "current":"temperature_2m,apparent_temperature,relative_humidity_2m,precipitation,weather_code,wind_speed_10m",
        "daily":"temperature_2m_max,temperature_2m_min,precipitation_probability_max,weather_code",
        "timezone":"America/Sao_Paulo","forecast_days":5},timeout=8).json()
    return loc,w

def wdesc(c):
    c=int(c or -1)
    if c==0:return "Céu limpo"
    if c in [1,2]:return "Parcialmente nublado"
    if c==3:return "Nublado"
    if c in [45,48]:return "Neblina"
    if c in [51,53,55]:return "Garoa"
    if c in [61,63,65,80,81,82]:return "Chuva"
    if c in [95,96,99]:return "Trovoadas"
    return "Condição variável"

def kpi(label,value,cls):
    st.markdown(f'<div class="kpi {cls}"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div></div>',unsafe_allow_html=True)

# ---------- LOGIN ----------
if "auth" not in st.session_state: st.session_state.auth=None
if st.session_state.auth is None:
    st.markdown('<div class="hero"><div class="hero-title">🏗️ Central de Gestão de Obras</div><div class="hero-sub">RDO • HH • HM • Máquinas • Caminhões • Combustível • Produtividade • Clima</div></div>',unsafe_allow_html=True)
    st.subheader("🔐 Acesso ao sistema")
    perfil=st.selectbox("Perfil",["Engenharia — somente visualização","Administrador — atualizar dados"])
    senha=st.text_input("Senha",type="password")
    if st.button("ENTRAR",type="primary",use_container_width=True):
        ok=(perfil.startswith("Administrador") and senha==ADMIN_PASSWORD) or (perfil.startswith("Engenharia") and senha==VIEWER_PASSWORD)
        if ok:
            st.session_state.auth="admin" if perfil.startswith("Administrador") else "viewer"; st.rerun()
        st.error("Senha incorreta.")
    st.stop()

# ---------- SIDEBAR ----------
with st.sidebar:
    st.markdown("## 🏗️ Central de Gestão")
    st.caption("Gestão de Obras")
    st.success("👤 Administrador" if st.session_state.auth=="admin" else "👤 Engenharia")
    st.markdown("### 🧭 Navegação")
    st.markdown("🏠 **Dashboard**  \n📋 RDO  \n🚚 Máquinas / Caminhões  \n👥 Funcionários  \n⛽ Combustível  \n🌤️ Clima  \n📊 Relatórios")
    if st.session_state.auth=="admin":
        st.divider(); st.markdown("### ⬆️ Importar dados")
        ur=st.file_uploader("Planilha RDO",type=["xlsx","xlsm"],key="rdo")
        um=st.file_uploader("Máquinas / Caminhões",type=["xlsx","xlsm"],key="maq")
        if st.button("PUBLICAR ATUALIZAÇÃO",type="primary",use_container_width=True):
            if ur:
                RDO_FILE.write_bytes(ur.getbuffer())
            if um:
                MAQ_FILE.write_bytes(um.getbuffer())
            st.cache_data.clear(); st.success("Dados publicados."); st.rerun()
    st.divider()
    if st.button("🚪 Sair",use_container_width=True):
        st.session_state.auth=None; st.rerun()

# ---------- DATA ----------
st.markdown('<div class="hero"><div class="hero-title">🏢 Central de Gestão de Obras</div><div class="hero-sub">RDO • Horas-Homem • Máquinas • Caminhões • Combustível • Produtividade • Clima</div></div>',unsafe_allow_html=True)

if not RDO_FILE.exists() and not MAQ_FILE.exists():
    st.info("O administrador precisa importar as planilhas RDO e Máquinas/Caminhões para iniciar o painel.")
    st.stop()

rdo=load_rdo(RDO_FILE) if RDO_FILE.exists() else pd.DataFrame()
maq=load_maq(MAQ_FILE) if MAQ_FILE.exists() else pd.DataFrame()
cli=load_clima(RDO_FILE) if RDO_FILE.exists() else pd.DataFrame()

obras=sorted(set((rdo["Obra"].dropna().tolist() if not rdo.empty else [])+(maq["Obra"].dropna().tolist() if not maq.empty else [])))
obras=[x for x in obras if x and x.lower()!="nan"]

# ---------- FILTERS ----------
all_dates=[]
if not rdo.empty: all_dates += rdo["Data"].dropna().tolist()
if not maq.empty: all_dates += maq["Data"].dropna().tolist()
maxd=max(all_dates).date() if all_dates else datetime.now().date()
mind=min(all_dates).date() if all_dates else maxd-timedelta(days=30)
default_start=max(mind,maxd-timedelta(days=29))

f1,f2,f3=st.columns([2.2,1,1])
obra=f1.selectbox("Obra",["Todas as obras"]+obras)
de=f2.date_input("De",value=default_start,min_value=mind,max_value=maxd)
ate=f3.date_input("Até",value=maxd,min_value=mind,max_value=maxd)

def filt(df):
    if df.empty:return df
    z=df[(df["Data"].dt.date>=de)&(df["Data"].dt.date<=ate)].copy()
    if obra!="Todas as obras": z=z[z["Obra"]==obra]
    return z
rr,mm,cc=filt(rdo),filt(maq),filt(cli)

hh=rr["HH"].sum() if not rr.empty else 0
hm=mm["HM"].sum() if not mm.empty else 0
func=rr.loc[rr["HH"]>0,"Nome"].nunique() if not rr.empty else 0
equip=mm.loc[mm["Prefixo"].str.len()>0,"Prefixo"].nunique() if not mm.empty else 0
par=mm["Paradas"].sum() if not mm.empty else 0
km=mm["KM"].sum() if not mm.empty else 0
fuel=mm["Combustivel"].sum() if not mm.empty else 0
prod=hh/hm if hm>0 else 0

# ---------- KPIs ----------
cols=st.columns(4)
with cols[0]: kpi("👥 Horas-Homem (HH)",f"{hh:,.1f}".replace(",", "X").replace(".",",").replace("X","."),"blue")
with cols[1]: kpi("🏗️ Horas-Máquina (HM)",f"{hm:,.1f}".replace(",", "X").replace(".",",").replace("X","."),"orange")
with cols[2]: kpi("👤 Funcionários únicos",str(func),"green")
with cols[3]: kpi("🚜 Equipamentos únicos",str(equip),"purple")
st.write("")
cols=st.columns(4)
with cols[0]: kpi("🕒 Horas paradas",f"{par:,.1f}".replace(".",","),"red")
with cols[1]: kpi("🛣️ KM rodados",f"{km:,.1f}".replace(",", "X").replace(".",",").replace("X","."),"yellow")
with cols[2]: kpi("⛽ Combustível (L)",f"{fuel:,.1f}".replace(",", "X").replace(".",",").replace("X","."),"cyan")
with cols[3]: kpi("📊 HH / HM — produtividade",f"{prod:.2f}".replace(".",","),"sky")

# ---------- CHARTS ----------
import plotly.express as px
import plotly.graph_objects as go
left,mid,right=st.columns([1.25,.9,.9])
with left:
    st.subheader("📅 Horas-Homem por dia")
    if not rr.empty:
        g=rr.groupby(rr["Data"].dt.date)["HH"].sum().reset_index()
        fig=px.bar(g,x="Data",y="HH",labels={"HH":"Horas","Data":""})
        fig.update_traces(marker_color="#2584ef")
        fig.update_layout(template="plotly_white",height=330,margin=dict(l=15,r=10,t=10,b=20),paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig,use_container_width=True)
    else: st.info("Sem HH no período.")
with mid:
    st.subheader("📊 Horas por tipo")
    vals=[hh,par,max(hm,0)]
    labs=["Trabalhadas","Paradas","Máquina"]
    fig=go.Figure(go.Pie(labels=labs,values=vals,hole=.62,marker=dict(colors=["#1677e8","#ff6848","#ffc43d"])))
    fig.update_layout(template="plotly_white",height=330,margin=dict(l=5,r=5,t=10,b=10),showlegend=True,paper_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(fig,use_container_width=True)
with right:
    st.subheader("🌤️ Clima em tempo real")
    if obra!="Todas as obras":
        try:
            w=weather(obra)
            if w:
                loc,d=w; cur=d["current"]
                st.markdown(f"### 📍 {loc.get('name','')} - {loc.get('admin1','')}")
                st.markdown(f"# ☀️ {cur.get('temperature_2m',0):.0f}°C")
                st.write(wdesc(cur.get("weather_code")))
                a,b=st.columns(2)
                a.metric("Umidade",f"{cur.get('relative_humidity_2m',0):.0f}%")
                b.metric("Vento",f"{cur.get('wind_speed_10m',0):.0f} km/h")
                a.metric("Sensação",f"{cur.get('apparent_temperature',0):.0f}°C")
                b.metric("Chuva agora",f"{cur.get('precipitation',0):.1f} mm")
            else: st.info("Cidade não identificada.")
        except: st.info("Clima online indisponível neste momento.")
    else: st.info("Selecione uma obra para ver o clima da cidade.")

# ---------- TABLES ----------
c1,c2,c3=st.columns(3)
with c1:
    st.subheader("👥 Top 10 Funcionários (HH)")
    if not rr.empty:
        t=rr.groupby("Nome").agg(Horas=("HH","sum"),Dias=("Data",lambda x:x.dt.date.nunique())).reset_index().sort_values("Horas",ascending=False).head(10)
        t["Horas"]=t["Horas"].round(1)
        st.dataframe(t,use_container_width=True,hide_index=True)
    else: st.info("Sem dados.")
with c2:
    st.subheader("🚚 Top 10 Máquinas / Caminhões (HM)")
    if not mm.empty:
        t=mm.groupby("Prefixo").agg(Horas=("HM","sum"),KM=("KM","sum")).reset_index().sort_values("Horas",ascending=False).head(10)
        t[["Horas","KM"]]=t[["Horas","KM"]].round(1)
        st.dataframe(t,use_container_width=True,hide_index=True)
    else: st.info("Sem dados.")
with c3:
    st.subheader("☁️ Clima registrado no RDO")
    if not cc.empty:
        t=cc.assign(Condicao=cc["Condicao"].replace({"":"Não informado"})).groupby("Condicao")["Data"].apply(lambda x:x.dt.date.nunique()).reset_index(name="Dias").sort_values("Dias",ascending=False)
        total=max(t["Dias"].sum(),1); t["%"]=(100*t["Dias"]/total).round(0).astype(int).astype(str)+"%"
        st.dataframe(t,use_container_width=True,hide_index=True)
        pratic=cc["Praticavel"].map(norm)
        st.caption(f"Dias impraticáveis registrados: {cc.loc[pratic.str.contains('NAO',na=False),'Data'].dt.date.nunique()}")
    else: st.info("Sem registros climáticos no período.")

st.caption("Central de Gestão de Obras • Dados exibidos conforme as planilhas publicadas pelo administrador.")
