
import streamlit as st
import requests
import pandas as pd
import numpy as np
from datetime import timedelta
import io, re, unicodedata
import os, hashlib
from pathlib import Path

st.set_page_config(page_title="Central de Gestão de Obras", page_icon="🏗️", layout="wide")

st.markdown("""
<style>
.stApp {background:#07101f; color:#f8fafc;}
.block-container {padding-top:1.2rem; max-width:1500px;}
[data-testid="stSidebar"] {background:#0d1728;}
div[data-testid="stMetric"] {
    background:linear-gradient(145deg,#111d31,#0d1728);
    border:1px solid #24334d; border-radius:16px; padding:18px;
}
div[data-testid="stMetricLabel"] {color:#94a3b8;}
div[data-testid="stMetricValue"] {color:#f8fafc;}
h1,h2,h3 {color:#f8fafc;}
.small {color:#94a3b8;font-size:13px}
.ok {padding:10px 14px;border-radius:10px;background:#0b2b25;border:1px solid #166534}
</style>
""", unsafe_allow_html=True)

def norm(s):
    if pd.isna(s): return ""
    s=str(s).strip().upper()
    s="".join(c for c in unicodedata.normalize("NFD",s) if unicodedata.category(c)!="Mn")
    return re.sub(r"\s+"," ",s)

def hours(v):
    if pd.isna(v): return 0.0
    if isinstance(v, timedelta): return v.total_seconds()/3600
    if hasattr(v, "hour") and not isinstance(v, str):
        return v.hour + v.minute/60 + getattr(v,"second",0)/3600
    if isinstance(v,(int,float,np.number)):
        x=float(v)
        return x*24 if 0 <= x < 2 else x
    s=str(v).strip()
    m=re.match(r"^(\d+):(\d+)(?::(\d+))?$",s)
    if m: return int(m.group(1))+int(m.group(2))/60+int(m.group(3) or 0)/3600
    try: return float(s.replace(",",".")) 
    except: return 0.0

def find_sheet(xls, preferred):
    for p in preferred:
        for s in xls.sheet_names:
            if norm(s)==norm(p): return s
    return None

def read_rdo(file):
    xls=pd.ExcelFile(file)
    sheet=find_sheet(xls,["Base de Dados Ajustada","Base de Dados Ajustada "])
    if not sheet: raise ValueError("Não encontrei a aba 'Base de Dados Ajustada'.")
    df=pd.read_excel(file,sheet_name=sheet)
    cols={norm(c):c for c in df.columns}
    def c(*names):
        for n in names:
            if norm(n) in cols:return cols[norm(n)]
        return None
    cn,co,cd,ch=c("Nome do Colaborador"),c("Obra"),c("Inicio","Início"),c("Total Horas do Dia")
    if not all([cn,co,cd]): raise ValueError("RDO sem colunas essenciais: Colaborador, Obra e Início.")
    out=pd.DataFrame({"Data":pd.to_datetime(df[cd],errors="coerce"),"Obra":df[co],"Colaborador":df[cn]})
    out["HH"]=df[ch].map(hours) if ch else 0.0
    out=out.dropna(subset=["Data","Obra","Colaborador"])
    out["Obra"]=out["Obra"].astype(str).str.strip()
    out["Colaborador"]=out["Colaborador"].astype(str).str.strip()
    out["KEY"]=out["Data"].dt.strftime("%Y-%m-%d")+"|"+out["Obra"].map(norm)+"|"+out["Colaborador"].map(norm)
    # Evita duplicar uma mesma pessoa/data/obra em reimportações/linhas idênticas:
    out=out.groupby(["KEY","Data","Obra","Colaborador"],as_index=False)["HH"].max()
    return out

def read_maquinas(file):
    xls=pd.ExcelFile(file)
    sheet=find_sheet(xls,["Lançamentos"])
    if not sheet: raise ValueError("Não encontrei a aba 'Lançamentos'.")
    # A planilha original possui cabeçalho algumas linhas abaixo; localizar "Data"
    raw=pd.read_excel(file,sheet_name=sheet,header=None,nrows=15)
    header=0
    for i,row in raw.iterrows():
        vals=[norm(x) for x in row.tolist()]
        if "DATA" in vals and any("OBRA" in x for x in vals):
            header=i; break
    df=pd.read_excel(file,sheet_name=sheet,header=header)
    cols={norm(c):c for c in df.columns}
    def c(*names):
        for n in names:
            if norm(n) in cols:return cols[norm(n)]
        return None
    cd=c("Data"); co=c("Obra / local","Obra"); cp=c("Prefixo")
    ct=c("Total ligada (h:mm)"); ci=c("Parada ligada (h:mm)")
    ckm=c("Distância (km)"); cf=c("Combustível total (L)")
    if not all([cd,co,cp]): raise ValueError("Base de máquinas sem Data, Obra/local ou Prefixo.")
    out=pd.DataFrame({"Data":pd.to_datetime(df[cd],errors="coerce"),"Obra":df[co],"Prefixo":df[cp]})
    out["Ligada"]=df[ct].map(hours) if ct else 0.0
    out["Parada"]=df[ci].map(hours) if ci else 0.0
    out["HM"]=(out["Ligada"]-out["Parada"]).clip(lower=0)
    out["KM"]=pd.to_numeric(df[ckm],errors="coerce").fillna(0) if ckm else 0
    out["Combustivel"]=pd.to_numeric(df[cf],errors="coerce").fillna(0) if cf else 0
    out=out.dropna(subset=["Data","Obra","Prefixo"])
    out["Obra"]=out["Obra"].astype(str).str.strip()
    out["Prefixo"]=out["Prefixo"].astype(str).str.strip()
    out["KEY"]=out["Data"].dt.strftime("%Y-%m-%d")+"|"+out["Obra"].map(norm)+"|"+out["Prefixo"].map(norm)
    # Uma linha consolidada por data/obra/prefixo; max evita dobrar ao importar cópia repetida
    out=out.groupby(["KEY","Data","Obra","Prefixo"],as_index=False).agg(
        Ligada=("Ligada","max"),Parada=("Parada","max"),HM=("HM","max"),
        KM=("KM","max"),Combustivel=("Combustivel","max"))
    return out


st.markdown("""
<style>
/* V2 - títulos e filtros com alto contraste */
div[data-testid="stMetric"] {
 background:linear-gradient(145deg,#111d31,#0d1728)!important;
 border:1px solid #2d405f!important;border-radius:16px!important;padding:18px!important;
}
div[data-testid="stMetricLabel"], div[data-testid="stMetricLabel"] p {
 color:#FFFFFF!important;font-weight:800!important;font-size:0.90rem!important;
}
div[data-testid="stMetricValue"], div[data-testid="stMetricValue"] * {
 color:#FFFFFF!important;font-weight:800!important;
}
div[data-testid="stWidgetLabel"] p, label[data-testid="stWidgetLabel"] p {
 color:#E2E8F0!important;font-weight:800!important;
}
</style>
""", unsafe_allow_html=True)


st.markdown("""
<style>
[data-testid="stMetricLabel"],[data-testid="stMetricLabel"] * {
 color:#FFFFFF !important; opacity:1 !important; font-weight:800 !important;
}
[data-testid="stMetricValue"],[data-testid="stMetricValue"] * {
 color:#FFFFFF !important; opacity:1 !important; font-weight:800 !important;
}
[data-testid="stWidgetLabel"],[data-testid="stWidgetLabel"] *,label,label * {
 color:#E2E8F0 !important; opacity:1 !important; font-weight:700 !important;
}
[data-testid="stCaptionContainer"],[data-testid="stCaptionContainer"] * {
 color:#CBD5E1 !important; opacity:1 !important;
}
</style>
""", unsafe_allow_html=True)


# ---------- TEMPO ATUAL / PREVISÃO ONLINE ----------
@st.cache_data(ttl=900, show_spinner=False)
def obter_clima_online(cidade):
    """Consulta Open-Meteo sem chave de API. Atualiza a cada ~15 min."""
    if not cidade or str(cidade).strip() == "":
        return None
    nome = str(cidade).strip()
    # Limpa nomes comuns vindos do cadastro de obras
    for termo in ["OPUB", "25.11"]:
        nome = nome.replace(termo, " ")
    import re as _re
    nome = _re.sub(r"\([^)]*\)", " ", nome)
    nome = _re.sub(r"\s+", " ", nome).strip(" -")
    # Tenta extrair a parte final mais provável como município
    partes = [x.strip() for x in nome.split(" - ") if x.strip()]
    if partes:
        nome = partes[-1]

    g = requests.get(
        "https://geocoding-api.open-meteo.com/v1/search",
        params={"name": nome, "count": 1, "language": "pt", "format": "json", "countryCode": "BR"},
        timeout=10,
    )
    g.raise_for_status()
    resultados = g.json().get("results", [])
    if not resultados:
        return None
    loc = resultados[0]
    lat, lon = loc["latitude"], loc["longitude"]

    w = requests.get(
        "https://api.open-meteo.com/v1/forecast",
        params={
            "latitude": lat, "longitude": lon,
            "current": "temperature_2m,apparent_temperature,relative_humidity_2m,precipitation,rain,weather_code,wind_speed_10m",
            "daily": "weather_code,temperature_2m_max,temperature_2m_min,precipitation_probability_max,precipitation_sum",
            "timezone": "America/Sao_Paulo",
            "forecast_days": 5,
        },
        timeout=10,
    )
    w.raise_for_status()
    return {"local": loc, "dados": w.json()}

def descricao_tempo(codigo):
    mapa = {
        0:"Céu limpo",1:"Predominantemente limpo",2:"Parcialmente nublado",3:"Nublado",
        45:"Neblina",48:"Neblina com geada",51:"Garoa leve",53:"Garoa",55:"Garoa forte",
        61:"Chuva leve",63:"Chuva moderada",65:"Chuva forte",71:"Neve leve",73:"Neve",
        75:"Neve forte",80:"Pancadas leves",81:"Pancadas de chuva",82:"Pancadas fortes",
        95:"Trovoadas",96:"Trovoadas com granizo",99:"Trovoadas fortes com granizo"
    }
    return mapa.get(int(codigo) if codigo is not None else -1, "Condição não identificada")

def painel_clima_online(nome_obra):
    st.markdown("## 🌤️ Tempo atual e previsão da obra")
    st.caption("Consulta meteorológica online pela cidade da obra. O histórico oficial continua sendo o clima registrado no RDO.")
    try:
        clima = obter_clima_online(nome_obra)
        if not clima:
            st.info("Não consegui identificar automaticamente a cidade desta obra para consultar o tempo.")
            return
        loc, dados = clima["local"], clima["dados"]
        atual = dados.get("current", {})
        st.markdown(f"**📍 {loc.get('name','')} / {loc.get('admin1','')}**")
        a,b,c,d,e = st.columns(5)
        a.metric("TEMPERATURA", f"{atual.get('temperature_2m',0):.1f} °C")
        b.metric("SENSAÇÃO", f"{atual.get('apparent_temperature',0):.1f} °C")
        c.metric("UMIDADE", f"{atual.get('relative_humidity_2m',0):.0f}%")
        d.metric("CHUVA AGORA", f"{atual.get('precipitation',0):.1f} mm")
        e.metric("VENTO", f"{atual.get('wind_speed_10m',0):.1f} km/h")
        st.info("☁️ " + descricao_tempo(atual.get("weather_code")))

        daily=dados.get("daily",{})
        if daily.get("time"):
            import pandas as _pd
            prev=_pd.DataFrame({
                "Data": _pd.to_datetime(daily["time"]).strftime("%d/%m/%Y"),
                "Condição": [descricao_tempo(x) for x in daily["weather_code"]],
                "Mín. °C": daily["temperature_2m_min"],
                "Máx. °C": daily["temperature_2m_max"],
                "Prob. chuva": [f"{x:.0f}%" for x in daily["precipitation_probability_max"]],
                "Chuva prevista (mm)": daily["precipitation_sum"],
            })
            st.markdown("### 📅 Previsão — próximos 5 dias")
            st.dataframe(prev, use_container_width=True, hide_index=True)
    except Exception:
        st.warning("O serviço de clima não respondeu agora. Os demais dados do sistema continuam disponíveis.")


# ---------- V6 VISUAL CLARO / COLORIDO ----------
st.markdown("""
<style>
:root { color-scheme: light; }
.stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"] {
    background: linear-gradient(135deg,#f8fbff 0%,#f4f8fc 55%,#ffffff 100%) !important;
    color:#10213d !important;
}
[data-testid="stHeader"] { background:rgba(255,255,255,.92) !important; }
[data-testid="stSidebar"] {
    background:linear-gradient(180deg,#eef6ff 0%,#ffffff 100%) !important;
    border-right:1px solid #dce8f5;
}
[data-testid="stSidebar"] * { color:#17345c !important; }
h1,h2,h3,h4,h5,h6,p,span,div { color:#10213d; }
h1 { font-weight:850 !important; letter-spacing:-.02em; }
h2,h3 { font-weight:800 !important; }

[data-testid="stMetric"] {
    background:#ffffff !important;
    border:1px solid #dbe7f3 !important;
    border-radius:18px !important;
    padding:18px 20px !important;
    box-shadow:0 6px 22px rgba(34,80,130,.08) !important;
    min-height:120px;
}
[data-testid="stMetric"]:nth-of-type(4n+1){background:linear-gradient(135deg,#edf7ff,#ffffff)!important;}
[data-testid="stMetric"]:nth-of-type(4n+2){background:linear-gradient(135deg,#fff4e8,#ffffff)!important;}
[data-testid="stMetric"]:nth-of-type(4n+3){background:linear-gradient(135deg,#eafbf2,#ffffff)!important;}
[data-testid="stMetric"]:nth-of-type(4n+4){background:linear-gradient(135deg,#f4edff,#ffffff)!important;}
[data-testid="stMetricLabel"],[data-testid="stMetricLabel"] * {
    color:#31577f !important; opacity:1 !important; font-weight:800 !important;
    text-transform:uppercase; letter-spacing:.02em;
}
[data-testid="stMetricValue"],[data-testid="stMetricValue"] * {
    color:#0c2344 !important; opacity:1 !important; font-weight:900 !important;
}
[data-testid="stMetricDelta"],[data-testid="stMetricDelta"] * { opacity:1 !important; }

div[data-baseweb="select"] > div,
[data-testid="stDateInput"] > div > div,
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input {
    background:#ffffff !important; color:#10213d !important;
    border-color:#cddceb !important; border-radius:12px !important;
}
[data-testid="stWidgetLabel"],[data-testid="stWidgetLabel"] *,label,label * {
    color:#17345c !important; opacity:1 !important; font-weight:750 !important;
}
.stButton > button, .stDownloadButton > button {
    border-radius:12px !important; font-weight:800 !important;
    box-shadow:0 4px 12px rgba(20,91,200,.12);
}
.stButton > button[kind="primary"] {
    background:linear-gradient(90deg,#1677ff,#0b5fe5) !important;
    color:white !important; border:0 !important;
}
[data-testid="stDataFrame"], [data-testid="stTable"] {
    background:#ffffff !important; border-radius:16px !important;
    box-shadow:0 5px 18px rgba(34,80,130,.06);
}
[data-testid="stAlert"] { border-radius:14px !important; }
hr { border-color:#dce8f5 !important; }
[data-testid="stCaptionContainer"],[data-testid="stCaptionContainer"] * {
    color:#60758f !important; opacity:1 !important;
}
.block-container { max-width:1500px; padding-top:1.4rem; padding-bottom:3rem; }
</style>
""", unsafe_allow_html=True)


# Tema visual claro para gráficos Plotly
_plotly_chart_original = st.plotly_chart
def _plotly_chart_claro(fig, *args, **kwargs):
    try:
        fig.update_layout(
            template="plotly_white",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#17345c"),
            margin=dict(l=20,r=20,t=45,b=20),
        )
    except Exception:
        pass
    return _plotly_chart_original(fig, *args, **kwargs)
st.plotly_chart = _plotly_chart_claro


# ---------- V7 DASHBOARD EXECUTIVO CLARO ----------
st.markdown("""
<style>
:root{color-scheme:light}
html,body,[class*="css"]{font-family:Inter,Arial,sans-serif}
.stApp,[data-testid="stAppViewContainer"]{
 background:#f6f9fd!important;color:#10213d!important
}
[data-testid="stHeader"]{background:rgba(246,249,253,.92)!important}
.block-container{max-width:1540px!important;padding:1.1rem 2rem 3rem!important}
h1{font-size:2.45rem!important;color:#0b2245!important;font-weight:900!important;letter-spacing:-.04em}
h2{color:#132d50!important;font-size:1.45rem!important;font-weight:850!important}
h3{color:#17345c!important;font-weight:800!important}
p,span,label,div{color:#17345c}
[data-testid="stCaptionContainer"],[data-testid="stCaptionContainer"] *{color:#66809d!important}

/* filtros */
[data-baseweb="select"]>div,[data-testid="stDateInput"]>div>div,[data-testid="stTextInput"] input{
 background:#fff!important;border:1px solid #d7e4f1!important;border-radius:12px!important;
 box-shadow:0 2px 8px rgba(31,77,124,.04)!important;color:#10213d!important
}
[data-testid="stWidgetLabel"],[data-testid="stWidgetLabel"] *,label,label *{
 color:#17345c!important;font-weight:800!important;opacity:1!important
}

/* cards KPI */
[data-testid="stMetric"]{
 background:#fff!important;border:1px solid #dce8f4!important;border-radius:18px!important;
 padding:18px 20px!important;box-shadow:0 7px 24px rgba(34,79,124,.08)!important;
 min-height:118px;position:relative;overflow:hidden
}
[data-testid="stMetric"]::before{
 content:"";position:absolute;left:0;top:0;bottom:0;width:6px;background:#2484ff
}
[data-testid="column"]:nth-of-type(2n) [data-testid="stMetric"]::before{background:#ff9e2a}
[data-testid="column"]:nth-of-type(3n) [data-testid="stMetric"]::before{background:#20b86a}
[data-testid="column"]:nth-of-type(4n) [data-testid="stMetric"]::before{background:#8b5cf6}
[data-testid="stMetricLabel"],[data-testid="stMetricLabel"] *{
 color:#3d6388!important;font-size:.84rem!important;font-weight:900!important;
 text-transform:uppercase;opacity:1!important
}
[data-testid="stMetricValue"],[data-testid="stMetricValue"] *{
 color:#0b2245!important;font-size:2rem!important;font-weight:900!important;opacity:1!important
}
[data-testid="stMetricDelta"],[data-testid="stMetricDelta"] *{font-weight:800!important;opacity:1!important}

/* botões */
.stButton>button,.stDownloadButton>button{
 border-radius:12px!important;font-weight:850!important;min-height:42px;
 border:1px solid #cfe0f2!important
}
.stButton>button[kind="primary"]{
 background:linear-gradient(90deg,#237cff,#0a64e8)!important;color:#fff!important;border:0!important
}
.stButton>button[kind="primary"] *{color:#fff!important}

/* tabelas, alerts e containers */
[data-testid="stDataFrame"],[data-testid="stTable"],[data-testid="stExpander"]{
 background:#fff!important;border:1px solid #dce8f4!important;border-radius:16px!important;
 box-shadow:0 5px 20px rgba(34,79,124,.06)!important;overflow:hidden
}
[data-testid="stAlert"]{border-radius:14px!important}
hr{border-color:#dce8f4!important}

/* abas */
button[data-baseweb="tab"]{font-weight:800!important}
button[data-baseweb="tab"][aria-selected="true"]{color:#126bdf!important}

/* upload */
[data-testid="stFileUploaderDropzone"]{
 background:#fff!important;border:1.5px dashed #9fc4ec!important;border-radius:16px!important
}

/* Plotly container */
[data-testid="stPlotlyChart"]{
 background:#fff;border:1px solid #dce8f4;border-radius:18px;
 box-shadow:0 5px 20px rgba(34,79,124,.06);padding:8px
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div style="background:linear-gradient(100deg,#edf6ff,#ffffff);border:1px solid #dce8f4;
border-radius:20px;padding:18px 24px;margin-bottom:12px;box-shadow:0 6px 22px rgba(34,79,124,.06)">
<div style="font-size:13px;font-weight:850;color:#2474d7;letter-spacing:.08em">PAINEL EXECUTIVO</div>
<div style="font-size:15px;color:#60758f;margin-top:4px">RDO • Horas-Homem • Máquinas • Caminhões • Combustível • Produtividade • Clima</div>
</div>
""", unsafe_allow_html=True)


_dashboard_plotly_original = st.plotly_chart
def _dashboard_plotly(fig,*args,**kwargs):
    try:
        fig.update_layout(
            template="plotly_white",
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#17345c",family="Arial"),
            title_font=dict(color="#10213d",size=18),
            legend=dict(bgcolor="rgba(0,0,0,0)"),
            margin=dict(l=25,r=25,t=55,b=30)
        )
        fig.update_xaxes(gridcolor="#edf2f7",linecolor="#dbe6f0")
        fig.update_yaxes(gridcolor="#edf2f7",linecolor="#dbe6f0")
    except Exception: pass
    return _dashboard_plotly_original(fig,*args,**kwargs)
st.plotly_chart=_dashboard_plotly

st.title("🏗️ Central de Gestão de Obras")
st.markdown('<div class="small">RDO • Horas-Homem • Máquinas • Caminhões • Combustível • Produtividade</div>',unsafe_allow_html=True)




def read_clima(file):
    """Lê a aba Lançamento Diário e consolida 1 registro por obra + data."""
    df=pd.read_excel(file,sheet_name="Lançamento Diário",header=2,engine="openpyxl")
    df.columns=[str(c).strip() for c in df.columns]
    needed=["Data","Nome da Obra","Condição Climática","Obra Praticável?"]
    missing=[c for c in needed if c not in df.columns]
    if missing:
        raise ValueError("Campos climáticos não encontrados: "+", ".join(missing))
    keep=needed+[c for c in ["Horas Paralisadas","Motivo / Observação","Atividade Afetada"] if c in df.columns]
    d=df[keep].copy()
    d["Data"]=pd.to_datetime(d["Data"],errors="coerce").dt.normalize()
    d["Obra"]=d["Nome da Obra"].fillna("").astype(str).str.strip()
    d["Clima"]=d["Condição Climática"].fillna("").astype(str).str.strip().str.title()
    d["Praticavel"]=d["Obra Praticável?"].fillna("").astype(str).str.strip().str.title()
    if "Horas Paralisadas" in d:
        d["Horas_Paralisadas"]=pd.to_numeric(d["Horas Paralisadas"],errors="coerce").fillna(0)
    else:
        d["Horas_Paralisadas"]=0.0
    d=d[(d["Data"].notna()) & (d["Obra"]!="")].copy()

    # Um dia só pode contar uma vez por obra. Em conflito, prevalece o estado mais restritivo.
    def consolidate(g):
        climates=set(x for x in g["Clima"] if x)
        pratic=set(x for x in g["Praticavel"] if x)
        if "Chuva" in climates: clima="Chuva"
        elif "Nublado" in climates: clima="Nublado"
        elif "Ensolarado" in climates: clima="Ensolarado"
        else: clima=next(iter(climates),"Não informado")
        if "Não" in pratic: pr="Não"
        elif "Parcial" in pratic: pr="Parcial"
        elif "Sim" in pratic: pr="Sim"
        else: pr="Não informado"
        return pd.Series({"Clima":clima,"Praticavel":pr,
                          "Horas_Paralisadas":g["Horas_Paralisadas"].max()})
    return d.groupby(["Data","Obra"],as_index=False).apply(
        consolidate,include_groups=False
    ).reset_index(drop=True)


# --- ACESSO ONLINE / PERFIS ---
DATA_DIR=Path("dados_online")
DATA_DIR.mkdir(exist_ok=True)
RDO_PATH=DATA_DIR/"rdo_atual.xlsx"
MAQ_PATH=DATA_DIR/"maquinas_atual.xlsx"

def secret(name, default):
    try:
        return st.secrets.get(name, os.environ.get(name, default))
    except Exception:
        return os.environ.get(name, default)

ADMIN_PASSWORD=secret("ADMIN_PASSWORD","trocar-admin-123")
VIEWER_PASSWORD=secret("VIEWER_PASSWORD","engenharia-123")

if "role" not in st.session_state:
    st.session_state.role=None

if st.session_state.role is None:
    st.markdown("### 🔐 Acesso ao sistema")
    perfil=st.selectbox("Perfil",["Engenharia — somente visualização","Administrador — atualizar dados"])
    senha=st.text_input("Senha",type="password")
    if st.button("ENTRAR",use_container_width=True):
        if perfil.startswith("Administrador") and senha==ADMIN_PASSWORD:
            st.session_state.role="admin"; st.rerun()
        elif perfil.startswith("Engenharia") and senha==VIEWER_PASSWORD:
            st.session_state.role="viewer"; st.rerun()
        else:
            st.error("Senha incorreta.")
    st.stop()

with st.sidebar:
    if st.session_state.role=="admin":
        st.success("🔑 Administrador")
        st.header("Atualizar dados")
        rdo_up=st.file_uploader("📄 Planilha RDO",type=["xlsx","xlsm"],key="rdo")
        maq_up=st.file_uploader("🚜 Máquinas / Caminhões",type=["xlsx","xlsm"],key="maq")
        if st.button("💾 PUBLICAR ATUALIZAÇÃO",use_container_width=True):
            if not rdo_up or not maq_up:
                st.error("Selecione as duas planilhas.")
            else:
                # Validate before publishing
                try:
                    rdo_up.seek(0); _=read_rdo(rdo_up)
                    maq_up.seek(0); _=read_maquinas(maq_up)
                    rdo_up.seek(0); maq_up.seek(0)
                    RDO_PATH.write_bytes(rdo_up.read())
                    MAQ_PATH.write_bytes(maq_up.read())
                    st.success("Dados publicados. A engenharia já pode visualizar.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Não publiquei porque encontrei um problema: {e}")
        st.caption("Somente o Administrador consegue importar e publicar planilhas.")
    else:
        st.info("👷 Engenharia • somente consulta")
        st.caption("Você pode filtrar e consultar o BI. A atualização das planilhas é exclusiva do Administrador.")
    st.divider()
    if st.button("Sair"):
        st.session_state.role=None; st.rerun()

if not RDO_PATH.exists() or not MAQ_PATH.exists():
    if st.session_state.role=="admin":
        st.info("Importe as duas planilhas no menu lateral e clique em PUBLICAR ATUALIZAÇÃO.")
    else:
        st.info("A base ainda não foi publicada pelo Administrador.")
    st.stop()

rdo_file=str(RDO_PATH)
maq_file=str(MAQ_PATH)

try:
    hh=read_rdo(rdo_file)
    mq=read_maquinas(maq_file)
    clima=read_clima(rdo_file)
except Exception as e:
    st.error(f"Erro ao processar: {e}")
    st.stop()

obras=sorted(set(hh["Obra"]).union(set(mq["Obra"])))
c1,c2,c3=st.columns([2.5,1,1])
with c1: obra=st.selectbox("OBRA",obras)
mind=min([x for x in [hh["Data"].min(),mq["Data"].min()] if pd.notna(x)])
maxd=max([x for x in [hh["Data"].max(),mq["Data"].max()] if pd.notna(x)])
with c2: inicio=st.date_input("DE",mind.date(),min_value=mind.date(),max_value=maxd.date())
with c3: fim=st.date_input("ATÉ",maxd.date(),min_value=mind.date(),max_value=maxd.date())

h=hh[(hh.Obra==obra)&(hh.Data.dt.date>=inicio)&(hh.Data.dt.date<=fim)].copy()
m=mq[(mq.Obra==obra)&(mq.Data.dt.date>=inicio)&(mq.Data.dt.date<=fim)].copy()

hh_total=h.HH.sum(); hm=m.HM.sum(); par=m.Parada.sum(); km=m.KM.sum(); comb=m.Combustivel.sum()
func=h["Colaborador"].map(norm).nunique(); equip=m["Prefixo"].map(norm).nunique()
ratio=hh_total/hm if hm else 0

a,b,c,d=st.columns(4)
a.metric("HORAS-HOMEM (HH)",f"{hh_total:,.1f}".replace(",","."))
b.metric("HORAS-MÁQUINA (HM)",f"{hm:,.1f}".replace(",","."))
c.metric("FUNCIONÁRIOS ÚNICOS",func)
d.metric("EQUIPAMENTOS ÚNICOS",equip)
a,b,c,d=st.columns(4)
a.metric("HORAS PARADAS",f"{par:,.1f}".replace(",","."))
b.metric("KM RODADOS",f"{km:,.1f}".replace(",","."))
c.metric("COMBUSTÍVEL (L)",f"{comb:,.1f}".replace(",","."))
d.metric("HH / HM — PRODUTIVIDADE",f"{ratio:.2f}")

st.markdown('<div class="ok">✓ Dados processados. Funcionários e equipamentos dos cards são contados sem repetição no período selecionado.</div>',unsafe_allow_html=True)
st.write("")

left,right=st.columns(2)
with left:
    st.subheader("Horas-máquina por equipamento")
    x=m.groupby("Prefixo",as_index=False)["HM"].sum().sort_values("HM",ascending=False).head(15)
    st.bar_chart(x.set_index("Prefixo"),horizontal=True)
with right:
    st.subheader("Horas-homem por colaborador")
    x=h.groupby("Colaborador",as_index=False)["HH"].sum().sort_values("HH",ascending=False).head(15)
    st.bar_chart(x.set_index("Colaborador"),horizontal=True)

left,right=st.columns(2)
with left:
    st.subheader("Evolução diária • HH")
    x=h.groupby(h.Data.dt.date)["HH"].sum()
    st.line_chart(x)
with right:
    st.subheader("Evolução diária • HM")
    x=m.groupby(m.Data.dt.date)["HM"].sum()
    st.line_chart(x)

st.subheader("Detalhamento da obra")
tab1,tab2=st.tabs(["👷 Funcionários","🚜 Equipamentos"])
with tab1:
    x=h.groupby("Colaborador",as_index=False).agg(HH=("HH","sum"),Dias=("Data","nunique")).sort_values("HH",ascending=False)
    st.dataframe(x,use_container_width=True,hide_index=True)
with tab2:
    x=m.groupby("Prefixo",as_index=False).agg(HM=("HM","sum"),Parada=("Parada","sum"),KM=("KM","sum"),Combustível=("Combustivel","sum"),Dias=("Data","nunique")).sort_values("HM",ascending=False)
    st.dataframe(x,use_container_width=True,hide_index=True)

st.download_button("⬇️ Exportar resumo da obra (CSV)",
    pd.DataFrame([{"Obra":obra,"De":inicio,"Até":fim,"HH":hh_total,"HM":hm,"Funcionários":func,"Equipamentos":equip,"Horas Paradas":par,"KM":km,"Combustível L":comb,"HH/HM":ratio}]).to_csv(index=False).encode("utf-8-sig"),
    file_name="resumo_obra.csv",mime="text/csv")


st.divider()
st.subheader("🌦️ Clima e Praticabilidade da Obra")
cf=clima.copy()
if obra!="TODAS":
    cf=cf[cf["Obra"]==obra]
cf=cf[(cf["Data"].dt.date>=inicio) & (cf["Data"].dt.date<=fim)]

dias_total=int(cf["Data"].nunique())
dias_sol=int(cf.loc[cf["Clima"].eq("Ensolarado"),"Data"].nunique())
dias_chuva=int(cf.loc[cf["Clima"].eq("Chuva"),"Data"].nunique())
dias_nublado=int(cf.loc[cf["Clima"].eq("Nublado"),"Data"].nunique())
dias_pratic=int(cf.loc[cf["Praticavel"].isin(["Sim","Parcial"]),"Data"].nunique())
dias_impratic=int(cf.loc[cf["Praticavel"].eq("Não"),"Data"].nunique())
horas_par=float(cf["Horas_Paralisadas"].sum()) if not cf.empty else 0.0

a,b,c,d,e,f=st.columns(6)
a.metric("☀️ DIAS ENSOLARADOS",dias_sol)
b.metric("🌧️ DIAS DE CHUVA",dias_chuva)
c.metric("☁️ DIAS NUBLADOS",dias_nublado)
d.metric("✅ DIAS PRATICÁVEIS",dias_pratic)
e.metric("⛔ DIAS IMPRATICÁVEIS",dias_impratic)
f.metric("⏱️ HORAS PARALISADAS",f"{horas_par:,.1f}".replace(",", "X").replace(".", ",").replace("X","."))

st.caption("Contagem climática consolidada por OBRA + DATA: o mesmo dia não é contado várias vezes.")
if not cf.empty:
    mensal=cf.assign(Mês=cf["Data"].dt.to_period("M").astype(str)).groupby("Mês").agg(
        Dias=("Data","nunique"),
        Ensolarados=("Clima",lambda x:(x=="Ensolarado").sum()),
        Chuva=("Clima",lambda x:(x=="Chuva").sum()),
        Nublados=("Clima",lambda x:(x=="Nublado").sum()),
        Praticáveis=("Praticavel",lambda x:x.isin(["Sim","Parcial"]).sum()),
        Impraticáveis=("Praticavel",lambda x:(x=="Não").sum()),
        Horas_Paralisadas=("Horas_Paralisadas","sum")
    ).reset_index()
    st.markdown("#### Resumo climático mensal")
    st.dataframe(mensal,use_container_width=True,hide_index=True)
    with st.expander("Ver dias e condições climáticas"):
        st.dataframe(cf.sort_values("Data"),use_container_width=True,hide_index=True)
else:
    st.info("Não há lançamentos climáticos para a obra/período selecionado.")



# ---------- PAINEL METEOROLÓGICO ONLINE ----------
try:
    _obra_tempo = None
    for _nome_var in ["obra_sel", "obra_selecionada", "obra_filtro", "obra"]:
        if _nome_var in globals():
            _v = globals()[_nome_var]
            if isinstance(_v, str) and _v and _v.lower() not in ["todas", "todos", "geral"]:
                _obra_tempo = _v
                break
    if _obra_tempo:
        painel_clima_online(_obra_tempo)
    else:
        st.markdown("## 🌤️ Tempo atual das obras")
        st.caption("Selecione uma obra específica nos filtros para consultar automaticamente o tempo atual e a previsão da cidade.")
except Exception:
    pass
