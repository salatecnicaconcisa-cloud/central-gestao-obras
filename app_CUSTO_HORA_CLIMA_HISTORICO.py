import streamlit as st
import pandas as pd
import numpy as np
import requests, re, unicodedata, json
from pathlib import Path
from datetime import datetime, timedelta

st.set_page_config(page_title="Central de Gestão de Obras", page_icon="🏗️", layout="wide", initial_sidebar_state="expanded")

# ---------- CONFIG ----------
DATA_DIR = Path("dados_online")
DATA_DIR.mkdir(exist_ok=True)
RDO_FILE = DATA_DIR / "rdo_atual.xlsx"
MAQ_FILE = DATA_DIR / "maquinas_atual.xlsx"
CUSTO_FILE = DATA_DIR / "custos_hora_atual.xlsx"

RDO_DRIVE_ID = "1sWEG7A6KtJ3i-qwdgtaC-xVAMM7_XJnR"
MAQ_DRIVE_ID = "1XmKPzsCNVY80sK37PyRSYy_3n4gkKNkB"
CUSTO_DRIVE_ID = "17upHY_21Ki6Vpkj7q64ajcu4WhPAWnTo"

def google_credentials():
    from google.oauth2.service_account import Credentials
    raw = st.secrets.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")
    if not raw:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON não configurado nos Secrets.")
    info = json.loads(raw) if isinstance(raw, str) else dict(raw)
    return Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive.readonly"]
    )

def download_drive_file(file_id, destino):
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    import io
    service = build("drive", "v3", credentials=google_credentials(), cache_discovery=False)
    req = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    dl = MediaIoBaseDownload(fh, req, chunksize=1024*1024)
    done = False
    while not done:
        _, done = dl.next_chunk()
    destino.write_bytes(fh.getvalue())
    return service.files().get(fileId=file_id, fields="name,modifiedTime").execute()

def sincronizar_google_drive():
    return (
        download_drive_file(RDO_DRIVE_ID, RDO_FILE),
        download_drive_file(MAQ_DRIVE_ID, MAQ_FILE),
        download_drive_file(CUSTO_DRIVE_ID, CUSTO_FILE),
    )

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
    x["ObraKey"]=x["Obra"].map(norm)
    x["Nome"]=x.get("Nome do Colaborador","").astype(str).str.strip()
    x["NomeKey"]=x["Nome"].map(norm)
    x["HH"]=x.get("Total Horas do Dia",0).apply(hours)

    # Valor do dia registrado na própria base do RDO.
    valor_col=next((c for c in x.columns if norm(c)=="VALOR DIA"),None)
    if valor_col:
        x["ValorDia"]=pd.to_numeric(x[valor_col],errors="coerce")
    else:
        x["ValorDia"]=np.nan

    # Cadastro do RDO: traz Função e serve como fallback para o valor/dia.
    try:
        cad=pd.read_excel(path,sheet_name="cadastro",header=1,engine="openpyxl")
        cad.columns=[str(c).strip() for c in cad.columns]
        mp={norm(c):c for c in cad.columns}
        c_nome=mp.get("FUNCIONARIO")
        c_func=mp.get("FUNCAO")
        c_dia=mp.get("DIA")
        if c_nome:
            cad["NomeKey"]=cad[c_nome].map(norm)
            cad["FuncaoCad"]=cad[c_func].astype(str).str.strip() if c_func else ""
            cad["ValorCad"]=pd.to_numeric(cad[c_dia],errors="coerce") if c_dia else np.nan
            cad=cad[cad["NomeKey"].ne("") & cad["NomeKey"].ne("NAN")].copy()

            # Um funcionário pode aparecer mais de uma vez por setor.
            # Mantém uma função válida e o último valor/dia válido do cadastro.
            def primeira_funcao(s):
                vals=[str(v).strip() for v in s if pd.notna(v) and str(v).strip() and str(v).strip().lower()!="nan"]
                return vals[0] if vals else "Não informado"
            resumo=(cad.groupby("NomeKey",as_index=False)
                      .agg(Funcao=("FuncaoCad",primeira_funcao),
                           ValorCadastro=("ValorCad",lambda s: s.dropna().iloc[-1] if not s.dropna().empty else np.nan)))
            x=x.merge(resumo,on="NomeKey",how="left")
        else:
            x["Funcao"]="Não informado"
            x["ValorCadastro"]=np.nan
    except Exception:
        x["Funcao"]="Não informado"
        x["ValorCadastro"]=np.nan

    x["Funcao"]=x.get("Funcao","Não informado").fillna("Não informado").astype(str).str.strip()
    x["ValorDia"]=x["ValorDia"].fillna(x.get("ValorCadastro",np.nan))
    x["ValorDia"]=pd.to_numeric(x["ValorDia"],errors="coerce").fillna(0.0)

    # Só há custo do dia quando existe trabalho registrado naquele dia.
    x["CustoDia"]=np.where(x["HH"]>0,x["ValorDia"],0.0)
    return x.dropna(subset=["Data"])

@st.cache_data(show_spinner=False)
def load_custos(path):
    # Localiza automaticamente a linha do cabeçalho, pois a planilha possui títulos antes da tabela.
    bruto=pd.read_excel(path,sheet_name=0,header=None,engine="openpyxl")
    header_idx=None
    for i in range(min(30,len(bruto))):
        vals=[norm(v) for v in bruto.iloc[i].tolist()]
        tem_nome=any(v in ["FUNCIONARIO","FUNCIONÁRIO","NOME","COLABORADOR","NOME DO COLABORADOR"] for v in vals)
        tem_custo=any(("CUSTO" in v and ("/H" in v or "HORA" in v)) for v in vals)
        if tem_nome and tem_custo:
            header_idx=i; break
    if header_idx is None:
        # fallback: procura a linha que contenha CUSTO/H e usa a mesma linha como cabeçalho
        for i in range(min(50,len(bruto))):
            vals=[norm(v) for v in bruto.iloc[i].tolist()]
            if any("CUSTO/H" in v or "CUSTO H" in v or "CUSTO/HORA" in v for v in vals):
                header_idx=i; break
    if header_idx is None:
        raise ValueError("Não foi possível localizar o cabeçalho da planilha de custo/hora.")

    x=pd.read_excel(path,sheet_name=0,header=header_idx,engine="openpyxl")
    x.columns=[str(c).strip() for c in x.columns]
    mp={norm(c):c for c in x.columns}

    def acha(pred):
        for c in x.columns:
            if pred(norm(c)): return c
        return None

    c_nome=acha(lambda v: v in ["FUNCIONARIO","FUNCIONÁRIO","NOME","COLABORADOR","NOME DO COLABORADOR"] or "FUNCION" in v)
    c_func=acha(lambda v: "FUNCAO" in v or "FUNÇÃO" in v or "CARGO" in v)
    c_custo=acha(lambda v: "CUSTO" in v and ("/H" in v or "HORA" in v))

    if not c_nome or not c_custo:
        raise ValueError("A planilha de custos precisa conter Funcionário/Nome e CUSTO/H.")

    z=pd.DataFrame()
    z["Nome"]=x[c_nome].astype(str).str.strip()
    z["NomeKey"]=z["Nome"].map(norm)
    z["FuncaoCusto"]=x[c_func].astype(str).str.strip() if c_func else "Não informado"

    # Aceita número do Excel ou texto como R$ 20,41.
    s=x[c_custo]
    def dinheiro(v):
        if pd.isna(v): return np.nan
        if isinstance(v,(int,float,np.number)): return float(v)
        t=str(v).strip().replace("R$","").replace(" ","")
        if "," in t:
            t=t.replace(".","").replace(",",".")
        try:return float(t)
        except:return np.nan
    z["CustoHora"]=s.apply(dinheiro)
    z=z[(z["NomeKey"]!="") & (z["NomeKey"]!="NAN")].copy()
    z=z.dropna(subset=["CustoHora"])
    # Em caso de repetição do funcionário, usa o último custo válido.
    z=z.drop_duplicates(subset=["NomeKey"],keep="last")
    return z

@st.cache_data(show_spinner=False)
def load_maq(path):
    # A planilha de máquinas tem títulos nas linhas 2-4 e cabeçalho real na linha 5.
    # Detecta o cabeçalho automaticamente para não depender de uma posição fixa.
    bruto=pd.read_excel(path,sheet_name="Lançamentos",header=None,engine="openpyxl")
    header_idx=None
    for i in range(min(15,len(bruto))):
        vals=[norm(v) for v in bruto.iloc[i].tolist()]
        if "DATA" in vals and any("PREFIXO" in v for v in vals) and any("OBRA" in v for v in vals):
            header_idx=i
            break
    if header_idx is None:
        raise ValueError("Não foi possível localizar o cabeçalho da aba Lançamentos.")
    x=pd.read_excel(path,sheet_name="Lançamentos",header=header_idx,engine="openpyxl")
    x.columns=[str(c).strip() for c in x.columns]

    def col(names, default=""):
        mp={norm(c):c for c in x.columns}
        for name in names:
            if norm(name) in mp:
                return x[mp[norm(name)]]
        return pd.Series([default]*len(x),index=x.index)

    x["Data"]=pd.to_datetime(col(["Data"]),errors="coerce",dayfirst=True)
    x["Obra"]=col(["Obra / local","Obra/local","Obra","Local"]).astype(str).str.strip()
    x["ObraKey"]=x["Obra"].map(norm)
    x["Prefixo"]=col(["Prefixo","Frota","Equipamento"]).astype(str).str.strip()

    total=col(["Total ligada (h:mm)","Total ligada","Tempo ligado rastreador (h:mm)"],0).apply(hours)
    parada=col(["Parada ligada (h:mm)","Parada ligada"],0).apply(hours)
    trab_raw=col(["Trabalhando / Deslocamento (h:mm)","Trabalhando / Deslocamento","Trabalhando/Deslocamento"],0)

    def hm_val(v, idx):
        # Na base original esta coluna pode conter fórmula Excel.
        # Nesse caso HM = Total ligada - Parada ligada.
        if isinstance(v,str) and v.strip().startswith("="):
            return max(float(total.loc[idx])-float(parada.loc[idx]),0)
        h=hours(v)
        return h if h>0 else max(float(total.loc[idx])-float(parada.loc[idx]),0)

    x["TotalLigada"]=total
    x["HM"]=[hm_val(v,i) for i,v in trab_raw.items()]
    x["Paradas"]=parada
    x["KM"]=pd.to_numeric(col(["Distância (km)","Distancia (km)","Distância","Distancia","KM"],0),errors="coerce").fillna(0)

    # Combustível total: usa a coluna pronta quando houver; se for fórmula sem valor,
    # soma combustível trabalhando + parado ligado.
    comb_total=col(["Combustível total (L)","Combustivel total (L)"],0)
    comb_trab=pd.to_numeric(col(["Combustível andando / trabalhando (L)","Combustivel andando / trabalhando (L)"],0),errors="coerce").fillna(0)
    comb_par=pd.to_numeric(col(["Combustível só parado ligado (L)","Combustivel so parado ligado (L)"],0),errors="coerce").fillna(0)
    vals=[]
    for i,v in comb_total.items():
        if isinstance(v,str) and v.strip().startswith("="):
            vals.append(float(comb_trab.loc[i])+float(comb_par.loc[i]))
        else:
            n=pd.to_numeric(pd.Series([v]),errors="coerce").iloc[0]
            vals.append(float(n) if pd.notna(n) else float(comb_trab.loc[i])+float(comb_par.loc[i]))
    x["Combustivel"]=vals
    return x.dropna(subset=["Data"])

@st.cache_data(show_spinner=False)
def load_clima(path):
    try:
        x=pd.read_excel(path,sheet_name="Lançamento Diário",header=2,engine="openpyxl")
        x.columns=[str(c).strip() for c in x.columns]
        x["Data"]=pd.to_datetime(x.get("Data"),errors="coerce",dayfirst=True)
        x["Obra"]=x.get("Nome da Obra","").astype(str).str.strip()
        x["ObraKey"]=x["Obra"].map(norm)
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


WEATHER_LOG_FILE = DATA_DIR / "clima_online_historico.csv"

def classifica_weather_code(c):
    c=int(c if pd.notna(c) else -1)
    if c==0:return "☀️ Ensolarado"
    if c in [1,2,3,45,48]:return "☁️ Nublado"
    if c in [51,53,55,61,63,65,80,81,82]:return "🌧️ Chuva"
    if c in [95,96,99]:return "⛈️ Temporal"
    return "⚪ Não informado"

def registra_clima_online(obra_nome, loc, dados):
    if not obra_nome or obra_nome=="Todas as obras" or not dados:return
    hoje=datetime.now().date().isoformat()
    cur=dados.get("current",{})
    row={
        "Data":hoje,
        "Obra":obra_nome,
        "Cidade":loc.get("name",""),
        "UF":loc.get("admin1",""),
        "Condicao":classifica_weather_code(cur.get("weather_code",-1)),
        "Temperatura":cur.get("temperature_2m",np.nan),
        "Precipitacao_mm":cur.get("precipitation",np.nan),
        "Fonte":"Open-Meteo"
    }
    if WEATHER_LOG_FILE.exists():
        hist=pd.read_csv(WEATHER_LOG_FILE)
    else:
        hist=pd.DataFrame(columns=row.keys())
    hist=pd.concat([hist,pd.DataFrame([row])],ignore_index=True)
    hist["Data"]=hist["Data"].astype(str)
    hist["Obra"]=hist["Obra"].astype(str)
    hist=hist.drop_duplicates(subset=["Data","Obra"],keep="last")
    hist.to_csv(WEATHER_LOG_FILE,index=False)

def le_clima_online():
    if not WEATHER_LOG_FILE.exists():return pd.DataFrame()
    h=pd.read_csv(WEATHER_LOG_FILE)
    h["Data"]=pd.to_datetime(h["Data"],errors="coerce")
    h["ObraKey"]=h["Obra"].map(norm)
    return h.dropna(subset=["Data"])

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

def brl(v):
    try:
        return "R$ "+f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "R$ 0,00"

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
        if st.button("🔄 SINCRONIZAR GOOGLE DRIVE",type="primary",use_container_width=True):
            try:
                sincronizar_google_drive()
                st.cache_data.clear(); st.success("Dados atualizados do Google Drive."); st.rerun()
            except Exception as e:
                st.error(f"Não foi possível sincronizar o Google Drive: {e}")
        st.caption("Upload manual abaixo fica disponível como plano B.")
        if st.button("PUBLICAR UPLOAD MANUAL",use_container_width=True):
            if ur: RDO_FILE.write_bytes(ur.getbuffer())
            if um: MAQ_FILE.write_bytes(um.getbuffer())
            st.cache_data.clear(); st.success("Upload manual publicado."); st.rerun()
    st.divider()
    if st.button("🚪 Sair",use_container_width=True):
        st.session_state.auth=None; st.rerun()

# ---------- DATA ----------
st.markdown('<div class="hero"><div class="hero-title">🏢 Central de Gestão de Obras</div><div class="hero-sub">RDO • Horas-Homem • Máquinas • Caminhões • Combustível • Produtividade • Clima</div></div>',unsafe_allow_html=True)

drive_meta=None
try:
    drive_meta=sincronizar_google_drive()
except Exception as e:
    if not RDO_FILE.exists() and not MAQ_FILE.exists():
        st.error(f"Não foi possível carregar os dados do Google Drive: {e}")
        st.stop()
    st.warning("Google Drive temporariamente indisponível. Exibindo a última cópia carregada.")

rdo=load_rdo(RDO_FILE) if RDO_FILE.exists() else pd.DataFrame()
maq=load_maq(MAQ_FILE) if MAQ_FILE.exists() else pd.DataFrame()
cli=load_clima(RDO_FILE) if RDO_FILE.exists() else pd.DataFrame()
custos=load_custos(CUSTO_FILE) if CUSTO_FILE.exists() else pd.DataFrame()

# Cruza RDO x planilha de custos pelo nome normalizado.
if not rdo.empty:
    if "NomeKey" not in rdo.columns:
        rdo["NomeKey"]=rdo["Nome"].map(norm)
    if not custos.empty:
        rdo=rdo.merge(custos[["NomeKey","FuncaoCusto","CustoHora"]],on="NomeKey",how="left")
        rdo["Funcao"]=rdo["FuncaoCusto"].where(rdo["FuncaoCusto"].notna() & rdo["FuncaoCusto"].astype(str).str.strip().ne(""),rdo.get("Funcao","Não informado"))
        rdo["CustoHora"]=pd.to_numeric(rdo["CustoHora"],errors="coerce").fillna(0.0)
    else:
        rdo["CustoHora"]=0.0
    rdo["CustoMaoObra"]=rdo["HH"]*rdo["CustoHora"]

if drive_meta:
    try:
        dt1=pd.to_datetime(drive_meta[0]["modifiedTime"],utc=True).tz_convert("America/Sao_Paulo")
        dt2=pd.to_datetime(drive_meta[1]["modifiedTime"],utc=True).tz_convert("America/Sao_Paulo")
        dt3=pd.to_datetime(drive_meta[2]["modifiedTime"],utc=True).tz_convert("America/Sao_Paulo")
        ultima=max(dt1,dt2,dt3)
        st.caption(f"🟢 Google Drive conectado • Última alteração das fontes: {ultima.strftime('%d/%m/%Y %H:%M')}")
    except Exception:
        st.caption("🟢 Google Drive conectado")

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
    if obra!="Todas as obras":
        if "ObraKey" in z.columns:
            z=z[z["ObraKey"]==norm(obra)]
        else:
            z=z[z["Obra"].map(norm)==norm(obra)]
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

# Custo de mão de obra = HH da RDO x custo total/hora da planilha de custos.
if not rr.empty and "CustoMaoObra" in rr.columns:
    custo_mao_obra=float(rr["CustoMaoObra"].sum())
else:
    custo_mao_obra=0.0

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
st.write("")
custo_col,_,_,_=st.columns(4)
with custo_col:
    kpi("💰 Custo total mão de obra",brl(custo_mao_obra),"green")

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
                registra_clima_online(obra,loc,d)
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
c1=st.container()
c2=st.container()
c3=st.container()
with c1:
    st.subheader("👥 Funcionários • Função • Custo/Hora")
    if not rr.empty:
        trab=rr[rr["HH"]>0].copy()
        if not trab.empty:
            t=(trab.groupby(["Nome","NomeKey"],as_index=False)
               .agg(**{
                   "Função":("Funcao",lambda s: next((str(v).strip() for v in s if pd.notna(v) and str(v).strip() and str(v).strip().lower()!="nan"),"Não informado")),
                   "Custo/H":("CustoHora","last"),
                   "Horas":("HH","sum"),
                   "Custo Total":("CustoMaoObra","sum")
               }))
            t["Horas"]=pd.to_numeric(t["Horas"],errors="coerce").fillna(0).round(1)
            t["Custo/H"]=t["Custo/H"].apply(brl)
            t["Custo Total"]=t["Custo Total"].apply(brl)
            t=t[["Nome","Função","Custo/H","Horas","Custo Total"]].sort_values("Horas",ascending=False)
            st.dataframe(t,use_container_width=True,hide_index=True,height=min(900,80+35*len(t)))
            st.caption(f"💰 Total de mão de obra no período: {brl(custo_mao_obra)}")
        else:
            st.info("Sem funcionários com horas trabalhadas no período.")
    else: st.info("Sem dados.")

with c2:
    st.subheader("🚚 Máquinas / Caminhões da Obra")
    if not mm.empty:
        t=(mm.groupby("Prefixo")
             .agg(**{
                 "Total ligada (h)":("TotalLigada","sum"),
                 "Ligada parada (h)":("Paradas","sum"),
                 "Ligada andando (h)":("HM","sum"),
                 "KM":("KM","sum")
             })
             .reset_index()
             .sort_values("Total ligada (h)",ascending=False))
        for c in ["Total ligada (h)","Ligada parada (h)","Ligada andando (h)","KM"]:
            t[c]=pd.to_numeric(t[c],errors="coerce").fillna(0).round(1)
        st.dataframe(t,use_container_width=True,hide_index=True,height=min(900,80+35*len(t)))
    else: st.info("Sem dados.")
with c3:
    st.subheader("☁️ Clima registrado no RDO")
    if not cc.empty:
        base_clima=cc.copy()
        base_clima["Dia"]=base_clima["Data"].dt.date

        # Normaliza condição climática e considera também "chuva" indicada no campo de praticabilidade.
        def classifica_clima(row):
            c=norm(row.get("Condicao",""))
            p=norm(row.get("Praticavel",""))
            if any(k in c for k in ["CHUVA","CHUVOS","PRECIPIT","GAROA","TEMPORAL"]):
                return "🌧️ Chuvoso"
            if any(k in c for k in ["NUBL","ENCoberto".upper()]):
                return "☁️ Nublado"
            if any(k in c for k in ["SOL","ENSOLAR","LIMPO"]):
                return "☀️ Ensolarado"
            return "⚪ Não informado"

        base_clima["Clima"]=base_clima.apply(classifica_clima,axis=1)
        t=(base_clima.groupby("Clima",dropna=False)
           .agg(Dias=("Dia","nunique"))
           .reset_index()
           .sort_values("Dias",ascending=False))
        t["Dias"]=pd.to_numeric(t["Dias"],errors="coerce").fillna(0).astype(int)
        total=max(int(t["Dias"].sum()),1)
        t["%"]=(100*t["Dias"]/total).round(0).astype(int).astype(str)+"%"
        st.dataframe(t,use_container_width=True,hide_index=True,height=220)

        pratic=base_clima["Praticavel"].map(norm)
        dias_impraticaveis=base_clima.loc[pratic.str.contains("NAO",na=False),"Dia"].nunique()
        # "Sim" = praticável. "Parcial" fica separado e não entra como dia totalmente praticável.
        dias_praticaveis=base_clima.loc[pratic.str.match(r"^SIM$",na=False),"Dia"].nunique()
        dias_parcial=base_clima.loc[pratic.str.contains("PARCIAL",na=False),"Dia"].nunique()
        dias_chuva=base_clima.loc[base_clima["Clima"].str.contains("Chuvoso",case=False,na=False),"Dia"].nunique()

        a,b=st.columns(2)
        a.metric("✅ Dias praticáveis",int(dias_praticaveis))
        b.metric("🚫 Dias impraticáveis",int(dias_impraticaveis))
        if dias_parcial:
            st.caption(f"⚠️ Dias parcialmente praticáveis: {int(dias_parcial)}")
    else:
        st.info("Sem registros climáticos no período.")


st.write("")
st.subheader("🌐 Histórico climático online")
hist_online=le_clima_online()
if not hist_online.empty:
    ho=hist_online[(hist_online["Data"].dt.date>=de)&(hist_online["Data"].dt.date<=ate)].copy()
    if obra!="Todas as obras":
        ho=ho[ho["ObraKey"]==norm(obra)]
    if not ho.empty:
        resumo=(ho.groupby("Condicao",as_index=False)
                .agg(Dias=("Data",lambda s:s.dt.date.nunique()))
                .sort_values("Dias",ascending=False))
        a,b=st.columns([1.1,1.9])
        with a:
            st.dataframe(resumo,use_container_width=True,hide_index=True,height=min(420,80+35*len(resumo)))
        with b:
            det=ho[["Data","Obra","Cidade","Condicao","Temperatura","Precipitacao_mm","Fonte"]].copy()
            det["Data"]=det["Data"].dt.strftime("%d/%m/%Y")
            det=det.rename(columns={"Temperatura":"Temp. °C","Precipitacao_mm":"Chuva mm"})
            st.dataframe(det.sort_values("Data",ascending=False),use_container_width=True,hide_index=True,height=420)
        st.caption("Fonte: Open-Meteo • Registro online salvo automaticamente quando o painel consulta o clima da obra.")
    else:
        st.info("Ainda não há registros online para a obra/período selecionado.")
else:
    st.info("O histórico online começará a ser formado a partir das próximas consultas do clima.")


st.caption("Central de Gestão de Obras • Fontes: RDO, Máquinas/Caminhões e Custos/Hora no Google Drive • Clima online: Open-Meteo.")
