# lulu.py
import re
from pathlib import Path
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="CIV – Port Arrivals (Cacao)", layout="wide")

# ---- Chemin du classeur (.xlsm) ----
DATA_FILE = Path(r"C:\Users\s.soro\Touton SA Dropbox\STATISTIQUES\ResearchFiles\Fiches_Pays.xlsm")

SHEET_DAILY  = "CIV_Arrivals_Ports_BDD"  # journalier par ports
SHEET_WEEKLY = "CIV_Arrivals_BDD"        # hebdo/cumul (Weekly_Stat / Cumul_Stat)

def _to_float(x):
    if pd.isna(x): return 0.0
    s = str(x).strip().replace("\u00A0","").replace(" ","").replace(",",".")
    if s in {"","-","—","–"}: return 0.0
    if s.count(".")>1: s = s.replace(".","")
    try: return float(s)
    except: return 0.0

def _parse_date(v):
    if isinstance(v, pd.Timestamp): return v
    s = str(v).replace("\u00A0"," ")
    s = re.sub(r"^(lundi|mardi|mercredi|jeudi|vendredi|samedi|dimanche)\s+","",s,flags=re.I)
    return pd.to_datetime(s, dayfirst=True, errors="coerce")

@st.cache_data
def load_daily(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=SHEET_DAILY, engine="openpyxl", dtype={"Tonnage":object})
    df = df.rename(columns={"Année Cacao":"AnneeCacao","Numéro Semaine":"NumeroSemaine","Numéro Jour":"NumeroJour"})
    df["Date"] = df["Date"].apply(_parse_date)
    df = df.dropna(subset=["Date"])
    df["Tonnage"] = df["Tonnage"].apply(_to_float)
    return df

@st.cache_data
def load_weekly(path: Path) -> pd.DataFrame:
    dfw = pd.read_excel(path, sheet_name=SHEET_WEEKLY, engine="openpyxl",
                        dtype={"Weekly_Stat":object,"Cumul_Stat":object})
    dfw = dfw.rename(columns={"Week_number":"NumeroSemaine","Month_number":"NumeroMois",
                              "cocoayear":"AnneeCacao","cocoayer 0000":"CocoaYear0000","cocoayear 0000":"CocoaYear0000"})
    dfw["Date"] = dfw["Date"].apply(_parse_date)
    dfw = dfw.dropna(subset=["Date"])
    dfw["Weekly_Stat"] = dfw["Weekly_Stat"].apply(_to_float)
    dfw["Cumul_Stat"]  = dfw["Cumul_Stat"].apply(_to_float)
    return dfw

# ---- Charge les données ----
try:
    df = load_daily(DATA_FILE)
    st.success(f"Journalier chargé ({len(df):,} lignes) – {SHEET_DAILY}")
except Exception as e:
    st.error(f"Erreur chargement {SHEET_DAILY}: {e}")
    st.stop()

try:
    dfw = load_weekly(DATA_FILE)
    st.info(f"Hebdo/Cumul chargé ({len(dfw):,} lignes) – {SHEET_WEEKLY}")
except Exception as e:
    st.error(f"Erreur chargement {SHEET_WEEKLY}: {e}")
    dfw = pd.DataFrame()

# ---- UI simple ----
st.title("Côte d’Ivoire – Port Arrivals")

# KPI (somme hebdo officielle)
if not dfw.empty:
    cur_years = sorted(dfw["AnneeCacao"].dropna().unique().tolist())
    annee_sel = st.selectbox("Année cacao (hebdo)", cur_years, index=len(cur_years)-1)
    curw = dfw[dfw["AnneeCacao"]==annee_sel]
    cumul_officiel = float(curw["Weekly_Stat"].sum())
    st.metric("Cumul hebdo (somme Weekly_Stat)", f"{cumul_officiel:,.1f} t")

# Aperçu journalier
st.subheader("Aperçu journalier (premières lignes)")
st.dataframe(df.head(20), use_container_width=True)

# Petit graphe hebdo si dispo
if not dfw.empty:
    curw = dfw[dfw["AnneeCacao"]==annee_sel].sort_values("NumeroSemaine")
    fig = px.bar(curw, x="Date", y="Weekly_Stat", title=f"Hebdo – {annee_sel}",
                 labels={"Weekly_Stat":"Tonnage (t)"})
    fig.update_xaxes(tickformat="%d/%m/%Y")
    st.plotly_chart(fig, use_container_width=True)
