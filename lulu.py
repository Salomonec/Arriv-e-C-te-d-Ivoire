# lulu.py — Dashboard Côte d'Ivoire (Cacao) + Prévisions BOTA
# -----------------------------------------------------------------------------
# 1) Cumul campagne (hebdo/mensuel) + superposition multi-années + BOTA
# 2) Comparaison Main/Mid (hebdo officiel) — HISTOGRAMMES + slider de semaines + LTA
# 3) Cumuls hebdomadaires – Multi-années + LTA (officiel) + BOTA
# 4) Répartitions (camemberts) : ports, jours de la semaine (moyenne hebdo)
# Style : fond blanc, grille grise, cadre noir ; titres/axes en gras
# Chargement robuste : chemin local OU lien Dropbox lecture seule
# -----------------------------------------------------------------------------

import re
from io import BytesIO
from pathlib import Path
import requests
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="CIV – Port Arrivals (Cacao)", layout="wide")

# ========= PARAM FICHIER =========
LOCAL_XLSM = r"C:\Users\s.soro\Touton SA Dropbox\STATISTIQUES\ResearchFiles\Fiches_Pays.xlsm"
DROPBOX_URL = (
    "https://www.dropbox.com/scl/fi/wyzhzkzx5ddxug3qgdvwt/Fiches_Pays.xlsm"
    "?rlkey=vqauw3m1v3c1tc8r0blx7rng3&dl=1"
)

SHEET_DAILY  = "CIV_Arrivals_Ports_BDD"
SHEET_WEEKLY = "CIV_Arrivals_BDD"
SHEET_BOTA   = "CIV_Bota_Arrivals_Treatments"

# ========= PALETTE & STYLES =========
def build_palette_long():
    names = [
        "Set3","Set1","Set2","Pastel1","Pastel2","Paired","Bold","Colorblind",
        "Dark24","Light24","Prism","Vivid","Antique","D3","G10","T10",
        "Alphabet","Safe","Picnic","Plotly"
    ]
    out = []
    for n in names:
        pal = getattr(px.colors.qualitative, n, None)
        if pal: out += pal
    return out or px.colors.qualitative.Plotly

BOLD_FONT = "Arial Black, Arial, sans-serif"
WEEK_MS = 7 * 24 * 60 * 60 * 1000

def _next_sunday(d: pd.Timestamp) -> pd.Timestamp:
    d = pd.Timestamp(d).normalize()
    off = (6 - d.weekday()) % 7
    return d + pd.Timedelta(days=off)

def weekly_xaxis_on_sundays(anchor: pd.Timestamp) -> dict:
    tick0 = _next_sunday(anchor)
    return dict(title="Date", tickformat="%d/%m/%Y", tickmode="linear", tick0=tick0, dtick=WEEK_MS)

def style_fig(fig, *, title=None, xaxis=None, yaxis=None, bg="transparent"):
    paper = "rgba(0,0,0,0)" if bg=="transparent" else "white"
    fig.update_layout(
        paper_bgcolor=paper, plot_bgcolor=paper,
        font=dict(family=BOLD_FONT, size=12),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
                    font=dict(family=BOLD_FONT, size=11), bgcolor="rgba(255,255,255,0)")
    )
    if title is not None:
        fig.update_layout(title=f"<b>{title}</b>", title_font=dict(family=BOLD_FONT, size=16))
    if xaxis is not None:
        xa = dict(xaxis); xa["title"] = f"<b>{xa.get('title','')}</b>"
        xa["tickfont"]  = dict(family=BOLD_FONT, size=11)
        xa["title_font"] = dict(family=BOLD_FONT, size=12)
        fig.update_xaxes(**xa)
    if yaxis is not None:
        ya = dict(yaxis); ya["title"] = f"<b>{ya.get('title','')}</b>"
        ya["tickfont"]  = dict(family=BOLD_FONT, size=11)
        ya["title_font"] = dict(family=BOLD_FONT, size=12)
        fig.update_yaxes(**ya)
    return fig

def style_excel_like(fig, *, title, base_start, ylabel, y_max_hint=None):
    xax = weekly_xaxis_on_sundays(base_start)
    xax["tickangle"] = 45
    fig.update_layout(
        paper_bgcolor="white", plot_bgcolor="white",
        title=f"<b>{title}</b>", title_font=dict(family=BOLD_FONT, size=16),
        font=dict(family=BOLD_FONT, size=12),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
                    bgcolor="rgba(255,255,255,0)", font=dict(family=BOLD_FONT, size=11)),
        margin=dict(l=70, r=40, t=70, b=70),
    )
    fig.update_xaxes(**xax, showgrid=True, gridcolor="#cccccc")
    if y_max_hint is None or y_max_hint <= 0:
        fig.update_yaxes(title=f"<b>{ylabel}</b>", showgrid=True, gridcolor="#cccccc",
                         tickformat=",.0f", tickfont=dict(family=BOLD_FONT, size=11),
                         title_font=dict(family=BOLD_FONT, size=12))
    else:
        top = int((y_max_hint + 499_999) // 500_000) * 500_000
        fig.update_yaxes(title=f"<b>{ylabel}</b>", range=[0, top], dtick=500_000,
                         tickformat=",.0f", showgrid=True, gridcolor="#cccccc",
                         tickfont=dict(family=BOLD_FONT, size=11),
                         title_font=dict(family=BOLD_FONT, size=12))
    fig.add_shape(type="rect", xref="paper", yref="paper", x0=0, y0=0, x1=1, y1=1,
                  line=dict(color="black", width=1))
    return fig

def add_value_labels(fig, unit=" t", extra_headroom=0.18):
    fig.update_traces(
        texttemplate="<b>%{y:,.0f}" + unit + "</b>",
        textposition="outside",
        cliponaxis=False,
        textfont=dict(family=BOLD_FONT, color="black", size=12),
        hovertemplate="%{x}: %{y:,.0f}" + unit
    )
    ymax = 0.0
    for tr in fig.data:
        try:
            vals = [v for v in tr.y if v is not None]
            if vals: ymax = max(ymax, float(max(vals)))
        except Exception: pass
    if ymax > 0:
        fig.update_yaxes(range=[0, ymax * (1.0 + extra_headroom)])
    return fig

# ========= OUTILS LECTURE EXCEL =========
def _open_excel(path_or_url: str) -> BytesIO:
    p = str(path_or_url)
    if p.lower().startswith(("c:\\", "d:\\", "\\\\", "/")) or Path(p).exists():
        with open(p, "rb") as f: return BytesIO(f.read())
    if "dropbox.com" in p and "dl=1" not in p:
        if "dl=0" in p: p = p.replace("dl=0", "dl=1")
        else: p = p + ("&" if "?" in p else "?") + "dl=1"
    resp = requests.get(p, timeout=60); resp.raise_for_status()
    return BytesIO(resp.content)

# ========= HELPERS =========
def _to_float(x):
    if pd.isna(x): return 0.0
    s = str(x).strip()
    if s in {"-", "—", "–", ""}: return 0.0
    s = s.replace("\u00A0", "").replace(" ", "").replace(",", ".")
    if s.count(".") > 1: s = s.replace(".", "")
    try: return float(s)
    except Exception: return 0.0

def _parse_date(v):
    if isinstance(v, pd.Timestamp): return v
    s = str(v).replace("\u00A0", " ")
    s = re.sub(r"^(lundi|mardi|mercredi|jeudi|vendredi|samedi|dimanche)\s+", "", s, flags=re.I)
    return pd.to_datetime(s, dayfirst=True, errors="coerce")

def _start_year_from_label(label: str) -> int:
    s = str(label).strip()
    if s.lower() in {"nan", "", "none"}: return pd.Timestamp.today().year
    try:
        yy = int(s.split("/")[0]); return 2000 + yy if yy < 50 else 1900 + yy
    except Exception:
        try: return int(s[:4])
        except Exception: return pd.Timestamp.today().year

def _safe_index(lst, value):
    try: return lst.index(value)
    except ValueError: return len(lst) - 1 if lst else 0

def safe_concat(dfs, cols=None):
    cleaned = []
    for d in dfs:
        if isinstance(d, pd.DataFrame) and not d.empty:
            if cols is not None and all(c in d.columns for c in cols):
                cleaned.append(d[cols].copy())
            elif cols is None:
                cleaned.append(d.copy())
    if cleaned: return pd.concat(cleaned, ignore_index=True)
    return pd.DataFrame(columns=cols or [])

# ========= LOADERS =========
@st.cache_data
def load_daily_ports(source_path_or_url: str, sheet: str) -> pd.DataFrame:
    bio = _open_excel(source_path_or_url)
    df = pd.read_excel(bio, sheet_name=sheet, engine="openpyxl", dtype={"Tonnage": object})
    df = df.rename(columns={"Année Cacao": "AnneeCacao",
                            "Numéro Semaine": "NumeroSemaine",
                            "Numéro Jour": "NumeroJour"})
    df["Date"] = df["Date"].apply(_parse_date)
    df = df.dropna(subset=["Date"])
    df["Tonnage"] = df["Tonnage"].apply(_to_float)
    df["NumeroSemaine"] = pd.to_numeric(df["NumeroSemaine"], errors="coerce").fillna(0).astype(int)
    df["NumeroJour"]    = pd.to_numeric(df["NumeroJour"],    errors="coerce").fillna(0).astype(int)
    df["CocoaYearStart"] = df["AnneeCacao"].map(_start_year_from_label)
    start_dates = pd.to_datetime(dict(year=df["CocoaYearStart"], month=10, day=1))
    df["JourCacao"] = (df["Date"].dt.normalize() - start_dates).dt.days + 1
    df["JourSemaine"] = df["Date"].dt.dayofweek + 1
    return df

@st.cache_data
def load_weekly_cumul(source_path_or_url: str, sheet: str) -> pd.DataFrame:
    bio = _open_excel(source_path_or_url)
    dfw = pd.read_excel(bio, sheet_name=sheet, engine="openpyxl",
                        dtype={"Weekly_Stat": object, "Cumul_Stat": object})
    dfw = dfw.rename(columns={
        "cocoayear": "AnneeCacao",
        "cocoayer 0000": "CocoaYear0000",
        "cocoayear 0000": "CocoaYear0000",
        "Week_number": "NumeroSemaine",
        "Month_number": "NumeroMois",
    })
    dfw["Date"] = dfw["Date"].apply(_parse_date)
    dfw = dfw.dropna(subset=["Date"])
    dfw["Weekly_Stat"] = dfw["Weekly_Stat"].apply(_to_float)
    dfw["Cumul_Stat"]  = dfw["Cumul_Stat"].apply(_to_float)
    dfw["NumeroSemaine"] = pd.to_numeric(dfw["NumeroSemaine"], errors="coerce").fillna(0).astype(int)

    if "AnneeCacao" in dfw.columns and dfw["AnneeCacao"].notna().any():
        dfw["CocoaYearStart"] = dfw["AnneeCacao"].map(_start_year_from_label)
    else:
        dfw["CocoaYearStart"] = pd.to_numeric(dfw["CocoaYear0000"], errors="coerce").fillna(0).astype(int)

    base = pd.to_datetime(dict(year=dfw["CocoaYearStart"], month=10, day=1))
    dfw["BaseDate"] = base + pd.to_timedelta((dfw["NumeroSemaine"] - 1) * 7, unit="D")
    return dfw

@st.cache_data
def load_bota(source_path_or_url: str, sheet: str) -> pd.DataFrame:
    bio = _open_excel(source_path_or_url)
    raw = pd.read_excel(bio, sheet_name=sheet, engine="openpyxl", header=0)
    cols = list(raw.columns)
    block_starts = []
    for i, c in enumerate(cols):
        if str(c).strip().lower().startswith("date"):
            win = [str(x) for x in cols[i:i+8]]
            if any("Week_number" in w for w in win):
                block_starts.append(i)

    frames = []
    for i in block_starts:
        block = cols[i:i+8]
        sub = raw[block].copy()
        sub = sub.rename(columns={c: str(c).split(".")[0] for c in sub.columns})
        for nm in ["Date","cocoayear","Week_number","Weekly_Bota","Cumul_Bota"]:
            if nm not in sub.columns: sub[nm] = None
        sub["Date"] = sub["Date"].apply(_parse_date)
        sub = sub.dropna(subset=["Date","cocoayear"])
        sub["AnneeCacao"]    = sub["cocoayear"].astype(str)
        sub["NumeroSemaine"] = pd.to_numeric(sub["Week_number"], errors="coerce").fillna(0).astype(int)
        sub["Weekly_Bota"] = sub["Weekly_Bota"].apply(_to_float) * 1000.0
        sub["Cumul_Bota"]  = sub["Cumul_Bota"].apply(_to_float)  * 1000.0
        sub["CocoaYearStart"] = sub["AnneeCacao"].map(_start_year_from_label)
        sub["BaseDate"] = (
            pd.to_datetime(dict(year=sub["CocoaYearStart"], month=10, day=1))
            + pd.to_timedelta((sub["NumeroSemaine"]-1)*7, unit="D")
        )
        frames.append(
            sub[["Date","AnneeCacao","CocoaYearStart","NumeroSemaine",
                 "Weekly_Bota","Cumul_Bota","BaseDate"]]
        )

    if not frames:
        return pd.DataFrame(columns=["Date","AnneeCacao","CocoaYearStart","NumeroSemaine","Weekly_Bota","Cumul_Bota","BaseDate"])
    out = pd.concat(frames, ignore_index=True)
    out = out.dropna(subset=["AnneeCacao"])
    return out

# ========= CHARGEMENT =========
SOURCE = LOCAL_XLSM if Path(LOCAL_XLSM).exists() else DROPBOX_URL

try:
    df = load_daily_ports(SOURCE, SHEET_DAILY)
except Exception as e:
    st.error(f"Erreur chargement journalier: {e}")
    df = pd.DataFrame()

try:
    dfw = load_weekly_cumul(SOURCE, SHEET_WEEKLY)
except Exception as e:
    st.error(f"Erreur chargement hebdo/cumul: {e}")
    dfw = pd.DataFrame()

try:
    dfb = load_bota(SOURCE, SHEET_BOTA)
except Exception as e:
    st.warning(f"Prévisions BOTA non chargées: {e}")
    dfb = pd.DataFrame()

# ========= SIDEBAR =========
with st.sidebar:
    st.header("Filtres – Côte d’Ivoire")

    uni = safe_concat([df, dfw], cols=["CocoaYearStart", "AnneeCacao"]).dropna().drop_duplicates()
    if not uni.empty and "CocoaYearStart" in uni.columns:
        uni = uni.sort_values("CocoaYearStart")

    if uni.empty:
        y = pd.Timestamp.today().year % 100
        labels_all = [f"{y:02d}/{(y+1)%100:02d}"]
    else:
        labels_all = uni["AnneeCacao"].astype(str).tolist()

    idx_default = max(0, len(labels_all) - 1)
    annee_sel = st.selectbox("Année cacao (référence)", labels_all, index=idx_default)

    years_all = []
    if not dfw.empty and "AnneeCacao" in dfw.columns:
        try:
            years_all = (
                dfw["AnneeCacao"].dropna().astype(str).drop_duplicates()
                .sort_values(key=lambda s: s.map(lambda x: int(str(x).split("/")[0])))
                .tolist()
            )
        except Exception:
            years_all = dfw["AnneeCacao"].dropna().astype(str).drop_duplicates().tolist()

    default_years = [annee_sel]
    if annee_sel in years_all:
        i = years_all.index(annee_sel)
        if i > 0: default_years.append(years_all[i-1])

    years_overlay = st.multiselect(
        "Années à superposer (hebdo multi-années)",
        options=years_all,
        default=list(dict.fromkeys(default_years))
    )

    ports = sorted(df["Port"].dropna().unique()) if not df.empty and "Port" in df.columns else []
    ports_sel = st.multiselect("Ports (journalier)", ports, default=ports)

    freq = st.radio("Vue campagne", ["Hebdomadaire (officiel)", "Mensuelle (calendaire)"], index=0)
    show_cum = st.checkbox("Afficher le cumul (sinon : hebdo/ mensuel)", value=True)

    show_bota = st.checkbox("Afficher prévisions BOTA", value=(not dfb.empty))
    bota_year_options = sorted(dfb["AnneeCacao"].dropna().unique()) if not dfb.empty and "AnneeCacao" in dfb.columns else []
    bota_years_sel = st.multiselect(
        "Années BOTA à afficher",
        bota_year_options,
        default=[annee_sel] if annee_sel in bota_year_options else []
    )

# Sous-ensemble journalier (ports)
fdf = df[df["AnneeCacao"] == annee_sel].copy() if not df.empty else pd.DataFrame()
if ports_sel and not fdf.empty:
    fdf = fdf[fdf["Port"].isin(ports_sel)]

# ---------------------------------------------------------------------
# 1) CUMUL CAMPAGNE
# ---------------------------------------------------------------------
st.header("Cumul campagne")
curw = dfw[dfw["AnneeCacao"] == annee_sel].copy() if not dfw.empty else pd.DataFrame()
official_total = float(curw["Weekly_Stat"].sum()) if not curw.empty else 0.0
ports_actifs = int(fdf["Port"].nunique()) if not fdf.empty else 0
k1, k2 = st.columns(2)
k1.metric("Cumul hebdo (somme Weekly_Stat)", f"{official_total:,.1f} t")
k2.metric("Ports actifs (journalier)", f"{ports_actifs}")
if not curw.empty:
    st.caption(f"Période hebdo couverte : {curw['Date'].min().date()} → {curw['Date'].max().date()}  •  Source : {SHEET_WEEKLY} (col. G)")

if freq.startswith("Hebdo"):
    if dfw.empty:
        st.warning("Feuille hebdo/cumul indisponible.")
    else:
        if (dfw["AnneeCacao"]==annee_sel).any():
            base_year = int(dfw.loc[dfw["AnneeCacao"]==annee_sel, "CocoaYearStart"].iloc[0])
        else:
            base_year = _start_year_from_label(annee_sel)
        base_start = pd.Timestamp(base_year, 10, 1)

        years_sel = years_overlay if years_overlay else [annee_sel]

        def series_for(label: str):
            d = dfw[dfw["AnneeCacao"]==label].sort_values("NumeroSemaine").copy()
            if label == annee_sel and not d.empty:
                mask_valid = (d["Weekly_Stat"].fillna(0) > 0) | (d["Cumul_Stat"].fillna(0) > 0)
                last_date = d.loc[mask_valid, "Date"].max()
                if pd.notna(last_date): d = d[d["Date"] <= last_date]
            ycol = "Cumul_Stat" if show_cum else "Weekly_Stat"
            out = d[["NumeroSemaine", ycol]].rename(columns={ycol: "y"}).copy()
            out["BaseDate"] = base_start + pd.to_timedelta((out["NumeroSemaine"]-1)*7, unit="D")
            return out

        fig1 = go.Figure()
        palette_main = build_palette_long()
        y_max_hint = 0.0

        for i, lab in enumerate(years_sel):
            s = series_for(lab)
            if s.empty: continue
            y_max_hint = max(y_max_hint, float(s["y"].max()))
            fig1.add_trace(go.Scatter(
                x=s["BaseDate"], y=s["y"], mode="lines+markers", name=str(lab),
                line=dict(width=3 if lab == annee_sel else 2),
                marker=dict(size=7),
                line_color=palette_main[i % len(palette_main)]
            ))

        if show_bota and not dfb.empty and bota_years_sel:
            for lab in [y for y in years_sel if y in set(bota_years_sel)]:
                sb = dfb[dfb["AnneeCacao"]==lab].sort_values("NumeroSemaine").copy()
                if sb.empty: continue
                x_ = base_start + pd.to_timedelta((sb["NumeroSemaine"]-1)*7, unit="D")
                y_ = sb["Cumul_Bota"] if show_cum else sb["Weekly_Bota"]
                if y_.notna().any():
                    y_max_hint = max(y_max_hint, float(y_.max()))
                    fig1.add_trace(go.Scatter(
                        x=x_, y=y_, mode="lines+markers",
                        name=f"{lab} (Prévision)",
                        line=dict(width=2, dash="dot", color="#6a51a3"),
                        marker=dict(size=6)
                    ))

        ttl  = "Cumul hebdomadaire" if show_cum else "Tonnage hebdomadaire"
        ylab = "Cumul (tons)" if show_cum else "Tonnage (tons)"
        style_excel_like(fig1,
            title=f"{ttl} – campagnes {', '.join(map(str, years_sel))} (source {SHEET_WEEKLY})",
            base_start=base_start, ylabel=ylab, y_max_hint=y_max_hint)
        st.plotly_chart(fig1, use_container_width=True)

else:
    if fdf.empty:
        st.info("Pas de données mensuelles disponibles.")
    else:
        fdf["Mois"] = fdf["Date"].dt.to_period("M").dt.to_timestamp()
        ts = fdf.groupby("Mois", as_index=False)["Tonnage"].sum().sort_values("Mois")
        if show_cum:
            ts["Cumul"] = ts["Tonnage"].cumsum()
            fig1m = px.line(ts, x="Mois", y="Cumul", markers=True)
            style_fig(fig1m, title=f"Cumul mensuel – Campagne {annee_sel}",
                      xaxis=dict(title="Mois", showgrid=True, gridcolor="#dddddd", tickformat="%m/%Y"),
                      yaxis=dict(title="Tonnage cumulé (t)", showgrid=True, gridcolor="#dddddd", tickformat=",.0f"),
                      bg="white")
        else:
            fig1m = px.line(ts, x="Mois", y="Tonnage", markers=True)
            style_fig(fig1m, title=f"Tonnage mensuel – Campagne {annee_sel}",
                      xaxis=dict(title="Mois", showgrid=True, gridcolor="#dddddd", tickformat="%m/%Y"),
                      yaxis=dict(title="Tonnage (t)", showgrid=True, gridcolor="#dddddd", tickformat=",.0f"),
                      bg="white")
        st.plotly_chart(fig1m, use_container_width=True)

# ---------------------------------------------------------------------
# 2) COMPARAISON PAR SOUS-CAMPAGNE — HISTOGRAMMES + slider semaines
# ---------------------------------------------------------------------
st.header("Comparaison par sous-campagne – Histogrammes (hebdo officiel)")

if dfw.empty:
    st.info("La feuille hebdo/cumul est indisponible.")
else:
    colp, coly, collta = st.columns([1.0, 1.2, 1.2])
    with colp:
        part = st.radio("Sous-campagne", ["MAIN CROP (01/10 → 31/03)", "MID CROP (01/04 → 30/09)"])
        is_main = part.startswith("MAIN")

    order = (dfw[["CocoaYearStart", "AnneeCacao"]]
             .dropna().drop_duplicates()
             .sort_values("CocoaYearStart"))
    labels = order["AnneeCacao"].astype(str).tolist()
    if annee_sel not in labels: labels.append(annee_sel)
    idx_cur = _safe_index(labels, annee_sel)

    with coly:
        default_compare = [labels[idx_cur-1]] if idx_cur > 0 else []
        compare_years = st.multiselect(
            "Comparer à (1..n années)",
            options=[y for y in labels if y != annee_sel],
            default=default_compare
        )
    with collta:
        prev_all = labels[:idx_cur]
        default_lta_season = prev_all[-4:] if prev_all else []
        lta_years_season = st.multiselect(
            "LTA (moyenne) – années",
            options=prev_all, default=default_lta_season
        )

    truncate_to_latest = st.checkbox(
        "Tronquer la campagne courante à la dernière semaine disponible",
        value=True
    )

    def _season_window(label: str, main: bool):
        y0 = _start_year_from_label(label)
        return (pd.Timestamp(y0,10,1), pd.Timestamp(y0+1,3,31,23,59,59)) if main \
               else (pd.Timestamp(y0+1,4,1), pd.Timestamp(y0+1,9,30,23,59,59))

    s0_cur, e0_cur = _season_window(annee_sel, is_main)
    cur_season = dfw[(dfw["AnneeCacao"]==annee_sel) & (dfw["Date"]>=s0_cur) & (dfw["Date"]<=e0_cur)].copy()

    max_week = int(cur_season["NumeroSemaine"].max()) if not cur_season.empty else 1
    if truncate_to_latest and not cur_season.empty:
        mask_valid = (cur_season["Weekly_Stat"].fillna(0) > 0) | (cur_season["Cumul_Stat"].fillna(0) > 0)
        if mask_valid.any():
            max_week = int(cur_season.loc[mask_valid, "NumeroSemaine"].max())
        else:
            max_week = 1

    if max_week < 2:
        st.caption("🔎 Pas assez de semaines pour un intervalle : seule la semaine 1 est disponible.")
        wk_start, wk_end = 1, 1
    else:
        wk_start, wk_end = st.slider(
            "Plage de semaines",
            min_value=1, max_value=max_week,
            value=(1, max_week)
        )

    def window_sum(label: str, main: bool, wk_lo: int, wk_hi: int) -> float:
        s0, e0 = _season_window(label, main)
        d = dfw[(dfw["AnneeCacao"]==label) & (dfw["Date"]>=s0) & (dfw["Date"]<=e0)]
        if d.empty: return 0.0
        d = d[(d["NumeroSemaine"]>=wk_lo) & (d["NumeroSemaine"]<=wk_hi)]
        return float(d["Weekly_Stat"].sum())

    rows = [{"Campagne": annee_sel, "Type": "Courante",
             "Tonnage": window_sum(annee_sel, is_main, wk_start, wk_end)}]
    for lab in compare_years:
        rows.append({"Campagne": lab, "Type": "Historique",
                     "Tonnage": window_sum(lab, is_main, wk_start, wk_end)})

    if lta_years_season:
        vals = [window_sum(lab, is_main, wk_start, wk_end) for lab in lta_years_season]
        if vals:
            rows.append({"Campagne": f"LTA ({lta_years_season[0]}–{lta_years_season[-1]})" if len(lta_years_season)>=2 else f"LTA ({lta_years_season[0]})",
                         "Type": "LTA", "Tonnage": float(pd.Series(vals).mean())})

    df_hist = pd.DataFrame(rows)
    palette = {"Courante": "#b71c1c", "Historique": "#3569a6", "LTA": "#2ca02c"}
    fig_hist = px.bar(df_hist, x="Campagne", y="Tonnage", color="Type",
                      color_discrete_map=palette)

    title_hist = (f"Cumul {'MAIN' if is_main else 'MID'} CROP — Histogrammes<br>"
                  f"<sup>Semaine {wk_start} → {wk_end} ({s0_cur.strftime('%d/%m/%Y')} → {e0_cur.strftime('%d/%m/%Y')})</sup>")
    style_fig(fig_hist, title=title_hist,
              xaxis=dict(title="Campagne", showgrid=False),
              yaxis=dict(title="Tonnage (t)", showgrid=True, gridcolor="#dddddd", tickformat=",.0f"),
              bg="white")

    # >>> LÉGENDE À DROITE (verticale, centrée)
    fig_hist.update_layout(
        legend=dict(
            orientation="v",
            y=0.5, yanchor="middle",
            x=1.02, xanchor="left",
            bgcolor="rgba(255,255,255,0)",
            font=dict(family=BOLD_FONT, size=11)
        )
    )

    fig_hist = add_value_labels(fig_hist, unit=" t")
    st.plotly_chart(fig_hist, use_container_width=True)

# ---------------------------------------------------------------------
# 3) CUMULS HEBDOMADAIRES — MULTI-ANNÉES + LTA (OFFICIEL) + BOTA
# ---------------------------------------------------------------------
st.header("Cumuls hebdomadaires – Multi-années + LTA (officiel)")

if dfw.empty:
    st.warning("Feuille hebdo/cumul indisponible.")
else:
    order = (dfw[["CocoaYearStart","AnneeCacao"]].drop_duplicates()
             .sort_values("CocoaYearStart"))
    labels = list(order["AnneeCacao"].astype(str))
    if annee_sel not in labels: labels.append(annee_sel)
    idx = _safe_index(labels, annee_sel)
    prev_all = labels[:idx]

    default_lta = prev_all[-4:] if len(prev_all) >= 1 else []
    lta_years = st.multiselect("LTA (Long Term Average) – choisir les années",
                               options=prev_all, default=default_lta)

    base_year  = _start_year_from_label(annee_sel)
    base_start = pd.Timestamp(base_year, 10, 1)

    def curve(label: str, limit_to_last=False, last_date=None) -> pd.DataFrame:
        d = dfw[dfw["AnneeCacao"]==label].sort_values("NumeroSemaine").copy()
        if limit_to_last and last_date is not None: d = d[d["Date"] <= last_date]
        out = d[["NumeroSemaine", "Cumul_Stat"]].copy()
        out["BaseDate"] = base_start + pd.to_timedelta((out["NumeroSemaine"]-1)*7, unit="D")
        return out

    cur_full = dfw[dfw["AnneeCacao"]==annee_sel].sort_values("NumeroSemaine")
    mask_valid = (cur_full["Weekly_Stat"].fillna(0) > 0) | (cur_full["Cumul_Stat"].fillna(0) > 0)
    last_valid_date = cur_full.loc[mask_valid, "Date"].max() if not cur_full.empty else None

    cur   = curve(annee_sel, limit_to_last=True, last_date=last_valid_date)
    prev1 = curve(labels[idx-1]) if idx-1 >= 0 else pd.DataFrame()
    prev2 = curve(labels[idx-2]) if idx-2 >= 0 else pd.DataFrame()

    lta_df = pd.DataFrame()
    if lta_years:
        tmp = None
        for lab in lta_years:
            c = curve(lab)[["NumeroSemaine", "Cumul_Stat"]].rename(columns={"Cumul_Stat": lab})
            tmp = c if tmp is None else tmp.merge(c, on="NumeroSemaine", how="outer")
        if tmp is not None:
            tmp = tmp.sort_values("NumeroSemaine").reset_index(drop=True)
            lta_df = pd.DataFrame({
                "NumeroSemaine": tmp["NumeroSemaine"],
                "BaseDate": base_start + pd.to_timedelta((tmp["NumeroSemaine"]-1)*7, unit="D"),
                "Cumul_Stat": tmp.drop(columns=["NumeroSemaine"]).mean(axis=1, skipna=True)
            })

    figc = go.Figure()
    if not cur.empty:
        figc.add_trace(go.Scatter(x=cur["BaseDate"], y=cur["Cumul_Stat"],
                                  mode="lines+markers", name=annee_sel,
                                  line=dict(color="#7f0000", width=3),
                                  marker=dict(symbol="circle", size=7)))
    if not prev1.empty:
        lab1 = labels[idx-1]
        figc.add_trace(go.Scatter(x=prev1["BaseDate"], y=prev1["Cumul_Stat"],
                                  mode="lines+markers", name=lab1,
                                  line=dict(color="#1f77b4", width=2.5),
                                  marker=dict(symbol="triangle-up", size=7)))
    if not prev2.empty:
        lab2 = labels[idx-2]
        figc.add_trace(go.Scatter(x=prev2["BaseDate"], y=prev2["Cumul_Stat"],
                                  mode="lines+markers", name=lab2,
                                  line=dict(color="#ff7f0e", width=2.5),
                                  marker=dict(symbol="square", size=7)))
    if not lta_df.empty:
        lblavg = f"LTA ({lta_years[0]}–{lta_years[-1]})" if len(lta_years) >= 2 else f"LTA ({lta_years[0]})"
        figc.add_trace(go.Scatter(x=lta_df["BaseDate"], y=lta_df["Cumul_Stat"],
                                  mode="lines+markers", name=lblavg,
                                  line=dict(color="#2ca02c", width=2.5, dash="dash"),
                                  marker=dict(symbol="diamond", size=6)))

    y_max_hint = 0.0
    for ddf in (cur, prev1, prev2, lta_df):
        if not ddf.empty: y_max_hint = max(y_max_hint, float(ddf["Cumul_Stat"].max()))

    if show_bota and not dfb.empty and bota_years_sel:
        for lab in bota_years_sel:
            sb = dfb[dfb["AnneeCacao"]==lab].sort_values("NumeroSemaine").copy()
            if sb.empty or sb["Cumul_Bota"].isna().all(): continue
            x_ = base_start + pd.to_timedelta((sb["NumeroSemaine"]-1)*7, unit="D")
            y_ = sb["Cumul_Bota"]
            y_max_hint = max(y_max_hint, float(y_.max()))
            figc.add_trace(go.Scatter(
                x=x_, y=y_, mode="lines+markers", name=f"{lab} (Prévision)",
                line=dict(color="#6a51a3", width=2, dash="dot"), marker=dict(size=6)
            ))

    style_excel_like(figc, title="Côte d'Ivoire – Cumul hebdomadaire (tons)",
                     base_start=base_start, ylabel="Cumul (tons)", y_max_hint=y_max_hint)
    st.plotly_chart(figc, use_container_width=True)

# ---------------------------------------------------------------------
# 4) RÉPARTITIONS — camemberts (ports / jours de la semaine)
# ---------------------------------------------------------------------
st.header("Répartitions — camemberts")

c1, c2 = st.columns(2)

with c1:
    st.subheader("Par port (journalier filtré)")
    if fdf.empty:
        st.info("Aucune donnée pour les ports avec les filtres actuels.")
    else:
        by_port = (fdf.groupby("Port", as_index=False)["Tonnage"].sum()
                     .sort_values("Tonnage", ascending=False))
        fig2 = px.pie(by_port, names="Port", values="Tonnage")
        style_fig(fig2, title="Répartition par port", bg="white")
        fig2.update_traces(textinfo="percent+label",
                           hovertemplate="%{label}: %{value:,.0f} t<br>%{percent}")
        st.plotly_chart(fig2, use_container_width=True)

with c2:
    st.subheader("Moyenne par jour de la semaine (part hebdo)")
    if fdf.empty:
        st.info("Pas de données pour la répartition journalière.")
    else:
        wk = fdf.groupby(["NumeroSemaine","JourSemaine"], as_index=False)["Tonnage"].sum()
        if wk.empty:
            st.info("Pas de données pour la répartition journalière.")
        else:
            wk_tot = wk.groupby("NumeroSemaine", as_index=False)["Tonnage"].sum().rename(columns={"Tonnage":"TotalSemaine"})
            wk = wk.merge(wk_tot, on="NumeroSemaine", how="left")
            wk["Pct"] = wk["Tonnage"] / wk["TotalSemaine"]
            avg = wk.groupby("JourSemaine", as_index=False)["Pct"].mean().sort_values("JourSemaine")
            jour_map = {1:"Lun",2:"Mar",3:"Mer",4:"Jeu",5:"Ven",6:"Sam",7:"Dim"}
            avg["Jour"] = avg["JourSemaine"].map(jour_map)
            figd2 = px.pie(avg, names="Jour", values="Pct")
            style_fig(figd2, title="Répartition journalière (moyenne des parts hebdo)", bg="white")
            figd2.update_traces(textinfo="percent+label",
                                hovertemplate="%{label}: %{percent}")
            st.plotly_chart(figd2, use_container_width=True)
