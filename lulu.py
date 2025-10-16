# lulu.py — Dashboard Côte d'Ivoire (Cacao) + Prévisions BOTA
# ------------------------------------------------------------
# 1) Cumul campagne (hebdo/mensuel) + superposition multi-années + BOTA
# 2) Comparaison Main/Mid – HISTOGRAMMES UNIQUEMENT (plage de semaines)
# 3) Cumuls hebdomadaires – Multi-années + LTA (officiel) + BOTA
# 4) Répartitions (ports & jours de semaine) + export
# Style : fond blanc, grille grise, cadre noir ; texte en gras
# ------------------------------------------------------------

import re
from pathlib import Path
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="CIV – Port Arrivals (Cacao)", layout="wide")

# ========= PALETTE =========
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

# ========= STYLES =========
BOLD_FONT = "Arial Black, Arial, sans-serif"
WEEK_MS = 7 * 24 * 60 * 60 * 1000  # 7 jours en ms pour dtick Plotly

def _next_sunday(d: pd.Timestamp) -> pd.Timestamp:
    d = pd.Timestamp(d).normalize()
    off = (6 - d.weekday()) % 7  # Mon=0..Sun=6
    return d + pd.Timedelta(days=off)

def weekly_xaxis_on_sundays(anchor: pd.Timestamp) -> dict:
    tick0 = _next_sunday(anchor)
    return dict(title="Date", tickformat="%d/%m/%Y", tickmode="linear", tick0=tick0, dtick=WEEK_MS)

def style_fig(fig, *, title=None, xaxis=None, yaxis=None, bg="transparent"):
    paper = "rgba(0,0,0,0)" if bg=="transparent" else "white"
    fig.update_layout(
        paper_bgcolor=paper, plot_bgcolor=paper,
        font=dict(family=BOLD_FONT, size=12),
        legend=dict(font=dict(family=BOLD_FONT, size=11))
    )
    if title is not None:
        fig.update_layout(title=f"<b>{title}</b>", title_font=dict(family=BOLD_FONT, size=14))
    if xaxis is not None:
        xa = dict(xaxis)
        xa["title"] = f"<b>{xa.get('title','')}</b>"
        xa["tickfont"]  = dict(family=BOLD_FONT, size=11)
        xa["title_font"] = dict(family=BOLD_FONT, size=12)
        fig.update_xaxes(**xa)
    if yaxis is not None:
        ya = dict(yaxis)
        ya["title"] = f"<b>{ya.get('title','')}</b>"
        ya["tickfont"]  = dict(family=BOLD_FONT, size=11)
        ya["title_font"] = dict(family=BOLD_FONT, size=12)
        fig.update_yaxes(**ya)
    return fig

def style_excel_like(fig, *, title, base_start, ylabel, y_max_hint=None):
    """Fond blanc, grille, cadre; ticks les dimanches. Autorange Y si y_max_hint<=0."""
    xax = weekly_xaxis_on_sundays(base_start)
    xax["tickangle"] = 45
    xax.pop("showgrid", None); xax.pop("gridcolor", None)

    fig.update_layout(
        paper_bgcolor="white", plot_bgcolor="white",
        title=f"<b>{title}</b>", title_font=dict(family=BOLD_FONT, size=16),
        font=dict(family=BOLD_FONT, size=12),
        legend=dict(orientation="v", y=1, yanchor="top", x=1.02, xanchor="left",
                    bgcolor="rgba(255,255,255,0)", font=dict(family=BOLD_FONT, size=11)),
        margin=dict(l=70, r=160, t=70, b=70),
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

    # cadre noir
    fig.add_shape(type="rect", xref="paper", yref="paper", x0=0, y0=0, x1=1, y1=1,
                  line=dict(color="black", width=1))
    return fig

# ========= FICHIERS =========
DATA_FILE = Path(r"C:\Users\s.soro\Touton SA Dropbox\STATISTIQUES\ResearchFiles\Fiches_Pays.xlsm")
SHEET_DAILY  = "CIV_Arrivals_Ports_BDD"
SHEET_WEEKLY = "CIV_Arrivals_BDD"
SHEET_BOTA   = "CIV_Bota_Arrivals_Treatments"  # Prévisions hebdo

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
    if s.lower() in {"nan", "", "none"}:
        return pd.Timestamp.today().year
    try:
        yy = int(s.split("/")[0])
        return 2000 + yy if yy < 50 else 1900 + yy
    except Exception:
        try:
            y = int(s[:4]); return y
        except Exception:
            return pd.Timestamp.today().year

def _safe_index(lst, value):
    try:
        return lst.index(value)
    except ValueError:
        return len(lst) - 1 if lst else 0

# ========= LOADERS =========
@st.cache_data
def load_daily_ports(path: Path, sheet: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl", dtype={"Tonnage": object})
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
def load_weekly_cumul(path: Path, sheet: str) -> pd.DataFrame:
    dfw = pd.read_excel(path, sheet_name=sheet, engine="openpyxl",
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
def load_bota(path: Path, sheet: str) -> pd.DataFrame:
    """
    Feuille 'CIV_Bota_Arrivals_Treatments' = blocs de 8 colonnes:
    Date | cocoayear | Week_number | Month_number | Weekly_Stat | Cumul_Stat | Weekly_Bota | Cumul_Bota
    BOTA en 'kt' -> conversion en tonnes (×1000).
    """
    raw = pd.read_excel(path, sheet_name=sheet, engine="openpyxl", header=0)
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

        # kt -> t
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
try:
    df = load_daily_ports(DATA_FILE, SHEET_DAILY)
except Exception as e:
    st.error(f"Erreur chargement journalier: {e}")
    st.stop()

try:
    dfw = load_weekly_cumul(DATA_FILE, SHEET_WEEKLY)
except Exception as e:
    st.error(f"Erreur chargement hebdo/cumul: {e}")
    dfw = pd.DataFrame()

try:
    dfb = load_bota(DATA_FILE, SHEET_BOTA)
except Exception as e:
    st.warning(f"Prévisions BOTA non chargées: {e}")
    dfb = pd.DataFrame()

# ========= SIDEBAR =========
with st.sidebar:
    st.header("Filtres – Côte d’Ivoire")

    # Union d'années (journalier + hebdo) pour être robuste
    uni = pd.concat(
        [
            df[["CocoaYearStart", "AnneeCacao"]],
            dfw[["CocoaYearStart", "AnneeCacao"]] if "AnneeCacao" in dfw.columns else pd.DataFrame(columns=["CocoaYearStart","AnneeCacao"])
        ],
        ignore_index=True
    ).dropna().drop_duplicates().sort_values("CocoaYearStart")

    labels_all = uni["AnneeCacao"].tolist()
    idx_default = len(labels_all)-1 if labels_all else 0
    annee_sel = st.selectbox("Année cacao (référence)", labels_all, index=idx_default)

    # Années à superposer dans la vue campagne (section 1)
    years_all = []
    if not dfw.empty and "AnneeCacao" in dfw.columns:
        years_all = (dfw["AnneeCacao"].dropna().drop_duplicates()
                       .sort_values(key=lambda s: s.map(lambda x: int(str(x).split("/")[0]))).tolist())
    default_years = [annee_sel] + ([years_all[years_all.index(annee_sel)-1]] if annee_sel in years_all and years_all.index(annee_sel)>0 else [])
    years_overlay = st.multiselect("Années à superposer (hebdo multi-années)", options=years_all, default=default_years)

    ports = sorted(df["Port"].dropna().unique())
    ports_sel = st.multiselect("Ports (journalier)", ports, default=ports)

    freq = st.radio("Vue campagne", ["Hebdomadaire (officiel)", "Mensuelle (calendaire)"], index=0)
    show_cum = st.checkbox("Afficher le cumul (sinon : hebdo/ mensuel)", value=True)

    # BOTA
    show_bota = st.checkbox("Afficher prévisions BOTA", value=(not dfb.empty))
    bota_year_options = sorted(dfb["AnneeCacao"].dropna().unique()) if not dfb.empty else []
    bota_years_sel = st.multiselect("Années BOTA à afficher", bota_year_options,
                                    default=[annee_sel] if annee_sel in bota_year_options else [])

# Sous-ensemble journalier (ports)
fdf = df[df["AnneeCacao"] == annee_sel].copy()
if ports_sel:
    fdf = fdf[fdf["Port"].isin(ports_sel)]

# ---------------------------------------------------------------------
# 1) CUMUL CAMPAGNE
# ---------------------------------------------------------------------
st.header("Cumul campagne")
curw = dfw[dfw["AnneeCacao"] == annee_sel].copy()
official_total = float(curw["Weekly_Stat"].sum()) if not curw.empty else 0.0
ports_actifs = int(fdf["Port"].nunique())
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

        # Overlay BOTA
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
# 2) COMPARAISON PAR SOUS-CAMPAGNE (MAIN / MID) — Histogrammes uniquement
#     Sélection par plage de semaines (slider), comparaisons, LTA
# ---------------------------------------------------------------------
st.header("Comparaison par sous-campagne (Main / Mid) – Histogrammes")

if dfw.empty:
    st.info("La feuille hebdo/cumul est indisponible.")
else:
    # --- UI (partie, années, LTA)
    colp, coly, collta = st.columns([1.0, 1.4, 1.4])
    with colp:
        part = st.radio("Sous-campagne", ["MAIN CROP (01/10 → 31/03)", "MID CROP (01/04 → 30/09)"])
        is_main = part.startswith("MAIN")

    order = (dfw[["CocoaYearStart", "AnneeCacao"]]
             .dropna().drop_duplicates()
             .sort_values("CocoaYearStart"))
    labels = order["AnneeCacao"].tolist()
    if annee_sel not in labels:
        labels.append(annee_sel)
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
            options=prev_all,
            default=default_lta_season
        )

    truncate_to_latest = st.checkbox(
        "Tronquer la campagne courante à la dernière semaine disponible",
        value=True
    )

    # Saison windows & helpers
    def _season_window(label: str, main: bool):
        y0 = _start_year_from_label(label)
        return (pd.Timestamp(y0,10,1), pd.Timestamp(y0+1,3,31,23,59,59)) if main \
               else (pd.Timestamp(y0+1,4,1), pd.Timestamp(y0+1,9,30,23,59,59))

    def _last_valid_week(label: str, main: bool) -> int:
        s0, e0 = _season_window(label, main)
        d = dfw[(dfw["AnneeCacao"]==label) & (dfw["Date"]>=s0) & (dfw["Date"]<=e0)].copy()
        if d.empty: return 1
        mask = (d["Weekly_Stat"].fillna(0) > 0) | (d["Cumul_Stat"].fillna(0) > 0)
        last_d = d.loc[mask, "Date"].max()
        if pd.isna(last_d): last_d = d["Date"].max()
        return int(((last_d.normalize() - s0.normalize()).days // 7) + 1)

    # Nombre de semaines théoriques
    s0_cur, e0_cur = _season_window(annee_sel, is_main)
    total_weeks = int(((e0_cur.normalize() - s0_cur.normalize()).days // 7) + 1)

    # Max slider quand on tronque
    w_last_cur = _last_valid_week(annee_sel, is_main)
    w_max_slider = w_last_cur if truncate_to_latest else total_weeks

    # Slider de semaines
    w0, w1 = st.slider(
        "Plage de semaines",
        min_value=1, max_value=total_weeks,
        value=(1, max(1, min(w_max_slider, total_weeks))),
        step=1
    )
    s_date = (s0_cur + pd.to_timedelta((w0-1)*7, unit="D")).date()
    e_date = (s0_cur + pd.to_timedelta((w1)*7 - 1, unit="D")).date()

    # Agrégations
    def _sum_weeks(label: str, main: bool, wk0: int, wk1: int) -> float:
        s0, e0 = _season_window(label, main)
        if label == annee_sel and truncate_to_latest:
            wk1 = min(wk1, _last_valid_week(label, main))
            if wk1 < wk0: wk1 = wk0
        start = s0 + pd.to_timedelta((wk0-1)*7, unit="D")
        end   = s0 + pd.to_timedelta((wk1)*7 - 1, unit="D")
        d = dfw[(dfw["AnneeCacao"]==label) & (dfw["Date"]>=start) & (dfw["Date"]<=end)]
        return float(d["Weekly_Stat"].apply(_to_float).sum()) if not d.empty else 0.0

    rows = []
    rows.append({"Campagne": annee_sel, "Type": "Courante",
                 "Tonnage": _sum_weeks(annee_sel, is_main, w0, w1)})
    for lab in compare_years:
        rows.append({"Campagne": lab, "Type": "Historique",
                     "Tonnage": _sum_weeks(lab, is_main, w0, w1)})

    if lta_years_season:
        vals = [_sum_weeks(lab, is_main, w0, w1) for lab in lta_years_season]
        if len(vals) > 0:
            rows.append({
                "Campagne": (f"LTA ({lta_years_season[0]}–{lta_years_season[-1]})"
                             if len(lta_years_season) > 1
                             else f"LTA ({lta_years_season[0]})"),
                "Type": "LTA",
                "Tonnage": float(pd.Series(vals).mean())
            })

    bar_df = pd.DataFrame(rows)

    if bar_df.empty:
        st.info("Aucune donnée pour la configuration choisie.")
    else:
        cmap = {"Courante":"#b22222","Historique":"#4e79a7","LTA":"#2ca02c"}
        fig_bar = px.bar(bar_df, x="Campagne", y="Tonnage", color="Type",
                         color_discrete_map=cmap, text=None)
        fig_bar.update_traces(hovertemplate="%{x} — %{y:,.0f} t")
        fig_bar.update_layout(
            paper_bgcolor="white", plot_bgcolor="white",
            title=(f"<b>Cumul {'MAIN' if is_main else 'MID'} CROP — Histogrammes</b>"
                   f"<br><sup>Semaine {w0} → {w1}  "
                   f"({s_date.strftime('%d/%m/%Y')} → {e_date.strftime('%d/%m/%Y')})</sup>"),
            font=dict(family=BOLD_FONT, size=12),
            legend=dict(orientation="v", y=1, yanchor="top",
                        x=1.02, xanchor="left", bgcolor="rgba(255,255,255,0)",
                        font=dict(family=BOLD_FONT, size=11)),
            margin=dict(l=70, r=160, t=70, b=70),
            xaxis=dict(title="<b>Campagne</b>"),
            yaxis=dict(title="<b>Tonnage (t)</b>", tickformat=",.0f",
                       showgrid=True, gridcolor="#cccccc"),
        )
        fig_bar.add_shape(type="rect", xref="paper", yref="paper",
                          x0=0, y0=0, x1=1, y1=1, line=dict(color="black", width=1))
        st.plotly_chart(fig_bar, use_container_width=True)

# ---------------------------------------------------------------------
# 3) CUMULS HEBDOMADAIRES — MULTI-ANNÉES + LTA (OFFICIEL) + BOTA
# ---------------------------------------------------------------------
st.header("Cumuls hebdomadaires – Multi-années + LTA (officiel)")

if dfw.empty:
    st.warning("Feuille hebdo/cumul indisponible.")
else:
    order = (dfw[["CocoaYearStart","AnneeCacao"]].drop_duplicates()
             .sort_values("CocoaYearStart"))
    labels = list(order["AnneeCacao"])
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

    # Overlay BOTA (campagne(s) choisie(s))
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

    # (IMPORTANT) — Aucune grosse annotation au milieu du graphe
    st.plotly_chart(figc, use_container_width=True)

# ---------------------------------------------------------------------
# 4) AUTRES SECTIONS : Répartition & Comparaison journalière + Export
# ---------------------------------------------------------------------

# Répartition par port
st.subheader("🧭 Répartition par port (journalier filtré)")
by_port = (fdf.groupby("Port", as_index=False)["Tonnage"].sum()
             .sort_values("Tonnage", ascending=False))
if by_port.empty:
    st.info("Aucune donnée pour les ports avec les filtres actuels.")
else:
    fig2 = px.pie(by_port, names="Port", values="Tonnage")
    style_fig(fig2, title="Répartition par port", xaxis=dict(title=""), yaxis=dict(title=""), bg="white")
    fig2.update_traces(textinfo="percent+label",
                       hovertemplate="%{label}: %{value:,.0f} t<br>%{percent}")
    st.plotly_chart(fig2, use_container_width=True)

# Répartition journalière (somme / moyenne hebdo)
st.subheader("🥧 Répartition journalière (jours de la semaine)")
tabs = st.tabs(["Part sur l’ensemble (somme)", "Moyenne par semaine"])
with tabs[0]:
    sum_dow = (fdf.groupby("JourSemaine", as_index=False)["Tonnage"].sum()
                 .rename(columns={"Tonnage": "Total"}))
    if sum_dow.empty:
        st.info("Pas de données pour la répartition journalière (somme).")
    else:
        jour_map = {1:"Lun",2:"Mar",3:"Mer",4:"Jeu",5:"Ven",6:"Sam",7:"Dim"}
        sum_dow["Jour"] = sum_dow["JourSemaine"].map(jour_map)
        figd1 = px.pie(sum_dow, names="Jour", values="Total")
        style_fig(figd1, title="Répartition journalière (somme)", bg="white")
        figd1.update_traces(textinfo="percent+label",
                            hovertemplate="%{label}: %{value:,.0f} t<br>%{percent}")
        st.plotly_chart(figd1, use_container_width=True)
with tabs[1]:
    wk = fdf.groupby(["NumeroSemaine","JourSemaine"], as_index=False)["Tonnage"].sum()
    if wk.empty:
        st.info("Pas de données pour la répartition journalière (moyenne hebdo).")
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
                            hovertemplate="%{label}: %{percent} (part moyenne)")
        st.plotly_chart(figd2, use_container_width=True)

# Comparaison journalière (YoY / années passées)
st.subheader("📅 Comparaison journalière (YoY / années passées)")
df_ports = df[df["Port"].isin(ports_sel)] if ports_sel else df
date_defaut = fdf["Date"].max().date() if not fdf.empty else (df_ports["Date"].max().date() if not df_ports.empty else pd.Timestamp.today().date())
c_date = st.date_input("Date (dans la campagne sélectionnée)", value=date_defaut, key="cmp_day")
match_mode = st.radio("Correspondance", ["Calendrier (même date - N ans)", "Campagne (même jour de campagne)"], horizontal=True)
camp_order = (df_ports[["CocoaYearStart","AnneeCacao"]].drop_duplicates().sort_values("CocoaYearStart"))
labels_all_cmp = list(camp_order["AnneeCacao"])
idx_cur_cmp = _safe_index(labels_all_cmp, annee_sel)
labels_prev = labels_all_cmp[max(0, idx_cur_cmp-4):idx_cur_cmp]
compare_sel = st.multiselect("Comparer à", labels_prev[::-1], default=labels_prev[-1:] if labels_prev else [], key="cmp_day_sel")

d_cur = pd.Timestamp(c_date)
if (df_ports["AnneeCacao"]==annee_sel).any():
    start_cur = int(df_ports.loc[df_ports["AnneeCacao"]==annee_sel, "CocoaYearStart"].iloc[0])
else:
    start_cur = _start_year_from_label(annee_sel)
jour_cacao_sel = int((d_cur.normalize() - pd.Timestamp(start_cur,10,1)).days) + 1
sum_cur = df_ports[(df_ports["AnneeCacao"]==annee_sel) & (df_ports["Date"].dt.date==c_date)]["Tonnage"].sum()

rows_prev = []
for lab in compare_sel:
    start_prev = int(camp_order.loc[camp_order["AnneeCacao"]==lab, "CocoaYearStart"].iloc[0])
    comp_date = (d_cur - pd.DateOffset(years=int(annee_sel.split("/")[0]) - int(lab.split("/")[0]))) if match_mode.startswith("Calendrier") \
                else (pd.Timestamp(start_prev,10,1) + pd.Timedelta(days=jour_cacao_sel-1))
    s = df_ports[(df_ports["AnneeCacao"]==lab) & (df_ports["Date"]==comp_date)]["Tonnage"].sum()
    rows_prev.append({"Campagne": lab, "Date": comp_date.date(), "Tonnage": s})

row_cur = {"Campagne": annee_sel, "Date": c_date, "Tonnage": float(sum_cur)}
res_prev = pd.DataFrame(rows_prev)
df_all = pd.concat([pd.DataFrame([row_cur]), res_prev], ignore_index=True)

cda, cdb = st.columns([1,1])
with cda: st.metric("Tonnage du jour (campagne courante)", f"{sum_cur:,.1f} t")
with cdb:
    if not df_all.empty: st.dataframe(df_all.sort_values("Campagne"), use_container_width=True)

if not df_all.empty:
    palette_long = build_palette_long()
    fig_daily = px.bar(df_all, x="Campagne", y="Tonnage", color="Campagne",
                       labels={"Tonnage":"t"}, color_discrete_sequence=palette_long)
    style_fig(fig_daily, title=f"Comparaison journalière – {pd.Timestamp(c_date).strftime('%d/%m/%Y')}",
              xaxis=dict(title="Campagne", showgrid=False),
              yaxis=dict(title="Tonnage (t)", showgrid=True, gridcolor="#dddddd", tickformat=",.1f"),
              bg="white")
    fig_daily.update_traces(hovertemplate="%{x}: %{y:,.1f} t")
    for tr in fig_daily.data:
        if tr.name == annee_sel:
            tr.update(marker=dict(line=dict(width=2.5, color="black")))
    st.plotly_chart(fig_daily, use_container_width=True)

# Détail / export (journalier filtré)
with st.expander("Voir le détail / Exporter (journalier filtré)"):
    cols = ["Date","AnneeCacao","JourCacao","NumeroSemaine","NumeroJour","Port","Tonnage","Semaine","Key"]
    cols = [c for c in cols if c in fdf.columns]
    st.dataframe(fdf.sort_values("Date", ascending=False)[cols], use_container_width=True)
    csv = fdf[cols].to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ Export CSV", csv, file_name=f"CIV_daily_{annee_sel}.csv", mime="text/csv")
