# lulu.py — Dashboard Côte d'Ivoire (Cacao)
# ------------------------------------------------------------
# 2 vues séparées :
#   - Daily  : like-for-like semaine (WTD) + tableau détaillé + export
#   - Weekly : cumul hebdo multi-années + vs N-1 + projection STAT + sous-campagnes + exports
# ------------------------------------------------------------

from __future__ import annotations

from itertools import cycle
from typing import Optional, Tuple, List

import pandas as pd
import plotly.graph_objects as go
import psycopg2
from psycopg2.extras import DictCursor
import streamlit as st


# ================== CONFIG ==================

st.set_page_config(page_title="CIV – Port Arrivals (Cacao)", layout="wide")

PG_HOST = "localhost"
PG_PORT = 5432
PG_DB = "postgres"
PG_USER = "postgres"
PG_PASSWORD = "Touton0812"

TABLE_DAILY = "cocoa_arrivals1"          # date_arrival, nb_of_day, nb_of_week, cocoa_year, port, arrivals
TABLE_WEEKLY = "cocoa_weekly"            # date_arrival, cocoa_year, nb_of_week, nb_of_month, weekly_stat
TABLE_FORECAST = "cocoa_weekly_forecast" # date_arrival, cocoa_year, nb_of_week, weekly_stat (et parfois cumul_stat)

BOLD_FONT = "Arial Black, Arial, sans-serif"
WEEK_MS = 7 * 24 * 60 * 60 * 1000


# ================== PG HELPERS ==================

@st.cache_resource
def get_pg_conn():
    return psycopg2.connect(
        host=PG_HOST,
        port=PG_PORT,
        dbname=PG_DB,
        user=PG_USER,
        password=PG_PASSWORD,
        cursor_factory=DictCursor,
    )

def sql_to_df(sql: str, params: dict | None = None) -> pd.DataFrame:
    conn = get_pg_conn()
    return pd.read_sql(sql, conn, params=params)

def try_sql_to_df(sql_list: List[str], params: dict | None = None) -> pd.DataFrame:
    last_err = None
    for q in sql_list:
        try:
            return sql_to_df(q, params=params)
        except Exception as e:
            last_err = e
    raise last_err  # type: ignore[misc]


# ================== CORE HELPERS ==================

def _start_year_from_label(label: str) -> int:
    s = str(label).strip()
    if s.lower() in {"nan", "", "none"}:
        return pd.Timestamp.today().year
    try:
        yy = int(s.split("/")[0])
        return 2000 + yy if yy < 50 else 1900 + yy
    except Exception:
        try:
            return int(s[:4])
        except Exception:
            return pd.Timestamp.today().year

def _sort_cocoa_years(labels: List[str]) -> List[str]:
    def keyf(x: str) -> int:
        return _start_year_from_label(x)
    return sorted([str(x) for x in labels if str(x).strip()], key=keyf)

def _previous_campaign(labels_sorted: List[str], current: str) -> Optional[str]:
    if current not in labels_sorted:
        return None
    i = labels_sorted.index(current)
    return labels_sorted[i - 1] if i > 0 else None

def _base_start_for_campaign(label: str) -> pd.Timestamp:
    y0 = _start_year_from_label(label)
    return pd.Timestamp(y0, 10, 1)

def _next_sunday(d: pd.Timestamp) -> pd.Timestamp:
    d = pd.Timestamp(d).normalize()
    off = (6 - d.weekday()) % 7
    return d + pd.Timedelta(days=off)

def weekly_xaxis_on_sundays(anchor: pd.Timestamp) -> dict:
    tick0 = _next_sunday(anchor)
    return dict(title="Date", tickformat="%d/%m/%Y", tickmode="linear", tick0=tick0, dtick=WEEK_MS)


# ================== TABLE STYLING (DERNIERS HELPERS) ==================

def style_table(df: pd.DataFrame, *, delta_cols=None, header_colors=None) -> "pd.io.formats.style.Styler":
    """
    - Entêtes avec couleurs différentes par colonne
    - Toutes les valeurs centrées + en gras
    - Zébrage
    - Colonnes delta (si delta_cols) en vert/rouge
    """
    delta_cols = delta_cols or []

    default_palette = [
        "#0B3954", "#087E8B", "#BFD7EA", "#FFB703", "#FB8500",
        "#6A4C93", "#2A9D8F", "#264653", "#8D99AE", "#D62828"
    ]
    if header_colors is None:
        header_colors = list(default_palette)

    col_color = cycle(header_colors)
    sty = df.style

    table_styles = []
    for i, _ in enumerate(df.columns):
        table_styles.append({
            "selector": f"th.col_heading.level0.col{i}",
            "props": [
                ("background-color", next(col_color)),
                ("color", "white"),
                ("font-weight", "800"),
                ("text-align", "center"),
                ("border", "1px solid #ffffff"),
                ("padding", "8px 10px"),
            ]
        })

    table_styles.append({
        "selector": "th.blank, th.row_heading",
        "props": [
            ("background-color", "#0B3954"),
            ("color", "white"),
            ("font-weight", "800"),
            ("text-align", "center"),
            ("border", "1px solid #ffffff"),
        ]
    })

    table_styles.append({
        "selector": "td",
        "props": [
            ("text-align", "center"),
            ("font-weight", "800"),
            ("padding", "6px 8px"),
            ("border", "1px solid #f0f0f0"),
        ]
    })

    table_styles.append({
        "selector": "tbody tr:nth-child(even) td",
        "props": [("background-color", "#FFF3E0")]
    })
    table_styles.append({
        "selector": "tbody tr:nth-child(odd) td",
        "props": [("background-color", "#FFE0B2")]
    })

    sty = sty.set_table_styles(table_styles)

    def _delta_color(v):
        try:
            x = float(v)
        except Exception:
            return ""
        if x > 0:
            return "color: #1B7F2A; font-weight: 900;"
        if x < 0:
            return "color: #B00020; font-weight: 900;"
        return "color: #333333; font-weight: 900;"

    for c in delta_cols:
        if c in df.columns:
            sty = sty.map(lambda v: _delta_color(v), subset=[c])

    return sty


# ================== LOADERS ==================

@st.cache_data
def load_daily() -> pd.DataFrame:
    """
    cocoa_arrivals1 : PAS de nb_of_month
    colonnes utilisées: date_arrival, nb_of_day, nb_of_week, cocoa_year, port, arrivals
    """
    sql = f"""
        SELECT
            date_arrival,
            nb_of_day,
            nb_of_week,
            cocoa_year,
            port,
            arrivals
        FROM {TABLE_DAILY}
    """
    df = sql_to_df(sql)
    if df.empty:
        return df

    df = df.rename(columns={
        "date_arrival": "Date",
        "nb_of_day": "Day_Number",
        "nb_of_week": "Week_Number",
        "cocoa_year": "CocoaYear",
        "port": "Port",
        "arrivals": "Tonnage",
    })
    df["Date"] = pd.to_datetime(df["Date"])
    df["Tonnage"] = pd.to_numeric(df["Tonnage"], errors="coerce").fillna(0.0)
    df["Week_Number"] = pd.to_numeric(df["Week_Number"], errors="coerce").fillna(0).astype(int)
    df["Day_Number"] = pd.to_numeric(df["Day_Number"], errors="coerce").fillna(0).astype(int)

    df["CocoaYearStart"] = df["CocoaYear"].map(_start_year_from_label)
    return df

@st.cache_data
def load_weekly() -> pd.DataFrame:
    """
    cocoa_weekly : la colonne cumul_stat N'EXISTE PAS => on calcule Cum_From_Weekly
    colonnes: date_arrival, cocoa_year, nb_of_week, nb_of_month, weekly_stat
    """
    sql = f"""
        SELECT
            date_arrival,
            cocoa_year,
            nb_of_week,
            nb_of_month,
            weekly_stat
        FROM {TABLE_WEEKLY}
    """
    dfw = sql_to_df(sql)
    if dfw.empty:
        return dfw

    dfw = dfw.rename(columns={
        "date_arrival": "Date",
        "cocoa_year": "CocoaYear",
        "nb_of_week": "Week_Number",
        "nb_of_month": "Month_Number",
        "weekly_stat": "Weekly_Stat",
    })
    dfw["Date"] = pd.to_datetime(dfw["Date"])
    dfw["Week_Number"] = pd.to_numeric(dfw["Week_Number"], errors="coerce").fillna(0).astype(int)
    dfw["Month_Number"] = pd.to_numeric(dfw["Month_Number"], errors="coerce").fillna(0).astype(int)
    dfw["Weekly_Stat"] = pd.to_numeric(dfw["Weekly_Stat"], errors="coerce").fillna(0.0)

    dfw["CocoaYearStart"] = dfw["CocoaYear"].map(_start_year_from_label)

    # Cum_From_Weekly (par campagne, en incluant bien la semaine 1)
    dfw = dfw.sort_values(["CocoaYear", "Week_Number"])
    dfw["Cum_From_Weekly"] = dfw.groupby("CocoaYear")["Weekly_Stat"].cumsum()

    return dfw

@st.cache_data
def load_forecast_2526() -> pd.DataFrame:
    """
    Prévisions 25/26 : on prend weekly_stat et on calcule Cumul_Forecast.
    (On essaie aussi cumul_stat si dispo, mais on n'en dépend pas.)
    """
    sql_list = [
        f"""
        SELECT
            date_arrival,
            cocoa_year,
            nb_of_week,
            weekly_stat,
            cumul_stat
        FROM {TABLE_FORECAST}
        WHERE cocoa_year='25/26'
        """,
        f"""
        SELECT
            date_arrival,
            cocoa_year,
            nb_of_week,
            weekly_stat
        FROM {TABLE_FORECAST}
        WHERE cocoa_year='25/26'
        """
    ]
    df = try_sql_to_df(sql_list)
    if df.empty:
        return df

    df = df.rename(columns={
        "date_arrival": "Date",
        "cocoa_year": "CocoaYear",
        "nb_of_week": "Week_Number",
        "weekly_stat": "Week_Stat_Forecast",
    })
    df["Date"] = pd.to_datetime(df["Date"])
    df["Week_Number"] = pd.to_numeric(df["Week_Number"], errors="coerce").fillna(0).astype(int)
    df["Week_Stat_Forecast"] = pd.to_numeric(df["Week_Stat_Forecast"], errors="coerce").fillna(0.0)

    df = df.sort_values("Week_Number")
    df["Cumul_Forecast"] = df["Week_Stat_Forecast"].cumsum()  # cumul forecast brut (depuis début campagne)
    df["CocoaYearStart"] = df["CocoaYear"].map(_start_year_from_label)

    return df


# ================== DAILY PAGE ==================

def page_daily(df: pd.DataFrame, annee_sel: str, ports_sel: List[str], years_sorted: List[str]):
    st.title("Daily – Vue hebdomadaire like-for-like")

    # Filtrage campagne + ports
    f = df[df["CocoaYear"] == annee_sel].copy()
    if ports_sel:
        f = f[f["Port"].isin(ports_sel)]

    if f.empty:
        st.warning("Aucune donnée Daily pour cette campagne / ports.")
        return

    # --- Sélection date (petit calendrier) ---
    default_date = f["Date"].max().date()
    chosen_date = st.date_input(
        "Choisir une date (campagne sélectionnée)",
        value=default_date,
        key="daily_date_picker",
    )

    # Trouver la ligne correspondante, sinon la date la plus proche <=
    f_by_date = f[f["Date"].dt.date == chosen_date]
    if f_by_date.empty:
        prev_dates = f[f["Date"].dt.date <= chosen_date]["Date"]
        if prev_dates.empty:
            st.warning("Pas de date disponible <= la date choisie.")
            return
        nearest = prev_dates.max().date()
        st.info(f"Aucune donnée le {chosen_date} — on prend la date la plus proche: {nearest}")
        chosen_date = nearest
        f_by_date = f[f["Date"].dt.date == chosen_date]

    # Week/Day de référence sur cette date
    week_cur = int(f_by_date["Week_Number"].iloc[0])
    day_cur = int(f_by_date["Day_Number"].iloc[0])

    # Campagne précédente
    prev_label = _previous_campaign(years_sorted, annee_sel)

    # règle spéciale like-for-like daily : 25/26 semaine n = 24/25 semaine n+1
    week_shift = 0
    if prev_label == "24/25" and annee_sel == "25/26":
        week_shift = 1

    week_prev = week_cur + week_shift

    # Agrégations semaine courante (par jour de semaine)
    cur_week = f[f["Week_Number"] == week_cur].groupby(["Date", "Day_Number"], as_index=False)["Tonnage"].sum()
    cur_week = cur_week.sort_values("Day_Number")

    prev_week = pd.DataFrame(columns=["Date", "Day_Number", "Tonnage"])
    if prev_label is not None:
        fprev = df[(df["CocoaYear"] == prev_label)].copy()
        if ports_sel:
            fprev = fprev[fprev["Port"].isin(ports_sel)]
        prev_week = (
            fprev[fprev["Week_Number"] == week_prev]
            .groupby(["Date", "Day_Number"], as_index=False)["Tonnage"].sum()
            .sort_values("Day_Number")
        )

    # Construire table (jours 1..day_cur)
    days = list(range(1, max(1, day_cur) + 1))

    def _val_for_day(dfi: pd.DataFrame, d: int) -> Tuple[str, float]:
        sub = dfi[dfi["Day_Number"] == d]
        if sub.empty:
            return ("", 0.0)
        return (pd.to_datetime(sub["Date"].iloc[0]).strftime("%d/%m/%Y"), float(sub["Tonnage"].sum()))

    rows = []
    for d in days:
        date_str, vcur = _val_for_day(cur_week, d)
        _, vprev = _val_for_day(prev_week, d) if prev_label else ("", 0.0)

        rows.append({
            "Date": date_str,
            "Cocoayear": annee_sel,
            "Week_Number": week_cur,
            "Day": d,
            "Current": vcur,
            "Last_Year": vprev,
            "Delta": vcur - vprev,
        })

    df_table = pd.DataFrame(rows)

    # Cumul (WTD)
    cur_wtd = float(df_table["Current"].sum()) if not df_table.empty else 0.0
    prev_wtd = float(df_table["Last_Year"].sum()) if not df_table.empty else 0.0
    delta_wtd = cur_wtd - prev_wtd

    # Projection fin de semaine (profil N-1)
    proj = cur_wtd
    if prev_label is not None:
        # total semaine N-1 (jours 1..7)
        prev_total_week = float(prev_week["Tonnage"].sum()) if not prev_week.empty else 0.0
        prev_wtd_week = float(prev_week[prev_week["Day_Number"].isin(days)]["Tonnage"].sum()) if not prev_week.empty else 0.0
        if prev_wtd_week > 0 and prev_total_week > 0:
            ratio = prev_wtd_week / prev_total_week
            if ratio > 0:
                proj = cur_wtd / ratio

    # Metrics : comparaison uniquement sous la valeur de cette année
    c1, c2 = st.columns([1.2, 1.0])
    with c1:
        st.metric(
            f"Cumul semaine {week_cur} (WTD) – {annee_sel}",
            f"{cur_wtd:,.0f} t",
            delta=f"{delta_wtd:,.0f} t vs {prev_label or 'N/A'} (sem. {week_prev})",
        )
        if prev_label is not None:
            st.caption(f"Référence N-1 : **{prev_label}** semaine **{week_prev}** (WTD) = **{prev_wtd:,.0f} t**")
        else:
            st.caption("Pas de campagne précédente disponible pour comparer.")
    with c2:
        st.metric(
            "Projection fin de semaine (profil N-1)",
            f"{proj:,.0f} t",
        )

    st.markdown("---")

    # Ajouter ligne Cumul au tableau
    cum_row = {
        "Date": "",
        "Cocoayear": "",
        "Week_Number": "",
        "Day": "Cumul",
        "Current": cur_wtd,
        "Last_Year": prev_wtd,
        "Delta": delta_wtd,
    }
    df_table_out = pd.concat([df_table, pd.DataFrame([cum_row])], ignore_index=True)

    # Format affichage (entiers)
    def _fmt_int(x):
        try:
            if x == "" or x is None:
                return ""
            return f"{float(x):,.0f}"
        except Exception:
            return x

    df_show = df_table_out.copy()
    for c in ["Current", "Last_Year", "Delta"]:
        if c in df_show.columns:
            df_show[c] = df_show[c].apply(_fmt_int)

    st.subheader("Détail journalier de la semaine (like-for-like)")
    st.dataframe(
        style_table(df_show, delta_cols=["Delta"]),
        use_container_width=True,
        hide_index=True,
    )

    # Export CSV du tableau daily
    csv_daily = df_table_out.copy()
    csv_daily["Date"] = csv_daily["Date"].astype(str)
    csv_bytes = csv_daily.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ Export CSV (détail semaine like-for-like)",
        data=csv_bytes,
        file_name=f"CIV_daily_likeforlike_{annee_sel}_week{week_cur}.csv",
        mime="text/csv",
        key="dl_daily_table",
    )


# ================== WEEKLY PAGE ==================

def page_weekly(dfw: pd.DataFrame, df_fc: pd.DataFrame, annee_sel: str, years_sorted: List[str]):
    st.title("Weekly – Cumul hebdomadaire & comparaisons Côte d'Ivoire")

    if dfw is None or dfw.empty:
        st.error("La table cocoa_weekly est vide / non chargée.")
        return

    cur = dfw[dfw["CocoaYear"] == annee_sel].copy().sort_values("Week_Number")
    if cur.empty:
        st.warning(f"Aucune donnée hebdo pour {annee_sel}.")
        return

    # dernière semaine dispo (on évite de couper la semaine 1)
    valid = cur[cur["Weekly_Stat"].fillna(0) > 0]
    last_week = int(valid["Week_Number"].max()) if not valid.empty else int(cur["Week_Number"].max())

    cur_last = cur[cur["Week_Number"] <= last_week].copy().sort_values("Week_Number")
    cur_cum = float(cur_last["Weekly_Stat"].cumsum().iloc[-1]) if not cur_last.empty else 0.0

    prev_label = _previous_campaign(years_sorted, annee_sel)
    prev_cum = None
    if prev_label is not None:
        prev = dfw[dfw["CocoaYear"] == prev_label].copy().sort_values("Week_Number")
        prev_last = prev[prev["Week_Number"] <= last_week]
        if not prev_last.empty:
            prev_cum = float(prev_last["Weekly_Stat"].cumsum().iloc[-1])

    # Metric : comparaison uniquement sous la valeur de cette année
    if prev_cum is not None:
        st.metric(
            f"Cumul campagne {annee_sel} (jusqu'à semaine {last_week})",
            f"{cur_cum:,.0f} t",
            delta=f"{(cur_cum - prev_cum):,.0f} t vs {prev_label} (même semaine)",
        )
        st.caption(f"Référence N-1 : **{prev_label}** cumul à semaine **{last_week}** = **{prev_cum:,.0f} t**")
    else:
        st.metric(
            f"Cumul campagne {annee_sel} (jusqu'à semaine {last_week})",
            f"{cur_cum:,.0f} t",
        )
        st.caption("Pas de campagne précédente disponible pour comparer.")

    st.markdown("---")
    st.subheader("Cumul hebdomadaire – multi-années + projection STAT (25/26)")

    # Choix LTA
    prev_all = [y for y in years_sorted if y != annee_sel]
    default_lta = prev_all[-4:] if len(prev_all) >= 1 else []
    lta_years = st.multiselect(
        "LTA (moyenne) – années",
        options=prev_all,
        default=default_lta,
        key="weekly_lta_years",
    )

    base_start = _base_start_for_campaign(annee_sel)

    def curve(label: str) -> pd.DataFrame:
        d = dfw[dfw["CocoaYear"] == label].copy().sort_values("Week_Number")
        if d.empty:
            return d
        d["Cum"] = d["Weekly_Stat"].cumsum()
        d["BaseDate"] = base_start + pd.to_timedelta((d["Week_Number"] - 1) * 7, unit="D")
        return d[["Week_Number", "BaseDate", "Cum"]]

    fig = go.Figure()

    # Courbe courant (jusqu'à last_week)
    ccur = curve(annee_sel)
    ccur = ccur[ccur["Week_Number"] <= last_week]
    if not ccur.empty:
        fig.add_trace(go.Scatter(
            x=ccur["BaseDate"], y=ccur["Cum"],
            mode="lines+markers",
            name=annee_sel
        ))

    # N-1
    if prev_label is not None:
        cprev = curve(prev_label)
        if not cprev.empty:
            fig.add_trace(go.Scatter(
                x=cprev["BaseDate"], y=cprev["Cum"],
                mode="lines",
                name=prev_label
            ))

    # LTA
    if lta_years:
        tmp = None
        for lab in lta_years:
            c = curve(lab)
            if c.empty:
                continue
            c = c[["Week_Number", "Cum"]].rename(columns={"Cum": lab})
            tmp = c if tmp is None else tmp.merge(c, on="Week_Number", how="outer")
        if tmp is not None and not tmp.empty:
            tmp = tmp.sort_values("Week_Number")
            tmp["LTA"] = tmp.drop(columns=["Week_Number"]).mean(axis=1, skipna=True)
            tmp["BaseDate"] = base_start + pd.to_timedelta((tmp["Week_Number"] - 1) * 7, unit="D")
            fig.add_trace(go.Scatter(
                x=tmp["BaseDate"], y=tmp["LTA"],
                mode="lines",
                name=f"LTA ({lta_years[0]}–{lta_years[-1]})" if len(lta_years) >= 2 else f"LTA ({lta_years[0]})",
                line=dict(dash="dash")
            ))

    # Projection STAT : doit commencer APRES le current
    if annee_sel == "25/26" and df_fc is not None and not df_fc.empty:
        next_week = last_week + 1

        fc = df_fc.copy().sort_values("Week_Number")
        fc = fc[fc["Week_Number"] >= next_week]

        if not fc.empty:
            # cumul "projeté" = cumul actuel + cumsum forecast à partir de next_week
            fc["Cum_Proj"] = cur_cum + fc["Week_Stat_Forecast"].cumsum()
            fc["BaseDate"] = base_start + pd.to_timedelta((fc["Week_Number"] - 1) * 7, unit="D")

            fig.add_trace(go.Scatter(
                x=fc["BaseDate"], y=fc["Cum_Proj"],
                mode="lines+markers",
                name="STAT 25/26 (projection)",
                line=dict(dash="dot")
            ))

    fig.update_layout(
        font=dict(family=BOLD_FONT, size=12),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1.0),
        margin=dict(l=60, r=30, t=60, b=60),
        title="<b>Côte d'Ivoire – Cumul hebdomadaire (tons)</b>",
        paper_bgcolor="white",
        plot_bgcolor="white",
    )
    fig.update_xaxes(**weekly_xaxis_on_sundays(base_start), tickangle=45, showgrid=True, gridcolor="#dddddd")
    fig.update_yaxes(title="<b>Cumul (t)</b>", showgrid=True, gridcolor="#dddddd", tickformat=",.0f")
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    # ================== SOUS-CAMPAGNES (GLISSIÈRE SEMAINE 1 -> ...) ==================
    st.subheader("Comparaison sous-campagnes – glissière (hebdo)")

    part = st.radio(
        "Sous-campagne",
        ["MAIN CROP (sem. 1 → 26)", "MID CROP (sem. 27 → fin)"],
        key="weekly_part_radio",
        horizontal=True
    )
    is_main = part.startswith("MAIN")

    # bornes
    max_week_cur = int(cur["Week_Number"].max())
    if is_main:
        wk_min = 1
        wk_max = min(26, max_week_cur)
    else:
        wk_min = 27
        wk_max = max(27, max_week_cur)

    if wk_max < wk_min:
        st.info("Pas de semaines disponibles pour cette sous-campagne.")
        return

    wk_start, wk_end = st.slider(
        "Plage de semaines",
        min_value=wk_min,
        max_value=wk_max,
        value=(wk_min, wk_max),
        key="weekly_week_slider"
    )

    compare_years = st.multiselect(
        "Comparer à",
        options=[y for y in years_sorted if y != annee_sel],
        default=[prev_label] if prev_label else [],
        key="weekly_compare_years"
    )

    def sum_weeks(label: str, w1: int, w2: int) -> float:
        d = dfw[(dfw["CocoaYear"] == label) & (dfw["Week_Number"] >= w1) & (dfw["Week_Number"] <= w2)]
        if d.empty:
            return 0.0
        return float(d["Weekly_Stat"].sum())

    rows = [{"Campagne": annee_sel, "Type": "Courante", "Tonnage": sum_weeks(annee_sel, wk_start, wk_end)}]
    for lab in compare_years:
        rows.append({"Campagne": lab, "Type": "Historique", "Tonnage": sum_weeks(lab, wk_start, wk_end)})

    if lta_years:
        vals = [sum_weeks(lab, wk_start, wk_end) for lab in lta_years]
        if vals:
            rows.append({
                "Campagne": f"LTA ({lta_years[0]}–{lta_years[-1]})" if len(lta_years) >= 2 else f"LTA ({lta_years[0]})",
                "Type": "LTA",
                "Tonnage": float(pd.Series(vals).mean())
            })

    st.dataframe(
        style_table(pd.DataFrame(rows)),
        use_container_width=True,
        hide_index=True
    )

    st.markdown("---")

    # ================== TABLE WEEKLY + FORECAST (à partir semaine suivante) ==================
    if annee_sel == "25/26" and df_fc is not None and not df_fc.empty:
        st.subheader("Table Weekly + Forecast (projection à partir de la semaine suivante)")

        prev_label = _previous_campaign(years_sorted, annee_sel)

        # actual jusqu'à last_week
        actual = cur.copy().sort_values("Week_Number")
        actual_last = actual[actual["Week_Number"] <= last_week]
        cum_actual_last = float(actual_last["Weekly_Stat"].sum()) if not actual_last.empty else 0.0

        next_week = last_week + 1

        # forecast à partir de next_week
        fc = df_fc.copy().sort_values("Week_Number")
        fc = fc[fc["Week_Number"] >= next_week].copy()

        # last_year weekly_stat (même Week_Number) depuis cocoa_weekly N-1
        if prev_label is not None:
            prev = dfw[dfw["CocoaYear"] == prev_label].copy().sort_values("Week_Number")
            prev_byweek = prev.set_index("Week_Number")["Weekly_Stat"].to_dict()

            # cumul N-1 (jusqu'à chaque semaine)
            prev["Cum"] = prev["Weekly_Stat"].cumsum()
            prev_cum_byweek = prev.set_index("Week_Number")["Cum"].to_dict()
        else:
            prev_byweek = {}
            prev_cum_byweek = {}

        # construire table combinée
        # Date / Month_Number : on utilise forecast.Date et son month, ou Month_Number du weekly si dispo
        base_rows = []
        run_fc_cum = 0.0
        for _, r in fc.iterrows():
            wk = int(r["Week_Number"])
            week_fc = float(r["Week_Stat_Forecast"])
            run_fc_cum += week_fc

            last_y = float(prev_byweek.get(wk, 0.0))
            delta_w = week_fc - last_y

            cum_fc = cum_actual_last + run_fc_cum
            cum_ly = float(prev_cum_byweek.get(wk, 0.0))
            delta_c = cum_fc - cum_ly

            base_rows.append({
                "Date": pd.to_datetime(r["Date"]).strftime("%d/%m/%Y"),
                "Cocoayear": "25/26",
                "Week_Number": wk,
                "Month_number": int(pd.to_datetime(r["Date"]).month),
                "Week_Stat_Forecast": week_fc,
                "Last_Year": last_y,
                "Delta_Week": delta_w,
                "Cum_Forecast": cum_fc,
                "Cum_Last_Year": cum_ly,
                "Delta_Cumul": delta_c,
            })

        df_wf = pd.DataFrame(base_rows)

        # format affichage (avec séparateurs)
        def f0(x):
            try:
                return f"{float(x):,.0f}"
            except Exception:
                return x

        df_wf_show = df_wf.copy()
        for c in ["Week_Stat_Forecast", "Last_Year", "Delta_Week", "Cum_Forecast", "Cum_Last_Year", "Delta_Cumul"]:
            if c in df_wf_show.columns:
                df_wf_show[c] = df_wf_show[c].apply(f0)

        st.dataframe(
            style_table(df_wf_show, delta_cols=["Delta_Week", "Delta_Cumul"]),
            use_container_width=True,
            hide_index=True
        )

        csv_wf = df_wf.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ Export CSV (Weekly + Forecast)",
            data=csv_wf,
            file_name=f"CIV_weekly_forecast_from_week{next_week}_2526.csv",
            mime="text/csv",
            key="dl_weekly_forecast_table"
        )

    st.markdown("---")

    # ================== EXPORTS WEEKLY ==================
    st.header("Exports – Weekly")

    all_years_weekly = _sort_cocoa_years(dfw["CocoaYear"].dropna().unique().tolist())
    export_years = st.multiselect(
        "Années à exporter (cocoa_weekly)",
        options=all_years_weekly,
        default=[annee_sel] if annee_sel in all_years_weekly else all_years_weekly[-1:],
        key="weekly_export_years"
    )

    wexp = dfw[dfw["CocoaYear"].isin(export_years)].copy().sort_values(["CocoaYear", "Week_Number"])
    if wexp.empty:
        st.info("Aucune donnée à exporter.")
        return

    wexp_out = wexp[["Date", "CocoaYear", "Week_Number", "Month_Number", "Weekly_Stat", "Cum_From_Weekly"]].copy()
    wexp_show = wexp_out.copy()
    wexp_show["Date"] = pd.to_datetime(wexp_show["Date"]).dt.strftime("%d/%m/%Y")
    for c in ["Weekly_Stat", "Cum_From_Weekly"]:
        wexp_show[c] = wexp_show[c].apply(lambda x: f"{float(x):,.0f}")

    st.dataframe(
        style_table(wexp_show),
        use_container_width=True,
        hide_index=True
    )

    csv_weekly = wexp_out.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ Export CSV (cocoa_weekly)",
        data=csv_weekly,
        file_name=f"CIV_weekly_{'_'.join(export_years)}.csv",
        mime="text/csv",
        key="dl_weekly_export"
    )

    st.markdown("---")

    # ================== EXPORT FORECAST STAT (avec cumul forecast) ==================
    st.header("Prévision STAT 25/26 (avec cumul forecast)")

    if df_fc is None or df_fc.empty:
        st.info("Pas de données forecast (25/26).")
        return

    fc_out = df_fc.copy().sort_values("Week_Number")
    fc_out["Cumul_Forecast"] = fc_out["Week_Stat_Forecast"].cumsum()

    fc_show = fc_out[["Date", "CocoaYear", "Week_Number", "Week_Stat_Forecast", "Cumul_Forecast"]].copy()
    fc_show["Date"] = pd.to_datetime(fc_show["Date"]).dt.strftime("%d/%m/%Y")
    for c in ["Week_Stat_Forecast", "Cumul_Forecast"]:
        fc_show[c] = fc_show[c].apply(lambda x: f"{float(x):,.0f}")

    st.dataframe(
        style_table(fc_show),
        use_container_width=True,
        hide_index=True
    )

    csv_fc = fc_out.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ Export CSV (forecast STAT 25/26)",
        data=csv_fc,
        file_name="CIV_forecast_STAT_2526.csv",
        mime="text/csv",
        key="dl_forecast_export"
    )


# ================== MAIN ==================

def main():
    with st.sidebar:
        st.header("Données & Vue")

        if st.button("🔄 Actualiser les données", use_container_width=True):
            load_daily.clear()
            load_weekly.clear()
            load_forecast_2526.clear()
            st.rerun()

        view = st.radio("Vue", ["Daily", "Weekly"], index=0, key="view_radio")

    # Load data
    try:
        df_daily = load_daily()
    except Exception as e:
        st.error(f"Erreur chargement journalier (PostgreSQL): {e}")
        return

    try:
        df_weekly = load_weekly()
    except Exception as e:
        st.warning(f"Hebdo non chargé (PostgreSQL): {e}")
        df_weekly = pd.DataFrame()

    try:
        df_fc = load_forecast_2526()
    except Exception as e:
        st.warning(f"Forecast non chargé (PostgreSQL): {e}")
        df_fc = pd.DataFrame()

    # Sidebar filters
    with st.sidebar:
        st.header("Filtres – Côte d’Ivoire")

        years = []
        if df_daily is not None and not df_daily.empty:
            years += df_daily["CocoaYear"].dropna().unique().tolist()
        if df_weekly is not None and not df_weekly.empty:
            years += df_weekly["CocoaYear"].dropna().unique().tolist()

        years_sorted = _sort_cocoa_years(list(set(years)))
        if not years_sorted:
            st.error("Aucune année cacao trouvée dans les données.")
            return

        annee_sel = st.selectbox("Année cacao (référence)", years_sorted, index=len(years_sorted) - 1, key="year_select")

        ports = sorted(df_daily["Port"].dropna().unique().tolist()) if df_daily is not None and not df_daily.empty else []
        ports_sel = st.multiselect("Ports (journalier)", ports, default=ports, key="ports_select")

    # Routing
    if view == "Daily":
        page_daily(df_daily, annee_sel, ports_sel, years_sorted)
    else:
        page_weekly(df_weekly, df_fc, annee_sel, years_sorted)


if __name__ == "__main__":
    main()
