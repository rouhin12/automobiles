"""
Vahan Vehicle Registration Dashboard — Institutional / Broking Client Edition.
Supports all sheets: Fuel, Maker, Norms, State, Vehicle Category, Vehicle Class.
Run: streamlit run dashboard.py
"""
import os
import re
import io
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from dashboard_config import (
    get_maker_to_category,
    get_maker_category_for_key,
    normalize_key,
)

# Plotly config for institutional use: enable download as PNG, clean toolbar
PLOTLY_CONFIG = {
    "displayModeBar": True,
    "displaylogo": False,
    "modeBarButtonsToRemove": ["lasso2d", "select2d"],
    "toImageButtonOptions": {"format": "png", "filename": "vahan_chart", "scale": 2},
}

DOWNLOAD_DIR = os.path.join(os.path.dirname(__file__), "downloads")
MASTER_PATH = os.path.join(DOWNLOAD_DIR, "master_sheet.xlsx")

# Month-year column pattern: "Jan 2018", "Feb 2018", ...
MONTH_YEAR_PATTERN = re.compile(r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})$", re.I)
# Total yearly: "Total 2018" or source 13th column "Col_12 2018" / "Col_13 2018" (normalized to "Total YYYY")
TOTAL_YEAR_PATTERN = re.compile(r"^(Total|Col_\d+)\s+(\d{4})$", re.I)


def normalize_total_column_names(df):
    """Rename Col_N YYYY columns to Total YYYY so they display consistently."""
    rename = {}
    for c in df.columns:
        m = TOTAL_YEAR_PATTERN.match(str(c).strip())
        if m and m.group(1).upper().startswith("COL_"):
            rename[c] = f"Total {m.group(2)}"
    return df.rename(columns=rename) if rename else df

# Human-readable labels for each sheet's key column and title
SHEET_LABELS = {
    "Fuel": {"key_label": "Fuel Type", "title": "Fuel Wise"},
    "Maker": {"key_label": "Maker", "title": "Maker Wise"},
    "Norms": {"key_label": "Norm", "title": "Norms Wise"},
    "State": {"key_label": "State", "title": "Statewise"},
    "Vehicle Category": {"key_label": "Vehicle Category", "title": "Vehicle Category Wise"},
    "Vehicle Class": {"key_label": "Vehicle Class", "title": "Vehicle Class Wise"},
}


def load_master_sheet():
    if not os.path.isfile(MASTER_PATH):
        return None
    return pd.ExcelFile(MASTER_PATH)


def get_month_columns(df):
    """Return list of column names that match 'Month Year' (e.g. Jan 2018)."""
    return [c for c in df.columns if MONTH_YEAR_PATTERN.match(str(c).strip())]


def get_total_year_columns(df):
    """Return list of 'Total YYYY' or 'Col_N YYYY' columns, sorted by year."""
    total_cols = [c for c in df.columns if TOTAL_YEAR_PATTERN.match(str(c).strip())]
    years = []
    for c in total_cols:
        m = TOTAL_YEAR_PATTERN.match(str(c).strip())
        if m:
            years.append((int(m.group(2)), c))  # group(2) is the year
    return [c for _, c in sorted(years)]


def parse_month_year(col):
    """Return (year, month_num) for sorting. col e.g. 'Jan 2018'."""
    m = MONTH_YEAR_PATTERN.match(str(col).strip())
    if not m:
        return None
    month_str, year_str = m.group(1), m.group(2)
    months = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()
    try:
        month_num = months.index(month_str[:3].capitalize()) + 1
        return (int(year_str), month_num)
    except (ValueError, IndexError):
        return None


def sort_month_columns(cols):
    """Sort column names by (year, month)."""
    return sorted(cols, key=lambda c: parse_month_year(c) or (0, 0))


def ensure_numeric(df, cols):
    """Convert cols to numeric (strip commas first). Only 1st row (header) and 1st column (Key) are non-numeric."""
    if not cols:
        return df
    df = df.copy()
    def parse_numeric_series(s):
        return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False), errors="coerce")
    df[cols] = df[cols].apply(parse_numeric_series).fillna(0).astype("int64")
    return df


def aggregate_by_maker_classification(df, key_col, maker_to_cat, month_cols_sorted):
    """Aggregate numeric month columns by maker category. Returns df with index = category."""
    df = df.copy()
    df["_category"] = df[key_col].apply(lambda k: get_maker_category_for_key(k, maker_to_cat))
    classified = df[df["_category"].notna()].copy()
    if classified.empty:
        return pd.DataFrame()
    classified[month_cols_sorted] = classified[month_cols_sorted].apply(pd.to_numeric, errors="coerce").fillna(0).astype("int64")
    agg = classified.groupby("_category", as_index=True)[month_cols_sorted].sum()
    agg.index.name = "Category"
    return agg


def df_to_time_series(df, month_cols_sorted, key_label="Key"):
    """Convert wide df to long: Date, Value, and key column for legend."""
    if df.empty or not month_cols_sorted:
        return pd.DataFrame()
    index_name = df.index.name or key_label
    out = df[month_cols_sorted].stack().reset_index()
    out.columns = [index_name, "Month_Year", "Registrations"]
    def to_date(s):
        p = parse_month_year(s)
        if p:
            y, m = p
            return pd.Timestamp(year=y, month=m, day=1)
        return pd.NaT
    out["Date"] = out["Month_Year"].apply(to_date)
    out["Registrations"] = pd.to_numeric(out["Registrations"], errors="coerce").fillna(0).astype("int64")
    return out


def load_sheet_data(sheet_name):
    """Load one sheet, normalize columns, return (df, key_col, month_cols_sorted, total_cols, key_label)."""
    if not os.path.isfile(MASTER_PATH):
        return None, None, [], [], None
    df = pd.read_excel(MASTER_PATH, sheet_name=sheet_name)
    key_col = df.columns[0]
    key_label = SHEET_LABELS.get(sheet_name, {}).get("key_label", key_col)
    df = df.rename(columns={key_col: key_label})
    key_col = key_label
    df = normalize_total_column_names(df)
    month_cols = get_month_columns(df)
    total_cols = get_total_year_columns(df)
    numeric_cols = [c for c in df.columns if c != key_col]
    df = ensure_numeric(df, numeric_cols)
    month_cols_sorted = sort_month_columns(month_cols) if month_cols else []
    return df, key_col, month_cols_sorted, total_cols, key_label


def get_available_years(month_cols_sorted, total_cols):
    """Return sorted list of years present in data."""
    years = set()
    for c in month_cols_sorted or []:
        p = parse_month_year(c)
        if p:
            years.add(p[0])
    for c in total_cols or []:
        m = TOTAL_YEAR_PATTERN.match(str(c).strip())
        if m:
            years.add(int(m.group(2)))
    return sorted(years) if years else []


def compute_yoy_growth(plot_df, month_cols_sorted, this_year):
    """YoY % = (This year month - Last year same month) / Last year same month * 100. Returns long df: Month_Year, Registrations_This, Registrations_Last, YoY_Pct."""
    last_year = this_year - 1
    months_short = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()
    rows = []
    for m_idx, month_name in enumerate(months_short, 1):
        col_this = f"{month_name} {this_year}"
        col_last = f"{month_name} {last_year}"
        if col_this not in plot_df.columns or col_last not in plot_df.columns:
            continue
        total_this = plot_df[col_this].sum()
        total_last = plot_df[col_last].sum()
        if total_last and total_last > 0:
            yoy_pct = round((total_this - total_last) / total_last * 100, 1)
        else:
            yoy_pct = None
        rows.append({"Month": month_name, "This_Year": this_year, "Last_Year": last_year, "Registrations_This": total_this, "Registrations_Last": total_last, "YoY_Pct": yoy_pct})
    return pd.DataFrame(rows)


def compute_mom_growth(plot_df, month_cols_sorted, year):
    """MoM % = (Month_i - Month_{i-1}) / Month_{i-1} * 100 for consecutive months. Returns df: Month, Volume, MoM_Pct."""
    months_short = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()
    rows = []
    prev_total = None
    for month_name in months_short:
        col = f"{month_name} {year}"
        if col not in plot_df.columns:
            continue
        total = plot_df[col].sum()
        mom_pct = None
        if prev_total is not None and prev_total > 0:
            mom_pct = round((total - prev_total) / prev_total * 100, 1)
        rows.append({"Month": month_name, "Volume": total, "MoM_Pct": mom_pct})
        prev_total = total
    return pd.DataFrame(rows)


def compute_cagr(start_vol, end_vol, n_years):
    """CAGR = (end/start)^(1/n) - 1. Returns percentage or None."""
    if start_vol is None or end_vol is None or start_vol <= 0 or n_years <= 0:
        return None
    return round((pow(end_vol / start_vol, 1 / n_years) - 1) * 100, 1)


def render_footer():
    st.markdown("---")
    st.caption(
        "**Source:** Vahan Parivahan (MoRTH). Data is compiled from the Vahan dashboard. "
        "For institutional use only. Not a recommendation. Past performance does not guarantee future results."
    )


def run_dashboard():
    st.set_page_config(page_title="Vahan Vehicle Registration — Institutional", layout="wide", initial_sidebar_state="expanded")
    # Professional styling for broking clients
    st.markdown("""
        <style>
        .metric-card { background: linear-gradient(135deg, #1e3a5f 0%, #0d2137 100%); color: white; padding: 1rem 1.25rem; border-radius: 8px; margin: 0.25rem 0; }
        .metric-card h3 { margin: 0; font-size: 0.85rem; opacity: 0.9; }
        .metric-card p { margin: 0.25rem 0 0 0; font-size: 1.5rem; font-weight: 600; }
        .stTabs [data-baseweb="tab-list"] { gap: 0.5rem; }
        .stTabs [data-baseweb="tab"] { padding: 0.75rem 1.25rem; font-weight: 500; }
        </style>
    """, unsafe_allow_html=True)

    xl = load_master_sheet()
    if xl is None:
        st.error(f"Master sheet not found at {MASTER_PATH}. Run: python vahan_full_pipeline.py --compile")
        return

    sheet_names = xl.sheet_names
    labels = SHEET_LABELS

    st.title("India Vehicle Registration Dashboard")
    st.caption("Vahan Parivahan (MoRTH) · Maker, Fuel, State & segment")

    tab_exec, tab_segment, tab_trends, tab_rankings = st.tabs([
        "Overview",
        "Charts",
        "Growth",
        "Rankings"
    ])

    # ---------- Overview (per-sheet) ----------
    with tab_exec:
        overview_sheet = st.selectbox(
            "Overview for",
            sheet_names,
            index=sheet_names.index("Maker") if "Maker" in sheet_names else 0,
            key="overview_sheet",
            format_func=lambda x: labels.get(x, {}).get("title", x),
        )
        sheet_title = labels.get(overview_sheet, {}).get("title", overview_sheet)
        df_ov, key_col_ov, months_ov, totals_ov, key_label_ov = load_sheet_data(overview_sheet)

        if df_ov is None or df_ov.empty:
            st.info("No data for this sheet.")
            render_footer()
        elif months_ov:
            # Same-period comparison using selected sheet's month columns
            last_month_col = months_ov[-1]
            last_parsed = parse_month_year(last_month_col)
            latest_year = last_parsed[0] if last_parsed else None
            latest_month_num = last_parsed[1] if last_parsed else 12
            prev_year = latest_year - 1 if latest_year else None
            month_names = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()
            period_label = f"Jan–{month_names[latest_month_num - 1]} {latest_year}" if latest_month_num else str(latest_year)

            cols_this = [c for c in months_ov if parse_month_year(c) and parse_month_year(c)[0] == latest_year and parse_month_year(c)[1] <= latest_month_num]
            cols_last = [c for c in months_ov if parse_month_year(c) and parse_month_year(c)[0] == prev_year and parse_month_year(c)[1] <= latest_month_num] if prev_year else []

            total_vol = df_ov[cols_this].sum().sum() if cols_this else 0
            total_vol_prev = df_ov[cols_last].sum().sum() if cols_last else 0
            yoy_pct = round((total_vol - total_vol_prev) / total_vol_prev * 100, 1) if total_vol_prev and total_vol_prev > 0 else None

            vol_by_entity = df_ov[cols_this].sum(axis=1) if cols_this else pd.Series()
            n_active = (vol_by_entity > 0).sum()
            entity_label = key_label_ov if n_active == 1 else (key_label_ov + "s")

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Registrations (same period)", f"{int(total_vol):,}", delta=f"{yoy_pct}% YoY" if yoy_pct is not None else None)
            with col2:
                st.metric("Period", period_label, delta=f"vs same months {prev_year}" if prev_year else None)
            with col3:
                st.metric(f"Active {entity_label}", str(n_active), None)
            with col4:
                st.metric("Sheet", sheet_title, None)

            # Chart: Maker = segment mix (grouped bar); others = top 10 breakdown (grouped bar)
            if overview_sheet == "Maker":
                maker_to_cat = get_maker_to_category()
                df_ov_c = df_ov.copy()
                df_ov_c["_cat"] = df_ov_c[key_col_ov].apply(lambda k: get_maker_category_for_key(k, maker_to_cat))
                seg_df = df_ov_c[df_ov_c["_cat"].notna()]
                if not seg_df.empty and cols_this:
                    seg_vol = seg_df.groupby("_cat")[cols_this].sum().sum(axis=1).sort_values(ascending=False)
                    if seg_vol.sum() > 0:
                        st.subheader(f"Segment mix — {sheet_title}")
                        colors = ["#1e3a5f", "#2e5a8f", "#3d7ab8", "#5a9fd4", "#7eb8e0"]
                        fig_seg = go.Figure()
                        for i, cat in enumerate(seg_vol.index):
                            fig_seg.add_trace(go.Bar(x=[cat], y=[seg_vol.loc[cat]], name=cat, marker_color=colors[i % len(colors)], showlegend=False))
                        fig_seg.update_layout(barmode="group", height=350, title=f"PV / 2W / CV / EV-PV / Tractors — {period_label}", margin=dict(t=40), xaxis_title="", yaxis_title="Registrations")
                        st.plotly_chart(fig_seg, use_container_width=True, config=PLOTLY_CONFIG)
            else:
                top_n_ov = min(10, len(df_ov))
                df_ov["_vol"] = df_ov[cols_this].sum(axis=1)
                top_ov = df_ov.nlargest(top_n_ov, "_vol")
                if not top_ov.empty and total_vol > 0:
                    st.subheader(f"Top {top_n_ov} — {sheet_title} ({period_label})")
                    colors = ["#1e3a5f", "#2e5a8f", "#3d7ab8", "#5a9fd4", "#7eb8e0", "#9ecae1", "#c6dbef", "#deebf7", "#3182bd", "#6baed6"]
                    fig_ov = go.Figure()
                    fig_ov.add_trace(go.Bar(x=top_ov[key_col_ov].tolist(), y=top_ov["_vol"].tolist(), marker_color=colors[:len(top_ov)], showlegend=False))
                    fig_ov.update_layout(height=350, margin=dict(t=40), xaxis_title="", yaxis_title="Registrations", xaxis_tickangle=-35)
                    st.plotly_chart(fig_ov, use_container_width=True, config=PLOTLY_CONFIG)

            c1, c2 = st.columns(2)
            with c1:
                df_ov["_vol"] = df_ov[cols_this].sum(axis=1)
                top5 = df_ov.nlargest(5, "_vol")[[key_col_ov, "_vol"]].copy()
                top5.columns = [key_label_ov, "Volume"]
                top5["Share %"] = (top5["Volume"] / total_vol * 100).round(1) if total_vol else 0
                st.markdown(f"**Top 5 {key_label_ov.lower()}s**")
                st.dataframe(top5, use_container_width=True, hide_index=True)
            with c2:
                if overview_sheet == "Maker":
                    state_df, st_key, st_months, _, _ = load_sheet_data("State")
                    if state_df is not None and not state_df.empty and st_months:
                        st_cols_this = [c for c in st_months if parse_month_year(c) and parse_month_year(c)[0] == latest_year and parse_month_year(c)[1] <= latest_month_num]
                        if st_cols_this:
                            state_df = state_df.copy()
                            state_df["_vol"] = state_df[st_cols_this].sum(axis=1)
                            top5s = state_df.nlargest(5, "_vol")[[st_key, "_vol"]].copy()
                            top5s.columns = ["State", "Volume"]
                            st.markdown("**Top 5 states**")
                            st.dataframe(top5s, use_container_width=True, hide_index=True)
                        else:
                            st.caption("No state data for this period.")
                    else:
                        st.caption("State sheet not available.")
                else:
                    st.caption(f"Select Maker in Overview for state comparison.")

            # Mini trend: last 24 months for this sheet
            last_24 = months_ov[-24:] if len(months_ov) >= 24 else months_ov
            months_parsed = [parse_month_year(c) for c in last_24]
            dates = [pd.Timestamp(y, m, 1) for y, m in months_parsed if (y and m)]
            if dates and len(dates) == len(last_24):
                trend_df = pd.DataFrame({"Date": dates, "Registrations": [df_ov[c].sum() for c in last_24]})
                st.subheader(f"Monthly trend — {sheet_title}")
                fig_t = px.line(trend_df, x="Date", y="Registrations", title=f"Total registrations ({sheet_title})")
                fig_t.update_layout(height=280, margin=dict(t=30))
                st.plotly_chart(fig_t, use_container_width=True, config=PLOTLY_CONFIG)
        elif totals_ov:
            # No month columns: use full latest year vs full prev year
            latest_year = int(totals_ov[-1].split()[-1])
            prev_year = latest_year - 1
            year_col = next((c for c in totals_ov if str(latest_year) in c), None)
            prev_col = next((c for c in totals_ov if str(prev_year) in c), None)
            if year_col:
                total_vol = df_ov[year_col].sum()
                total_vol_prev = df_ov[prev_col].sum() if prev_col else 0
                yoy_pct = round((total_vol - total_vol_prev) / total_vol_prev * 100, 1) if total_vol_prev > 0 else None
                st.metric(f"Registrations — {sheet_title}", f"{int(total_vol):,}", delta=f"{yoy_pct}% YoY" if yoy_pct else None)
                st.caption(f"Full year {latest_year} vs {prev_year} (no month-level data)")
                vol_by_entity = df_ov[year_col]
                n_active = (vol_by_entity > 0).sum()
                st.metric(f"Active {key_label_ov}s", str(n_active), None)
        else:
            st.info("This sheet has no month or yearly columns.")
        render_footer()

    # ---------- By Segment ----------
    with tab_segment:
        sheet_options = sheet_names
        selected_sheet = st.sidebar.selectbox("Select sheet", sheet_options, index=0)
        key_label = labels.get(selected_sheet, {}).get("key_label", "Key")
        sheet_title = labels.get(selected_sheet, {}).get("title", selected_sheet)

        df = pd.read_excel(MASTER_PATH, sheet_name=selected_sheet)
        key_col = df.columns[0]
        df = df.rename(columns={key_col: key_label})
        key_col = key_label
        df = normalize_total_column_names(df)

        month_cols = get_month_columns(df)
        total_cols = get_total_year_columns(df)
        numeric_cols = [c for c in df.columns if c != key_col]
        df = ensure_numeric(df, numeric_cols)

        st.subheader(sheet_title)

        if not month_cols and not total_cols:
            st.warning("No month or yearly columns in this sheet.")
            st.dataframe(df.head(20), use_container_width=True)
        else:
            month_cols_sorted = sort_month_columns(month_cols) if month_cols else []

            show_maker_classification = False
            top_n = 15 if len(df) > 20 else 0
            with st.sidebar.expander("Options", expanded=False):
                if selected_sheet == "Maker" and month_cols_sorted:
                    show_maker_classification = st.checkbox("Group by segment (PV, 2W, CV, EV, Tractors)", value=False)
                top_n = st.slider("Show top N (0 = all)", 0, 50, 15 if len(df) > 20 else 0)
            if show_maker_classification:
                maker_to_cat = get_maker_to_category()
                plot_df = aggregate_by_maker_classification(df, key_col, maker_to_cat, month_cols_sorted)
                if plot_df.empty:
                    st.warning("No makers matched the classification. Showing raw data.")
                    plot_df = df.set_index(key_col)[month_cols_sorted] if month_cols_sorted else pd.DataFrame()
                    key_label_plot = key_label
                else:
                    key_label_plot = "Classification"
            else:
                if top_n > 0 and (month_cols_sorted or total_cols):
                    if total_cols:
                        df["_total"] = df[total_cols].sum(axis=1)
                    else:
                        df["_total"] = df[month_cols_sorted].sum(axis=1)
                    top_keys = df.nlargest(top_n, "_total")[key_col].tolist()
                    df = df[df[key_col].isin(top_keys)].drop(columns=["_total"], errors="ignore")
                plot_df = df.set_index(key_col)
                key_label_plot = key_label
                cols_for_plot = (month_cols_sorted or []) + [c for c in total_cols if c in plot_df.columns]
                if cols_for_plot:
                    plot_df = plot_df[cols_for_plot]

            chart_options = [
                "Monthly trend (line)", "Monthly trend (area)", "Yearly totals (stacked bar)",
                "Volume (Maker & Category)", "YoY growth (This Jan vs Last Jan)", "MoM growth (Jan vs Feb)",
                "Market share", "Share by year (pie)", "Data table"
            ]
            chart_labels = ["Monthly trend", "Monthly (stacked)", "Yearly totals", "Volume table", "YoY growth", "MoM growth", "Market share", "Share (pie)", "Raw data"]
            chart_choice = st.sidebar.selectbox("Chart", chart_options, format_func=lambda x: chart_labels[chart_options.index(x)], index=0)
            available_years = get_available_years(month_cols_sorted, total_cols)

            if chart_choice == "Monthly trend (line)" and month_cols_sorted and not plot_df.empty:
                long = df_to_time_series(plot_df, month_cols_sorted, key_label_plot)
                if long.columns[0] != key_label_plot:
                    long = long.rename(columns={long.columns[0]: key_label_plot})
                fig = px.line(long, x="Date", y="Registrations", color=key_label_plot, title=f"{sheet_title} — Monthly Registrations")
                fig.update_layout(height=500, xaxis_title="Month", yaxis_title="Number of Registrations", legend_title=key_label_plot)
                fig.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02))
                st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

            elif chart_choice == "Monthly trend (area)" and month_cols_sorted and not plot_df.empty:
                long = df_to_time_series(plot_df, month_cols_sorted, key_label_plot)
                if long.columns[0] != key_label_plot:
                    long = long.rename(columns={long.columns[0]: key_label_plot})
                fig = px.bar(long, x="Date", y="Registrations", color=key_label_plot, barmode="stack", title=f"{sheet_title} — Monthly Registrations (Stacked)")
                fig.update_layout(height=500, xaxis_title="Month", yaxis_title="Number of Registrations", legend_title=key_label_plot)
                fig.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02))
                st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

            elif chart_choice == "Yearly totals (stacked bar)" and not plot_df.empty:
                total_cols_in_plot = [c for c in total_cols if c in plot_df.columns]
                chart_total_cols = None
                if total_cols_in_plot:
                    yearly_df = plot_df[total_cols_in_plot].copy()
                    chart_total_cols = total_cols_in_plot
                elif month_cols_sorted:
                    years = sorted(set(parse_month_year(c)[0] for c in month_cols_sorted if parse_month_year(c)))
                    yearly_df = pd.DataFrame(index=plot_df.index)
                    for y in years:
                        y_cols = [c for c in month_cols_sorted if parse_month_year(c) and parse_month_year(c)[0] == y]
                        if y_cols:
                            yearly_df[f"Total {y}"] = plot_df[y_cols].sum(axis=1)
                    chart_total_cols = list(yearly_df.columns)
                else:
                    yearly_df = pd.DataFrame()
                if not yearly_df.empty and chart_total_cols:
                    long = yearly_df.reset_index().melt(id_vars=[yearly_df.index.name or key_label_plot], value_vars=chart_total_cols, var_name="Year", value_name="Registrations")
                    long["Year"] = long["Year"].str.extract(r"(\d{4})$", expand=False).astype(int)
                    id_col = [c for c in long.columns if c not in ("Year", "Registrations")][0]
                    if id_col != key_label_plot:
                        long = long.rename(columns={id_col: key_label_plot})
                        id_col = key_label_plot
                    fig = px.bar(long, x="Year", y="Registrations", color=id_col, barmode="stack", title=f"{sheet_title} — Yearly Registrations")
                    fig.update_layout(height=500, xaxis_title="Year", yaxis_title="Number of Registrations", legend_title=key_label_plot)
                    st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

            elif chart_choice == "Share by year (pie)" and not plot_df.empty:
                total_cols_in_plot = [c for c in total_cols if c in plot_df.columns]
                chart_total_cols = None
                if total_cols_in_plot:
                    yearly_df = plot_df[total_cols_in_plot]
                    chart_total_cols = total_cols_in_plot
                elif month_cols_sorted:
                    years = sorted(set(parse_month_year(c)[0] for c in month_cols_sorted if parse_month_year(c)))
                    yearly_df = pd.DataFrame(index=plot_df.index)
                    for y in years:
                        y_cols = [c for c in month_cols_sorted if parse_month_year(c) and parse_month_year(c)[0] == y]
                        if y_cols:
                            yearly_df[f"Total {y}"] = plot_df[y_cols].sum(axis=1)
                    chart_total_cols = list(yearly_df.columns)
                else:
                    yearly_df = pd.DataFrame()
                if not yearly_df.empty and chart_total_cols:
                    year_sel = st.sidebar.selectbox("Year", chart_total_cols, index=len(chart_total_cols) - 1, key="year_share")
                    s = yearly_df[year_sel].sort_values(ascending=False)
                    s = s[s > 0]
                    if s.empty:
                        st.info("No data for selected year.")
                    else:
                        pie_df = s.reset_index()
                        name_col, value_col = pie_df.columns[0], pie_df.columns[1]
                        fig = px.pie(pie_df, names=name_col, values=value_col, title=f"{sheet_title} — Share of Registrations ({year_sel})")
                        fig.update_layout(height=500, legend_title=key_label_plot)
                        st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

            elif chart_choice == "YoY growth (This Jan vs Last Jan)" and not plot_df.empty and month_cols_sorted and len(available_years) >= 2:
                sel_year = st.sidebar.selectbox("Year", available_years[1:], index=len(available_years[1:]) - 1, key="year_yoy")
                yoy_df = compute_yoy_growth(plot_df, month_cols_sorted, sel_year)
                if not yoy_df.empty:
                    st.markdown(f"**Year-over-Year growth:** {sel_year} vs {sel_year - 1} (same month)")
                    fig = px.bar(yoy_df, x="Month", y="YoY_Pct", title=f"{sheet_title} — YoY % by month ({sel_year} vs {sel_year - 1})")
                    fig.update_layout(height=400, xaxis_title="Month", yaxis_title="YoY %", yaxis_tickformat=".1f")
                    fig.add_hline(y=0, line_dash="dash", line_color="gray")
                    st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)
                    yoy_display = yoy_df.rename(columns={"Registrations_This": f"Vol {sel_year}", "Registrations_Last": f"Vol {sel_year - 1}", "YoY_Pct": "YoY %"})
                    st.dataframe(yoy_display[["Month", f"Vol {sel_year - 1}", f"Vol {sel_year}", "YoY %"]], use_container_width=True, hide_index=True)
                else:
                    st.info("No overlapping months for selected years.")

            elif chart_choice == "MoM growth (Jan vs Feb)" and not plot_df.empty and month_cols_sorted and available_years:
                sel_year = st.sidebar.selectbox("Year", available_years, index=len(available_years) - 1, key="year_mom")
                mom_df = compute_mom_growth(plot_df, month_cols_sorted, sel_year)
                if not mom_df.empty:
                    st.markdown(f"**Month-over-Month growth:** {sel_year} (each month vs previous month)")
                    fig = px.bar(mom_df, x="Month", y="MoM_Pct", title=f"{sheet_title} — MoM % ({sel_year})")
                    fig.update_layout(height=400, xaxis_title="Month", yaxis_title="MoM %", yaxis_tickformat=".1f")
                    fig.add_hline(y=0, line_dash="dash", line_color="gray")
                    st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)
                    st.dataframe(mom_df.rename(columns={"Volume": "Registrations", "MoM_Pct": "MoM %"}), use_container_width=True, hide_index=True)
                else:
                    st.info("No monthly data for selected year.")

            elif chart_choice == "Market share" and not plot_df.empty:
                total_cols_in_plot = [c for c in total_cols if c in plot_df.columns]
                period_options = []
                if total_cols_in_plot:
                    period_options = [(c, f"Full year {c}") for c in total_cols_in_plot]
                if month_cols_sorted:
                    for c in month_cols_sorted[-24:]:
                        period_options.append((c, c))
                if not period_options:
                    st.warning("No period available for market share.")
                else:
                    period_labels = [p[1] for p in period_options]
                    period_keys = [p[0] for p in period_options]
                    sel_period_label = st.sidebar.selectbox("Period", period_labels, index=len(period_labels) - 1, key="period_share")
                    sel_period_col = period_keys[period_labels.index(sel_period_label)]
                    if sel_period_col not in plot_df.columns:
                        st.warning("Selected period column not in data.")
                    else:
                        share_s = plot_df[sel_period_col]
                        total = share_s.sum()
                        if total and total > 0:
                            share_pct = (share_s / total * 100).round(1)
                            share_df = share_pct.sort_values(ascending=False).reset_index()
                            share_df.columns = [key_label_plot, "Share %"]
                            share_df["Volume"] = share_s.reindex(share_df[key_label_plot]).values
                            fig = px.bar(share_df.head(20), x="Share %", y=key_label_plot, orientation="h", title=f"{sheet_title} — Market share ({sel_period_label})")
                            fig.update_layout(height=500, xaxis_title="Share %", yaxis_title="")
                            st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)
                            st.dataframe(share_df[[key_label_plot, "Volume", "Share %"]], use_container_width=True, hide_index=True)
                        else:
                            st.info("No volume in selected period.")

            elif chart_choice == "Volume (Maker & Category)" and not plot_df.empty:
                if not available_years:
                    st.warning("No year data for volume.")
                else:
                    sel_year_vol = st.sidebar.selectbox("Year", available_years, index=len(available_years) - 1, key="year_vol")
                    total_cols_in_plot = [c for c in total_cols if c in plot_df.columns]
                    year_col = next((c for c in total_cols_in_plot if str(sel_year_vol) in str(c)), None) if total_cols_in_plot else None
                    if year_col:
                        total_vol = plot_df[year_col].sum()
                        st.metric("Total volume (selected year)", f"{total_vol:,}")
                        vol_by_key = plot_df[year_col].sort_values(ascending=False)
                        vol_df = vol_by_key.reset_index()
                        vol_df.columns = [key_label_plot, "Volume"]
                        vol_df["Share %"] = (vol_df["Volume"] / total_vol * 100).round(1)
                        st.dataframe(vol_df, use_container_width=True, hide_index=True)
                    else:
                        y_cols = [c for c in (month_cols_sorted or []) if parse_month_year(c) and parse_month_year(c)[0] == sel_year_vol]
                        if y_cols:
                            total_vol = plot_df[y_cols].sum().sum()
                            st.metric("Total volume (selected year)", f"{total_vol:,}")
                            vol_by_key = plot_df[y_cols].sum(axis=1).sort_values(ascending=False)
                            vol_df = vol_by_key.reset_index()
                            vol_df.columns = [key_label_plot, "Volume"]
                            vol_df["Share %"] = (vol_df["Volume"] / total_vol * 100).round(1)
                            st.dataframe(vol_df, use_container_width=True, hide_index=True)
                        else:
                            st.info("No data for selected year.")

            st.subheader("Data")
            display_df = df.copy()
            for c in month_cols + total_cols:
                if c in display_df.columns:
                    display_df[c] = pd.to_numeric(display_df[c], errors="coerce").fillna(0).astype("int64")
            st.dataframe(display_df, use_container_width=True, height=400)
        render_footer()

    # ---------- Growth ----------
    with tab_trends:
        st.subheader("Growth")
        st.sidebar.markdown("### Growth")
        sheet_t = st.sidebar.selectbox("Data", sheet_names, key="trends_sheet")
        df_t, key_t, months_t, totals_t, key_label_t = load_sheet_data(sheet_t)
        if df_t is not None and not df_t.empty and (totals_t or months_t):
            years_t = get_available_years(months_t, totals_t)
            if len(years_t) >= 2:
                start_y = st.sidebar.selectbox("Start year", years_t, key="start_y")
                end_y = st.sidebar.selectbox("End year", years_t, index=min(years_t.index(max(years_t)), len(years_t) - 1), key="end_y")
                if start_y >= end_y:
                    st.warning("Pick Start year before End year.")
                else:
                    start_col = next((c for c in totals_t if str(start_y) in c), None) if totals_t else None
                    end_col = next((c for c in totals_t if str(end_y) in c), None) if totals_t else None
                    if not start_col and months_t:
                        start_cols = [c for c in months_t if parse_month_year(c) and parse_month_year(c)[0] == start_y]
                        start_col = start_cols[0] if start_cols else None
                    if not end_col and months_t:
                        end_cols = [c for c in months_t if parse_month_year(c) and parse_month_year(c)[0] == end_y]
                        end_col = end_cols[0] if end_cols else None
                    if start_col and end_col:
                        plot_t = df_t.set_index(key_t)
                        if start_col in plot_t.columns and end_col in plot_t.columns:
                            start_vols = plot_t[start_col].replace(0, pd.NA) if start_col in plot_t.columns else pd.Series()
                            end_vols = plot_t[end_col]
                            total_start = plot_t[start_col].sum()
                            total_end = plot_t[end_col].sum()
                            cagr_total = compute_cagr(total_start, total_end, end_y - start_y)
                            st.metric("Industry CAGR (%)", f"{cagr_total}%" if cagr_total is not None else "—", delta=f"{start_y} → {end_y}")
                            cagr_by_entity = []
                            for idx in plot_t.index:
                                s, e = plot_t.loc[idx, start_col], plot_t.loc[idx, end_col]
                                cagr = compute_cagr(s, e, end_y - start_y)
                                if cagr is not None:
                                    cagr_by_entity.append({key_label_t: idx, "CAGR %": cagr, f"Vol {start_y}": s, f"Vol {end_y}": e})
                            if cagr_by_entity:
                                cagr_df = pd.DataFrame(cagr_by_entity).sort_values("CAGR %", ascending=False)
                                st.dataframe(cagr_df.head(20), use_container_width=True, hide_index=True)
                    # Dual-axis in expander
                    if months_t and len(years_t) >= 2:
                        with st.expander("One entity: volume & YoY %", expanded=False):
                            plot_t = df_t.set_index(key_t)
                            cols_plot = (months_t or []) + [c for c in totals_t if c in plot_t.columns]
                            if cols_plot:
                                plot_t = plot_t[cols_plot]
                            entities = plot_t.index.tolist()
                            sel_entity = st.selectbox("Entity", entities[:50], key="dual_entity")
                            if sel_entity in plot_t.index:
                                row = plot_t.loc[sel_entity]
                                if totals_t:
                                    year_vals = [row[c] for c in totals_t if c in row.index]
                                    years_axis = [int(c.split()[-1]) for c in totals_t if c in row.index]
                                else:
                                    years_axis = sorted(set(parse_month_year(c)[0] for c in months_t if parse_month_year(c)))
                                    year_vals = [row[[c for c in months_t if parse_month_year(c) and parse_month_year(c)[0] == y]].sum() for y in years_axis]
                                if len(years_axis) >= 2 and len(year_vals) >= 2:
                                    yoy_pcts = [None] + [round((year_vals[i] - year_vals[i - 1]) / year_vals[i - 1] * 100, 1) if year_vals[i - 1] else None for i in range(1, len(year_vals))]
                                    fig_dual = make_subplots(specs=[[{"secondary_y": True}]])
                                    fig_dual.add_trace(go.Bar(x=years_axis, y=year_vals, name="Volume"), secondary_y=False)
                                    fig_dual.add_trace(go.Scatter(x=years_axis, y=yoy_pcts, name="YoY %", line=dict(dash="dash")), secondary_y=True)
                                    fig_dual.update_layout(title=f"{sel_entity} — Volume & YoY %", height=400)
                                    fig_dual.update_yaxes(title_text="Volume", secondary_y=False)
                                    fig_dual.update_yaxes(title_text="YoY %", secondary_y=True)
                                    st.plotly_chart(fig_dual, use_container_width=True, config=PLOTLY_CONFIG)
        else:
            st.info("Choose Data and a range of years.")
        render_footer()

    # ---------- Rankings ----------
    with tab_rankings:
        st.subheader("Rankings")
        st.sidebar.markdown("### Rankings")
        sheet_r = st.sidebar.selectbox("Data", sheet_names, key="rank_sheet")
        df_r, key_r, months_r, totals_r, key_label_r = load_sheet_data(sheet_r)
        if df_r is not None and not df_r.empty and (totals_r or months_r):
            years_r = get_available_years(months_r, totals_r)
            year_rank = st.sidebar.selectbox("Year", years_r, index=len(years_r) - 1 if years_r else 0, key="year_rank")
            year_col_r = next((c for c in totals_r if str(year_rank) in c), None) if totals_r else None
            if not year_col_r and months_r:
                y_cols = [c for c in months_r if parse_month_year(c) and parse_month_year(c)[0] == year_rank]
                if y_cols:
                    df_r["_rank_vol"] = df_r[y_cols].sum(axis=1)
                    year_col_r = "_rank_vol"
            prev_col_r = next((c for c in totals_r if str(year_rank - 1) in c), None) if totals_r and year_rank - 1 in years_r else None
            if year_col_r and year_col_r in df_r.columns:
                total_r = df_r[year_col_r].sum()
                rank_df = df_r[[key_r, year_col_r]].copy()
                rank_df.columns = [key_label_r, "Volume"]
                rank_df["Share %"] = (rank_df["Volume"] / total_r * 100).round(1)
                rank_df["YoY %"] = None
                if prev_col_r and prev_col_r in df_r.columns:
                    for i, row in rank_df.iterrows():
                        prev_val = df_r.loc[df_r[key_r] == row[key_label_r], prev_col_r].sum()
                        if prev_val and prev_val > 0:
                            rank_df.loc[rank_df[key_label_r] == row[key_label_r], "YoY %"] = round((row["Volume"] - prev_val) / prev_val * 100, 1)
                rank_df = rank_df.sort_values("Volume", ascending=False).reset_index(drop=True)
                rank_df.insert(0, "Rank", range(1, len(rank_df) + 1))
                st.dataframe(rank_df, use_container_width=True, height=400, hide_index=True)
                csv = rank_df.to_csv(index=False).encode("utf-8")
                st.download_button("Download CSV", csv, file_name=f"vahan_rankings_{sheet_r}_{year_rank}.csv", mime="text/csv", key="dl_rank")
            with st.expander("Compare entities", expanded=False):
                plot_r = df_r.set_index(key_r)
                cols_r = (months_r or []) + [c for c in totals_r if c in plot_r.columns]
                if cols_r:
                    plot_r = plot_r[cols_r]
                    entities_r = plot_r.index.tolist()
                    compare_sel = st.multiselect("Pick 2–5", entities_r, default=entities_r[:2] if len(entities_r) >= 2 else entities_r[:1], key="compare_sel", max_selections=5)
                    if len(compare_sel) >= 1 and months_r:
                        comp_df = plot_r.loc[compare_sel, months_r].T
                        comp_df = comp_df.reset_index()
                        comp_df.columns = ["Month_Year", *compare_sel]
                        comp_long = comp_df.melt(id_vars=["Month_Year"], var_name=key_label_r, value_name="Registrations")
                        def to_date(my):
                            p = parse_month_year(my)
                            return pd.Timestamp(p[0], p[1], 1) if p and len(p) == 2 else pd.NaT
                        comp_long["Date"] = comp_long["Month_Year"].apply(to_date)
                        comp_long = comp_long.dropna(subset=["Date"])
                        fig_c = px.line(comp_long, x="Date", y="Registrations", color=key_label_r, title="Compare over time")
                        fig_c.update_layout(height=400)
                        st.plotly_chart(fig_c, use_container_width=True, config=PLOTLY_CONFIG)
        else:
            st.info("Choose Data with yearly/monthly columns.")
        render_footer()


if __name__ == "__main__":
    run_dashboard()
