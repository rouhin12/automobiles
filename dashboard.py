"""
Vahan Vehicle Registration Dashboard — Institutional / Broking Client Edition.
Supports all sheets: Fuel, Maker, Norms, State, Vehicle Category, Vehicle Class.
Run: streamlit run dashboard.py
"""
import os
import re
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

PLOTLY_CONFIG = {
    "displayModeBar": True,
    "displaylogo": False,
    "modeBarButtonsToRemove": ["lasso2d", "select2d"],
    "toImageButtonOptions": {"format": "png", "filename": "vahan_chart", "scale": 2},
}

# Chart suggestions per sheet: (chart_key, display_label). First = default for that sheet.
CHART_SUGGESTIONS = {
    "Fuel": [
        ("treemap_period", "Share by fuel (treemap)"),
        ("stacked_area", "Monthly trend by fuel (stacked area)"),
        ("line_multi", "Monthly trend by fuel (line)"),
        ("pie_period", "Share by fuel (pie)"),
        ("yoy_growth", "YoY growth by month"),
        ("market_share_hbar", "Market share (horizontal bar)"),
        ("data_table", "Data table"),
    ],
    "Maker": [
        ("treemap_period", "Share by maker (treemap)"),
        ("stacked_area", "Monthly trend by maker (stacked area)"),
        ("line_multi", "Monthly trend by maker (line)"),
        ("pie_period", "Share by maker (pie)"),
        ("yearly_stacked", "Yearly totals (stacked)"),
        ("segment_treemap", "Segment mix PV/2W/CV/EV (treemap)"),
        ("data_table", "Data table"),
    ],
    "State": [
        ("treemap_period", "Share by state (treemap)"),
        ("stacked_area", "Monthly trend by state (stacked area)"),
        ("line_multi", "Monthly trend by state (line)"),
        ("market_share_hbar", "Market share (horizontal bar)"),
        ("pie_period", "Share by state (pie)"),
        ("data_table", "Data table"),
    ],
    "Norms": [
        ("treemap_period", "Share by norm (treemap)"),
        ("stacked_area", "Monthly trend by norm (stacked area)"),
        ("pie_period", "Share by norm (pie)"),
        ("data_table", "Data table"),
    ],
    "Vehicle Category": [
        ("treemap_period", "Share by category (treemap)"),
        ("stacked_area", "Monthly trend (stacked area)"),
        ("pie_period", "Share (pie)"),
        ("data_table", "Data table"),
    ],
    "Vehicle Class": [
        ("treemap_period", "Share by class (treemap)"),
        ("stacked_area", "Monthly trend (stacked area)"),
        ("pie_period", "Share (pie)"),
        ("data_table", "Data table"),
    ],
}

DOWNLOAD_DIR = os.path.join(os.path.dirname(__file__), "downloads")
MASTER_PATH = os.path.join(DOWNLOAD_DIR, "master_sheet.xlsx")

MONTH_YEAR_PATTERN = re.compile(r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})$", re.I)
TOTAL_YEAR_PATTERN = re.compile(r"^(Total|Col_\d+)\s+(\d{4})$", re.I)
MONTHS_ORDER = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()

SHEET_LABELS = {
    "Fuel": {"key_label": "Fuel Type", "title": "Fuel Wise"},
    "Maker": {"key_label": "Maker", "title": "Maker Wise"},
    "Norms": {"key_label": "Norm", "title": "Norms Wise"},
    "State": {"key_label": "State", "title": "Statewise"},
    "Vehicle Category": {"key_label": "Vehicle Category", "title": "Vehicle Category Wise"},
    "Vehicle Class": {"key_label": "Vehicle Class", "title": "Vehicle Class Wise"},
}


def normalize_total_column_names(df):
    rename = {}
    for c in df.columns:
        m = TOTAL_YEAR_PATTERN.match(str(c).strip())
        if m and m.group(1).upper().startswith("COL_"):
            rename[c] = f"Total {m.group(2)}"
    return df.rename(columns=rename) if rename else df


def load_master_sheet():
    if not os.path.isfile(MASTER_PATH):
        return None
    return pd.ExcelFile(MASTER_PATH)


def get_month_columns(df):
    return [c for c in df.columns if MONTH_YEAR_PATTERN.match(str(c).strip())]


def get_total_year_columns(df):
    total_cols = [c for c in df.columns if TOTAL_YEAR_PATTERN.match(str(c).strip())]
    years = []
    for c in total_cols:
        m = TOTAL_YEAR_PATTERN.match(str(c).strip())
        if m:
            years.append((int(m.group(2)), c))
    return [c for _, c in sorted(years)]


def parse_month_year(col):
    m = MONTH_YEAR_PATTERN.match(str(col).strip())
    if not m:
        return None
    month_str, year_str = m.group(1), m.group(2)
    try:
        month_num = MONTHS_ORDER.index(month_str[:3].capitalize()) + 1
        return (int(year_str), month_num)
    except (ValueError, IndexError):
        return None


def sort_month_columns(cols):
    return sorted(cols, key=lambda c: parse_month_year(c) or (0, 0))


def ensure_numeric(df, cols):
    if not cols:
        return df
    df = df.copy()
    def parse_numeric_series(s):
        return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False), errors="coerce")
    df[cols] = df[cols].apply(parse_numeric_series).fillna(0).astype("int64")
    return df


def aggregate_by_maker_classification(df, key_col, maker_to_cat, month_cols_sorted):
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


def get_period_options(month_cols_sorted):
    """Return list of (label, list of columns) for period selector. Option: full months range or YTD-style."""
    if not month_cols_sorted:
        return []
    parsed = [(c, parse_month_year(c)) for c in month_cols_sorted]
    parsed = [(c, p) for c, p in parsed if p]
    if not parsed:
        return []
    by_year_month = {}
    for col, (y, m) in parsed:
        by_year_month[(y, m)] = col
    keys_sorted = sorted(by_year_month.keys())
    options = []
    for i, (y, m) in enumerate(keys_sorted):
        label = f"{MONTHS_ORDER[m-1]} {y}"
        cols = [by_year_month[k] for k in keys_sorted[: i + 1]]
        options.append((label, cols))
    return options


def get_from_to_options(month_cols_sorted):
    """Return list of (label, year, month) for From/To date pickers. Label e.g. 'Jan 2024'."""
    if not month_cols_sorted:
        return []
    parsed = [(c, parse_month_year(c)) for c in month_cols_sorted]
    parsed = [(c, p) for c, p in parsed if p]
    if not parsed:
        return []
    by_year_month = {}
    for col, (y, m) in parsed:
        by_year_month[(y, m)] = col
    keys_sorted = sorted(by_year_month.keys())
    return [(f"{MONTHS_ORDER[m-1]} {y}", y, m) for (y, m) in keys_sorted]


def cols_in_range(month_cols_sorted, from_year, from_month, to_year, to_month):
    """Columns where (year, month) is in [from_ym, to_ym] inclusive."""
    out = []
    for c in month_cols_sorted:
        p = parse_month_year(c)
        if not p:
            continue
        y, m = p
        if (y, m) >= (from_year, from_month) and (y, m) <= (to_year, to_month):
            out.append(c)
    return out


def compute_yoy_growth(plot_df, month_cols_sorted, this_year):
    last_year = this_year - 1
    rows = []
    for month_name in MONTHS_ORDER:
        col_this = f"{month_name} {this_year}"
        col_last = f"{month_name} {last_year}"
        if col_this not in plot_df.columns or col_last not in plot_df.columns:
            continue
        total_this = plot_df[col_this].sum()
        total_last = plot_df[col_last].sum()
        yoy_pct = round((total_this - total_last) / total_last * 100, 1) if total_last and total_last > 0 else None
        rows.append({"Month": month_name, "This_Year": this_year, "Last_Year": last_year, "Registrations_This": total_this, "Registrations_Last": total_last, "YoY_Pct": yoy_pct})
    return pd.DataFrame(rows)


def compute_mom_growth(plot_df, month_cols_sorted, year):
    rows = []
    prev_total = None
    for month_name in MONTHS_ORDER:
        col = f"{month_name} {year}"
        if col not in plot_df.columns:
            continue
        total = plot_df[col].sum()
        mom_pct = round((total - prev_total) / prev_total * 100, 1) if prev_total and prev_total > 0 else None
        rows.append({"Month": month_name, "Volume": total, "MoM_Pct": mom_pct})
        prev_total = total
    return pd.DataFrame(rows)


def compute_cagr(start_vol, end_vol, n_years):
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
    st.set_page_config(page_title="Vahan Vehicle Registration", layout="wide", initial_sidebar_state="collapsed")
    st.markdown("""
        <style>
        .metric-card { background: linear-gradient(135deg, #1a365d 0%, #2c5282 100%); color: white; padding: 1rem 1.25rem; border-radius: 10px; margin: 0.25rem 0; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
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

    tab_overview, tab_charts, tab_trends, tab_rankings = st.tabs(["Overview", "Charts", "Growth", "Rankings"])

    # ---- Overview: no selector, briefs of all aspects, no View ----
    with tab_overview:
        # Use Maker (or first sheet) for date range and total/active
        ref_sheet = "Maker" if "Maker" in sheet_names else sheet_names[0]
        df_ref, key_ref, months_ref, totals_ref, key_label_ref = load_sheet_data(ref_sheet)
        if df_ref is None or df_ref.empty or not months_ref:
            st.info("No monthly data available.")
            render_footer()
        else:
            month_cols_sorted = sort_month_columns(months_ref)
            from_to_options = get_from_to_options(month_cols_sorted)
            if not from_to_options:
                st.warning("No month columns found.")
                render_footer()
            else:
                opt_labels = [x[0] for x in from_to_options]
                default_to_idx = len(from_to_options) - 1
                default_from_idx = max(0, default_to_idx - 11)
                col_from, col_to = st.columns(2)
                with col_from:
                    from_idx = st.selectbox("**From** (month–year)", range(len(opt_labels)), format_func=lambda i: opt_labels[i], index=default_from_idx, key="overview_from")
                with col_to:
                    to_idx = st.selectbox("**To** (month–year)", range(len(opt_labels)), format_func=lambda i: opt_labels[i], index=default_to_idx, key="overview_to")
                if from_idx > to_idx:
                    st.warning("From must be ≤ To. Using swapped range.")
                from_idx, to_idx = min(from_idx, to_idx), max(from_idx, to_idx)
                from_label, from_year, from_month = from_to_options[from_idx]
                to_label, to_year, to_month = from_to_options[to_idx]
                cols_this = cols_in_range(month_cols_sorted, from_year, from_month, to_year, to_month)
                period_label = f"{from_label} – {to_label}" if from_idx != to_idx else to_label

                n_months = len(cols_this)
                prev_year = to_year - 1
                cols_last = cols_in_range(month_cols_sorted, prev_year, from_month, prev_year, to_month) if n_months <= 12 else []

                total_vol = df_ref[cols_this].sum().sum() if cols_this else 0
                total_vol_prev = df_ref[cols_last].sum().sum() if cols_last else 0
                yoy_pct = round((total_vol - total_vol_prev) / total_vol_prev * 100, 1) if total_vol_prev and total_vol_prev > 0 else None

                if to_label in df_ref.columns:
                    n_active = (df_ref[to_label] > 0).sum()
                else:
                    n_active = 0

                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Vehicle registrations (period)", f"{int(total_vol):,}", delta=f"{yoy_pct}% YoY" if yoy_pct is not None else None)
                with col2:
                    st.metric("Period", period_label, delta=f"vs same range {prev_year}" if cols_last else None)
                with col3:
                    st.metric(f"Active makers (in {to_label})", str(n_active), None)

                # Segment mix (from Maker data)
                maker_to_cat = get_maker_to_category()
                df_maker, key_maker, months_maker, _, _ = load_sheet_data("Maker")
                if df_maker is not None and not df_maker.empty and cols_this and all(c in df_maker.columns for c in cols_this):
                    df_maker = df_maker.copy()
                    df_maker["_cat"] = df_maker[key_maker].apply(lambda k: get_maker_category_for_key(k, maker_to_cat))
                    seg_df = df_maker[df_maker["_cat"].notna()]
                    if not seg_df.empty:
                        seg_vol = seg_df.groupby("_cat")[cols_this].sum().sum(axis=1).sort_values(ascending=False)
                        if seg_vol.sum() > 0:
                            st.subheader("Segment mix")
                            st.caption("Registrations by segment (PV, 2W, CV, EV-PV, Tractors) for the selected period.")
                            seg_tbl = seg_vol.reset_index()
                            seg_tbl.columns = ["Segment", "Registrations"]
                            seg_tbl["Share %"] = (seg_tbl["Registrations"] / seg_tbl["Registrations"].sum() * 100).round(1)
                            st.dataframe(seg_tbl, use_container_width=True, hide_index=True)

                # Briefs by aspect — single table
                st.subheader("Briefs by aspect")
                from dashboard_config import MAKER_CLASSIFICATION
                brief_rows = []
                for sheet in sheet_names:
                    title = labels.get(sheet, {}).get("title", sheet)
                    df_s, k, months_s, totals_s, kl = load_sheet_data(sheet)
                    if df_s is None or df_s.empty:
                        brief_rows.append({"Aspect": title, "Registrations": 0, "Active": 0, "Top 3": "—"})
                        continue
                    cols_s = cols_in_range(sort_month_columns(months_s), from_year, from_month, to_year, to_month) if months_s else []
                    if not cols_s and totals_s:
                        y_col = next((c for c in totals_s if str(to_year) in c), None)
                        cols_s = [y_col] if y_col and y_col in df_s.columns else []
                    if not cols_s:
                        brief_rows.append({"Aspect": title, "Registrations": 0, "Active": 0, "Top 3": "No data"})
                        continue
                    df_s = df_s.copy()
                    df_s["_vol"] = df_s[cols_s].sum(axis=1)
                    vol = df_s["_vol"].sum()
                    n_ent = int((df_s["_vol"] > 0).sum())
                    top3 = df_s.nlargest(3, "_vol")
                    top3_str = "; ".join(f"{str(r[k])[:24]}{'…' if len(str(r[k])) > 24 else ''} ({r['_vol']/vol*100:.1f}%)" for _, r in top3.iterrows()) if vol else "—"
                    brief_rows.append({"Aspect": title, "Registrations": int(vol), "Active": n_ent, "Top 3": top3_str})
                brief_df = pd.DataFrame(brief_rows)
                st.dataframe(brief_df, use_container_width=True, hide_index=True, column_config={"Registrations": st.column_config.NumberColumn(format="%d")})

                # Top 3 makers by segment + Top 5 states (compact)
                c1, c2 = st.columns(2)
                with c1:
                    df_maker, key_maker, months_maker, _, _ = load_sheet_data("Maker")
                    if df_maker is not None and not df_maker.empty and cols_this:
                        df_maker = df_maker.copy()
                        df_maker["_vol"] = df_maker[cols_this].sum(axis=1)
                        df_maker["_cat"] = df_maker[key_maker].apply(lambda k: get_maker_category_for_key(k, maker_to_cat))
                        seg_order = list(MAKER_CLASSIFICATION.keys())
                        top3_rows = []
                        for seg in seg_order:
                            df_seg = df_maker[df_maker["_cat"] == seg]
                            if df_seg.empty:
                                continue
                            top3 = df_seg.nlargest(3, "_vol")[[key_maker, "_vol"]].copy()
                            top3.columns = ["Maker", "Registrations"]
                            top3["Segment"] = seg
                            top3["Share %"] = (top3["Registrations"] / total_vol * 100).round(1) if total_vol else 0
                            top3_rows.append(top3)
                        if top3_rows:
                            top3_df = pd.concat(top3_rows, ignore_index=True)
                            top3_df = top3_df[["Segment", "Maker", "Registrations", "Share %"]]
                            st.markdown("**Top 3 makers by segment**")
                            st.dataframe(top3_df, use_container_width=True, hide_index=True)
                with c2:
                    state_df, st_key, st_months, _, _ = load_sheet_data("State")
                    if state_df is not None and not state_df.empty and st_months:
                        st_cols = cols_in_range(sort_month_columns(st_months), from_year, from_month, to_year, to_month)
                        if st_cols:
                            state_df = state_df.copy()
                            state_df["_vol"] = state_df[st_cols].sum(axis=1)
                            top5s = state_df.nlargest(5, "_vol")[[st_key, "_vol"]].copy()
                            top5s.columns = ["State", "Registrations"]
                            st.markdown("**Top 5 states**")
                            st.dataframe(top5s, use_container_width=True, hide_index=True)

                # Monthly trend (total registrations)
                last_24 = month_cols_sorted[-24:] if len(month_cols_sorted) >= 24 else month_cols_sorted
                months_parsed = [parse_month_year(c) for c in last_24]
                dates = [pd.Timestamp(y, m, 1) for y, m in months_parsed if (y and m)]
                if dates and len(dates) == len(last_24):
                    monthly_totals = [df_ref[c].sum() for c in last_24]
                    trend_df = pd.DataFrame({"Month–Year": list(last_24), "Date": dates, "Vehicle registrations": monthly_totals})
                    st.subheader("Monthly vehicle registration trend")
                    st.caption("Total vehicle registrations by month (last 24 months).")
                    fig_t = px.line(trend_df, x="Date", y="Vehicle registrations", title="Monthly vehicle registrations")
                    fig_t.update_layout(height=300, margin=dict(t=30), xaxis_title="Month", yaxis_title="Vehicle registrations")
                    st.plotly_chart(fig_t, use_container_width=True, config=PLOTLY_CONFIG)
                    with st.expander("View monthly registration table"):
                        st.dataframe(trend_df[["Month–Year", "Vehicle registrations"]], use_container_width=True, hide_index=True)
        render_footer()

    # ---- Charts tab: all controls in main area ----
    with tab_charts:
        # Single row: Sheet | Chart type (minimal selectors)
        r1, r2 = st.columns([1, 1])
        with r1:
            selected_sheet = st.selectbox("Sheet", sheet_names, format_func=lambda x: labels.get(x, {}).get("title", x), key="charts_sheet")
        sheet_title = labels.get(selected_sheet, {}).get("title", selected_sheet)
        key_label = labels.get(selected_sheet, {}).get("key_label", "Key")

        df, key_col, month_cols_sorted, total_cols, key_label = load_sheet_data(selected_sheet)
        if df is None:
            df = pd.DataFrame()
        else:
            df = df.copy()

        if not month_cols_sorted and not total_cols:
            st.warning("No month or yearly columns in this sheet.")
            if not df.empty:
                st.dataframe(df.head(20), use_container_width=True)
        else:
            suggestions = CHART_SUGGESTIONS.get(selected_sheet, CHART_SUGGESTIONS["Fuel"])
            chart_keys = [s[0] for s in suggestions]
            chart_display = [s[1] for s in suggestions]
            with r2:
                chart_choice = st.selectbox("Chart type", chart_keys, format_func=lambda x: chart_display[chart_keys.index(x)], key="charts_type")

            period_options = get_period_options(month_cols_sorted)
            period_label_for_chart = period_options[-1][0] if period_options else None
            cols_period = period_options[-1][1] if period_options else []
            available_years = get_available_years(month_cols_sorted, total_cols)
            sel_year = available_years[-1] if len(available_years) >= 2 else None

            with st.expander("Options", expanded=False):
                top_n = st.slider("Show top N (0 = all)", 0, 50, 15 if len(df) > 20 else 0, key="charts_topn")
                show_maker_classification = False
                if selected_sheet == "Maker" and month_cols_sorted:
                    show_maker_classification = st.checkbox("Group by segment (PV, 2W, CV, EV, Tractors)", value=False, key="charts_segment")
                if period_options and len(period_options) > 1:
                    period_sel = st.selectbox("Period (for share charts)", [p[0] for p in period_options], index=len(period_options) - 1, key="charts_period")
                    period_label_for_chart = period_sel
                    cols_period = next(p[1] for p in period_options if p[0] == period_sel)
                if chart_choice == "yoy_growth" and len(available_years) >= 2:
                    sel_year = st.selectbox("Year (for YoY)", available_years[1:], index=len(available_years[1:]) - 1, key="yoy_year")

            if show_maker_classification and selected_sheet == "Maker":
                maker_to_cat = get_maker_to_category()
                plot_df = aggregate_by_maker_classification(df, key_col, maker_to_cat, month_cols_sorted)
                if plot_df.empty:
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

            if chart_choice == "treemap_period" and not plot_df.empty and cols_period:
                s = plot_df[cols_period].sum(axis=1)
                s = s[s > 0].sort_values(ascending=False)
                if not s.empty:
                    plot_treemap = s.reset_index()
                    plot_treemap.columns = [key_label_plot, "Registrations"]
                    fig = px.treemap(plot_treemap, path=[key_label_plot], values="Registrations", title=f"{sheet_title} — Share of registrations ({period_label_for_chart})")
                    fig.update_traces(textinfo="label+value+percent root")
                    fig.update_layout(height=500)
                    st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

            elif chart_choice == "stacked_area" and month_cols_sorted and not plot_df.empty:
                long = df_to_time_series(plot_df, month_cols_sorted, key_label_plot)
                if long.columns[0] != key_label_plot:
                    long = long.rename(columns={long.columns[0]: key_label_plot})
                fig = px.area(long, x="Date", y="Registrations", color=key_label_plot, title=f"{sheet_title} — Monthly vehicle registrations (stacked)")
                fig.update_layout(height=500, xaxis_title="Month", yaxis_title="Vehicle registrations", legend_title=key_label_plot, legend=dict(orientation="h", yanchor="bottom", y=1.02))
                st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

            elif chart_choice == "line_multi" and month_cols_sorted and not plot_df.empty:
                long = df_to_time_series(plot_df, month_cols_sorted, key_label_plot)
                if long.columns[0] != key_label_plot:
                    long = long.rename(columns={long.columns[0]: key_label_plot})
                fig = px.line(long, x="Date", y="Registrations", color=key_label_plot, title=f"{sheet_title} — Monthly vehicle registrations")
                fig.update_layout(height=500, xaxis_title="Month", yaxis_title="Vehicle registrations", legend_title=key_label_plot, legend=dict(orientation="h", yanchor="bottom", y=1.02))
                st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

            elif chart_choice == "pie_period" and not plot_df.empty and cols_period:
                s = plot_df[cols_period].sum(axis=1)
                s = s[s > 0].sort_values(ascending=False).head(15)
                if not s.empty:
                    pie_df = s.reset_index()
                    pie_df.columns = [key_label_plot, "Registrations"]
                    fig = px.pie(pie_df, names=key_label_plot, values="Registrations", title=f"{sheet_title} — Share ({period_label_for_chart})")
                    fig.update_layout(height=500)
                    st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

            elif chart_choice == "yearly_stacked" and not plot_df.empty:
                total_cols_in_plot = [c for c in total_cols if c in plot_df.columns]
                if total_cols_in_plot:
                    yearly_df = plot_df[total_cols_in_plot]
                elif month_cols_sorted:
                    years = sorted(set(parse_month_year(c)[0] for c in month_cols_sorted if parse_month_year(c)))
                    yearly_df = pd.DataFrame(index=plot_df.index)
                    for y in years:
                        y_cols = [c for c in month_cols_sorted if parse_month_year(c) and parse_month_year(c)[0] == y]
                        if y_cols:
                            yearly_df[f"Total {y}"] = plot_df[y_cols].sum(axis=1)
                else:
                    yearly_df = pd.DataFrame()
                if not yearly_df.empty:
                    long = yearly_df.reset_index().melt(id_vars=[yearly_df.index.name or key_label_plot], value_vars=yearly_df.columns, var_name="Year", value_name="Registrations")
                    long["Year"] = long["Year"].str.extract(r"(\d{4})", expand=False).astype(int)
                    id_col = [c for c in long.columns if c not in ("Year", "Registrations")][0]
                    if id_col != key_label_plot:
                        long = long.rename(columns={id_col: key_label_plot})
                    fig = px.area(long, x="Year", y="Registrations", color=key_label_plot, title=f"{sheet_title} — Yearly registrations (stacked)")
                    fig.update_layout(height=500, xaxis_title="Year", yaxis_title="Vehicle registrations")
                    st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

            elif chart_choice == "segment_treemap" and selected_sheet == "Maker" and not plot_df.empty and cols_period:
                maker_to_cat = get_maker_to_category()
                df_m = df.set_index(key_col)
                df_m["_cat"] = df_m.index.map(lambda k: get_maker_category_for_key(k, maker_to_cat))
                df_m = df_m[df_m["_cat"].notna()]
                if not df_m.empty:
                    seg_vol = df_m[cols_period].sum(axis=1).groupby(df_m["_cat"]).sum().sort_values(ascending=False)
                    seg_df = seg_vol.reset_index()
                    seg_df.columns = ["Segment", "Registrations"]
                    fig = px.treemap(seg_df, path=["Segment"], values="Registrations", title=f"Maker segment mix — {period_label_for_chart}")
                    fig.update_traces(textinfo="label+value+percent root")
                    fig.update_layout(height=450)
                    st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

            elif chart_choice == "market_share_hbar" and not plot_df.empty and cols_period:
                s = plot_df[cols_period].sum(axis=1)
                total = s.sum()
                if total > 0:
                    share_pct = (s / total * 100).round(1).sort_values(ascending=True).tail(20)
                    share_df = share_pct.reset_index()
                    share_df.columns = [key_label_plot, "Share %"]
                    share_df["Registrations"] = s.reindex(share_df[key_label_plot]).values
                    fig = px.bar(share_df, x="Share %", y=key_label_plot, orientation="h", title=f"{sheet_title} — Market share ({period_label_for_chart})")
                    fig.update_layout(height=500, xaxis_title="Share %", yaxis_title="")
                    st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

            elif chart_choice == "yoy_growth" and not plot_df.empty and month_cols_sorted and len(available_years) >= 2 and sel_year:
                yoy_df = compute_yoy_growth(plot_df, month_cols_sorted, sel_year)
                if not yoy_df.empty:
                    fig = go.Figure()
                    fig.add_trace(go.Scatter(x=yoy_df["Month"], y=yoy_df["YoY_Pct"], mode="lines+markers", name="YoY %", line=dict(color="#2c5282")))
                    fig.add_hline(y=0, line_dash="dash", line_color="gray")
                    fig.update_layout(height=400, title=f"{sheet_title} — YoY % ({sel_year} vs {sel_year-1})", xaxis_title="Month", yaxis_title="YoY %")
                    st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)
                    st.dataframe(yoy_df.rename(columns={"Registrations_This": f"Vol {sel_year}", "Registrations_Last": f"Vol {sel_year-1}", "YoY_Pct": "YoY %"})[["Month", f"Vol {sel_year-1}", f"Vol {sel_year}", "YoY %"]], use_container_width=True, hide_index=True)

            elif chart_choice == "data_table":
                display_df = df.copy()
                for c in (month_cols_sorted or []) + total_cols:
                    if c in display_df.columns:
                        display_df[c] = pd.to_numeric(display_df[c], errors="coerce").fillna(0).astype("int64")
                st.dataframe(display_df, use_container_width=True, height=500)

            if chart_choice != "data_table":
                st.subheader("Data table")
                display_df = df.copy()
                for c in (month_cols_sorted or []) + total_cols:
                    if c in display_df.columns:
                        display_df[c] = pd.to_numeric(display_df[c], errors="coerce").fillna(0).astype("int64")
                st.dataframe(display_df.head(100), use_container_width=True, height=300)
        render_footer()

    # ---- Growth ----
    with tab_trends:
        g1, g2, g3 = st.columns(3)
        with g1:
            sheet_t = st.selectbox("Data", sheet_names, key="trends_sheet")
        df_t, key_t, months_t, totals_t, key_label_t = load_sheet_data(sheet_t)
        if df_t is not None and not df_t.empty and (totals_t or months_t):
            years_t = get_available_years(months_t, totals_t)
            if len(years_t) >= 2:
                with g2:
                    start_y = st.selectbox("Start year", years_t, key="start_y")
                with g3:
                    end_y = st.selectbox("End year", years_t, index=min(years_t.index(max(years_t)), len(years_t) - 1), key="end_y")
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
                                    fig_dual.add_trace(go.Scatter(x=years_axis, y=year_vals, name="Volume", fill="tozeroy"), secondary_y=False)
                                    fig_dual.add_trace(go.Scatter(x=years_axis, y=yoy_pcts, name="YoY %", line=dict(dash="dash")), secondary_y=True)
                                    fig_dual.update_layout(title=f"{sel_entity} — Volume & YoY %", height=400)
                                    fig_dual.update_yaxes(title_text="Volume", secondary_y=False)
                                    fig_dual.update_yaxes(title_text="YoY %", secondary_y=True)
                                    st.plotly_chart(fig_dual, use_container_width=True, config=PLOTLY_CONFIG)
        else:
            st.info("Choose Data and a range of years.")
        render_footer()

    # ---- Rankings ----
    with tab_rankings:
        ra1, ra2 = st.columns(2)
        with ra1:
            sheet_r = st.selectbox("Data", sheet_names, key="rank_sheet")
        df_r, key_r, months_r, totals_r, key_label_r = load_sheet_data(sheet_r)
        if df_r is not None and not df_r.empty and (totals_r or months_r):
            years_r = get_available_years(months_r, totals_r)
            with ra2:
                year_rank = st.selectbox("Year", years_r, index=len(years_r) - 1 if years_r else 0, key="year_rank")
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
            with st.expander("Compare entities over time", expanded=False):
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
                        fig_c = px.line(comp_long, x="Date", y="Registrations", color=key_label_r, title="Registrations over time")
                        fig_c.update_layout(height=400, yaxis_title="Vehicle registrations")
                        st.plotly_chart(fig_c, use_container_width=True, config=PLOTLY_CONFIG)
        else:
            st.info("Choose Data with yearly/monthly columns.")
        render_footer()


if __name__ == "__main__":
    run_dashboard()
