import os
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="출고 이력 검색", layout="wide")

st.markdown("""
<style>
.block-container {
    padding-top: 0.65rem;
    padding-bottom: 1rem;
    padding-left: 0.7rem;
    padding-right: 0.7rem;
    max-width: 100%;
}
div[data-testid="stHorizontalBlock"] {
    gap: 0.6rem;
    align-items: flex-end !important;
}
div[data-testid="column"] > div {
    width: 100% !important;
}
label[data-testid="stWidgetLabel"] {
    margin-bottom: 0.18rem !important;
    padding-bottom: 0 !important;
}
label[data-testid="stWidgetLabel"] p {
    margin: 0 !important;
    font-size: 0.78rem !important;
    line-height: 1.1 !important;
    color: #4b5563 !important;
    font-weight: 600 !important;
}
div[data-testid="stMetric"] {
    background: #fafafa;
    border: 1px solid #eeeeee;
    border-radius: 10px;
    padding: 0.28rem 0.62rem !important;
    min-height: 56px !important;
    height: 56px !important;
    display: flex;
    flex-direction: column;
    justify-content: center;
    box-sizing: border-box;
}
div[data-testid="stMetric"] > div {
    justify-content: center !important;
}
div[data-testid="stMetricLabel"] {
    font-size: 0.70rem !important;
    line-height: 1.0 !important;
    margin-bottom: 0.08rem !important;
    color: #6b7280 !important;
}
div[data-testid="stMetricValue"] {
    font-size: 0.98rem !important;
    line-height: 1.05 !important;
    color: #111827 !important;
}
div[data-baseweb="select"] {
    min-height: 56px !important;
}
div[data-baseweb="select"] > div {
    min-height: 56px !important;
    height: 56px !important;
    display: flex !important;
    align-items: center !important;
    box-sizing: border-box !important;
    border-radius: 10px !important;
    padding-top: 0 !important;
    padding-bottom: 0 !important;
}
div[data-baseweb="select"] div[class*="valueContainer"] {
    min-height: 56px !important;
    height: 56px !important;
    display: flex !important;
    align-items: center !important;
    padding-top: 0 !important;
    padding-bottom: 0 !important;
}
div[data-baseweb="select"] div[class*="singleValue"] {
    display: flex !important;
    align-items: center !important;
    color: #111827 !important;
    font-size: 0.95rem !important;
    line-height: 1.25 !important;
    margin: 0 !important;
}
div[data-baseweb="select"] input {
    color: #111827 !important;
    line-height: 1.25 !important;
    font-size: 0.95rem !important;
    margin: 0 !important;
    padding: 0 !important;
}
div[data-baseweb="select"] input::placeholder {
    color: #6b7280 !important;
    opacity: 1 !important;
}
div[data-baseweb="select"] div[class*="placeholder"] {
    color: #6b7280 !important;
    opacity: 1 !important;
    font-size: 0.95rem !important;
    line-height: 1.25 !important;
    margin: 0 !important;
    display: flex !important;
    align-items: center !important;
}
div[data-baseweb="tag"] {
    display: flex !important;
    align-items: center !important;
    margin-top: 0 !important;
    margin-bottom: 0 !important;
}
section[data-testid="stSidebar"] label[data-testid="stWidgetLabel"] p {
    color: #374151 !important;
    font-weight: 700 !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
    background: #ffffff !important;
    border: 1px solid #e5e7eb !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] div[class*="valueContainer"] {
    min-height: 56px !important;
    height: 56px !important;
    display: flex !important;
    align-items: center !important;
    padding-top: 0 !important;
    padding-bottom: 0 !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] div[class*="singleValue"] {
    color: #111827 !important;
    font-size: 0.96rem !important;
    line-height: 1.25 !important;
    display: flex !important;
    align-items: center !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] div[class*="placeholder"] {
    color: #6b7280 !important;
    opacity: 1 !important;
    font-size: 0.96rem !important;
    line-height: 1.25 !important;
    display: flex !important;
    align-items: center !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] input {
    color: #111827 !important;
    font-size: 0.96rem !important;
    line-height: 1.25 !important;
    padding: 0 !important;
    margin: 0 !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] input::placeholder {
    color: #6b7280 !important;
    opacity: 1 !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] * {
    overflow: visible !important;
}
div[data-testid="stDataFrame"] {
    width: 100%;
}
[data-testid="column"] {
    width: 100% !important;
}
.app-main-title {
    display: block;
    font-size: 1.8rem;
    font-weight: 800;
    line-height: 1.2;
    margin: 0.1rem 0 0.45rem 0;
    padding: 0;
    color: #1f2937;
}
.section-title {
    font-size: 1.45rem;
    font-weight: 700;
    margin-bottom: 0.2rem;
}
@media (max-width: 1024px) {
    .block-container {
        padding-left: 0.45rem;
        padding-right: 0.45rem;
        padding-top: 0.4rem;
    }
    html, body, [class*="css"] {
        font-size: 14px;
    }
    .app-main-title {
        font-size: 1.35rem;
        line-height: 1.15;
        margin-top: 0.1rem;
    }
    .section-title {
        font-size: 1.2rem;
    }
    div[data-testid="stMetric"] {
        min-height: 52px !important;
        height: 52px !important;
        padding: 0.22rem 0.5rem !important;
    }
    div[data-testid="stMetricLabel"] {
        font-size: 0.66rem !important;
    }
    div[data-testid="stMetricValue"] {
        font-size: 0.88rem !important;
    }
    div[data-baseweb="select"] {
        min-height: 52px !important;
    }
    div[data-baseweb="select"] > div {
        min-height: 52px !important;
        height: 52px !important;
    }
    div[data-baseweb="select"] div[class*="valueContainer"] {
        min-height: 52px !important;
        height: 52px !important;
    }
    section[data-testid="stSidebar"] div[data-baseweb="select"] div[class*="singleValue"],
    section[data-testid="stSidebar"] div[data-baseweb="select"] div[class*="placeholder"],
    section[data-testid="stSidebar"] div[data-baseweb="select"] input {
        font-size: 0.90rem !important;
    }
}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# 공통 유틸
# ══════════════════════════════════════════════════════════
def safe_numeric(series, default=0):
    s = pd.to_numeric(series, errors="coerce")
    return s.replace([np.inf, -np.inf], np.nan).fillna(default)


def safe_str(series):
    return pd.Series(series).fillna("").astype(str).str.strip()


def safe_date(series):
    return pd.to_datetime(series, errors="coerce")


def calc_table_height(df, min_rows=3, max_rows=18, row_px=35, header_px=38):
    if df is None or len(df) == 0:
        rows = min_rows
    else:
        rows = max(min_rows, min(len(df), max_rows))
    return header_px + rows * row_px


def safe_make_product_label(df):
    temp = df.copy()
    if "품목코드" not in temp.columns:
        temp["품목코드"] = ""
    if "품목명(공식)" not in temp.columns:
        temp["품목명(공식)"] = ""

    temp["품목코드"] = temp["품목코드"].astype(str)
    temp["품목명(공식)"] = temp["품목명(공식)"].fillna("").astype(str)

    temp["품목표시"] = np.where(
        temp["품목명(공식)"].str.strip() != "",
        temp["품목코드"] + " | " + temp["품목명(공식)"],
        temp["품목코드"]
    )
    return temp


def clean_and_safe_display(
    df,
    height=None,
    key=None,
    editable=False,
    pinned_cols=None,
    text_cols=None,
    disabled_cols=None,
):
    if df is None:
        df = pd.DataFrame()

    display_df = df.copy().reset_index(drop=True)
    display_df.columns = [str(c) for c in display_df.columns]

    pinned_cols = [str(c) for c in (pinned_cols or [])]
    text_cols = set(str(c) for c in (text_cols or []))
    disabled_cols = disabled_cols if disabled_cols is not None else False

    if display_df.empty:
        empty_height = calc_table_height(display_df)
        if editable and key:
            return st.data_editor(
                display_df,
                width="stretch",
                hide_index=True,
                key=key,
                disabled=disabled_cols,
                height=empty_height,
            )
        st.dataframe(display_df, width="stretch", hide_index=True, height=empty_height)
        return None

    for col in display_df.columns:
        s = display_df[col]

        if pd.api.types.is_datetime64_any_dtype(s):
            display_df[col] = pd.to_datetime(s, errors="coerce").dt.strftime("%Y-%m-%d")
            display_df[col] = display_df[col].fillna("")
            continue

        if pd.api.types.is_numeric_dtype(s) and col not in text_cols:
            display_df[col] = pd.to_numeric(s, errors="coerce")
            display_df[col] = display_df[col].replace([np.inf, -np.inf], np.nan).fillna(0)
            continue

        display_df[col] = s.astype(str)
        display_df[col] = display_df[col].replace(
            ["nan", "NaN", "None", "<NA>", "NaT"], ""
        )
        display_df[col] = display_df[col].fillna("")

    column_config = {}
    fixed_text_like_cols = {
        "거래처", "품목코드", "품목명(공식)", "품목표시", "점착제코드", "점착제명",
        "최근날짜", "가로폭이력", "분석_내역", "AI분석", "업체성향", "최근추세",
        "월", "구분", "월표기", "연도", "대표구분", "반품원인", "주요반품원인",
        "원인추정", "감소원인", "비고", "AI반품분석", "담당부서", "영업담당부서", "담당자",
        "분석요약", "추천자료"
    }

    for col in display_df.columns:
        pinned = col in pinned_cols

        if col in text_cols or col in fixed_text_like_cols:
            column_config[col] = st.column_config.TextColumn(
                col,
                width="large" if col in [
                    "품목명(공식)", "품목표시", "가로폭이력", "분석_내역",
                    "AI분석", "원인추정", "감소원인", "비고", "주요반품원인",
                    "AI반품분석", "분석요약", "추천자료"
                ] else "medium",
                pinned=pinned,
            )
            continue

        if pd.api.types.is_numeric_dtype(display_df[col]):
            if any(k in col for k in ["하락률", "증감률", "비율", "변화율", "CV", "반품율"]):
                column_config[col] = st.column_config.NumberColumn(col, format="%,.1f", pinned=pinned)
            elif any(k in col for k in ["M2", "수량", "판매량", "총량", "출고량"]):
                column_config[col] = st.column_config.NumberColumn(col, format="%,.1f", pinned=pinned)
            elif any(k in col for k in ["단가", "금액", "매출", "총매출", "평균", "원"]):
                column_config[col] = st.column_config.NumberColumn(col, format="%,.0f", pinned=pinned)
            elif any(k in col for k in ["점수", "AI", "우선순위", "통계", "종합"]):
                column_config[col] = st.column_config.NumberColumn(col, format="%,.1f", pinned=pinned)
            else:
                column_config[col] = st.column_config.NumberColumn(col, format="%,.0f", pinned=pinned)
        else:
            column_config[col] = st.column_config.TextColumn(col, width="medium", pinned=pinned)

    final_height = height if height is not None else calc_table_height(display_df)

    if editable and key:
        safe_editor_df = display_df.copy()

        for col in safe_editor_df.columns:
            if col in text_cols or col in fixed_text_like_cols:
                safe_editor_df[col] = safe_editor_df[col].astype(str).fillna("")
            else:
                tmp = pd.to_numeric(safe_editor_df[col], errors="coerce")
                if tmp.notna().sum() > 0:
                    safe_editor_df[col] = tmp.fillna(0)
                else:
                    safe_editor_df[col] = safe_editor_df[col].astype(str).fillna("")

        return st.data_editor(
            safe_editor_df,
            column_config=column_config,
            width="stretch",
            height=final_height,
            num_rows="fixed",
            key=key,
            hide_index=True,
            disabled=disabled_cols,
        )

    st.dataframe(
        display_df,
        column_config=column_config,
        width="stretch",
        height=final_height,
        hide_index=True,
    )
    return None


def render_banded_table(df, title=None, height=None, pinned_cols=None, text_cols=None):
    if title:
        st.markdown(title)

    if df is None or df.empty:
        clean_and_safe_display(
            pd.DataFrame(),
            height=height,
            pinned_cols=pinned_cols,
            text_cols=text_cols,
        )
        return

    temp = df.copy().reset_index(drop=True)
    temp.columns = [str(c) for c in temp.columns]

    if "품목코드" not in temp.columns:
        clean_and_safe_display(
            temp,
            height=height,
            pinned_cols=pinned_cols,
            text_cols=text_cols,
        )
        return

    item_order = temp["품목코드"].astype(str).fillna("").drop_duplicates().tolist()
    shade_map = {
        item: ("background-color: rgba(0,0,0,0.04);" if i % 2 == 1 else "")
        for i, item in enumerate(item_order)
    }

    def _row_style(row):
        key = str(row.get("품목코드", ""))
        return [shade_map.get(key, "")] * len(row)

    styled = temp.style.apply(_row_style, axis=1)

    num_format = {}
    for c in temp.columns:
        if any(k in c for k in ["M2", "수량", "판매량", "총량", "출고량"]):
            num_format[c] = "{:,.1f}"
        elif any(k in c for k in ["단가", "금액", "매출", "총매출", "평균", "원"]):
            num_format[c] = "{:,.0f}"
        elif any(k in c for k in ["하락률", "증감률", "비율", "변화율", "CV", "점수", "반품율"]):
            num_format[c] = "{:,.1f}"

    if num_format:
        styled = styled.format(num_format)

    final_height = height if height is not None else calc_table_height(temp, max_rows=20)
    st.dataframe(styled, width="stretch", height=final_height, hide_index=True)


def sorted_unique(series):
    if series is None:
        return []

    s = pd.Series(series)
    s = s.dropna()
    s = s.astype(str).str.strip()
    s = s[~s.isin(["", "0", "0.0", "nan", "NaN", "None", "<NA>", "NaT"])]

    vals = s.tolist()
    vals = list(dict.fromkeys(vals))
    return sorted(vals, key=lambda x: str(x))


def sorted_unique_safe(series):
    if series is None:
        return []

    s = pd.Series(series)
    s = s.dropna()
    s = s.astype(str).str.strip()
    s = s[~s.isin(["", "0", "0.0", "nan", "NaN", "None", "<NA>", "NaT"])]

    vals = s.tolist()
    vals = list(dict.fromkeys(vals))
    return sorted(vals, key=lambda x: str(x))


def add_year_month_axis(fig, x_dates):
    dt = pd.to_datetime(pd.Series(x_dates), errors="coerce").dropna().sort_values().unique()
    if len(dt) == 0:
        return fig

    dt = pd.to_datetime(dt)
    tickvals = list(dt)
    ticktext = [pd.Timestamp(d).strftime("%m") for d in dt]

    fig.update_xaxes(
        tickmode="array",
        tickvals=tickvals,
        ticktext=ticktext,
        tickangle=0,
        title="",
        showgrid=True,
        zeroline=False,
        automargin=True
    )

    years = sorted(pd.Series(dt).dt.year.unique())
    for y in years:
        year_dates = [d for d in dt if pd.Timestamp(d).year == y]
        if not year_dates:
            continue
        mid = year_dates[len(year_dates) // 2]
        fig.add_annotation(
            x=mid,
            y=-0.20,
            xref="x",
            yref="paper",
            text=str(y),
            showarrow=False,
            font=dict(size=12, color="black"),
        )
        fig.add_vline(
            x=year_dates[0],
            line_width=1,
            line_dash="dot",
            line_color="rgba(120,120,120,0.5)"
        )

    fig.update_layout(margin=dict(l=20, r=40, t=35, b=90))
    return fig


def apply_mobile_friendly_line_layout(fig, x_dates, y_title="판매금액(원)", height=380):
    x_dates = pd.to_datetime(pd.Series(x_dates), errors="coerce").dropna().sort_values()
    fig.update_traces(cliponaxis=False)

    if len(x_dates) > 0:
        x_min = x_dates.min() - pd.Timedelta(days=10)
        x_max = x_dates.max() + pd.Timedelta(days=12)
    else:
        x_min = None
        x_max = None

    fig.update_layout(
        height=height,
        yaxis_title=y_title,
        yaxis_tickformat=",",
        margin=dict(l=20, r=45, t=35, b=90),
        xaxis=dict(
            automargin=True,
            range=[x_min, x_max] if x_min is not None and x_max is not None else None,
            fixedrange=False
        ),
        yaxis=dict(automargin=True),
    )
    return add_year_month_axis(fig, x_dates)


def build_month_axis_frame(months):
    try:
        month_list = sorted([str(m) for m in months if pd.notna(m) and str(m).strip() != ""])
        axis_df = pd.DataFrame({"월": month_list})
        if axis_df.empty:
            axis_df["날짜축"] = pd.NaT
            return axis_df
        axis_df["날짜축"] = pd.to_datetime(axis_df["월"] + "-01", errors="coerce")
        axis_df = axis_df.dropna(subset=["날짜축"]).sort_values("날짜축").reset_index(drop=True)
        return axis_df
    except Exception:
        return pd.DataFrame(columns=["월", "날짜축"])


def align_monthly_series(base_month_df, data_df, value_col):
    try:
        if base_month_df is None or base_month_df.empty:
            return pd.DataFrame(columns=["월", "날짜축", value_col])

        out = base_month_df.copy()

        if data_df is None or data_df.empty or value_col not in data_df.columns or "월" not in data_df.columns:
            out[value_col] = 0
            return out

        temp = data_df.copy()
        temp["월"] = temp["월"].astype(str)
        temp[value_col] = pd.to_numeric(temp[value_col], errors="coerce").fillna(0)
        temp = temp.groupby("월", as_index=False)[value_col].sum()
        out = out.merge(temp, on="월", how="left")
        out[value_col] = pd.to_numeric(out[value_col], errors="coerce").fillna(0)
        return out
    except Exception:
        out = base_month_df.copy() if base_month_df is not None else pd.DataFrame(columns=["월", "날짜축"])
        out[value_col] = 0
        return out


def sales_to_manwon_label(value):
    try:
        if pd.isna(value):
            return ""
        return f"{int(round(float(value) / 10000.0, 0)):,}"
    except Exception:
        return ""


def make_text_position_map(items):
    positions = ["top center", "bottom center", "middle right", "middle left"]
    return {str(item): positions[i % len(positions)] for i, item in enumerate(items)}


def make_indexed_series(df, group_col, value_col, time_col):
    if df is None or df.empty:
        return pd.DataFrame()

    temp = df.copy()
    required = {group_col, value_col, time_col}
    if not required.issubset(set(temp.columns)):
        return pd.DataFrame()

    temp = temp.sort_values([group_col, time_col])

    def _calc(g):
        g = g.copy()
        vals = pd.to_numeric(g[value_col], errors="coerce")
        base_candidates = vals.replace(0, np.nan).dropna()
        if len(base_candidates) == 0:
            g["지수값"] = np.nan
        else:
            base = base_candidates.iloc[0]
            if pd.isna(base) or base == 0:
                g["지수값"] = np.nan
            else:
                g["지수값"] = (vals / base) * 100
        return g

    try:
        return temp.groupby(group_col, group_keys=False).apply(_calc)
    except Exception:
        return pd.DataFrame()


def should_show_helper_chart(top_product_monthly):
    try:
        if top_product_monthly is None or top_product_monthly.empty:
            return False, 0.0

        if "품목표시" not in top_product_monthly.columns or "금액(원)" not in top_product_monthly.columns:
            return False, 0.0

        temp = (
            top_product_monthly.groupby("품목표시", as_index=False)["금액(원)"]
            .sum()
            .rename(columns={"금액(원)": "총매출"})
        )
        temp["총매출"] = pd.to_numeric(temp["총매출"], errors="coerce").fillna(0)
        temp = temp[temp["총매출"] > 0].copy()

        if len(temp) < 2:
            return False, 0.0

        max_v = temp["총매출"].max()
        min_v = temp["총매출"].min()

        if min_v <= 0:
            return True, np.inf

        ratio = max_v / min_v
        return bool(ratio >= 5), float(ratio)
    except Exception:
        return False, 0.0


def calc_cv(series):
    try:
        s = pd.to_numeric(pd.Series(series), errors="coerce").replace([np.inf, -np.inf], np.nan).dropna()
        if len(s) == 0:
            return 0.0
        m = s.mean()
        if m == 0 or pd.isna(m):
            return 0.0
        std = s.std(ddof=0)
        if pd.isna(std):
            return 0.0
        return float(std / m)
    except Exception:
        return 0.0


def calc_slope(values):
    try:
        y = pd.to_numeric(pd.Series(values), errors="coerce").replace([np.inf, -np.inf], np.nan).fillna(0).values.astype(float)
        if len(y) < 2:
            return 0.0
        if np.all(np.isclose(y, y[0], equal_nan=True)):
            return 0.0
        x = np.arange(len(y))
        coef = np.polyfit(x, y, 1)
        if len(coef) < 1 or pd.isna(coef[0]):
            return 0.0
        return float(coef[0])
    except Exception:
        return 0.0


def scale_to_100(series, reverse=False):
    s = pd.to_numeric(pd.Series(series), errors="coerce").fillna(0)
    if len(s) == 0:
        return s
    if s.nunique() <= 1:
        return pd.Series([50.0] * len(s), index=s.index)
    mn, mx = s.min(), s.max()
    if mx == mn:
        out = pd.Series([50.0] * len(s), index=s.index)
    else:
        out = ((s - mn) / (mx - mn)) * 100.0
    if reverse:
        out = 100.0 - out
    return out.clip(0, 100)


def build_color_map(items):
    palette = [
        "#1f77b4", "#d62728", "#2ca02c", "#9467bd", "#ff7f0e",
        "#17becf", "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22"
    ]
    items = [str(x) for x in items]
    return {item: palette[i % len(palette)] for i, item in enumerate(items)}


def build_return_base(df):
    temp = df.copy()
    if "비고" not in temp.columns:
        temp["비고"] = ""

    temp["금액(원)"] = safe_numeric(temp["금액(원)"] if "금액(원)" in temp.columns else 0, default=0)
    if "수량(M2)" in temp.columns:
        temp["수량(M2)"] = safe_numeric(temp["수량(M2)"], default=0)
    else:
        temp["수량(M2)"] = 0

    if "거래처" in temp.columns:
        temp["거래처"] = safe_str(temp["거래처"])

    temp["반품여부_표준"] = temp["금액(원)"] < 0
    temp["반품금액_표준"] = np.where(temp["반품여부_표준"], temp["금액(원)"].abs(), 0)
    temp["반품수량_표준"] = np.where(temp["반품여부_표준"], temp["수량(M2)"].abs(), 0)
    temp["반품원인_표준"] = np.where(
        temp["반품여부_표준"],
        temp["비고"].fillna("").astype(str),
        ""
    )
    return temp


def summarize_return_reason_text(texts):
    vals = [str(x).strip() for x in texts if str(x).strip() not in ["", "nan", "None"]]
    if len(vals) == 0:
        return "반품 사유 기재 없음"

    s = pd.Series(vals)
    top = s.value_counts().head(3)
    return " / ".join([f"{idx}({cnt})" for idx, cnt in top.items()])


def infer_customer_item_decline_reason(row):
    reasons = []

    if row.get("감소금액", 0) > 0:
        reasons.append(f"감소금액 {int(row.get('감소금액', 0)):,}원")
    if row.get("하락률(%)", 0) >= 30:
        reasons.append("하락률 큼")
    if row.get("반품금액", 0) > 0:
        reasons.append("반품 발생")
    if row.get("반품율(%)", 0) >= 5:
        reasons.append("반품 비중 높음")
    if str(row.get("주요반품원인", "")).strip() not in ["", "반품 없음", "반품 사유 기재 없음"]:
        reasons.append("반품 사유 확인 필요")
    if row.get("최근기울기", 0) < 0:
        reasons.append("최근 추세 하락")

    if len(reasons) == 0:
        reasons.append("일반 매출 감소 가능성")

    return " | ".join(reasons)


def infer_ai_return_analysis(row):
    reasons = []
    return_amt = row.get("반품금액", 0)
    return_rate = row.get("반품율(%)", 0)
    reason_text = str(row.get("주요반품원인", "")).strip()

    if return_amt > 0:
        reasons.append(f"반품금액 {int(return_amt):,}원 발생")
    if return_rate >= 5:
        reasons.append(f"반품율 {return_rate:.1f}%로 높음")
    elif return_amt > 0:
        reasons.append(f"반품율 {return_rate:.1f}% 수준")

    lower_reason = reason_text.lower()
    if reason_text not in ["", "반품 없음", "반품 사유 기재 없음"]:
        if any(k in lower_reason for k in ["불량", "기포", "들뜸", "점착", "미부착", "번짐", "오염", "스크래치"]):
            reasons.append("품질/점착 이슈 가능성")
        elif any(k in lower_reason for k in ["취소", "오주문", "변경", "납기", "지연"]):
            reasons.append("주문/납기 이슈 가능성")
        else:
            reasons.append("비고 사유 검토 필요")

    if len(reasons) == 0:
        return "반품 이슈 없음"
    return " | ".join(reasons)


@st.cache_data
def load_excel(file_bytes):
    xls = pd.ExcelFile(BytesIO(file_bytes))

    def read_sheet(name, required=True):
        try:
            return pd.read_excel(xls, name)
        except Exception:
            if required:
                st.error(f"필수 시트가 없습니다: {name}")
                st.stop()
            return pd.DataFrame()

    rec = read_sheet("출고기록")
    alias = read_sheet("별칭맵핑")
    prod = read_sheet("품목마스터", required=False)
    adh = read_sheet("점착제마스터", required=False)
    cust = read_sheet("거래처마스터", required=False)

    if rec.empty:
        st.error("출고기록 시트가 비어있습니다.")
        st.stop()

    rec.columns = [str(c).strip() for c in rec.columns]

    essential_defaults = {
        "거래처": "",
        "품목코드": "",
        "점착제코드": "",
        "비고": "",
    }
    for c, default_val in essential_defaults.items():
        if c not in rec.columns:
            rec[c] = default_val

    if "날짜" in rec.columns:
        rec["날짜"] = pd.to_datetime(rec["날짜"], errors="coerce")
    else:
        rec["날짜"] = pd.NaT

    for c in ["가로폭(mm)", "수량(M2)", "단가(원/M2)", "금액(원)"]:
        if c not in rec.columns:
            rec[c] = 0
        rec[c] = pd.to_numeric(rec[c], errors="coerce").replace([np.inf, -np.inf], np.nan)

    if "금액(원)" not in rec.columns and {"수량(M2)", "단가(원/M2)"} <= set(rec.columns):
        rec["금액(원)"] = (safe_numeric(rec["수량(M2)"]) * safe_numeric(rec["단가(원/M2)"])).round(0)

    if "비고" not in rec.columns:
        rec["비고"] = ""

    def normalize(df, col):
        if col in df.columns:
            s = df[col].astype(str).str.strip()
            df[col] = s.replace({
                "": pd.NA, "0": pd.NA, "0.0": pd.NA,
                "nan": pd.NA, "NaN": pd.NA, "None": pd.NA, "<NA>": pd.NA,
            })

    for c in ["거래처", "품목코드", "점착제코드", "품목명_고객표현", "점착제_고객표현", "담당부서", "영업담당부서", "담당자"]:
        if c in rec.columns:
            rec[c] = rec[c].astype(str).str.strip()

    for c in ["품목코드", "점착제코드"]:
        normalize(rec, c)
    normalize(rec, "거래처")
    if "담당부서" in rec.columns:
        normalize(rec, "담당부서")
    if "영업담당부서" in rec.columns:
        normalize(rec, "영업담당부서")
    if "담당자" in rec.columns:
        normalize(rec, "담당자")

    def map_alias(series, alias_df, typ):
        if alias_df is None or alias_df.empty or series is None:
            return pd.Series([None] * len(series), index=series.index)

        df = alias_df.copy()
        if not {"유형", "별칭", "공식코드"} <= set(df.columns):
            return pd.Series([None] * len(series), index=series.index)

        df = df[df["유형"].astype(str) == typ].dropna(subset=["별칭", "공식코드"])
        al = df["별칭"].astype(str).str.strip().tolist()
        cl = df["공식코드"].astype(str).str.strip().tolist()

        out = []
        for v in series.astype(str).fillna("").str.strip():
            m = next((c for a, c in zip(al, cl) if v == a), None)
            if m is None:
                m = next((c for a, c in zip(al, cl) if a and a in v), None)
            out.append(m)
        return pd.Series(out, index=series.index)

    for col, typ in [("품목코드", "품목"), ("점착제코드", "점착제")]:
        src = "품목명_고객표현" if typ == "품목" else "점착제_고객표현"
        if col not in rec.columns or rec[col].isna().any():
            if col not in rec.columns:
                rec[col] = pd.NA
            if src in rec.columns:
                rec[col] = rec[col].fillna(map_alias(rec[src], alias, typ))

    for c in ["품목코드", "점착제코드"]:
        normalize(rec, c)

    if not prod.empty and {"품목코드", "품목명(공식)"} <= set(prod.columns):
        prod_cols = ["품목코드", "품목명(공식)"]
        if "품목비고" in prod.columns:
            prod_cols.append("품목비고")
        rec = rec.merge(prod[prod_cols].drop_duplicates(), on="품목코드", how="left")

    if not adh.empty and {"점착제코드", "점착제명"} <= set(adh.columns):
        rec = rec.merge(
            adh[["점착제코드", "점착제명"]].drop_duplicates(),
            on="점착제코드",
            how="left",
        )

    tmp = rec.copy()
    tmp["_cs"] = tmp["거래처"].astype(str).fillna("") if "거래처" in tmp.columns else ""
    tmp["_ps"] = tmp["품목코드"].astype(str).fillna("") if "품목코드" in tmp.columns else ""
    sc = ["_cs", "_ps"] + (["날짜"] if "날짜" in tmp.columns else [])
    tmp = tmp.sort_values(sc, kind="mergesort")

    if "단가(원/M2)" in tmp.columns and {"거래처", "품목코드"} <= set(tmp.columns):
        _t = tmp.dropna(subset=["단가(원/M2)"]).groupby(
            ["거래처", "품목코드"], as_index=False, dropna=False
        ).tail(1)
        _ec = ["거래처", "품목코드", "단가(원/M2)"] + (["날짜"] if "날짜" in _t.columns else [])
        _rm = {"단가(원/M2)": "최근단가"}
        if "날짜" in _t.columns:
            _rm["날짜"] = "최근날짜"
        lp = _t[_ec].rename(columns=_rm)
        if "최근날짜" in lp.columns:
            lp["최근날짜"] = pd.to_datetime(lp["최근날짜"], errors="coerce").dt.strftime("%Y-%m-%d")
        rec = rec.merge(lp, on=["거래처", "품목코드"], how="left")

    def join_unique(s):
        vals = []
        for x in pd.unique(s.dropna()):
            try:
                xf = float(x)
                vals.append(str(int(xf)) if xf.is_integer() else str(xf))
            except Exception:
                vals.append(str(x))
        return ", ".join(vals)

    if "가로폭(mm)" in rec.columns and {"거래처", "품목코드"} <= set(rec.columns):
        wh = (
            rec.groupby(["거래처", "품목코드"], dropna=False)["가로폭(mm)"]
            .apply(join_unique)
            .reset_index()
            .rename(columns={"가로폭(mm)": "가로폭이력"})
        )
        rec = rec.merge(wh, on=["거래처", "품목코드"], how="left")
    else:
        rec["가로폭이력"] = pd.NA

    return rec, alias, prod, adh, cust


@st.cache_data(show_spinner=False)
def build_customer_sales_analysis(q):
    if q is None or q.empty:
        return {
            "df": pd.DataFrame(),
            "all_months": [],
            "customer_summary": pd.DataFrame(),
            "customer_monthly": pd.DataFrame(),
            "customer_item_monthly": pd.DataFrame(),
            "customer_item_summary": pd.DataFrame(),
        }

    df = q.copy()
    if "날짜" not in df.columns:
        df["날짜"] = pd.NaT
    df["날짜"] = pd.to_datetime(df["날짜"], errors="coerce")
    df = df[df["날짜"].notna()].copy()
    if df.empty:
        return {
            "df": pd.DataFrame(),
            "all_months": [],
            "customer_summary": pd.DataFrame(),
            "customer_monthly": pd.DataFrame(),
            "customer_item_monthly": pd.DataFrame(),
            "customer_item_summary": pd.DataFrame(),
        }

    df["월"] = df["날짜"].dt.strftime("%Y-%m")
    df = safe_make_product_label(df)

    for c in ["금액(원)", "수량(M2)", "단가(원/M2)"]:
        if c not in df.columns:
            df[c] = 0
        df[c] = safe_numeric(df[c], default=0)

    if "가로폭(mm)" not in df.columns:
        df["가로폭(mm)"] = np.nan
    if "거래처" not in df.columns:
        df["거래처"] = ""
    if "품목코드" not in df.columns:
        df["품목코드"] = ""

    customer_monthly = (
        df.groupby(["거래처", "월"], dropna=False)
        .agg(
            매출액=("금액(원)", "sum"),
            판매량=("수량(M2)", "sum"),
        )
        .reset_index()
    )
    customer_monthly["날짜축"] = pd.to_datetime(customer_monthly["월"] + "-01", errors="coerce")

    customer_item_monthly = (
        df.groupby(["거래처", "품목코드", "품목표시", "월"], dropna=False)
        .agg(
            매출액=("금액(원)", "sum"),
            판매량=("수량(M2)", "sum"),
        )
        .reset_index()
    )
    customer_item_monthly["날짜축"] = pd.to_datetime(customer_item_monthly["월"] + "-01", errors="coerce")

    item_latest = (
        df.sort_values("날짜")
        .groupby(["거래처", "품목코드", "품목표시"], as_index=False)
        .tail(1)[["거래처", "품목코드", "품목표시", "단가(원/M2)", "날짜"]]
        .rename(columns={"단가(원/M2)": "최근단가", "날짜": "최근날짜"})
    )
    item_latest["최근날짜"] = pd.to_datetime(item_latest["최근날짜"], errors="coerce").dt.strftime("%Y-%m-%d")

    def join_unique_width(s):
        vals = []
        for x in pd.unique(pd.Series(s).dropna()):
            try:
                xf = float(x)
                vals.append(str(int(xf)) if xf.is_integer() else str(xf))
            except Exception:
                vals.append(str(x))
        return ", ".join(vals)

    width_hist = (
        df.groupby(["거래처", "품목코드", "품목표시"], dropna=False)["가로폭(mm)"]
        .apply(join_unique_width)
        .reset_index()
        .rename(columns={"가로폭(mm)": "가로폭이력"})
    )

    customer_item_summary = (
        df.groupby(["거래처", "품목코드", "품목표시"], dropna=False)
        .agg(
            출고횟수=("수량(M2)", "count"),
            총판매량=("수량(M2)", "sum"),
            총매출액=("금액(원)", "sum"),
            평균단가=("단가(원/M2)", "mean"),
            최근출고일=("날짜", "max"),
            월수=("월", "nunique"),
        )
        .reset_index()
    )

    customer_item_summary["월평균_판매량"] = np.where(
        customer_item_summary["월수"] > 0,
        customer_item_summary["총판매량"] / customer_item_summary["월수"],
        0
    )
    customer_item_summary["월평균_매출"] = np.where(
        customer_item_summary["월수"] > 0,
        customer_item_summary["총매출액"] / customer_item_summary["월수"],
        0
    )
    customer_item_summary["최근출고일"] = pd.to_datetime(
        customer_item_summary["최근출고일"], errors="coerce"
    ).dt.strftime("%Y-%m-%d")

    customer_item_summary = customer_item_summary.merge(
        item_latest[["거래처", "품목코드", "품목표시", "최근단가", "최근날짜"]],
        on=["거래처", "품목코드", "품목표시"],
        how="left"
    )
    customer_item_summary = customer_item_summary.merge(
        width_hist,
        on=["거래처", "품목코드", "품목표시"],
        how="left"
    )

    total_by_customer = (
        df.groupby("거래처", dropna=False)
        .agg(
            총매출액=("금액(원)", "sum"),
            총판매량=("수량(M2)", "sum"),
            거래개월수=("월", "nunique"),
            품목수=("품목코드", "nunique"),
            출고건수=("수량(M2)", "count"),
            최근출고일=("날짜", "max"),
        )
        .reset_index()
    )
    total_by_customer["최근출고일"] = pd.to_datetime(total_by_customer["최근출고일"], errors="coerce").dt.strftime("%Y-%m-%d")
    total_by_customer["월평균_매출"] = np.where(
        total_by_customer["거래개월수"] > 0,
        total_by_customer["총매출액"] / total_by_customer["거래개월수"],
        0
    )
    total_by_customer["월평균_판매량"] = np.where(
        total_by_customer["거래개월수"] > 0,
        total_by_customer["총판매량"] / total_by_customer["거래개월수"],
        0
    )

    slope_rows = []
    for cust_name, g in customer_monthly.groupby("거래처"):
        g = g.sort_values("월")
        slope_rows.append({
            "거래처": cust_name,
            "매출기울기": calc_slope(g["매출액"].tolist()),
            "매출CV": calc_cv(g["매출액"]),
        })
    slope_df = pd.DataFrame(slope_rows)

    customer_summary = total_by_customer.merge(slope_df, on="거래처", how="left")

    if not customer_summary.empty:
        customer_summary["최근추세"] = np.where(
            customer_summary["매출기울기"] > 0, "성장",
            np.where(customer_summary["매출기울기"] < 0, "감소", "안정")
        )

        customer_summary["분석요약"] = customer_summary.apply(
            lambda r: " / ".join([
                f"총매출 {int(r['총매출액']):,}원",
                f"품목수 {int(r['품목수'])}개",
                f"월평균매출 {int(r['월평균_매출']):,}원",
                f"최근추세 {r['최근추세']}",
                f"최근출고 {r['최근출고일']}",
            ]),
            axis=1
        )

    all_months = sorted(df["월"].dropna().unique().tolist())

    return {
        "df": df,
        "all_months": all_months,
        "customer_summary": customer_summary,
        "customer_monthly": customer_monthly,
        "customer_item_monthly": customer_item_monthly,
        "customer_item_summary": customer_item_summary,
    }


@st.cache_data(show_spinner=False)
def build_return_decline_item_analysis(q):
    if q is None or q.empty:
        return {}

    df = q.copy()
    if "날짜" not in df.columns:
        return {}

    df["월"] = pd.to_datetime(df["날짜"], errors="coerce").dt.strftime("%Y-%m")
    df = df[df["월"].notna() & (df["월"] != "")].copy()
    if df.empty:
        return {}

    df = safe_make_product_label(df)
    df = build_return_base(df)

    all_months = sorted(df["월"].unique().tolist())
    if len(all_months) < 2:
        return {
            "all_months": all_months,
            "item_rank": pd.DataFrame(),
            "item_monthly": pd.DataFrame(),
            "item_customer_monthly": pd.DataFrame(),
            "return_reason_df": pd.DataFrame(),
            "customer_return_reason_df": pd.DataFrame(),
        }

    mid_idx = len(all_months) // 2
    first_half = all_months[:mid_idx] if mid_idx > 0 else [all_months[0]]
    last_half = all_months[mid_idx:] if mid_idx < len(all_months) else [all_months[-1]]

    item_monthly = (
        df.groupby(["품목코드", "품목표시", "월"], dropna=False)
        .agg(
            매출액=("금액(원)", "sum"),
            출고량=("수량(M2)", "sum"),
            반품금액=("반품금액_표준", "sum"),
            반품수량=("반품수량_표준", "sum"),
        )
        .reset_index()
    )
    item_monthly["날짜축"] = pd.to_datetime(item_monthly["월"] + "-01", errors="coerce")

    item_customer_monthly = (
        df.groupby(["품목코드", "품목표시", "거래처", "월"], dropna=False)
        .agg(
            매출액=("금액(원)", "sum"),
            출고량=("수량(M2)", "sum"),
            반품금액=("반품금액_표준", "sum"),
            반품수량=("반품수량_표준", "sum"),
        )
        .reset_index()
    )
    item_customer_monthly["거래처"] = item_customer_monthly["거래처"].astype(str).str.strip()
    item_customer_monthly["날짜축"] = pd.to_datetime(item_customer_monthly["월"] + "-01", errors="coerce")

    return_reason_df = (
        df[df["반품여부_표준"]]
        .groupby(["품목코드", "품목표시", "거래처"], as_index=False)
        .agg(
            반품금액=("반품금액_표준", "sum"),
            반품수량=("반품수량_표준", "sum"),
            주요반품원인=("반품원인_표준", summarize_return_reason_text),
        )
    )
    if not return_reason_df.empty:
        return_reason_df["거래처"] = return_reason_df["거래처"].astype(str).str.strip()

    customer_return_reason_df = (
        df[df["반품여부_표준"]]
        .groupby(["품목표시", "거래처"], as_index=False)
        .agg(주요반품원인=("반품원인_표준", summarize_return_reason_text))
    )
    if not customer_return_reason_df.empty:
        customer_return_reason_df["거래처"] = customer_return_reason_df["거래처"].astype(str).str.strip()

    rows = []
    for (prod_code, prod_label), g in item_monthly.groupby(["품목코드", "품목표시"]):
        g = g.sort_values("월").copy()
        first_avg = g[g["월"].isin(first_half)]["매출액"].mean() if len(g[g["월"].isin(first_half)]) > 0 else 0
        last_avg = g[g["월"].isin(last_half)]["매출액"].mean() if len(g[g["월"].isin(last_half)]) > 0 else 0
        decline_amt = max(0.0, first_avg - last_avg)
        decline_rate = ((decline_amt / first_avg) * 100) if first_avg > 0 else 0
        slope = calc_slope(g["매출액"].tolist())

        cust_sub = item_customer_monthly[item_customer_monthly["품목코드"].astype(str) == str(prod_code)].copy()
        first_c = (
            cust_sub[cust_sub["월"].isin(first_half)]
            .groupby("거래처", as_index=False)["매출액"]
            .mean()
            .rename(columns={"매출액": "전반평균"})
        )
        last_c = (
            cust_sub[cust_sub["월"].isin(last_half)]
            .groupby("거래처", as_index=False)["매출액"]
            .mean()
            .rename(columns={"매출액": "후반평균"})
        )
        cc = first_c.merge(last_c, on="거래처", how="outer").fillna(0)
        cc["감소금액"] = cc["전반평균"] - cc["후반평균"]
        decline_customer_cnt = int((cc["감소금액"] > 0).sum()) if not cc.empty else 0

        total_sales = float(g["매출액"].sum())
        total_return_amt = float(g["반품금액"].sum())
        return_rate = (total_return_amt / total_sales * 100) if total_sales > 0 else 0

        rr = return_reason_df[return_reason_df["품목코드"].astype(str) == str(prod_code)].copy()
        top_reason = "반품 없음" if rr.empty else summarize_return_reason_text(rr["주요반품원인"])

        rows.append({
            "품목코드": str(prod_code),
            "품목표시": str(prod_label),
            "전체매출액": total_sales,
            "전반부_평균매출": int(round(first_avg, 0)),
            "후반부_평균매출": int(round(last_avg, 0)),
            "감소금액": int(round(decline_amt, 0)),
            "하락률(%)": round(decline_rate, 1),
            "감소고객사수": decline_customer_cnt,
            "반품금액": int(round(total_return_amt, 0)),
            "반품율(%)": round(return_rate, 1),
            "최근기울기": slope,
            "주요반품원인": top_reason,
        })

    item_rank = pd.DataFrame(rows)
    if not item_rank.empty:
        item_rank["품목하락점수"] = (
            scale_to_100(item_rank["감소금액"]) * 0.45 +
            scale_to_100(item_rank["하락률(%)"]) * 0.20 +
            scale_to_100(item_rank["감소고객사수"]) * 0.15 +
            scale_to_100(item_rank["반품금액"]) * 0.10 +
            scale_to_100(item_rank["최근기울기"], reverse=True) * 0.10
        ).round(1)
        item_rank["AI분석"] = item_rank.apply(
            lambda r: infer_ai_return_analysis(r) + " | " + infer_customer_item_decline_reason(r),
            axis=1
        )
        item_rank = item_rank.sort_values(
            ["품목하락점수", "감소금액", "하락률(%)"],
            ascending=[False, False, False]
        ).reset_index(drop=True)
        item_rank["순위"] = range(1, len(item_rank) + 1)

    return {
        "all_months": all_months,
        "first_half": first_half,
        "last_half": last_half,
        "item_rank": item_rank,
        "item_monthly": item_monthly,
        "item_customer_monthly": item_customer_monthly,
        "return_reason_df": return_reason_df,
        "customer_return_reason_df": customer_return_reason_df,
    }


DEFAULT_FILE = "data.xlsx"

st.markdown('<div class="app-main-title">출고 이력 검색(거래처/품목/가로폭/점착제)</div>', unsafe_allow_html=True)

uploaded = st.file_uploader("📂 다른 파일 업로드 (미업로드 시 기본 데이터 자동 로드)", type=["xlsx"])

if uploaded:
    file_bytes = uploaded.getvalue()
    st.success("✅ 업로드 파일 사용")
elif os.path.exists(DEFAULT_FILE):
    with open(DEFAULT_FILE, "rb") as f:
        file_bytes = f.read()
    st.info(f"📌 기본 데이터({DEFAULT_FILE}) 자동 로드")
else:
    st.info("GitHub 레포에 data.xlsx를 추가하거나 파일을 업로드하세요.")
    st.stop()

rec, alias, prod, adh, cust = load_excel(file_bytes)

st.sidebar.header("검색 필터")

dept_col = None
manager_col = None

if "담당부서" in rec.columns:
    dept_col = "담당부서"
elif "영업담당부서" in rec.columns:
    dept_col = "영업담당부서"

if "담당자" in rec.columns:
    manager_col = "담당자"

sel_dept = st.sidebar.multiselect(
    "담당부서",
    sorted_unique_safe(rec[dept_col]) if dept_col else [],
    placeholder="Choose options"
)

sel_manager = st.sidebar.multiselect(
    "담당자",
    sorted_unique_safe(rec[manager_col]) if manager_col else [],
    placeholder="Choose options"
)

sel_cust = st.sidebar.multiselect(
    "거래처",
    sorted_unique(rec["거래처"]) if "거래처" in rec.columns else [],
    placeholder="Choose options"
)
sel_prod = st.sidebar.multiselect(
    "품목코드",
    sorted_unique(rec["품목코드"]) if "품목코드" in rec.columns else [],
    placeholder="Choose options"
)
sel_adh = st.sidebar.multiselect(
    "점착제코드",
    sorted_unique(rec["점착제코드"]) if "점착제코드" in rec.columns else [],
    placeholder="Choose options"
)

date_min = pd.to_datetime(rec["날짜"], errors="coerce").min() if "날짜" in rec.columns else None
date_max = pd.to_datetime(rec["날짜"], errors="coerce").max() if "날짜" in rec.columns else None

sdate, edate = None, None
if date_min is not None and pd.notna(date_min) and date_max is not None and pd.notna(date_max):
    try:
        date_value = [date_min.date(), date_max.date()] if date_min <= date_max else [date_max.date(), date_max.date()]
        picked = st.sidebar.date_input("기간", value=date_value)
        if isinstance(picked, (list, tuple)) and len(picked) == 2:
            sdate, edate = picked[0], picked[1]
        elif picked:
            sdate = edate = picked
    except Exception:
        sdate, edate = None, None

st.sidebar.markdown("---")
st.sidebar.caption("💡 견적 레퍼런스: 품목코드·점착제코드·기간 필터 위주로 사용하세요.")

q = rec.copy()

if dept_col and sel_dept:
    q = q[q[dept_col].astype(str).str.strip().isin(sel_dept)]

if manager_col and sel_manager:
    q = q[q[manager_col].astype(str).str.strip().isin(sel_manager)]

if sel_cust and "거래처" in q.columns:
    q = q[q["거래처"].astype(str).isin(sel_cust)]
if sel_prod and "품목코드" in q.columns:
    q = q[q["품목코드"].astype(str).isin(sel_prod)]
if sel_adh and "점착제코드" in q.columns:
    q = q[q["점착제코드"].astype(str).isin(sel_adh)]
if sdate and edate and "날짜" in q.columns:
    q_date = pd.to_datetime(q["날짜"], errors="coerce")
    q = q[(q_date >= pd.to_datetime(sdate)) & (q_date <= pd.to_datetime(edate))]

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "거래처별 검색",
    "품목별 검색",
    "🏷️ 견적 레퍼런스",
    "📉 매출 하락 분석",
    "📊 거래처별 매출 분석",
    "매출 감소 품목 분석",
    "원자료",
])

with tab1:
    st.subheader("거래처별 → 품목별 출고 이력/최근 단가/가로폭")
    if q.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
    else:
        q1 = q.copy()
        q1["월"] = pd.to_datetime(q1["날짜"], errors="coerce").dt.strftime("%Y-%m") if "날짜" in q1.columns else ""

        for c in ["거래처", "품목코드", "점착제코드"]:
            if c in q1.columns:
                q1[c] = q1[c].astype(str)

        cols = ["거래처", "품목코드", "점착제코드", "점착제명", "가로폭이력", "최근날짜", "최근단가"]
        uc = [c for c in cols if c in q1.columns]

        if len(uc) == 0:
            st.info("표시할 컬럼이 없습니다.")
        else:
            g = (
                q1.groupby(uc, dropna=False)
                .agg(
                    출고횟수=("수량(M2)", "count"),
                    월평균_출고량=("수량(M2)", "mean"),
                    총량_M2=("수량(M2)", "sum"),
                    월평균_매출=("금액(원)", "mean"),
                    매출액=("금액(원)", "sum"),
                )
                .reset_index()
            )

            g["가중평균단가"] = np.where(g["총량_M2"] > 0, (g["매출액"] / g["총량_M2"]).round(0), 0)

            ordered_cols = [
                "거래처", "품목코드", "점착제코드", "점착제명", "가로폭이력",
                "최근날짜", "최근단가", "출고횟수", "월평균_출고량", "월평균_매출",
                "총량_M2", "매출액", "가중평균단가"
            ]
            ordered_cols = [c for c in ordered_cols if c in g.columns]
            sc = [c for c in ["거래처", "품목코드"] if c in g.columns]

            clean_and_safe_display(
                g[ordered_cols].sort_values(sc) if sc else g[ordered_cols],
                pinned_cols=["거래처", "품목코드"],
                text_cols=["거래처", "품목코드", "점착제코드", "점착제명", "가로폭이력", "최근날짜"],
                height=None,
            )

with tab2:
    st.subheader("품목별 → 거래처별 출고 이력/총량/단가/매출")
    if q.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
    else:
        q2 = q.copy()
        q2["월"] = pd.to_datetime(q2["날짜"], errors="coerce").dt.strftime("%Y-%m") if "날짜" in q2.columns else ""

        for c in ["거래처", "품목코드"]:
            if c in q2.columns:
                q2[c] = q2[c].astype(str)

        cols = ["품목코드", "거래처", "최근날짜", "최근단가"]
        uc = [c for c in cols if c in q2.columns]

        if len(uc) == 0:
            st.info("표시할 컬럼이 없습니다.")
        else:
            g2 = (
                q2.groupby(uc, dropna=False)
                .agg(
                    출고횟수=("수량(M2)", "count"),
                    월평균_출고량=("수량(M2)", "mean"),
                    총량_M2=("수량(M2)", "sum"),
                    월평균_매출=("금액(원)", "mean"),
                    매출액=("금액(원)", "sum"),
                )
                .reset_index()
            )

            g2["가중평균단가"] = np.where(g2["총량_M2"] > 0, (g2["매출액"] / g2["총량_M2"]).round(0), 0)

            ordered_cols = [
                "품목코드", "거래처", "최근날짜", "최근단가", "출고횟수",
                "월평균_출고량", "월평균_매출", "총량_M2", "매출액", "가중평균단가"
            ]
            ordered_cols = [c for c in ordered_cols if c in g2.columns]
            sc = [c for c in ["품목코드", "거래처"] if c in g2.columns]

            clean_and_safe_display(
                g2[ordered_cols].sort_values(sc) if sc else g2[ordered_cols],
                pinned_cols=["품목코드", "거래처"],
                text_cols=["품목코드", "거래처", "최근날짜"],
                height=None,
            )

with tab3:
    st.subheader("🏷️ 견적 레퍼런스 — 기준 견적가 & 판매 동향")
    q_ref = q.copy()

    if "단가(원/M2)" in q_ref.columns:
        q_ref = q_ref[q_ref["단가(원/M2)"].notna() & (q_ref["단가(원/M2)"] > 0)]

    for col in ["품목코드", "품목명(공식)"]:
        if col in q_ref.columns:
            q_ref = q_ref[~q_ref[col].astype(str).str.contains("샘플", case=False, na=False)]

    if q_ref.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
    else:
        overview, rep_ref, special_ref = build_quote_reference(q_ref)

        st.markdown("### 1) 품목 기준 견적 레퍼런스")
        overview_cols = [
            "품목코드", "점착제코드", "최저단가", "최고단가", "거래처수",
            "총출고횟수", "월평균_출고량", "월평균_매출", "총량_M2", "총매출액"
        ]
        overview_cols = [c for c in overview_cols if c in overview.columns]

        clean_and_safe_display(
            overview[overview_cols] if overview_cols else pd.DataFrame(),
            pinned_cols=["품목코드"],
            text_cols=["품목코드", "점착제코드"],
            height=None,
        )

        st.markdown("### 2) 업체 성향 AI 분석 기반 대표 레퍼런스")
        rep_cols = [
            "품목코드", "거래처", "업체성향", "AI분석",
            "최근단가", "최저단가", "최고단가", "최근날짜",
            "월평균_출고량", "월평균_매출", "최근추세", "총매출액"
        ]
        rep_cols = [c for c in rep_cols if c in rep_ref.columns]

        clean_and_safe_display(
            rep_ref[rep_cols] if rep_cols else pd.DataFrame(),
            pinned_cols=["품목코드", "거래처"],
            text_cols=["품목코드", "거래처", "업체성향", "AI분석", "최근날짜", "최근추세"],
            height=None,
        )

        st.markdown("### 3) 대표 업체 레퍼런스 확장")
        special_cols = [
            "대표구분", "품목코드", "거래처", "업체성향", "최근단가",
            "최저단가", "최고단가", "월평균_출고량", "월평균_매출",
            "최근추세", "총매출액", "AI분석"
        ]
        special_cols = [c for c in special_cols if c in special_ref.columns]

        render_banded_table(
            special_ref[special_cols] if (not special_ref.empty and special_cols) else pd.DataFrame(columns=special_cols),
            pinned_cols=["대표구분", "품목코드", "거래처"],
            text_cols=["대표구분", "품목코드", "거래처", "업체성향", "최근추세", "AI분석"],
            height=None,
        )

        st.markdown("### 4) 대표 업체 최근단가 비교")
        draw_quote_reference_chart(special_ref)

        export_df = special_ref[special_cols] if (not special_ref.empty and special_cols) else pd.DataFrame(columns=special_cols)
        rep_csv = export_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button(
            "📥 대표 업체 레퍼런스 CSV 다운로드",
            data=rep_csv,
            file_name="대표업체_레퍼런스.csv",
            mime="text/csv",
        )

with tab4:
    st.subheader("매출 하락 업체 분석")
    st.caption("감소금액 중심 75% + 감소추세 15% + 품목감소 10%")

    if q.empty or "날짜" not in q.columns or "금액(원)" not in q.columns or "거래처" not in q.columns:
        st.warning("분석에 필요한 데이터가 부족합니다.")
    else:
        detail_df, monthly_sales, customer_total_monthly_all, product_monthly_all, all_months = build_analysis_cache(q)
        priority_df, first_half, last_half = build_priority_results(monthly_sales, detail_df, all_months)

        if len(all_months) < 2:
            st.info("분석 기간이 너무 짧습니다. (최소 2개월 필요)")
        elif priority_df.empty:
            st.success("✅ 설정 기간 동안 매출이 하락한 업체가 없습니다.")
        else:
            st.info(
                "순위 산정 방식: "
                "① 감소금액 규모를 가장 크게 반영(75%), "
                "② 감소 추세/속도 반영(15%), "
                "③ 품목 감소 확산 반영(10%)"
            )

            top_count = max(1, int(np.ceil(len(priority_df) * 0.35)))
            top_priority = priority_df.head(top_count).copy()

            st.markdown("### 🎯 매출감소 추이 업체 LIST")
            display_cols = [
                "순위", "거래처", "AI_우선순위점수", "감소규모점수", "추세하락점수", "품목감소점수",
                "전체_매출액", "전반부_평균매출", "후반부_평균매출", "실제감소액", "하락률(%)", "분석_내역"
            ]

            edited_priority = clean_and_safe_display(
                top_priority[display_cols],
                key="priority_customers_editor",
                editable=True,
                pinned_cols=["순위", "거래처"],
                text_cols=["거래처", "분석_내역"],
                disabled_cols=display_cols,
            )

            if edited_priority is not None:
                csv_priority = edited_priority.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                st.download_button(
                    "📥 매출감소 추이 업체 LIST CSV 다운로드",
                    data=csv_priority,
                    file_name="매출감소추이_업체_LIST.csv",
                    mime="text/csv",
                )

            st.markdown("---")
            st.markdown("### 🔍 업체별 상세 분석")

            selected_customer = st.selectbox(
                "분석할 업체를 선택하세요",
                options=["선택하세요"] + [
                    f"{row['거래처']} (점수: {row['AI_우선순위점수']:.1f}점, 감소액: {row['실제감소액']:,}원)"
                    for _, row in top_priority.iterrows()
                ],
                key="customer_detail_select",
            )

            if selected_customer != "선택하세요":
                selected_cust_name = selected_customer.split(" (점수:")[0]

                customer_total_monthly = customer_total_monthly_all[
                    customer_total_monthly_all["거래처"] == selected_cust_name
                ].copy()
                product_monthly = product_monthly_all[
                    product_monthly_all["거래처"] == selected_cust_name
                ].copy()

                if customer_total_monthly.empty:
                    st.warning("해당 업체 데이터가 없습니다.")
                else:
                    st.markdown(f"## 📉 [{selected_cust_name}] 매출 추이 분석")
                    st.caption("전체 매출 감소 여부와 감소 기여 품목 분석")

                    pivot_prod = product_monthly.pivot_table(
                        index=["품목코드", "품목명(공식)", "품목표시"],
                        columns="월",
                        values="금액(원)",
                        aggfunc="sum",
                        fill_value=0
                    ) if not product_monthly.empty else pd.DataFrame()

                    contribution_rows = []
                    if not pivot_prod.empty:
                        for idx in pivot_prod.index:
                            prod_code, _, prod_label = idx
                            row = pivot_prod.loc[idx]

                            first_vals = [row[c] for c in first_half if c in row.index]
                            last_vals = [row[c] for c in last_half if c in row.index]

                            first_avg = float(np.mean(first_vals)) if len(first_vals) > 0 else 0.0
                            last_avg = float(np.mean(last_vals)) if len(last_vals) > 0 else 0.0
                            decline_amt = first_avg - last_avg

                            monthly_vals = product_monthly[
                                product_monthly["품목표시"] == prod_label
                            ].sort_values("월")["금액(원)"]

                            total_sales = float(row.sum())
                            growth_amt = max(0.0, last_avg - first_avg)

                            contribution_rows.append({
                                "품목코드": str(prod_code),
                                "품목표시": str(prod_label),
                                "전반부_평균": int(round(first_avg, 0)),
                                "후반부_평균": int(round(last_avg, 0)),
                                "감소액": int(round(decline_amt, 0)),
                                "변화액": int(round(last_avg - first_avg, 0)),
                                "변화율(%)": round(((last_avg - first_avg) / first_avg) * 100, 1) if first_avg > 0 else 0.0,
                                "총매출": int(round(total_sales, 0)),
                                "성장매출": int(round(growth_amt, 0)),
                                "CV": round(calc_cv(monthly_vals), 3),
                            })

                    contribution_df = pd.DataFrame(contribution_rows)
                    if not contribution_df.empty:
                        contribution_df = contribution_df.sort_values("감소액", ascending=False).reset_index(drop=True)

                    st.markdown("### 📋 품목별 변화율 상세")
                    detail_cols = [
                        "품목코드", "전반부_평균", "후반부_평균",
                        "감소액", "변화액", "변화율(%)", "총매출", "성장매출", "CV"
                    ]
                    clean_and_safe_display(
                        contribution_df[detail_cols] if not contribution_df.empty else pd.DataFrame(columns=detail_cols),
                        pinned_cols=["품목코드"],
                        text_cols=["품목코드"]
                    )

                    st.markdown("---")
                    st.markdown("### 📈 매출 및 품목별 변동 추이")

                    fig_total = go.Figure()
                    fig_total.add_trace(go.Scatter(
                        x=customer_total_monthly["날짜축"],
                        y=customer_total_monthly["금액(원)"],
                        mode="lines+markers+text",
                        name="월별 총매출",
                        line=dict(color="#1f77b4", width=3),
                        marker=dict(size=8),
                        text=[sales_to_manwon_label(v) for v in customer_total_monthly["금액(원)"]],
                        textposition="top center",
                        textfont=dict(size=10, color="#1f77b4"),
                        hovertemplate="월: %{x|%Y-%m}<br>매출: %{y:,.0f}원<br>만원단위: %{text}<extra></extra>",
                    ))

                    if len(customer_total_monthly) >= 2:
                        try:
                            x_num = np.arange(len(customer_total_monthly))
                            y_num = customer_total_monthly["금액(원)"].values.astype(float)
                            coef = np.polyfit(x_num, y_num, 1)
                            trend = coef[0] * x_num + coef[1]
                            fig_total.add_trace(go.Scatter(
                                x=customer_total_monthly["날짜축"],
                                y=trend,
                                mode="lines",
                                name="추세선",
                                line=dict(color="red", dash="dash", width=2),
                                hoverinfo="skip",
                            ))
                        except Exception:
                            pass

                    fig_total = apply_mobile_friendly_line_layout(
                        fig_total,
                        customer_total_monthly["날짜축"],
                        y_title="매출액(원)",
                        height=430
                    )
                    fig_total.update_layout(title="1️⃣ 업체 전체 월별 매출 추이")
                    st.plotly_chart(fig_total, use_container_width=True, key=f"tab4_total_{selected_cust_name}")

with tab5:
    st.subheader("📊 거래처별 매출 분석")
    st.caption("거래처의 전체 매출 흐름, 품목별 변화율, 최근 단가/가로폭 이력, 선택 품목의 원자료를 확인합니다.")

    if q.empty or "거래처" not in q.columns or "날짜" not in q.columns or "금액(원)" not in q.columns:
        st.warning("거래처별 매출 분석에 필요한 데이터가 부족합니다.")
    else:
        pack = build_customer_sales_analysis(q)

        customer_summary = pack.get("customer_summary", pd.DataFrame())
        customer_monthly = pack.get("customer_monthly", pd.DataFrame())
        customer_item_monthly = pack.get("customer_item_monthly", pd.DataFrame())
        customer_item_summary = pack.get("customer_item_summary", pd.DataFrame())
        all_months = pack.get("all_months", [])
        base_df = pack.get("df", pd.DataFrame())

        if customer_summary.empty:
            st.info("분석 가능한 거래처 데이터가 없습니다.")
        else:
            st.markdown("### 1) 매출분석 자료")
            st.caption("거래처별 총매출, 월평균 매출, 품목 수, 최근 추세를 한 번에 확인하는 요약 자료입니다.")

            summary_cols = [
                "거래처", "총매출액", "월평균_매출", "총판매량", "월평균_판매량",
                "품목수", "거래개월수", "출고건수", "최근출고일", "매출기울기", "매출CV", "최근추세", "분석요약"
            ]
            summary_cols = [c for c in summary_cols if c in customer_summary.columns]

            clean_and_safe_display(
                customer_summary[summary_cols].sort_values(["총매출액", "거래처"], ascending=[False, True]),
                pinned_cols=["거래처"],
                text_cols=["거래처", "최근출고일", "최근추세", "분석요약"],
                height=None,
            )

            st.markdown("---")
            st.markdown("### 2) 업체별 상세 분석")
            st.caption("선택한 업체의 월별 매출 흐름과 품목별 변화율 상세를 확인합니다.")

            customer_options = customer_summary.sort_values(["총매출액", "거래처"], ascending=[False, True])["거래처"].dropna().astype(str).tolist()

            if len(customer_options) == 0:
                st.info("선택 가능한 업체가 없습니다.")
            else:
                selected_customer_analysis = st.selectbox(
                    "분석할 업체를 선택하세요",
                    options=customer_options,
                    key="customer_sales_analysis_select"
                )

                selected_customer_analysis = str(selected_customer_analysis)

                cust_month = customer_monthly[customer_monthly["거래처"].astype(str) == selected_customer_analysis].copy()
                cust_item_month = customer_item_monthly[customer_item_monthly["거래처"].astype(str) == selected_customer_analysis].copy()
                cust_item_sum = customer_item_summary[customer_item_summary["거래처"].astype(str) == selected_customer_analysis].copy()
                cust_raw = base_df[base_df["거래처"].astype(str) == selected_customer_analysis].copy()

                if cust_month.empty:
                    st.info("선택한 업체의 월별 매출 데이터가 없습니다.")
                else:
                    top_info = customer_summary[customer_summary["거래처"].astype(str) == selected_customer_analysis].copy()
                    if not top_info.empty:
                        r = top_info.iloc[0]
                        m1, m2, m3, m4 = st.columns(4)
                        with m1:
                            st.metric("총매출액", f"{int(r['총매출액']):,} 원")
                        with m2:
                            st.metric("월평균 매출", f"{int(r['월평균_매출']):,} 원")
                        with m3:
                            st.metric("품목 수", f"{int(r['품목수'])} 개")
                        with m4:
                            st.metric("최근 추세", str(r["최근추세"]))

                    st.markdown("#### 업체 전체 월별 매출 추이")
                    st.caption("업체의 전체 월 매출 흐름과 추세선을 확인합니다.")

                    cust_month_axis = build_month_axis_frame(cust_month["월"].unique().tolist())
                    cust_month_plot = align_monthly_series(
                        cust_month_axis,
                        cust_month[["월", "매출액"]].copy(),
                        "매출액"
                    )

                    fig_cust = go.Figure()
                    fig_cust.add_trace(go.Scatter(
                        x=cust_month_plot["날짜축"],
                        y=cust_month_plot["매출액"],
                        mode="lines+markers+text",
                        name="월매출",
                        text=[sales_to_manwon_label(v) for v in cust_month_plot["매출액"]],
                        textposition="top center",
                        line=dict(width=3, color="#1f77b4"),
                        marker=dict(size=8),
                        cliponaxis=False,
                        hovertemplate="월: %{x|%Y-%m}<br>매출액: %{y:,.0f}원<br>만원: %{text}<extra></extra>",
                    ))

                    if len(cust_month_plot) >= 2:
                        try:
                            x_num = np.arange(len(cust_month_plot))
                            y_num = cust_month_plot["매출액"].astype(float).values
                            coef = np.polyfit(x_num, y_num, 1)
                            trend = coef[0] * x_num + coef[1]
                            fig_cust.add_trace(go.Scatter(
                                x=cust_month_plot["날짜축"],
                                y=trend,
                                mode="lines",
                                name="추세선",
                                line=dict(color="red", dash="dash", width=2),
                                hoverinfo="skip"
                            ))
                        except Exception:
                            pass

                    fig_cust = apply_mobile_friendly_line_layout(
                        fig_cust,
                        cust_month_plot["날짜축"],
                        y_title="매출액(원)",
                        height=430
                    )
                    st.plotly_chart(fig_cust, use_container_width=True, key=f"customer_sales_monthly_{selected_customer_analysis}")

                    if not cust_item_sum.empty:
                        total_sales_customer = float(cust_item_sum["총매출액"].sum()) if "총매출액" in cust_item_sum.columns else 0.0
                        total_qty_customer = float(cust_item_sum["총판매량"].sum()) if "총판매량" in cust_item_sum.columns else 0.0

                        cust_item_sum["매출비중(%)"] = np.where(
                            total_sales_customer > 0,
                            cust_item_sum["총매출액"] / total_sales_customer * 100,
                            0
                        )
                        cust_item_sum["판매량비중(%)"] = np.where(
                            total_qty_customer > 0,
                            cust_item_sum["총판매량"] / total_qty_customer * 100,
                            0
                        )

                        item_change_rows = []
                        half_idx = len(all_months) // 2
                        first_half = all_months[:half_idx] if half_idx > 0 else all_months[:1]
                        last_half = all_months[half_idx:] if half_idx < len(all_months) else all_months[-1:]

                        for _, row in cust_item_sum.iterrows():
                            prod_label = str(row["품목표시"])
                            sub = cust_item_month[cust_item_month["품목표시"].astype(str) == prod_label].copy().sort_values("월")

                            first_avg_sales = sub[sub["월"].isin(first_half)]["매출액"].mean() if len(sub[sub["월"].isin(first_half)]) > 0 else 0
                            last_avg_sales = sub[sub["월"].isin(last_half)]["매출액"].mean() if len(sub[sub["월"].isin(last_half)]) > 0 else 0

                            first_avg_qty = sub[sub["월"].isin(first_half)]["판매량"].mean() if len(sub[sub["월"].isin(first_half)]) > 0 else 0
                            last_avg_qty = sub[sub["월"].isin(last_half)]["판매량"].mean() if len(sub[sub["월"].isin(last_half)]) > 0 else 0

                            item_change_rows.append({
                                "품목코드": row["품목코드"],
                                "품목표시": row["품목표시"],
                                "총매출액": row["총매출액"],
                                "총판매량": row["총판매량"],
                                "매출비중(%)": row["매출비중(%)"],
                                "판매량비중(%)": row["판매량비중(%)"],
                                "전반부_평균매출": round(first_avg_sales, 0),
                                "후반부_평균매출": round(last_avg_sales, 0),
                                "매출변화액": round(last_avg_sales - first_avg_sales, 0),
                                "매출변화율(%)": round(((last_avg_sales - first_avg_sales) / first_avg_sales) * 100, 1) if first_avg_sales > 0 else 0,
                                "전반부_평균판매량": round(first_avg_qty, 1),
                                "후반부_평균판매량": round(last_avg_qty, 1),
                                "판매량변화율(%)": round(((last_avg_qty - first_avg_qty) / first_avg_qty) * 100, 1) if first_avg_qty > 0 else 0,
                                "최근단가": row.get("최근단가", np.nan),
                                "최근날짜": row.get("최근날짜", ""),
                                "가로폭이력": row.get("가로폭이력", ""),
                            })

                        item_change_df = pd.DataFrame(item_change_rows)
                        item_change_df = item_change_df.sort_values(["총매출액", "품목코드"], ascending=[False, True]).reset_index(drop=True)

                        st.markdown("#### 품목별 변화율 상세 (총매출순)")
                        st.caption("총매출이 큰 품목 순으로, 전반부 대비 후반부의 매출/판매량 변화율을 확인합니다.")

                        change_cols = [
                            "품목코드", "총매출액", "총판매량",
                            "매출비중(%)", "판매량비중(%)",
                            "전반부_평균매출", "후반부_평균매출", "매출변화액", "매출변화율(%)",
                            "전반부_평균판매량", "후반부_평균판매량", "판매량변화율(%)",
                            "최근단가", "최근날짜", "가로폭이력"
                        ]
                        change_cols = [c for c in change_cols if c in item_change_df.columns]

                        clean_and_safe_display(
                            item_change_df[change_cols],
                            pinned_cols=["품목코드"],
                            text_cols=["품목코드", "최근날짜", "가로폭이력"],
                            height=None,
                        )

                        st.markdown("#### 선택 품목 원자료")
                        st.caption("품목별 변화율 상세에서 선택한 품목의 해당 거래처 원자료를 표시합니다.")

                        item_options = item_change_df["품목코드"].dropna().astype(str).tolist()
                        if len(item_options) == 0:
                            st.info("선택 가능한 품목이 없습니다.")
                        else:
                            selected_item_code_detail = st.selectbox(
                                "원자료를 볼 품목 선택",
                                options=item_options,
                                key="customer_analysis_item_raw_select"
                            )

                            raw_item_df = cust_raw[cust_raw["품목코드"].astype(str) == str(selected_item_code_detail)].copy()
                            raw_item_df = raw_item_df.sort_values("날짜", ascending=False) if "날짜" in raw_item_df.columns else raw_item_df

                            raw_cols = [
                                c for c in [
                                    "날짜", "거래처", "품목코드", "품목명(공식)", "점착제코드", "점착제명",
                                    "가로폭(mm)", "수량(M2)", "단가(원/M2)", "금액(원)", "비고",
                                    "담당부서", "영업담당부서", "담당자"
                                ] if c in raw_item_df.columns
                            ]

                            clean_and_safe_display(
                                raw_item_df[raw_cols] if raw_cols else pd.DataFrame(),
                                pinned_cols=["날짜", "거래처", "품목코드"],
                                text_cols=["날짜", "거래처", "품목코드", "품목명(공식)", "점착제코드", "점착제명", "비고", "담당부서", "영업담당부서", "담당자"],
                                height=None,
                            )
                    else:
                        st.info("선택한 업체의 품목별 자료가 없습니다.")

with tab6:
    st.subheader("품목별 하락 원인 분석")
    st.caption("반품은 금액(원) 음수 기준, 반품 원인은 음수 행의 비고 내용을 사용합니다.")

    if q.empty or "날짜" not in q.columns or "금액(원)" not in q.columns or "품목코드" not in q.columns:
        st.warning("분석에 필요한 데이터가 부족합니다.")
    else:
        pack = build_return_decline_item_analysis(q)
        item_rank = pack.get("item_rank", pd.DataFrame())
        item_monthly = pack.get("item_monthly", pd.DataFrame())
        item_customer_monthly = pack.get("item_customer_monthly", pd.DataFrame())
        return_reason_df = pack.get("return_reason_df", pd.DataFrame())
        customer_return_reason_df = pack.get("customer_return_reason_df", pd.DataFrame())
        first_half = pack.get("first_half", [])
        last_half = pack.get("last_half", [])

        if item_rank.empty:
            st.info("분석할 품목 데이터가 없습니다.")
        else:
            top_count = max(1, int(np.ceil(len(item_rank) * 0.35)))
            top_items = item_rank.head(top_count).copy()

            st.info(
                "품목 우선순위: 감소금액, 하락률, 감소고객사수, 반품금액, 최근 하락추세를 반영하여 상위 35% 품목을 우선 분석합니다."
            )

            st.markdown("### 1) 감소 품목 우선순위 LIST (상위 35%)")
            rank_cols = [
                "순위", "품목코드", "품목하락점수",
                "전체매출액", "전반부_평균매출", "후반부_평균매출",
                "감소금액", "하락률(%)", "감소고객사수",
                "반품금액", "반품율(%)", "주요반품원인", "AI분석"
            ]
            rank_cols = [c for c in rank_cols if c in top_items.columns]

            clean_and_safe_display(
                top_items[rank_cols],
                pinned_cols=["순위", "품목코드"],
                text_cols=["품목코드", "주요반품원인", "AI분석"],
                height=None,
            )

            st.markdown("---")
            st.markdown("### 2) 선택 품목별 업체 감소현황")

            item_options = top_items["품목표시"].astype(str).tolist()
            selected_item = item_options[0] if len(item_options) > 0 else None

            if selected_item:
                select_col, c1, c2, c3, c4 = st.columns([2.6, 1, 1, 1, 1])

                with select_col:
                    selected_item = st.selectbox(
                        "품목 선택",
                        options=item_options,
                        index=0,
                        key="decline_item_select_top_aligned_v5"
                    )

                item_row = top_items[top_items["품목표시"].astype(str) == str(selected_item)].copy()
                item_month = item_monthly[item_monthly["품목표시"].astype(str) == str(selected_item)].copy()
                item_cust = item_customer_monthly[item_customer_monthly["품목표시"].astype(str) == str(selected_item)].copy()
                item_cust["거래처"] = item_cust["거래처"].astype(str).str.strip()
                item_return = return_reason_df[return_reason_df["품목표시"].astype(str) == str(selected_item)].copy()
                reason_by_customer = customer_return_reason_df[
                    customer_return_reason_df["품목표시"].astype(str) == str(selected_item)
                ].copy()

                if not item_row.empty:
                    rr = item_row.iloc[0]
                    with c1:
                        st.metric("품목하락점수", f"{rr['품목하락점수']:.1f}")
                    with c2:
                        st.metric("감소금액", f"{int(rr['감소금액']):,} 원")
                    with c3:
                        st.metric("하락률", f"{rr['하락률(%)']:.1f}%")
                    with c4:
                        st.metric("반품금액", f"{int(rr['반품금액']):,} 원")

                common_month_axis = build_month_axis_frame(
                    item_month["월"].unique().tolist() if not item_month.empty else []
                )

                st.markdown("#### 선택 품목 월별 추이")
                if not item_month.empty and not common_month_axis.empty:
                    item_month_plot = align_monthly_series(
                        common_month_axis,
                        item_month[["월", "매출액"]].copy(),
                        "매출액"
                    )

                    fig_item = go.Figure()
                    fig_item.add_trace(go.Scatter(
                        x=item_month_plot["날짜축"],
                        y=item_month_plot["매출액"],
                        mode="lines+markers+text",
                        text=[sales_to_manwon_label(v) for v in item_month_plot["매출액"]],
                        textposition="top center",
                        name="월매출",
                        line=dict(width=3, color="#1f77b4"),
                        marker=dict(size=8),
                        cliponaxis=False,
                        hovertemplate="월: %{x|%Y-%m}<br>판매금액: %{y:,.0f}원<br>만원: %{text}<extra></extra>",
                    ))
                    fig_item = apply_mobile_friendly_line_layout(
                        fig_item,
                        common_month_axis["날짜축"],
                        y_title="판매금액(원)",
                        height=380
                    )
                    st.plotly_chart(
                        fig_item,
                        use_container_width=True,
                        key=f"item_monthly_chart_{selected_item}_v5"
                    )
                else:
                    st.info("선택 품목의 월별 추이 데이터가 없습니다.")

                if not item_cust.empty:
                    first_c = (
                        item_cust[item_cust["월"].isin(first_half)]
                        .groupby("거래처", as_index=False)
                        .agg(
                            전반부_평균매출=("매출액", "mean"),
                            전반부_반품금액=("반품금액", "sum"),
                        )
                    )
                    last_c = (
                        item_cust[item_cust["월"].isin(last_half)]
                        .groupby("거래처", as_index=False)
                        .agg(
                            후반부_평균매출=("매출액", "mean"),
                            후반부_반품금액=("반품금액", "sum"),
                        )
                    )

                    cust_summary = first_c.merge(last_c, on="거래처", how="outer").fillna(0)
                    cust_summary["거래처"] = cust_summary["거래처"].astype(str).str.strip()
                    cust_summary["감소금액"] = cust_summary["전반부_평균매출"] - cust_summary["후반부_평균매출"]
                    cust_summary["하락률(%)"] = np.where(
                        cust_summary["전반부_평균매출"] > 0,
                        ((cust_summary["전반부_평균매출"] - cust_summary["후반부_평균매출"]) / cust_summary["전반부_평균매출"]) * 100,
                        0
                    )
                    cust_summary["반품금액"] = cust_summary["전반부_반품금액"] + cust_summary["후반부_반품금액"]
                    cust_summary["반품율(%)"] = np.where(
                        cust_summary["전반부_평균매출"].abs() + cust_summary["후반부_평균매출"].abs() > 0,
                        cust_summary["반품금액"] / (cust_summary["전반부_평균매출"].abs() + cust_summary["후반부_평균매출"].abs()) * 100,
                        0
                    )

                    if not reason_by_customer.empty:
                        reason_by_customer["거래처"] = reason_by_customer["거래처"].astype(str).str.strip()
                        cust_summary = cust_summary.merge(
                            reason_by_customer[["거래처", "주요반품원인"]].drop_duplicates(),
                            on="거래처",
                            how="left"
                        )
                    else:
                        cust_summary["주요반품원인"] = np.nan

                    cust_summary["주요반품원인"] = cust_summary["주요반품원인"].fillna("반품 없음")

                    slope_map = (
                        item_cust.sort_values(["거래처", "월"])
                        .groupby("거래처")["매출액"]
                        .apply(lambda s: calc_slope(s.tolist()))
                        .reset_index(name="최근기울기")
                    )
                    cust_summary = cust_summary.merge(slope_map, on="거래처", how="left")

                    cust_summary["감소원인"] = cust_summary.apply(infer_customer_item_decline_reason, axis=1)
                    cust_summary["AI반품분석"] = cust_summary.apply(infer_ai_return_analysis, axis=1)

                    cust_summary = cust_summary.sort_values(
                        ["감소금액", "하락률(%)", "반품금액"],
                        ascending=[False, False, False]
                    ).reset_index(drop=True)
                    cust_summary["순위"] = range(1, len(cust_summary) + 1)

                    st.markdown("#### 업체별 월판매 현황")
                    customer_options = cust_summary.sort_values("순위")["거래처"].dropna().astype(str).str.strip().tolist()

                    if len(customer_options) > 0:
                        left_col, m1, m2, m3 = st.columns([2.9, 1.2, 1.2, 1.2])

                        with left_col:
                            selected_customer = st.selectbox(
                                "업체 선택",
                                options=customer_options,
                                index=0,
                                key="decline_item_customer_select_v5"
                            )

                        selected_customer = str(selected_customer).strip()
                        selected_row = cust_summary[
                            cust_summary["거래처"].astype(str).str.strip() == selected_customer
                        ].copy()

                        if not selected_row.empty:
                            sr = selected_row.iloc[0]
                            with m1:
                                st.metric("순위", f"{int(sr['순위'])}")
                            with m2:
                                st.metric("감소금액", f"{int(sr['감소금액']):,} 원")
                            with m3:
                                st.metric("반품금액", f"{int(sr['반품금액']):,} 원")
                        else:
                            with m1:
                                st.metric("순위", "-")
                            with m2:
                                st.metric("감소금액", "-")
                            with m3:
                                st.metric("반품금액", "-")

                        st.markdown(f"#### {selected_item} 월별 판매 추이")

                        selected_customer_month = item_cust.copy()
                        selected_customer_month["거래처"] = selected_customer_month["거래처"].astype(str).str.strip()
                        selected_customer_month = selected_customer_month[
                            selected_customer_month["거래처"] == selected_customer
                        ].copy()

                        if not common_month_axis.empty:
                            selected_customer_month_plot = align_monthly_series(
                                common_month_axis,
                                selected_customer_month[["월", "매출액"]].copy() if not selected_customer_month.empty else pd.DataFrame(columns=["월", "매출액"]),
                                "매출액"
                            )

                            fig_item_customer = go.Figure()
                            fig_item_customer.add_trace(go.Scatter(
                                x=selected_customer_month_plot["날짜축"],
                                y=selected_customer_month_plot["매출액"],
                                mode="lines+markers+text",
                                text=[sales_to_manwon_label(v) for v in selected_customer_month_plot["매출액"]],
                                textposition="top center",
                                name=selected_customer,
                                line=dict(width=3, color="#1f77b4"),
                                marker=dict(size=8),
                                cliponaxis=False,
                                hovertemplate="월: %{x|%Y-%m}<br>판매금액: %{y:,.0f}원<br>만원: %{text}<extra></extra>",
                            ))
                            fig_item_customer = apply_mobile_friendly_line_layout(
                                fig_item_customer,
                                common_month_axis["날짜축"],
                                y_title="판매금액(원)",
                                height=380
                            )
                            st.plotly_chart(
                                fig_item_customer,
                                use_container_width=True,
                                key=f"item_customer_monthly_chart_{selected_item}_{selected_customer}_v5"
                            )
                        else:
                            st.info("선택한 업체의 해당 품목 판매금액 월별 데이터가 없습니다.")
                    else:
                        st.info("선택 가능한 업체가 없습니다.")

                    st.markdown("#### 업체별 감소현황 Data")
                    customer_display_cols = [
                        "순위", "거래처", "전반부_평균매출", "후반부_평균매출",
                        "감소금액", "하락률(%)", "반품금액", "반품율(%)",
                        "주요반품원인", "AI반품분석", "감소원인"
                    ]
                    customer_display_cols = [c for c in customer_display_cols if c in cust_summary.columns]

                    clean_and_safe_display(
                        cust_summary[customer_display_cols],
                        pinned_cols=["순위", "거래처"],
                        text_cols=["거래처", "주요반품원인", "AI반품분석", "감소원인"],
                        height=None,
                    )

                    st.markdown("#### 품목 반품 원인 요약")
                    if item_return.empty:
                        st.info("해당 품목의 반품 데이터가 없습니다.")
                    else:
                        item_return_sorted = item_return.copy()
                        item_return_sorted["거래처"] = item_return_sorted["거래처"].astype(str).str.strip()
                        if "반품금액" in item_return_sorted.columns:
                            item_return_sorted["반품금액"] = pd.to_numeric(item_return_sorted["반품금액"], errors="coerce").fillna(0)
                            item_return_sorted = item_return_sorted.sort_values(
                                ["반품금액", "반품수량", "거래처"],
                                ascending=[False, False, True]
                            ).reset_index(drop=True)

                        clean_and_safe_display(
                            item_return_sorted[["거래처", "반품금액", "반품수량", "주요반품원인"]],
                            text_cols=["거래처", "주요반품원인"],
                            height=None,
                        )
                else:
                    st.info("업체별 감소현황 데이터가 없습니다.")

with tab7:
    st.subheader("원자료(필터 적용됨)")
    raw_cols = [c for c in q.columns if c != "품목명(공식)"]
    clean_and_safe_display(
        q[raw_cols],
        pinned_cols=["거래처", "품목코드"],
        text_cols=[
            "거래처", "품목코드", "점착제코드", "점착제명", "가로폭이력",
            "최근날짜", "월", "비고", "담당부서", "영업담당부서", "담당자"
        ],
    )
