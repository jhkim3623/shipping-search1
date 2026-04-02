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

/* 가로 row 정렬 */
div[data-testid="stHorizontalBlock"] {
    gap: 0.6rem;
    align-items: flex-end !important;
}

/* column 내부 wrapper 높이 안정화 */
div[data-testid="column"] > div {
    width: 100% !important;
}

/* 라벨 공통 */
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

/* metric box */
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

/* metric 내부 */
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

/* select 공통 wrapper */
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

/* select 내부 값 세로 중앙 */
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

/* choose options 복구 */
div[data-baseweb="select"] div[class*="placeholder"] {
    color: #6b7280 !important;
    opacity: 1 !important;
    font-size: 0.95rem !important;
    line-height: 1.25 !important;
    margin: 0 !important;
    display: flex !important;
    align-items: center !important;
}

/* multiselect tag 정렬 */
div[data-baseweb="tag"] {
    display: flex !important;
    align-items: center !important;
    margin-top: 0 !important;
    margin-bottom: 0 !important;
}

/* 사이드바 필터 글씨 선명하게 */
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

/* 멀티셀렉트 내부 텍스트 잘림 방지 */
section[data-testid="stSidebar"] div[data-baseweb="select"] * {
    overflow: visible !important;
}

/* dataframe */
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

    temp["품목코드"] = temp["품목코드"].fillna("").astype(str)
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
        "원인추정", "감소원인", "비고", "AI반품분석", "담당부서", "영업담당부서", "담당자"
    }

    for col in display_df.columns:
        pinned = col in pinned_cols

        if col in text_cols or col in fixed_text_like_cols:
            column_config[col] = st.column_config.TextColumn(
                col,
                width="large" if col in [
                    "품목명(공식)", "품목표시", "가로폭이력", "분석_내역",
                    "AI분석", "원인추정", "감소원인", "비고", "주요반품원인", "AI반품분석"
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
    try:
        s = pd.Series(series).dropna().astype(str).str.strip()
        s = s[~s.isin(["", "0", "0.0", "nan", "NaN", "None", "<NA>", "NaT"])]
        vals = list(dict.fromkeys(s.tolist()))
        return sorted(vals, key=lambda x: str(x))
    except Exception:
        return []


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
    month_list = sorted([str(m) for m in months if pd.notna(m) and str(m).strip() != ""])
    axis_df = pd.DataFrame({"월": month_list})
    if axis_df.empty:
        axis_df["날짜축"] = pd.NaT
        return axis_df
    axis_df["날짜축"] = pd.to_datetime(axis_df["월"] + "-01", errors="coerce")
    axis_df = axis_df.dropna(subset=["날짜축"]).sort_values("날짜축").reset_index(drop=True)
    return axis_df


def align_monthly_series(base_month_df, data_df, value_col):
    if base_month_df is None or base_month_df.empty:
        return pd.DataFrame(columns=["월", "날짜축", value_col])

    out = base_month_df.copy()

    if data_df is None or data_df.empty or value_col not in data_df.columns or "월" not in data_df.columns:
        out[value_col] = 0
        return out

    temp = data_df.copy()
    temp["월"] = temp["월"].astype(str)
    temp = temp.groupby("월", as_index=False)[value_col].sum()
    out = out.merge(temp, on="월", how="left")
    out[value_col] = pd.to_numeric(out[value_col], errors="coerce").fillna(0)
    return out


def sales_to_manwon_label(value):
    if pd.isna(value):
        return ""
    return f"{int(round(float(value) / 10000.0, 0)):,}"


def make_text_position_map(items):
    positions = ["top center", "bottom center", "middle right", "middle left"]
    return {str(item): positions[i % len(positions)] for i, item in enumerate(items)}


def make_indexed_series(df, group_col, value_col, time_col):
    if df is None or df.empty:
        return pd.DataFrame(columns=[group_col, time_col, value_col, "지수값"])

    required = {group_col, value_col, time_col}
    if not required.issubset(set(df.columns)):
        return pd.DataFrame(columns=[group_col, time_col, value_col, "지수값"])

    temp = df.copy()
    temp[group_col] = temp[group_col].astype(str)
    temp[value_col] = pd.to_numeric(temp[value_col], errors="coerce").fillna(0)
    temp[time_col] = pd.to_datetime(temp[time_col], errors="coerce")
    temp = temp.dropna(subset=[time_col]).sort_values([group_col, time_col]).reset_index(drop=True)

    out_list = []
    for gname, g in temp.groupby(group_col, dropna=False):
        sub = g.copy()
        base_candidates = sub[value_col].replace(0, np.nan).dropna()
        if len(base_candidates) == 0:
            sub["지수값"] = np.nan
        else:
            base = base_candidates.iloc[0]
            if pd.isna(base) or base == 0:
                sub["지수값"] = np.nan
            else:
                sub["지수값"] = (sub[value_col] / base) * 100.0
        out_list.append(sub)

    if len(out_list) == 0:
        return pd.DataFrame(columns=[group_col, time_col, value_col, "지수값"])

    return pd.concat(out_list, ignore_index=True)


def should_show_helper_chart(top_product_monthly):
    if top_product_monthly is None or top_product_monthly.empty:
        return False, 0.0

    if "품목표시" not in top_product_monthly.columns:
        return False, 0.0

    value_col = "금액(원)" if "금액(원)" in top_product_monthly.columns else ("매출액" if "매출액" in top_product_monthly.columns else None)
    if value_col is None:
        return False, 0.0

    temp = (
        top_product_monthly.groupby("품목표시", as_index=False)[value_col]
        .sum()
        .rename(columns={value_col: "총매출"})
    )
    temp = temp[temp["총매출"] > 0].copy()

    if len(temp) < 2:
        return False, 0.0

    max_v = temp["총매출"].max()
    min_v = temp["총매출"].min()

    if min_v <= 0:
        return True, np.inf

    ratio = max_v / min_v
    return bool(ratio >= 5), ratio


def calc_cv(series):
    s = pd.to_numeric(pd.Series(series), errors="coerce").dropna()
    if len(s) == 0:
        return 0.0
    m = s.mean()
    if m == 0 or pd.isna(m):
        return 0.0
    return float(s.std(ddof=0) / m)


def calc_slope(values):
    y = pd.to_numeric(pd.Series(values), errors="coerce").fillna(0).values.astype(float)
    if len(y) < 2:
        return 0.0
    if np.all(np.isclose(y, y[0])):
        return 0.0
    x = np.arange(len(y))
    return float(np.polyfit(x, y, 1)[0])


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

    temp["금액(원)"] = pd.to_numeric(temp["금액(원)"], errors="coerce").fillna(0)
    if "수량(M2)" in temp.columns:
        temp["수량(M2)"] = pd.to_numeric(temp["수량(M2)"], errors="coerce").fillna(0)
    else:
        temp["수량(M2)"] = 0

    if "거래처" in temp.columns:
        temp["거래처"] = temp["거래처"].astype(str).str.strip()

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

    if "날짜" in rec.columns:
        rec["날짜"] = pd.to_datetime(rec["날짜"], errors="coerce")

    for c in ["가로폭(mm)", "수량(M2)", "단가(원/M2)", "금액(원)"]:
        if c in rec.columns:
            rec[c] = pd.to_numeric(rec[c], errors="coerce")

    if "금액(원)" not in rec.columns and {"수량(M2)", "단가(원/M2)"} <= set(rec.columns):
        rec["금액(원)"] = (rec["수량(M2)"] * rec["단가(원/M2)"]).round(0)

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
def build_analysis_cache(q):
    df = q.copy()
    df["월"] = pd.to_datetime(df["날짜"], errors="coerce").dt.strftime("%Y-%m")
    df = df[df["월"].notna() & (df["월"] != "")].copy()
    df = safe_make_product_label(df)

    monthly_sales = (
        df.groupby(["거래처", "월"], dropna=False)["금액(원)"]
        .sum()
        .reset_index()
    )

    customer_total_monthly = monthly_sales.copy()
    customer_total_monthly["날짜축"] = pd.to_datetime(customer_total_monthly["월"] + "-01")

    product_monthly = (
        df.groupby(["거래처", "품목코드", "품목명(공식)", "품목표시", "월"], dropna=False)["금액(원)"]
        .sum()
        .reset_index()
    )
    product_monthly["날짜축"] = pd.to_datetime(product_monthly["월"] + "-01")
    product_monthly["만원라벨"] = product_monthly["금액(원)"].apply(sales_to_manwon_label)

    all_months = sorted(df["월"].unique().tolist())
    return df, monthly_sales, customer_total_monthly, product_monthly, all_months


def build_priority_results(monthly_sales, detail_df, all_months):
    if len(all_months) < 2:
        return pd.DataFrame(), [], []

    mid_idx = len(all_months) // 2
    first_half = all_months[:mid_idx] if mid_idx > 0 else [all_months[0]]
    last_half = all_months[mid_idx:] if mid_idx < len(all_months) else [all_months[-1]]

    rows = []
    for cust_name in monthly_sales["거래처"].dropna().unique():
        cust_monthly = monthly_sales[monthly_sales["거래처"] == cust_name].sort_values("월").copy()
        cust_detail = detail_df[detail_df["거래처"] == cust_name].copy()

        first_data = cust_monthly[cust_monthly["월"].isin(first_half)]["금액(원)"]
        last_data = cust_monthly[cust_monthly["월"].isin(last_half)]["금액(원)"]

        avg_first = float(first_data.mean()) if len(first_data) > 0 else 0.0
        avg_last = float(last_data.mean()) if len(last_data) > 0 else 0.0
        total_sales = float(cust_monthly["금액(원)"].sum())

        decline_amount = max(0.0, avg_first - avg_last)
        decline_rate = (decline_amount / avg_first) if avg_first > 0 else 0.0

        monthly_vals = cust_monthly["금액(원)"].astype(float).tolist()
        slope = calc_slope(monthly_vals)
        cv = calc_cv(monthly_vals)

        recent_months = all_months[-3:] if len(all_months) >= 3 else all_months
        recent_data = cust_monthly[cust_monthly["월"].isin(recent_months)].sort_values("월")
        if len(recent_data) >= 2:
            recent_trend = recent_data["금액(원)"].pct_change().replace([np.inf, -np.inf], np.nan).dropna()
            recent_neg_ratio = float((recent_trend < 0).mean()) if len(recent_trend) > 0 else 0.0
            recent_avg_change = float(recent_trend.mean()) if len(recent_trend) > 0 else 0.0
        else:
            recent_neg_ratio = 0.0
            recent_avg_change = 0.0

        first_products = cust_detail[cust_detail["월"].isin(first_half)]["품목코드"].nunique() if "품목코드" in cust_detail.columns else 0
        last_products = cust_detail[cust_detail["월"].isin(last_half)]["품목코드"].nunique() if "품목코드" in cust_detail.columns else 0
        product_decline = max(0, first_products - last_products)
        product_decline_ratio = (product_decline / max(1, first_products)) if first_products > 0 else 0.0

        rows.append({
            "거래처": str(cust_name),
            "전반부_평균매출": int(round(avg_first, 0)),
            "후반부_평균매출": int(round(avg_last, 0)),
            "실제감소액": int(round(decline_amount, 0)),
            "하락률(%)": round(decline_rate * 100, 1) if avg_first > 0 else 0.0,
            "전체_매출액": int(round(total_sales, 0)),
            "기울기": slope,
            "CV": cv,
            "최근음수비중": recent_neg_ratio,
            "최근평균증감": recent_avg_change,
            "품목감소확산도": product_decline_ratio,
        })

    base_df = pd.DataFrame(rows)
    if base_df.empty:
        return base_df, first_half, last_half

    amount_component = (
        scale_to_100(base_df["실제감소액"]) * 0.55 +
        scale_to_100(base_df["하락률(%)"]) * 0.20
    )
    base_df["감소규모점수"] = amount_component

    trend_component = (
        scale_to_100(base_df["기울기"], reverse=True) * 0.50 +
        scale_to_100(base_df["최근음수비중"]) * 0.30 +
        scale_to_100(base_df["최근평균증감"], reverse=True) * 0.20
    ) * 0.15
    base_df["추세하락점수"] = trend_component

    spread_component = scale_to_100(base_df["품목감소확산도"]) * 0.10
    base_df["품목감소점수"] = spread_component

    base_df["AI_우선순위점수"] = (
        base_df["감소규모점수"] +
        base_df["추세하락점수"] +
        base_df["품목감소점수"]
    ).round(1)

    comments = []
    for _, r in base_df.iterrows():
        msg = []
        if r["실제감소액"] > 0:
            msg.append(f"감소금액 {int(r['실제감소액']):,}원")
        if r["하락률(%)"] >= 20:
            msg.append(f"하락률 {r['하락률(%)']:.1f}%")
        if r["기울기"] < 0:
            msg.append("감소 추세 빠름")
        if r["최근음수비중"] >= 0.5:
            msg.append("최근 감소 지속")
        if r["품목감소확산도"] >= 0.3:
            msg.append("품목 이탈 동반")
        comments.append(" | ".join(msg) if msg else "추세 안정")

    base_df["분석_내역"] = comments

    result_df = base_df.sort_values(
        ["AI_우선순위점수", "실제감소액", "기울기", "품목감소확산도"],
        ascending=[False, False, True, False]
    ).reset_index(drop=True)
    result_df["순위"] = range(1, len(result_df) + 1)
    return result_df, first_half, last_half


@st.cache_data(show_spinner=False)
def build_quote_reference(q_ref):
    if q_ref is None or q_ref.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    df = q_ref.copy()
    required_cols = {"날짜", "품목코드", "거래처", "수량(M2)", "금액(원)", "단가(원/M2)"}
    if not required_cols.issubset(set(df.columns)):
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    if "품목명(공식)" not in df.columns:
        df["품목명(공식)"] = ""
    if "점착제코드" not in df.columns:
        df["점착제코드"] = ""
    if "점착제명" not in df.columns:
        df["점착제명"] = ""

    df["월"] = pd.to_datetime(df["날짜"], errors="coerce").dt.strftime("%Y-%m")
    df = df[df["월"].notna() & (df["월"] != "")].copy()

    df["품목코드"] = df["품목코드"].astype(str)
    df["거래처"] = df["거래처"].astype(str)
    df["품목명(공식)"] = df["품목명(공식)"].fillna("").astype(str)
    df["점착제코드"] = df["점착제코드"].fillna("").astype(str)
    df["점착제명"] = df["점착제명"].fillna("").astype(str)

    overview = (
        df.groupby(["품목코드", "점착제코드", "점착제명"], dropna=False)
        .agg(
            최저단가=("단가(원/M2)", "min"),
            최고단가=("단가(원/M2)", "max"),
            거래처수=("거래처", "nunique"),
            총출고횟수=("수량(M2)", "count"),
            총량_M2=("수량(M2)", "sum"),
            총매출액=("금액(원)", "sum"),
            개월수=("월", "nunique"),
        )
        .reset_index()
    )
    overview["월평균_출고량"] = np.where(overview["개월수"] > 0, overview["총량_M2"] / overview["개월수"], 0)
    overview["월평균_매출"] = np.where(overview["개월수"] > 0, overview["총매출액"] / overview["개월수"], 0)

    monthly_pc = (
        df.groupby(["품목코드", "거래처", "월"], dropna=False)
        .agg(월출고량=("수량(M2)", "sum"), 월매출=("금액(원)", "sum"))
        .reset_index()
        .sort_values(["품목코드", "거래처", "월"])
    )

    recent_unit = (
        df.dropna(subset=["단가(원/M2)"])
        .sort_values("날짜")
        .groupby(["품목코드", "거래처"], as_index=False)
        .tail(1)[["품목코드", "거래처", "단가(원/M2)", "날짜"]]
        .rename(columns={"단가(원/M2)": "최근단가", "날짜": "최근날짜"})
    )

    unit_extreme = (
        df.dropna(subset=["단가(원/M2)"])
        .groupby(["품목코드", "거래처"], dropna=False)
        .agg(
            최저단가=("단가(원/M2)", "min"),
            최고단가=("단가(원/M2)", "max"),
            총매출액=("금액(원)", "sum"),
            총량_M2=("수량(M2)", "sum"),
            개월수=("월", "nunique"),
        )
        .reset_index()
    )
    unit_extreme["월평균_출고량"] = np.where(unit_extreme["개월수"] > 0, unit_extreme["총량_M2"] / unit_extreme["개월수"], 0)
    unit_extreme["월평균_매출"] = np.where(unit_extreme["개월수"] > 0, unit_extreme["총매출액"] / unit_extreme["개월수"], 0)

    rows = []
    for (prod_code, cust_name), g in monthly_pc.groupby(["품목코드", "거래처"]):
        g = g.sort_values("월").copy()
        month_count = g["월"].nunique()
        avg_qty = float(g["월출고량"].mean()) if len(g) > 0 else 0.0
        avg_sales = float(g["월매출"].mean()) if len(g) > 0 else 0.0
        total_sales = float(g["월매출"].sum()) if len(g) > 0 else 0.0
        cv_sales = calc_cv(g["월매출"])
        slope_sales = calc_slope(g["월매출"].tolist())

        trend = "성장"
        if slope_sales < 0:
            trend = "감소"
        elif abs(slope_sales) < max(1, avg_sales * 0.02):
            trend = "안정"

        rows.append({
            "품목코드": str(prod_code),
            "거래처": str(cust_name),
            "개월수": int(month_count),
            "월평균_출고량": avg_qty,
            "월평균_매출": avg_sales,
            "총매출액": total_sales,
            "매출CV": cv_sales,
            "매출기울기": slope_sales,
            "최근추세": trend,
        })

    ref_detail = pd.DataFrame(rows)
    if ref_detail.empty:
        return overview, pd.DataFrame(), pd.DataFrame()

    ref_detail = ref_detail.merge(recent_unit, on=["품목코드", "거래처"], how="left")
    ref_detail = ref_detail.merge(
        unit_extreme[["품목코드", "거래처", "최저단가", "최고단가"]],
        on=["품목코드", "거래처"], how="left"
    )

    prod_bench = ref_detail.groupby("품목코드").agg(
        qty_p70=("월평균_출고량", lambda s: np.nanpercentile(s, 70) if len(s.dropna()) > 0 else 0),
        sales_p70=("월평균_매출", lambda s: np.nanpercentile(s, 70) if len(s.dropna()) > 0 else 0),
        unit_p30=("최근단가", lambda s: np.nanpercentile(s.dropna(), 30) if len(s.dropna()) > 0 else 0),
        cv_p30=("매출CV", lambda s: np.nanpercentile(s, 30) if len(s.dropna()) > 0 else 0),
    ).reset_index()

    ref_detail = ref_detail.merge(prod_bench, on="품목코드", how="left")

    types = []
    comments = []
    for _, r in ref_detail.iterrows():
        tags = []
        desc = []

        is_bulk = r["월평균_출고량"] >= r["qty_p70"] and r["qty_p70"] > 0
        is_high_sales = r["월평균_매출"] >= r["sales_p70"] and r["sales_p70"] > 0
        is_growth = r["최근추세"] == "성장"
        is_price_sensitive = pd.notna(r.get("최근단가")) and r["unit_p30"] > 0 and r["최근단가"] <= r["unit_p30"]
        is_small_test = r["월평균_매출"] < max(500000, r["sales_p70"] * 0.3)
        is_stable = r["매출CV"] <= r["cv_p30"]

        if is_bulk and is_high_sales:
            tags.append("대량출고형+고매출핵심형")
            desc.append("월평균 출고량과 매출 기여도가 모두 높은 핵심 거래처")
        if is_price_sensitive:
            tags.append("가격민감형")
            desc.append("최근단가가 상대적으로 낮은 편")
        if is_growth:
            tags.append("성장형")
            desc.append("최근 월매출 추세가 상승")
        if is_small_test:
            tags.append("소량테스트형")
            desc.append("월매출 규모가 낮아 테스트성 거래 가능성")
        if is_stable:
            tags.append("안정거래형")
            desc.append("월별 변동이 낮아 거래가 안정적")

        if len(tags) == 0:
            tags.append("일반거래형")
            desc.append("평균적인 거래 패턴")

        types.append(tags[0])
        comments.append(" / ".join(desc[:3]))

    ref_detail["업체성향"] = types
    ref_detail["AI분석"] = comments
    ref_detail["대표점수"] = (
        scale_to_100(ref_detail["총매출액"]).values * 0.35 +
        scale_to_100(ref_detail["월평균_출고량"]).values * 0.25 +
        scale_to_100(ref_detail["월평균_매출"]).values * 0.20 +
        scale_to_100(ref_detail["매출기울기"]).values * 0.10 +
        scale_to_100(ref_detail["최근단가"], reverse=True).values * 0.10
    )

    representative = (
        ref_detail.sort_values(["품목코드", "대표점수", "총매출액"], ascending=[True, False, False])
        .groupby("품목코드", as_index=False)
        .head(5)
        .reset_index(drop=True)
    )
    representative["최근날짜"] = pd.to_datetime(representative["최근날짜"], errors="coerce").dt.strftime("%Y-%m-%d")

    special_rows = []

    def pick_top(df_in, cond, label, sort_cols, ascending):
        sub = df_in[cond].copy()
        if sub.empty:
            return
        picked = (
            sub.sort_values(sort_cols, ascending=ascending)
            .groupby("품목코드", as_index=False)
            .head(3)
            .copy()
        )
        picked["대표구분"] = label
        special_rows.append(picked)

    pick_top(ref_detail, ref_detail["업체성향"] == "대량출고형+고매출핵심형", "대량출고형+고매출핵심형", ["품목코드", "총매출액", "월평균_출고량"], [True, False, False])
    pick_top(ref_detail, ref_detail["업체성향"] == "가격민감형", "가격민감형", ["품목코드", "최근단가", "총매출액"], [True, True, False])
    pick_top(ref_detail, ref_detail["업체성향"] == "성장형", "성장형", ["품목코드", "매출기울기", "총매출액"], [True, False, False])
    pick_top(ref_detail, ref_detail["업체성향"] == "소량테스트형", "소량테스트형", ["품목코드", "월평균_매출", "최근단가"], [True, True, True])

    lowest_unit = ref_detail.sort_values(["품목코드", "최저단가", "총매출액"], ascending=[True, True, False]).groupby("품목코드", as_index=False).head(3).copy()
    lowest_unit["대표구분"] = "최저단가 대표업체"
    special_rows.append(lowest_unit)

    highest_unit = ref_detail.sort_values(["품목코드", "최고단가", "총매출액"], ascending=[True, False, False]).groupby("품목코드", as_index=False).head(3).copy()
    highest_unit["대표구분"] = "최고단가 대표업체"
    special_rows.append(highest_unit)

    special_reference = pd.concat(special_rows, ignore_index=True) if special_rows else pd.DataFrame()
    if not special_reference.empty:
        special_reference = special_reference.sort_values(
            ["품목코드", "대표구분", "총매출액", "거래처"],
            ascending=[True, True, False, True]
        ).reset_index(drop=True)

    return overview, representative, special_reference


def draw_quote_reference_chart(special_df):
    if special_df is None or special_df.empty:
        st.info("차트로 표시할 대표 업체 데이터가 없습니다.")
        return

    required_cols = {"품목코드", "거래처", "최근단가"}
    if not required_cols.issubset(set(special_df.columns)):
        st.info("차트에 필요한 컬럼이 부족합니다.")
        return

    chart_df = special_df.copy()
    items = chart_df["품목코드"].astype(str).unique().tolist()
    color_map = build_color_map(items)
    fig = go.Figure()

    for item in items:
        sub = chart_df[chart_df["품목코드"].astype(str) == str(item)].copy()
        fig.add_trace(go.Scatter(
