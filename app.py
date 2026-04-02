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
                use_container_width=True,
                hide_index=True,
                key=key,
                disabled=disabled_cols,
                height=empty_height,
            )
        st.dataframe(display_df, use_container_width=True, hide_index=True, height=empty_height)
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
        "진행현황", "최근진행현황"
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
            # 퍼센트/비율 계열
            if any(k in col for k in ["하락률", "증감률", "비율", "변화율", "CV", "반품율"]):
                column_config[col] = st.column_config.NumberColumn(col, format="%,.1f", pinned=pinned)
            # 수량/M2 계열
            elif any(k in col for k in ["M2", "수량", "판매량", "총량", "출고량"]):
                column_config[col] = st.column_config.NumberColumn(col, format="%,.1f", pinned=pinned)
            # 점수 계열
            elif any(k in col for k in ["점수", "AI", "우선순위", "통계", "종합"]):
                column_config[col] = st.column_config.NumberColumn(col, format="%,.1f", pinned=pinned)
            # 금액/단가/매출/평균 계열
            elif any(k in col for k in ["단가", "금액", "매출", "총매출", "평균", "원"]):
                column_config[col] = st.column_config.NumberColumn(col, format="%,.0f", pinned=pinned)
            # 그 외 숫자형도 기본 콤마 표시
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
            use_container_width=True,
            height=final_height,
            num_rows="fixed",
            key=key,
            hide_index=True,
            disabled=disabled_cols,
        )

    st.dataframe(
        display_df,
        column_config=column_config,
        use_container_width=True,
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
        else:
            if pd.api.types.is_numeric_dtype(temp[c]):
                num_format[c] = "{:,.0f}"

    if num_format:
        styled = styled.format(num_format)

    final_height = height if height is not None else calc_table_height(temp, max_rows=20)
    st.dataframe(styled, use_container_width=True, height=final_height, hide_index=True)


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
    empty_cols = [group_col, time_col, value_col, "지수값"]

    if df is None or len(df) == 0:
        return pd.DataFrame(columns=empty_cols)

    required = {group_col, value_col, time_col}
    if not required.issubset(set(df.columns)):
        return pd.DataFrame(columns=empty_cols)

    temp = df.copy()

    temp[group_col] = temp[group_col].astype(str)
    temp[time_col] = pd.to_datetime(temp[time_col], errors="coerce")
    temp[value_col] = pd.to_numeric(temp[value_col], errors="coerce").fillna(0)

    temp = temp.dropna(subset=[time_col]).copy()
    if temp.empty:
        return pd.DataFrame(columns=empty_cols)

    temp = temp.sort_values([group_col, time_col]).reset_index(drop=True)

    out_frames = []
    for _, sub in temp.groupby(group_col, dropna=False):
        sub = sub[[group_col, time_col, value_col]].copy()

        base_candidates = sub[value_col].replace(0, np.nan).dropna()
        if len(base_candidates) == 0:
            sub["지수값"] = np.nan
        else:
            base = base_candidates.iloc[0]
            if pd.isna(base) or base == 0:
                sub["지수값"] = np.nan
            else:
                sub["지수값"] = (sub[value_col] / base) * 100.0

        out_frames.append(sub)

    if not out_frames:
        return pd.DataFrame(columns=empty_cols)

    result = pd.concat(out_frames, ignore_index=True)

    for col in empty_cols:
        if col not in result.columns:
            result[col] = np.nan

    result = result[empty_cols].copy()
    return result


def should_show_helper_chart(top_product_monthly):
    if top_product_monthly is None or top_product_monthly.empty:
        return False, 0.0

    if "품목표시" not in top_product_monthly.columns or "금액(원)" not in top_product_monthly.columns:
        return False, 0.0

    temp = (
        top_product_monthly.groupby("품목표시", as_index=False)["금액(원)"]
        .sum()
        .rename(columns={"금액(원)": "총매출"})
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


def infer_customer_sales_status(row):
    first_avg = float(row.get("전반부_평균매출", 0) or 0)
    last_avg = float(row.get("후반부_평균매출", 0) or 0)
    slope = float(row.get("기울기", 0) or 0)
    recent_avg_change = float(row.get("최근평균증감", 0) or 0)

    if first_avg <= 0 and last_avg > 0:
        return "신규상승"
    if first_avg > 0 and last_avg <= 0:
        return "거래중단위험"

    if last_avg > first_avg * 1.08 and slope > 0:
        return "상승추세"
    if last_avg < first_avg * 0.92 and slope < 0:
        return "감소추세"
    if recent_avg_change > 0.05:
        return "최근회복"
    if recent_avg_change < -0.05:
        return "최근둔화"
    return "안정"


def infer_customer_sales_analysis(row):
    comments = []

    status = str(row.get("진행현황", "")).strip()
    total_sales = float(row.get("전체_매출액", 0) or 0)
    decline_amt = float(row.get("실제감소액", 0) or 0)
    decline_rate = float(row.get("하락률(%)", 0) or 0)
    slope = float(row.get("기울기", 0) or 0)
    recent_neg_ratio = float(row.get("최근음수비중", 0) or 0)
    product_decline_ratio = float(row.get("품목감소확산도", 0) or 0)
    recent_avg_change = float(row.get("최근평균증감", 0) or 0)
    first_products = int(row.get("전반부_품목수", 0) or 0)
    last_products = int(row.get("후반부_품목수", 0) or 0)

    if status:
        comments.append(f"현재 흐름은 {status}")

    if total_sales >= 100000000:
        comments.append("누적 매출 규모가 큰 핵심 거래처")
    elif total_sales < 5000000:
        comments.append("누적 매출 규모가 작은 편")

    if decline_amt > 0:
        comments.append(f"전반 대비 평균 매출 {int(decline_amt):,}원 감소")
    elif decline_amt < 0:
        comments.append(f"전반 대비 평균 매출 {int(abs(decline_amt)):,}원 증가")

    if decline_rate >= 20:
        comments.append(f"하락률 {decline_rate:.1f}%로 감소폭 큼")
    elif decline_rate <= -20:
        comments.append(f"증가율 {abs(decline_rate):.1f}%로 성장폭 큼")

    if slope < 0:
        comments.append("월별 추세선 기준 하락 방향")
    elif slope > 0:
        comments.append("월별 추세선 기준 상승 방향")

    if recent_neg_ratio >= 0.67:
        comments.append("최근 구간에서 감소가 반복됨")
    elif recent_avg_change > 0.03:
        comments.append("최근 월간 증감은 회복 흐름")

    if product_decline_ratio >= 0.3:
        comments.append("매출 변화가 일부 품목이 아닌 품목 전반으로 확산")
    elif first_products > 0 and last_products > first_products:
        comments.append("거래 품목 수가 확대됨")

    if len(comments) == 0:
        comments.append("전반적으로 거래 흐름 안정")

    return " | ".join(comments[:5])


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
            x=sub["거래처"],
            y=sub["최근단가"],
            mode="markers+text",
            name=str(item),
            text=sub["대표구분"] if "대표구분" in sub.columns else "",
            textposition="top center",
            marker=dict(size=11, color=color_map[str(item)], line=dict(width=1, color="white")),
            hovertemplate=(
                f"품목코드: {item}<br>"
                "거래처: %{x}<br>"
                "최근단가: %{y:,.0f}원/M2<br>"
                "구분: %{text}<extra></extra>"
            )
        ))

    fig.update_layout(
        height=460,
        margin=dict(l=20, r=45, t=30, b=90),
        xaxis=dict(title="", tickangle=-35, automargin=True),
        yaxis=dict(title="최근단가(원/M2)", tickformat=",.0f", automargin=True),
        legend=dict(orientation="h", yanchor="top", y=-0.28, x=0, xanchor="left")
    )
    st.plotly_chart(fig, use_container_width=True, key="quote_reference_chart_main")


@st.cache_data(show_spinner=False)
def build_return_decline_item_analysis(q):
    if q is None or q.empty:
        return {}

    df = q.copy()
    df["월"] = pd.to_datetime(df["날짜"], errors="coerce").dt.strftime("%Y-%m")
    df = df[df["월"].notna() & (df["월"] != "")].copy()
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


@st.cache_data(show_spinner=False)
def build_customer_sales_analysis(q):
    if q is None or q.empty:
        return {
            "customer_summary": pd.DataFrame(),
            "customer_monthly": pd.DataFrame(),
            "customer_item_summary": pd.DataFrame(),
            "customer_item_monthly": pd.DataFrame(),
            "all_months": [],
        }

    df = q.copy()
    df["날짜"] = pd.to_datetime(df["날짜"], errors="coerce")
    df["월"] = df["날짜"].dt.strftime("%Y-%m")
    df = df[df["월"].notna() & (df["월"] != "")].copy()
    df = safe_make_product_label(df)

    for c in ["금액(원)", "수량(M2)", "단가(원/M2)", "가로폭(mm)"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        else:
            df[c] = 0

    all_months = sorted(df["월"].dropna().astype(str).unique().tolist())

    customer_summary = (
        df.groupby("거래처", as_index=False)
        .agg(
            총매출액=("금액(원)", "sum"),
            총판매량=("수량(M2)", "sum"),
            품목수=("품목표시", "nunique"),
            최근일자=("날짜", "max"),
        )
        .sort_values(["총매출액", "거래처"], ascending=[False, True])
        .reset_index(drop=True)
    )

    customer_monthly = (
        df.groupby(["거래처", "월"], as_index=False)
        .agg(매출액=("금액(원)", "sum"))
        .sort_values(["거래처", "월"])
        .reset_index(drop=True)
    )
    customer_monthly["날짜축"] = pd.to_datetime(customer_monthly["월"] + "-01", errors="coerce")

    customer_item_monthly = (
        df.groupby(["거래처", "품목표시", "월"], as_index=False)
        .agg(
            매출액=("금액(원)", "sum"),
            판매량=("수량(M2)", "sum"),
        )
        .sort_values(["거래처", "품목표시", "월"])
        .reset_index(drop=True)
    )

    item_summary = (
        df.groupby(["거래처", "품목표시"], as_index=False)
        .agg(
            총매출액=("금액(원)", "sum"),
            총판매량=("수량(M2)", "sum"),
            평균단가=("단가(원/M2)", "mean"),
            최근가로폭=("가로폭(mm)", lambda s: s.dropna().iloc[-1] if len(s.dropna()) > 0 else np.nan),
            최근일자=("날짜", "max"),
        )
    )

    if not customer_item_monthly.empty:
        base_map = (
            customer_item_monthly.sort_values(["거래처", "품목표시", "월"])
            .groupby(["거래처", "품목표시"], as_index=False)
            .first()[["거래처", "품목표시", "매출액"]]
            .rename(columns={"매출액": "기준월매출"})
        )
        last_map = (
            customer_item_monthly.sort_values(["거래처", "품목표시", "월"])
            .groupby(["거래처", "품목표시"], as_index=False)
            .last()[["거래처", "품목표시", "매출액"]]
            .rename(columns={"매출액": "최근월매출"})
        )
        item_summary = item_summary.merge(base_map, on=["거래처", "품목표시"], how="left")
        item_summary = item_summary.merge(last_map, on=["거래처", "품목표시"], how="left")
    else:
        item_summary["기준월매출"] = 0
        item_summary["최근월매출"] = 0

    recent_price = (
        df.sort_values("날짜")
        .groupby(["거래처", "품목표시"], as_index=False)
        .tail(1)[["거래처", "품목표시", "단가(원/M2)"]]
        .rename(columns={"단가(원/M2)": "최근단가"})
    )
    item_summary = item_summary.merge(recent_price, on=["거래처", "품목표시"], how="left")

    width_hist = (
        df.groupby(["거래처", "품목표시"], as_index=False)["가로폭(mm)"]
        .apply(lambda s: ", ".join([
            str(int(v)) if pd.notna(v) and float(v).is_integer() else str(v)
            for v in pd.Series(s).dropna().unique()
        ]))
        .reset_index()
        .rename(columns={"가로폭(mm)": "가로폭이력"})
    )
    item_summary = item_summary.merge(width_hist, on=["거래처", "품목표시"], how="left")

    item_summary["기준월매출"] = pd.to_numeric(item_summary["기준월매출"], errors="coerce").fillna(0)
    item_summary["최근월매출"] = pd.to_numeric(item_summary["최근월매출"], errors="coerce").fillna(0)
    item_summary["변화율(%)"] = np.where(
        item_summary["기준월매출"] == 0,
        np.nan,
        ((item_summary["최근월매출"] - item_summary["기준월매출"]) / item_summary["기준월매출"]) * 100.0
    )

    # ── 거래처 분석 강화: tab4 스타일의 분석 자료 생성 ──
    analysis_rows = []
    if len(all_months) >= 2 and not customer_monthly.empty:
        mid_idx = len(all_months) // 2
        first_half = all_months[:mid_idx] if mid_idx > 0 else [all_months[0]]
        last_half = all_months[mid_idx:] if mid_idx < len(all_months) else [all_months[-1]]

        for cust_name in customer_monthly["거래처"].dropna().astype(str).unique():
            cust_month = customer_monthly[customer_monthly["거래처"].astype(str) == str(cust_name)].sort_values("월").copy()
            cust_detail = df[df["거래처"].astype(str) == str(cust_name)].copy()

            first_data = cust_month[cust_month["월"].isin(first_half)]["매출액"]
            last_data = cust_month[cust_month["월"].isin(last_half)]["매출액"]

            avg_first = float(first_data.mean()) if len(first_data) > 0 else 0.0
            avg_last = float(last_data.mean()) if len(last_data) > 0 else 0.0
            delta = avg_last - avg_first
            decline_amount = avg_first - avg_last
            total_sales = float(cust_month["매출액"].sum())

            if avg_first > 0:
                decline_rate = (decline_amount / avg_first) * 100.0
            else:
                decline_rate = 0.0

            monthly_vals = cust_month["매출액"].astype(float).tolist()
            slope = calc_slope(monthly_vals)
            cv = calc_cv(monthly_vals)

            recent_months = all_months[-3:] if len(all_months) >= 3 else all_months
            recent_data = cust_month[cust_month["월"].isin(recent_months)].sort_values("월")
            if len(recent_data) >= 2:
                recent_trend = recent_data["매출액"].pct_change().replace([np.inf, -np.inf], np.nan).dropna()
                recent_neg_ratio = float((recent_trend < 0).mean()) if len(recent_trend) > 0 else 0.0
                recent_avg_change = float(recent_trend.mean()) if len(recent_trend) > 0 else 0.0
            else:
                recent_neg_ratio = 0.0
                recent_avg_change = 0.0

            first_products = cust_detail[cust_detail["월"].isin(first_half)]["품목표시"].nunique() if "품목표시" in cust_detail.columns else 0
            last_products = cust_detail[cust_detail["월"].isin(last_half)]["품목표시"].nunique() if "품목표시" in cust_detail.columns else 0
            product_decline = max(0, first_products - last_products)
            product_decline_ratio = (product_decline / max(1, first_products)) if first_products > 0 else 0.0

            analysis_rows.append({
                "거래처": str(cust_name),
                "전체_매출액": int(round(total_sales, 0)),
                "전반부_평균매출": int(round(avg_first, 0)),
                "후반부_평균매출": int(round(avg_last, 0)),
                "평균증감액": int(round(delta, 0)),
                "실제감소액": int(round(max(0.0, decline_amount), 0)),
                "하락률(%)": round(decline_rate, 1),
                "기울기": slope,
                "CV": cv,
                "최근음수비중": recent_neg_ratio,
                "최근평균증감": recent_avg_change,
                "전반부_품목수": int(first_products),
                "후반부_품목수": int(last_products),
                "품목감소확산도": round(product_decline_ratio, 3),
            })

    customer_analysis = pd.DataFrame(analysis_rows)

    if not customer_analysis.empty:
        base_score = (
            scale_to_100(customer_analysis["전체_매출액"]) * 0.35 +
            scale_to_100(customer_analysis["후반부_평균매출"]) * 0.20 +
            scale_to_100(customer_analysis["기울기"]) * 0.15 +
            scale_to_100(customer_analysis["최근평균증감"]) * 0.10 +
            scale_to_100(customer_analysis["품목감소확산도"], reverse=True) * 0.10 +
            scale_to_100(customer_analysis["CV"], reverse=True) * 0.10
        ).round(1)

        customer_analysis["AI_평가점수"] = base_score
        customer_analysis["진행현황"] = customer_analysis.apply(infer_customer_sales_status, axis=1)
        customer_analysis["분석_내역"] = customer_analysis.apply(infer_customer_sales_analysis, axis=1)

        customer_analysis["AI분석"] = customer_analysis.apply(
            lambda r: (
                f"진행현황 {r.get('진행현황', '안정')} | "
                f"매출규모 {int(r.get('전체_매출액', 0)):,}원 | "
                f"전반 {int(r.get('전반부_평균매출', 0)):,}원 → 후반 {int(r.get('후반부_평균매출', 0)):,}원 | "
                f"{'감소 압력 존재' if float(r.get('기울기', 0)) < 0 else '상승/회복 흐름'}"
            ),
            axis=1
        )

        customer_summary = customer_summary.merge(
            customer_analysis,
            on="거래처",
            how="left"
        )
    else:
        customer_summary["전체_매출액"] = customer_summary["총매출액"]
        customer_summary["전반부_평균매출"] = 0
        customer_summary["후반부_평균매출"] = 0
        customer_summary["평균증감액"] = 0
        customer_summary["실제감소액"] = 0
        customer_summary["하락률(%)"] = 0.0
        customer_summary["기울기"] = 0.0
        customer_summary["CV"] = 0.0
        customer_summary["최근음수비중"] = 0.0
        customer_summary["최근평균증감"] = 0.0
        customer_summary["전반부_품목수"] = 0
        customer_summary["후반부_품목수"] = 0
        customer_summary["품목감소확산도"] = 0.0
        customer_summary["AI_평가점수"] = 50.0
        customer_summary["진행현황"] = "안정"
        customer_summary["분석_내역"] = "분석 기간 부족 또는 거래 데이터 부족"
        customer_summary["AI분석"] = "분석 기간 부족"

    item_summary = item_summary.sort_values(["거래처", "총매출액", "품목표시"], ascending=[True, False, True]).reset_index(drop=True)

    return {
        "customer_summary": customer_summary,
        "customer_monthly": customer_monthly,
        "customer_item_summary": item_summary,
        "customer_item_monthly": customer_item_monthly,
        "all_months": all_months,
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

dept_col = "담당부서" if "담당부서" in rec.columns else ("영업담당부서" if "영업담당부서" in rec.columns else None)
manager_col = "담당자" if "담당자" in rec.columns else None

sel_dept = st.sidebar.multiselect(
    "담당부서",
    sorted_unique(rec[dept_col]) if dept_col else [],
    placeholder="Choose options"
)
sel_manager = st.sidebar.multiselect(
    "담당자",
    sorted_unique(rec[manager_col]) if manager_col else [],
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

date_min = pd.to_datetime(rec["날짜"].min()) if "날짜" in rec.columns else None
date_max = pd.to_datetime(rec["날짜"].max()) if "날짜" in rec.columns else None

if date_min is not None and pd.notna(date_min) and date_max is not None and pd.notna(date_max):
    picked = st.sidebar.date_input("기간", [date_min.date(), date_max.date()])
    if isinstance(picked, (list, tuple)) and len(picked) == 2:
        sdate, edate = picked
    else:
        sdate = edate = None
else:
    sdate = edate = None

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
    q = q[(q["날짜"] >= pd.to_datetime(sdate)) & (q["날짜"] <= pd.to_datetime(edate))]

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
        if "날짜" in q1.columns:
            q1["월"] = pd.to_datetime(q1["날짜"], errors="coerce").dt.strftime("%Y-%m")
        else:
            q1["월"] = ""

        for c in ["거래처", "품목코드", "점착제코드"]:
            if c in q1.columns:
                q1[c] = q1[c].astype(str)

        cols = ["거래처", "품목코드", "점착제코드", "점착제명", "가로폭이력", "최근날짜", "최근단가"]
        uc = [c for c in cols if c in q1.columns]

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
        if "날짜" in q2.columns:
            q2["월"] = pd.to_datetime(q2["날짜"], errors="coerce").dt.strftime("%Y-%m")
        else:
            q2["월"] = ""

        for c in ["거래처", "품목코드"]:
            if c in q2.columns:
                q2[c] = q2[c].astype(str)

        cols = ["품목코드", "거래처", "최근날짜", "최근단가"]
        uc = [c for c in cols if c in q2.columns]

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

                customer_data = detail_df[detail_df["거래처"] == selected_cust_name].copy()
                customer_total_monthly = customer_total_monthly_all[
                    customer_total_monthly_all["거래처"] == selected_cust_name
                ].copy()
                product_monthly = product_monthly_all[
                    product_monthly_all["거래처"] == selected_cust_name
                ].copy()

                if customer_data.empty:
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
                    )

                    contribution_rows = []
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

                    fig_total = apply_mobile_friendly_line_layout(
                        fig_total,
                        customer_total_monthly["날짜축"],
                        y_title="매출액(원)",
                        height=430
                    )
                    fig_total.update_layout(title="1️⃣ 업체 전체 월별 매출 추이")
                    st.plotly_chart(fig_total, use_container_width=True, key=f"tab4_total_{selected_cust_name}")

                    top_contrib = contribution_df.head(12).copy()
                    if not top_contrib.empty:
                        fig_contrib = go.Figure()
                        fig_contrib.add_trace(go.Bar(
                            x=top_contrib["품목표시"],
                            y=top_contrib["감소액"],
                            marker_color=[
                                "#d62728" if v > 0 else "#2ca02c"
                                for v in top_contrib["감소액"]
                            ],
                            text=[f"{v:,}" for v in top_contrib["감소액"]],
                            textposition="outside",
                            name="감소액",
                            hovertemplate="품목: %{x}<br>감소액: %{y:,.0f}원<extra></extra>",
                        ))
                        fig_contrib.update_layout(
                            title="2️⃣ 품목별 매출 감소 기여도 (감소 큰 순)",
                            height=450,
                            yaxis_tickformat=",",
                            xaxis_title="품목",
                            yaxis_title="감소액(원)",
                            margin=dict(l=20, r=45, t=35, b=120),
                            xaxis=dict(automargin=True),
                            yaxis=dict(automargin=True)
                        )
                        st.plotly_chart(fig_contrib, use_container_width=True, key=f"tab4_contrib_{selected_cust_name}")

                    top_products = contribution_df.head(5)["품목표시"].astype(str).tolist() if not contribution_df.empty else []
                    top_product_monthly = product_monthly[
                        product_monthly["품목표시"].astype(str).isin(top_products)
                    ].copy() if len(top_products) > 0 else pd.DataFrame()

                    if not top_product_monthly.empty:
                        fig_products = go.Figure()
                        pos_map = make_text_position_map(top_products)

                        for prod_label in top_products:
                            sub = top_product_monthly[
                                top_product_monthly["품목표시"] == prod_label
                            ].sort_values("날짜축").copy()

                            if sub.empty:
                                continue

                            fig_products.add_trace(go.Scatter(
                                x=sub["날짜축"],
                                y=sub["금액(원)"],
                                mode="lines+markers+text",
                                name=prod_label,
                                line=dict(width=3),
                                marker=dict(size=8),
                                text=sub["만원라벨"],
                                textposition=pos_map.get(prod_label, "top center"),
                                textfont=dict(size=9),
                                cliponaxis=False,
                                hovertemplate=f"품목: {prod_label}<br>월: %{{x|%Y-%m}}<br>매출: %{{y:,.0f}}원<br>만원단위: %{{text}}<extra></extra>",
                            ))

                        fig_products = apply_mobile_friendly_line_layout(
                            fig_products,
                            top_product_monthly["날짜축"],
                            y_title="매출액(원)",
                            height=460
                        )
                        fig_products.update_layout(
                            title="3️⃣ 감소 주도 품목 월별 매출 추이 (Top 5)",
                            legend=dict(
                                orientation="h",
                                yanchor="bottom",
                                y=-0.35,
                                x=0,
                                xanchor="left"
                            )
                        )
                        st.plotly_chart(fig_products, use_container_width=True, key=f"tab4_products_{selected_cust_name}")
                        st.caption("※ 각 포인트 값은 만원 단위입니다. 예: 4,500 = 4천5백만원")

                        show_helper, scale_ratio = should_show_helper_chart(top_product_monthly)

                        if show_helper:
                            st.markdown("#### 3-1) 변화율 보조 그래프")

                            indexed_df = make_indexed_series(
                                top_product_monthly,
                                group_col="품목표시",
                                value_col="금액(원)",
                                time_col="날짜축"
                            )

                            required_cols = {"품목표시", "날짜축", "지수값"}
                            if indexed_df is not None and not indexed_df.empty and required_cols.issubset(set(indexed_df.columns)):
                                indexed_df["지수라벨"] = indexed_df["지수값"].apply(
                                    lambda v: "" if pd.isna(v) else f"{v:,.0f}"
                                )

                                fig_idx = go.Figure()
                                for prod_label in top_products:
                                    sub = indexed_df[
                                        indexed_df["품목표시"].astype(str) == str(prod_label)
                                    ].sort_values("날짜축").copy()

                                    if sub.empty:
                                        continue

                                    fig_idx.add_trace(go.Scatter(
                                        x=sub["날짜축"],
                                        y=sub["지수값"],
                                        mode="lines+markers+text",
                                        name=prod_label,
                                        line=dict(width=2),
                                        marker=dict(size=7),
                                        text=sub["지수라벨"],
                                        textposition=pos_map.get(prod_label, "top center"),
                                        textfont=dict(size=9),
                                        cliponaxis=False,
                                        hovertemplate=f"품목: {prod_label}<br>월: %{{x|%Y-%m}}<br>지수: %{{y:,.1f}}<extra></extra>",
                                        showlegend=True
                                    ))

                                if len(fig_idx.data) > 0:
                                    fig_idx.add_hline(
                                        y=100,
                                        line_dash="dash",
                                        line_color="gray",
                                        opacity=0.7
                                    )

                                    fig_idx = apply_mobile_friendly_line_layout(
                                        fig_idx,
                                        indexed_df["날짜축"],
                                        y_title="지수값",
                                        height=460
                                    )
                                    fig_idx.update_layout(
                                        title=f"품목간 매출 규모 차이가 커서 추가 표시 (최대/최소 약 {scale_ratio:.1f}배)",
                                        legend=dict(
                                            orientation="h",
                                            yanchor="bottom",
                                            y=-0.35,
                                            x=0,
                                            xanchor="left"
                                        )
                                    )
                                    st.plotly_chart(fig_idx, use_container_width=True, key=f"tab4_helper_{selected_cust_name}")
                                else:
                                    st.info("변화율 보조 그래프를 생성할 데이터가 없습니다.")
                            else:
                                st.info("변화율 보조 그래프를 생성할 데이터가 없습니다.")

                    if not contribution_df.empty:
                        comp_df = contribution_df.head(12).copy()
                        fig_bar = go.Figure()
                        fig_bar.add_trace(go.Bar(
                            x=comp_df["품목표시"],
                            y=comp_df["전반부_평균"],
                            name="전반부 평균",
                            marker_color="#3498db",
                            text=[f"{v:,.0f}" for v in comp_df["전반부_평균"]],
                            textposition="outside"
                        ))
                        fig_bar.add_trace(go.Bar(
                            x=comp_df["품목표시"],
                            y=comp_df["후반부_평균"],
                            name="후반부 평균",
                            marker_color="#e74c3c",
                            text=[f"{v:,.0f}" for v in comp_df["후반부_평균"]],
                            textposition="outside"
                        ))
                        fig_bar.update_layout(
                            title="4️⃣ 품목별 전반부 vs 후반부 평균 매출 비교",
                            barmode="group",
                            height=500,
                            yaxis_tickformat=",",
                            xaxis_title="품목",
                            yaxis_title="매출액(원)",
                            margin=dict(l=20, r=45, t=35, b=120),
                            xaxis=dict(automargin=True),
                            yaxis=dict(automargin=True)
                        )
                        st.plotly_chart(fig_bar, use_container_width=True, key=f"tab4_compare_{selected_cust_name}")

with tab5:
    st.subheader("📊 거래처별 매출 분석")
    st.caption("기존 매출하락/품목감소 분석에는 영향 없이 별도 탭으로 분리된 거래처 분석입니다.")

    if q.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
    else:
        pack = build_customer_sales_analysis(q)
        customer_summary = pack["customer_summary"]
        customer_monthly = pack["customer_monthly"]
        customer_item_summary = pack["customer_item_summary"]
        customer_item_monthly = pack["customer_item_monthly"]
        all_months = pack["all_months"]

        if customer_summary.empty:
            st.info("거래처 분석 데이터가 없습니다.")
        else:
            st.markdown("### 1) 매출분석 자료")
            show_summary = customer_summary.copy()
            if "최근일자" in show_summary.columns:
                show_summary["최근일자"] = pd.to_datetime(show_summary["최근일자"], errors="coerce").dt.strftime("%Y-%m-%d")

            summary_cols = [
                "거래처", "AI_평가점수", "진행현황", "전체_매출액",
                "전반부_평균매출", "후반부_평균매출", "평균증감액",
                "실제감소액", "하락률(%)", "전반부_품목수", "후반부_품목수",
                "품목감소확산도", "총판매량", "품목수", "최근일자", "분석_내역", "AI분석"
            ]
            summary_cols = [c for c in summary_cols if c in show_summary.columns]

            clean_and_safe_display(
                show_summary[summary_cols].sort_values(
                    ["AI_평가점수", "전체_매출액", "거래처"],
                    ascending=[False, False, True]
                ),
                pinned_cols=["거래처"],
                text_cols=["거래처", "진행현황", "최근일자", "분석_내역", "AI분석"]
            )

            customer_options = customer_summary["거래처"].dropna().astype(str).tolist()

            if customer_options:
                selected_customer = st.selectbox(
                    "분석할 업체를 선택하세요",
                    options=customer_options,
                    key="customer_sales_analysis_select_v2"
                )

                st.markdown("### 2) 업체별 상세 분석")
                st.markdown("#### 업체 전체 월별 매출 추이")

                cust_month = customer_monthly[customer_monthly["거래처"] == selected_customer].copy()
                if not cust_month.empty:
                    month_axis = build_month_axis_frame(all_months if all_months else cust_month["월"].tolist())
                    series = align_monthly_series(month_axis, cust_month[["월", "매출액"]], "매출액")

                    fig = go.Figure()
                    fig.add_trace(go.Scatter(
                        x=series["날짜축"],
                        y=series["매출액"],
                        mode="lines+markers+text",
                        name="월매출",
                        line=dict(color="#1f77b4", width=3),
                        marker=dict(size=8),
                        text=[sales_to_manwon_label(v) for v in series["매출액"]],
                        textposition="top center",
                        textfont=dict(size=10, color="#1f77b4"),
                        hovertemplate="월: %{x|%Y-%m}<br>매출: %{y:,.0f}원<br>만원단위: %{text}<extra></extra>",
                    ))

                    if len(series) >= 2:
                        x_num = np.arange(len(series))
                        y_num = pd.to_numeric(series["매출액"], errors="coerce").fillna(0).values.astype(float)
                        coef = np.polyfit(x_num, y_num, 1)
                        trend = coef[0] * x_num + coef[1]
                        fig.add_trace(go.Scatter(
                            x=series["날짜축"],
                            y=trend,
                            mode="lines",
                            name="추세선",
                            line=dict(color="red", dash="dash", width=2),
                            hoverinfo="skip",
                        ))

                    fig = apply_mobile_friendly_line_layout(
                        fig,
                        series["날짜축"],
                        y_title="매출액(원)",
                        height=430
                    )
                    fig.update_layout(title="업체 전체 월별 매출 추이")
                    st.plotly_chart(fig, use_container_width=True, key=f"tab5_total_{selected_customer}")
                else:
                    st.info("해당 업체의 월별 매출 데이터가 없습니다.")

                st.markdown("#### 품목별 변화율 상세 (총매출순)")
                item_detail = customer_item_summary[customer_item_summary["거래처"] == selected_customer].copy()

                if not item_detail.empty:
                    show_cols = [
                        c for c in [
                            "품목표시", "총매출액", "총판매량", "최근단가",
                            "기준월매출", "최근월매출", "변화율(%)", "가로폭이력", "최근일자"
                        ] if c in item_detail.columns
                    ]
                    show_df = item_detail[show_cols].sort_values(["총매출액", "품목표시"], ascending=[False, True]).reset_index(drop=True)

                    if "최근일자" in show_df.columns:
                        show_df["최근일자"] = pd.to_datetime(show_df["최근일자"], errors="coerce").dt.strftime("%Y-%m-%d")

                    clean_and_safe_display(
                        show_df,
                        pinned_cols=["품목표시"],
                        text_cols=["품목표시", "가로폭이력", "최근일자"]
                    )

                    total_sales = pd.to_numeric(show_df["총매출액"], errors="coerce").fillna(0).sum()
                    top70 = show_df.copy()
                    if total_sales > 0:
                        top70["누적매출"] = pd.to_numeric(top70["총매출액"], errors="coerce").fillna(0).cumsum()
                        top70["누적비중"] = top70["누적매출"] / total_sales
                        top70 = top70[top70["누적비중"] <= 0.70].copy()
                        if top70.empty:
                            top70 = show_df.head(1).copy()
                    top70 = top70.head(5).copy()

                    if not top70.empty:
                        st.markdown("#### 매출 70% 해당 품목 월별 매출 추이 (TOP5)")
                        top_items = top70["품목표시"].astype(str).tolist()

                        top_month = customer_item_monthly[
                            (customer_item_monthly["거래처"] == selected_customer) &
                            (customer_item_monthly["품목표시"].astype(str).isin(top_items))
                        ].copy()

                        if not top_month.empty:
                            month_axis = build_month_axis_frame(all_months)
                            fig_top = go.Figure()

                            for item_name in top_items:
                                sub = top_month[top_month["품목표시"].astype(str) == str(item_name)].copy()
                                aligned = align_monthly_series(month_axis, sub[["월", "매출액"]], "매출액")
                                fig_top.add_trace(go.Scatter(
                                    x=aligned["날짜축"],
                                    y=aligned["매출액"],
                                    mode="lines+markers",
                                    name=item_name
                                ))

                            fig_top = apply_mobile_friendly_line_layout(
                                fig_top,
                                month_axis["날짜축"],
                                y_title="매출액(원)",
                                height=430
                            )
                            fig_top.update_layout(
                                title="매출 70% 해당 품목 월별 매출 추이 (TOP5)",
                                hovermode="x unified"
                            )
                            st.plotly_chart(fig_top, use_container_width=True)

                            st.markdown("#### 변화율 보조 그래프")
                            aux = top_month.copy()
                            aux["날짜축"] = pd.to_datetime(aux["월"] + "-01", errors="coerce")
                            indexed_df = make_indexed_series(aux, "품목표시", "매출액", "날짜축")

                            required_cols = {"품목표시", "날짜축", "지수값"}
                            if indexed_df is None or indexed_df.empty or not required_cols.issubset(set(indexed_df.columns)):
                                st.info("변화율 보조 그래프를 생성할 데이터가 없습니다.")
                            else:
                                fig_idx = go.Figure()

                                for item_name in top_items:
                                    sub = indexed_df[indexed_df["품목표시"].astype(str) == str(item_name)].copy()
                                    if sub.empty:
                                        continue

                                    fig_idx.add_trace(go.Scatter(
                                        x=sub["날짜축"],
                                        y=sub["지수값"],
                                        mode="lines+markers",
                                        name=item_name
                                    ))

                                if len(fig_idx.data) == 0:
                                    st.info("변화율 보조 그래프를 생성할 데이터가 없습니다.")
                                else:
                                    fig_idx.add_hline(y=100, line_dash="dash", line_color="gray")
                                    fig_idx = apply_mobile_friendly_line_layout(
                                        fig_idx,
                                        indexed_df["날짜축"],
                                        y_title="지수(기준월=100)",
                                        height=430
                                    )
                                    fig_idx.update_layout(hovermode="x unified")
                                    st.plotly_chart(fig_idx, use_container_width=True)

                    product_options = show_df["품목표시"].dropna().astype(str).tolist()
                    if product_options:
                        selected_product = st.selectbox(
                            "원자료를 확인할 품목을 선택하세요",
                            options=product_options,
                            key="customer_item_raw_select_v2"
                        )

                        raw_cols = [c for c in [
                            "날짜", "거래처", "품목코드", "품목명(공식)", "품목표시",
                            "점착제코드", "가로폭(mm)", "수량(M2)", "단가(원/M2)", "금액(원)", "비고"
                        ] if c in q.columns]

                        raw_df = q.copy()
                        raw_df = safe_make_product_label(raw_df)

                        if "거래처" in raw_df.columns:
                            raw_df = raw_df[raw_df["거래처"].astype(str).str.strip() == str(selected_customer).strip()]
                        if "품목표시" in raw_df.columns:
                            raw_df = raw_df[raw_df["품목표시"].astype(str).str.strip() == str(selected_product).strip()]

                        st.markdown("#### 선택 품목 원자료")
                        if raw_df.empty:
                            st.info("선택한 품목의 원자료가 없습니다.")
                        else:
                            if "날짜" in raw_df.columns:
                                raw_df = raw_df.sort_values("날짜", ascending=False)
                            clean_and_safe_display(
                                raw_df[raw_cols],
                                pinned_cols=["날짜", "거래처", "품목코드"],
                                text_cols=["날짜", "거래처", "품목코드", "품목명(공식)", "품목표시", "점착제코드", "비고"]
                            )
                else:
                    st.info("해당 업체의 품목별 분석 데이터가 없습니다.")

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
                        st.metric("품목하락점수", f"{rr['품목하락점수']:,.1f}")
                    with c2:
                        st.metric("감소금액", f"{int(rr['감소금액']):,} 원")
                    with c3:
                        st.metric("하락률", f"{rr['하락률(%)']:,.1f}%")
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
                                st.metric("순위", f"{int(sr['순위']):,}")
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
        text_cols=["거래처", "품목코드", "점착제코드", "점착제명", "가로폭이력", "최근날짜", "월", "비고", "담당부서", "영업담당부서", "담당자"],
    )
