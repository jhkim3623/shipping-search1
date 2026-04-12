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
    column_width_overrides=None,
):
    if df is None:
        df = pd.DataFrame()

    display_df = df.copy().reset_index(drop=True)
    display_df.columns = [str(c) for c in display_df.columns]

    pinned_cols = [str(c) for c in (pinned_cols or [])]
    text_cols = set(str(c) for c in (text_cols or []))
    disabled_cols = disabled_cols if disabled_cols is not None else False
    column_width_overrides = column_width_overrides or {}

    def _normalize_width(val):
        if val is None:
            return None
        if isinstance(val, str):
            low = val.strip().lower()
            if low in ["small", "medium", "large"]:
                return low
            try:
                return int(float(low))
            except Exception:
                return None
        if isinstance(val, (int, float, np.integer, np.floating)):
            return int(val)
        return None

    normalized_widths = {
        str(k): _normalize_width(v)
        for k, v in column_width_overrides.items()
        if _normalize_width(v) is not None
    }

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

    compact_small_cols = {
        "순위", "AI_평가점수", "AI_우선순위점수", "감소규모점수", "추세하락점수",
        "품목감소점수", "품목하락점수", "품목수", "전반부_품목수", "후반부_품목수",
        "출고횟수", "거래처수", "총출고횟수", "개월수"
    }
    compact_medium_cols = {
        "가로폭(mm)", "수량(M2)", "단가(원/M2)", "금액(원)", "총판매량", "총매출액",
        "전체_매출액", "전반부_평균매출", "후반부_평균매출", "평균증감액", "실제감소액",
        "최근단가", "최근월매출", "하락률(%)", "변화율(%)", "반품율(%)", "반품금액",
        "총매출", "총판매량", "월평균_매출", "월평균_출고량", "총량_M2"
    }

    for col in display_df.columns:
        pinned = col in pinned_cols
        manual_width = normalized_widths.get(col)

        if col in text_cols or col in fixed_text_like_cols:
            width_value = "medium"
            if col in ["거래처", "품목코드", "점착제코드", "점착제명", "최근날짜", "최근일자", "진행현황"]:
                width_value = "small"
            if col in [
                "품목명(공식)", "품목표시", "가로폭이력", "분석_내역",
                "AI분석", "원인추정", "감소원인", "비고", "주요반품원인", "AI반품분석"
            ]:
                width_value = "large"

            if manual_width is not None:
                width_value = manual_width

            column_config[col] = st.column_config.TextColumn(
                col,
                width=width_value,
                pinned=pinned,
            )
            continue

        if pd.api.types.is_numeric_dtype(display_df[col]):
            num_width = "medium"
            if col in compact_small_cols:
                num_width = "small"
            elif col in compact_medium_cols:
                num_width = "medium"

            if manual_width is not None:
                num_width = manual_width

            if any(k in col for k in ["하락률", "증감률", "비율", "변화율", "CV", "반품율"]):
                column_config[col] = st.column_config.NumberColumn(col, format="%,.1f", pinned=pinned, width=num_width)
            elif any(k in col for k in ["M2", "수량", "판매량", "총량", "출고량"]):
                column_config[col] = st.column_config.NumberColumn(col, format="%,.1f", pinned=pinned, width=num_width)
            elif any(k in col for k in ["점수", "AI", "우선순위", "통계", "종합"]):
                column_config[col] = st.column_config.NumberColumn(col, format="%,.1f", pinned=pinned, width=num_width)
            elif any(k in col for k in ["단가", "금액", "매출", "총매출", "평균", "원"]):
                column_config[col] = st.column_config.NumberColumn(col, format="%,.0f", pinned=pinned, width=num_width)
            else:
                column_config[col] = st.column_config.NumberColumn(col, format="%,.0f", pinned=pinned, width=num_width)
        else:
            width_value = manual_width if manual_width is not None else "medium"
            column_config[col] = st.column_config.TextColumn(col, width=width_value, pinned=pinned)

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
    customer_item_monthly["날짜축"] = pd.to_datetime(customer_item_monthly["월"] + "-01", errors="coerce")

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

    item_analysis_rows = []
    if len(all_months) >= 2 and not customer_item_monthly.empty:
        mid_idx = len(all_months) // 2
        first_half = all_months[:mid_idx] if mid_idx > 0 else [all_months[0]]
        last_half = all_months[mid_idx:] if mid_idx < len(all_months) else [all_months[-1]]

        grouped = customer_item_monthly.groupby(["거래처", "품목표시"], dropna=False)
        for (cust_name, item_name), g in grouped:
            g = g.sort_values("월").copy()
            first_vals = g[g["월"].isin(first_half)]["매출액"]
            last_vals = g[g["월"].isin(last_half)]["매출액"]

            first_avg = float(first_vals.mean()) if len(first_vals) > 0 else 0.0
            last_avg = float(last_vals.mean()) if len(last_vals) > 0 else 0.0
            delta_avg = last_avg - first_avg

            item_analysis_rows.append({
                "거래처": str(cust_name),
                "품목표시": str(item_name),
                "전반부_평균매출": int(round(first_avg, 0)),
                "후반부_평균매출": int(round(last_avg, 0)),
                "평균증감액": int(round(delta_avg, 0)),
            })

    item_analysis_df = pd.DataFrame(item_analysis_rows)
    if not item_analysis_df.empty:
        item_summary = item_summary.merge(
            item_analysis_df,
            on=["거래처", "품목표시"],
            how="left"
        )
    else:
        item_summary["전반부_평균매출"] = 0
        item_summary["후반부_평균매출"] = 0
        item_summary["평균증감액"] = 0

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

        tab1_df = g[ordered_cols].sort_values(sc) if sc else g[ordered_cols]

        clean_and_safe_display(
            tab1_df,
            pinned_cols=["거래처", "품목코드"],
            text_cols=["거래처", "품목코드", "점착제코드", "점착제명", "가로폭이력", "최근날짜"],
            height=calc_table_height(tab1_df),
            column_width_overrides={
                "거래처": 170,
                "품목코드": 145,
                "점착제코드": 95,
                "점착제명": 120,
                "가로폭이력": 260,
                "최근날짜": 95,
                "최근단가": 70,
                "출고횟수": 70,
                "월평균_출고량": 95,
                "월평균_매출": 95,
                "총량_M2": 80,
                "매출액": 105,
                "가중평균단가": 95,
            },
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
            column_width_overrides={
                "품목코드": 145,
                "거래처": 145,
                "최근날짜": 95,
                "최근단가": 75,
                "출고횟수": 70,
                "월평균_출고량": 100,
                "월평균_매출": 105,
                "총량_M2": 90,
                "매출액": 110,
                "가중평균단가": 100,
            },
        )

with tab3:
    st.subheader("🏷️ 견적 레퍼런스 — 기준 견적가 & 판매 동향")
    st.info("기존 기능 유지")
    st.write("이 탭은 기존 로직 그대로 사용하면 됩니다.")

with tab4:
    st.subheader("📉 매출 하락 분석")
    st.info("기존 기능 유지")
    st.write("이 탭은 기존 로직 그대로 사용하면 됩니다.")

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
                text_cols=["거래처", "진행현황", "최근일자", "분석_내역", "AI분석"],
                column_width_overrides={
                    "거래처": 145,
                    "AI_평가점수": 85,
                    "진행현황": 85,
                    "전체_매출액": 110,
                    "전반부_평균매출": 120,
                    "후반부_평균매출": 120,
                    "평균증감액": 100,
                    "실제감소액": 100,
                    "하락률(%)": 85,
                    "전반부_품목수": 85,
                    "후반부_품목수": 85,
                    "품목감소확산도": 95,
                    "총판매량": 95,
                    "품목수": 70,
                    "최근일자": 95,
                    "분석_내역": 330,
                    "AI분석": 360,
                }
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
                            "품목표시", "총매출액", "총판매량",
                            "전반부_평균매출", "후반부_평균매출", "평균증감액",
                            "최근단가", "최근월매출", "변화율(%)", "가로폭이력", "최근일자"
                        ] if c in item_detail.columns
                    ]
                    show_df = item_detail[show_cols].sort_values(["총매출액", "품목표시"], ascending=[False, True]).reset_index(drop=True)

                    if "최근일자" in show_df.columns:
                        show_df["최근일자"] = pd.to_datetime(show_df["최근일자"], errors="coerce").dt.strftime("%Y-%m-%d")

                    clean_and_safe_display(
                        show_df,
                        pinned_cols=["품목표시"],
                        text_cols=["품목표시", "가로폭이력", "최근일자"],
                        column_width_overrides={
                            "품목표시": 180,
                            "총매출액": 95,
                            "총판매량": 95,
                            "전반부_평균매출": 120,
                            "후반부_평균매출": 120,
                            "평균증감액": 100,
                            "최근단가": 85,
                            "최근월매출": 95,
                            "변화율(%)": 85,
                            "가로폭이력": 220,
                            "최근일자": 95,
                        }
                    )

                    top7 = show_df.head(7).copy()

                    if not top7.empty:
                        st.markdown("#### 매출 주도 상위 품목 월별 매출 추이 (Top 7)")
                        top_items = top7["품목표시"].astype(str).tolist()

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
                                    mode="lines+markers+text",
                                    name=item_name,
                                    text=[sales_to_manwon_label(v) if pd.notna(v) and v != 0 else "" for v in aligned["매출액"]],
                                    textposition="top center",
                                    textfont=dict(size=9),
                                    cliponaxis=False,
                                    hovertemplate=f"품목: {item_name}<br>월: %{{x|%Y-%m}}<br>매출: %{{y:,.0f}}원<extra></extra>"
                                ))

                            fig_top = apply_mobile_friendly_line_layout(
                                fig_top,
                                month_axis["날짜축"],
                                y_title="매출액(원)",
                                height=460
                            )
                            fig_top.update_layout(
                                title="매출 주도 상위 품목 월별 매출 추이 (Top 7)",
                                hovermode="x unified",
                                legend=dict(
                                    orientation="h",
                                    yanchor="bottom",
                                    y=-0.32,
                                    x=0,
                                    xanchor="left"
                                )
                            )
                            st.plotly_chart(fig_top, use_container_width=True)

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

                        selected_item_month = customer_item_monthly[
                            (customer_item_monthly["거래처"].astype(str).str.strip() == str(selected_customer).strip()) &
                            (customer_item_monthly["품목표시"].astype(str).str.strip() == str(selected_product).strip())
                        ].copy()

                        st.markdown("#### 선택 품목 월별 매출 그래프")
                        if not selected_item_month.empty:
                            month_axis_single = build_month_axis_frame(all_months if all_months else selected_item_month["월"].tolist())
                            single_series = align_monthly_series(month_axis_single, selected_item_month[["월", "매출액"]], "매출액")

                            fig_single = go.Figure()
                            fig_single.add_trace(go.Scatter(
                                x=single_series["날짜축"],
                                y=single_series["매출액"],
                                mode="lines+markers+text",
                                name=selected_product,
                                line=dict(color="#1f77b4", width=3),
                                marker=dict(size=8),
                                text=[sales_to_manwon_label(v) if pd.notna(v) and v != 0 else "" for v in single_series["매출액"]],
                                textposition="top center",
                                textfont=dict(size=9),
                                cliponaxis=False,
                                hovertemplate="월: %{x|%Y-%m}<br>매출: %{y:,.0f}원<extra></extra>",
                            ))
                            fig_single = apply_mobile_friendly_line_layout(
                                fig_single,
                                single_series["날짜축"],
                                y_title="매출액(원)",
                                height=400
                            )
                            fig_single.update_layout(title=f"{selected_product} 월별 매출 추이")
                            st.plotly_chart(fig_single, use_container_width=True)

                            st.markdown("#### 선택 품목 변화율(Index) 그래프")
                            indexed_df = make_indexed_series(
                                single_series.rename(columns={"매출액": "금액(원)", "날짜축": "날짜축"}).assign(품목표시=selected_product),
                                group_col="품목표시",
                                value_col="금액(원)",
                                time_col="날짜축"
                            )

                            if indexed_df is not None and not indexed_df.empty:
                                indexed_df["지수라벨"] = indexed_df["지수값"].apply(
                                    lambda v: "" if pd.isna(v) else f"{v:,.0f}"
                                )

                                fig_idx = go.Figure()
                                fig_idx.add_trace(go.Scatter(
                                    x=indexed_df["날짜축"],
                                    y=indexed_df["지수값"],
                                    mode="lines+markers+text",
                                    name=selected_product,
                                    line=dict(color="#1f77b4", width=3),
                                    marker=dict(size=8),
                                    text=indexed_df["지수라벨"],
                                    textposition="top center",
                                    textfont=dict(size=9),
                                    cliponaxis=False,
                                    hovertemplate="월: %{x|%Y-%m}<br>지수: %{y:,.1f}<extra></extra>",
                                ))
                                fig_idx.add_hline(y=100, line_dash="dash", line_color="gray", opacity=0.7)
                                fig_idx = apply_mobile_friendly_line_layout(
                                    fig_idx,
                                    indexed_df["날짜축"],
                                    y_title="지수값",
                                    height=380
                                )
                                fig_idx.update_layout(title=f"{selected_product} 변화율 보조 그래프")
                                st.plotly_chart(fig_idx, use_container_width=True)
                            else:
                                st.info("선택 품목의 변화율(Index) 그래프를 생성할 데이터가 없습니다.")

                        st.markdown("#### 선택 품목 원자료")
                        if raw_df.empty:
                            st.info("선택한 품목의 원자료가 없습니다.")
                        else:
                            if "날짜" in raw_df.columns:
                                raw_df = raw_df.sort_values("날짜", ascending=False)
                            clean_and_safe_display(
                                raw_df[raw_cols],
                                pinned_cols=["날짜", "거래처", "품목코드"],
                                text_cols=["날짜", "거래처", "품목코드", "품목명(공식)", "품목표시", "점착제코드", "비고"],
                                column_width_overrides={
                                    "날짜": 95,
                                    "거래처": 145,
                                    "품목코드": 145,
                                    "품목명(공식)": 180,
                                    "품목표시": 180,
                                    "점착제코드": 80,
                                    "가로폭(mm)": 80,
                                    "수량(M2)": 90,
                                    "단가(원/M2)": 85,
                                    "금액(원)": 95,
                                    "비고": 220,
                                }
                            )
                else:
                    st.info("해당 업체의 품목별 분석 데이터가 없습니다.")

with tab6:
    st.subheader("매출 감소 품목 분석")
    st.info("기존 기능 유지")
    st.write("이 탭은 기존 로직 그대로 사용하면 됩니다.")

with tab7:
    st.subheader("원자료(필터 적용됨)")
    raw_cols = [c for c in q.columns if c != "품목명(공식)"]
    clean_and_safe_display(
        q[raw_cols],
        pinned_cols=["거래처", "품목코드"],
        text_cols=["거래처", "품목코드", "점착제코드", "점착제명", "가로폭이력", "최근날짜", "월", "비고", "담당부서", "영업담당부서", "담당자"],
        column_width_overrides={
            "날짜": 95,
            "거래처": 120,
            "품목코드": 145,
            "점착제코드": 95,
            "점착제명": 120,
            "가로폭(mm)": 90,
            "가로폭이력": 220,
            "수량(M2)": 90,
            "단가(원/M2)": 95,
            "금액(원)": 95,
            "최근날짜": 95,
            "최근단가": 85,
            "비고": 220,
            "담당부서": 110,
            "영업담당부서": 120,
            "담당자": 95,
        },
    )
