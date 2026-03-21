import os
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# =========================================================
# 기본 설정
# =========================================================
st.set_page_config(
    page_title="출고 이력 검색 / 매출 하락 분석",
    layout="wide"
)

st.title("📦 출고 이력 검색 / 매출 하락 분석")


# =========================================================
# 공통 유틸
# =========================================================
MONEY_KEYWORDS = [
    "매출", "금액", "공급가액", "합계", "후반부", "전반부", "감소액", "평균", "누적", "실적"
]
RATE_KEYWORDS = [
    "비율", "률", "증감", "하락", "변화", "기여도", "지수", "CV"
]
TEXT_PRIORITY_COLS = [
    "거래처", "품목코드", "품목명", "품목명(공식)", "점착제코드", "점착제명",
    "가로폭이력", "최근날짜", "월", "구분", "분석_내역"
]


def norm_colname(col):
    return str(col).strip()


def to_yymm(x):
    try:
        dt = pd.to_datetime(x, errors="coerce")
        if pd.isna(dt):
            return str(x)
        return dt.strftime("%y%m")
    except Exception:
        return str(x)


def format_manwon(v):
    try:
        if pd.isna(v):
            return ""
        return f"{int(round(float(v) / 10000, 0)):,}"
    except Exception:
        return ""


def safe_num(s):
    return pd.to_numeric(s, errors="coerce")


def infer_amount_column(df):
    candidates = [
        "매출", "매출액", "공급가액", "금액", "합계금액", "출고금액", "판매금액"
    ]
    for c in candidates:
        if c in df.columns:
            return c

    numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    for c in numeric_cols:
        if any(k in str(c) for k in ["매출", "금액", "공급가액"]):
            return c
    return numeric_cols[0] if numeric_cols else None


def infer_date_column(df):
    candidates = ["날짜", "출고일", "일자", "매출일", "date", "Date"]
    for c in candidates:
        if c in df.columns:
            return c
    return None


def infer_customer_column(df):
    candidates = ["거래처", "업체", "고객", "고객사", "매출처"]
    for c in candidates:
        if c in df.columns:
            return c
    return None


def infer_product_code_column(df):
    candidates = ["품목코드", "제품코드", "품번", "품목"]
    for c in candidates:
        if c in df.columns:
            return c
    return None


def infer_product_name_column(df):
    candidates = ["품목명(공식)", "품목명", "제품명", "상품명"]
    for c in candidates:
        if c in df.columns:
            return c
    return None


def coerce_month_col(df, date_col):
    temp = df.copy()
    if date_col and date_col in temp.columns:
        temp[date_col] = pd.to_datetime(temp[date_col], errors="coerce")
        temp = temp[~temp[date_col].isna()].copy()
        temp["월"] = temp[date_col].dt.to_period("M").astype(str)
        temp["월표시"] = temp[date_col].dt.strftime("%y%m")
    else:
        if "월" in temp.columns:
            temp["월"] = temp["월"].astype(str)
            temp["월표시"] = temp["월"].apply(to_yymm)
        else:
            temp["월"] = ""
            temp["월표시"] = ""
    return temp


def calc_cv(series):
    s = pd.to_numeric(series, errors="coerce").dropna()
    if len(s) == 0:
        return 0.0
    mean_v = s.mean()
    std_v = s.std(ddof=0)
    if mean_v == 0 or pd.isna(mean_v):
        return 0.0
    return float(std_v / mean_v)


def calc_trend_slope(values):
    s = pd.to_numeric(pd.Series(values), errors="coerce").fillna(0)
    n = len(s)
    if n <= 1:
        return 0.0
    x = np.arange(n)
    slope = np.polyfit(x, s.values, 1)[0]
    return float(slope)


def scale_0_100(series, reverse=False):
    s = pd.to_numeric(series, errors="coerce").fillna(0)
    if s.nunique() <= 1:
        return pd.Series([50.0] * len(s), index=s.index)
    min_v = s.min()
    max_v = s.max()
    scaled = ((s - min_v) / (max_v - min_v)) * 100
    if reverse:
        scaled = 100 - scaled
    return scaled.clip(0, 100)


def make_text_position_map(items):
    positions = ["top center", "bottom center", "middle right", "middle left"]
    return {item: positions[i % len(positions)] for i, item in enumerate(items)}


def make_indexed_series(df, group_col, value_col, time_col):
    temp = df.copy().sort_values([group_col, time_col])

    def _calc(g):
        base_candidates = g[value_col].replace(0, np.nan).dropna()
        if len(base_candidates) == 0:
            g["지수값"] = np.nan
        else:
            base = base_candidates.iloc[0]
            if pd.isna(base) or base == 0:
                g["지수값"] = np.nan
            else:
                g["지수값"] = (g[value_col] / base) * 100
        return g

    return temp.groupby(group_col, group_keys=False).apply(_calc)


# =========================================================
# 안전한 표 출력
# =========================================================
def clean_and_safe_display(
    df,
    height=None,
    key=None,
    editable=False,
    pinned_cols=None,
    text_cols=None,
    disabled_cols=None
):
    if df is None:
        df = pd.DataFrame()

    display_df = df.copy().reset_index(drop=True)
    display_df.columns = [str(c) for c in display_df.columns]

    pinned_cols = [str(c) for c in (pinned_cols or [])]
    text_cols = set(str(c) for c in (text_cols or [])) | set(TEXT_PRIORITY_COLS)

    for col in display_df.columns:
        s = display_df[col]

        if pd.api.types.is_datetime64_any_dtype(s):
            display_df[col] = pd.to_datetime(s, errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
            continue

        if col in text_cols:
            display_df[col] = s.astype(str).replace(
                ["nan", "NaN", "None", "<NA>", "NaT"], ""
            ).fillna("")
            continue

        num_s = pd.to_numeric(s, errors="coerce")
        if num_s.notna().sum() >= max(1, int(len(display_df) * 0.6)):
            display_df[col] = num_s.replace([np.inf, -np.inf], np.nan).fillna(0)
        else:
            display_df[col] = s.astype(str).replace(
                ["nan", "NaN", "None", "<NA>", "NaT"], ""
            ).fillna("")

    column_config = {}
    for col in display_df.columns:
        pinned = col in pinned_cols

        if col in text_cols or not pd.api.types.is_numeric_dtype(display_df[col]):
            column_config[col] = st.column_config.TextColumn(
                col,
                width="large" if col in ["품목명(공식)", "품목명", "가로폭이력", "분석_내역"] else "medium",
                pinned=pinned
            )
        else:
            if any(k in col for k in RATE_KEYWORDS):
                fmt = "%.1f"
            elif any(k in col for k in MONEY_KEYWORDS):
                fmt = "%,d"
            else:
                if pd.api.types.is_float_dtype(display_df[col]):
                    fmt = "%,.1f"
                else:
                    fmt = "%,d"

            column_config[col] = st.column_config.NumberColumn(
                col,
                format=fmt,
                pinned=pinned
            )

    if editable and key:
        safe_editor_df = display_df.copy()
        return st.data_editor(
            safe_editor_df,
            column_config=column_config,
            width="stretch",
            height=height or 420,
            num_rows="fixed",
            key=key,
            hide_index=True,
            disabled=disabled_cols if disabled_cols is not None else False
        )
    else:
        st.dataframe(
            display_df,
            column_config=column_config,
            width="stretch",
            height=height or 420,
            hide_index=True
        )
        return None


# =========================================================
# 데이터 로드
# =========================================================
@st.cache_data(show_spinner=False)
def load_excel_file(file_bytes):
    xls = pd.ExcelFile(BytesIO(file_bytes))
    sheets = xls.sheet_names

    preferred_sheet = None
    for s in sheets:
        if any(k in s for k in ["출고", "매출", "data", "Data", "Sheet1", "원자료"]):
            preferred_sheet = s
            break
    if preferred_sheet is None:
        preferred_sheet = sheets[0]

    df = pd.read_excel(BytesIO(file_bytes), sheet_name=preferred_sheet)
    df.columns = [norm_colname(c) for c in df.columns]

    # 흔한 문자열 결측치 정리
    for c in df.columns:
        if df[c].dtype == "object":
            df[c] = df[c].replace(
                ["nan", "NaN", "None", "<NA>", "NaT"], np.nan
            )

    return df, preferred_sheet, sheets


def preprocess_data(df_raw):
    df = df_raw.copy()
    df.columns = [norm_colname(c) for c in df.columns]

    customer_col = infer_customer_column(df)
    product_code_col = infer_product_code_column(df)
    product_name_col = infer_product_name_column(df)
    amount_col = infer_amount_column(df)
    date_col = infer_date_column(df)

    if customer_col is None:
        df["거래처"] = "미분류"
        customer_col = "거래처"

    if product_code_col is None:
        df["품목코드"] = "미지정"
        product_code_col = "품목코드"

    if product_name_col is None:
        df["품목명(공식)"] = df[product_code_col].astype(str)
        product_name_col = "품목명(공식)"

    if amount_col is None:
        # 숫자 컬럼이 없으면 강제로 생성
        df["매출"] = 0
        amount_col = "매출"

    df[amount_col] = pd.to_numeric(df[amount_col], errors="coerce").fillna(0)

    # 날짜/월
    df = coerce_month_col(df, date_col)

    # 텍스트 정리
    df[customer_col] = df[customer_col].astype(str).fillna("").replace("nan", "")
    df[product_code_col] = df[product_code_col].astype(str).fillna("").replace("nan", "")
    df[product_name_col] = df[product_name_col].astype(str).fillna("").replace("nan", "")

    # 표준 컬럼명 별칭 생성
    if customer_col != "거래처":
        df["거래처"] = df[customer_col]
    if product_code_col != "품목코드":
        df["품목코드"] = df[product_code_col]
    if product_name_col != "품목명(공식)":
        df["품목명(공식)"] = df[product_name_col]
    if amount_col != "매출":
        df["매출"] = df[amount_col]

    # 표시용 날짜
    if date_col and date_col in df.columns:
        df["최근날짜"] = pd.to_datetime(df[date_col], errors="coerce").dt.strftime("%Y-%m-%d")
    else:
        df["최근날짜"] = ""

    return {
        "df": df,
        "customer_col": "거래처",
        "product_code_col": "품목코드",
        "product_name_col": "품목명(공식)",
        "amount_col": "매출",
        "date_col": date_col,
    }


# =========================================================
# 우선순위 분석
# =========================================================
def build_customer_priority(df):
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    month_order = sorted(df["월"].dropna().astype(str).unique().tolist())
    if len(month_order) == 0:
        return pd.DataFrame(), pd.DataFrame()

    monthly = (
        df.groupby(["거래처", "월", "월표시"], as_index=False)["매출"]
        .sum()
        .sort_values(["거래처", "월"])
    )

    month_list = sorted(monthly["월"].dropna().unique().tolist())
    half = max(1, len(month_list) // 2)
    first_half = month_list[:half]
    second_half = month_list[half:] if len(month_list[half:]) > 0 else month_list[-1:]

    summary = (
        monthly.groupby("거래처")
        .agg(
            전반부매출=("매출", lambda x: monthly.loc[x.index][monthly.loc[x.index, "월"].isin(first_half)]["매출"].sum()),
            후반부매출=("매출", lambda x: monthly.loc[x.index][monthly.loc[x.index, "월"].isin(second_half)]["매출"].sum()),
            월수=("매출", "count"),
            평균월매출=("매출", "mean"),
        )
        .reset_index()
    )

    # 통계 지표 계산
    stat_rows = []
    for cust, g in monthly.groupby("거래처"):
        g = g.sort_values("월")
        vals = g["매출"].tolist()
        slope = calc_trend_slope(vals)
        cv = calc_cv(vals)

        first_avg = g[g["월"].isin(first_half)]["매출"].mean() if len(first_half) > 0 else 0
        second_avg = g[g["월"].isin(second_half)]["매출"].mean() if len(second_half) > 0 else 0

        decrease_amt = max(first_avg - second_avg, 0)
        decrease_rate = ((first_avg - second_avg) / first_avg * 100) if first_avg not in [0, np.nan] and first_avg != 0 else 0

        stat_rows.append({
            "거래처": cust,
            "CV": cv,
            "추세기울기": slope,
            "전반부평균": first_avg if not pd.isna(first_avg) else 0,
            "후반부평균": second_avg if not pd.isna(second_avg) else 0,
            "감소액": decrease_amt if not pd.isna(decrease_amt) else 0,
            "감소율": decrease_rate if not pd.isna(decrease_rate) else 0,
        })

    stat_df = pd.DataFrame(stat_rows)
    result = summary.merge(stat_df, on="거래처", how="left")

    # 60%: 매출 감소, 20%: 통계 감소추이, 20%: AI 분석
    result["매출감소점수"] = (
        scale_0_100(result["감소액"]) * 0.7 +
        scale_0_100(result["감소율"]) * 0.3
    )

    # 통계 감소추이 점수: 감소 slope 클수록, CV 적당히 높을수록
    # slope가 더 음수일수록 하락이므로 reverse 사용
    result["통계추세점수"] = (
        scale_0_100(result["추세기울기"], reverse=True) * 0.7 +
        scale_0_100(result["CV"]) * 0.3
    )

    # AI 분석 점수(업무 해석형 룰 기반)
    ai_score = []
    ai_comment = []

    for _, row in result.iterrows():
        score = 0
        comments = []

        if row["감소액"] > 0:
            score += 35
            comments.append("전반부 대비 후반부 매출 감소 확인")
        if row["감소율"] >= 20:
            score += 25
            comments.append("감소율이 높아 우선 대응 필요")
        if row["추세기울기"] < 0:
            score += 20
            comments.append("월별 추세가 하락 방향")
        if row["CV"] >= 0.4:
            score += 10
            comments.append("월간 변동성 큼")
        if row["후반부평균"] < row["전반부평균"] * 0.7:
            score += 10
            comments.append("최근 매출 레벨 급감")

        score = min(score, 100)

        if len(comments) == 0:
            comments.append("전반적 매출 흐름 안정")

        ai_score.append(score)
        ai_comment.append(" / ".join(comments))

    result["AI점수"] = ai_score
    result["분석_내역"] = ai_comment

    result["우선순위점수"] = (
        result["매출감소점수"] * 0.6 +
        result["통계추세점수"] * 0.2 +
        result["AI점수"] * 0.2
    ).round(1)

    result = result.sort_values(
        ["우선순위점수", "감소액", "감소율"],
        ascending=[False, False, False]
    ).reset_index(drop=True)

    result["순위"] = np.arange(1, len(result) + 1)

    return result, monthly


def build_product_analysis(df):
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    monthly = (
        df.groupby(["품목코드", "품목명(공식)", "월", "월표시"], as_index=False)["매출"]
        .sum()
        .sort_values(["품목코드", "월"])
    )

    month_list = sorted(monthly["월"].dropna().unique().tolist())
    if len(month_list) == 0:
        return pd.DataFrame(), monthly

    half = max(1, len(month_list) // 2)
    first_half = month_list[:half]
    second_half = month_list[half:] if len(month_list[half:]) > 0 else month_list[-1:]

    rows = []
    for (pcode, pname), g in monthly.groupby(["품목코드", "품목명(공식)"]):
        g = g.sort_values("월")
        first_avg = g[g["월"].isin(first_half)]["매출"].mean() if len(first_half) > 0 else 0
        second_avg = g[g["월"].isin(second_half)]["매출"].mean() if len(second_half) > 0 else 0
        first_sum = g[g["월"].isin(first_half)]["매출"].sum()
        second_sum = g[g["월"].isin(second_half)]["매출"].sum()
        decline_amt = max(first_sum - second_sum, 0)
        decline_rate = ((first_avg - second_avg) / first_avg * 100) if first_avg not in [0, np.nan] and first_avg != 0 else 0
        slope = calc_trend_slope(g["매출"].tolist())
        cv = calc_cv(g["매출"].tolist())

        rows.append({
            "품목코드": pcode,
            "품목명(공식)": pname,
            "전반부매출": first_sum if not pd.isna(first_sum) else 0,
            "후반부매출": second_sum if not pd.isna(second_sum) else 0,
            "감소액": decline_amt if not pd.isna(decline_amt) else 0,
            "감소율": decline_rate if not pd.isna(decline_rate) else 0,
            "추세기울기": slope,
            "CV": cv,
        })

    product_summary = pd.DataFrame(rows)
    product_summary = product_summary.sort_values(
        ["감소액", "감소율"],
        ascending=[False, False]
    ).reset_index(drop=True)

    return product_summary, monthly


# =========================================================
# 그래프
# =========================================================
def draw_customer_sales_trend(customer_monthly):
    if customer_monthly is None or customer_monthly.empty:
        st.info("표시할 데이터가 없습니다.")
        return

    plot_df = customer_monthly.copy().sort_values("월")
    plot_df["월표시"] = plot_df["월표시"].apply(str)
    plot_df["라벨"] = plot_df["매출"].apply(format_manwon)

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=plot_df["월표시"],
            y=plot_df["매출"],
            mode="lines+markers+text",
            text=plot_df["라벨"],
            textposition="top center",
            textfont=dict(size=10),
            line=dict(width=3, color="#1f77b4"),
            marker=dict(size=8),
            name="월별 매출",
            hovertemplate="월: %{x}<br>매출: %{y:,.0f}원<br>만원단위: %{text}<extra></extra>"
        )
    )

    # 추세선
    if len(plot_df) >= 2:
        x_num = np.arange(len(plot_df))
        z = np.polyfit(x_num, plot_df["매출"].values, 1)
        trend = np.poly1d(z)(x_num)
        fig.add_trace(
            go.Scatter(
                x=plot_df["월표시"],
                y=trend,
                mode="lines",
                line=dict(width=2, dash="dash", color="red"),
                name="추세선",
                hoverinfo="skip"
            )
        )

    fig.update_layout(
        height=420,
        margin=dict(l=20, r=20, t=20, b=20),
        xaxis=dict(title="월(YYMM)", type="category"),
        yaxis=dict(title="매출액(원)", tickformat=",.0f", separatethousands=True),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    )
    st.plotly_chart(fig, use_container_width=True)
    st.caption("※ 각 포인트의 숫자는 만원 단위입니다. 예: 4,500 = 4천5백만원")


def draw_decline_contribution_bar(product_summary, top_n=10):
    if product_summary is None or product_summary.empty:
        st.info("감소 기여 품목 데이터가 없습니다.")
        return pd.DataFrame()

    top_items = product_summary.head(top_n).copy()
    top_items["표시명"] = top_items["품목명(공식)"].fillna(top_items["품목코드"]).astype(str)

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=top_items["표시명"],
            y=top_items["감소액"],
            text=top_items["감소액"].apply(lambda x: f"{int(x):,}"),
            textposition="outside",
            marker_color="#d62728",
            hovertemplate="품목: %{x}<br>감소액: %{y:,.0f}원<extra></extra>"
        )
    )
    fig.update_layout(
        height=420,
        margin=dict(l=20, r=20, t=20, b=80),
        xaxis=dict(title="품목", tickangle=-25),
        yaxis=dict(title="감소액(원)", tickformat=",.0f", separatethousands=True),
    )
    st.plotly_chart(fig, use_container_width=True)
    return top_items


def draw_decline_driver_monthly_chart(top_decline_items_monthly):
    st.markdown("#### 감소주도 품목 월별 추이")

    if top_decline_items_monthly is None or top_decline_items_monthly.empty:
        st.info("감소주도 품목 데이터가 없습니다.")
        return

    decline_plot_df = top_decline_items_monthly.copy()
    decline_plot_df["월표시"] = decline_plot_df["월표시"].astype(str)

    item_col = "품목명(공식)" if "품목명(공식)" in decline_plot_df.columns else "품목코드"
    value_col = "매출"

    decline_plot_df[value_col] = pd.to_numeric(
        decline_plot_df[value_col], errors="coerce"
    ).fillna(0)

    item_order = decline_plot_df[item_col].fillna("").astype(str).unique().tolist()
    text_pos_map = make_text_position_map(item_order)

    decline_plot_df["금액라벨"] = decline_plot_df[value_col].apply(format_manwon)

    # 메인 차트
    fig_main = go.Figure()

    for item in item_order:
        sub = decline_plot_df[decline_plot_df[item_col].astype(str) == str(item)].copy()
        sub = sub.sort_values("월")

        fig_main.add_trace(
            go.Scatter(
                x=sub["월표시"],
                y=sub[value_col],
                mode="lines+markers+text",
                name=str(item),
                text=sub["금액라벨"],
                textposition=text_pos_map.get(str(item), "top center"),
                textfont=dict(size=10),
                line=dict(width=3),
                marker=dict(size=8),
                hovertemplate=(
                    f"{item_col}: {str(item)}<br>"
                    "월: %{x}<br>"
                    "매출: %{y:,.0f}원<br>"
                    "만원단위: %{text}<extra></extra>"
                ),
            )
        )

    fig_main.update_layout(
        height=520,
        margin=dict(l=20, r=20, t=20, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        xaxis=dict(title="월(YYMM)", type="category"),
        yaxis=dict(title="매출액(원)", tickformat=",.0f", separatethousands=True),
    )

    st.plotly_chart(fig_main, use_container_width=True)
    st.caption("※ 각 포인트 값은 만원 단위로 표시했습니다. 예: 4,500 = 4천5백만원")

    # 보조 차트: 지수화
    st.markdown("##### 변화율 보조 그래프 (첫 달=100 기준)")
    st.caption("금액 차이가 큰 품목이 함께 있을 때, 작은 금액 품목의 변화도 비교하기 쉽도록 지수화했습니다.")

    indexed_df = make_indexed_series(
        decline_plot_df,
        group_col=item_col,
        value_col=value_col,
        time_col="월"
    )
    indexed_df["지수라벨"] = indexed_df["지수값"].apply(
        lambda v: "" if pd.isna(v) else f"{v:.0f}"
    )

    fig_index = go.Figure()
    for item in item_order:
        sub = indexed_df[indexed_df[item_col].astype(str) == str(item)].copy()
        sub = sub.sort_values("월")

        fig_index.add_trace(
            go.Scatter(
                x=sub["월표시"],
                y=sub["지수값"],
                mode="lines+markers+text",
                name=str(item),
                text=sub["지수라벨"],
                textposition=text_pos_map.get(str(item), "top center"),
                textfont=dict(size=9),
                line=dict(width=2),
                marker=dict(size=7),
                hovertemplate=(
                    f"{item_col}: {str(item)}<br>"
                    "월: %{x}<br>"
                    "지수: %{y:.1f}<br>(첫 달=100)<extra></extra>"
                ),
                showlegend=False
            )
        )

    fig_index.add_hline(y=100, line_dash="dash", line_color="gray", opacity=0.7)
    fig_index.update_layout(
        height=420,
        margin=dict(l=20, r=20, t=10, b=20),
        xaxis=dict(title="월(YYMM)", type="category"),
        yaxis=dict(title="지수(첫 달=100)"),
    )
    st.plotly_chart(fig_index, use_container_width=True)


def draw_half_compare_bar(product_summary, top_n=10):
    st.markdown("#### 전반부 vs 후반부 직접 비교")

    if product_summary is None or product_summary.empty:
        st.info("비교할 데이터가 없습니다.")
        return

    comp = product_summary.head(top_n).copy()
    comp["표시명"] = comp["품목명(공식)"].fillna(comp["품목코드"]).astype(str)

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=comp["표시명"],
            y=comp["전반부매출"],
            name="전반부",
            marker_color="#1f77b4",
            text=comp["전반부매출"].apply(lambda x: f"{int(x):,}"),
            textposition="outside"
        )
    )
    fig.add_trace(
        go.Bar(
            x=comp["표시명"],
            y=comp["후반부매출"],
            name="후반부",
            marker_color="#ff7f0e",
            text=comp["후반부매출"].apply(lambda x: f"{int(x):,}"),
            textposition="outside"
        )
    )
    fig.update_layout(
        barmode="group",
        height=460,
        margin=dict(l=20, r=20, t=20, b=80),
        xaxis=dict(title="품목", tickangle=-25),
        yaxis=dict(title="매출액(원)", tickformat=",.0f", separatethousands=True),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    )
    st.plotly_chart(fig, use_container_width=True)


# =========================================================
# 파일 로드
# =========================================================
DEFAULT_FILE = "data.xlsx"

uploaded_file = st.sidebar.file_uploader(
    "📂 엑셀 파일 업로드",
    type=["xlsx"]
)

file_bytes = None
file_name = None

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    file_name = uploaded_file.name
elif os.path.exists(DEFAULT_FILE):
    with open(DEFAULT_FILE, "rb") as f:
        file_bytes = f.read()
    file_name = DEFAULT_FILE

if file_bytes is None:
    st.warning("엑셀 파일을 업로드하거나 프로젝트 폴더에 data.xlsx 파일을 넣어주세요.")
    st.stop()

raw_df, used_sheet, all_sheets = load_excel_file(file_bytes)
prep = preprocess_data(raw_df)
df = prep["df"]


# =========================================================
# 사이드바 필터
# =========================================================
st.sidebar.markdown("### ⚙️ 필터")

month_values = sorted(df["월"].dropna().astype(str).unique().tolist())
customer_values = sorted(df["거래처"].dropna().astype(str).unique().tolist())
product_values = sorted(df["품목명(공식)"].dropna().astype(str).unique().tolist())

selected_months = st.sidebar.multiselect(
    "월 선택",
    month_values,
    default=month_values
)

selected_customers = st.sidebar.multiselect(
    "거래처 선택",
    customer_values,
    default=customer_values
)

selected_products = st.sidebar.multiselect(
    "품목 선택",
    product_values,
    default=product_values
)

q = df.copy()
if selected_months:
    q = q[q["월"].astype(str).isin(selected_months)]
if selected_customers:
    q = q[q["거래처"].astype(str).isin(selected_customers)]
if selected_products:
    q = q[q["품목명(공식)"].astype(str).isin(selected_products)]

st.caption(f"사용 파일: {file_name} / 시트: {used_sheet} / 원본행수: {len(df):,} / 필터적용행수: {len(q):,}")


# =========================================================
# 탭 구성
# =========================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "거래처 리스트",
    "품목별 월간 매출 분석",
    "매출 추이 분석",
    "우선 대응업체",
    "원자료"
])


# =========================================================
# TAB1 거래처 리스트
# =========================================================
with tab1:
    st.subheader("거래처 리스트")

    cust_summary = (
        q.groupby("거래처", as_index=False)
        .agg(
            총매출=("매출", "sum"),
            건수=("매출", "count"),
            평균매출=("매출", "mean")
        )
        .sort_values("총매출", ascending=False)
        .reset_index(drop=True)
    )
    cust_summary["순위"] = np.arange(1, len(cust_summary) + 1)
    cust_summary = cust_summary[["순위", "거래처", "총매출", "건수", "평균매출"]]

    clean_and_safe_display(
        cust_summary,
        height=500,
        pinned_cols=["순위", "거래처"],
        text_cols=["거래처"]
    )


# =========================================================
# TAB2 품목별 월간 매출 분석
# =========================================================
with tab2:
    st.subheader("품목별 월간 매출 분석")

    product_summary, product_monthly = build_product_analysis(q)

    if product_summary.empty:
        st.info("분석할 품목 데이터가 없습니다.")
    else:
        st.caption("※ 매출이 가장 많이 빠진 품목 순으로 정렬했습니다.")
        clean_and_safe_display(
            product_summary,
            height=520,
            pinned_cols=["품목코드"],
            text_cols=["품목코드", "품목명(공식)"]
        )

        st.markdown("#### 품목별 월간 매출 피벗")
        pivot_df = (
            product_monthly.pivot_table(
                index=["품목코드", "품목명(공식)"],
                columns="월표시",
                values="매출",
                aggfunc="sum",
                fill_value=0
            )
            .reset_index()
        )

        clean_and_safe_display(
            pivot_df,
            height=520,
            pinned_cols=["품목코드"],
            text_cols=["품목코드", "품목명(공식)"]
        )


# =========================================================
# TAB3 매출 추이 분석
# =========================================================
with tab3:
    st.subheader("매출 추이 분석")
    st.write("전체 매출 감소 여부와 감소 기여 품목 분석")

    monthly_total = (
        q.groupby(["월", "월표시"], as_index=False)["매출"]
        .sum()
        .sort_values("월")
    )
    draw_customer_sales_trend(monthly_total)

    st.markdown("#### 감소 기여 품목")
    product_summary, product_monthly = build_product_analysis(q)
    top_items = draw_decline_contribution_bar(product_summary, top_n=10)

    if top_items is not None and not top_items.empty:
        top_codes = top_items["품목코드"].astype(str).tolist()
        top_decline_items_monthly = product_monthly[
            product_monthly["품목코드"].astype(str).isin(top_codes)
        ].copy()

        draw_decline_driver_monthly_chart(top_decline_items_monthly)
        draw_half_compare_bar(product_summary, top_n=10)


# =========================================================
# TAB4 우선 대응업체
# =========================================================
with tab4:
    st.subheader("우선 대응업체 순위")

    priority_df, customer_monthly = build_customer_priority(q)

    if priority_df.empty:
        st.info("우선순위 분석 대상 데이터가 없습니다.")
    else:
        display_cols = [
            "순위", "거래처", "우선순위점수",
            "매출감소점수", "통계추세점수", "AI점수",
            "전반부평균", "후반부평균", "감소액", "감소율",
            "CV", "추세기울기", "분석_내역"
        ]
        display_cols = [c for c in display_cols if c in priority_df.columns]

        edited_priority = clean_and_safe_display(
            priority_df[display_cols],
            key="priority_customers_editor",
            editable=True,
            height=520,
            pinned_cols=["순위", "거래처"],
            text_cols=["거래처", "분석_내역"],
            disabled_cols=display_cols
        )

        csv_bytes = priority_df[display_cols].to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button(
            "📥 우선 대응업체 CSV 다운로드",
            data=csv_bytes,
            file_name="priority_customers.csv",
            mime="text/csv"
        )

        st.markdown("#### 상위 거래처 월간 매출 추이")
        top_customers = priority_df.head(5)["거래처"].tolist()
        top_customer_monthly = customer_monthly[customer_monthly["거래처"].isin(top_customers)].copy()

        if not top_customer_monthly.empty:
            fig = go.Figure()
            pos_map = make_text_position_map(top_customers)

            for cust in top_customers:
                sub = top_customer_monthly[top_customer_monthly["거래처"] == cust].sort_values("월")
                sub["라벨"] = sub["매출"].apply(format_manwon)

                fig.add_trace(
                    go.Scatter(
                        x=sub["월표시"],
                        y=sub["매출"],
                        mode="lines+markers+text",
                        name=cust,
                        text=sub["라벨"],
                        textposition=pos_map.get(cust, "top center"),
                        textfont=dict(size=9),
                        line=dict(width=3),
                        marker=dict(size=8),
                        hovertemplate=f"거래처: {cust}<br>월: %{{x}}<br>매출: %{{y:,.0f}}원<extra></extra>"
                    )
                )

            fig.update_layout(
                height=480,
                margin=dict(l=20, r=20, t=20, b=20),
                xaxis=dict(title="월(YYMM)", type="category"),
                yaxis=dict(title="매출액(원)", tickformat=",.0f", separatethousands=True),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
            )
            st.plotly_chart(fig, use_container_width=True)


# =========================================================
# TAB5 원자료
# =========================================================
with tab5:
    st.subheader("원자료 (필터 적용)")

    raw_display_cols = list(q.columns)

    clean_and_safe_display(
        q[raw_display_cols],
        height=520,
        pinned_cols=["거래처", "품목코드"],
        text_cols=["거래처", "품목코드", "품목명(공식)", "점착제코드", "점착제명", "가로폭이력", "최근날짜", "월", "월표시"]
    )

    raw_csv = q.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        "📥 원자료 CSV 다운로드",
        data=raw_csv,
        file_name="raw_filtered_data.csv",
        mime="text/csv"
    )
