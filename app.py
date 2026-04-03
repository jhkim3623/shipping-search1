import os
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# -------------------------------------------------
# Page Config
# -------------------------------------------------
st.set_page_config(
    page_title="출고 이력 검색",
    layout="wide"
)


# -------------------------------------------------
# Style
# -------------------------------------------------
st.markdown("""
<style>
.block-container {
    padding-top: 1.2rem;
    padding-bottom: 2rem;
    max-width: 100%;
}

.app-main-title {
    font-size: 28px;
    font-weight: 800;
    margin-bottom: 12px;
}

.small-help {
    color: #666;
    font-size: 13px;
    margin-bottom: 6px;
}

.section-gap {
    margin-top: 14px;
    margin-bottom: 14px;
}

div[data-testid="stDataFrame"] {
    border: 1px solid #e5e7eb;
    border-radius: 10px;
    overflow: hidden;
}
</style>
""", unsafe_allow_html=True)


# -------------------------------------------------
# Utility
# -------------------------------------------------
def calc_table_height(df, min_height=220, max_height=700):
    if df is None:
        return min_height
    try:
        n = len(df)
    except Exception:
        return min_height

    if n <= 5:
        return 220
    elif n <= 10:
        return 300
    elif n <= 20:
        return 420
    elif n <= 40:
        return 560
    return max_height


def safe_make_product_label(row):
    candidates = ["품목", "품명", "제품명", "품목명", "상품명"]
    for c in candidates:
        if c in row and pd.notna(row[c]):
            return str(row[c])
    return ""


def clean_and_safe_display(
    df,
    use_container_width=True,
    hide_index=True,
    height=None,
    pinned_cols=None,
    text_cols=None,
    key=None,
):
    import pandas as pd
    import numpy as np
    import streamlit as st

    if df is None:
        st.info("표시할 데이터가 없습니다.")
        return

    if not isinstance(df, pd.DataFrame):
        try:
            df = pd.DataFrame(df)
        except Exception:
            st.warning("데이터를 표 형식으로 표시할 수 없습니다.")
            return

    if df.empty:
        st.info("표시할 데이터가 없습니다.")
        return

    display_df = df.copy()

    pinned_cols = pinned_cols or []
    text_cols = text_cols or []

    # ---------------------------------
    # 기본 표시용 값 정리
    # ---------------------------------
    for col in display_df.columns:
        if pd.api.types.is_datetime64_any_dtype(display_df[col]):
            try:
                display_df[col] = pd.to_datetime(display_df[col]).dt.strftime("%Y-%m-%d")
            except Exception:
                display_df[col] = display_df[col].astype(str)
        elif pd.api.types.is_object_dtype(display_df[col]) or pd.api.types.is_string_dtype(display_df[col]):
            try:
                display_df[col] = display_df[col].replace({np.nan: ""})
            except Exception:
                display_df[col] = display_df[col].astype(str)

    # ---------------------------------
    # 컬럼 폭 계산 함수
    # ---------------------------------
    def _safe_strlen(x):
        try:
            if pd.isna(x):
                return 0
        except Exception:
            pass
        return len(str(x))

    def _sample_series(series, sample_size=300):
        try:
            if len(series) <= sample_size:
                return series
            idx = np.linspace(0, len(series) - 1, sample_size, dtype=int)
            return series.iloc[idx]
        except Exception:
            return series

    def _estimate_text_width(series, header):
        header_len = len(str(header))
        s = series.fillna("").astype(str)
        s = _sample_series(s)

        if len(s) == 0:
            max_data_len = 0
        else:
            try:
                max_data_len = s.map(_safe_strlen).max()
            except Exception:
                max_data_len = header_len

        max_len = max(header_len, max_data_len)

        if max_len <= 4:
            return 80
        elif max_len <= 6:
            return 95
        elif max_len <= 8:
            return 110
        elif max_len <= 10:
            return 130
        elif max_len <= 12:
            return 150
        elif max_len <= 16:
            return 180
        elif max_len <= 20:
            return 220
        elif max_len <= 26:
            return 260
        elif max_len <= 34:
            return 320
        else:
            return 400

    def _estimate_numeric_width(series, header):
        header_len = len(str(header))
        s = pd.to_numeric(series, errors="coerce")
        s = _sample_series(s)

        def _fmt_num(x):
            if pd.isna(x):
                return ""
            try:
                if float(x).is_integer():
                    return f"{x:,.0f}"
                return f"{x:,.2f}"
            except Exception:
                return str(x)

        if len(s) == 0:
            max_data_len = 0
        else:
            try:
                formatted = s.map(_fmt_num)
                max_data_len = formatted.map(_safe_strlen).max()
            except Exception:
                max_data_len = header_len

        max_len = max(header_len, max_data_len)

        if max_len <= 4:
            return 80
        elif max_len <= 6:
            return 95
        elif max_len <= 8:
            return 110
        elif max_len <= 10:
            return 125
        elif max_len <= 12:
            return 145
        elif max_len <= 15:
            return 170
        elif max_len <= 18:
            return 200
        else:
            return 230

    def _estimate_date_width(series, header):
        header_len = len(str(header))
        max_len = max(header_len, 10)
        if max_len <= 7:
            return 95
        elif max_len <= 10:
            return 115
        else:
            return 130

    def _is_probably_text_col(col_name, series):
        if col_name in text_cols:
            return True
        if pd.api.types.is_object_dtype(series):
            return True
        if pd.api.types.is_string_dtype(series):
            return True
        return False

    def _infer_number_format(col_name, series):
        col_str = str(col_name)

        if "%" in col_str or "비율" in col_str or "율" in col_str:
            return "%,.2f"

        s = pd.to_numeric(series, errors="coerce").dropna()
        if len(s) == 0:
            return None

        try:
            has_decimal = ((s % 1) != 0).any()
        except Exception:
            has_decimal = False

        money_keywords = ["매출", "금액", "단가", "원가", "이익", "합계", "총액", "공급가"]
        qty_keywords = ["수량", "판매량", "출고량", "재고", "건수", "횟수"]

        if any(k in col_str for k in money_keywords):
            return "%,.0f" if not has_decimal else "%,.2f"

        if any(k in col_str for k in qty_keywords):
            return "%,.0f" if not has_decimal else "%,.2f"

        return "%,.0f" if not has_decimal else "%,.2f"

    # ---------------------------------
    # column_config 생성
    # ---------------------------------
    column_config = {}

    for col in display_df.columns:
        pinned = col in pinned_cols
        series = display_df[col]

        try:
            if pd.api.types.is_numeric_dtype(series) and col not in text_cols:
                width = _estimate_numeric_width(series, col)
                number_format = _infer_number_format(col, series)

                if number_format:
                    column_config[col] = st.column_config.NumberColumn(
                        label=str(col),
                        width=width,
                        format=number_format,
                        pinned=pinned,
                    )
                else:
                    column_config[col] = st.column_config.NumberColumn(
                        label=str(col),
                        width=width,
                        pinned=pinned,
                    )

            elif pd.api.types.is_datetime64_any_dtype(series):
                width = _estimate_date_width(series, col)
                column_config[col] = st.column_config.TextColumn(
                    label=str(col),
                    width=width,
                    pinned=pinned,
                )

            elif _is_probably_text_col(col, series):
                width = _estimate_text_width(series, col)
                column_config[col] = st.column_config.TextColumn(
                    label=str(col),
                    width=width,
                    pinned=pinned,
                )

            else:
                width = _estimate_text_width(series.astype(str), col)
                column_config[col] = st.column_config.TextColumn(
                    label=str(col),
                    width=width,
                    pinned=pinned,
                )

        except Exception:
            column_config[col] = st.column_config.TextColumn(
                label=str(col),
                width="medium",
                pinned=pinned,
            )

    # ---------------------------------
    # 높이 자동 계산
    # ---------------------------------
    if height is None:
        row_count = len(display_df)
        if row_count <= 5:
            height = 220
        elif row_count <= 10:
            height = 300
        elif row_count <= 20:
            height = 420
        elif row_count <= 40:
            height = 560
        else:
            height = 700

    # ---------------------------------
    # dataframe 출력
    # ---------------------------------
    try:
        st.dataframe(
            display_df,
            use_container_width=use_container_width,
            hide_index=hide_index,
            height=height,
            column_config=column_config,
            key=key,
        )
    except TypeError:
        st.dataframe(
            display_df,
            use_container_width=use_container_width,
            hide_index=hide_index,
            height=height,
            column_config=column_config,
        )


def render_banded_table(df, title=None, pinned_cols=None, text_cols=None, key=None):
    if title:
        st.markdown(f"#### {title}")
    clean_and_safe_display(
        df,
        pinned_cols=pinned_cols or [],
        text_cols=text_cols or [],
        height=calc_table_height(df),
        key=key,
    )


def add_year_month_axis(fig, x_dates):
    if x_dates is None or len(x_dates) == 0:
        return fig

    try:
        x_dates = pd.to_datetime(pd.Series(x_dates)).dropna().sort_values().unique()
        if len(x_dates) == 0:
            return fig

        tickvals = list(x_dates)
        ticktext = [pd.to_datetime(d).strftime("%m") for d in x_dates]

        fig.update_xaxes(
            tickmode="array",
            tickvals=tickvals,
            ticktext=ticktext
        )

        years = sorted(set(pd.to_datetime(d).year for d in x_dates))
        annotations = []

        for y in years:
            year_dates = [pd.to_datetime(d) for d in x_dates if pd.to_datetime(d).year == y]
            if not year_dates:
                continue
            mid_idx = len(year_dates) // 2
            mid_date = year_dates[mid_idx]

            annotations.append(
                dict(
                    x=mid_date,
                    y=-0.18,
                    xref="x",
                    yref="paper",
                    text=str(y),
                    showarrow=False,
                    font=dict(size=12, color="#666")
                )
            )

        fig.update_layout(annotations=annotations)
    except Exception:
        pass

    return fig


def apply_mobile_friendly_line_layout(fig, x_dates, y_title="판매금액(원)", height=380):
    fig = add_year_month_axis(fig, x_dates)
    fig.update_layout(
        height=height,
        margin=dict(l=20, r=20, t=50, b=70),
        xaxis_title="",
        yaxis_title=y_title,
        hovermode="x unified",
        legend=dict(orientation="h")
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=True, gridcolor="rgba(0,0,0,0.08)")
    return fig


# -------------------------------------------------
# Data Loading
# -------------------------------------------------
DEFAULT_FILE = "data.xlsx"


@st.cache_data(show_spinner=False)
def load_excel(file_bytes):
    try:
        xl = pd.ExcelFile(BytesIO(file_bytes))
        sheet_map = {}
        for sheet in xl.sheet_names:
            try:
                sheet_map[sheet] = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet)
            except Exception:
                pass
        return sheet_map
    except Exception:
        return {}


def get_primary_df(sheet_map):
    if not sheet_map:
        return pd.DataFrame()

    priority_names = [
        "Sheet1", "sheet1", "출고", "출고이력", "data", "DATA", "원자료"
    ]
    for name in priority_names:
        if name in sheet_map:
            return sheet_map[name]

    return next(iter(sheet_map.values()))


def normalize_date_column(df):
    if df.empty:
        return df.copy()

    out = df.copy()
    date_candidates = ["출고일", "일자", "날짜", "매출일", "년월", "월"]
    for c in date_candidates:
        if c in out.columns:
            try:
                out[c] = pd.to_datetime(out[c], errors="coerce")
                return out
            except Exception:
                pass
    return out


def ensure_month_column(df):
    out = df.copy()
    if out.empty:
        return out

    if "년월" not in out.columns:
        date_candidates = ["출고일", "일자", "날짜", "매출일", "월"]
        for c in date_candidates:
            if c in out.columns:
                try:
                    dt = pd.to_datetime(out[c], errors="coerce")
                    if dt.notna().any():
                        out["년월"] = dt.dt.to_period("M").astype(str)
                        break
                except Exception:
                    continue
    return out


# -------------------------------------------------
# Sample Analysis Helpers
# -------------------------------------------------
@st.cache_data(show_spinner=False)
def build_customer_sales_analysis(df):
    if df is None or df.empty:
        return {
            "customer_summary": pd.DataFrame(),
            "customer_monthly": pd.DataFrame(),
            "customer_item_summary": pd.DataFrame(),
            "customer_item_monthly": pd.DataFrame(),
            "all_months": [],
        }

    q = df.copy()
    q = ensure_month_column(q)

    customer_col = None
    sales_col = None
    qty_col = None
    item_col = None

    for c in ["거래처", "업체", "고객", "거래처명"]:
        if c in q.columns:
            customer_col = c
            break

    for c in ["매출액", "판매금액", "금액", "공급가액", "합계금액"]:
        if c in q.columns:
            sales_col = c
            break

    for c in ["판매량", "수량", "출고량"]:
        if c in q.columns:
            qty_col = c
            break

    for c in ["품목", "품목명", "품명", "제품명", "상품명"]:
        if c in q.columns:
            item_col = c
            break

    if customer_col is None or sales_col is None or "년월" not in q.columns:
        return {
            "customer_summary": pd.DataFrame(),
            "customer_monthly": pd.DataFrame(),
            "customer_item_summary": pd.DataFrame(),
            "customer_item_monthly": pd.DataFrame(),
            "all_months": [],
        }

    q[sales_col] = pd.to_numeric(q[sales_col], errors="coerce").fillna(0)

    if qty_col and qty_col in q.columns:
        q[qty_col] = pd.to_numeric(q[qty_col], errors="coerce").fillna(0)

    customer_monthly = (
        q.groupby([customer_col, "년월"], dropna=False)[sales_col]
        .sum()
        .reset_index()
        .rename(columns={customer_col: "거래처", sales_col: "매출액"})
    )

    all_months = sorted(customer_monthly["년월"].dropna().unique().tolist())

    summary = (
        customer_monthly.groupby("거래처", dropna=False)["매출액"]
        .sum()
        .reset_index()
        .rename(columns={"매출액": "전체매출액"})
        .sort_values("전체매출액", ascending=False)
    )

    def _half_stats(name):
        sub = customer_monthly[customer_monthly["거래처"] == name].sort_values("년월").copy()
        vals = sub["매출액"].tolist()
        if not vals:
            return pd.Series([0, 0, 0, "분석불가"])
        mid = len(vals) // 2
        front = vals[:mid] if mid > 0 else vals
        back = vals[mid:] if mid < len(vals) else vals
        front_avg = float(np.mean(front)) if len(front) else 0
        back_avg = float(np.mean(back)) if len(back) else 0
        diff = back_avg - front_avg
        if diff > 0:
            status = "상승추세"
        elif diff < 0:
            status = "감소추세"
        else:
            status = "보합"
        return pd.Series([front_avg, back_avg, diff, status])

    if not summary.empty:
        summary[["전반부_평균매출", "후반부_평균매출", "평균증감액", "진행현황"]] = summary["거래처"].apply(_half_stats)

    if item_col and item_col in q.columns:
        customer_item_summary = (
            q.groupby([customer_col, item_col], dropna=False)
            .agg(
                매출액=(sales_col, "sum"),
                총판매량=(qty_col, "sum") if qty_col else (sales_col, "size")
            )
            .reset_index()
            .rename(columns={customer_col: "거래처", item_col: "품목"})
            .sort_values(["거래처", "매출액"], ascending=[True, False])
        )

        customer_item_monthly = (
            q.groupby([customer_col, item_col, "년월"], dropna=False)[sales_col]
            .sum()
            .reset_index()
            .rename(columns={customer_col: "거래처", item_col: "품목", sales_col: "매출액"})
        )
    else:
        customer_item_summary = pd.DataFrame()
        customer_item_monthly = pd.DataFrame()

    return {
        "customer_summary": summary,
        "customer_monthly": customer_monthly,
        "customer_item_summary": customer_item_summary,
        "customer_item_monthly": customer_item_monthly,
        "all_months": all_months,
    }


# -------------------------------------------------
# App Header
# -------------------------------------------------
st.markdown('<div class="app-main-title">출고 이력 검색(거래처/품목/가로폭/점착제)</div>', unsafe_allow_html=True)
st.markdown('<div class="small-help">헤더와 데이터 값 길이에 맞춰 표 컬럼 폭을 자동 조정합니다.</div>', unsafe_allow_html=True)


# -------------------------------------------------
# File Load
# -------------------------------------------------
uploaded = st.file_uploader("📂 다른 파일 업로드 (미업로드 시 기본 데이터 자동 로드)", type=["xlsx"])

file_bytes = None

if uploaded is not None:
    file_bytes = uploaded.read()
else:
    if os.path.exists(DEFAULT_FILE):
        with open(DEFAULT_FILE, "rb") as f:
            file_bytes = f.read()

if file_bytes is None:
    st.warning("불러올 엑셀 파일이 없습니다. `data.xlsx`를 두거나 파일을 업로드해주세요.")
    st.stop()

sheet_map = load_excel(file_bytes)
df = get_primary_df(sheet_map)
df = normalize_date_column(df)
df = ensure_month_column(df)


# -------------------------------------------------
# Sidebar Filters
# -------------------------------------------------
with st.sidebar:
    st.header("검색 조건")

    search_text = st.text_input("검색어", value="")

    candidate_customer_cols = [c for c in ["거래처", "업체", "고객", "거래처명"] if c in df.columns]
    candidate_item_cols = [c for c in ["품목", "품목명", "품명", "제품명", "상품명"] if c in df.columns]

    customer_col = candidate_customer_cols[0] if candidate_customer_cols else None
    item_col = candidate_item_cols[0] if candidate_item_cols else None

    if customer_col:
        customer_options = ["전체"] + sorted(df[customer_col].dropna().astype(str).unique().tolist())
        selected_customer = st.selectbox("거래처 선택", customer_options, index=0)
    else:
        selected_customer = "전체"

    if item_col:
        item_options = ["전체"] + sorted(df[item_col].dropna().astype(str).unique().tolist())
        selected_item = st.selectbox("품목 선택", item_options, index=0)
    else:
        selected_item = "전체"


# -------------------------------------------------
# Filtered Data
# -------------------------------------------------
q = df.copy()

if search_text:
    mask = pd.Series(False, index=q.index)
    for c in q.columns:
        try:
            mask = mask | q[c].astype(str).str.contains(search_text, case=False, na=False)
        except Exception:
            pass
    q = q[mask]

if customer_col and selected_customer != "전체":
    q = q[q[customer_col].astype(str) == selected_customer]

if item_col and selected_item != "전체":
    q = q[q[item_col].astype(str) == selected_item]


# -------------------------------------------------
# Tabs
# -------------------------------------------------
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "거래처별 검색",
    "품목별 검색",
    "🏷️ 견적 레퍼런스",
    "📉 매출 하락 분석",
    "📊 거래처별 매출 분석",
    "품목별 하락 원인 분석",
    "원자료",
])


# -------------------------------------------------
# Tab 1
# -------------------------------------------------
with tab1:
    st.subheader("거래처별 검색")

    if q.empty:
        st.info("조회 결과가 없습니다.")
    else:
        if customer_col:
            cust_summary = q.groupby(customer_col, dropna=False).size().reset_index(name="건수")
            render_banded_table(
                cust_summary.sort_values("건수", ascending=False),
                title="거래처 요약",
                pinned_cols=[customer_col],
                text_cols=[customer_col],
                key="tab1_customer_summary"
            )
        else:
            clean_and_safe_display(q, key="tab1_raw")


# -------------------------------------------------
# Tab 2
# -------------------------------------------------
with tab2:
    st.subheader("품목별 검색")

    if q.empty:
        st.info("조회 결과가 없습니다.")
    else:
        if item_col:
            item_summary = q.groupby(item_col, dropna=False).size().reset_index(name="건수")
            render_banded_table(
                item_summary.sort_values("건수", ascending=False),
                title="품목 요약",
                pinned_cols=[item_col],
                text_cols=[item_col],
                key="tab2_item_summary"
            )
        else:
            clean_and_safe_display(q, key="tab2_raw")


# -------------------------------------------------
# Tab 3
# -------------------------------------------------
with tab3:
    st.subheader("🏷️ 견적 레퍼런스")
    if q.empty:
        st.info("조회 결과가 없습니다.")
    else:
        ref_df = q.head(100).copy()
        clean_and_safe_display(
            ref_df,
            pinned_cols=[c for c in ["거래처", "품목"] if c in ref_df.columns],
            text_cols=[c for c in ["거래처", "품목", "품목명", "품명"] if c in ref_df.columns],
            key="tab3_ref"
        )


# -------------------------------------------------
# Tab 4
# -------------------------------------------------
with tab4:
    st.subheader("📉 매출 하락 분석")

    if q.empty:
        st.info("조회 결과가 없습니다.")
    else:
        pack = build_customer_sales_analysis(q)
        customer_summary = pack["customer_summary"]

        if customer_summary.empty:
            st.info("분석 가능한 매출 데이터가 없습니다.")
        else:
            show_cols = [c for c in ["거래처", "전체매출액", "전반부_평균매출", "후반부_평균매출", "평균증감액", "진행현황"] if c in customer_summary.columns]
            clean_and_safe_display(
                customer_summary[show_cols],
                pinned_cols=["거래처"] if "거래처" in customer_summary.columns else [],
                text_cols=["거래처", "진행현황"],
                key="tab4_decline_table"
            )


# -------------------------------------------------
# Tab 5
# -------------------------------------------------
with tab5:
    st.subheader("📊 거래처별 매출 분석")

    if q.empty:
        st.info("조회 결과가 없습니다.")
    else:
        pack = build_customer_sales_analysis(q)
        customer_summary = pack["customer_summary"]
        customer_monthly = pack["customer_monthly"]
        customer_item_summary = pack["customer_item_summary"]

        if customer_summary.empty:
            st.info("분석 가능한 데이터가 없습니다.")
        else:
            summary_cols = [c for c in ["거래처", "전체매출액", "전반부_평균매출", "후반부_평균매출", "평균증감액", "진행현황"] if c in customer_summary.columns]
            clean_and_safe_display(
                customer_summary[summary_cols],
                pinned_cols=["거래처"],
                text_cols=["거래처", "진행현황"],
                key="tab5_customer_summary"
            )

            customer_options = customer_summary["거래처"].dropna().astype(str).tolist()
            selected_customer_tab5 = st.selectbox("업체 선택", customer_options, key="tab5_customer_select")

            if selected_customer_tab5:
                cust_month = customer_monthly[customer_monthly["거래처"].astype(str) == str(selected_customer_tab5)].copy()
                if not cust_month.empty:
                    cust_month["날짜축"] = pd.to_datetime(cust_month["년월"], errors="coerce")
                    cust_month = cust_month.sort_values("날짜축")

                    fig = go.Figure()
                    fig.add_trace(go.Scatter(
                        x=cust_month["날짜축"],
                        y=cust_month["매출액"],
                        mode="lines+markers+text",
                        text=[f"{v:,.0f}" if pd.notna(v) else "" for v in cust_month["매출액"]],
                        textposition="top center",
                        name="매출액",
                        line=dict(color="#1f77b4", width=3),
                        marker=dict(size=7),
                    ))
                    fig = apply_mobile_friendly_line_layout(fig, cust_month["날짜축"], y_title="매출액(원)")
                    fig.update_layout(title="업체 전체 월별 매출 추이")
                    st.plotly_chart(fig, use_container_width=True, key="tab5_customer_chart")

                if not customer_item_summary.empty:
                    sub_item = customer_item_summary[customer_item_summary["거래처"].astype(str) == str(selected_customer_tab5)].copy()
                    if not sub_item.empty:
                        clean_and_safe_display(
                            sub_item,
                            pinned_cols=["거래처", "품목"] if "품목" in sub_item.columns else ["거래처"],
                            text_cols=[c for c in ["거래처", "품목"] if c in sub_item.columns],
                            key="tab5_customer_item_table"
                        )


# -------------------------------------------------
# Tab 6
# -------------------------------------------------
with tab6:
    st.subheader("품목별 하락 원인 분석")
    if q.empty:
        st.info("조회 결과가 없습니다.")
    else:
        sample = q.head(200).copy()
        clean_and_safe_display(
            sample,
            pinned_cols=[c for c in ["거래처", "품목"] if c in sample.columns],
            text_cols=[c for c in ["거래처", "품목", "품목명", "품명"] if c in sample.columns],
            key="tab6_sample"
        )


# -------------------------------------------------
# Tab 7
# -------------------------------------------------
with tab7:
    st.subheader("원자료")
    if df.empty:
        st.info("원자료가 없습니다.")
    else:
        clean_and_safe_display(
            df,
            pinned_cols=[c for c in ["거래처", "품목"] if c in df.columns],
            text_cols=[c for c in ["거래처", "품목", "품목명", "품명", "업체", "고객"] if c in df.columns],
            key="tab7_raw"
        )
