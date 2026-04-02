import os
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# =========================================================
# Page Config
# =========================================================
st.set_page_config(page_title="출고 이력 검색", layout="wide")

# =========================================================
# Style
# =========================================================
st.markdown(
    """
<style>
.block-container {
    padding-top: 0.8rem;
    padding-bottom: 1rem;
    max-width: 100%;
}
.app-main-title {
    font-size: 28px;
    font-weight: 800;
    margin-bottom: 8px;
}
.small-caption {
    color: #666;
    font-size: 13px;
    margin-top: -4px;
    margin-bottom: 10px;
}
.section-title {
    font-size: 20px;
    font-weight: 800;
    margin: 14px 0 8px 0;
}
.card-title {
    font-size: 16px;
    font-weight: 700;
    margin-bottom: 8px;
}
.stMetric {
    background: #fafafa;
    border: 1px solid #eee;
    border-radius: 12px;
    padding: 8px;
}
div[data-baseweb="select"] > div {
    min-height: 42px !important;
}
section[data-testid="stSidebar"] .stMultiSelect label,
section[data-testid="stSidebar"] .stDateInput label,
section[data-testid="stSidebar"] .stSelectbox label {
    font-weight: 700 !important;
}
</style>
""",
    unsafe_allow_html=True,
)

# =========================================================
# Utility Functions
# =========================================================
def calc_table_height(df, min_rows=3, max_rows=18, row_px=35, header_px=38):
    if df is None or len(df) == 0:
        rows = min_rows
    else:
        rows = min(max_rows, max(min_rows, len(df)))
    return header_px + rows * row_px


def safe_numeric(series, default=0):
    try:
        s = pd.to_numeric(series, errors="coerce")
        return s.replace([np.inf, -np.inf], np.nan).fillna(default)
    except Exception:
        return pd.Series([default] * len(series)) if series is not None else pd.Series(dtype=float)


def safe_str(series):
    try:
        return pd.Series(series).fillna("").astype(str).str.strip()
    except Exception:
        return pd.Series(dtype=str)


def safe_date(series):
    try:
        return pd.to_datetime(series, errors="coerce")
    except Exception:
        return pd.Series(dtype="datetime64[ns]")


def sorted_unique(series):
    """
    mixed type 정렬 오류 방지용
    """
    if series is None:
        return []
    try:
        s = pd.Series(series).dropna().astype(str).str.strip()
        s = s[~s.isin(["", "0", "0.0", "nan", "NaN", "None", "<NA>", "NaT"])]
        vals = list(dict.fromkeys(s.tolist()))
        return sorted(vals, key=lambda x: str(x))
    except Exception:
        return []


def sales_to_manwon_label(value):
    try:
        if pd.isna(value):
            return ""
        return f"{int(round(float(value) / 10000.0)):,}"
    except Exception:
        return ""


def add_year_month_axis(df, date_col="날짜", out_col="월"):
    temp = df.copy()
    if date_col not in temp.columns:
        temp[out_col] = ""
        return temp
    temp[date_col] = pd.to_datetime(temp[date_col], errors="coerce")
    temp[out_col] = temp[date_col].dt.strftime("%Y-%m").fillna("")
    return temp


def build_month_axis_frame(months):
    month_list = sorted([str(m) for m in months if pd.notna(m) and str(m).strip() != ""])
    df = pd.DataFrame({"월": month_list})
    if df.empty:
        df["날짜축"] = pd.NaT
        return df
    df["날짜축"] = pd.to_datetime(df["월"] + "-01", errors="coerce")
    return df.dropna(subset=["날짜축"]).sort_values("날짜축").reset_index(drop=True)


def align_monthly_series(base_month_df, data_df, value_col):
    if base_month_df is None or base_month_df.empty:
        return pd.DataFrame(columns=["월", "날짜축", value_col])

    out = base_month_df.copy()

    if data_df is None or data_df.empty or "월" not in data_df.columns or value_col not in data_df.columns:
        out[value_col] = 0
        return out

    temp = data_df.copy()
    temp["월"] = temp["월"].astype(str)
    temp[value_col] = pd.to_numeric(temp[value_col], errors="coerce").fillna(0)
    temp = temp.groupby("월", as_index=False)[value_col].sum()

    out = out.merge(temp, on="월", how="left")
    out[value_col] = pd.to_numeric(out[value_col], errors="coerce").fillna(0)
    return out


def safe_make_product_label(df):
    temp = df.copy()
    if "품목코드" not in temp.columns:
        temp["품목코드"] = ""
    if "품목명(공식)" not in temp.columns:
        temp["품목명(공식)"] = ""
    temp["품목코드"] = temp["품목코드"].fillna("").astype(str).str.strip()
    temp["품목명(공식)"] = temp["품목명(공식)"].fillna("").astype(str).str.strip()

    temp["품목표시"] = np.where(
        temp["품목명(공식)"] != "",
        temp["품목코드"] + " | " + temp["품목명(공식)"],
        temp["품목코드"]
    )
    return temp


def clean_and_safe_display(df, height=None, key=None):
    if df is None or df.empty:
        st.info("표시할 데이터가 없습니다.")
        return
    temp = df.copy()
    for c in temp.columns:
        if pd.api.types.is_datetime64_any_dtype(temp[c]):
            temp[c] = temp[c].dt.strftime("%Y-%m-%d")
    st.dataframe(temp, use_container_width=True, height=height or calc_table_height(temp), key=key)


def safe_download_button(df, file_name, label):
    try:
        csv = df.to_csv(index=False, encoding="utf-8-sig")
        st.download_button(label=label, data=csv, file_name=file_name, mime="text/csv")
    except Exception:
        st.warning("다운로드 파일 생성 중 오류가 발생했습니다.")


def calc_slope(values):
    try:
        y = pd.Series(values).astype(float).replace([np.inf, -np.inf], np.nan).dropna()
        if len(y) < 2:
            return 0.0
        x = np.arange(len(y))
        slope = np.polyfit(x, y, 1)[0]
        return float(slope)
    except Exception:
        return 0.0


def calc_cv(values):
    try:
        y = pd.Series(values).astype(float).replace([np.inf, -np.inf], np.nan).dropna()
        if len(y) == 0:
            return 0.0
        mean_v = y.mean()
        if mean_v == 0:
            return 0.0
        return float(y.std(ddof=0) / mean_v)
    except Exception:
        return 0.0


def make_indexed_series(df, group_col, value_col, time_col):
    if df is None or df.empty:
        return pd.DataFrame(columns=[group_col, time_col, value_col, "지수값"])

    required = {group_col, value_col, time_col}
    if not required.issubset(set(df.columns)):
        return pd.DataFrame(columns=[group_col, time_col, value_col, "지수값"])

    tmp = df.copy()
    tmp[group_col] = tmp[group_col].astype(str)
    tmp[value_col] = pd.to_numeric(tmp[value_col], errors="coerce").fillna(0)
    tmp[time_col] = pd.to_datetime(tmp[time_col], errors="coerce")
    tmp = tmp.dropna(subset=[time_col]).sort_values([group_col, time_col]).reset_index(drop=True)

    out = []
    for g, sub in tmp.groupby(group_col, dropna=False):
        sub = sub.copy()
        non_zero = sub[value_col].replace(0, np.nan).dropna()
        base = non_zero.iloc[0] if not non_zero.empty else np.nan
        if pd.isna(base) or base == 0:
            sub["지수값"] = np.nan
        else:
            sub["지수값"] = (sub[value_col] / base) * 100.0
        out.append(sub)

    if not out:
        return pd.DataFrame(columns=[group_col, time_col, value_col, "지수값"])
    return pd.concat(out, ignore_index=True)


# =========================================================
# Data Loading
# =========================================================
@st.cache_data(show_spinner=False)
def load_excel(file_bytes):
    xls = pd.ExcelFile(BytesIO(file_bytes))
    sheet_names = xls.sheet_names

    rec = pd.read_excel(xls, "출고기록") if "출고기록" in sheet_names else pd.DataFrame()
    alias = pd.read_excel(xls, "별칭맵핑") if "별칭맵핑" in sheet_names else pd.DataFrame()
    prod = pd.read_excel(xls, "품목마스터") if "품목마스터" in sheet_names else pd.DataFrame()
    adh = pd.read_excel(xls, "점착제마스터") if "점착제마스터" in sheet_names else pd.DataFrame()
    cust = pd.read_excel(xls, "거래처마스터") if "거래처마스터" in sheet_names else pd.DataFrame()

    # -------- 출고기록 기본 컬럼 보정 --------
    if rec.empty:
        return rec, alias, prod, adh, cust

    rename_candidates = {
        "출고일": "날짜",
        "일자": "날짜",
        "거래처명": "거래처",
        "품명": "품목명(공식)",
        "금액": "매출액",
        "출고수량": "수량",
        "판매수량": "수량",
        "폭": "가로폭",
    }
    for old, new in rename_candidates.items():
        if old in rec.columns and new not in rec.columns:
            rec = rec.rename(columns={old: new})

    required_defaults = {
        "날짜": pd.NaT,
        "거래처": "",
        "품목코드": "",
        "품목명(공식)": "",
        "점착제코드": "",
        "매출액": 0,
        "수량": 0,
        "단가": 0,
        "가로폭": 0,
        "담당부서": "",
        "담당자": "",
    }
    for col, default_val in required_defaults.items():
        if col not in rec.columns:
            rec[col] = default_val

    # -------- 마스터 병합(가능한 경우만) --------
    if not prod.empty:
        prod2 = prod.copy()
        if "품목코드" in prod2.columns:
            if "품목명(공식)" not in prod2.columns:
                # 후보 컬럼 탐색
                name_candidates = [c for c in prod2.columns if "품목명" in str(c)]
                if name_candidates:
                    prod2["품목명(공식)"] = prod2[name_candidates[0]]
            keep_cols = [c for c in ["품목코드", "품목명(공식)"] if c in prod2.columns]
            if "품목코드" in keep_cols and len(keep_cols) >= 1:
                prod2 = prod2[keep_cols].drop_duplicates()
                rec = rec.merge(prod2, on="품목코드", how="left", suffixes=("", "_prod"))
                if "품목명(공식)_prod" in rec.columns:
                    rec["품목명(공식)"] = np.where(
                        safe_str(rec["품목명(공식)"]) == "",
                        safe_str(rec["품목명(공식)_prod"]),
                        safe_str(rec["품목명(공식)"])
                    )
                    rec = rec.drop(columns=["품목명(공식)_prod"], errors="ignore")

    # -------- 타입 정리 --------
    rec["날짜"] = safe_date(rec["날짜"])
    for c in ["매출액", "수량", "단가", "가로폭"]:
        rec[c] = safe_numeric(rec[c], default=0)

    for c in ["거래처", "품목코드", "품목명(공식)", "점착제코드", "담당부서", "담당자"]:
        rec[c] = safe_str(rec[c])

    rec = safe_make_product_label(rec)
    rec = add_year_month_axis(rec, "날짜", "월")

    # 최근단가 계산
    if not rec.empty:
        temp_recent = rec.dropna(subset=["날짜"]).sort_values("날짜")
        recent_price = (
            temp_recent.groupby(["거래처", "품목표시"], as_index=False)
            .tail(1)[["거래처", "품목표시", "단가", "가로폭", "날짜"]]
            .rename(columns={"단가": "최근단가", "가로폭": "최근가로폭", "날짜": "최근출고일"})
        )
        rec = rec.merge(recent_price, on=["거래처", "품목표시"], how="left")

    return rec, alias, prod, adh, cust


# =========================================================
# Analysis Builders
# =========================================================
@st.cache_data(show_spinner=False)
def build_customer_sales_analysis(df):
    result = {
        "df": pd.DataFrame(),
        "all_months": pd.DataFrame(),
        "customer_summary": pd.DataFrame(),
        "customer_monthly": pd.DataFrame(),
        "customer_item_monthly": pd.DataFrame(),
        "customer_item_summary": pd.DataFrame(),
    }

    if df is None or df.empty:
        return result

    temp = df.copy()

    for col in ["거래처", "품목표시", "월"]:
        if col not in temp.columns:
            temp[col] = ""

    if "매출액" not in temp.columns:
        temp["매출액"] = 0

    temp["거래처"] = safe_str(temp["거래처"])
    temp["품목표시"] = safe_str(temp["품목표시"])
    temp["월"] = safe_str(temp["월"])
    temp["매출액"] = safe_numeric(temp["매출액"], default=0)

    temp = temp[temp["거래처"] != ""].copy()

    all_months = build_month_axis_frame(temp["월"].dropna().unique())

    customer_monthly = (
        temp.groupby(["거래처", "월"], as_index=False)["매출액"]
        .sum()
        .rename(columns={"매출액": "월매출"})
    )

    customer_summary = (
        temp.groupby("거래처", as_index=False)["매출액"]
        .sum()
        .rename(columns={"매출액": "총매출"})
        .sort_values("총매출", ascending=False)
        .reset_index(drop=True)
    )

    # 변화율, 추세, 변동성
    change_rates = []
    for cust_name in customer_summary["거래처"].tolist():
        sub = customer_monthly[customer_monthly["거래처"] == cust_name].copy()
        aligned = align_monthly_series(all_months, sub.rename(columns={"월매출": "값"}), "값")
        vals = aligned["값"].tolist()
        non_zero = [v for v in vals if pd.notna(v)]
        first_val = non_zero[0] if len(non_zero) >= 1 else 0
        last_val = non_zero[-1] if len(non_zero) >= 1 else 0

        if first_val == 0:
            change_pct = np.nan if last_val == 0 else 100.0
        else:
            change_pct = ((last_val - first_val) / first_val) * 100.0

        change_rates.append(
            {
                "거래처": cust_name,
                "변화율(%)": round(change_pct, 2) if pd.notna(change_pct) else np.nan,
                "추세기울기": round(calc_slope(vals), 4),
                "변동계수(CV)": round(calc_cv(vals), 4),
                "최근월매출": float(last_val) if pd.notna(last_val) else 0,
            }
        )

    rate_df = pd.DataFrame(change_rates)
    customer_summary = customer_summary.merge(rate_df, on="거래처", how="left")

    customer_item_monthly = (
        temp.groupby(["거래처", "품목표시", "월"], as_index=False)["매출액"]
        .sum()
        .rename(columns={"매출액": "월매출"})
    )

    # 품목별 요약
    item_summary = (
        temp.groupby(["거래처", "품목표시"], as_index=False)
        .agg(
            총매출=("매출액", "sum"),
            최근단가=("최근단가", "max") if "최근단가" in temp.columns else ("매출액", "size"),
            최근가로폭=("최근가로폭", "max") if "최근가로폭" in temp.columns else ("매출액", "size"),
        )
    )

    item_rates = []
    for (cust_name, item_name), sub in customer_item_monthly.groupby(["거래처", "품목표시"], dropna=False):
        aligned = align_monthly_series(all_months, sub.rename(columns={"월매출": "값"}), "값")
        vals = aligned["값"].tolist()
        non_zero = [v for v in vals if pd.notna(v)]
        first_val = non_zero[0] if len(non_zero) >= 1 else 0
        last_val = non_zero[-1] if len(non_zero) >= 1 else 0

        if first_val == 0:
            change_pct = np.nan if last_val == 0 else 100.0
        else:
            change_pct = ((last_val - first_val) / first_val) * 100.0

        item_rates.append(
            {
                "거래처": cust_name,
                "품목표시": item_name,
                "변화율(%)": round(change_pct, 2) if pd.notna(change_pct) else np.nan,
                "최근월매출": float(last_val) if pd.notna(last_val) else 0,
                "추세기울기": round(calc_slope(vals), 4),
                "변동계수(CV)": round(calc_cv(vals), 4),
            }
        )

    item_rate_df = pd.DataFrame(item_rates)
    item_summary = item_summary.merge(item_rate_df, on=["거래처", "품목표시"], how="left")

    result["df"] = temp
    result["all_months"] = all_months
    result["customer_summary"] = customer_summary.sort_values("총매출", ascending=False).reset_index(drop=True)
    result["customer_monthly"] = customer_monthly
    result["customer_item_monthly"] = customer_item_monthly
    result["customer_item_summary"] = item_summary.sort_values(["거래처", "총매출"], ascending=[True, False]).reset_index(drop=True)

    return result


def add_trendline(fig, x_values, y_values, name="추세선", color="rgba(255,0,0,0.7)"):
    try:
        x = pd.to_datetime(pd.Series(x_values), errors="coerce")
        y = pd.to_numeric(pd.Series(y_values), errors="coerce")
        mask = (~x.isna()) & (~y.isna())
        x = x[mask]
        y = y[mask]

        if len(x) < 2:
            return fig

        idx = np.arange(len(x))
        slope, intercept = np.polyfit(idx, y, 1)
        trend = slope * idx + intercept

        fig.add_trace(
            go.Scatter(
                x=x,
                y=trend,
                mode="lines",
                name=name,
                line=dict(color=color, dash="dash", width=2),
            )
        )
        return fig
    except Exception:
        return fig


# =========================================================
# Main
# =========================================================
DEFAULT_FILE = "data.xlsx"

st.markdown('<div class="app-main-title">출고 이력 검색 / 매출 분석</div>', unsafe_allow_html=True)
st.caption("기본 파일(data.xlsx) 자동 로드 / 업로드 파일 우선 적용")

uploaded = st.file_uploader("📂 다른 파일 업로드 (xlsx)", type=["xlsx"])

file_bytes = None
if uploaded is not None:
    file_bytes = uploaded.getvalue()
    st.success("업로드 파일을 사용합니다.")
elif os.path.exists(DEFAULT_FILE):
    with open(DEFAULT_FILE, "rb") as f:
        file_bytes = f.read()
    st.info(f"기본 파일({DEFAULT_FILE})을 자동 로드했습니다.")
else:
    st.warning("data.xlsx 파일이 없습니다. 파일을 업로드해주세요.")
    st.stop()

rec, alias, prod, adh, cust = load_excel(file_bytes)

if rec is None or rec.empty:
    st.warning("출고기록 데이터가 비어 있습니다.")
    st.stop()

# =========================================================
# Sidebar Filters
# =========================================================
with st.sidebar:
    st.markdown("## 검색 필터")

    dept_options = sorted_unique(rec["담당부서"]) if "담당부서" in rec.columns else []
    emp_options = sorted_unique(rec["담당자"]) if "담당자" in rec.columns else []
    cust_options = sorted_unique(rec["거래처"]) if "거래처" in rec.columns else []
    item_code_options = sorted_unique(rec["품목코드"]) if "품목코드" in rec.columns else []
    adh_options = sorted_unique(rec["점착제코드"]) if "점착제코드" in rec.columns else []

    sel_dept = st.multiselect("담당부서", dept_options, default=[])
    sel_emp = st.multiselect("담당자", emp_options, default=[])
    sel_cust = st.multiselect("거래처", cust_options, default=[])
    sel_item_code = st.multiselect("품목코드", item_code_options, default=[])
    sel_adh = st.multiselect("점착제코드", adh_options, default=[])

    min_date = rec["날짜"].min() if "날짜" in rec.columns and rec["날짜"].notna().any() else pd.Timestamp.today()
    max_date = rec["날짜"].max() if "날짜" in rec.columns and rec["날짜"].notna().any() else pd.Timestamp.today()

    date_range = st.date_input(
        "기간",
        value=(min_date.date(), max_date.date()),
        min_value=min_date.date(),
        max_value=max_date.date(),
    )

# =========================================================
# Apply Filters
# =========================================================
q = rec.copy()

if sel_dept and "담당부서" in q.columns:
    q = q[q["담당부서"].astype(str).isin(sel_dept)]

if sel_emp and "담당자" in q.columns:
    q = q[q["담당자"].astype(str).isin(sel_emp)]

if sel_cust and "거래처" in q.columns:
    q = q[q["거래처"].astype(str).isin(sel_cust)]

if sel_item_code and "품목코드" in q.columns:
    q = q[q["품목코드"].astype(str).isin(sel_item_code)]

if sel_adh and "점착제코드" in q.columns:
    q = q[q["점착제코드"].astype(str).isin(sel_adh)]

if isinstance(date_range, (list, tuple)) and len(date_range) == 2 and "날짜" in q.columns:
    start_date = pd.to_datetime(date_range[0], errors="coerce")
    end_date = pd.to_datetime(date_range[1], errors="coerce")
    q = q[(q["날짜"] >= start_date) & (q["날짜"] <= end_date)]

q = q.copy()
q = add_year_month_axis(q, "날짜", "월")

# =========================================================
# Tabs
# =========================================================
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(
    [
        "거래처별 검색",
        "품목별 검색",
        "견적 레퍼런스",
        "매출 하락 분석",
        "거래처별 매출 분석",
        "매출 감소 품목 분석",
        "원자료",
    ]
)

# =========================================================
# Tab1: 거래처별 검색
# =========================================================
with tab1:
    st.markdown('<div class="section-title">거래처별 검색</div>', unsafe_allow_html=True)

    if q.empty:
        st.info("필터 조건에 맞는 데이터가 없습니다.")
    else:
        summary = (
            q.groupby("거래처", as_index=False)
            .agg(
                총매출=("매출액", "sum"),
                출고건수=("거래처", "size"),
                품목수=("품목표시", "nunique"),
            )
            .sort_values("총매출", ascending=False)
            .reset_index(drop=True)
        )
        clean_and_safe_display(summary)
        safe_download_button(summary, "거래처별검색.csv", "거래처별 검색 결과 다운로드")

# =========================================================
# Tab2: 품목별 검색
# =========================================================
with tab2:
    st.markdown('<div class="section-title">품목별 검색</div>', unsafe_allow_html=True)

    if q.empty:
        st.info("필터 조건에 맞는 데이터가 없습니다.")
    else:
        summary = (
            q.groupby("품목표시", as_index=False)
            .agg(
                총매출=("매출액", "sum"),
                출고건수=("품목표시", "size"),
                거래처수=("거래처", "nunique"),
                평균단가=("단가", "mean"),
            )
            .sort_values("총매출", ascending=False)
            .reset_index(drop=True)
        )
        clean_and_safe_display(summary)
        safe_download_button(summary, "품목별검색.csv", "품목별 검색 결과 다운로드")

# =========================================================
# Tab3: 견적 레퍼런스
# =========================================================
with tab3:
    st.markdown('<div class="section-title">견적 레퍼런스</div>', unsafe_allow_html=True)

    if q.empty:
        st.info("필터 조건에 맞는 데이터가 없습니다.")
    else:
        quote_ref = (
            q.groupby("품목표시", as_index=False)
            .agg(
                평균단가=("단가", "mean"),
                최근단가=("최근단가", "max") if "최근단가" in q.columns else ("단가", "mean"),
                거래처수=("거래처", "nunique"),
                총매출=("매출액", "sum"),
            )
            .sort_values("총매출", ascending=False)
            .reset_index(drop=True)
        )
        clean_and_safe_display(quote_ref)
        safe_download_button(quote_ref, "견적레퍼런스.csv", "견적 레퍼런스 다운로드")

# =========================================================
# Tab4: 매출 하락 분석
# =========================================================
with tab4:
    st.markdown('<div class="section-title">매출 하락 분석</div>', unsafe_allow_html=True)

    if q.empty:
        st.info("필터 조건에 맞는 데이터가 없습니다.")
    else:
        monthly = (
            q.groupby(["거래처", "월"], as_index=False)["매출액"]
            .sum()
            .rename(columns={"매출액": "월매출"})
        )

        all_months = build_month_axis_frame(monthly["월"].dropna().unique())
        rows = []

        for cust_name, sub in monthly.groupby("거래처", dropna=False):
            aligned = align_monthly_series(all_months, sub.rename(columns={"월매출": "값"}), "값")
            vals = aligned["값"].tolist()

            if len(vals) == 0:
                continue

            first_val = vals[0] if len(vals) >= 1 else 0
            last_val = vals[-1] if len(vals) >= 1 else 0

            if first_val == 0:
                change_pct = np.nan if last_val == 0 else 100.0
            else:
                change_pct = ((last_val - first_val) / first_val) * 100.0

            rows.append(
                {
                    "거래처": cust_name,
                    "첫월매출": first_val,
                    "최근월매출": last_val,
                    "변화율(%)": round(change_pct, 2) if pd.notna(change_pct) else np.nan,
                    "추세기울기": round(calc_slope(vals), 4),
                    "변동계수(CV)": round(calc_cv(vals), 4),
                    "총매출": float(pd.to_numeric(sub["월매출"], errors="coerce").fillna(0).sum()),
                }
            )

        decline_df = pd.DataFrame(rows)

        if decline_df.empty:
            st.info("분석할 월별 데이터가 없습니다.")
        else:
            decline_df = decline_df.sort_values(["변화율(%)", "총매출"], ascending=[True, False]).reset_index(drop=True)
            st.caption("매출감소 추이 업체 LIST")
            clean_and_safe_display(decline_df)
            safe_download_button(decline_df, "매출하락분석.csv", "매출 하락 분석 다운로드")

            selected_decline_customer = st.selectbox(
                "상세 분석 거래처 선택",
                options=decline_df["거래처"].tolist(),
                index=0 if len(decline_df) > 0 else None,
                key="tab4_customer_select"
            )

            if selected_decline_customer:
                sub = monthly[monthly["거래처"] == selected_decline_customer].copy()
                aligned = align_monthly_series(all_months, sub.rename(columns={"월매출": "매출액"}), "매출액")

                fig = go.Figure()
                fig.add_trace(
                    go.Scatter(
                        x=aligned["날짜축"],
                        y=aligned["매출액"],
                        mode="lines+markers+text",
                        text=[sales_to_manwon_label(v) for v in aligned["매출액"]],
                        textposition="top center",
                        name="월별 매출"
                    )
                )
                fig = add_trendline(fig, aligned["날짜축"], aligned["매출액"], name="추세선")
                fig.update_layout(
                    height=420,
                    margin=dict(l=20, r=20, t=20, b=20),
                    xaxis_title="월",
                    yaxis_title="매출액",
                )
                st.plotly_chart(fig, use_container_width=True)

# =========================================================
# Tab5: 거래처별 매출 분석
# =========================================================
with tab5:
    st.markdown('<div class="section-title">거래처별 매출 분석</div>', unsafe_allow_html=True)

    analysis = build_customer_sales_analysis(q)
    customer_summary = analysis["customer_summary"]
    customer_monthly = analysis["customer_monthly"]
    customer_item_summary = analysis["customer_item_summary"]
    customer_item_monthly = analysis["customer_item_monthly"]
    all_months = analysis["all_months"]

    if customer_summary.empty:
        st.info("분석할 데이터가 없습니다.")
    else:
        st.caption("매출감소 추이 업체 LIST와 유사한 형식의 거래처별 요약")
        display_summary = customer_summary.copy()
        clean_and_safe_display(display_summary)
        safe_download_button(display_summary, "거래처별매출분석_요약.csv", "거래처별 매출 분석 요약 다운로드")

        customer_list = customer_summary["거래처"].tolist()
        selected_customer = st.selectbox("거래처 선택", customer_list, key="tab5_customer")

        if selected_customer:
            cust_month = customer_monthly[customer_monthly["거래처"] == selected_customer].copy()
            aligned = align_monthly_series(all_months, cust_month.rename(columns={"월매출": "매출액"}), "매출액")

            st.markdown("#### 업체 전체 월별 매출 추이")
            fig = go.Figure()
            fig.add_trace(
                go.Scatter(
                    x=aligned["날짜축"],
                    y=aligned["매출액"],
                    mode="lines+markers+text",
                    text=[sales_to_manwon_label(v) for v in aligned["매출액"]],
                    textposition="top center",
                    name="월별 매출"
                )
            )
            fig = add_trendline(fig, aligned["날짜축"], aligned["매출액"], name="추세선")
            fig.update_layout(
                height=430,
                margin=dict(l=20, r=20, t=20, b=20),
                xaxis_title="월",
                yaxis_title="매출액",
            )
            st.plotly_chart(fig, use_container_width=True)

            st.markdown("#### 품목별 변화율 상세 (총매출순)")
            item_detail = customer_item_summary[customer_item_summary["거래처"] == selected_customer].copy()
            item_detail = item_detail.sort_values("총매출", ascending=False).reset_index(drop=True)

            # 최근가로폭을 가로폭이력 형식으로 보여주기 위한 간단 구성
            cust_raw = q[q["거래처"] == selected_customer].copy()
            if not cust_raw.empty and "품목표시" in cust_raw.columns:
                width_hist = (
                    cust_raw.sort_values("날짜")
                    .groupby("품목표시")["가로폭"]
                    .apply(lambda s: " / ".join([str(int(v)) if pd.notna(v) and float(v).is_integer() else str(v) for v in pd.Series(s).dropna().unique()[:10]]))
                    .reset_index()
                    .rename(columns={"가로폭": "가로폭이력"})
                )
                item_detail = item_detail.merge(width_hist, on="품목표시", how="left")

            show_cols = [c for c in ["품목표시", "총매출", "변화율(%)", "최근월매출", "최근단가", "가로폭이력", "추세기울기", "변동계수(CV)"] if c in item_detail.columns]
            clean_and_safe_display(item_detail[show_cols] if show_cols else item_detail, key="tab5_item_detail")
            safe_download_button(item_detail, f"{selected_customer}_품목별변화율상세.csv", "품목별 변화율 상세 다운로드")

            # 매출 70% 해당 품목 TOP5
            st.markdown("#### 매출 70% 해당 품목의 월별 매출 추이 (TOP5)")
            top_items = item_detail.copy()
            if not top_items.empty:
                total_sales = pd.to_numeric(top_items["총매출"], errors="coerce").fillna(0).sum()
                if total_sales > 0:
                    top_items["누적매출"] = pd.to_numeric(top_items["총매출"], errors="coerce").fillna(0).cumsum()
                    top_items["누적비중"] = top_items["누적매출"] / total_sales
                    top_items_70 = top_items[top_items["누적비중"] <= 0.70].copy()
                    if top_items_70.empty:
                        top_items_70 = top_items.head(1).copy()
                else:
                    top_items_70 = top_items.head(5).copy()

                top_items_70 = top_items_70.head(5).copy()
                st.write("대상 품목:", ", ".join(top_items_70["품목표시"].astype(str).tolist()))

                top_month = customer_item_monthly[
                    (customer_item_monthly["거래처"] == selected_customer) &
                    (customer_item_monthly["품목표시"].isin(top_items_70["품목표시"].tolist()))
                ].copy()

                fig2 = go.Figure()
                for item_name in top_items_70["품목표시"].tolist():
                    sub = top_month[top_month["품목표시"] == item_name].copy()
                    aligned_item = align_monthly_series(all_months, sub.rename(columns={"월매출": "매출액"}), "매출액")
                    fig2.add_trace(
                        go.Scatter(
                            x=aligned_item["날짜축"],
                            y=aligned_item["매출액"],
                            mode="lines+markers",
                            name=item_name
                        )
                    )

                fig2.update_layout(
                    height=450,
                    margin=dict(l=20, r=20, t=20, b=20),
                    xaxis_title="월",
                    yaxis_title="매출액",
                )
                st.plotly_chart(fig2, use_container_width=True)

                st.markdown("#### 변화율 보조 그래프")
                aux_df = top_month.copy()
                aux_df["날짜축"] = pd.to_datetime(aux_df["월"] + "-01", errors="coerce")
                indexed_df = make_indexed_series(aux_df, "품목표시", "월매출", "날짜축")

                if not indexed_df.empty and "품목표시" in indexed_df.columns and "지수값" in indexed_df.columns:
                    fig3 = go.Figure()
                    for item_name, sub in indexed_df.groupby("품목표시", dropna=False):
                        fig3.add_trace(
                            go.Scatter(
                                x=sub["날짜축"],
                                y=sub["지수값"],
                                mode="lines+markers",
                                name=str(item_name)
                            )
                        )
                    fig3.update_layout(
                        height=420,
                        margin=dict(l=20, r=20, t=20, b=20),
                        xaxis_title="월",
                        yaxis_title="지수(기준월=100)"
                    )
                    st.plotly_chart(fig3, use_container_width=True)
                else:
                    st.info("변화율 보조 그래프를 생성할 데이터가 없습니다.")

# =========================================================
# Tab6: 매출 감소 품목 분석
# =========================================================
with tab6:
    st.markdown('<div class="section-title">매출 감소 품목 분석</div>', unsafe_allow_html=True)

    if q.empty:
        st.info("필터 조건에 맞는 데이터가 없습니다.")
    else:
        item_monthly = (
            q.groupby(["품목표시", "월"], as_index=False)["매출액"]
            .sum()
            .rename(columns={"매출액": "월매출"})
        )

        all_months = build_month_axis_frame(item_monthly["월"].dropna().unique())
        rows = []

        for item_name, sub in item_monthly.groupby("품목표시", dropna=False):
            aligned = align_monthly_series(all_months, sub.rename(columns={"월매출": "값"}), "값")
            vals = aligned["값"].tolist()

            if len(vals) == 0:
                continue

            first_val = vals[0] if len(vals) >= 1 else 0
            last_val = vals[-1] if len(vals) >= 1 else 0

            if first_val == 0:
                change_pct = np.nan if last_val == 0 else 100.0
            else:
                change_pct = ((last_val - first_val) / first_val) * 100.0

            rows.append(
                {
                    "품목표시": item_name,
                    "첫월매출": first_val,
                    "최근월매출": last_val,
                    "변화율(%)": round(change_pct, 2) if pd.notna(change_pct) else np.nan,
                    "추세기울기": round(calc_slope(vals), 4),
                    "변동계수(CV)": round(calc_cv(vals), 4),
                    "총매출": float(pd.to_numeric(sub["월매출"], errors="coerce").fillna(0).sum()),
                }
            )

        item_decline_df = pd.DataFrame(rows)

        if item_decline_df.empty:
            st.info("분석할 품목 월별 데이터가 없습니다.")
        else:
            item_decline_df = item_decline_df.sort_values(["변화율(%)", "총매출"], ascending=[True, False]).reset_index(drop=True)
            clean_and_safe_display(item_decline_df)
            safe_download_button(item_decline_df, "매출감소품목분석.csv", "매출 감소 품목 분석 다운로드")

            selected_item = st.selectbox(
                "상세 분석 품목 선택",
                options=item_decline_df["품목표시"].tolist(),
                index=0 if len(item_decline_df) > 0 else None,
                key="tab6_item_select"
            )

            if selected_item:
                sub = item_monthly[item_monthly["품목표시"] == selected_item].copy()
                aligned = align_monthly_series(all_months, sub.rename(columns={"월매출": "매출액"}), "매출액")

                fig = go.Figure()
                fig.add_trace(
                    go.Scatter(
                        x=aligned["날짜축"],
                        y=aligned["매출액"],
                        mode="lines+markers+text",
                        text=[sales_to_manwon_label(v) for v in aligned["매출액"]],
                        textposition="top center",
                        name="월별 매출"
                    )
                )
                fig = add_trendline(fig, aligned["날짜축"], aligned["매출액"], name="추세선")
                fig.update_layout(
                    height=420,
                    margin=dict(l=20, r=20, t=20, b=20),
                    xaxis_title="월",
                    yaxis_title="매출액",
                )
                st.plotly_chart(fig, use_container_width=True)

                raw_item = q[q["품목표시"] == selected_item].copy()
                st.markdown("#### 선택 품목 원자료")
                clean_and_safe_display(raw_item)
                safe_download_button(raw_item, f"{selected_item}_원자료.csv", "선택 품목 원자료 다운로드")

# =========================================================
# Tab7: 원자료
# =========================================================
with tab7:
    st.markdown('<div class="section-title">원자료</div>', unsafe_allow_html=True)
    clean_and_safe_display(q, height=520)
    safe_download_button(q, "원자료.csv", "원자료 다운로드")
