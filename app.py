import os
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="출고 이력 검색", layout="wide")

# =========================================================
# 스타일
# =========================================================
st.markdown("""
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
section[data-testid="stSidebar"] .stMultiSelect label,
section[data-testid="stSidebar"] .stDateInput label,
section[data-testid="stSidebar"] .stSelectbox label {
    font-weight: 700 !important;
}
div[data-baseweb="select"] > div {
    min-height: 44px !important;
}
div[data-baseweb="select"] input::placeholder {
    color: #888 !important;
    opacity: 1 !important;
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
</style>
""", unsafe_allow_html=True)

# =========================================================
# 공통 유틸
# =========================================================
def calc_table_height(df, min_rows=3, max_rows=18, row_px=35, header_px=38):
    rows = min_rows if df is None or len(df) == 0 else max(min_rows, min(len(df), max_rows))
    return header_px + rows * row_px

def safe_numeric(series, default=0):
    try:
        s = pd.to_numeric(series, errors="coerce")
        return s.replace([np.inf, -np.inf], np.nan).fillna(default)
    except Exception:
        try:
            return pd.Series([default] * len(series))
        except Exception:
            return pd.Series(dtype=float)

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
    if series is None:
        return []
    try:
        s = pd.Series(series).dropna().astype(str).str.strip()
        s = s[~s.isin(["", "0", "0.0", "nan", "NaN", "None", "<NA>", "NaT"])]
        vals = list(dict.fromkeys(s.tolist()))
        return sorted(vals, key=lambda x: str(x))
    except Exception:
        return []

def sorted_unique_safe(series):
    return sorted_unique(series)

def sales_to_manwon_label(value):
    try:
        if pd.isna(value):
            return ""
        return f"{int(round(float(value) / 10000.0, 0)):,}"
    except Exception:
        return ""

def add_year_month_axis(df, date_col="날짜", out_col="월"):
    temp = df.copy()
    if date_col not in temp.columns:
        temp[out_col] = ""
        return temp
    temp[date_col] = pd.to_datetime(temp[date_col], errors="coerce")
    temp[out_col] = temp[date_col].dt.strftime("%Y-%m")
    temp[out_col] = temp[out_col].fillna("")
    return temp

def build_month_axis_frame(months):
    try:
        month_list = sorted([str(m) for m in months if pd.notna(m) and str(m).strip() != ""])
        df = pd.DataFrame({"월": month_list})
        if df.empty:
            df["날짜축"] = pd.NaT
            return df
        df["날짜축"] = pd.to_datetime(df["월"] + "-01", errors="coerce")
        return df.dropna(subset=["날짜축"]).sort_values("날짜축").reset_index(drop=True)
    except Exception:
        return pd.DataFrame(columns=["월", "날짜축"])

def align_monthly_series(base_month_df, data_df, value_col):
    try:
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
    except Exception:
        out = base_month_df.copy() if base_month_df is not None else pd.DataFrame(columns=["월", "날짜축"])
        out[value_col] = 0
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
    st.dataframe(
        temp,
        use_container_width=True,
        height=height or calc_table_height(temp),
        key=key
    )

def safe_download_button(df, file_name, label):
    try:
        down_df = df.copy() if df is not None else pd.DataFrame()
        csv = down_df.to_csv(index=False, encoding="utf-8-sig")
        st.download_button(
            label=label,
            data=csv,
            file_name=file_name,
            mime="text/csv"
        )
    except Exception:
        st.warning("다운로드 파일 생성 중 오류가 발생했습니다.")

# =========================================================
# 엑셀 로드
# =========================================================
@st.cache_data
def load_excel(file_bytes):
    xls = pd.ExcelFile(BytesIO(file_bytes))
    sheet_names = xls.sheet_names

    rec = pd.read_excel(xls, "출고기록") if "출고기록" in sheet_names else pd.DataFrame()
    alias = pd.read_excel(xls, "별칭맵핑") if "별칭맵핑" in sheet_names else pd.DataFrame()
    prod = pd.read_excel(xls, "품목마스터") if "품목마스터" in sheet_names else pd.DataFrame()
    adh = pd.read_excel(xls, "점착제마스터") if "점착제마스터" in sheet_names else pd.DataFrame()
    cust = pd.read_excel(xls, "거래처마스터") if "거래처마스터" in sheet_names else pd.DataFrame()

    if rec is None or rec.empty:
        rec = pd.DataFrame()

    rec.columns = [str(c).strip() for c in rec.columns]

    essential_defaults = {
        "거래처": "",
        "품목코드": "",
        "점착제코드": "",
        "비고": "",
    }
    for c, v in essential_defaults.items():
        if c not in rec.columns:
            rec[c] = v

    numeric_cols = ["가로폭(mm)", "수량(M2)", "단가(원/M2)", "금액(원)"]
    for c in numeric_cols:
        if c not in rec.columns:
            rec[c] = 0
        rec[c] = pd.to_numeric(rec[c], errors="coerce").replace([np.inf, -np.inf], np.nan).fillna(0)

    if "날짜" not in rec.columns:
        rec["날짜"] = pd.NaT
    rec["날짜"] = pd.to_datetime(rec["날짜"], errors="coerce")

    for d in [alias, prod, adh, cust]:
        if d is not None and not d.empty:
            d.columns = [str(c).strip() for c in d.columns]

    if not prod.empty and "품목코드" in prod.columns:
        prod["품목코드"] = prod["품목코드"].astype(str).str.strip()
        if "품목명(공식)" not in prod.columns:
            if "품목명" in prod.columns:
                prod["품목명(공식)"] = prod["품목명"]
            else:
                prod["품목명(공식)"] = ""
        prod_map = prod[["품목코드", "품목명(공식)"]].drop_duplicates()
        rec["품목코드"] = rec["품목코드"].astype(str).str.strip()
        rec = rec.merge(prod_map, on="품목코드", how="left")
    else:
        if "품목명(공식)" not in rec.columns:
            rec["품목명(공식)"] = ""

    rec = safe_make_product_label(rec)
    return rec, alias, prod, adh, cust

# =========================================================
# 거래처별 매출 분석 데이터 생성
# =========================================================
def build_customer_sales_analysis(df):
    if df is None or df.empty:
        return {
            "df": pd.DataFrame(),
            "all_months": [],
            "customer_summary": pd.DataFrame(),
            "customer_monthly": pd.DataFrame(),
            "customer_item_summary": pd.DataFrame(),
        }

    temp = df.copy()

    need_cols = ["거래처", "품목표시", "날짜", "금액(원)", "수량(M2)", "단가(원/M2)", "가로폭(mm)"]
    for c in need_cols:
        if c not in temp.columns:
            if c in ["금액(원)", "수량(M2)", "단가(원/M2)", "가로폭(mm)"]:
                temp[c] = 0
            elif c == "날짜":
                temp[c] = pd.NaT
            else:
                temp[c] = ""

    temp["거래처"] = temp["거래처"].fillna("").astype(str).str.strip()
    temp["품목표시"] = temp["품목표시"].fillna("").astype(str).str.strip()
    temp["금액(원)"] = pd.to_numeric(temp["금액(원)"], errors="coerce").fillna(0)
    temp["수량(M2)"] = pd.to_numeric(temp["수량(M2)"], errors="coerce").fillna(0)
    temp["단가(원/M2)"] = pd.to_numeric(temp["단가(원/M2)"], errors="coerce").fillna(0)
    temp["가로폭(mm)"] = pd.to_numeric(temp["가로폭(mm)"], errors="coerce")
    temp["날짜"] = pd.to_datetime(temp["날짜"], errors="coerce")
    temp = add_year_month_axis(temp, "날짜", "월")
    temp = temp[temp["거래처"] != ""].copy()

    all_months = sorted_unique(temp["월"])

    customer_summary = (
        temp.groupby("거래처", as_index=False)
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
        temp.groupby(["거래처", "월"], as_index=False)
        .agg(매출액=("금액(원)", "sum"))
        .sort_values(["거래처", "월"])
        .reset_index(drop=True)
    )

    customer_item_monthly = (
        temp.groupby(["거래처", "품목표시", "월"], as_index=False)
        .agg(
            매출액=("금액(원)", "sum"),
            판매량=("수량(M2)", "sum"),
        )
        .sort_values(["거래처", "품목표시", "월"])
        .reset_index(drop=True)
    )

    item_summary = (
        temp.groupby(["거래처", "품목표시"], as_index=False)
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

    item_summary["기준월매출"] = pd.to_numeric(item_summary["기준월매출"], errors="coerce").fillna(0)
    item_summary["최근월매출"] = pd.to_numeric(item_summary["최근월매출"], errors="coerce").fillna(0)

    item_summary["변화율(%)"] = np.where(
        item_summary["기준월매출"] == 0,
        np.nan,
        ((item_summary["최근월매출"] - item_summary["기준월매출"]) / item_summary["기준월매출"]) * 100.0
    )

    item_summary = item_summary.sort_values(["거래처", "총매출액", "품목표시"], ascending=[True, False, True]).reset_index(drop=True)

    return {
        "df": temp,
        "all_months": all_months,
        "customer_summary": customer_summary,
        "customer_monthly": customer_monthly,
        "customer_item_summary": item_summary,
    }

# =========================================================
# 파일 로드
# =========================================================
DEFAULT_FILE = "data.xlsx"

st.markdown('<div class="app-main-title">출고 이력 검색(거래처/품목/가로폭/점착제)</div>', unsafe_allow_html=True)

uploaded = st.file_uploader(
    "📂 다른 파일 업로드 (미업로드 시 기본 데이터 자동 로드)",
    type=["xlsx"]
)

file_bytes = None
if uploaded is not None:
    file_bytes = uploaded.getvalue()
    st.success("✅ 업로드 파일 사용")
elif os.path.exists(DEFAULT_FILE):
    with open(DEFAULT_FILE, "rb") as f:
        file_bytes = f.read()
    st.info(f"📌 기본 데이터({DEFAULT_FILE}) 자동 로드")
else:
    st.warning("data.xlsx 파일이 없습니다. 파일을 업로드해주세요.")
    st.stop()

try:
    rec, alias, prod, adh, cust = load_excel(file_bytes)
except Exception as e:
    st.error(f"엑셀 로드 중 오류가 발생했습니다: {e}")
    st.stop()

if rec is None or rec.empty:
    st.warning("출고기록 데이터가 비어 있습니다.")
    st.stop()

# =========================================================
# 사이드바 필터
# =========================================================
st.sidebar.header("검색 필터")

dept_col = "담당부서" if "담당부서" in rec.columns else ("영업담당부서" if "영업담당부서" in rec.columns else None)
manager_col = "담당자" if "담당자" in rec.columns else None

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

q = rec.copy()

if dept_col and sel_dept:
    q = q[q[dept_col].astype(str).str.strip().isin(sel_dept)]

if manager_col and sel_manager:
    q = q[q[manager_col].astype(str).str.strip().isin(sel_manager)]

if sel_cust and "거래처" in q.columns:
    q = q[q["거래처"].astype(str).str.strip().isin(sel_cust)]

if sel_prod and "품목코드" in q.columns:
    q = q[q["품목코드"].astype(str).str.strip().isin(sel_prod)]

if sel_adh and "점착제코드" in q.columns:
    q = q[q["점착제코드"].astype(str).str.strip().isin(sel_adh)]

date_min = pd.to_datetime(rec["날짜"], errors="coerce").min() if "날짜" in rec.columns else None
date_max = pd.to_datetime(rec["날짜"], errors="coerce").max() if "날짜" in rec.columns else None

sdate, edate = None, None
if pd.notna(date_min) and pd.notna(date_max):
    try:
        default_range = [date_min.date(), date_max.date()] if date_min <= date_max else [date_max.date(), date_max.date()]
        picked = st.sidebar.date_input("기간", value=default_range)
        if isinstance(picked, (list, tuple)) and len(picked) == 2:
            sdate, edate = picked
        elif picked:
            sdate = picked
            edate = picked
    except Exception:
        sdate, edate = None, None

if sdate and edate and "날짜" in q.columns:
    q_date = pd.to_datetime(q["날짜"], errors="coerce")
    q = q[(q_date >= pd.to_datetime(sdate)) & (q_date <= pd.to_datetime(edate))]

# =========================================================
# 탭
# =========================================================
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "거래처별 검색",
    "품목별 검색",
    "🏷️ 견적 레퍼런스",
    "📉 매출 하락 분석",
    "📊 거래처별 매출 분석",
    "매출 감소 품목 분석",
    "원자료",
])

# =========================================================
# 1. 거래처별 검색
# =========================================================
with tab1:
    st.subheader("거래처별 검색")
    if q.empty:
        st.info("검색 결과가 없습니다.")
    else:
        temp = (
            q.groupby("거래처", as_index=False)
            .agg(
                매출액=("금액(원)", "sum"),
                판매량=("수량(M2)", "sum"),
                품목수=("품목코드", "nunique"),
            )
            .sort_values("매출액", ascending=False)
        )
        clean_and_safe_display(temp, key="tab1_customer")
        safe_download_button(temp, "거래처별_검색.csv", "거래처별 검색 결과 다운로드")

# =========================================================
# 2. 품목별 검색
# =========================================================
with tab2:
    st.subheader("품목별 검색")
    if q.empty:
        st.info("검색 결과가 없습니다.")
    else:
        temp = (
            q.groupby("품목표시", as_index=False)
            .agg(
                매출액=("금액(원)", "sum"),
                판매량=("수량(M2)", "sum"),
                거래처수=("거래처", "nunique"),
            )
            .sort_values("매출액", ascending=False)
        )
        clean_and_safe_display(temp, key="tab2_product")
        safe_download_button(temp, "품목별_검색.csv", "품목별 검색 결과 다운로드")

# =========================================================
# 3. 견적 레퍼런스
# =========================================================
with tab3:
    st.subheader("🏷️ 견적 레퍼런스")
    st.caption("현재 필터 기준 원자료를 확인합니다.")
    if q.empty:
        st.info("검색 결과가 없습니다.")
    else:
        cols = [c for c in ["거래처", "품목코드", "품목명(공식)", "점착제코드", "가로폭(mm)", "수량(M2)", "단가(원/M2)", "금액(원)", "날짜"] if c in q.columns]
        ref_df = q[cols].copy()
        if "날짜" in ref_df.columns:
            ref_df = ref_df.sort_values("날짜", ascending=False)
        clean_and_safe_display(ref_df, key="tab3_quote")
        safe_download_button(ref_df, "견적_레퍼런스.csv", "견적 레퍼런스 다운로드")

# =========================================================
# 4. 매출 하락 분석
# =========================================================
with tab4:
    st.subheader("📉 매출 하락 분석")
    st.caption("현재 필터 기준 거래처 매출 증감 현황입니다.")
    if q.empty:
        st.info("검색 결과가 없습니다.")
    else:
        tmp = add_year_month_axis(q, "날짜", "월")
        monthly = (
            tmp.groupby(["거래처", "월"], as_index=False)
            .agg(매출액=("금액(원)", "sum"))
        )
        monthly = monthly[monthly["월"].astype(str).str.strip() != ""]

        if monthly.empty:
            st.info("월별 분석 데이터가 없습니다.")
        else:
            summary = (
                monthly.sort_values(["거래처", "월"])
                .groupby("거래처", as_index=False)
                .agg(
                    기준월매출=("매출액", "first"),
                    최근월매출=("매출액", "last"),
                    평균월매출=("매출액", "mean"),
                )
            )
            summary["변화율(%)"] = np.where(
                summary["기준월매출"] == 0,
                np.nan,
                ((summary["최근월매출"] - summary["기준월매출"]) / summary["기준월매출"]) * 100.0
            )
            summary = summary.sort_values(["변화율(%)", "최근월매출"], ascending=[True, False])
            clean_and_safe_display(summary, key="tab4_decline")
            safe_download_button(summary, "매출_하락_분석.csv", "매출 하락 분석 다운로드")

# =========================================================
# 5. 거래처별 매출 분석
# =========================================================
with tab5:
    st.subheader("📊 거래처별 매출 분석")
    st.caption("업체의 전체 월별 매출 추이와 품목별 변화율 상세를 확인합니다.")

    if q.empty:
        st.info("검색 결과가 없습니다.")
    else:
        pack = build_customer_sales_analysis(q)
        customer_summary = pack["customer_summary"]
        customer_monthly = pack["customer_monthly"]
        customer_item_summary = pack["customer_item_summary"]
        all_months = pack["all_months"]

        if customer_summary.empty:
            st.info("거래처 분석 데이터가 없습니다.")
        else:
            st.markdown('<div class="section-title">1. 매출분석 자료</div>', unsafe_allow_html=True)
            show_summary = customer_summary.copy()
            if "최근일자" in show_summary.columns:
                show_summary["최근일자"] = pd.to_datetime(show_summary["최근일자"], errors="coerce").dt.strftime("%Y-%m-%d")
            clean_and_safe_display(show_summary, key="tab5_summary")
            safe_download_button(show_summary, "거래처별_매출분석_요약.csv", "매출분석 자료 다운로드")

            customer_options = (
                customer_summary.sort_values(["총매출액", "거래처"], ascending=[False, True])["거래처"]
                .dropna().astype(str).tolist()
            )

            if not customer_options:
                st.info("선택 가능한 업체가 없습니다.")
            else:
                selected_customer = st.selectbox(
                    "분석할 업체를 선택하세요",
                    options=customer_options,
                    key="customer_sales_analysis_select"
                )

                st.markdown('<div class="section-title">2. 업체별 상세 분석</div>', unsafe_allow_html=True)

                st.markdown("#### 업체 전체 월별 매출 추이")
                cust_month = customer_monthly[customer_monthly["거래처"] == selected_customer].copy()

                if cust_month.empty:
                    st.info("해당 업체의 월별 매출 데이터가 없습니다.")
                else:
                    month_axis = build_month_axis_frame(all_months if all_months else cust_month["월"].tolist())
                    series = align_monthly_series(month_axis, cust_month[["월", "매출액"]], "매출액")

                    fig = go.Figure()
                    fig.add_trace(go.Scatter(
                        x=series["날짜축"],
                        y=series["매출액"],
                        mode="lines+markers+text",
                        text=[sales_to_manwon_label(v) for v in series["매출액"]],
                        textposition="top center",
                        name="월매출"
                    ))
                    fig.update_layout(
                        height=420,
                        margin=dict(l=30, r=30, t=30, b=30),
                        xaxis_title="월",
                        yaxis_title="매출액(원)",
                        hovermode="x unified"
                    )
                    st.plotly_chart(fig, use_container_width=True)

                st.markdown("#### 품목별 변화율 상세 (총매출순)")
                item_detail = customer_item_summary[customer_item_summary["거래처"] == selected_customer].copy()

                if item_detail.empty:
                    st.info("해당 업체의 품목별 분석 데이터가 없습니다.")
                else:
                    item_detail["총매출액"] = pd.to_numeric(item_detail["총매출액"], errors="coerce").fillna(0)
                    item_detail["총판매량"] = pd.to_numeric(item_detail["총판매량"], errors="coerce").fillna(0)
                    item_detail["평균단가"] = pd.to_numeric(item_detail["평균단가"], errors="coerce").fillna(0)
                    item_detail["변화율(%)"] = pd.to_numeric(item_detail["변화율(%)"], errors="coerce")

                    show_cols = [
                        c for c in [
                            "품목표시", "총매출액", "총판매량", "평균단가",
                            "기준월매출", "최근월매출", "변화율(%)", "최근가로폭", "최근일자"
                        ] if c in item_detail.columns
                    ]

                    show_df = item_detail[show_cols].sort_values(["총매출액", "품목표시"], ascending=[False, True]).reset_index(drop=True)

                    if "최근일자" in show_df.columns:
                        show_df["최근일자"] = pd.to_datetime(show_df["최근일자"], errors="coerce").dt.strftime("%Y-%m-%d")

                    clean_and_safe_display(show_df, key="tab5_item_detail")
                    safe_download_button(show_df, f"거래처별_품목변화율_{selected_customer}.csv", "품목별 변화율 상세 다운로드")

                    product_options = show_df["품목표시"].dropna().astype(str).tolist() if "품목표시" in show_df.columns else []

                    if product_options:
                        selected_product = st.selectbox(
                            "원자료를 확인할 품목을 선택하세요",
                            options=product_options,
                            key="customer_item_raw_select"
                        )

                        raw_cols = [c for c in [
                            "날짜", "거래처", "품목코드", "품목명(공식)", "품목표시",
                            "점착제코드", "가로폭(mm)", "수량(M2)", "단가(원/M2)", "금액(원)", "비고"
                        ] if c in q.columns]

                        raw_df = q.copy()

                        if "거래처" in raw_df.columns:
                            raw_df = raw_df[raw_df["거래처"].astype(str).str.strip() == str(selected_customer).strip()]

                        if "품목표시" in raw_df.columns:
                            raw_df = raw_df[raw_df["품목표시"].astype(str).str.strip() == str(selected_product).strip()]

                        if raw_df.empty:
                            st.info("선택한 품목의 원자료가 없습니다.")
                        else:
                            st.markdown("#### 선택 품목 원자료")
                            if "날짜" in raw_df.columns:
                                raw_df = raw_df.sort_values("날짜", ascending=False)
                            display_raw_df = raw_df[raw_cols].copy()
                            clean_and_safe_display(display_raw_df, key="tab5_raw_detail")
                            safe_download_button(
                                display_raw_df,
                                f"원자료_{selected_customer}_{selected_product}.csv",
                                "선택 품목 원자료 다운로드"
                            )
                    else:
                        st.info("선택 가능한 품목이 없습니다.")

# =========================================================
# 6. 매출 감소 품목 분석
# =========================================================
with tab6:
    st.subheader("매출 감소 품목 분석")
    if q.empty:
        st.info("검색 결과가 없습니다.")
    else:
        tmp = add_year_month_axis(q, "날짜", "월")
        monthly = (
            tmp.groupby(["품목표시", "월"], as_index=False)
            .agg(매출액=("금액(원)", "sum"))
        )
        monthly = monthly[monthly["월"].astype(str).str.strip() != ""]

        if monthly.empty:
            st.info("분석할 데이터가 없습니다.")
        else:
            summary = (
                monthly.sort_values(["품목표시", "월"])
                .groupby("품목표시", as_index=False)
                .agg(
                    기준월매출=("매출액", "first"),
                    최근월매출=("매출액", "last"),
                    평균월매출=("매출액", "mean"),
                )
            )
            summary["변화율(%)"] = np.where(
                summary["기준월매출"] == 0,
                np.nan,
                ((summary["최근월매출"] - summary["기준월매출"]) / summary["기준월매출"]) * 100.0
            )
            summary = summary.sort_values(["변화율(%)", "최근월매출"], ascending=[True, False])
            clean_and_safe_display(summary, key="tab6_product_decline")
            safe_download_button(summary, "매출_감소_품목_분석.csv", "매출 감소 품목 분석 다운로드")

# =========================================================
# 7. 원자료
# =========================================================
with tab7:
    st.subheader("원자료")
    if q.empty:
        st.info("검색 결과가 없습니다.")
    else:
        temp = q.copy()
        if "날짜" in temp.columns:
            temp = temp.sort_values("날짜", ascending=False)
        clean_and_safe_display(temp, key="tab7_raw")
        safe_download_button(temp, "원자료.csv", "원자료 다운로드")
