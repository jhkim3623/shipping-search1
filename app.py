import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import os

st.set_page_config(page_title="출고 이력 검색", layout="wide")

def sorted_unique(series):
    if series is None:
        return []
    s = pd.Series(series).astype(str).str.strip()
    s = s[~s.isin(["", "0", "0.0", "nan", "NaN", "None", "<NA>"])]
    return sorted(s.unique())

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

    rec   = read_sheet("출고기록")
    alias = read_sheet("별칭맵핑")
    prod  = read_sheet("품목마스터", required=False)
    adh   = read_sheet("점착제마스터", required=False)
    cust  = read_sheet("거래처마스터", required=False)

    if rec.empty:
        st.error("출고기록 시트가 비어있습니다.")
        st.stop()

    if "날짜" in rec.columns:
        rec["날짜"] = pd.to_datetime(rec["날짜"], errors="coerce")

    for c in ["가로폭(mm)", "수량(M2)", "단가(원/M2)", "금액(원)"]:
        if c in rec.columns:
            rec[c] = pd.to_numeric(rec[c], errors="coerce")

    if "금액(원)" not in rec.columns and {"수량(M2)", "단가(원/M2)"} <= set(rec.columns):
        rec["금액(원)"] = (rec["수량(M2)"] * rec["단가(원/M2)"]).round(0)

    def normalize_code_col(df, col):
        if col in df.columns:
            s = df[col].astype(str).str.strip()
            s = s.replace({"": pd.NA, "0": pd.NA, "0.0": pd.NA,
                           "nan": pd.NA, "NaN": pd.NA, "None": pd.NA, "<NA>": pd.NA})
            df[col] = s

    for c in ["거래처", "품목코드", "점착제코드", "품목명_고객표현", "점착제_고객표현"]:
        if c in rec.columns:
            rec[c] = rec[c].astype(str).str.strip()
    for c in ["품목코드", "점착제코드"]:
        normalize_code_col(rec, c)
    normalize_code_col(rec, "거래처")

    def map_by_alias(series, alias_df, typ):
        if alias_df is None or alias_df.empty or series is None:
            return pd.Series([None] * len(series))
        df = alias_df.copy()
        if "유형" not in df.columns or "별칭" not in df.columns or "공식코드" not in df.columns:
            return pd.Series([None] * len(series))
        df["유형"] = df["유형"].astype(str)
        df = df[df["유형"] == typ].dropna(subset=["별칭", "공식코드"])
        alias_list = df["별칭"].astype(str).str.strip().tolist()
        code_list  = df["공식코드"].astype(str).str.strip().tolist()
        out = []
        vals = series.astype(str).fillna("").str.strip()
        for v in vals:
            mapped = None
            for a, c in zip(alias_list, code_list):
                if v == a:
                    mapped = c
                    break
            if mapped is None:
                for a, c in zip(alias_list, code_list):
                    if a and a in v:
                        mapped = c
                        break
            out.append(mapped)
        return pd.Series(out)

    if "품목코드" not in rec.columns or rec["품목코드"].isna().any():
        if "품목코드" not in rec.columns:
            rec["품목코드"] = pd.NA
        if "품목명_고객표현" in rec.columns:
            fill = map_by_alias(rec["품목명_고객표현"], alias, "품목")
            rec["품목코드"] = rec["품목코드"].fillna(fill)
    if "점착제코드" not in rec.columns or rec["점착제코드"].isna().any():
        if "점착제코드" not in rec.columns:
            rec["점착제코드"] = pd.NA
        if "점착제_고객표현" in rec.columns:
            fill = map_by_alias(rec["점착제_고객표현"], alias, "점착제")
            rec["점착제코드"] = rec["점착제코드"].fillna(fill)

    for c in ["품목코드", "점착제코드"]:
        normalize_code_col(rec, c)

    if not prod.empty and {"품목코드", "품목명(공식)"} <= set(prod.columns):
        rec = rec.merge(
            prod[["품목코드", "품목명(공식)", "품목비고"]].drop_duplicates(),
            on="품목코드", how="left"
        )
    if not adh.empty and {"점착제코드", "점착제명"} <= set(adh.columns):
        rec = rec.merge(
            adh[["점착제코드", "점착제명"]].drop_duplicates(),
            on="점착제코드", how="left"
        )

    # ✅ [수정] 최근단가 + 최근날짜 함께 추출
    tmp = rec.copy()
    tmp["_거래처s"] = tmp["거래처"].astype(str).fillna("")
    tmp["_품목s"]  = tmp["품목코드"].astype(str).fillna("")
    sort_cols = ["_거래처s", "_품목s"]
    if "날짜" in tmp.columns:
        sort_cols.append("날짜")
    tmp = tmp.sort_values(sort_cols, kind="mergesort")

    if "단가(원/M2)" in tmp.columns:
        _tail = (
            tmp.dropna(subset=["단가(원/M2)"])
               .groupby(["거래처", "품목코드"], as_index=False, dropna=False)
               .tail(1)
        )
        _extract_cols = ["거래처", "품목코드", "단가(원/M2)"]
        _rename_map   = {"단가(원/M2)": "최근단가"}
        if "날짜" in _tail.columns:
            _extract_cols.append("날짜")
            _rename_map["날짜"] = "최근날짜"
        last_price = _tail[_extract_cols].rename(columns=_rename_map)
        # 날짜 → YYYY-MM-DD 문자열 포맷
        if "최근날짜" in last_price.columns:
            last_price["최근날짜"] = (
                pd.to_datetime(last_price["최근날짜"], errors="coerce")
                  .dt.strftime("%Y-%m-%d")
            )
        rec = rec.merge(last_price, on=["거래처", "품목코드"], how="left")

    def join_unique(s):
        vals = []
        for x in pd.unique(s.dropna()):
            try:
                xf = float(x)
                vals.append(str(int(xf)) if xf.is_integer() else str(xf))
            except Exception:
                vals.append(str(x))
        return ", ".join(vals)

    if "가로폭(mm)" in rec.columns:
        width_hist = (
            rec.groupby(["거래처", "품목코드"], dropna=False)["가로폭(mm)"]
               .apply(join_unique)
               .reset_index()
               .rename(columns={"가로폭(mm)": "가로폭이력"})
        )
        rec = rec.merge(width_hist, on=["거래처", "품목코드"], how="left")
    else:
        rec["가로폭이력"] = pd.NA

    return rec, alias, prod, adh, cust


# -------------------------------------------------------
# 기본 파일 자동 로드
DEFAULT_FILE = "data.xlsx"
# -------------------------------------------------------

st.title("출고 이력 검색(거래처/품목/가로폭/점착제)")

uploaded = st.file_uploader(
    "📂 다른 파일로 조회하려면 업로드 (미업로드 시 기본 데이터 자동 로드)",
    type=["xlsx"]
)

if uploaded:
    file_bytes = uploaded.getvalue()
    st.success("✅ 업로드한 파일을 사용합니다.")
elif os.path.exists(DEFAULT_FILE):
    with open(DEFAULT_FILE, "rb") as f:
        file_bytes = f.read()
    st.info(f"📌 기본 데이터({DEFAULT_FILE})를 자동으로 불러왔습니다.")
else:
    st.info("템플릿에 데이터 입력 후 업로드하거나, GitHub 레포에 data.xlsx를 추가하세요.")
    st.stop()

rec, alias, prod, adh, cust = load_excel(file_bytes)

# 사이드바 필터
st.sidebar.header("검색 필터")
cust_list = sorted_unique(rec["거래처"]) if "거래처" in rec.columns else []
prod_list = sorted_unique(rec["품목코드"]) if "품목코드" in rec.columns else []
adh_list  = sorted_unique(rec["점착제코드"]) if "점착제코드" in rec.columns else []

sel_cust = st.sidebar.multiselect("거래처", cust_list)
sel_prod = st.sidebar.multiselect("품목코드", prod_list)
sel_adh  = st.sidebar.multiselect("점착제코드", adh_list)

date_min = pd.to_datetime(rec["날짜"].min()) if "날짜" in rec.columns else None
date_max = pd.to_datetime(rec["날짜"].max()) if "날짜" in rec.columns else None
if date_min is not None and pd.notna(date_min) and date_max is not None and pd.notna(date_max):
    sdate, edate = st.sidebar.date_input("기간", [date_min.date(), date_max.date()])
else:
    sdate, edate = None, None

# 필터 적용
q = rec.copy()
if sel_cust:
    q = q[q["거래처"].astype(str).isin(sel_cust)]
if sel_prod:
    q = q[q["품목코드"].astype(str).isin(sel_prod)]
if sel_adh:
    q = q[q["점착제코드"].astype(str).isin(sel_adh)]
if sdate and edate and "날짜" in q.columns:
    q = q[(q["날짜"] >= pd.to_datetime(sdate)) & (q["날짜"] <= pd.to_datetime(edate))]

tab1, tab2, tab3 = st.tabs(["거래처별 검색", "품목별 검색", "원자료"])

with tab1:
    st.subheader("거래처별 → 품목별 출고 이력/최근 단가/가로폭")
    if q.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
    else:
        for c in ["거래처", "품목코드", "점착제코드"]:
            if c in q.columns:
                q[c] = q[c].astype(str)

        # ✅ [수정] 최근날짜를 최근단가 앞에 추가
        cols = ["거래처","품목코드","품목명(공식)","점착제코드","점착제명",
                "가로폭이력","최근날짜","최근단가"]
        use_cols = [c for c in cols if c in q.columns]
        grp = (
            q.groupby(use_cols, dropna=False)
             .agg(출고횟수=("수량(M2)", "count"),
                  총량_M2=("수량(M2)", "sum"),
                  매출액=("금액(원)", "sum"))
             .reset_index()
        )
        grp["가중평균단가"] = np.where(
            grp["총량_M2"] > 0,
            (grp["매출액"] / grp["총량_M2"]).round(0),
            np.nan
        )
        sort_cols = [c for c in ["거래처","품목코드"] if c in grp.columns]
        st.dataframe(grp.sort_values(sort_cols) if sort_cols else grp)

with tab2:
    st.subheader("품목별 → 거래처별 출고 이력/총량/단가/매출")
    if q.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
    else:
        for c in ["거래처", "품목코드"]:
            if c in q.columns:
                q[c] = q[c].astype(str)

        # ✅ [수정] 최근날짜, 최근단가 추가
        cols = ["품목코드","품목명(공식)","거래처","최근날짜","최근단가"]
        use_cols = [c for c in cols if c in q.columns]
        grp2 = (
            q.groupby(use_cols, dropna=False)
             .agg(출고횟수=("수량(M2)", "count"),
                  총량_M2=("수량(M2)", "sum"),
                  매출액=("금액(원)", "sum"))
             .reset_index()
        )
        grp2["가중평균단가"] = np.where(
            grp2["총량_M2"] > 0,
            (grp2["매출액"] / grp2["총량_M2"]).round(0),
            np.nan
        )
        sort_cols = [c for c in ["품목코드","거래처"] if c in grp2.columns]
        st.dataframe(grp2.sort_values(sort_cols) if sort_cols else grp2)

with tab3:
    st.subheader("원자료(필터 적용됨)")
    st.dataframe(q)
