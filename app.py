import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="출고 이력 검색", layout="wide")

# 목록용 유니크 정렬(문자열 통일 + 빈값/0/NAN 제거)
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

    rec  = read_sheet("출고기록")
    alias = read_sheet("별칭맵핑")
    prod = read_sheet("품목마스터", required=False)
    adh  = read_sheet("점착제마스터", required=False)
    cust = read_sheet("거래처마스터", required=False)

    if rec.empty:
        st.error("출고기록 시트가 비어있습니다.")
        st.stop()

    # 날짜/숫자 정리
    if "날짜" in rec.columns:
        rec["날짜"] = pd.to_datetime(rec["날짜"], errors="coerce")

    # 숫자 컬럼 캐스팅
    for c in ["가로폭(mm)", "수량(M2)", "단가(원/M2)", "금액(원)"]:
        if c in rec.columns:
            rec[c] = pd.to_numeric(rec[c], errors="coerce")

    # 금액(원) 자동 계산
    if "금액(원)" not in rec.columns and {"수량(M2)", "단가(원/M2)"} <= set(rec.columns):
        rec["금액(원)"] = (rec["수량(M2)"] * rec["단가(원/M2)"]).round(0)

    # 코드/텍스트 정규화(혼합 타입/0/빈값 처리)
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

    # 별칭 매핑 함수(정확일치 우선, 부분일치 보조)
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
            # 정확 일치
            for a, c in zip(alias_list, code_list):
                if v == a:
                    mapped = c
                    break
            # 부분 일치
            if mapped is None:
                for a, c in zip(alias_list, code_list):
                    if a and a in v:
                        mapped = c
                        break
            out.append(mapped)
        return pd.Series(out)

    # 공식코드 자동 보정(별칭 → 공식)
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

    # 재정규화(별칭 매핑 후 '0', 'nan' 등 제거)
    for c in ["품목코드", "점착제코드"]:
        normalize_code_col(rec, c)

    # 마스터명 매핑
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

    # 최근단가(거래처+품목) 계산 - 정렬 키 분리로 타입 충돌 방지
    tmp = rec.copy()
    tmp["_거래처s"] = tmp["거래처"].astype(str).fillna("")
    tmp["_품목s"]  = tmp["품목코드"].astype(str).fillna("")
    sort_cols = ["_거래처s", "_품목s"]
    if "날짜" in tmp.columns:
        sort_cols.append("날짜")
    tmp = tmp.sort_values(sort_cols, kind="mergesort")
    if "단가(원/M2)" in tmp.columns:
        last_price = (
            tmp.dropna(subset=["단가(원/M2)"])
              .groupby(["거래처", "품목코드"], as_index=False, dropna=False)
              .tail(1)[["거래처", "품목코드", "단가(원/M2)"]]
              .rename(columns={"단가(원/M2)": "최근단가"})
        )
        rec = rec.merge(last_price, on=["거래처", "품목코드"], how="left")

    # 가로폭 이력(거래처+품목)
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


st.title("출고 이력 검색(거래처/품목/가로폭/점착제)")
uploaded = st.file_uploader("엑셀 파일(.xlsx) 업로드: '출고기록', '별칭맵핑' 시트 필수", type=["xlsx"])
if not uploaded:
    st.info("템플릿에 데이터 입력 후 업로드하면 웹에서 즉시 검색할 수 있어요.")
    st.stop()

rec, alias, prod, adh, cust = load_excel(uploaded.getvalue())

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
        # 표시용 정렬 충돌 방지(문자열로 보기)
        for c in ["거래처", "품목코드", "점착제코드"]:
            if c in q.columns:
                q[c] = q[c].astype(str)

        cols = ["거래처","품목코드","품목명(공식)","점착제코드","점착제명","가로폭이력","최근단가"]
        use_cols = [c for c in cols if c in q.columns]
        grp = (
            q.groupby(use_cols, dropna=False)
             .agg(출고횟수=("수량(M2)", "count"),
                  총량_M2=("수량(M2)", "sum"),
                  매출액=("금액(원)", "sum"))
             .reset_index()
        )
        grp["가중평균단가"] = np.where(grp["총량_M2"] > 0, (grp["매출액"] / grp["총량_M2"]).round(0), np.nan)
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

        cols = ["품목코드","품목명(공식)","거래처"]
        use_cols = [c for c in cols if c in q.columns]
        grp2 = (
            q.groupby(use_cols, dropna=False)
             .agg(출고횟수=("수량(M2)", "count"),
                  총량_M2=("수량(M2)", "sum"),
                  매출액=("금액(원)", "sum"))
             .reset_index()
        )
        grp2["가중평균단가"] = np.where(grp2["총량_M2"] > 0, (grp2["매출액"] / grp2["총량_M2"]).round(0), np.nan)
        sort_cols = [c for c in ["품목코드","거래처"] if c in grp2.columns]
        st.dataframe(grp2.sort_values(sort_cols) if sort_cols else grp2)

with tab3:
    st.subheader("원자료(필터 적용됨)")
    st.dataframe(q)
