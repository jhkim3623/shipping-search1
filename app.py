import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import os

st.set_page_config(page_title="출고 이력 검색", layout="wide")

# ── 헬퍼: 숫자 포맷 자동 감지 ─────────────────────────
def auto_fmt(df):
    """컬럼명 키워드로 천단위 포맷 자동 적용"""
    fmt = {}
    for col in df.columns:
        if df[col].dtype not in [object, "datetime64[ns]", bool]:
            if any(k in col for k in ["단가", "금액", "매출", "원", "기준단가"]):
                fmt[col] = "{:,.0f}"
            elif any(k in col for k in ["M2", "총량", "수량", "기준수량", "월평균"]):
                fmt[col] = "{:,.1f}"
            elif any(k in col for k in ["횟수", "수", "개월"]):
                fmt[col] = "{:,.0f}"
    return df.style.format(fmt, na_rep="-")

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
        if not {"유형","별칭","공식코드"} <= set(df.columns):
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
                if v == a: mapped = c; break
            if mapped is None:
                for a, c in zip(alias_list, code_list):
                    if a and a in v: mapped = c; break
            out.append(mapped)
        return pd.Series(out)

    if "품목코드" not in rec.columns or rec["품목코드"].isna().any():
        if "품목코드" not in rec.columns: rec["품목코드"] = pd.NA
        if "품목명_고객표현" in rec.columns:
            rec["품목코드"] = rec["품목코드"].fillna(map_by_alias(rec["품목명_고객표현"], alias, "품목"))
    if "점착제코드" not in rec.columns or rec["점착제코드"].isna().any():
        if "점착제코드" not in rec.columns: rec["점착제코드"] = pd.NA
        if "점착제_고객표현" in rec.columns:
            rec["점착제코드"] = rec["점착제코드"].fillna(map_by_alias(rec["점착제_고객표현"], alias, "점착제"))
    for c in ["품목코드", "점착제코드"]:
        normalize_code_col(rec, c)

    if not prod.empty and {"품목코드","품목명(공식)"} <= set(prod.columns):
        rec = rec.merge(prod[["품목코드","품목명(공식)","품목비고"]].drop_duplicates(), on="품목코드", how="left")
    if not adh.empty and {"점착제코드","점착제명"} <= set(adh.columns):
        rec = rec.merge(adh[["점착제코드","점착제명"]].drop_duplicates(), on="점착제코드", how="left")

    # 최근단가 + 최근날짜 (거래처+품목 기준)
    tmp = rec.copy()
    tmp["_cs"] = tmp["거래처"].astype(str).fillna("")
    tmp["_ps"] = tmp["품목코드"].astype(str).fillna("")
    sc = ["_cs","_ps"]
    if "날짜" in tmp.columns: sc.append("날짜")
    tmp = tmp.sort_values(sc, kind="mergesort")
    if "단가(원/M2)" in tmp.columns:
        _t = tmp.dropna(subset=["단가(원/M2)"]).groupby(["거래처","품목코드"], as_index=False, dropna=False).tail(1)
        _ec = ["거래처","품목코드","단가(원/M2)"]
        _rm = {"단가(원/M2)": "최근단가"}
        if "날짜" in _t.columns: _ec.append("날짜"); _rm["날짜"] = "최근날짜"
        lp = _t[_ec].rename(columns=_rm)
        if "최근날짜" in lp.columns:
            lp["최근날짜"] = pd.to_datetime(lp["최근날짜"], errors="coerce").dt.strftime("%Y-%m-%d")
        rec = rec.merge(lp, on=["거래처","품목코드"], how="left")

    def join_unique(s):
        vals = []
        for x in pd.unique(s.dropna()):
            try:
                xf = float(x); vals.append(str(int(xf)) if xf.is_integer() else str(xf))
            except: vals.append(str(x))
        return ", ".join(vals)

    if "가로폭(mm)" in rec.columns:
        wh = rec.groupby(["거래처","품목코드"], dropna=False)["가로폭(mm)"].apply(join_unique).reset_index().rename(columns={"가로폭(mm)":"가로폭이력"})
        rec = rec.merge(wh, on=["거래처","품목코드"], how="left")
    else:
        rec["가로폭이력"] = pd.NA

    return rec, alias, prod, adh, cust


# ── 기본 파일 자동 로드 ────────────────────────────────
DEFAULT_FILE = "data.xlsx"

st.title("출고 이력 검색(거래처/품목/가로폭/점착제)")
uploaded = st.file_uploader("📂 다른 파일로 조회하려면 업로드 (미업로드 시 기본 데이터 자동 로드)", type=["xlsx"])

if uploaded:
    file_bytes = uploaded.getvalue()
    st.success("✅ 업로드한 파일을 사용합니다.")
elif os.path.exists(DEFAULT_FILE):
    with open(DEFAULT_FILE, "rb") as f:
        file_bytes = f.read()
    st.info(f"📌 기본 데이터({DEFAULT_FILE})를 자동으로 불러왔습니다.")
else:
    st.info("GitHub 레포에 data.xlsx를 추가하거나 파일을 업로드하세요.")
    st.stop()

rec, alias, prod, adh, cust = load_excel(file_bytes)

# ── 사이드바 필터 ──────────────────────────────────────
st.sidebar.header("검색 필터")
cust_list = sorted_unique(rec["거래처"])   if "거래처"   in rec.columns else []
prod_list = sorted_unique(rec["품목코드"]) if "품목코드" in rec.columns else []
adh_list  = sorted_unique(rec["점착제코드"]) if "점착제코드" in rec.columns else []

sel_cust = st.sidebar.multiselect("거래처",    cust_list)
sel_prod = st.sidebar.multiselect("품목코드",  prod_list)
sel_adh  = st.sidebar.multiselect("점착제코드", adh_list)

date_min = pd.to_datetime(rec["날짜"].min()) if "날짜" in rec.columns else None
date_max = pd.to_datetime(rec["날짜"].max()) if "날짜" in rec.columns else None
if date_min is not None and pd.notna(date_min) and date_max is not None and pd.notna(date_max):
    sdate, edate = st.sidebar.date_input("기간", [date_min.date(), date_max.date()])
else:
    sdate, edate = None, None

st.sidebar.markdown("---")
st.sidebar.caption("💡 견적 레퍼런스: 품목코드·점착제코드·기간 위주로 필터하세요.")

# 필터 적용
q = rec.copy()
if sel_cust: q = q[q["거래처"].astype(str).isin(sel_cust)]
if sel_prod: q = q[q["품목코드"].astype(str).isin(sel_prod)]
if sel_adh:  q = q[q["점착제코드"].astype(str).isin(sel_adh)]
if sdate and edate and "날짜" in q.columns:
    q = q[(q["날짜"] >= pd.to_datetime(sdate)) & (q["날짜"] <= pd.to_datetime(edate))]

# ── 탭 구성 ───────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs(["거래처별 검색", "품목별 검색", "🏷️ 견적 레퍼런스", "원자료"])

# ══════════════════════════════════════════════════════
# TAB 1 : 거래처별 검색
# ══════════════════════════════════════════════════════
with tab1:
    st.subheader("거래처별 → 품목별 출고 이력/최근 단가/가로폭")
    if q.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
    else:
        for c in ["거래처","품목코드","점착제코드"]:
            if c in q.columns: q[c] = q[c].astype(str)
        cols = ["거래처","품목코드","품목명(공식)","점착제코드","점착제명","가로폭이력","최근날짜","최근단가"]
        use_cols = [c for c in cols if c in q.columns]
        grp = (q.groupby(use_cols, dropna=False)
                .agg(출고횟수=("수량(M2)","count"), 총량_M2=("수량(M2)","sum"), 매출액=("금액(원)","sum"))
                .reset_index())
        grp["가중평균단가"] = np.where(grp["총량_M2"]>0,(grp["매출액"]/grp["총량_M2"]).round(0),np.nan)
        sc = [c for c in ["거래처","품목코드"] if c in grp.columns]
        st.dataframe(auto_fmt(grp.sort_values(sc) if sc else grp), use_container_width=True)

# ══════════════════════════════════════════════════════
# TAB 2 : 품목별 검색
# ══════════════════════════════════════════════════════
with tab2:
    st.subheader("품목별 → 거래처별 출고 이력/총량/단가/매출")
    if q.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
    else:
        for c in ["거래처","품목코드"]:
            if c in q.columns: q[c] = q[c].astype(str)
        cols = ["품목코드","품목명(공식)","거래처","최근날짜","최근단가"]
        use_cols = [c for c in cols if c in q.columns]
        grp2 = (q.groupby(use_cols, dropna=False)
                 .agg(출고횟수=("수량(M2)","count"), 총량_M2=("수량(M2)","sum"), 매출액=("금액(원)","sum"))
                 .reset_index())
        grp2["가중평균단가"] = np.where(grp2["총량_M2"]>0,(grp2["매출액"]/grp2["총량_M2"]).round(0),np.nan)
        sc = [c for c in ["품목코드","거래처"] if c in grp2.columns]
        st.dataframe(auto_fmt(grp2.sort_values(sc) if sc else grp2), use_container_width=True)

# ══════════════════════════════════════════════════════
# TAB 3 : 견적 레퍼런스 (핵심 신규)
# ══════════════════════════════════════════════════════
with tab3:
    st.subheader("🏷️ 견적 레퍼런스 — 품목별 기준 견적가 & 판매 동향")
    st.caption(
        "📌 **단가 0 및 샘플 품목 자동 제외** | "
        "소형/중형/대형은 해당 품목 거래처들의 월판매수량 하위33%/중위33%/상위33% 기준으로 자동 분류"
    )

    # ── ① 견적 레퍼런스 전용 필터 (0단가·샘플 제외) ──
    q_ref = q.copy()
    if "단가(원/M2)" in q_ref.columns:
        q_ref = q_ref[q_ref["단가(원/M2)"].notna() & (q_ref["단가(원/M2)"] > 0)]
    for col in ["품목코드","품목명(공식)"]:
        if col in q_ref.columns:
            q_ref = q_ref[~q_ref[col].astype(str).str.contains("샘플", case=False, na=False)]

    if q_ref.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
    else:
        # ── ② 그룹 기준 컬럼 ──────────────────────────
        grp_base = ["품목코드","품목명(공식)","점착제코드","점착제명"]
        grp_cols = [c for c in grp_base if c in q_ref.columns]

        # ── ③ 기간 내 월 수 ───────────────────────────
        if "날짜" in q_ref.columns:
            vd = q_ref["날짜"].dropna()
            if len(vd) > 0:
                d0, d1 = vd.min(), vd.max()
                n_months = max(1, (d1.year-d0.year)*12 + d1.month-d0.month+1)
            else:
                n_months = 1
        else:
            n_months = 1

        # ── ④ 전체 집계 (단가범위·거래처수·총량 등) ──
        ref = q_ref.groupby(grp_cols, dropna=False).agg(
            최저단가=("단가(원/M2)","min"),
            최고단가=("단가(원/M2)","max"),
            거래처수=("거래처","nunique"),
            총출고횟수=("수량(M2)","count"),
            총량_M2=("수량(M2)","sum"),
            총매출액=("금액(원)","sum"),
        ).reset_index()
        ref["월평균총량_M2"] = (ref["총량_M2"] / n_months).round(1)
        ref["단가범위(원/M2)"] = ref.apply(
            lambda r: (f"{int(r['최저단가']):,} ~ {int(r['최고단가']):,}"
                       if pd.notna(r["최저단가"]) and pd.notna(r["최고단가"]) else "-"),
            axis=1
        )

        # ── ⑤ 거래처별 월판매수량 + 최근단가 계산 ────
        if "거래처" in q_ref.columns and "수량(M2)" in q_ref.columns:
            cust_vol = (q_ref.groupby(grp_cols + ["거래처"], dropna=False)
                             .agg(c_총량=("수량(M2)","sum")).reset_index())
            cust_vol["c_월평균수량"] = (cust_vol["c_총량"] / n_months).round(1)

            # 거래처별 최근단가
            if "날짜" in q_ref.columns and "단가(원/M2)" in q_ref.columns:
                rp = (q_ref.dropna(subset=["단가(원/M2)","날짜"])
                           .sort_values("날짜")
                           .groupby(grp_cols+["거래처"], dropna=False)
                           .tail(1)[grp_cols+["거래처","단가(원/M2)"]]
                           .rename(columns={"단가(원/M2)":"c_최근단가"}))
                cust_vol = cust_vol.merge(rp, on=grp_cols+["거래처"], how="left")

            # ── ⑥ 소형/중형/대형 분류 (품목별 분위수) ─
            tier_rows = []
            for keys, grp in cust_vol.groupby(grp_cols, dropna=False):
                grp = grp.dropna(subset=["c_월평균수량"]).copy()
                n = len(grp)
                if n == 0: continue
                if n < 3:
                    grp["tier"] = "중형"
                else:
                    p33 = grp["c_월평균수량"].quantile(1/3)
                    p67 = grp["c_월평균수량"].quantile(2/3)
                    grp["tier"] = grp["c_월평균수량"].apply(
                        lambda v: "소형" if v <= p33 else ("대형" if v > p67 else "중형")
                    )
                tier_rows.append(grp)

            if tier_rows:
                ct = pd.concat(tier_rows, ignore_index=True)

                # 티어별 집계
                tier_agg = (ct.groupby(grp_cols+["tier"], dropna=False)
                              .agg(기준수량=("c_월평균수량","median"),
                                   기준단가=("c_최근단가","median"),
                                   업체수=("거래처","count"))
                              .reset_index())
                tier_agg["기준수량"] = tier_agg["기준수량"].round(0)
                tier_agg["기준단가"] = tier_agg["기준단가"].round(0)

                # 소형/중형/대형 컬럼으로 펼치기
                for tier_name in ["소형","중형","대형"]:
                    td = tier_agg[tier_agg["tier"]==tier_name][grp_cols+["기준수량","기준단가","업체수"]].copy()
                    td = td.rename(columns={
                        "기준수량": f"[{tier_name}]기준수량(M2/월)",
                        "기준단가": f"[{tier_name}]기준단가(원/M2)",
                        "업체수":   f"[{tier_name}]업체수",
                    })
                    ref = ref.merge(td, on=grp_cols, how="left")

        # ── ⑦ 최종 컬럼 순서 및 표시 ─────────────────
        ordered = grp_cols + [
            "단가범위(원/M2)", "최저단가", "최고단가",
            "거래처수", "월평균총량_M2",
            "[소형]업체수", "[소형]기준수량(M2/월)", "[소형]기준단가(원/M2)",
            "[중형]업체수", "[중형]기준수량(M2/월)", "[중형]기준단가(원/M2)",
            "[대형]업체수", "[대형]기준수량(M2/월)", "[대형]기준단가(원/M2)",
            "총량_M2", "총출고횟수", "총매출액",
        ]
        show_cols = [c for c in ordered if c in ref.columns]
        sc = [c for c in ["품목코드","점착제코드"] if c in ref.columns]
        out_df = ref[show_cols].sort_values(sc) if sc else ref[show_cols]

        st.dataframe(auto_fmt(out_df), use_container_width=True)

        # 요약 카드
        st.markdown("---")
        ca, cb, cc, cd = st.columns(4)
        ca.metric("조회 품목 수", f"{len(ref):,}종")
        cb.metric("적용 기간", f"{n_months:,}개월")
        if "총량_M2" in ref.columns:
            cc.metric("총 출고량", f"{ref['총량_M2'].sum():,.1f} M²")
        if "총매출액" in ref.columns:
            cd.metric("총 매출액", f"{ref['총매출액'].sum():,.0f} 원")

# ══════════════════════════════════════════════════════
# TAB 4 : 원자료
# ══════════════════════════════════════════════════════
with tab4:
    st.subheader("원자료(필터 적용됨)")
    st.dataframe(auto_fmt(q), use_container_width=True)
