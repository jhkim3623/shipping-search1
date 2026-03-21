import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import os

st.set_page_config(page_title="출고 이력 검색", layout="wide")

# ── 천단위 콤마 포맷 자동 감지 ────────────────────────
def auto_fmt(df):
    fmt = {}
    for col in df.columns:
        if df[col].dtype not in [object, bool] and not str(df[col].dtype).startswith("datetime"):
            if any(k in col for k in ["단가","금액","매출","원","기준단가","기준월","차액","감소액"]):
                fmt[col] = "{:,.0f}"
            elif any(k in col for k in ["M2","총량","수량","기준수량","월평균","판매량"]):
                fmt[col] = "{:,.1f}"
            elif any(k in col for k in ["횟수","업체수","거래처수","개월","순위"]):
                fmt[col] = "{:,.0f}"
            elif any(k in col for k in ["하락률","증감률","비율"]):
                fmt[col] = "{:.1f}%"
    return df.style.format(fmt, na_rep="-")

# ── 피벗 테이블 전용 포맷 ──
def fmt_pivot(df):
    fmt = {}
    for col in df.columns:
        if col in ["품목코드","품목명(공식)","거래처"]:
            continue
        if pd.api.types.is_numeric_dtype(df[col]):
            fmt[col] = "{:,.0f}"
    return df.style.format(fmt, na_rep="-")

# ── 구분별 기준견적가 — 품목 그룹별 교번 배경색 스타일 ──
TIER_COLORS = ["#EBF4FF", "#FFFFFF"]

def style_tier_grouped(df, gc_cols):
    df_reset = df.reset_index(drop=True)
    seen = {}
    group_idx = []
    for _, row in df_reset.iterrows():
        key = tuple(str(row[c]) for c in gc_cols if c in df_reset.columns)
        if key not in seen:
            seen[key] = len(seen)
        group_idx.append(seen[key])

    def row_style(row):
        i = df_reset.index.get_loc(row.name)
        color = TIER_COLORS[group_idx[i] % len(TIER_COLORS)]
        return [f"background-color: {color}"] * len(row)

    num_fmt = {}
    for col in df.select_dtypes(include="number").columns:
        if any(k in col for k in ["단가","금액","매출","원"]):
            num_fmt[col] = "{:,.0f}"
        elif any(k in col for k in ["M2","수량","판매량"]):
            num_fmt[col] = "{:,.1f}"
        else:
            num_fmt[col] = "{:,.0f}"

    return df.style.apply(row_style, axis=1).format(num_fmt, na_rep="-")

def sorted_unique(series):
    if series is None: return []
    s = pd.Series(series).astype(str).str.strip()
    s = s[~s.isin(["","0","0.0","nan","NaN","None","<NA>"])]
    return sorted(s.unique())

@st.cache_data
def load_excel(file_bytes):
    xls = pd.ExcelFile(BytesIO(file_bytes))
    def read_sheet(name, required=True):
        try: return pd.read_excel(xls, name)
        except:
            if required: st.error(f"필수 시트가 없습니다: {name}"); st.stop()
            return pd.DataFrame()

    rec   = read_sheet("출고기록")
    alias = read_sheet("별칭맵핑")
    prod  = read_sheet("품목마스터",  required=False)
    adh   = read_sheet("점착제마스터", required=False)
    cust  = read_sheet("거래처마스터", required=False)

    if rec.empty: st.error("출고기록 시트가 비어있습니다."); st.stop()

    if "날짜" in rec.columns:
        rec["날짜"] = pd.to_datetime(rec["날짜"], errors="coerce")
    for c in ["가로폭(mm)","수량(M2)","단가(원/M2)","금액(원)"]:
        if c in rec.columns: rec[c] = pd.to_numeric(rec[c], errors="coerce")
    if "금액(원)" not in rec.columns and {"수량(M2)","단가(원/M2)"} <= set(rec.columns):
        rec["금액(원)"] = (rec["수량(M2)"] * rec["단가(원/M2)"]).round(0)

    def normalize(df, col):
        if col in df.columns:
            s = df[col].astype(str).str.strip()
            df[col] = s.replace({"":pd.NA,"0":pd.NA,"0.0":pd.NA,"nan":pd.NA,
                                  "NaN":pd.NA,"None":pd.NA,"<NA>":pd.NA})

    for c in ["거래처","품목코드","점착제코드","품목명_고객표현","점착제_고객표현"]:
        if c in rec.columns: rec[c] = rec[c].astype(str).str.strip()
    for c in ["품목코드","점착제코드"]: normalize(rec, c)
    normalize(rec, "거래처")

    def map_alias(series, alias_df, typ):
        if alias_df is None or alias_df.empty or series is None:
            return pd.Series([None]*len(series))
        df = alias_df.copy()
        if not {"유형","별칭","공식코드"} <= set(df.columns):
            return pd.Series([None]*len(series))
        df = df[df["유형"].astype(str)==typ].dropna(subset=["별칭","공식코드"])
        al = df["별칭"].astype(str).str.strip().tolist()
        cl = df["공식코드"].astype(str).str.strip().tolist()
        out = []
        for v in series.astype(str).fillna("").str.strip():
            m = next((c for a,c in zip(al,cl) if v==a), None)
            if m is None: m = next((c for a,c in zip(al,cl) if a and a in v), None)
            out.append(m)
        return pd.Series(out)

    for col,typ in [("품목코드","품목"),("점착제코드","점착제")]:
        src = "품목명_고객표현" if typ=="품목" else "점착제_고객표현"
        if col not in rec.columns or rec[col].isna().any():
            if col not in rec.columns: rec[col] = pd.NA
            if src in rec.columns:
                rec[col] = rec[col].fillna(map_alias(rec[src], alias, typ))
    for c in ["품목코드","점착제코드"]: normalize(rec, c)

    if not prod.empty and {"품목코드","품목명(공식)"} <= set(prod.columns):
        rec = rec.merge(prod[["품목코드","품목명(공식)","품목비고"]].drop_duplicates(),
                        on="품목코드", how="left")
    if not adh.empty and {"점착제코드","점착제명"} <= set(adh.columns):
        rec = rec.merge(adh[["점착제코드","점착제명"]].drop_duplicates(),
                        on="점착제코드", how="left")

    tmp = rec.copy()
    tmp["_cs"] = tmp["거래처"].astype(str).fillna("")
    tmp["_ps"] = tmp["품목코드"].astype(str).fillna("")
    sc = ["_cs","_ps"] + (["날짜"] if "날짜" in tmp.columns else [])
    tmp = tmp.sort_values(sc, kind="mergesort")
    if "단가(원/M2)" in tmp.columns:
        _t  = tmp.dropna(subset=["단가(원/M2)"]).groupby(
            ["거래처","품목코드"], as_index=False, dropna=False).tail(1)
        _ec = ["거래처","품목코드","단가(원/M2)"] + (["날짜"] if "날짜" in _t.columns else [])
        _rm = {"단가(원/M2)":"최근단가"}
        if "날짜" in _t.columns: _rm["날짜"] = "최근날짜"
        lp = _t[_ec].rename(columns=_rm)
        if "최근날짜" in lp.columns:
            lp["최근날짜"] = pd.to_datetime(lp["최근날짜"], errors="coerce").dt.strftime("%Y-%m-%d")
        rec = rec.merge(lp, on=["거래처","품목코드"], how="left")

    def join_unique(s):
        vals = []
        for x in pd.unique(s.dropna()):
            try: xf=float(x); vals.append(str(int(xf)) if xf.is_integer() else str(xf))
            except: vals.append(str(x))
        return ", ".join(vals)

    if "가로폭(mm)" in rec.columns:
        wh = (rec.groupby(["거래처","품목코드"], dropna=False)["가로폭(mm)"]
                 .apply(join_unique).reset_index().rename(columns={"가로폭(mm)":"가로폭이력"}))
        rec = rec.merge(wh, on=["거래처","품목코드"], how="left")
    else: rec["가로폭이력"] = pd.NA
    return rec, alias, prod, adh, cust

DEFAULT_FILE = "data.xlsx"
st.title("출고 이력 검색(거래처/품목/가로폭/점착제)")
uploaded = st.file_uploader("📂 다른 파일 업로드 (미업로드 시 기본 데이터 자동 로드)", type=["xlsx"])

if uploaded:
    file_bytes = uploaded.getvalue(); st.success("✅ 업로드 파일 사용")
elif os.path.exists(DEFAULT_FILE):
    with open(DEFAULT_FILE,"rb") as f: file_bytes = f.read()
    st.info(f"📌 기본 데이터({DEFAULT_FILE}) 자동 로드")
else:
    st.info("GitHub 레포에 data.xlsx를 추가하거나 파일을 업로드하세요."); st.stop()

rec, alias, prod, adh, cust = load_excel(file_bytes)

st.sidebar.header("검색 필터")
sel_cust = st.sidebar.multiselect("거래처",     sorted_unique(rec.get("거래처",    pd.Series())))
sel_prod = st.sidebar.multiselect("품목코드",   sorted_unique(rec.get("품목코드",  pd.Series())))
sel_adh  = st.sidebar.multiselect("점착제코드", sorted_unique(rec.get("점착제코드",pd.Series())))

date_min = pd.to_datetime(rec["날짜"].min()) if "날짜" in rec.columns else None
date_max = pd.to_datetime(rec["날짜"].max()) if "날짜" in rec.columns else None
if date_min and pd.notna(date_min) and date_max and pd.notna(date_max):
    sdate, edate = st.sidebar.date_input("기간", [date_min.date(), date_max.date()])
else: sdate = edate = None

st.sidebar.markdown("---")
st.sidebar.caption("💡 견적 레퍼런스: 품목코드·점착제코드·기간 필터 위주로 사용하세요.")

q = rec.copy()
if sel_cust: q = q[q["거래처"].astype(str).isin(sel_cust)]
if sel_prod: q = q[q["품목코드"].astype(str).isin(sel_prod)]
if sel_adh:  q = q[q["점착제코드"].astype(str).isin(sel_adh)]
if sdate and edate and "날짜" in q.columns:
    q = q[(q["날짜"]>=pd.to_datetime(sdate)) & (q["날짜"]<=pd.to_datetime(edate))]

# ── [변경] 탭 구성: 매출 하락 분석 탭 추가 ──
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "거래처별 검색",
    "품목별 검색", 
    "🏷️ 견적 레퍼런스",
    "📉 매출 하락 분석",
    "원자료"
])

# ══════════════════════════════════════════════════════════
# TAB 1 — 거래처별 검색
# ══════════════════════════════════════════════════════════
with tab1:
    st.subheader("거래처별 → 품목별 출고 이력/최근 단가/가로폭")
    if q.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
    else:
        for c in ["거래처","품목코드","점착제코드"]:
            if c in q.columns: q[c] = q[c].astype(str)
        cols = ["거래처","품목코드","품목명(공식)","점착제코드","점착제명",
                "가로폭이력","최근날짜","최근단가"]
        uc = [c for c in cols if c in q.columns]
        g = (q.groupby(uc, dropna=False)
              .agg(출고횟수=("수량(M2)","count"),
                   총량_M2=("수량(M2)","sum"),
                   매출액=("금액(원)","sum"))
              .reset_index())
        g["가중평균단가"] = np.where(
            g["총량_M2"]>0, (g["매출액"]/g["총량_M2"]).round(0), np.nan)
        sc = [c for c in ["거래처","품목코드"] if c in g.columns]
        st.dataframe(
            auto_fmt(g.sort_values(sc) if sc else g),
            use_container_width=True)

# ══════════════════════════════════════════════════════════
# TAB 2 — 품목별 검색
# ══════════════════════════════════════════════════════════
with tab2:
    st.subheader("품목별 → 거래처별 출고 이력/총량/단가/매출")
    if q.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
    else:
        for c in ["거래처","품목코드"]:
            if c in q.columns: q[c] = q[c].astype(str)
        cols = ["품목코드","품목명(공식)","거래처","최근날짜","최근단가"]
        uc = [c for c in cols if c in q.columns]
        g2 = (q.groupby(uc, dropna=False)
               .agg(출고횟수=("수량(M2)","count"),
                    총량_M2=("수량(M2)","sum"),
                    매출액=("금액(원)","sum"))
               .reset_index())
        g2["가중평균단가"] = np.where(
            g2["총량_M2"]>0, (g2["매출액"]/g2["총량_M2"]).round(0), np.nan)
        sc = [c for c in ["품목코드","거래처"] if c in g2.columns]
        st.dataframe(
            auto_fmt(g2.sort_values(sc) if sc else g2),
            use_container_width=True)

# ══════════════════════════════════════════════════════════
# TAB 3 — 견적 레퍼런스
# ══════════════════════════════════════════════════════════
with tab3:
    import plotly.express as px
    import plotly.graph_objects as go

    st.subheader("🏷️ 견적 레퍼런스 — 기준 견적가 & 판매 동향")
    st.caption(
        "단가 0 및 샘플 품목 자동 제외 | "
        "소형 하위35% / 중형 중간40% / 대형 상위25% | "
        "복합지표: 월판매수량 60% + 월판매금액 40%")

    q_ref = q.copy()
    if "단가(원/M2)" in q_ref.columns:
        q_ref = q_ref[q_ref["단가(원/M2)"].notna() & (q_ref["단가(원/M2)"] > 0)]
    for col in ["품목코드","품목명(공식)"]:
        if col in q_ref.columns:
            q_ref = q_ref[~q_ref[col].astype(str).str.contains("샘플", case=False, na=False)]

    if q_ref.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
    else:
        grp_base = ["품목코드","품목명(공식)","점착제코드","점착제명"]
        GC = [c for c in grp_base if c in q_ref.columns]

        if "날짜" in q_ref.columns:
            vd = q_ref["날짜"].dropna()
            if len(vd):
                d0, d1 = vd.min(), vd.max()
                n_months = max(1,(d1.year-d0.year)*12+d1.month-d0.month+1)
            else: n_months = 1
        else: n_months = 1

        overview = q_ref.groupby(GC, dropna=False).agg(
            최저단가=("단가(원/M2)","min"),
            최고단가=("단가(원/M2)","max"),
            거래처수=("거래처","nunique"),
            총출고횟수=("수량(M2)","count"),
            총량_M2=("수량(M2)","sum"),
            총매출액=("금액(원)","sum"),
        ).reset_index()
        overview["월평균판매량_M2"] = (overview["총량_M2"]/n_months).round(1)
        overview["월평균판매액_원"] = (overview["총매출액"]/n_months).round(0)
        overview["단가범위(원/M2)"] = overview.apply(
            lambda r: f"{int(r['최저단가']):,} ~ {int(r['최고단가']):,}"
                      if pd.notna(r["최저단가"]) and pd.notna(r["최고단가"]) else "-",
            axis=1)

        st.markdown("#### 📋 제품별 개요")
        ov_show = [c for c in GC + ["단가범위(원/M2)","최저단가","최고단가","거래처수",
                                     "월평균판매량_M2","월평균판매액_원","총량_M2","총매출액"]
                   if c in overview.columns]
        sc = [c for c in ["품목코드","점착제코드"] if c in overview.columns]
        st.dataframe(
            auto_fmt(overview[ov_show].sort_values(sc) if sc else overview[ov_show]),
            use_container_width=True)

        cv = (q_ref.groupby(GC+["거래처"], dropna=False)
                   .agg(c_총량=("수량(M2)","sum"), c_총금액=("금액(원)","sum"))
                   .reset_index())
        cv["c_월수량"] = (cv["c_총량"]/n_months).round(1)
        cv["c_월금액"] = (cv["c_총금액"]/n_months).round(0)
        if "날짜" in q_ref.columns and "단가(원/M2)" in q_ref.columns:
            rp = (q_ref.dropna(subset=["단가(원/M2)","날짜"]).sort_values("날짜")
                       .groupby(GC+["거래처"], dropna=False).tail(1)
                       [GC+["거래처","단가(원/M2)"]].rename(columns={"단가(원/M2)":"c_최근단가"}))
            cv = cv.merge(rp, on=GC+["거래처"], how="left")

        tier_parts = []
        for keys, grp in cv.groupby(GC, dropna=False):
            grp = grp.dropna(subset=["c_월수량"]).copy()
            n = len(grp)
            if n == 0: continue
            if n < 3:
                grp["tier"] = "중형"
            else:
                grp["r_vol"] = grp["c_월수량"].rank(pct=True)
                if "c_월금액" in grp.columns and grp["c_월금액"].notna().sum() > 0:
                    grp["r_amt"] = grp["c_월금액"].rank(pct=True)
                    grp["score"] = 0.6*grp["r_vol"] + 0.4*grp["r_amt"]
                else:
                    grp["score"] = grp["r_vol"]
                grp["tier"] = grp["score"].apply(
                    lambda s: "소형" if s<=0.35 else ("대형" if s>0.75 else "중형"))
            tier_parts.append(grp)

        ct = pd.concat(tier_parts, ignore_index=True) if tier_parts else pd.DataFrame()

        if not ct.empty:
            agg_cols = {"업체수":("거래처","count"),
                        "기준수량(M2/월)":("c_월수량","median"),
                        "기준월판매금액(원)":("c_월금액","median")}
            if "c_최근단가" in ct.columns:
                agg_cols["기준단가(원/M2)"] = ("c_최근단가","median")
            tier_agg = ct.groupby(GC+["tier"], dropna=False).agg(**agg_cols).reset_index()
            tier_agg["기준수량(M2/월)"] = tier_agg["기준수량(M2/월)"].round(0)
            if "기준단가(원/M2)" in tier_agg.columns:
                tier_agg["기준단가(원/M2)"] = tier_agg["기준단가(원/M2)"].round(0)
            if "기준월판매금액(원)" in tier_agg.columns:
                tier_agg["기준월판매금액(원)"] = tier_agg["기준월판매금액(원)"].round(0)

            tier_order = {"소형":1,"중형":2,"대형":3}
            tier_agg["_ord"] = tier_agg["tier"].map(tier_order)
            tier_agg = tier_agg.sort_values(GC+["_ord"]).drop(columns=["_ord"])
            tier_agg = tier_agg.rename(columns={"tier":"구분"})

            t_show = [c for c in GC+["구분","업체수","기준수량(M2/월)",
                                      "기준단가(원/M2)","기준월판매금액(원)"]
                      if c in tier_agg.columns]

            st.markdown("#### 🏷️ 구분별 기준 견적가 (세로형 — 모바일 최적화)")
            st.dataframe(
                style_tier_grouped(tier_agg[t_show].reset_index(drop=True), GC),
                use_container_width=True,
                height=400)

            with st.expander("✏️ 기준 견적가 — 수정 가능 버전 (복사·붙여넣기·직접 수정)"):
                st.caption("셀을 클릭하여 값을 직접 수정하거나, Ctrl+C / Ctrl+V 로 복사·붙여넣기 할 수 있습니다.")
                edited_tier = st.data_editor(
                    tier_agg[t_show].reset_index(drop=True),
                    use_container_width=True,
                    num_rows="dynamic",
                    key="tier_editor")
                col_dl1, col_dl2 = st.columns([1,5])
                with col_dl1:
                    csv_bytes = edited_tier.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                    st.download_button(
                        "📥 수정 데이터 CSV 다운로드",
                        data=csv_bytes,
                        file_name="기준견적가_수정본.csv",
                        mime="text/csv")

            st.markdown("---")

            summary_dict: dict = {}
            summary_dict["조회 품목 수"] = f"{len(overview):,}종"
            summary_dict["적용 기간"] = f"{n_months:,}개월"
            if "총량_M2" in overview.columns:
                summary_dict["총 출고량"] = f"{overview['총량_M2'].sum():,.1f} M²"
            if "총매출액" in overview.columns:
                summary_dict["총 매출액"] = f"{overview['총매출액'].sum():,.0f} 원"
            if "월평균판매량_M2" in overview.columns:
                summary_dict["월평균 판매량"] = f"{overview['월평균판매량_M2'].sum():,.1f} M²"
            if "월평균판매액_원" in overview.columns:
                summary_dict["월평균 판매액"] = f"{overview['월평균판매액_원'].sum():,.0f} 원"

            st.markdown("#### 📊 집계 요약")
            st.dataframe(
                pd.DataFrame([summary_dict]),
                use_container_width=True,
                hide_index=True)

            st.markdown("---")

            st.markdown("#### 📊 월별 품목별 판매 추이")
            if ("날짜" in q_ref.columns and "금액(원)" in q_ref.columns
                    and "품목코드" in q_ref.columns):
                cd = q_ref.copy()
                cd["월"] = cd["날짜"].dt.to_period("M").astype(str)
                cd = cd[cd["월"].notna()]
                mp = cd.groupby(["월","품목코드"], dropna=False).agg(
                    매출액=("금액(원)","sum"),
                    판매량=("수량(M2)","sum")).reset_index()
                mp = mp[mp["품목코드"].astype(str).str.lower() != "nan"]
                mt = cd.groupby("월").agg(
                    총매출액=("금액(원)","sum"),
                    총판매량=("수량(M2)","sum")).reset_index()

                col_l, col_r = st.columns([3,1])
                with col_l:
                    fig1 = px.bar(
                        mp, x="월", y="매출액", color="품목코드",
                        barmode="stack",
                        title="월별 품목별 판매금액 (누적 막대)",
                        labels={"매출액":"판매금액(원)","월":""},
                        color_discrete_sequence=px.colors.qualitative.Set2)
                    fig1.add_trace(go.Scatter(
                        x=mt["월"], y=mt["총매출액"],
                        mode="lines+markers+text",
                        name="월 총액",
                        text=[f"{v/1e6:.1f}M" for v in mt["총매출액"]],
                        textposition="top center",
                        line=dict(color="#e74c3c", width=2.5),
                        marker=dict(size=7)))
                    fig1.update_layout(
                        xaxis_tickangle=-45, height=420,
                        legend=dict(orientation="h", yanchor="bottom", y=1.02),
                        yaxis_tickformat=",")
                    st.plotly_chart(fig1, use_container_width=True)

                with col_r:
                    fig2 = px.line(
                        mp, x="월", y="판매량", color="품목코드",
                        title="월별 판매량 (M²)",
                        labels={"판매량":"M²","월":""},
                        markers=True,
                        color_discrete_sequence=px.colors.qualitative.Set2)
                    fig2.update_layout(
                        xaxis_tickangle=-45, height=420,
                        legend=dict(orientation="h", yanchor="bottom", y=1.02),
                        yaxis_tickformat=",")
                    st.plotly_chart(fig2, use_container_width=True)

                # ── [변경 1] 월별 수치 상세 — 행↔열 전치 (품목이 행, 월이 열) ──
                with st.expander("📋 월별 수치 상세 보기 (품목별 월간 판매금액)"):
                    st.caption("💡 행: 품목코드, 열: 월별 판매금액 (편집 및 복사 가능)")
                    
                    # 피벗 테이블: 품목이 행, 월이 열
                    pivot = mp.pivot_table(
                        index="품목코드", columns="월",
                        values="매출액", aggfunc="sum",
                        fill_value=0)
                    
                    # 합계 컬럼 추가
                    month_cols = list(pivot.columns)
                    pivot["합계"] = pivot[month_cols].sum(axis=1)
                    
                    # 합계 기준 내림차순 정렬
                    pivot = pivot.sort_values("합계", ascending=False)
                    
                    # 인덱스 리셋
                    pivot_reset = pivot.reset_index()
                    
                    # 편집 가능한 테이블
                    edited_pivot = st.data_editor(
                        pivot_reset,
                        use_container_width=True,
                        num_rows="dynamic",
                        key="monthly_detail_editor")
                    
                    # CSV 다운로드
                    csv_pivot = edited_pivot.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                    st.download_button(
                        "📥 월별 수치 CSV 다운로드",
                        data=csv_pivot,
                        file_name="월별_품목별_판매금액.csv",
                        mime="text/csv")

            st.markdown("---")

            st.markdown("#### 🔍 단가 역전 현상 분석 (대형 > 중형 단가 감지)")
            if "기준단가(원/M2)" in tier_agg.columns and "구분" in tier_agg.columns:
                anomalies = []
                for keys, grp in tier_agg.groupby(GC, dropna=False):
                    d_large  = grp[grp["구분"]=="대형"]["기준단가(원/M2)"].values
                    d_medium = grp[grp["구분"]=="중형"]["기준단가(원/M2)"].values
                    d_small  = grp[grp["구분"]=="소형"]["기준단가(원/M2)"].values
                    if len(d_large) and len(d_medium):
                        if pd.notna(d_large[0]) and pd.notna(d_medium[0]):
                            if d_large[0] > d_medium[0]:
                                keys_t = keys if isinstance(keys,tuple) else (keys,)
                                prod_label = " / ".join(str(k) for k in keys_t[:2])
                                anomalies.append({
                                    "품목": prod_label,
                                    "소형_단가": d_small[0] if len(d_small) else None,
                                    "중형_단가": d_medium[0],
                                    "대형_단가": d_large[0],
                                    "역전폭(원)": round(d_large[0]-d_medium[0],0)})
                if anomalies:
                    st.warning(
                        f"⚠️ {len(anomalies)}개 품목에서 대형 기준단가가 중형보다 높은 "
                        "역전 현상이 감지되었습니다.")
                    st.dataframe(pd.DataFrame(anomalies), use_container_width=True)
                else:
                    st.success("✅ 모든 품목에서 소형 ≥ 중형 ≥ 대형 단가 순서가 정상입니다.")
            else:
                st.info("단가 데이터가 없어 역전 분석을 수행할 수 없습니다.")

# ══════════════════════════════════════════════════════════
# TAB 4 — [신규] 매출 하락 분석
# ══════════════════════════════════════════════════════════
with tab4:
    st.subheader("📉 매출 하락 업체 분석")
    st.caption("설정 기간 내 전반부 대비 후반부 매출 감소가 큰 상위 35% 업체를 분석합니다.")

    if q.empty or "날짜" not in q.columns or "금액(원)" not in q.columns or "거래처" not in q.columns:
        st.warning("분석에 필요한 데이터(날짜, 금액, 거래처)가 부족합니다.")
    else:
        # 월별 데이터 준비
        decline_df = q.copy()
        decline_df["월"] = decline_df["날짜"].dt.to_period("M").astype(str)
        decline_df = decline_df[decline_df["월"].notna()]
        
        # 전체 월 리스트
        all_months = sorted(decline_df["월"].unique())
        if len(all_months) < 2:
            st.info("분석 기간이 너무 짧습니다. (최소 2개월 필요)")
        else:
            # 전반부/후반부 구분
            mid_idx = len(all_months) // 2
            first_half = all_months[:mid_idx] if mid_idx > 0 else [all_months[0]]
            last_half = all_months[mid_idx:] if mid_idx < len(all_months) else [all_months[-1]]
            
            st.info(f"📅 분석 기준: 전반부 {first_half[0]}~{first_half[-1]} vs 후반부 {last_half[0]}~{last_half[-1]}")
            
            # 거래처별 월별 매출 집계
            monthly_sales = decline_df.groupby(["거래처","월"], dropna=False)["금액(원)"].sum().reset_index()
            
            # 거래처별 전반부/후반부 평균 매출 및 감소액 계산
            decline_analysis = []
            for cust in monthly_sales["거래처"].unique():
                cust_data = monthly_sales[monthly_sales["거래처"] == cust]
                
                # 전반부 평균
                first_data = cust_data[cust_data["월"].isin(first_half)]["금액(원)"]
                avg_first = first_data.mean() if len(first_data) > 0 else 0
                
                # 후반부 평균
                last_data = cust_data[cust_data["월"].isin(last_half)]["금액(원)"]
                avg_last = last_data.mean() if len(last_data) > 0 else 0
                
                # 감소액 (음수면 하락)
                decrease_amount = avg_last - avg_first
                
                # 전체 매출액
                total_sales = cust_data["금액(원)"].sum()
                
                decline_analysis.append({
                    "거래처": cust,
                    "전반부_평균매출": round(avg_first, 0),
                    "후반부_평균매출": round(avg_last, 0),
                    "매출_감소액": round(decrease_amount, 0),
                    "전체_매출액": round(total_sales, 0)
                })
            
            decline_df_result = pd.DataFrame(decline_analysis)
            
            # [변경 2] 매출 감소액 기준 상위 35% (가장 많이 감소한 업체들)
            decline_customers = decline_df_result[decline_df_result["매출_감소액"] < 0].copy()
            
            if decline_customers.empty:
                st.success("✅ 설정 기간 동안 매출이 하락한 업체가 없습니다.")
            else:
                # 감소액 기준 정렬 (가장 많이 감소한 순)
                decline_customers = decline_customers.sort_values("매출_감소액")
                
                # 상위 35% 선정
                n_customers = len(decline_customers)
                top_35_count = max(1, int(np.ceil(n_customers * 0.35)))
                top_decline_customers = decline_customers.head(top_35_count).copy()
                top_decline_customers["순위"] = range(1, len(top_decline_customers) + 1)
                
                # 결과 표시
                col_list, col_detail = st.columns([2, 3])
                
                with col_list:
                    st.markdown(f"#### 🔻 매출 하락 상위 {len(top_decline_customers)}개 업체")
                    
                    # 편집 가능한 업체 목록
                    display_cols = ["순위","거래처","전체_매출액","전반부_평균매출","후반부_평균매출","매출_감소액"]
                    edited_decline = st.data_editor(
                        top_decline_customers[display_cols],
                        use_container_width=True,
                        num_rows="dynamic",
                        key="decline_customers_editor")
                    
                    # CSV 다운로드
                    csv_decline = edited_decline.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                    st.download_button(
                        "📥 하락 업체 목록 CSV 다운로드",
                        data=csv_decline,
                        file_name="매출하락업체_분석.csv",
                        mime="text/csv")
                
                with col_detail:
                    # [변경 3] 업체 선택 → 품목별 월간 분석
                    st.markdown("#### 📊 업체별 품목 분석")
                    
                    selected_customer = st.selectbox(
                        "업체를 선택하여 품목별 상세 분석",
                        options=["선택하세요"] + top_decline_customers["거래처"].tolist(),
                        key="customer_detail_select")
                    
                    if selected_customer != "선택하세요":
                        st.markdown(f"##### 📋 [{selected_customer}] 품목별 월간 매출 추이")
                        
                        # 선택 업체의 품목별 월별 데이터
                        customer_data = decline_df[decline_df["거래처"] == selected_customer]
                        
                        if "품목코드" in customer_data.columns:
                            # 품목별 월별 매출 피벗
                            customer_pivot = customer_data.pivot_table(
                                index="품목코드", columns="월",
                                values="금액(원)", aggfunc="sum",
                                fill_value=0)
                            
                            # [변경 3] 품목별 하락률 계산 및 정렬
                            product_declines = []
                            for prod in customer_pivot.index:
                                prod_data = customer_pivot.loc[prod]
                                
                                # 전반부/후반부 평균
                                first_avg = prod_data[prod_data.index.isin(first_half)].mean()
                                last_avg = prod_data[prod_data.index.isin(last_half)].mean()
                                
                                decline_amount = last_avg - first_avg
                                product_declines.append(decline_amount)
                            
                            # 하락률로 정렬
                            customer_pivot["_decline"] = product_declines
                            customer_pivot = customer_pivot.sort_values("_decline")
                            customer_pivot = customer_pivot.drop(columns=["_decline"])
                            
                            # 합계 컬럼 추가
                            month_cols = list(customer_pivot.columns)
                            customer_pivot["합계"] = customer_pivot[month_cols].sum(axis=1)
                            
                            # 인덱스 리셋
                            customer_pivot_reset = customer_pivot.reset_index()
                            
                            # [변경 4] 편집 가능한 테이블
                            st.caption(f"💡 {selected_customer}의 품목별 월간 매출 (하락폭 큰 순 정렬)")
                            edited_customer_pivot = st.data_editor(
                                customer_pivot_reset,
                                use_container_width=True,
                                num_rows="dynamic",
                                key="customer_product_editor")
                            
                            # CSV 다운로드
                            csv_customer = edited_customer_pivot.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                            st.download_button(
                                "📥 품목별 월간 매출 CSV 다운로드",
                                data=csv_customer,
                                file_name=f"{selected_customer}_품목별_월간매출.csv",
                                mime="text/csv")
                        else:
                            st.warning("품목코드 정보가 없어 품목별 분석을 수행할 수 없습니다.")

# ══════════════════════════════════════════════════════════
# TAB 5 — 원자료
# ══════════════════════════════════════════════════════════
with tab5:
    st.subheader("원자료(필터 적용됨)")
    st.dataframe(auto_fmt(q), use_container_width=True)
