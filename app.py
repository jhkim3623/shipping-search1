import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import os

st.set_page_config(page_title="출고 이력 검색", layout="wide")

# ── [FIX 2] 천단위 콤마 포맷 자동 감지 ────────────────────────
def auto_fmt(df):
    fmt = {}
    for col in df.columns:
        if df[col].dtype not in [object, bool] and not str(df[col].dtype).startswith("datetime"):
            if any(k in col for k in ["단가","금액","매출","원","기준단가","기준월"]):
                fmt[col] = "{:,.0f}"
            elif any(k in col for k in ["M2","총량","수량","기준수량","월평균","판매량"]):
                fmt[col] = "{:,.1f}"
            elif any(k in col for k in ["횟수","업체수","거래처수","개월"]):
                fmt[col] = "{:,.0f}"
    return df.style.format(fmt, na_rep="-")

# ── [FIX 2] 피벗 테이블 전용 포맷 (모든 숫자 컬럼 천단위 콤마) ──
def fmt_pivot(df):
    fmt = {}
    for col in df.columns:
        if col == "월":
            continue
        if pd.api.types.is_numeric_dtype(df[col]):
            fmt[col] = "{:,.0f}"
    return df.style.format(fmt, na_rep="-")

# ── [FIX 3] 구분별 기준견적가 — 품목 그룹별 교번 배경색 스타일 ──
TIER_COLORS = ["#EBF4FF", "#FFFFFF"]  # 파란계열 / 흰색 교번

def style_tier_grouped(df, gc_cols):
    """품목코드 그룹 단위로 행 배경색을 교번 적용"""
    # 그룹 키 → 색상 인덱스 매핑
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

    # 숫자 컬럼 포맷
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
uploaded = st.file_uploader("📂 다른 파일 업로드 (미업로드 시 기본 데이터 자동 로드)",
                             type=["xlsx"])

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

tab1, tab2, tab3, tab4 = st.tabs(["거래처별 검색","품목별 검색","🏷️ 견적 레퍼런스","원자료"])

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

        # ── 제품별 개요 ──────────────────────────────────
        st.markdown("#### 📋 제품별 개요")
        ov_show = [c for c in GC + ["단가범위(원/M2)","최저단가","최고단가","거래처수",
                                     "월평균판매량_M2","월평균판매액_원","총량_M2","총매출액"]
                   if c in overview.columns]
        sc = [c for c in ["품목코드","점착제코드"] if c in overview.columns]
        st.dataframe(
            auto_fmt(overview[ov_show].sort_values(sc) if sc else overview[ov_show]),
            use_container_width=True)

        # ── 거래처별 분류 계산 ───────────────────────────
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

            # ── [FIX 3] 구분별 기준견적가 — 품목 그룹별 교번 색상 ──
            st.markdown("#### 🏷️ 구분별 기준 견적가 (세로형 — 모바일 최적화)")
            st.dataframe(
                style_tier_grouped(tier_agg[t_show].reset_index(drop=True), GC),
                use_container_width=True,
                height=400)

            # ── [FIX 4] 편집 가능 버전 (복사/붙여넣기 & 수정) ──
            with st.expander("✏️ 기준 견적가 — 수정 가능 버전 (복사·붙여넣기·직접 수정)"):
                st.caption("셀을 클릭하여 값을 직접 수정하거나, Ctrl+C / Ctrl+V 로 복사·붙여넣기 할 수 있습니다.")
                edited_tier = st.data_editor(
                    tier_agg[t_show].reset_index(drop=True),
                    use_container_width=True,
                    num_rows="dynamic",
                    key="tier_editor")
                col_dl1, col_dl2 = st.columns([1,5])
                with col_dl1:
                    # 수정된 데이터 CSV 다운로드
                    csv_bytes = edited_tier.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                    st.download_button(
                        "📥 수정 데이터 CSV 다운로드",
                        data=csv_bytes,
                        file_name="기준견적가_수정본.csv",
                        mime="text/csv")

            st.markdown("---")

            # ── [FIX 1] 요약 지표 — st.metric → 컴팩트 데이터프레임 ──
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

            # ── 월별 판매 추이 ───────────────────────────
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

                # ── [FIX 2] 월별 수치 상세 — 천단위 콤마 적용 ──
                with st.expander("📋 월별 수치 상세 보기"):
                    pivot = mp.pivot_table(
                        index="월", columns="품목코드",
                        values="매출액", aggfunc="sum",
                        fill_value=0).reset_index()
                    pivot.columns.name = None  # 컬럼 레벨 이름 제거
                    pivot["월합계"] = pivot.drop(columns="월").sum(axis=1)
                    st.dataframe(fmt_pivot(pivot), use_container_width=True)

            st.markdown("---")

            # ── 단가 역전 분석 ───────────────────────────
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
                    st.markdown("""
**📌 단가 역전 현상의 주요 원인 분석:**
| 원인 | 설명 | 확인 방법 |
|------|------|-----------|
| **① 사양 차이** | 대형 거래처가 특수폭·고점착제 등 고가 사양 위주 구매 | 거래처별 검색 탭에서 가로폭이력·점착제코드 확인 |
| **② 최근 단가 인상** | 대형 거래처에 최근 단가가 인상된 경우 (원자재 가격 반영 등) | 원자료 탭에서 날짜·단가 추이 확인 |
| **③ 거래 집중도** | 대형 거래처라도 특정 품목은 소량만 구매하여 소형처럼 분류됨 | 복합지표 재검토 (수량 비중 높임 추천) |
| **④ 데이터 편향** | 중형 거래처의 할인 계약이 반영된 경우 | 원자료에서 이상치 단가 확인 |
| **⑤ 납기·조건 차이** | 긴급 발주·소량 다빈도 납품으로 인한 프리미엄 단가 | 출고횟수 대비 수량 확인 |
> 💡 **권장 조치**: 역전 품목은 거래처별 검색 탭에서 해당 거래처의 상세 내역을 확인하고,
> 필요시 단가 기준을 수동으로 재검토하세요.
                    """)
                else:
                    st.success("✅ 모든 품목에서 소형 ≥ 중형 ≥ 대형 단가 순서가 정상입니다.")
            else:
                st.info("단가 데이터가 없어 역전 분석을 수행할 수 없습니다.")

# ══════════════════════════════════════════════════════════
# TAB 4 — 원자료
# ══════════════════════════════════════════════════════════
with tab4:
    st.subheader("원자료(필터 적용됨)")
    st.dataframe(auto_fmt(q), use_container_width=True)
