import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import os
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="출고 이력 검색", layout="wide")

# ── [개선] 안전한 데이터프레임 표시 함수들 ────────────────────────
def format_currency(val):
    """통화 포맷 (천단위 콤마)"""
    try:
        if pd.isna(val) or val == "" or str(val).lower() in ["nan", "none"]:
            return "-"
        return f"{float(val):,.0f}"
    except:
        return str(val)

def format_decimal(val):
    """소수점 1자리 포맷"""
    try:
        if pd.isna(val) or val == "" or str(val).lower() in ["nan", "none"]:
            return "-"
        return f"{float(val):,.1f}"
    except:
        return str(val)

def format_percentage(val):
    """퍼센트 포맷"""
    try:
        if pd.isna(val) or val == "" or str(val).lower() in ["nan", "none"]:
            return "-"
        return f"{float(val):.1f}%"
    except:
        return str(val)

def safe_dataframe_display(df, height=None):
    """안전한 데이터프레임 표시 - 에러 방지"""
    # 인덱스 리셋으로 중복 문제 해결
    display_df = df.reset_index(drop=True).copy()
    
    # 컬럼별 포맷 설정
    column_config = {}
    for col in display_df.columns:
        if col in ["품목코드", "품목명(공식)", "거래처", "점착제코드", "점착제명", 
                   "가로폭이력", "최근날짜", "월", "구분", "분석_내역"]:
            continue
            
        if pd.api.types.is_numeric_dtype(display_df[col]):
            if any(k in col for k in ["단가", "금액", "매출", "원", "기준단가", "기준월", 
                                     "차액", "감소액", "합계", "전반부", "후반부", "전체_매출액"]):
                column_config[col] = st.column_config.NumberColumn(col, format="%,d원")
            elif any(k in col for k in ["M2", "총량", "수량", "기준수량", "월평균", "판매량"]):
                column_config[col] = st.column_config.NumberColumn(col, format="%,.1f")
            elif any(k in col for k in ["하락률", "증감률", "비율", "변화율"]):
                column_config[col] = st.column_config.NumberColumn(col, format="%.1f%%")
            elif any(k in col for k in ["점수", "순위"]):
                column_config[col] = st.column_config.NumberColumn(col, format="%.1f")
    
    # 안전한 표시
    if height:
        st.dataframe(display_df, column_config=column_config, height=height)
    else:
        st.dataframe(display_df, column_config=column_config)

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
        safe_dataframe_display(g.sort_values(sc) if sc else g)

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
        safe_dataframe_display(g2.sort_values(sc) if sc else g2)

# ══════════════════════════════════════════════════════════
# TAB 3 — 견적 레퍼런스 (기존 유지)
# ══════════════════════════════════════════════════════════
with tab3:
    st.subheader("🏷️ 견적 레퍼런스 — 기준 견적가 & 판매 동향")
    st.caption("단가 0 및 샘플 품목 자동 제외 | 소형 하위35% / 중형 중간40% / 대형 상위25% | 복합지표: 월판매수량 60% + 월판매금액 40%")

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

        st.markdown("#### 📋 제품별 개요")
        ov_show = [c for c in GC + ["최저단가","최고단가","거래처수",
                                     "월평균판매량_M2","월평균판매액_원","총량_M2","총매출액"]
                   if c in overview.columns]
        sc = [c for c in ["품목코드","점착제코드"] if c in overview.columns]
        safe_dataframe_display(overview[ov_show].sort_values(sc) if sc else overview[ov_show])

        # 월별 수치 상세보기 (행/열 전치)
        if ("날짜" in q_ref.columns and "금액(원)" in q_ref.columns and "품목코드" in q_ref.columns):
            cd = q_ref.copy()
            cd["월"] = cd["날짜"].dt.to_period("M").astype(str)
            cd = cd[cd["월"].notna()]
            mp = cd.groupby(["월","품목코드"], dropna=False).agg(매출액=("금액(원)","sum")).reset_index()
            mp = mp[mp["품목코드"].astype(str).str.lower() != "nan"]

            with st.expander("📋 월별 수치 상세 보기 (품목별 월간 판매금액)"):
                st.caption("💡 행: 품목코드, 열: 월별 판매금액 (편집 및 복사 가능)")
                
                pivot = mp.pivot_table(index="품목코드", columns="월", values="매출액", aggfunc="sum", fill_value=0)
                month_cols = list(pivot.columns)
                pivot["합계"] = pivot[month_cols].sum(axis=1)
                pivot = pivot.sort_values("합계", ascending=False)
                pivot_reset = pivot.reset_index()
                
                edited_pivot = st.data_editor(pivot_reset, num_rows="dynamic", key="monthly_detail_editor")
                
                csv_pivot = edited_pivot.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                st.download_button("📥 월별 수치 CSV 다운로드", data=csv_pivot, file_name="월별_품목별_판매금액.csv", mime="text/csv")

# ══════════════════════════════════════════════════════════
# TAB 4 — 매출 하락 분석 (완전 개선)
# ══════════════════════════════════════════════════════════
with tab4:
    st.subheader("📉 AI 기반 매출 하락 업체 분석")
    
    with st.expander("ℹ️ 전반부/후반부 매출 산출 방식 및 AI 분석 로직", expanded=False):
        st.markdown("""
### **📊 매출 산출 방식**

**기간 분할:** 전체 분석 기간을 시간 순서대로 정렬하여 중간 지점에서 분할

**수학적 공식:**


$$\\text{전반부 평균매출} = \\frac{\\sum \\text{전반부 월별매출}}{\\text{전반부 개월수}}$$


$$\\text{후반부 평균매출} = \\frac{\\sum \\text{후반부 월별매출}}{\\text{후반부 개월수}}$$


$$\\text{매출 감소액} = \\text{후반부 평균} - \\text{전반부 평균}$$

### **🤖 AI 우선순위 점수 시스템 (총 100점)**

| 지표 | 배점 | 설명 |
|------|------|------|
| **매출 규모** | 30점 | 전체 거래 규모 (큰 손 우선) |
| **하락 심각도** | 25점 | 전반부 대비 후반부 감소율 |
| **품목 다양성 감소** | 20점 | 취급 품목 수 변화 |
| **변동성 증가** | 15점 | 월별 매출 불안정도 |
| **최근 추세** | 10점 | 최근 3개월 동향 |
        """)

    if q.empty or "날짜" not in q.columns or "금액(원)" not in q.columns or "거래처" not in q.columns:
        st.warning("분석에 필요한 데이터(날짜, 금액, 거래처)가 부족합니다.")
    else:
        decline_df = q.copy()
        decline_df["월"] = decline_df["날짜"].dt.to_period("M").astype(str)
        decline_df = decline_df[decline_df["월"].notna()]
        
        all_months = sorted(decline_df["월"].unique())
        if len(all_months) < 2:
            st.info("분석 기간이 너무 짧습니다. (최소 2개월 필요)")
        else:
            mid_idx = len(all_months) // 2
            first_half = all_months[:mid_idx] if mid_idx > 0 else [all_months[0]]
            last_half = all_months[mid_idx:] if mid_idx < len(all_months) else [all_months[-1]]
            
            st.info(f"📅 **분석 기준:** 전반부 {first_half[0]}~{first_half[-1]} ({len(first_half)}개월) vs 후반부 {last_half[0]}~{last_half[-1]} ({len(last_half)}개월)")
            
            monthly_sales = decline_df.groupby(["거래처","월"], dropna=False)["금액(원)"].sum().reset_index()
            
            def calculate_ai_priority_score(cust, cust_monthly, all_cust_data):
                first_data = cust_monthly[cust_monthly["월"].isin(first_half)]["금액(원)"]
                last_data = cust_monthly[cust_monthly["월"].isin(last_half)]["금액(원)"]
                
                avg_first = float(first_data.mean()) if len(first_data) > 0 else 0.0
                avg_last = float(last_data.mean()) if len(last_data) > 0 else 0.0
                total_sales = float(cust_monthly["금액(원)"].sum())
                
                max_sales = all_cust_data.groupby("거래처")["금액(원)"].sum().max()
                sales_score = (total_sales / max_sales) * 30 if max_sales > 0 else 0
                
                decline_rate = ((avg_first - avg_last) / avg_first) if avg_first > 0 else 0
                decline_score = min(max(0, decline_rate) * 25, 25)
                
                if "품목코드" in all_cust_data.columns:
                    cust_detail = all_cust_data[all_cust_data["거래처"] == cust]
                    first_products = cust_detail[cust_detail["월"].isin(first_half)]["품목코드"].nunique()
                    last_products = cust_detail[cust_detail["월"].isin(last_half)]["품목코드"].nunique()
                    product_decline = max(0, first_products - last_products)
                    diversity_score = min((product_decline / max(1, first_products)) * 20, 20)
                else:
                    diversity_score = 0
                
                monthly_amounts = cust_monthly.set_index("월")["금액(원)"]
                if len(monthly_amounts) > 1 and monthly_amounts.mean() > 0:
                    cv = (monthly_amounts.std() / monthly_amounts.mean())
                    volatility_score = min(cv * 15, 15)
                else:
                    volatility_score = 0
                
                recent_months = all_months[-3:] if len(all_months) >= 3 else all_months
                recent_data = cust_monthly[cust_monthly["월"].isin(recent_months)]
                if len(recent_data) >= 2:
                    recent_trend = recent_data["금액(원)"].pct_change().mean()
                    trend_score = max(0, -recent_trend * 10) if not pd.isna(recent_trend) else 0
                    trend_score = min(trend_score, 10)
                else:
                    trend_score = 0
                
                total_score = sales_score + decline_score + diversity_score + volatility_score + trend_score
                
                analysis_text = f"""매출규모: {sales_score:.1f}점 (총 {total_sales:,.0f}원)
하락심각도: {decline_score:.1f}점 (하락률 {decline_rate*100:.1f}%)
품목다양성: {diversity_score:.1f}점
변동성: {volatility_score:.1f}점  
최근추세: {trend_score:.1f}점
종합평가: {"⚠️ 긴급" if total_score > 70 else "🔍 주의" if total_score > 50 else "📋 모니터링"}"""
                
                return {
                    "거래처": cust,
                    "전반부_평균매출": round(avg_first, 0),
                    "후반부_평균매출": round(avg_last, 0),
                    "매출_감소액": round(avg_last - avg_first, 0),
                    "전체_매출액": round(total_sales, 0),
                    "AI_우선순위점수": round(total_score, 1),
                    "분석_내역": analysis_text
                }
            
            st.markdown("#### 🤖 AI 분석 진행 중...")
            progress_bar = st.progress(0)
            
            analysis_results = []
            customers = monthly_sales["거래처"].unique()
            
            for idx, cust in enumerate(customers):
                cust_monthly = monthly_sales[monthly_sales["거래처"] == cust]
                result = calculate_ai_priority_score(cust, cust_monthly, decline_df)
                if result["매출_감소액"] < 0:
                    analysis_results.append(result)
                progress_bar.progress((idx + 1) / len(customers))
            
            progress_bar.empty()
            
            if not analysis_results:
                st.success("✅ 설정 기간 동안 매출이 하락한 업체가 없습니다.")
            else:
                results_df = pd.DataFrame(analysis_results)
                results_df = results_df.sort_values("AI_우선순위점수", ascending=False)
                
                top_count = max(1, int(np.ceil(len(results_df) * 0.35)))
                top_priority = results_df.head(top_count).copy()
                top_priority["순위"] = range(1, len(top_priority) + 1)
                
                # [수직 레이아웃] 우선 대응 업체 리스트
                st.markdown(f"### 🎯 우선 대응 필요 업체 (상위 {len(top_priority)}개)")
                st.caption("💡 AI가 분석한 종합 위험도 점수 기준으로 정렬되었습니다.")
                
                display_cols = ["순위", "거래처", "AI_우선순위점수", "전체_매출액", 
                               "전반부_평균매출", "후반부_평균매출", "매출_감소액", "분석_내역"]
                
                edited_priority = st.data_editor(
                    top_priority[display_cols],
                    column_config={
                        "AI_우선순위점수": st.column_config.NumberColumn("AI 우선순위점수", format="%.1f점"),
                        "전체_매출액": st.column_config.NumberColumn("전체 매출액", format="%,d원"),
                        "전반부_평균매출": st.column_config.NumberColumn("전반부 평균매출", format="%,d원"),
                        "후반부_평균매출": st.column_config.NumberColumn("후반부 평균매출", format="%,d원"),
                        "매출_감소액": st.column_config.NumberColumn("매출 감소액", format="%,d원"),
                        "분석_내역": st.column_config.TextColumn("AI 분석 내역", width="large")
                    },
                    num_rows="dynamic",
                    key="priority_customers_editor"
                )
                
                csv_priority = edited_priority.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                st.download_button("📥 우선 대응 업체 분석 결과 CSV 다운로드", data=csv_priority, file_name="AI분석_우선대응업체.csv", mime="text/csv")
                
                st.markdown("---")
                
                # [수직 레이아웃] 업체별 상세 분석
                st.markdown("### 🔍 업체별 상세 품목 분석")
                
                selected_customer = st.selectbox(
                    "분석할 업체를 선택하세요",
                    options=["선택하세요"] + [f"{row['거래처']} (점수: {row['AI_우선순위점수']:.1f}점)" for _, row in top_priority.iterrows()],
                    key="customer_detail_select"
                )
                
                if selected_customer != "선택하세요":
                    selected_cust_name = selected_customer.split(" (점수:")[0]
                    
                    st.markdown(f"#### 📋 [{selected_cust_name}] 품목별 월간 매출 분석")
                    
                    customer_data = decline_df[decline_df["거래처"] == selected_cust_name]
                    
                    if "품목코드" in customer_data.columns and not customer_data.empty:
                        customer_pivot = customer_data.pivot_table(index="품목코드", columns="월", values="금액(원)", aggfunc="sum", fill_value=0)
                        
                        product_declines = []
                        for prod in customer_pivot.index:
                            prod_data = customer_pivot.loc[prod]
                            first_avg = prod_data[prod_data.index.isin(first_half)].mean()
                            last_avg = prod_data[prod_data.index.isin(last_half)].mean()
                            decline_amount = last_avg - first_avg
                            product_declines.append(decline_amount)
                        
                        customer_pivot["_decline"] = product_declines
                        customer_pivot = customer_pivot.sort_values("_decline")
                        customer_pivot = customer_pivot.drop(columns=["_decline"])
                        
                        month_cols = list(customer_pivot.columns)
                        customer_pivot["합계"] = customer_pivot[month_cols].sum(axis=1)
                        customer_pivot_reset = customer_pivot.reset_index()
                        
                        st.caption(f"💡 {selected_cust_name}의 품목별 월간 매출 (하락 기여도 큰 순)")
                        
                        edited_customer_pivot = st.data_editor(customer_pivot_reset, num_rows="dynamic", key="customer_product_editor")
                        
                        csv_customer = edited_customer_pivot.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                        st.download_button("📥 품목별 월간 매출 CSV 다운로드", data=csv_customer, file_name=f"{selected_cust_name}_품목별_월간매출.csv", mime="text/csv")
                        
                        st.markdown("---")
                        
                        # [수직 레이아웃] 품목별 판매 추이 그래프
                        st.markdown(f"### 📈 [{selected_cust_name}] 품목별 판매 추이 시각화")
                        st.caption("💡 어떤 품목이 매출 하락을 주도하는지 시각적으로 확인하세요.")
                        
                        customer_monthly = customer_data.groupby(["월","품목코드"], dropna=False)["금액(원)"].sum().reset_index()
                        customer_monthly = customer_monthly[customer_monthly["품목코드"].astype(str).str.lower() != "nan"]
                        
                        # 라인 차트
                        fig_line = px.line(customer_monthly, x="월", y="금액(원)", color="품목코드",
                                         title=f"{selected_cust_name} - 품목별 매출 추세", 
                                         labels={"금액(원)":"매출액(원)", "월":""}, markers=True)
                        fig_line.update_layout(xaxis_tickangle=-45, height=400, yaxis_tickformat=",")
                        st.plotly_chart(fig_line, use_container_width=True)
                        
                        # 누적 영역 차트
                        fig_area = px.area(customer_monthly, x="월", y="금액(원)", color="품목코드",
                                         title=f"{selected_cust_name} - 품목별 매출 구성",
                                         labels={"금액(원)":"매출액(원)", "월":""})
                        fig_area.update_layout(xaxis_tickangle=-45, height=400, yaxis_tickformat=",")
                        st.plotly_chart(fig_area, use_container_width=True)
                        
                        # 전반부 vs 후반부 비교 막대 그래프
                        st.markdown("##### 📊 품목별 전반부 vs 후반부 평균 매출 비교")
                        
                        comparison_data = []
                        for prod in customer_pivot_reset["품목코드"]:
                            prod_row = customer_pivot_reset[customer_pivot_reset["품목코드"] == prod]
                            first_cols = [c for c in first_half if c in customer_pivot_reset.columns]
                            last_cols = [c for c in last_half if c in customer_pivot_reset.columns]
                            
                            if first_cols and last_cols:
                                first_avg = prod_row[first_cols].values[0].mean() if first_cols else 0
                                last_avg = prod_row[last_cols].values[0].mean() if last_cols else 0
                                
                                comparison_data.append({
                                    "품목코드": prod,
                                    "전반부_평균": first_avg,
                                    "후반부_평균": last_avg,
                                    "변화액": last_avg - first_avg,
                                    "변화율(%)": ((last_avg - first_avg) / first_avg * 100) if first_avg > 0 else 0
                                })
                        
                        if comparison_data:
                            comp_df = pd.DataFrame(comparison_data)
                            comp_df = comp_df.sort_values("변화액")
                            
                            fig_bar = go.Figure()
                            fig_bar.add_trace(go.Bar(x=comp_df["품목코드"], y=comp_df["전반부_평균"], name="전반부 평균", marker_color="#3498db"))
                            fig_bar.add_trace(go.Bar(x=comp_df["품목코드"], y=comp_df["후반부_평균"], name="후반부 평균", marker_color="#e74c3c"))
                            fig_bar.update_layout(title=f"{selected_cust_name} - 품목별 전반부/후반부 평균 매출 비교", 
                                                barmode="group", height=400, yaxis_tickformat=",")
                            st.plotly_chart(fig_bar, use_container_width=True)
                            
                            st.markdown("##### 📋 품목별 변화율 상세")
                            safe_dataframe_display(comp_df)
                    else:
                        st.warning("해당 업체의 품목별 데이터가 없습니다.")

# ══════════════════════════════════════════════════════════
# TAB 5 — 원자료
# ══════════════════════════════════════════════════════════
with tab5:
    st.subheader("원자료(필터 적용됨)")
    safe_dataframe_display(q)
