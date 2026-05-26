import os
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="출고 이력 검색", layout="wide")

st.markdown("""
<style>
.block-container {
    padding-top: 0.65rem;
    padding-bottom: 1rem;
    padding-left: 0.7rem;
    padding-right: 0.7rem;
    max-width: 100%;
}

/* 가로 row 정렬 */
div[data-testid=\"stHorizontalBlock\"] {
    gap: 0.6rem;
    align-items: flex-end !important;
}

/* column 내부 wrapper 높이 안정화 */
div[data-testid=\"column\"] > div {
    width: 100% !important;
}

/* 라벨 공통 */
label[data-testid=\"stWidgetLabel\"] {
    margin-bottom: 0.18rem !important;
    padding-bottom: 0 !important;
}
label[data-testid=\"stWidgetLabel\"] p {
    margin: 0 !important;
    font-size: 0.78rem !important;
    line-height: 1.1 !important;
    color: #4b5563 !important;
    font-weight: 600 !important;
}

/* metric box */
div[data-testid=\"stMetric\"] {
    background: #fafafa;
    border: 1px solid #eeeeee;
    border-radius: 10px;
    padding: 0.4rem 0.6rem !important;
}

/* 탭 폰트 스타일 */
button[data-baseweb=\"tab\"] p {
    font-size: 0.88rem !important;
    font-weight: 600 !important;
}
</style>
""", unsafe_allow_html=True)

# 안전한 데이터프레임 렌더링 도우미 함수
def clean_and_safe_display(df, pinned_cols=None, text_cols=None, column_width_overrides=None):
    if df.empty:
        st.info("표시할 데이터가 없습니다.")
        return
    
    tdf = df.copy()
    
    # 1. 날짜 데이터 형식 일치화
    for col in tdf.columns:
        if "날짜" in col or "일자" in col:
            try:
                tdf[col] = pd.to_datetime(tdf[col]).dt.strftime('%Y-%m-%d')
            except:
                pass

    # 2. 숫자 형식 콤마 포맷 및 NaN 예외 처리
    format_dict = {}
    for col in tdf.columns:
        if tdf[col].dtype in [np.int64, np.float64, int, float]:
            if not (col.endswith("코드") or col.endswith("ID") or col.endswith("년") or col.endswith("월")):
                format_dict[col] = "{:,.0f}"
                tdf[col] = tdf[col].fillna(0)
        else:
            tdf[col] = tdf[col].fillna("")

    # 3. 데이터프레임 컬럼 설정 빌드
    col_config = {}
    if text_cols:
        for c in text_cols:
            if c in tdf.columns:
                col_config[c] = st.column_config.TextColumn(c)
                
    if column_width_overrides:
        for c, w in column_width_overrides.items():
            if c in tdf.columns:
                if c in col_config:
                    col_config[c] = st.column_config.TextColumn(c, width=w)
                else:
                    # 기본 컬럼 타입에 너비만 적용
                    if tdf[c].dtype in [np.int64, np.float64, int, float] and not (c.endswith("코드") or c.endswith("ID")):
                        col_config[c] = st.column_config.NumberColumn(c, format="%,.0f", width=w)
                    else:
                        col_config[c] = st.column_config.Column(c, width=w)

    # 4. 고정 컬럼(Pinning) 안정화 처리
    if pinned_cols:
        actual_pinned = [c for c in pinned_cols if c in tdf.columns]
        other_cols = [c for c in tdf.columns if c not in actual_pinned]
        tdf = tdf[actual_pinned + other_cols]

    st.dataframe(
        tdf,
        use_container_width=True,
        hide_index=True,
        column_config=col_config
    )

# 엑셀 파일로드 세션 저장 및 가로 캐싱 최적화
@st.cache_data(ttl=3600)
def load_all_data(file_bytes):
    try:
        xls = pd.ExcelFile(BytesIO(file_bytes))
        sheets = xls.sheet_names
        
        df_records = pd.read_excel(xls, sheet_name="출고기록") if "출고기록" in sheets else pd.DataFrame()
        df_alias = pd.read_excel(xls, sheet_name="별칭맵핑") if "별칭맵핑" in sheets else pd.DataFrame()
        df_item_master = pd.read_excel(xls, sheet_name="품목마스터") if "품목마스터" in sheets else pd.DataFrame()
        df_glue_master = pd.read_excel(xls, sheet_name="점착제마스터") if "점착제마스터" in sheets else pd.DataFrame()
        df_cust_master = pd.read_excel(xls, sheet_name="거래처마스터") if "거래처마스터" in sheets else pd.DataFrame()
        
        # 컬럼 공백 제거 전처리
        for df in [df_records, df_alias, df_item_master, df_glue_master, df_cust_master]:
            if not df.empty:
                df.columns = [str(c).strip() for c in df.columns]
                
        return df_records, df_alias, df_item_master, df_glue_master, df_cust_master
    except Exception as e:
        st.error(f"데이터 파일 읽기 오류: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# ----------------- 앱 메인 레이아웃 -----------------
st.title("📦 출고 이력 종합 검색 시스템")

uploaded_file = st.sidebar.file_uploader("데이터 엑셀파일(data.xlsx) 업로드", type=["xlsx"])

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    df_records, df_alias, df_item_master, df_glue_master, df_cust_master = load_all_data(file_bytes)
    
    if df_records.empty:
        st.warning("'출고기록' 시트 데이터가 없거나 비어있습니다.")
        st.stop()
        
    # 날짜 필드 형변환 처리 및 통합 파생 변수 생성
    df_records["날짜"] = pd.to_datetime(df_records["날짜"])
    df_records["월"] = df_records["날짜"].dt.strftime("%Y-%m")
    
    # 마스터 데이터 결합 처리 (품목명 보완용)
    if not df_item_master.empty and "품목코드" in df_item_master.columns and "품목명(공식)" in df_item_master.columns:
        item_map = df_item_master.set_index("품목코드")["품목명(공식)"].to_dict()
        df_records["품목명(공식)"] = df_records["품목코드"].map(item_map).fillna(df_records["품목코드"])
    else:
        df_records["품목명(공식)"] = df_records["품목코드"]
        
    if not df_glue_master.empty and "점착제코드" in df_glue_master.columns and "점착제명" in df_glue_master.columns:
        glue_map = df_glue_master.set_index("점착제코드")["점착제명"].to_dict()
        df_records["점착제명"] = df_records["점착제코드"].map(glue_map).fillna(df_records["점착제코드"])
    else:
        df_records["점착제명"] = df_records["점착제코드"]

    # 기본 검색 필터 구성 사이드바
    st.sidebar.header("🔍 전역 기간 필터")
    all_months = sorted(df_records["월"].unique(), reverse=True)
    selected_months = st.sidebar.multiselect("조회 월 선택 (미선택시 전체)", options=all_months, default=[])
    
    # 글로벌 필터 컨텍스트 적용 데이터 생성
    q = df_records.copy()
    if selected_months:
        q = q[q["월"].isin(selected_months)]

    # 탭 메뉴 정의
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "🏭 거래처별 검색", "🏷️ 품목별 검색", "🧪 점착제별 분석", 
        "📈 월별 추이", "📊 대시보드", "🏢 업체별 증가현황", "📝 원자료(필터)"
    ])

    # =========================================================================
    # [TAB 1] 거래처별 검색 (수정 핵심 반영 부분)
    # =========================================================================
    with tab1:
        st.subheader("🏭 거래처별 출고 기록 및 단가 변동 내역")
        
        cust_list = sorted(q["거래처"].dropna().unique())
        c1, c2 = st.columns([2, 2])
        with c1:
            selected_cust = st.selectbox("거래처 선택", options=["선택하세요"] + cust_list)
            
        if selected_cust != "선택하세요":
            cust_df = q[q["거래처"] == selected_cust]
            
            with c2:
                cust_items = sorted(cust_df["품목코드"].dropna().unique())
                selected_item = st.selectbox("품목코드 선택 (선택 사항)", options=["전체 품목"] + cust_items)
                
            if selected_item != "전체 품목":
                cust_df = cust_df[cust_df["품목코드"] == selected_item]
                
            # 기본 통계 메트릭 산출
            m1, m2, m3 = st.columns(3)
            with m1:
                st.metric("총 출고 수량", f"{cust_df['수량(M2)'].sum():,.0f} M²")
            with m2:
                st.metric("총 매출 금액", f"{cust_df['금액(원)'].sum():,.0f} 원")
            with m3:
                avg_p = cust_df['금액(원)'].sum() / cust_df['수량(M2)'].sum() if cust_df['수량(M2)'].sum() > 0 else 0
                st.metric("가중 평균 단가", f"{avg_p:,.1f} 원/M²")
                
            st.markdown("---")
            
            # 1단계: 최근 단가 요약 테이블
            st.markdown("#### 📌 최근 출고 품목별 단가 요약")
            idx_cols = ["품목코드", "품목명(공식)", "점착제코드"]
            # 추가된 컬럼이 있으면 유연하게 그룹바이에 인입 처리 (에러 방지 안전장치)
            for extra_col in ["재단구분", "기준폭"]:
                if extra_col in cust_df.columns:
                    idx_cols.append(extra_col)
                    
            recent_idx = cust_df.groupby("품목코드")["날짜"].idxmax()
            df_recent_price = cust_df.loc[recent_idx, idx_cols + ["날짜", "단가(원/M2)"]].rename(columns={"날짜": "최근출고일"})
            
            clean_and_safe_display(
                df_recent_price,
                pinned_cols=["품목코드"],
                text_cols=["품목코드", "품목명(공식)", "점착제코드", "재단구분"],
                column_width_overrides={"품목코드": 130, "품목명(공식)": 220, "단가(원/M2)": 100}
            )
            
            # 2단계: 가로폭 이력 이력 요약 테이블 (★ 요청 사항 반영!)
            st.markdown("#### 📐 품목별 규격 및 가로폭 출고 이력 요약")
            
            # 그룹바이 대상 가변 지정 (재단구분, 기준폭 안전 추가)
            base_group_cols = ["품목코드", "품목명(공식)", "점착제코드"]
            for extra_col in ["재단구분", "기준폭"]:
                if extra_col in cust_df.columns:
                    base_group_cols.append(extra_col)
            
            # 가로폭 이력 데이터 애그리게이션 연산
            def get_width_history(group):
                # 날짜 역순 정렬
                group_sorted = group.sort_values(by="날짜", ascending=False)
                # 폭 유니크 리스트업
                widths = group_sorted["가로폭(mm)"].unique()
                width_str = ", ".join([f"{w:,.0f}mm" for w in widths])
                
                res = {
                    "가로폭이력": width_str,
                    "최근날짜": group_sorted["날짜"].max(),
                    "총출고량(M2)": group_sorted["수량(M2)"].sum(),
                    "총매출액(원)": group_sorted["금액(원)"].sum()
                }
                return pd.Series(res)
                
            df_width_summary = cust_df.groupby(base_group_cols, dropna=False).apply(get_width_history).reset_index()
            
            # 테이블 컬럼 표시 순서 재정의 (★ 사진 명시 조건: 가로폭이력 바로 앞에 재단구분, 기준폭 정렬 배치)
            desired_order = ["품목코드", "품목명(공식)", "점착제코드"]
            for ec in ["재단구분", "기준폭"]:
                if ec in df_width_summary.columns:
                    desired_order.append(ec)
            desired_order += ["가로폭이력", "최근날짜", "총출고량(M2)", "총매출액(원)"]
            
            # 존재하는 컬럼만 최종 순서 리스트로 확정
            final_order = [c for c in desired_order if c in df_width_summary.columns]
            df_width_summary = df_width_summary[final_order]
            
            # 가로폭 폭 커스텀 가로 길이 세팅 지정 및 렌더링
            width_overrides = {
                "품목코드": 130, "품목명(공식)": 200, "점착제코드": 90,
                "재단구분": 90, "기준폭": 85, "가로폭이력": 220,
                "최근날짜": 100, "총출고량(M2)": 100, "총매출액(원)": 110
            }
            clean_and_safe_display(
                df_width_summary,
                pinned_cols=["품목코드"],
                text_cols=["품목코드", "품목명(공식)", "점착제코드", "재단구분", "가로폭이력"],
                column_width_overrides=width_overrides
            )
            
            # 3단계: 필터 적용된 전체 세부 이력 출력
            st.markdown("#### 📋 상세 출고 내역 (필터 적용됨)")
            # 상세 표기 시에도 재단구분, 기준폭 포함 전체 자동 표기
            clean_and_safe_display(
                cust_df.sort_values(by="날짜", ascending=False),
                pinned_cols=["날짜", "품목코드"],
                text_cols=["품목코드", "재단구분", "점착제코드", "비고", "담당자"],
                column_width_overrides={"비고": 200, "품목코드": 130}
            )

    # =========================================================================
    # [TAB 2] 품목별 검색 (안전조치: 고정 인덱스 슬라이싱 제거하고 컬럼 매핑으로 전환)
    # =========================================================================
    with tab2:
        st.subheader("🏷️ 품목별 출고 트렌드 및 거래처 현황")
        item_list = sorted(q["품목코드"].dropna().unique())
        selected_item_tab2 = st.selectbox("검색할 품목코드 선택", options=item_list)
        
        if selected_item_tab2:
            item_df = q[q["품목코드"] == selected_item_tab2]
            
            # 공식마스터 매핑 기준 타이틀 노출
            p_name = item_df["품목명(공식)"].iloc[0] if not item_df.empty else ""
            st.info(f"💡 **선택 품목 공식명칭:** {p_name}")
            
            ti1, ti2 = st.columns(2)
            with ti1:
                cust_sum = item_df.groupby("거래처").agg(
                    총량_M2=("수량(M2)", "sum"),
                    매출액_원=("금액(원)", "sum")
                ).reset_index()
                cust_sum["가중평균단가"] = cust_sum["매출액_원"] / cust_sum["총량_M2"]
                cust_sum = cust_sum.sort_values(by="총량_M2", ascending=False)
                st.markdown("##### 🏢 거래처별 출고 요약")
                clean_and_safe_display(cust_sum, pinned_cols=["거래처"])
                
            with ti2:
                st.markdown("##### 📊 월별 출고량 추이")
                monthly_item = item_df.groupby("월")["수량(M2)"].sum().reset_index()
                fig_item = go.Figure(data=[go.Bar(x=monthly_item["월"], y=monthly_item["수량(M2)"], marker_color="#3b82f6")])
                fig_item.update_layout(margin=dict(l=20, r=20, t=20, b=20), height=240, xaxis_type='category')
                st.plotly_chart(fig_item, use_container_width=True)

    # =========================================================================
    # [TAB 3] 점착제별 분석
    # =========================================================================
    with tab3:
        st.subheader("🧪 점착제 유형별 출고 실적 통계")
        if "점착제코드" in q.columns:
            glue_list = sorted(q["점착제코드"].dropna().unique())
            selected_glue = st.selectbox("분석할 점착제코드 선택", options=glue_list)
            
            if selected_glue:
                glue_df = q[q["점착제코드"] == selected_glue]
                g_name = glue_df["점착제명"].iloc[0] if not glue_df.empty else ""
                st.caption(f"🧪 **점착제 공식명:** {g_name}")
                
                g_cust = glue_df.groupby("거래처")["수량(M2)"].sum().reset_index().sort_values(by="수량(M2)", ascending=False).head(15)
                g_item = glue_df.groupby("품목코드")["수량(M2)"].sum().reset_index().sort_values(by="수량(M2)", ascending=False).head(15)
                
                gc1, gc2 = st.columns(2)
                with gc1:
                    st.markdown("##### 🏆 주요 사용 거래처 (Top 15)")
                    clean_and_safe_display(g_cust)
                with gc2:
                    st.markdown("##### 🏆 주요 적용 품목코드 (Top 15)")
                    clean_and_safe_display(g_item)

    # =========================================================================
    # [TAB 4] 월별 추이
    # =========================================================================
    with tab4:
        st.subheader("📈 기업 전체 월별 출고 추이 분석")
        monthly_trend = q.groupby("월").agg(
            총출고량_M2=("수량(M2)", "sum"),
            총매출액_원=("금액(원)", "sum"),
            출고건수=("날짜", "count")
        ).reset_index().sort_values(by="월")
        
        tc1, tc2 = st.columns([3, 2])
        with tc1:
            fig_trend = go.Figure()
            fig_trend.add_trace(go.Scatter(x=monthly_trend["월"], y=monthly_trend["총매출액_원"], name="매출액(원)", mode="lines+markers", line=dict(color="#ef4444", width=3)))
            fig_trend.update_layout(title="월별 총 매출액 추이", height=300, margin=dict(t=40, b=20, l=20, r=20), xaxis_type='category')
            st.plotly_chart(fig_trend, use_container_width=True)
        with tc2:
            st.markdown("##### 📅 월별 실적 집계표")
            clean_and_safe_display(monthly_trend)

    # =========================================================================
    # [TAB 5] 대시보드
    # =========================================================================
    with tab5:
        st.subheader("📊 출고 현황 종합 요약 대시보드")
        db1, db2 = st.columns(2)
        with db1:
            st.markdown("##### 🏆 거래처별 매출 순위 (Top 15)")
            top_cust = q.groupby("거래처")["금액(원)"].sum().reset_index().sort_values(by="금액(원)", ascending=False).head(15)
            clean_and_safe_display(top_cust)
        with db2:
            st.markdown("##### 🏆 품목별 출고량 순위 (Top 15)")
            top_items = q.groupby(["품목코드", "품목명(공식)"])["수량(M2)"].sum().reset_index().sort_values(by="수량(M2)", ascending=False).head(15)
            clean_and_safe_display(top_items, pinned_cols=["품목코드"])

    # =========================================================================
    # [TAB 6] 업체별 증가현황
    # =========================================================================
    with tab6:
        st.subheader("🏢 전월 대비 업체별 출고 증가 현황 조회")
        if len(all_months) >= 2:
            c_m1, c_m2 = st.columns(2)
            with c_m1:
                m_base = st.selectbox("기준월(신규)", options=all_months, index=0)
            with c_m2:
                m_comp = st.selectbox("비교월(과거)", options=all_months, index=1)
                
            if m_base != m_comp:
                df_base = df_records[df_records["월"] == m_base].groupby("거래처")["수량(M2)"].sum()
                df_comp = df_records[df_records["월"] == m_comp].groupby("거래처")["수량(M2)"].sum()
                
                df_diff = pd.DataFrame(index=df_base.index.union(df_comp.index))
                df_diff[f"{m_base} 수량"] = df_diff.index.map(df_base).fillna(0)
                df_diff[f"{m_comp} 수량"] = df_diff.index.map(df_comp).fillna(0)
                df_diff["증감량(M2)"] = df_diff[f"{m_base} 수량"] - df_diff[f"{m_comp} 수량"]
                df_diff = df_diff.reset_index().rename(columns={"index": "거래처"}).sort_values(by="증감량(M2)", ascending=False)
                
                clean_and_safe_display(df_diff, pinned_cols=["거래처"])
        else:
            st.info("업체별 증가현황 비교를 위한 데이터 기간이 부족합니다.")

    # =========================================================================
    # [TAB 7] 원자료 (필터 적용됨 - ★ 요청 사항 반영)
    # =========================================================================
    with tab7:
        st.subheader("📝 원자료 (필터 및 신규 추가 규격 열 포함)")
        
        # 품목명(공식) 및 점착제명을 제외한 원본 데이터의 실시간 컬럼 목록 생성
        # 데이터에 재단구분, 기준폭이 추가되면 해당 위치에 동적으로 바인딩되어 출력됩니다.
        raw_cols = [c for c in q.columns if c not in ["품목명(공식)", "점착제명"]]
        
        # 원자료 탭의 가로폭 표시 오버라이드 및 고정 너비 설정
        raw_width_overrides = {
            "날짜": 95,
            "거래처": 140,
            "담당부서": 90,
            "담당자": 80,
            "재단구분": 85,  # 신규 추가 컬럼 배치 가시성 확보
            "품목코드": 140,
            "점착제코드": 90,
            "기준폭": 80,    # 신규 추가 컬럼 배치 가시성 확보
            "가로폭(mm)": 85,
            "수량(M2)": 95,
            "단가(원/M2)": 95,
            "금액(원)": 105,
            "비고": 220
        }
        
        # 문자열 텍스트 정렬 가이드용 리스트업
        raw_text_cols = ["거래처", "담당부서", "담당자", "재단구분", "품목코드", "점착제코드", "비고"]
        
        clean_and_safe_display(
            q[raw_cols].sort_values(by="날짜", ascending=False),
            pinned_cols=["날짜", "거래처", "품목코드"],
            text_cols=raw_text_cols,
            column_width_overrides=raw_width_overrides
        )
else:
    st.info("💡 왼쪽 사이드바에서 `data.xlsx` 엑셀 파일을 업로드하시면 시스템이 활성화됩니다.")
