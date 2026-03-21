# ══════════════════════════════════════════════════════════
# TAB 4 — [완전 개선] 매출 하락 분석 (AI 우선순위, 수직 레이아웃, 그래프)
# ══════════════════════════════════════════════════════════
with tab4:
    import plotly.express as px
    import plotly.graph_objects as go
    
    st.subheader("📉 AI 기반 매출 하락 업체 분석")
    
    # 산출 방식 설명
    with st.expander("ℹ️ 전반부/후반부 매출 산출 방식 및 AI 분석 로직", expanded=False):
        st.markdown("""
### **📊 매출 산출 방식**

**기간 분할:**
- 전체 분석 기간을 시간 순서대로 정렬하여 중간 지점에서 분할
- 전반부와 후반부 각각의 월평균 매출을 계산

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

**특징:**
- 단순 감소액이 아닌 **종합적 위험도** 기준 정렬
- 영업적으로 시급한 대응이 필요한 업체 우선 노출
- 모든 수치는 **천단위 콤마(,)** 표시
        """)

    if q.empty or "날짜" not in q.columns or "금액(원)" not in q.columns or "거래처" not in q.columns:
        st.warning("분석에 필요한 데이터(날짜, 금액, 거래처)가 부족합니다.")
    else:
        # 월별 데이터 준비
        decline_df = q.copy()
        decline_df["월"] = decline_df["날짜"].dt.to_period("M").astype(str)
        decline_df = decline_df[decline_df["월"].notna()]
        
        all_months = sorted(decline_df["월"].unique())
        if len(all_months) < 2:
            st.info("분석 기간이 너무 짧습니다. (최소 2개월 필요)")
        else:
            # 전반부/후반부 분할
            mid_idx = len(all_months) // 2
            first_half = all_months[:mid_idx] if mid_idx > 0 else [all_months[0]]
            last_half = all_months[mid_idx:] if mid_idx < len(all_months) else [all_months[-1]]
            
            st.info(f"📅 **분석 기준:** 전반부 {first_half[0]}~{first_half[-1]} ({len(first_half)}개월) vs 후반부 {last_half[0]}~{last_half[-1]} ({len(last_half)}개월)")
            
            # 거래처별 월별 매출 집계
            monthly_sales = decline_df.groupby(["거래처","월"], dropna=False)["금액(원)"].sum().reset_index()
            
            # AI 우선순위 점수 계산 함수
            def calculate_ai_priority_score(cust, cust_monthly, all_cust_data):
                """AI 기반 우선순위 점수 계산 (100점 만점)"""
                
                # 기본 통계
                first_data = cust_monthly[cust_monthly["월"].isin(first_half)]["금액(원)"]
                last_data = cust_monthly[cust_monthly["월"].isin(last_half)]["금액(원)"]
                
                avg_first = float(first_data.mean()) if len(first_data) > 0 else 0.0
                avg_last = float(last_data.mean()) if len(last_data) > 0 else 0.0
                total_sales = float(cust_monthly["금액(원)"].sum())
                
                # 1. 매출 규모 점수 (30점)
                max_sales = all_cust_data.groupby("거래처")["금액(원)"].sum().max()
                sales_score = (total_sales / max_sales) * 30 if max_sales > 0 else 0
                
                # 2. 하락 심각도 점수 (25점)
                decline_rate = ((avg_first - avg_last) / avg_first) if avg_first > 0 else 0
                decline_score = min(max(0, decline_rate) * 25, 25)
                
                # 3. 품목 다양성 감소 점수 (20점)
                if "품목코드" in all_cust_data.columns:
                    cust_detail = all_cust_data[all_cust_data["거래처"] == cust]
                    first_products = cust_detail[cust_detail["월"].isin(first_half)]["품목코드"].nunique()
                    last_products = cust_detail[cust_detail["월"].isin(last_half)]["품목코드"].nunique()
                    product_decline = max(0, first_products - last_products)
                    diversity_score = min((product_decline / max(1, first_products)) * 20, 20)
                else:
                    diversity_score = 0
                
                # 4. 변동성 증가 점수 (15점)
                monthly_amounts = cust_monthly.set_index("월")["금액(원)"]
                if len(monthly_amounts) > 1 and monthly_amounts.mean() > 0:
                    cv = (monthly_amounts.std() / monthly_amounts.mean())
                    volatility_score = min(cv * 15, 15)
                else:
                    volatility_score = 0
                
                # 5. 최근 추세 점수 (10점)
                recent_months = all_months[-3:] if len(all_months) >= 3 else all_months
                recent_data = cust_monthly[cust_monthly["월"].isin(recent_months)]
                if len(recent_data) >= 2:
                    recent_trend = recent_data["금액(원)"].pct_change().mean()
                    trend_score = max(0, -recent_trend * 10) if not pd.isna(recent_trend) else 0
                    trend_score = min(trend_score, 10)
                else:
                    trend_score = 0
                
                total_score = sales_score + decline_score + diversity_score + volatility_score + trend_score
                
                # 분석 내역 자동 생성
                analysis_text = f"""
**매출규모**: {sales_score:.1f}점 (총 {total_sales:,.0f}원)
**하락심각도**: {decline_score:.1f}점 (하락률 {decline_rate*100:.1f}%)
**품목다양성**: {diversity_score:.1f}점
**변동성**: {volatility_score:.1f}점  
**최근추세**: {trend_score:.1f}점
**종합평가**: {"⚠️ 긴급" if total_score > 70 else "🔍 주의" if total_score > 50 else "📋 모니터링"}
                """.strip()
                
                return {
                    "거래처": cust,
                    "전반부_평균매출": round(avg_first, 0),
                    "후반부_평균매출": round(avg_last, 0),
                    "매출_감소액": round(avg_last - avg_first, 0),
                    "전체_매출액": round(total_sales, 0),
                    "AI_우선순위점수": round(total_score, 1),
                    "분석_내역": analysis_text
                }
            
            # 거래처별 분석 실행
            st.markdown("#### 🤖 AI 분석 진행 중...")
            progress_bar = st.progress(0)
            
            analysis_results = []
            customers = monthly_sales["거래처"].unique()
            
            for idx, cust in enumerate(customers):
                cust_monthly = monthly_sales[monthly_sales["거래처"] == cust]
                result = calculate_ai_priority_score(cust, cust_monthly, decline_df)
                
                # 실제 하락한 업체만 포함
                if result["매출_감소액"] < 0:
                    analysis_results.append(result)
                
                progress_bar.progress((idx + 1) / len(customers))
            
            progress_bar.empty()
            
            if not analysis_results:
                st.success("✅ 설정 기간 동안 매출이 하락한 업체가 없습니다.")
            else:
                # AI 점수 기준 정렬 및 상위 35% 선정
                results_df = pd.DataFrame(analysis_results)
                results_df = results_df.sort_values("AI_우선순위점수", ascending=False)
                
                top_count = max(1, int(np.ceil(len(results_df) * 0.35)))
                top_priority = results_df.head(top_count).copy()
                top_priority["순위"] = range(1, len(top_priority) + 1)
                
                # ═══════════════════════════════════════════════════════════
                # [수직 레이아웃 1] 우선 대응 업체 리스트
                # ═══════════════════════════════════════════════════════════
                st.markdown(f"### 🎯 우선 대응 필요 업체 (상위 {len(top_priority)}개)")
                st.caption("💡 AI가 분석한 종합 위험도 점수 기준으로 정렬되었습니다. 분석 내역을 참고하여 영업 전략을 수립하세요.")
                
                # 표시할 컬럼 정의
                display_cols = ["순위", "거래처", "AI_우선순위점수", "전체_매출액", 
                               "전반부_평균매출", "후반부_평균매출", "매출_감소액", "분석_내역"]
                
                # 편집 가능한 테이블 (천단위 콤마 적용)
                edited_priority = st.data_editor(
                    top_priority[display_cols],
                    use_container_width=True,
                    column_config={
                        "AI_우선순위점수": st.column_config.NumberColumn(
                            "AI 우선순위점수", 
                            format="%.1f점",
                            help="100점 만점 (높을수록 긴급)"
                        ),
                        "전체_매출액": st.column_config.NumberColumn(
                            "전체 매출액", 
                            format="%d원"
                        ),
                        "전반부_평균매출": st.column_config.NumberColumn(
                            "전반부 평균매출", 
                            format="%d원"
                        ),
                        "후반부_평균매출": st.column_config.NumberColumn(
                            "후반부 평균매출", 
                            format="%d원"
                        ),
                        "매출_감소액": st.column_config.NumberColumn(
                            "매출 감소액", 
                            format="%d원"
                        ),
                        "분석_내역": st.column_config.TextColumn(
                            "AI 분석 내역",
                            width="large",
                            help="AI가 분석한 상세 내역"
                        )
                    },
                    num_rows="dynamic",
                    key="priority_customers_editor"
                )
                
                # CSV 다운로드
                csv_priority = edited_priority.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                st.download_button(
                    "📥 우선 대응 업체 분석 결과 CSV 다운로드",
                    data=csv_priority,
                    file_name="AI분석_우선대응업체.csv",
                    mime="text/csv"
                )
                
                st.markdown("---")
                
                # ═══════════════════════════════════════════════════════════
                # [수직 레이아웃 2] 업체별 상세 분석
                # ═══════════════════════════════════════════════════════════
                st.markdown("### 🔍 업체별 상세 품목 분석")
                
                selected_customer = st.selectbox(
                    "분석할 업체를 선택하세요",
                    options=["선택하세요"] + [
                        f"{row['거래처']} (점수: {row['AI_우선순위점수']:.1f}점)" 
                        for _, row in top_priority.iterrows()
                    ],
                    key="customer_detail_select"
                )
                
                if selected_customer != "선택하세요":
                    # 선택된 업체명 추출
                    selected_cust_name = selected_customer.split(" (점수:")[0]
                    
                    st.markdown(f"#### 📋 [{selected_cust_name}] 품목별 월간 매출 분석")
                    
                    # 해당 업체 데이터 필터링
                    customer_data = decline_df[decline_df["거래처"] == selected_cust_name]
                    
                    if "품목코드" in customer_data.columns and not customer_data.empty:
                        # 품목별 월별 피벗 테이블
                        customer_pivot = customer_data.pivot_table(
                            index="품목코드", columns="월",
                            values="금액(원)", aggfunc="sum", fill_value=0)
                        
                        # 품목별 하락 기여도 계산 및 정렬
                        product_declines = []
                        for prod in customer_pivot.index:
                            prod_data = customer_pivot.loc[prod]
                            first_avg = prod_data[prod_data.index.isin(first_half)].mean()
                            last_avg = prod_data[prod_data.index.isin(last_half)].mean()
                            decline_amount = last_avg - first_avg
                            product_declines.append(decline_amount)
                        
                        customer_pivot["_decline"] = product_declines
                        customer_pivot = customer_pivot.sort_values("_decline")  # 하락폭 큰 순
                        customer_pivot = customer_pivot.drop(columns=["_decline"])
                        
                        # 합계 컬럼 추가
                        month_cols = list(customer_pivot.columns)
                        customer_pivot["합계"] = customer_pivot[month_cols].sum(axis=1)
                        customer_pivot_reset = customer_pivot.reset_index()
                        
                        st.caption(f"💡 {selected_cust_name}의 품목별 월간 매출 (하락 기여도 큰 순)")
                        
                        # 편집 가능한 품목 분석 테이블
                        edited_customer_pivot = st.data_editor(
                            customer_pivot_reset,
                            use_container_width=True,
                            num_rows="dynamic",
                            key="customer_product_editor"
                        )
                        
                        # 품목별 분석 CSV 다운로드
                        csv_customer = edited_customer_pivot.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                        st.download_button(
                            "📥 품목별 월간 매출 CSV 다운로드",
                            data=csv_customer,
                            file_name=f"{selected_cust_name}_품목별_월간매출.csv",
                            mime="text/csv"
                        )
                        
                        st.markdown("---")
                        
                        # ═══════════════════════════════════════════════════════════
                        # [수직 레이아웃 3] 품목별 판매 추이 그래프
                        # ═══════════════════════════════════════════════════════════
                        st.markdown(f"### 📈 [{selected_cust_name}] 품목별 판매 추이 시각화")
                        st.caption("💡 어떤 품목이 매출 하락을 주도하는지 시각적으로 확인하세요.")
                        
                        # 그래프용 데이터 준비
                        customer_monthly = customer_data.groupby(["월","품목코드"], dropna=False)["금액(원)"].sum().reset_index()
                        customer_monthly = customer_monthly[customer_monthly["품목코드"].astype(str).str.lower() != "nan"]
                        
                        # 2개 그래프를 나란히 배치
                        col_graph1, col_graph2 = st.columns(2)
                        
                        with col_graph1:
                            # 라인 차트 (품목별 추세)
                            fig_line = px.line(
                                customer_monthly,
                                x="월", y="금액(원)", color="품목코드",
                                title=f"{selected_cust_name} - 품목별 매출 추세",
                                labels={"금액(원)":"매출액(원)", "월":""},
                                markers=True
                            )
                            fig_line.update_layout(
                                xaxis_tickangle=-45,
                                height=400,
                                yaxis_tickformat=",",
                                legend=dict(orientation="h", yanchor="bottom", y=-0.3)
                            )
                            st.plotly_chart(fig_line, use_container_width=True)
                        
                        with col_graph2:
                            # 누적 영역 차트 (전체 구성 파악)
                            fig_area = px.area(
                                customer_monthly,
                                x="월", y="금액(원)", color="품목코드",
                                title=f"{selected_cust_name} - 품목별 매출 구성",
                                labels={"금액(원)":"매출액(원)", "월":""}
                            )
                            fig_area.update_layout(
                                xaxis_tickangle=-45,
                                height=400,
                                yaxis_tickformat=",",
                                legend=dict(orientation="h", yanchor="bottom", y=-0.3)
                            )
                            st.plotly_chart(fig_area, use_container_width=True)
                        
                        # 전반부 vs 후반부 비교 막대 그래프
                        st.markdown("##### 📊 품목별 전반부 vs 후반부 평균 매출 비교")
                        
                        # 비교 데이터 생성
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
                            comp_df = comp_df.sort_values("변화액")  # 하락폭 큰 순
                            
                            # 막대 그래프
                            fig_bar = go.Figure()
                            fig_bar.add_trace(go.Bar(
                                x=comp_df["품목코드"],
                                y=comp_df["전반부_평균"],
                                name="전반부 평균",
                                marker_color="#3498db",
                                text=[f"{v:,.0f}" for v in comp_df["전반부_평균"]],
                                textposition="outside"
                            ))
                            fig_bar.add_trace(go.Bar(
                                x=comp_df["품목코드"],
                                y=comp_df["후반부_평균"],
                                name="후반부 평균",
                                marker_color="#e74c3c",
                                text=[f"{v:,.0f}" for v in comp_df["후반부_평균"]],
                                textposition="outside"
                            ))
                            
                            fig_bar.update_layout(
                                title=f"{selected_cust_name} - 품목별 전반부/후반부 평균 매출 비교",
                                xaxis_title="품목코드",
                                yaxis_title="평균 매출액(원)",
                                barmode="group",
                                height=400,
                                yaxis_tickformat=","
                            )
                            
                            st.plotly_chart(fig_bar, use_container_width=True)
                            
                            # 수치 비교 테이블
                            st.markdown("##### 📋 품목별 변화율 상세")
                            st.dataframe(
                                auto_fmt(comp_df),
                                use_container_width=True
                            )
                    else:
                        st.warning("해당 업체의 품목별 데이터가 없습니다.")
