import os
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="출고 이력 검색", layout="wide")


# ══════════════════════════════════════════════════════════
# 공통 유틸
# ══════════════════════════════════════════════════════════
def clean_and_safe_display(
    df,
    height=None,
    key=None,
    editable=False,
    pinned_cols=None,
    text_cols=None,
    disabled_cols=None,
):
    if df is None:
        df = pd.DataFrame()

    display_df = df.copy().reset_index(drop=True)
    display_df.columns = [str(c) for c in display_df.columns]

    pinned_cols = [str(c) for c in (pinned_cols or [])]
    text_cols = set(str(c) for c in (text_cols or []))
    disabled_cols = disabled_cols if disabled_cols is not None else False

    if display_df.empty:
        if editable and key:
            return st.data_editor(
                display_df,
                width="stretch",
                hide_index=True,
                key=key,
                disabled=disabled_cols,
                placeholder="",
            )
        st.dataframe(display_df, width="stretch", hide_index=True, placeholder="")
        return None

    for col in display_df.columns:
        s = display_df[col]

        if pd.api.types.is_datetime64_any_dtype(s):
            display_df[col] = pd.to_datetime(s, errors="coerce").dt.strftime("%Y-%m-%d")
            display_df[col] = display_df[col].fillna("")
            continue

        if pd.api.types.is_numeric_dtype(s) and col not in text_cols:
            display_df[col] = pd.to_numeric(s, errors="coerce")
            display_df[col] = display_df[col].replace([np.inf, -np.inf], np.nan).fillna(0)
            continue

        display_df[col] = s.astype(str)
        display_df[col] = display_df[col].replace(
            ["nan", "NaN", "None", "<NA>", "NaT"], ""
        )
        display_df[col] = display_df[col].fillna("")

    column_config = {}
    fixed_text_like_cols = {
        "거래처", "품목코드", "품목명(공식)", "점착제코드", "점착제명",
        "최근날짜", "가로폭이력", "분석_내역", "월", "구분", "월표기", "연도"
    }

    for col in display_df.columns:
        pinned = col in pinned_cols

        if col in text_cols or col in fixed_text_like_cols:
            column_config[col] = st.column_config.TextColumn(
                col,
                width="large" if col in ["품목명(공식)", "가로폭이력", "분석_내역"] else "medium",
                pinned=pinned,
            )
            continue

        if pd.api.types.is_numeric_dtype(display_df[col]):
            if any(k in col for k in ["하락률", "증감률", "비율", "변화율"]):
                column_config[col] = st.column_config.NumberColumn(
                    col, format="%.1f", pinned=pinned
                )
            elif any(k in col for k in ["M2", "수량", "판매량", "총량"]):
                column_config[col] = st.column_config.NumberColumn(
                    col, format="%,.1f", pinned=pinned
                )
            elif any(k in col for k in ["점수", "AI", "우선순위", "통계", "종합"]):
                column_config[col] = st.column_config.NumberColumn(
                    col, format="%.1f", pinned=pinned
                )
            else:
                column_config[col] = st.column_config.NumberColumn(
                    col, format="%,d", pinned=pinned
                )
        else:
            column_config[col] = st.column_config.TextColumn(
                col, width="medium", pinned=pinned
            )

    if editable and key:
        safe_editor_df = display_df.copy()

        for col in safe_editor_df.columns:
            if col in text_cols or col in fixed_text_like_cols:
                safe_editor_df[col] = safe_editor_df[col].astype(str).fillna("")
            else:
                tmp = pd.to_numeric(safe_editor_df[col], errors="coerce")
                if tmp.notna().sum() > 0:
                    safe_editor_df[col] = tmp.fillna(0)
                else:
                    safe_editor_df[col] = safe_editor_df[col].astype(str).fillna("")

        return st.data_editor(
            safe_editor_df,
            column_config=column_config,
            width="stretch",
            height=height if height else "auto",
            num_rows="fixed",
            key=key,
            hide_index=True,
            disabled=disabled_cols,
            placeholder="",
        )

    st.dataframe(
        display_df,
        column_config=column_config,
        width="stretch",
        height=height if height else "auto",
        hide_index=True,
        placeholder="",
    )
    return None


def sorted_unique(series):
    if series is None:
        return []
    s = pd.Series(series).astype(str).str.strip()
    s = s[~s.isin(["", "0", "0.0", "nan", "NaN", "None", "<NA>"])]
    return sorted(s.unique())


def to_yymm(series):
    dt = pd.to_datetime(series, errors="coerce")
    return dt.dt.strftime("%y%m")


def add_year_month_axis(fig, x_dates):
    """
    월은 ticktext로, 연도는 annotation으로 하단에 표시.
    사용자 요청의 '상단 월 / 하단 연도' 느낌을 최대한 유사하게 구현.
    """
    dt = pd.to_datetime(pd.Series(x_dates), errors="coerce").dropna().sort_values().unique()
    if len(dt) == 0:
        return fig

    dt = pd.to_datetime(dt)
    tickvals = list(dt)
    ticktext = [pd.Timestamp(d).strftime("%m") for d in dt]

    fig.update_xaxes(
        tickmode="array",
        tickvals=tickvals,
        ticktext=ticktext,
        tickangle=0,
        title="",
        showgrid=True,
        zeroline=False
    )

    # 연도 annotation
    years = sorted(pd.Series(dt).dt.year.unique())
    for y in years:
        year_dates = [d for d in dt if pd.Timestamp(d).year == y]
        if not year_dates:
            continue
        mid = year_dates[len(year_dates) // 2]
        fig.add_annotation(
            x=mid,
            y=-0.20,
            xref="x",
            yref="paper",
            text=str(y),
            showarrow=False,
            font=dict(size=12, color="black"),
        )

        # 연도 시작선
        fig.add_vline(
            x=year_dates[0],
            line_width=1,
            line_dash="dot",
            line_color="rgba(120,120,120,0.5)"
        )

    fig.update_layout(margin=dict(b=90))
    return fig


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

    rec = read_sheet("출고기록")
    alias = read_sheet("별칭맵핑")
    prod = read_sheet("품목마스터", required=False)
    adh = read_sheet("점착제마스터", required=False)
    cust = read_sheet("거래처마스터", required=False)

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

    def normalize(df, col):
        if col in df.columns:
            s = df[col].astype(str).str.strip()
            df[col] = s.replace({
                "": pd.NA,
                "0": pd.NA,
                "0.0": pd.NA,
                "nan": pd.NA,
                "NaN": pd.NA,
                "None": pd.NA,
                "<NA>": pd.NA,
            })

    for c in ["거래처", "품목코드", "점착제코드", "품목명_고객표현", "점착제_고객표현"]:
        if c in rec.columns:
            rec[c] = rec[c].astype(str).str.strip()

    for c in ["품목코드", "점착제코드"]:
        normalize(rec, c)
    normalize(rec, "거래처")

    def map_alias(series, alias_df, typ):
        if alias_df is None or alias_df.empty or series is None:
            return pd.Series([None] * len(series))

        df = alias_df.copy()
        if not {"유형", "별칭", "공식코드"} <= set(df.columns):
            return pd.Series([None] * len(series))

        df = df[df["유형"].astype(str) == typ].dropna(subset=["별칭", "공식코드"])
        al = df["별칭"].astype(str).str.strip().tolist()
        cl = df["공식코드"].astype(str).str.strip().tolist()

        out = []
        for v in series.astype(str).fillna("").str.strip():
            m = next((c for a, c in zip(al, cl) if v == a), None)
            if m is None:
                m = next((c for a, c in zip(al, cl) if a and a in v), None)
            out.append(m)
        return pd.Series(out)

    for col, typ in [("품목코드", "품목"), ("점착제코드", "점착제")]:
        src = "품목명_고객표현" if typ == "품목" else "점착제_고객표현"
        if col not in rec.columns or rec[col].isna().any():
            if col not in rec.columns:
                rec[col] = pd.NA
            if src in rec.columns:
                rec[col] = rec[col].fillna(map_alias(rec[src], alias, typ))

    for c in ["품목코드", "점착제코드"]:
        normalize(rec, c)

    if not prod.empty and {"품목코드", "품목명(공식)"} <= set(prod.columns):
        prod_cols = ["품목코드", "품목명(공식)"]
        if "품목비고" in prod.columns:
            prod_cols.append("품목비고")
        rec = rec.merge(prod[prod_cols].drop_duplicates(), on="품목코드", how="left")

    if not adh.empty and {"점착제코드", "점착제명"} <= set(adh.columns):
        rec = rec.merge(
            adh[["점착제코드", "점착제명"]].drop_duplicates(),
            on="점착제코드",
            how="left",
        )

    tmp = rec.copy()
    tmp["_cs"] = tmp["거래처"].astype(str).fillna("") if "거래처" in tmp.columns else ""
    tmp["_ps"] = tmp["품목코드"].astype(str).fillna("") if "품목코드" in tmp.columns else ""
    sc = ["_cs", "_ps"] + (["날짜"] if "날짜" in tmp.columns else [])
    tmp = tmp.sort_values(sc, kind="mergesort")

    if "단가(원/M2)" in tmp.columns and {"거래처", "품목코드"} <= set(tmp.columns):
        _t = tmp.dropna(subset=["단가(원/M2)"]).groupby(
            ["거래처", "품목코드"], as_index=False, dropna=False
        ).tail(1)
        _ec = ["거래처", "품목코드", "단가(원/M2)"] + (["날짜"] if "날짜" in _t.columns else [])
        _rm = {"단가(원/M2)": "최근단가"}
        if "날짜" in _t.columns:
            _rm["날짜"] = "최근날짜"
        lp = _t[_ec].rename(columns=_rm)
        if "최근날짜" in lp.columns:
            lp["최근날짜"] = pd.to_datetime(lp["최근날짜"], errors="coerce").dt.strftime("%Y-%m-%d")
        rec = rec.merge(lp, on=["거래처", "품목코드"], how="left")

    def join_unique(s):
        vals = []
        for x in pd.unique(s.dropna()):
            try:
                xf = float(x)
                vals.append(str(int(xf)) if xf.is_integer() else str(xf))
            except Exception:
                vals.append(str(x))
        return ", ".join(vals)

    if "가로폭(mm)" in rec.columns and {"거래처", "품목코드"} <= set(rec.columns):
        wh = (
            rec.groupby(["거래처", "품목코드"], dropna=False)["가로폭(mm)"]
            .apply(join_unique)
            .reset_index()
            .rename(columns={"가로폭(mm)": "가로폭이력"})
        )
        rec = rec.merge(wh, on=["거래처", "품목코드"], how="left")
    else:
        rec["가로폭이력"] = pd.NA

    return rec, alias, prod, adh, cust


DEFAULT_FILE = "data.xlsx"

st.title("출고 이력 검색(거래처/품목/가로폭/점착제)")
uploaded = st.file_uploader("📂 다른 파일 업로드 (미업로드 시 기본 데이터 자동 로드)", type=["xlsx"])

if uploaded:
    file_bytes = uploaded.getvalue()
    st.success("✅ 업로드 파일 사용")
elif os.path.exists(DEFAULT_FILE):
    with open(DEFAULT_FILE, "rb") as f:
        file_bytes = f.read()
    st.info(f"📌 기본 데이터({DEFAULT_FILE}) 자동 로드")
else:
    st.info("GitHub 레포에 data.xlsx를 추가하거나 파일을 업로드하세요.")
    st.stop()

rec, alias, prod, adh, cust = load_excel(file_bytes)

st.sidebar.header("검색 필터")
sel_cust = st.sidebar.multiselect("거래처", sorted_unique(rec.get("거래처", pd.Series())))
sel_prod = st.sidebar.multiselect("품목코드", sorted_unique(rec.get("품목코드", pd.Series())))
sel_adh = st.sidebar.multiselect("점착제코드", sorted_unique(rec.get("점착제코드", pd.Series())))

date_min = pd.to_datetime(rec["날짜"].min()) if "날짜" in rec.columns else None
date_max = pd.to_datetime(rec["날짜"].max()) if "날짜" in rec.columns else None

if date_min is not None and pd.notna(date_min) and date_max is not None and pd.notna(date_max):
    sdate, edate = st.sidebar.date_input("기간", [date_min.date(), date_max.date()])
else:
    sdate = edate = None

st.sidebar.markdown("---")
st.sidebar.caption("💡 견적 레퍼런스: 품목코드·점착제코드·기간 필터 위주로 사용하세요.")

q = rec.copy()
if sel_cust and "거래처" in q.columns:
    q = q[q["거래처"].astype(str).isin(sel_cust)]
if sel_prod and "품목코드" in q.columns:
    q = q[q["품목코드"].astype(str).isin(sel_prod)]
if sel_adh and "점착제코드" in q.columns:
    q = q[q["점착제코드"].astype(str).isin(sel_adh)]
if sdate and edate and "날짜" in q.columns:
    q = q[(q["날짜"] >= pd.to_datetime(sdate)) & (q["날짜"] <= pd.to_datetime(edate))]

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "거래처별 검색",
    "품목별 검색",
    "🏷️ 견적 레퍼런스",
    "📉 매출 하락 분석",
    "원자료",
])


# ══════════════════════════════════════════════════════════
# TAB 1
# ══════════════════════════════════════════════════════════
with tab1:
    st.subheader("거래처별 → 품목별 출고 이력/최근 단가/가로폭")

    if q.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
    else:
        q1 = q.copy()
        for c in ["거래처", "품목코드", "점착제코드"]:
            if c in q1.columns:
                q1[c] = q1[c].astype(str)

        cols = ["거래처", "품목코드", "품목명(공식)", "점착제코드", "점착제명",
                "가로폭이력", "최근날짜", "최근단가"]
        uc = [c for c in cols if c in q1.columns]

        g = (
            q1.groupby(uc, dropna=False)
            .agg(
                출고횟수=("수량(M2)", "count"),
                총량_M2=("수량(M2)", "sum"),
                매출액=("금액(원)", "sum"),
            )
            .reset_index()
        )

        g["가중평균단가"] = np.where(
            g["총량_M2"] > 0,
            (g["매출액"] / g["총량_M2"]).round(0),
            0,
        )

        sc = [c for c in ["거래처", "품목코드"] if c in g.columns]

        clean_and_safe_display(
            g.sort_values(sc) if sc else g,
            pinned_cols=["거래처", "품목코드"],
            text_cols=["거래처", "품목코드", "품목명(공식)", "점착제코드", "점착제명", "가로폭이력", "최근날짜"],
        )


# ══════════════════════════════════════════════════════════
# TAB 2
# ══════════════════════════════════════════════════════════
with tab2:
    st.subheader("품목별 → 거래처별 출고 이력/총량/단가/매출")

    if q.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
    else:
        q2 = q.copy()
        for c in ["거래처", "품목코드"]:
            if c in q2.columns:
                q2[c] = q2[c].astype(str)

        cols = ["품목코드", "품목명(공식)", "거래처", "최근날짜", "최근단가"]
        uc = [c for c in cols if c in q2.columns]

        g2 = (
            q2.groupby(uc, dropna=False)
            .agg(
                출고횟수=("수량(M2)", "count"),
                총량_M2=("수량(M2)", "sum"),
                매출액=("금액(원)", "sum"),
            )
            .reset_index()
        )

        g2["가중평균단가"] = np.where(
            g2["총량_M2"] > 0,
            (g2["매출액"] / g2["총량_M2"]).round(0),
            0,
        )

        sc = [c for c in ["품목코드", "거래처"] if c in g2.columns]

        clean_and_safe_display(
            g2.sort_values(sc) if sc else g2,
            pinned_cols=["품목코드", "거래처"],
            text_cols=["품목코드", "품목명(공식)", "거래처", "최근날짜"],
        )


# ══════════════════════════════════════════════════════════
# TAB 3
# ══════════════════════════════════════════════════════════
with tab3:
    st.subheader("🏷️ 견적 레퍼런스 — 기준 견적가 & 판매 동향")
    st.caption("단가 0 및 샘플 품목 자동 제외")

    q_ref = q.copy()

    if "단가(원/M2)" in q_ref.columns:
        q_ref = q_ref[q_ref["단가(원/M2)"].notna() & (q_ref["단가(원/M2)"] > 0)]

    for col in ["품목코드", "품목명(공식)"]:
        if col in q_ref.columns:
            q_ref = q_ref[~q_ref[col].astype(str).str.contains("샘플", case=False, na=False)]

    if q_ref.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
    else:
        grp_base = ["품목코드", "품목명(공식)", "점착제코드", "점착제명"]
        GC = [c for c in grp_base if c in q_ref.columns]

        if "날짜" in q_ref.columns:
            vd = q_ref["날짜"].dropna()
            if len(vd):
                d0, d1 = vd.min(), vd.max()
                n_months = max(1, (d1.year - d0.year) * 12 + d1.month - d0.month + 1)
            else:
                n_months = 1
        else:
            n_months = 1

        overview = q_ref.groupby(GC, dropna=False).agg(
            최저단가=("단가(원/M2)", "min"),
            최고단가=("단가(원/M2)", "max"),
            거래처수=("거래처", "nunique"),
            총출고횟수=("수량(M2)", "count"),
            총량_M2=("수량(M2)", "sum"),
            총매출액=("금액(원)", "sum"),
        ).reset_index()

        overview["월평균판매량_M2"] = (overview["총량_M2"] / n_months).round(1)
        overview["월평균판매액_원"] = (overview["총매출액"] / n_months).round(0)

        clean_and_safe_display(
            overview,
            pinned_cols=["품목코드"],
            text_cols=["품목코드", "품목명(공식)", "점착제코드", "점착제명"],
        )


# ══════════════════════════════════════════════════════════
# TAB 4
# ══════════════════════════════════════════════════════════
with tab4:
    st.subheader("📉 AI 기반 매출 하락 업체 분석")
    st.caption("감소규모 60% + 통계추세 20% + AI분석 20%")

    if q.empty or "날짜" not in q.columns or "금액(원)" not in q.columns or "거래처" not in q.columns:
        st.warning("분석에 필요한 데이터가 부족합니다.")
    else:
        decline_df = q.copy()
        decline_df["월"] = pd.to_datetime(decline_df["날짜"], errors="coerce").dt.strftime("%Y-%m")
        decline_df = decline_df[decline_df["월"].notna() & (decline_df["월"] != "")]

        all_months = sorted(decline_df["월"].unique().tolist())

        if len(all_months) < 2:
            st.info("분석 기간이 너무 짧습니다. (최소 2개월 필요)")
        else:
            mid_idx = len(all_months) // 2
            first_half = all_months[:mid_idx] if mid_idx > 0 else [all_months[0]]
            last_half = all_months[mid_idx:] if mid_idx < len(all_months) else [all_months[-1]]

            monthly_sales = decline_df.groupby(["거래처", "월"], dropna=False)["금액(원)"].sum().reset_index()

            all_decline_amounts = []
            for cust in monthly_sales["거래처"].unique():
                cust_monthly = monthly_sales[monthly_sales["거래처"] == cust]
                first_data = cust_monthly[cust_monthly["월"].isin(first_half)]["금액(원)"]
                last_data = cust_monthly[cust_monthly["월"].isin(last_half)]["금액(원)"]
                avg_first = float(first_data.mean()) if len(first_data) > 0 else 0.0
                avg_last = float(last_data.mean()) if len(last_data) > 0 else 0.0
                decline_amount = max(0.0, avg_first - avg_last)
                all_decline_amounts.append(decline_amount)

            max_decline_amount = max(all_decline_amounts) if all_decline_amounts else 0.0

            def calculate_priority_score(cust, cust_monthly, all_cust_data):
                cust_monthly = cust_monthly.sort_values("월").copy()

                first_data = cust_monthly[cust_monthly["월"].isin(first_half)]["금액(원)"]
                last_data = cust_monthly[cust_monthly["월"].isin(last_half)]["금액(원)"]

                avg_first = float(first_data.mean()) if len(first_data) > 0 else 0.0
                avg_last = float(last_data.mean()) if len(last_data) > 0 else 0.0
                total_sales = float(cust_monthly["금액(원)"].sum())

                decline_amount = max(0.0, avg_first - avg_last)
                decline_rate = (decline_amount / avg_first) if avg_first > 0 else 0.0

                amount_score = (decline_amount / max_decline_amount) * 60 if max_decline_amount > 0 else 0.0
                amount_score = min(amount_score, 60.0)

                monthly_vals = cust_monthly.groupby("월", as_index=False)["금액(원)"].sum().sort_values("월")
                monthly_vals["t"] = range(len(monthly_vals))

                if len(monthly_vals) >= 2:
                    x = monthly_vals["t"].values.astype(float)
                    y = monthly_vals["금액(원)"].values.astype(float)
                    slope = np.polyfit(x, y, 1)[0] if len(x) >= 2 else 0.0
                    slope_score = max(0.0, -slope)

                    cv = (np.std(y) / np.mean(y)) if np.mean(y) > 0 else 0.0
                    recent_changes = monthly_vals["금액(원)"].pct_change().dropna()
                    down_streak = int((recent_changes.tail(2) < 0).sum()) if len(recent_changes) > 0 else 0

                    stat_score = 0.0
                    if slope_score > 0:
                        stat_score += 12.0
                    stat_score += min(cv * 4, 4.0)
                    stat_score += min(down_streak * 2, 4.0)
                    stat_score = min(stat_score, 20.0)
                else:
                    slope = 0.0
                    cv = 0.0
                    stat_score = 0.0

                if "품목코드" in all_cust_data.columns:
                    cust_detail = all_cust_data[all_cust_data["거래처"] == cust].copy()
                    first_products = cust_detail[cust_detail["월"].isin(first_half)]["품목코드"].nunique()
                    last_products = cust_detail[cust_detail["월"].isin(last_half)]["품목코드"].nunique()
                    product_decline = max(0, first_products - last_products)
                    diversity_score = min((product_decline / max(1, first_products)) * 8, 8.0)
                else:
                    diversity_score = 0.0

                max_sales = all_cust_data.groupby("거래처")["금액(원)"].sum().max()
                scale_score = min((total_sales / max_sales) * 6, 6.0) if max_sales > 0 else 0.0

                recent_months = all_months[-3:] if len(all_months) >= 3 else all_months
                recent_data = cust_monthly[cust_monthly["월"].isin(recent_months)].sort_values("월")
                if len(recent_data) >= 2:
                    recent_trend = recent_data["금액(원)"].pct_change().mean()
                    trend_score = min(max(0.0, -recent_trend * 6), 6.0) if not pd.isna(recent_trend) else 0.0
                else:
                    trend_score = 0.0

                ai_score = min(diversity_score + scale_score + trend_score, 20.0)
                total_score = amount_score + stat_score + ai_score

                analysis_text = (
                    f"감소규모:{amount_score:.1f}/60 | "
                    f"통계추세:{stat_score:.1f}/20 | "
                    f"AI분석:{ai_score:.1f}/20 | "
                    f"감소액:{int(decline_amount):,}원 | "
                    f"하락률:{decline_rate*100:.1f}% | "
                    f"기울기:{int(round(slope,0)):,} | CV:{cv:.2f}"
                )

                return {
                    "거래처": str(cust),
                    "전반부_평균매출": int(round(avg_first, 0)),
                    "후반부_평균매출": int(round(avg_last, 0)),
                    "실제감소액": int(round(decline_amount, 0)),
                    "매출_감소액": int(round(avg_last - avg_first, 0)),
                    "하락률(%)": round(decline_rate * 100, 1) if avg_first > 0 else 0.0,
                    "전체_매출액": int(round(total_sales, 0)),
                    "감소규모점수": round(amount_score, 1),
                    "통계추세점수": round(stat_score, 1),
                    "AI분석점수": round(ai_score, 1),
                    "AI_우선순위점수": round(total_score, 1),
                    "분석_내역": analysis_text,
                }

            analysis_results = []
            for cust in monthly_sales["거래처"].unique():
                cust_monthly = monthly_sales[monthly_sales["거래처"] == cust]
                result = calculate_priority_score(cust, cust_monthly, decline_df)
                if result["실제감소액"] > 0:
                    analysis_results.append(result)

            if not analysis_results:
                st.success("✅ 설정 기간 동안 매출이 하락한 업체가 없습니다.")
            else:
                results_df = pd.DataFrame(analysis_results)
                results_df = results_df.sort_values(
                    ["AI_우선순위점수", "실제감소액"],
                    ascending=[False, False]
                ).reset_index(drop=True)

                top_count = max(1, int(np.ceil(len(results_df) * 0.35)))
                top_priority = results_df.head(top_count).copy()
                top_priority["순위"] = range(1, len(top_priority) + 1)

                st.markdown(f"### 🎯 우선 대응 필요 업체 (상위 {len(top_priority)}개)")

                display_cols = [
                    "순위", "거래처", "AI_우선순위점수", "감소규모점수", "통계추세점수", "AI분석점수",
                    "전체_매출액", "전반부_평균매출", "후반부_평균매출", "실제감소액", "하락률(%)", "분석_내역"
                ]

                edited_priority = clean_and_safe_display(
                    top_priority[display_cols],
                    key="priority_customers_editor",
                    editable=True,
                    pinned_cols=["순위", "거래처"],
                    text_cols=["거래처", "분석_내역"],
                    disabled_cols=display_cols,
                )

                if edited_priority is not None:
                    csv_priority = edited_priority.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                    st.download_button(
                        "📥 우선 대응 업체 분석 결과 CSV 다운로드",
                        data=csv_priority,
                        file_name="AI분석_우선대응업체.csv",
                        mime="text/csv",
                    )

                st.markdown("---")
                st.markdown("### 🔍 업체별 상세 분석")

                selected_customer = st.selectbox(
                    "분석할 업체를 선택하세요",
                    options=["선택하세요"] + [
                        f"{row['거래처']} (점수: {row['AI_우선순위점수']:.1f}점, 감소액: {row['실제감소액']:,}원)"
                        for _, row in top_priority.iterrows()
                    ],
                    key="customer_detail_select",
                )

                if selected_customer != "선택하세요":
                    selected_cust_name = selected_customer.split(" (점수:")[0]
                    customer_data = decline_df[decline_df["거래처"] == selected_cust_name].copy()

                    if customer_data.empty:
                        st.warning("해당 업체 데이터가 없습니다.")
                    else:
                        st.markdown(f"#### 📋 [{selected_cust_name}] 품목별 월간 매출 분석")

                        # 월별 업체 총매출
                        customer_total_monthly = (
                            customer_data.groupby("월", as_index=False)["금액(원)"]
                            .sum()
                            .sort_values("월")
                        )
                        customer_total_monthly["날짜축"] = pd.to_datetime(customer_total_monthly["월"] + "-01")
                        customer_total_monthly["월표기"] = customer_total_monthly["날짜축"].dt.strftime("%m")
                        customer_total_monthly["연도"] = customer_total_monthly["날짜축"].dt.strftime("%Y")

                        # 품목별 월간
                        product_monthly = (
                            customer_data.groupby(["품목코드", "월"], as_index=False)["금액(원)"]
                            .sum()
                            .sort_values(["품목코드", "월"])
                        )
                        product_monthly["날짜축"] = pd.to_datetime(product_monthly["월"] + "-01")

                        # 품목별 감소 기여도
                        pivot_prod = product_monthly.pivot_table(
                            index="품목코드",
                            columns="월",
                            values="금액(원)",
                            aggfunc="sum",
                            fill_value=0
                        )

                        contribution_rows = []
                        for prod in pivot_prod.index:
                            row = pivot_prod.loc[prod]
                            first_vals = [row[c] for c in first_half if c in row.index]
                            last_vals = [row[c] for c in last_half if c in row.index]
                            first_avg = float(np.mean(first_vals)) if len(first_vals) > 0 else 0.0
                            last_avg = float(np.mean(last_vals)) if len(last_vals) > 0 else 0.0
                            decline_amt = first_avg - last_avg

                            contribution_rows.append({
                                "품목코드": prod,
                                "전반부_평균": int(round(first_avg, 0)),
                                "후반부_평균": int(round(last_avg, 0)),
                                "감소액": int(round(decline_amt, 0)),
                                "변화액": int(round(last_avg - first_avg, 0)),
                                "변화율(%)": round(((last_avg - first_avg) / first_avg) * 100, 1) if first_avg > 0 else 0.0
                            })

                        contribution_df = pd.DataFrame(contribution_rows)
                        contribution_df = contribution_df.sort_values("감소액", ascending=False).reset_index(drop=True)

                        clean_and_safe_display(
                            contribution_df,
                            pinned_cols=["품목코드"],
                            text_cols=["품목코드"]
                        )

                        st.markdown("---")
                        st.markdown(f"### 📈 [{selected_cust_name}] 영업용 시각화")
                        st.caption("전체 매출 감소 여부와 감소 기여 품목을 한눈에 파악할 수 있게 구성했습니다.")

                        # ─────────────────────────────────────────────
                        # 그래프 1: 업체 전체 월별 매출 추이 + 추세선
                        # ─────────────────────────────────────────────
                        fig_total = go.Figure()

                        fig_total.add_trace(go.Scatter(
                            x=customer_total_monthly["날짜축"],
                            y=customer_total_monthly["금액(원)"],
                            mode="lines+markers",
                            name="월별 총매출",
                            line=dict(color="#1f77b4", width=3),
                            marker=dict(size=8)
                        ))

                        if len(customer_total_monthly) >= 2:
                            x_num = np.arange(len(customer_total_monthly))
                            y_num = customer_total_monthly["금액(원)"].values.astype(float)
                            coef = np.polyfit(x_num, y_num, 1)
                            trend = coef[0] * x_num + coef[1]
                            fig_total.add_trace(go.Scatter(
                                x=customer_total_monthly["날짜축"],
                                y=trend,
                                mode="lines",
                                name="추세선",
                                line=dict(color="red", dash="dash", width=2)
                            ))

                        fig_total.update_layout(
                            title="1️⃣ 업체 전체 월별 매출 추이",
                            height=430,
                            yaxis_tickformat=",",
                            yaxis_title="매출액(원)",
                            legend=dict(orientation="h", yanchor="bottom", y=1.02)
                        )
                        fig_total = add_year_month_axis(fig_total, customer_total_monthly["날짜축"])
                        st.plotly_chart(fig_total, use_container_width=True)

                        # ─────────────────────────────────────────────
                        # 그래프 2: 품목별 매출 감소 기여도
                        # ─────────────────────────────────────────────
                        top_contrib = contribution_df.head(12).copy()
                        if not top_contrib.empty:
                            fig_contrib = go.Figure()
                            fig_contrib.add_trace(go.Bar(
                                x=top_contrib["품목코드"],
                                y=top_contrib["감소액"],
                                marker_color=[
                                    "#d62728" if v > 0 else "#2ca02c"
                                    for v in top_contrib["감소액"]
                                ],
                                text=[f"{v:,}" for v in top_contrib["감소액"]],
                                textposition="outside",
                                name="감소액"
                            ))
                            fig_contrib.update_layout(
                                title="2️⃣ 품목별 매출 감소 기여도 (감소 큰 순)",
                                height=450,
                                yaxis_tickformat=",",
                                xaxis_title="품목코드",
                                yaxis_title="감소액(원)"
                            )
                            st.plotly_chart(fig_contrib, use_container_width=True)

                        # ─────────────────────────────────────────────
                        # 그래프 3: 감소 주도 품목 월별 추이
                        # ─────────────────────────────────────────────
                        top_products = contribution_df.head(5)["품목코드"].tolist()
                        top_product_monthly = product_monthly[product_monthly["품목코드"].isin(top_products)].copy()

                        if not top_product_monthly.empty:
                            fig_products = px.line(
                                top_product_monthly,
                                x="날짜축",
                                y="금액(원)",
                                color="품목코드",
                                category_orders={"품목코드": top_products},
                                markers=True,
                                title="3️⃣ 감소 주도 품목 월별 매출 추이 (Top 5)",
                                labels={"금액(원)": "매출액(원)", "날짜축": ""}
                            )
                            fig_products.update_layout(
                                height=460,
                                yaxis_tickformat=",",
                                legend=dict(orientation="h", yanchor="bottom", y=-0.35)
                            )
                            fig_products = add_year_month_axis(fig_products, top_product_monthly["날짜축"])
                            st.plotly_chart(fig_products, use_container_width=True)

                        # ─────────────────────────────────────────────
                        # 그래프 4: 전반부 vs 후반부 직접 비교
                        # ─────────────────────────────────────────────
                        if not contribution_df.empty:
                            comp_df = contribution_df.head(12).copy()
                            fig_bar = go.Figure()
                            fig_bar.add_trace(go.Bar(
                                x=comp_df["품목코드"],
                                y=comp_df["전반부_평균"],
                                name="전반부 평균",
                                marker_color="#3498db",
                                text=[f"{v:,}" for v in comp_df["전반부_평균"]],
                                textposition="outside"
                            ))
                            fig_bar.add_trace(go.Bar(
                                x=comp_df["품목코드"],
                                y=comp_df["후반부_평균"],
                                name="후반부 평균",
                                marker_color="#e74c3c",
                                text=[f"{v:,}" for v in comp_df["후반부_평균"]],
                                textposition="outside"
                            ))
                            fig_bar.update_layout(
                                title="4️⃣ 품목별 전반부 vs 후반부 평균 매출 비교",
                                barmode="group",
                                height=500,
                                yaxis_tickformat=",",
                                xaxis_title="품목코드",
                                yaxis_title="매출액(원)"
                            )
                            st.plotly_chart(fig_bar, use_container_width=True)

                            st.markdown("##### 📋 품목별 변화율 상세")
                            clean_and_safe_display(
                                contribution_df,
                                pinned_cols=["품목코드"],
                                text_cols=["품목코드"]
                            )


# ══════════════════════════════════════════════════════════
# TAB 5
# ══════════════════════════════════════════════════════════
with tab5:
    st.subheader("원자료(필터 적용됨)")
    clean_and_safe_display(
        q,
        pinned_cols=["거래처", "품목코드"],
        text_cols=["거래처", "품목코드", "품목명(공식)", "점착제코드", "점착제명", "가로폭이력", "최근날짜", "월"],
    )
