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
    padding-top: 1rem;
    padding-bottom: 1.2rem;
    padding-left: 0.7rem;
    padding-right: 0.7rem;
    max-width: 100%;
}
div[data-testid="stHorizontalBlock"] {
    gap: 0.6rem;
}
div[data-testid="stMetric"] {
    background: #fafafa;
    border: 1px solid #eeeeee;
    border-radius: 10px;
    padding: 0.5rem 0.7rem;
}
div[data-testid="stDataFrame"] {
    width: 100%;
}
[data-testid="column"] {
    width: 100% !important;
}
@media (max-width: 1024px) {
    .block-container {
        padding-left: 0.45rem;
        padding-right: 0.45rem;
    }
    html, body, [class*="css"]  {
        font-size: 14px;
    }
}
</style>
""", unsafe_allow_html=True)


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
            )
        st.dataframe(display_df, width="stretch", hide_index=True)
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
        "거래처", "품목코드", "품목표시", "점착제코드", "점착제명",
        "최근날짜", "가로폭이력", "분석_내역", "AI분석", "업체성향", "최근추세",
        "월", "구분", "월표기", "연도", "대표구분"
    }

    for col in display_df.columns:
        pinned = col in pinned_cols

        if col in text_cols or col in fixed_text_like_cols:
            column_config[col] = st.column_config.TextColumn(
                col,
                width="large" if col in ["품목표시", "가로폭이력", "분석_내역", "AI분석"] else "medium",
                pinned=pinned,
            )
            continue

        if pd.api.types.is_numeric_dtype(display_df[col]):
            if any(k in col for k in ["하락률", "증감률", "비율", "변화율", "CV"]):
                column_config[col] = st.column_config.NumberColumn(col, format="%.1f", pinned=pinned)
            elif any(k in col for k in ["M2", "수량", "판매량", "총량", "출고량"]):
                column_config[col] = st.column_config.NumberColumn(col, format="%,.1f", pinned=pinned)
            elif any(k in col for k in ["점수", "AI", "우선순위", "통계", "종합"]):
                column_config[col] = st.column_config.NumberColumn(col, format="%.1f", pinned=pinned)
            else:
                column_config[col] = st.column_config.NumberColumn(col, format="%,d", pinned=pinned)
        else:
            column_config[col] = st.column_config.TextColumn(col, width="medium", pinned=pinned)

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
        )

    st.dataframe(
        display_df,
        column_config=column_config,
        width="stretch",
        height=height if height else "auto",
        hide_index=True,
    )
    return None


def sorted_unique(series):
    if series is None:
        return []
    s = pd.Series(series).astype(str).str.strip()
    s = s[~s.isin(["", "0", "0.0", "nan", "NaN", "None", "<NA>"])]
    return sorted(s.unique())


def add_year_month_axis(fig, x_dates):
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
        fig.add_vline(
            x=year_dates[0],
            line_width=1,
            line_dash="dot",
            line_color="rgba(120,120,120,0.5)"
        )

    fig.update_layout(margin=dict(b=90))
    return fig


def sales_to_manwon_label(value):
    if pd.isna(value):
        return ""
    return f"{int(round(float(value) / 10000.0, 0)):,}"


def make_text_position_map(items):
    positions = ["top center", "bottom center", "middle right", "middle left"]
    return {str(item): positions[i % len(positions)] for i, item in enumerate(items)}


def make_indexed_series(df, group_col, value_col, time_col):
    temp = df.copy().sort_values([group_col, time_col])

    def _calc(g):
        base_candidates = g[value_col].replace(0, np.nan).dropna()
        if len(base_candidates) == 0:
            g["지수값"] = np.nan
        else:
            base = base_candidates.iloc[0]
            if pd.isna(base) or base == 0:
                g["지수값"] = np.nan
            else:
                g["지수값"] = (g[value_col] / base) * 100
        return g

    return temp.groupby(group_col, group_keys=False).apply(_calc)


def calc_cv(series):
    s = pd.to_numeric(pd.Series(series), errors="coerce").dropna()
    if len(s) == 0:
        return 0.0
    m = s.mean()
    if m == 0 or pd.isna(m):
        return 0.0
    return float(s.std(ddof=0) / m)


def calc_slope(values):
    y = pd.to_numeric(pd.Series(values), errors="coerce").fillna(0).values.astype(float)
    if len(y) < 2:
        return 0.0
    x = np.arange(len(y))
    return float(np.polyfit(x, y, 1)[0])


def scale_to_100(series, reverse=False):
    s = pd.to_numeric(pd.Series(series), errors="coerce").fillna(0)
    if len(s) == 0:
        return s
    if s.nunique() <= 1:
        return pd.Series([50.0] * len(s), index=s.index)
    mn, mx = s.min(), s.max()
    if mx == mn:
        out = pd.Series([50.0] * len(s), index=s.index)
    else:
        out = ((s - mn) / (mx - mn)) * 100.0
    if reverse:
        out = 100.0 - out
    return out.clip(0, 100)


def build_color_map(items):
    palette = [
        "#1f77b4", "#d62728", "#2ca02c", "#9467bd", "#ff7f0e",
        "#17becf", "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22"
    ]
    items = [str(x) for x in items]
    return {item: palette[i % len(palette)] for i, item in enumerate(items)}


def should_show_helper_chart(top_product_monthly):
    if top_product_monthly is None or top_product_monthly.empty:
        return False, 0.0

    temp = (
        top_product_monthly.groupby("품목표시", as_index=False)["금액(원)"]
        .sum()
        .rename(columns={"금액(원)": "총매출"})
    )
    temp = temp[temp["총매출"] > 0].copy()

    if len(temp) < 2:
        return False, 0.0

    max_v = temp["총매출"].max()
    min_v = temp["총매출"].min()

    if min_v <= 0:
        return True, np.inf

    ratio = max_v / min_v
    return bool(ratio >= 5), ratio


# ══════════════════════════════════════════════════════════
# 데이터 로드
# ══════════════════════════════════════════════════════════
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
            return pd.Series([None] * len(series), index=series.index)

        df = alias_df.copy()
        if not {"유형", "별칭", "공식코드"} <= set(df.columns):
            return pd.Series([None] * len(series), index=series.index)

        df = df[df["유형"].astype(str) == typ].dropna(subset=["별칭", "공식코드"])
        al = df["별칭"].astype(str).str.strip().tolist()
        cl = df["공식코드"].astype(str).str.strip().tolist()

        out = []
        for v in series.astype(str).fillna("").str.strip():
            m = next((c for a, c in zip(al, cl) if v == a), None)
            if m is None:
                m = next((c for a, c in zip(al, cl) if a and a in v), None)
            out.append(m)
        return pd.Series(out, index=series.index)

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


# ══════════════════════════════════════════════════════════
# 분석용 사전 계산
# ══════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def build_analysis_cache(q):
    df = q.copy()
    df["월"] = pd.to_datetime(df["날짜"], errors="coerce").dt.strftime("%Y-%m")
    df = df[df["월"].notna() & (df["월"] != "")].copy()

    if "품목명(공식)" not in df.columns:
        df["품목명(공식)"] = ""

    df["품목코드"] = df["품목코드"].astype(str)
    df["품목명(공식)"] = df["품목명(공식)"].fillna("").astype(str)
    df["품목표시"] = np.where(
        df["품목명(공식)"].str.strip() != "",
        df["품목코드"] + " | " + df["품목명(공식)"],
        df["품목코드"]
    )

    monthly_sales = (
        df.groupby(["거래처", "월"], dropna=False)["금액(원)"]
        .sum()
        .reset_index()
    )

    customer_total_monthly = (
        df.groupby(["거래처", "월"], dropna=False)["금액(원)"]
        .sum()
        .reset_index()
    )
    customer_total_monthly["날짜축"] = pd.to_datetime(customer_total_monthly["월"] + "-01")

    product_monthly = (
        df.groupby(["거래처", "품목코드", "품목명(공식)", "품목표시", "월"], dropna=False)["금액(원)"]
        .sum()
        .reset_index()
    )
    product_monthly["날짜축"] = pd.to_datetime(product_monthly["월"] + "-01")
    product_monthly["만원라벨"] = product_monthly["금액(원)"].apply(sales_to_manwon_label)

    all_months = sorted(df["월"].unique().tolist())

    return df, monthly_sales, customer_total_monthly, product_monthly, all_months


def build_priority_results(monthly_sales, detail_df, all_months):
    if len(all_months) < 2:
        return pd.DataFrame(), [], []

    mid_idx = len(all_months) // 2
    first_half = all_months[:mid_idx] if mid_idx > 0 else [all_months[0]]
    last_half = all_months[mid_idx:] if mid_idx < len(all_months) else [all_months[-1]]

    rows = []
    for cust_name in monthly_sales["거래처"].dropna().unique():
        cust_monthly = monthly_sales[monthly_sales["거래처"] == cust_name].sort_values("월").copy()
        cust_detail = detail_df[detail_df["거래처"] == cust_name].copy()

        first_data = cust_monthly[cust_monthly["월"].isin(first_half)]["금액(원)"]
        last_data = cust_monthly[cust_monthly["월"].isin(last_half)]["금액(원)"]

        avg_first = float(first_data.mean()) if len(first_data) > 0 else 0.0
        avg_last = float(last_data.mean()) if len(last_data) > 0 else 0.0
        total_sales = float(cust_monthly["금액(원)"].sum())

        decline_amount = max(0.0, avg_first - avg_last)
        decline_rate = (decline_amount / avg_first) if avg_first > 0 else 0.0

        monthly_vals = cust_monthly["금액(원)"].astype(float).tolist()
        slope = calc_slope(monthly_vals)
        cv = calc_cv(monthly_vals)

        recent_months = all_months[-3:] if len(all_months) >= 3 else all_months
        recent_data = cust_monthly[cust_monthly["월"].isin(recent_months)].sort_values("월")
        if len(recent_data) >= 2:
            recent_trend = recent_data["금액(원)"].pct_change().replace([np.inf, -np.inf], np.nan).dropna()
            recent_neg_ratio = float((recent_trend < 0).mean()) if len(recent_trend) > 0 else 0.0
            recent_avg_change = float(recent_trend.mean()) if len(recent_trend) > 0 else 0.0
        else:
            recent_neg_ratio = 0.0
            recent_avg_change = 0.0

        if "품목코드" in cust_detail.columns:
            first_products = cust_detail[cust_detail["월"].isin(first_half)]["품목코드"].nunique()
            last_products = cust_detail[cust_detail["월"].isin(last_half)]["품목코드"].nunique()
            product_decline = max(0, first_products - last_products)
            product_decline_ratio = (product_decline / max(1, first_products)) if first_products > 0 else 0.0
        else:
            product_decline_ratio = 0.0

        rows.append({
            "거래처": str(cust_name),
            "전반부_평균매출": int(round(avg_first, 0)),
            "후반부_평균매출": int(round(avg_last, 0)),
            "실제감소액": int(round(decline_amount, 0)),
            "하락률(%)": round(decline_rate * 100, 1) if avg_first > 0 else 0.0,
            "전체_매출액": int(round(total_sales, 0)),
            "기울기": slope,
            "CV": cv,
            "최근음수비중": recent_neg_ratio,
            "최근평균증감": recent_avg_change,
            "품목감소확산도": product_decline_ratio,
        })

    base_df = pd.DataFrame(rows)
    if base_df.empty:
        return base_df, first_half, last_half

    amount_component = scale_to_100(base_df["실제감소액"]) * 0.7 + scale_to_100(base_df["하락률(%)"]) * 0.3
    base_df["감소규모점수"] = amount_component * 0.60

    stat_component = (
        scale_to_100(base_df["기울기"], reverse=True) * 0.5 +
        scale_to_100(base_df["최근음수비중"]) * 0.3 +
        scale_to_100(base_df["CV"]) * 0.2
    )
    base_df["통계추세점수"] = stat_component * 0.20

    ai_component = (
        scale_to_100(base_df["품목감소확산도"]) * 0.35 +
        scale_to_100(base_df["전체_매출액"]) * 0.35 +
        scale_to_100(base_df["최근평균증감"], reverse=True) * 0.30
    )
    base_df["AI분석점수"] = ai_component * 0.20

    base_df["AI_우선순위점수"] = (
        base_df["감소규모점수"] +
        base_df["통계추세점수"] +
        base_df["AI분석점수"]
    ).round(1)

    comments = []
    for _, r in base_df.iterrows():
        msg = []
        if r["실제감소액"] > 0:
            msg.append(f"전후반 평균 감소 {int(r['실제감소액']):,}원")
        if r["하락률(%)"] >= 20:
            msg.append(f"하락률 {r['하락률(%)']:.1f}%")
        if r["기울기"] < 0:
            msg.append("월별 추세 하락")
        if r["최근음수비중"] >= 0.5:
            msg.append("최근 월별 악화 지속")
        if r["품목감소확산도"] >= 0.3:
            msg.append("감소가 여러 품목으로 확산")
        comments.append(" | ".join(msg) if msg else "추세 안정")

    base_df["분석_내역"] = comments

    result_df = base_df.sort_values(
        ["AI_우선순위점수", "실제감소액", "하락률(%)"],
        ascending=[False, False, False]
    ).reset_index(drop=True)

    result_df["순위"] = range(1, len(result_df) + 1)
    return result_df, first_half, last_half


# ══════════════════════════════════════════════════════════
# 견적 레퍼런스 AI 분석
# ══════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def build_quote_reference(q_ref):
    if q_ref.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    df = q_ref.copy()
    df["월"] = pd.to_datetime(df["날짜"], errors="coerce").dt.strftime("%Y-%m")
    df = df[df["월"].notna() & (df["월"] != "")].copy()

    if "품목명(공식)" not in df.columns:
        df["품목명(공식)"] = ""

    df["품목코드"] = df["품목코드"].astype(str)
    df["거래처"] = df["거래처"].astype(str)
    df["품목명(공식)"] = df["품목명(공식)"].fillna("").astype(str)

    base_group_cols = ["품목코드", "점착제코드", "점착제명"]
    base_group_cols = [c for c in base_group_cols if c in df.columns]

    overview = (
        df.groupby(base_group_cols, dropna=False)
        .agg(
            최저단가=("단가(원/M2)", "min"),
            최고단가=("단가(원/M2)", "max"),
            거래처수=("거래처", "nunique"),
            총출고횟수=("수량(M2)", "count"),
            총량_M2=("수량(M2)", "sum"),
            총매출액=("금액(원)", "sum"),
            개월수=("월", "nunique"),
        )
        .reset_index()
    )
    overview["월평균_출고량"] = np.where(
        overview["개월수"] > 0,
        overview["총량_M2"] / overview["개월수"],
        0
    )
    overview["월평균_매출"] = np.where(
        overview["개월수"] > 0,
        overview["총매출액"] / overview["개월수"],
        0
    )

    monthly_pc = (
        df.groupby(["품목코드", "거래처", "월"], dropna=False)
        .agg(
            월출고량=("수량(M2)", "sum"),
            월매출=("금액(원)", "sum"),
        )
        .reset_index()
        .sort_values(["품목코드", "거래처", "월"])
    )

    recent_unit = (
        df.dropna(subset=["단가(원/M2)"])
        .sort_values("날짜")
        .groupby(["품목코드", "거래처"], as_index=False)
        .tail(1)[["품목코드", "거래처", "단가(원/M2)", "날짜"]]
        .rename(columns={"단가(원/M2)": "최근단가", "날짜": "최근날짜"})
    )

    unit_extreme = (
        df.dropna(subset=["단가(원/M2)"])
        .groupby(["품목코드", "거래처"], dropna=False)
        .agg(
            최저단가=("단가(원/M2)", "min"),
            최고단가=("단가(원/M2)", "max"),
            총매출액=("금액(원)", "sum"),
            총량_M2=("수량(M2)", "sum"),
            개월수=("월", "nunique"),
        )
        .reset_index()
    )
    unit_extreme["월평균_출고량"] = np.where(unit_extreme["개월수"] > 0, unit_extreme["총량_M2"] / unit_extreme["개월수"], 0)
    unit_extreme["월평균_매출"] = np.where(unit_extreme["개월수"] > 0, unit_extreme["총매출액"] / unit_extreme["개월수"], 0)

    rows = []
    for (prod_code, cust_name), g in monthly_pc.groupby(["품목코드", "거래처"]):
        g = g.sort_values("월").copy()
        month_count = g["월"].nunique()
        avg_qty = float(g["월출고량"].mean()) if len(g) > 0 else 0.0
        avg_sales = float(g["월매출"].mean()) if len(g) > 0 else 0.0
        total_sales = float(g["월매출"].sum()) if len(g) > 0 else 0.0
        cv_sales = calc_cv(g["월매출"])
        slope_sales = calc_slope(g["월매출"].tolist())

        trend = "성장"
        if slope_sales < 0:
            trend = "감소"
        elif abs(slope_sales) < max(1, avg_sales * 0.02):
            trend = "안정"

        rows.append({
            "품목코드": str(prod_code),
            "거래처": str(cust_name),
            "개월수": int(month_count),
            "월평균_출고량": avg_qty,
            "월평균_매출": avg_sales,
            "총매출액": total_sales,
            "매출CV": cv_sales,
            "매출기울기": slope_sales,
            "최근추세": trend,
        })

    ref_detail = pd.DataFrame(rows)
    if ref_detail.empty:
        return overview, ref_detail, pd.DataFrame()

    ref_detail = ref_detail.merge(
        recent_unit,
        on=["품목코드", "거래처"],
        how="left"
    ).merge(
        unit_extreme[["품목코드", "거래처", "최저단가", "최고단가"]],
        on=["품목코드", "거래처"],
        how="left"
    )

    prod_bench = ref_detail.groupby("품목코드").agg(
        qty_p70=("월평균_출고량", lambda s: np.nanpercentile(s, 70) if len(s.dropna()) > 0 else 0),
        sales_p70=("월평균_매출", lambda s: np.nanpercentile(s, 70) if len(s.dropna()) > 0 else 0),
        unit_p30=("최근단가", lambda s: np.nanpercentile(s.dropna(), 30) if len(s.dropna()) > 0 else 0),
        cv_p30=("매출CV", lambda s: np.nanpercentile(s, 30) if len(s.dropna()) > 0 else 0),
    ).reset_index()

    ref_detail = ref_detail.merge(prod_bench, on="품목코드", how="left")

    types = []
    comments = []

    for _, r in ref_detail.iterrows():
        tags = []
        desc = []

        is_bulk = r["월평균_출고량"] >= r["qty_p70"] and r["qty_p70"] > 0
        is_high_sales = r["월평균_매출"] >= r["sales_p70"] and r["sales_p70"] > 0
        is_growth = r["최근추세"] == "성장"
        is_price_sensitive = pd.notna(r.get("최근단가")) and r["unit_p30"] > 0 and r["최근단가"] <= r["unit_p30"]
        is_small_test = r["월평균_매출"] < max(500000, r["sales_p70"] * 0.3)
        is_stable = r["매출CV"] <= r["cv_p30"]

        if is_bulk and is_high_sales:
            tags.append("대량출고형+고매출핵심형")
            desc.append("월평균 출고량과 매출 기여도가 모두 높은 핵심 거래처")
        if is_price_sensitive:
            tags.append("가격민감형")
            desc.append("최근단가가 상대적으로 낮은 편")
        if is_growth:
            tags.append("성장형")
            desc.append("최근 월매출 추세가 상승")
        if is_small_test:
            tags.append("소량테스트형")
            desc.append("월매출 규모가 낮아 테스트성 거래 가능성")
        if is_stable:
            tags.append("안정거래형")
            desc.append("월별 변동이 낮아 거래가 안정적")

        if len(tags) == 0:
            tags.append("일반거래형")
            desc.append("평균적인 거래 패턴")

        primary = tags[0]
        types.append(primary)
        comments.append(" / ".join(desc[:3]))

    ref_detail["업체성향"] = types
    ref_detail["AI분석"] = comments

    ref_detail["대표점수"] = (
        scale_to_100(ref_detail["총매출액"]).values * 0.35 +
        scale_to_100(ref_detail["월평균_출고량"]).values * 0.25 +
        scale_to_100(ref_detail["월평균_매출"]).values * 0.20 +
        scale_to_100(ref_detail["매출기울기"]).values * 0.10 +
        scale_to_100(ref_detail["최근단가"], reverse=True).values * 0.10
    )

    representative = (
        ref_detail.sort_values(
            ["품목코드", "대표점수", "총매출액"],
            ascending=[True, False, False]
        )
        .groupby("품목코드", as_index=False)
        .head(5)
        .reset_index(drop=True)
    )

    representative["최근날짜"] = pd.to_datetime(representative["최근날짜"], errors="coerce").dt.strftime("%Y-%m-%d")

    special_rows = []

    def pick_top(df_in, cond, label, sort_cols, ascending):
        sub = df_in[cond].copy()
        if sub.empty:
            return
        picked = (
            sub.sort_values(sort_cols, ascending=ascending)
            .groupby("품목코드", as_index=False)
            .head(3)
            .copy()
        )
        picked["대표구분"] = label
        special_rows.append(picked)

    pick_top(
        ref_detail,
        ref_detail["업체성향"] == "대량출고형+고매출핵심형",
        "대량출고형+고매출핵심형",
        ["품목코드", "총매출액", "월평균_출고량"],
        [True, False, False]
    )
    pick_top(
        ref_detail,
        ref_detail["업체성향"] == "가격민감형",
        "가격민감형",
        ["품목코드", "최근단가", "총매출액"],
        [True, True, False]
    )
    pick_top(
        ref_detail,
        ref_detail["업체성향"] == "성장형",
        "성장형",
        ["품목코드", "매출기울기", "총매출액"],
        [True, False, False]
    )
    pick_top(
        ref_detail,
        ref_detail["업체성향"] == "소량테스트형",
        "소량테스트형",
        ["품목코드", "월평균_매출", "최근단가"],
        [True, True, True]
    )

    lowest_unit = ref_detail.sort_values(["품목코드", "최저단가", "총매출액"], ascending=[True, True, False]).groupby("품목코드", as_index=False).head(3).copy()
    lowest_unit["대표구분"] = "최저단가 대표업체"
    special_rows.append(lowest_unit)

    highest_unit = ref_detail.sort_values(["품목코드", "최고단가", "총매출액"], ascending=[True, False, False]).groupby("품목코드", as_index=False).head(3).copy()
    highest_unit["대표구분"] = "최고단가 대표업체"
    special_rows.append(highest_unit)

    special_reference = pd.concat(special_rows, ignore_index=True) if special_rows else pd.DataFrame()

    return overview, representative, special_reference


def draw_quote_reference_chart(special_df):
    if special_df is None or special_df.empty:
        st.info("차트로 표시할 대표 업체 데이터가 없습니다.")
        return

    chart_df = special_df.copy()
    items = chart_df["품목코드"].astype(str).unique().tolist()
    color_map = build_color_map(items)

    fig = go.Figure()

    for item in items:
        sub = chart_df[chart_df["품목코드"].astype(str) == str(item)].copy()
        fig.add_trace(go.Scatter(
            x=sub["거래처"],
            y=sub["최근단가"],
            mode="markers+text",
            name=str(item),
            text=sub["대표구분"],
            textposition="top center",
            marker=dict(
                size=11,
                color=color_map[str(item)],
                line=dict(width=1, color="white")
            ),
            hovertemplate=(
                f"품목코드: {item}<br>"
                "거래처: %{x}<br>"
                "최근단가: %{y:,.0f}원/M2<br>"
                "구분: %{text}<extra></extra>"
            )
        ))

    fig.update_layout(
        height=460,
        margin=dict(l=20, r=20, t=30, b=90),
        xaxis=dict(title="", tickangle=-35),
        yaxis=dict(title="최근단가(원/M2)", tickformat=",.0f"),
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.28,
            x=0,
            xanchor="left"
        )
    )
    st.plotly_chart(fig, use_container_width=True)


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

with tab1:
    st.subheader("거래처별 → 품목별 출고 이력/최근 단가/가로폭")

    if q.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
    else:
        q1 = q.copy()
        for c in ["거래처", "품목코드", "점착제코드"]:
            if c in q1.columns:
                q1[c] = q1[c].astype(str)

        cols = ["거래처", "품목코드", "점착제코드", "점착제명", "가로폭이력", "최근날짜", "최근단가"]
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
            text_cols=["거래처", "품목코드", "점착제코드", "점착제명", "가로폭이력", "최근날짜"],
            height=460,
        )

with tab2:
    st.subheader("품목별 → 거래처별 출고 이력/총량/단가/매출")

    if q.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
    else:
        q2 = q.copy()
        for c in ["거래처", "품목코드"]:
            if c in q2.columns:
                q2[c] = q2[c].astype(str)

        cols = ["품목코드", "거래처", "최근날짜", "최근단가"]
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
            text_cols=["품목코드", "거래처", "최근날짜"],
            height=460,
        )

with tab3:
    st.subheader("🏷️ 견적 레퍼런스 — 기준 견적가 & 판매 동향")

    q_ref = q.copy()

    if "단가(원/M2)" in q_ref.columns:
        q_ref = q_ref[q_ref["단가(원/M2)"].notna() & (q_ref["단가(원/M2)"] > 0)]

    for col in ["품목코드", "품목명(공식)"]:
        if col in q_ref.columns:
            q_ref = q_ref[~q_ref[col].astype(str).str.contains("샘플", case=False, na=False)]

    if q_ref.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
    else:
        overview, rep_ref, special_ref = build_quote_reference(q_ref)

        st.markdown("### 1) 품목 기준 견적 레퍼런스")
        overview_cols = [
            "품목코드", "점착제코드", "점착제명",
            "최저단가", "최고단가", "거래처수", "총출고횟수",
            "총량_M2", "총매출액", "월평균_출고량", "월평균_매출"
        ]
        overview_cols = [c for c in overview_cols if c in overview.columns]

        clean_and_safe_display(
            overview[overview_cols],
            pinned_cols=["품목코드"],
            text_cols=["품목코드", "점착제코드", "점착제명"],
            height=360,
        )

        st.markdown("### 2) 업체 성향 AI 분석 기반 대표 레퍼런스")
        rep_cols = [
            "품목코드", "거래처", "업체성향", "AI분석",
            "최근단가", "최저단가", "최고단가", "최근날짜",
            "월평균_출고량", "월평균_매출", "최근추세", "총매출액"
        ]
        rep_cols = [c for c in rep_cols if c in rep_ref.columns]

        clean_and_safe_display(
            rep_ref[rep_cols],
            pinned_cols=["품목코드", "거래처"],
            text_cols=["품목코드", "거래처", "업체성향", "AI분석", "최근날짜", "최근추세"],
            height=520,
        )

        st.markdown("### 3) 대표 업체 레퍼런스 확장")
        special_cols = [
            "대표구분", "품목코드", "거래처", "업체성향", "최근단가",
            "최저단가", "최고단가", "월평균_출고량", "월평균_매출",
            "최근추세", "총매출액", "AI분석"
        ]
        special_cols = [c for c in special_cols if c in special_ref.columns]

        clean_and_safe_display(
            special_ref[special_cols] if not special_ref.empty else pd.DataFrame(columns=special_cols),
            pinned_cols=["대표구분", "품목코드", "거래처"],
            text_cols=["대표구분", "품목코드", "거래처", "업체성향", "최근추세", "AI분석"],
            height=520,
        )

        st.markdown("### 4) 대표 업체 최근단가 비교")
        draw_quote_reference_chart(special_ref)

        export_df = special_ref[special_cols] if not special_ref.empty else pd.DataFrame(columns=special_cols)
        rep_csv = export_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button(
            "📥 대표 업체 레퍼런스 CSV 다운로드",
            data=rep_csv,
            file_name="대표업체_레퍼런스.csv",
            mime="text/csv",
        )

with tab4:
    st.subheader("📉 AI 기반 매출 하락 업체 분석")
    st.caption("전후반기 매출감소 60% + 통계분석 20% + AI분석 20%")

    if q.empty or "날짜" not in q.columns or "금액(원)" not in q.columns or "거래처" not in q.columns:
        st.warning("분석에 필요한 데이터가 부족합니다.")
    else:
        detail_df, monthly_sales, customer_total_monthly_all, product_monthly_all, all_months = build_analysis_cache(q)
        priority_df, first_half, last_half = build_priority_results(monthly_sales, detail_df, all_months)

        if len(all_months) < 2:
            st.info("분석 기간이 너무 짧습니다. (최소 2개월 필요)")
        elif priority_df.empty:
            st.success("✅ 설정 기간 동안 매출이 하락한 업체가 없습니다.")
        else:
            top_count = max(1, int(np.ceil(len(priority_df) * 0.35)))
            top_priority = priority_df.head(top_count).copy()

            st.markdown("### 🎯 매출감소 추이 업체 LIST")

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
                height=520,
            )

            if edited_priority is not None:
                csv_priority = edited_priority.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                st.download_button(
                    "📥 매출감소 추이 업체 LIST CSV 다운로드",
                    data=csv_priority,
                    file_name="매출감소추이_업체_LIST.csv",
                    mime="text/csv",
                )

with tab5:
    st.subheader("원자료(필터 적용됨)")
    raw_cols = [c for c in q.columns if c != "품목명(공식)"]
    clean_and_safe_display(
        q[raw_cols],
        pinned_cols=["거래처", "품목코드"],
        text_cols=["거래처", "품목코드", "점착제코드", "점착제명", "가로폭이력", "최근날짜", "월"],
        height=540,
    )
